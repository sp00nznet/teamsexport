#!/usr/bin/env python3
"""Pull Microsoft Teams chat history out of the local client cache.

Teams keeps whatever it has synced in a Chromium IndexedDB (LevelDB) store on
disk. This reads that store and renders readable HTML. No admin, no eDiscovery,
no Graph app registration.

  collect   copy the LevelDB store to a zip   (stdlib only -- run on the Teams box)
  extract   zip/dir -> messages.jsonl         (needs ccl_chromium_reader)
  probe     what's in a snapshot: schema only, no message text
  render    messages.jsonl -> export.html
  watch     collect+extract+render on a loop, merging into one growing store
  selftest  asserts
"""
from __future__ import annotations

import argparse
import collections
import datetime as dt
import html
import json
import os
import pathlib
import platform
import re
import shutil
import sys
import tempfile
import time
import zipfile
from html.parser import HTMLParser

# ---------------------------------------------------------------- discovery

# Teams has moved this path twice (classic Electron -> MicrosoftTeams -> MSTeams)
# and the WebView2 profile dir varies (Default / WV2Profile_tfw / ...), so glob.
GLOBS = [
    "LOCALAPPDATA/Packages/MSTeams_*/LocalCache/Microsoft/MSTeams/EBWebView/*/IndexedDB/*.leveldb",
    "LOCALAPPDATA/Packages/MicrosoftTeams_*/LocalCache/Microsoft/MSTeams/EBWebView/*/IndexedDB/*.leveldb",
    "LOCALAPPDATA/Microsoft/Teams/EBWebView/*/IndexedDB/*.leveldb",
    "APPDATA/Microsoft/Teams/IndexedDB/*.leveldb",
]


def discover() -> list[pathlib.Path]:
    hits: list[pathlib.Path] = []
    for g in GLOBS:
        var, _, rest = g.partition("/")
        root = os.environ.get(var)
        if root:
            hits += [p for p in pathlib.Path(root).glob(rest) if p.is_dir()]
    if not hits:  # last resort: Teams renamed the path again
        root = os.environ.get("LOCALAPPDATA")
        if root:
            for pkg in pathlib.Path(root, "Packages").glob("*Teams*"):
                hits += [p for p in pkg.rglob("*.indexeddb.leveldb") if p.is_dir()]
    hits = sorted(set(hits))
    # Teams' WebView2 profile also holds SharePoint/OneDrive origins we have no
    # business copying. Keep the teams.microsoft.com ones; fall back to all only
    # if the origin gets renamed out from under us.
    return [p for p in hits if "teams" in p.name.lower()] or hits


def cmd_collect(args) -> pathlib.Path:
    dirs = discover()
    if not dirs:
        sys.exit("No Teams IndexedDB found. Is Teams installed and signed in on this machine?")
    out = pathlib.Path(args.out or f"teams-snapshot-{dt.datetime.now():%Y%m%d-%H%M%S}.zip")
    manifest = {"host": platform.node(), "user": os.environ.get("USERNAME", ""),
                "collected": dt.datetime.now().astimezone().isoformat(),
                "sources": [], "skipped": []}
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for i, d in enumerate(dirs):
            manifest["sources"].append(str(d))
            # the .blob sibling holds values too big to inline in the LevelDB
            for src in [d, d.with_name(d.name.replace(".leveldb", ".blob"))]:
                if not src.is_dir():
                    continue
                for f in src.rglob("*"):
                    if not f.is_file():
                        continue
                    try:
                        z.write(f, f"db{i}/{src.name}/{f.relative_to(src).as_posix()}")
                    except OSError as e:  # Teams holds LOCK/CURRENT open while running
                        manifest["skipped"].append(f"{f}: {e}")
        z.writestr("manifest.json", json.dumps(manifest, indent=2))
    print(f"{out}  ({len(dirs)} store(s), {len(manifest['skipped'])} locked file(s) skipped)")
    return out


# ---------------------------------------------------------------- extraction

MSG_KEYS = {"content", "body"}
TIME_KEYS = ("composetime", "originalarrivaltime", "createdtime", "version")


def keys(d: dict) -> set[str]:
    # V8 objects deserialize with non-string keys (ints, mostly), so str() first.
    return {str(x).lower() for x in d}


def looks_like_message(d: dict) -> bool:
    k = keys(d)
    if not (k & MSG_KEYS):
        return False
    return "messagetype" in k or d.get("type") == "Message" or bool(
        k & {"composetime", "originalarrivaltime"})


def looks_like_conversation(d: dict) -> bool:
    # Teams ids are prefixed (19: threads and meetings, 48: drafts/mentions/etc),
    # which is a tighter gate than any of the name fields -- the channel title is
    # often nested in threadProperties rather than sitting at the top level.
    if "id" not in keys(d) or not str(get(d, "id", default="")).startswith(("19:", "48:")):
        return False
    return bool(keys(d) & {"displayname", "title", "topic", "threadproperties",
                           "members", "lastmessage", "type"})


def walk(obj, out_msgs: list, out_convs: list, depth: int = 0) -> None:
    """Teams nests messages under wrapper keys that change between builds; duck-type."""
    if depth > 12:
        return
    if isinstance(obj, dict):
        if looks_like_message(obj):
            out_msgs.append(obj)
        elif looks_like_conversation(obj):
            out_convs.append(obj)
        for v in obj.values():
            walk(v, out_msgs, out_convs, depth + 1)
    elif isinstance(obj, (list, tuple)):
        for v in obj:
            walk(v, out_msgs, out_convs, depth + 1)


def get(d: dict, *names, default=None):
    low = {str(k).lower(): v for k, v in d.items()}
    for n in names:
        v = low.get(n.lower())
        if v not in (None, ""):
            return v
    return default


def to_iso(v) -> str:
    if isinstance(v, (int, float)) and v > 1e11:  # epoch ms
        return dt.datetime.fromtimestamp(v / 1000, dt.timezone.utc).isoformat()
    if isinstance(v, str):
        try:
            return dt.datetime.fromisoformat(
                v.replace("Z", "+00:00")).astimezone(dt.timezone.utc).isoformat()
        except ValueError:
            if v.isdigit() and len(v) >= 12:
                return to_iso(int(v))
    return ""


def normalize(m: dict) -> dict | None:
    content = get(m, "content", "body", default="")
    if not isinstance(content, str) or not content.strip():
        return None
    conv = get(m, "conversationId", "threadId", default="")
    if not conv:
        link = get(m, "conversationLink", default="") or ""
        conv = link.rsplit("/conversations/", 1)[-1] if "/conversations/" in link else ""
    ts = ""
    for k in TIME_KEYS:
        ts = to_iso(get(m, k))
        if ts:
            break
    return {
        "id": str(get(m, "id", "messageId", "clientmessageid", default="")),
        "conv": str(conv),
        "sender": str(get(m, "imdisplayname", "displayName", "creator", "from", default="?")),
        "ts": ts,
        "type": str(get(m, "messagetype", default="")),
        "content": content,
    }


def conv_name(c: dict) -> str:
    n = get(c, "displayName", "title", "topic", "spaceThreadTopic")
    if not n:
        tp = get(c, "threadProperties", default={})
        n = get(tp, "topic", "spaceThreadTopic", "title") if isinstance(tp, dict) else None
    if not n:
        members = get(c, "members", default=[]) or []
        names = [get(x, "displayName", "imdisplayname") for x in members if isinstance(x, dict)]
        n = ", ".join(x for x in names if x)
    return str(n or "") if isinstance(n, (str, int)) else ""


def iter_records(db_dir: pathlib.Path):
    from ccl_chromium_reader import ccl_chromium_indexeddb as ix
    blob = db_dir.with_name(db_dir.name.replace(".leveldb", ".blob"))
    # A snapshot taken while Teams is running is partial by definition -- locked
    # files, half-written tables. Never let one bad table abort the whole export.
    try:
        wrapped = ix.WrappedIndexDB(db_dir, blob if blob.is_dir() else None)
        db_ids = list(wrapped.database_ids)
    except Exception as e:
        print(f"  ! cannot open {db_dir}: {e!r}", file=sys.stderr)
        return
    for db_id in db_ids:
        try:
            db = wrapped[db_id.dbid_no]
            stores = list(db.object_store_names)
        except Exception as e:
            print(f"  ! db {db_id}: {e!r}", file=sys.stderr)
            continue
        for store_name in stores:
            try:
                yield from db[store_name].iterate_records(
                    bad_deserializer_data_handler=lambda k, v: None)
            except Exception as e:
                print(f"  ! {db_id.name}/{store_name}: {e!r}", file=sys.stderr)


def load_store(path: pathlib.Path) -> dict:
    if not path.exists():
        return {}
    with path.open(encoding="utf-8") as f:
        return {json.loads(l)["key"]: json.loads(l) for l in f if l.strip()}


ID_RE = re.compile(r"(?:19|48):[^\s\"'|,}\])]+")


def id_in(s: str) -> str:
    """Pull a Teams thread id out of a string. IndexedDB record keys are often
    the conversation id, which is where messages with no conversationId hide."""
    m = ID_RE.search(s)
    return m.group(0) if m else ""


def source_dbs(source: str) -> tuple[list[pathlib.Path], str | None]:
    src = pathlib.Path(source)
    tmp = None
    if src.is_file() and src.suffix == ".zip":
        tmp = tempfile.mkdtemp(prefix="teamsexport-")
        with zipfile.ZipFile(src) as z:
            z.extractall(tmp)
        src = pathlib.Path(tmp)
    dbs = [src] if (src / "CURRENT").exists() else [
        p for p in src.rglob("*") if p.is_dir() and (p / "CURRENT").exists()]
    if not dbs:
        sys.exit(f"No LevelDB store under {source}")
    return dbs, tmp


def cmd_extract(args) -> pathlib.Path:
    dbs, tmp = source_dbs(args.source)
    out = pathlib.Path(args.out)
    store = load_store(out)          # merge: Teams evicts old cache, we keep it
    names: dict[str, str] = {}
    before = len(store)
    bad = 0
    for db in dbs:
        print(f"reading {db}", file=sys.stderr)
        for rec in iter_records(db):
            try:
                msgs: list = []
                convs: list = []
                walk(rec.value, msgs, convs)
                for c in convs:
                    cid, n = str(get(c, "id", default="")), conv_name(c)
                    if cid and len(n) > len(names.get(cid, "")):
                        names[cid] = n
                for m in msgs:
                    r = normalize(m)
                    if not r:
                        continue
                    if not r["conv"]:  # last resort: the record key is often the thread
                        r["conv"] = id_in(str(rec.key))
                    r["key"] = f"{r['conv']}|{r['id'] or r['ts']}|{hash(r['content']) & 0xffffffff}"
                    store.setdefault(r["key"], r)
            except Exception as e:  # one odd record must not cost you the export
                bad += 1
                if bad == 1:
                    print(f"  ! skipping unreadable record: {e!r}", file=sys.stderr)
    if bad:
        print(f"  ! {bad} record(s) skipped", file=sys.stderr)

    meta_path = out.with_suffix(".convs.json")
    known = json.loads(meta_path.read_text(encoding="utf-8")) if meta_path.exists() else {}
    known.update({k: v for k, v in names.items() if v})
    meta_path.write_text(json.dumps(known, indent=1), encoding="utf-8")
    with out.open("w", encoding="utf-8") as f:
        for r in store.values():
            f.write(json.dumps(r) + "\n")
    if tmp:
        shutil.rmtree(tmp, ignore_errors=True)
    print(f"{out}: {len(store)} messages (+{len(store) - before} new), {len(known)} conversations")
    return out


# ---------------------------------------------------------------- probe

def keypat(k) -> str:
    """Record key with digit runs collapsed, so thousands of keys become a few
    patterns -- and any number in them is redacted on the way."""
    return re.sub(r"\d+", "#", str(k))[:80]


def top(c: collections.Counter, n: int, indent: str = "  ") -> str:
    return "".join(f"{indent}{v:>6} x  {k}\n" for k, v in c.most_common(n)) or f"{indent}(none)\n"


def cmd_probe(args) -> None:
    """What's actually in the database: store names, key names, message types,
    record-key shapes. Never message text, so the output is safe to paste."""
    dbs, _ = source_dbs(args.source)
    stores, orphan_stores = collections.Counter(), collections.Counter()
    sigs, types, pats = collections.Counter(), collections.Counter(), collections.Counter()
    total = shaped = orphans = 0
    for db in dbs:
        for rec in iter_records(db):
            total += 1
            where = f"{rec.database_name}/{rec.object_store_name}"
            stores[where] += 1
            msgs: list = []
            try:
                walk(rec.value, msgs, [])
            except Exception:
                continue
            for m in msgs:
                shaped += 1
                r = normalize(m)
                if not r or r["conv"] or id_in(str(rec.key)):
                    continue                      # already placed in a conversation
                orphans += 1
                orphan_stores[where] += 1
                sigs[", ".join(sorted(keys(m)))] += 1
                types[str(get(m, "messagetype", default="(none)"))] += 1
                pats[keypat(rec.key)] += 1

    print(f"\n{total} records, {shaped} message-shaped, {orphans} with no conversation id\n")
    print("stores:")
    print(top(stores, 15), end="")
    if not orphans:
        return
    print("\n--- the orphan bucket ---\n\nwhich store they came from:")
    print(top(orphan_stores, 10), end="")
    print("\ntheir message types:")
    print(top(types, 10), end="")
    print("\ntheir record-key patterns (digits replaced with #):")
    print(top(pats, 10), end="")
    print("\ntheir field names:")
    print(top(sigs, 6), end="")


# ---------------------------------------------------------------- rendering

ALLOWED = {"b", "strong", "i", "em", "u", "s", "br", "p", "div", "span", "ul", "ol", "li",
           "code", "pre", "blockquote", "h1", "h2", "h3", "h4", "table", "thead", "tbody",
           "tr", "td", "th", "hr"}
VOID = {"br", "hr"}


class Sanitizer(HTMLParser):
    """Message bodies are arbitrary HTML written by other people. Allowlist it."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.out: list[str] = []
        self.open: list[str] = []
        self.drop = 0

    def handle_starttag(self, tag, attrs):
        if tag in ("script", "style"):
            self.drop += 1
            return
        a = dict(attrs)
        if tag == "emoji":
            self.out.append(html.escape(a.get("alt") or a.get("title") or ""))
        elif tag == "img":
            self.out.append('<span class="att">[image]</span>')
        elif tag == "at":
            self.out.append('<span class="mention">')
            self.open.append("span")
        elif tag == "a":
            href = a.get("href", "")
            if href.split(":", 1)[0].lower() in ("http", "https", "mailto"):
                self.out.append(f'<a href="{html.escape(href, quote=True)}" '
                                'target="_blank" rel="noopener noreferrer">')
                self.open.append("a")
        elif tag in ALLOWED:
            self.out.append(f"<{tag}>")
            if tag not in VOID:
                self.open.append(tag)

    def handle_startendtag(self, tag, attrs):
        self.handle_starttag(tag, attrs)
        if tag not in VOID and self.open:
            self.handle_endtag(tag)

    def handle_endtag(self, tag):
        if tag in ("script", "style"):
            self.drop = max(0, self.drop - 1)
            return
        tag = {"at": "span"}.get(tag, tag)
        if tag in self.open:
            while self.open:
                t = self.open.pop()
                self.out.append(f"</{t}>")
                if t == tag:
                    break

    def result(self) -> str:
        return "".join(self.out) + "".join(f"</{t}>" for t in reversed(self.open))

    def handle_data(self, data):
        if not self.drop:
            self.out.append(html.escape(data))


def clean(s: str) -> str:
    p = Sanitizer()
    p.feed(s)
    p.close()
    return p.result()


PAGE = """<!doctype html><meta charset="utf-8"><title>Teams export</title>
<style>
:root{color-scheme:light dark;--bg:#fff;--fg:#1b1b1f;--dim:#6b6b76;--line:#e3e3e8;--accent:#5b5fc7;--hl:#fff3a3}
@media(prefers-color-scheme:dark){:root{--bg:#141419;--fg:#e8e8ee;--dim:#9a9aa8;--line:#2c2c35;--hl:#5c5320}}
*{box-sizing:border-box}
body{margin:0;font:14px/1.55 -apple-system,Segoe UI,system-ui,sans-serif;background:var(--bg);color:var(--fg);display:flex;height:100vh}
#side{width:290px;flex:none;border-right:1px solid var(--line);display:flex;flex-direction:column}
#q{margin:10px;padding:8px 10px;border:1px solid var(--line);border-radius:6px;background:transparent;color:inherit;font:inherit}
#list{overflow:auto;flex:1}
#list div{padding:9px 12px;cursor:pointer;border-bottom:1px solid var(--line)}
#list div:hover{background:var(--line)}
#list div.on{background:var(--accent);color:#fff}
#list small{display:block;color:var(--dim);font-size:11px}
#list div.on small{color:#dcdcf5}
#main{flex:1;overflow:auto;padding:20px 26px}
h2{margin:0 0 14px;font-size:17px}
.m{margin:0 0 14px;max-width:900px}
.h{color:var(--dim);font-size:12px}
.h b{color:var(--fg);font-size:13px}
.c{margin-top:2px;overflow-wrap:anywhere}
.c pre{background:var(--line);padding:8px;border-radius:5px;overflow:auto}
.mention{color:var(--accent);font-weight:600}
.att{color:var(--dim)}
.sys{opacity:.6;font-style:italic}
mark{background:var(--hl);color:inherit}
#meta{padding:8px 12px;color:var(--dim);font-size:11px;border-top:1px solid var(--line)}
</style>
<div id=side><input id=q placeholder="filter conversations / messages"><div id=list></div><div id=meta></div></div>
<div id=main><h2>Pick a conversation</h2></div>
<script id=data type=application/json>__DATA__</script>
<script>
const D=JSON.parse(document.getElementById('data').textContent);
const list=document.getElementById('list'),main=document.getElementById('main'),q=document.getElementById('q');
document.getElementById('meta').textContent=D.meta;
let cur=null;
const esc=s=>s.replace(/[&<>]/g,x=>({'&':'&amp;','<':'&lt;','>':'&gt;'}[x]));
function match(f){return D.convs.filter(c=>!f||c.name.toLowerCase().includes(f)||c.id.toLowerCase().includes(f)
  ||c.msgs.some(m=>m.c.toLowerCase().includes(f)||m.s.toLowerCase().includes(f)))}
function draw(){
  const f=q.value.trim().toLowerCase(),hits=match(f);
  list.innerHTML='';
  hits.forEach(c=>{
    const d=document.createElement('div');
    d.append(c.name);
    const s=document.createElement('small');
    s.textContent=c.msgs.length+' msgs · '+(c.last||'').slice(0,10);
    d.append(s);
    d.title=c.id;
    d.onclick=()=>{cur=c;draw();};
    if(c===cur)d.className='on';
    list.append(d);
  });
  if(cur&&hits.includes(cur))show(cur,f);
  else if(cur)main.innerHTML='<h2>No match</h2>';
}
function show(c,f){
  main.innerHTML='<h2>'+esc(c.name)+'</h2>'+c.msgs.map(m=>
    '<div class="m'+(m.t.startsWith('ThreadActivity')?' sys':'')+'"><div class=h><b>'+esc(m.s)+
    '</b> · '+esc(m.d)+'</div><div class=c>'+m.c+'</div></div>').join('');
  if(f)main.querySelectorAll('.c,.h b').forEach(n=>hl(n,f));
}
function hl(n,f){
  for(const t of [...n.childNodes]){
    if(t.nodeType===3){
      const i=t.data.toLowerCase().indexOf(f);
      if(i<0)continue;
      const rest=t.splitText(i),after=rest.splitText(f.length),mk=document.createElement('mark');
      mk.textContent=rest.data;rest.replaceWith(mk);
    } else hl(t,f);
  }
}
q.oninput=draw;draw();
</script>
"""


def label(cid: str, msgs: list, me: str) -> str:
    """Teams doesn't always cache a title, but the participants are right there
    in the messages. 'Alice, Bob (chat)' beats a raw thread GUID every time."""
    if not cid:
        return "(no conversation id)"
    if cid.startswith("48:"):
        return cid[3:].capitalize()          # 48:drafts, 48:mentions, ...
    seen = list(dict.fromkeys(m["sender"] for m in msgs if m["sender"] not in ("", "?")))
    others = [s for s in seen if s != me] or seen
    if not others:
        return cid
    kind = ("meeting" if "meeting_" in cid else
            "chat" if "unq.gbl.spaces" in cid else "channel")
    extra = f" +{len(others) - 3}" if len(others) > 3 else ""
    return f"{', '.join(others[:3])}{extra} ({kind})"


def cmd_render(args) -> pathlib.Path:
    src = pathlib.Path(args.source)
    rows = [json.loads(l) for l in src.read_text(encoding="utf-8").splitlines() if l.strip()]
    meta_path = src.with_suffix(".convs.json")
    names = json.loads(meta_path.read_text(encoding="utf-8")) if meta_path.exists() else {}

    by_conv: dict[str, list] = {}
    for r in rows:
        by_conv.setdefault(r["conv"], []).append(r)
    # whoever posts across the most conversations is you -- drop yourself from
    # participant-derived labels, otherwise every 1:1 reads "You, Alice"
    seen_in: dict[str, set] = {}
    for cid, msgs in by_conv.items():
        for m in msgs:
            seen_in.setdefault(m["sender"], set()).add(cid)
    me = max(seen_in, key=lambda s: len(seen_in[s]), default="")

    convs = []
    for cid, msgs in by_conv.items():
        msgs.sort(key=lambda m: (m["ts"], m["id"]))
        convs.append({
            "name": names.get(cid) or label(cid, msgs, me),
            "id": cid,
            "last": msgs[-1]["ts"],
            "msgs": [{"s": m["sender"], "d": m["ts"][:19].replace("T", " "),
                      "t": m["type"], "c": clean(m["content"])} for m in msgs],
        })
    convs.sort(key=lambda c: c["last"], reverse=True)
    data = {"convs": convs,
            "meta": f"{len(rows)} messages · {len(convs)} conversations · "
                    f"rendered {dt.datetime.now():%Y-%m-%d %H:%M}"}
    blob = json.dumps(data).replace("</", "<\\/")
    out = pathlib.Path(args.out)
    out.write_text(PAGE.replace("__DATA__", blob), encoding="utf-8")
    print(f"{out}  ({len(rows)} messages, {len(convs)} conversations)")
    return out


# ---------------------------------------------------------------- watch

def export_once(snapshots, store, out) -> None:
    """collect -> merge -> render. The whole job, one call. Shared with the GUI."""
    snaps = pathlib.Path(snapshots)
    snaps.mkdir(parents=True, exist_ok=True)
    z = cmd_collect(argparse.Namespace(out=str(snaps / f"snap-{dt.datetime.now():%Y%m%d-%H%M%S}.zip")))
    cmd_extract(argparse.Namespace(source=str(z), out=store))
    cmd_render(argparse.Namespace(source=store, out=out))


def cmd_watch(args) -> None:
    """Teams evicts old cache; snapshotting on a loop is how you keep history."""
    while True:
        try:
            export_once(args.snapshots, args.store, args.out)
        except SystemExit as e:
            print(f"skip: {e}", file=sys.stderr)
        except Exception as e:
            print(f"error: {e!r}", file=sys.stderr)
        print(f"-- sleeping {args.interval}s (ctrl-c to stop)", file=sys.stderr)
        time.sleep(args.interval)


# ---------------------------------------------------------------- selftest

def cmd_selftest(_args) -> None:
    assert clean("<b>hi</b><script>alert(1)</script>") == "<b>hi</b>"
    assert clean('<a href="javascript:x">no</a>') == "no"
    assert 'target="_blank"' in clean('<a href="https://x/y">y</a>')
    assert clean('<emoji alt="\U0001f600"></emoji>') == "\U0001f600"
    assert clean("<p>a<b>b") == "<p>a<b>b</b></p>"          # unclosed tags balanced
    assert clean("a & b <3") == "a &amp; b &lt;3"
    assert clean('<div onclick="x()">t</div>') == "<div>t</div>"

    msg = {"id": "1700000000000", "messagetype": "RichText/Html", "content": "<b>hi</b>",
           "imdisplayname": "Ada", "composetime": "2023-11-14T22:13:20.000Z",
           "conversationId": "19:abc@thread.v2"}
    m, c = [], []
    walk({"messageMap": {"x": msg}, "junk": [1, 2]}, m, c)  # nested under a wrapper key
    assert m == [msg], m
    n = normalize(msg)
    assert n["sender"] == "Ada" and n["conv"] == "19:abc@thread.v2"
    assert n["ts"].startswith("2023-11-14T22:13:20")
    assert to_iso(1700000000000).startswith("2023-11-14T22:13:20")
    assert normalize({"messagetype": "x", "content": "  "}) is None
    assert normalize({"messagetype": "x", "content": "hi",
                      "conversationLink": "https://x/conversations/19:z"})["conv"] == "19:z"

    weird = {0: "int key", "content": "hi", "messagetype": "Text"}   # V8 gives int keys
    m3, c3 = [], []
    walk({"a": weird}, m3, c3)
    assert m3 == [weird] and normalize(weird)["content"] == "hi"

    conv = {"id": "19:abc@thread.v2", "type": "Space", "displayName": "PM guild"}
    m2, c2 = [], []
    walk([conv], m2, c2)
    assert c2 == [conv] and conv_name(conv) == "PM guild"
    assert conv_name({"id": "19:1", "type": "Chat", "title": "",
                      "members": [{"displayName": "Bo"}, {"displayName": "Cy"}]}) == "Bo, Cy"
    nested = {"id": "19:x@thread.tacv2", "threadProperties": {"topic": "PM guild"}}
    assert looks_like_conversation(nested) and conv_name(nested) == "PM guild"
    assert not looks_like_conversation({"id": "abc", "title": "not a teams id"})

    chat = [{"sender": "Me"}, {"sender": "Alice"}]
    assert label("48:drafts", [], "") == "Drafts"
    assert label("", [], "") == "(no conversation id)"
    assert label("19:x_y@unq.gbl.spaces", chat, "Me") == "Alice (chat)"
    assert label("19:meeting_abc@thread.v2", chat, "Me") == "Alice (meeting)"
    assert label("19:abc@thread.tacv2", [{"sender": s} for s in "ABCDE"], "A") \
        == "B, C, D +1 (channel)"
    assert label("19:solo@unq.gbl.spaces", [{"sender": "Me"}], "Me") == "Me (chat)"

    assert id_in("IdbKey(value='19:abc@thread.tacv2')") == "19:abc@thread.tacv2"
    assert id_in("48:notifications") == "48:notifications"
    assert id_in("1700000000000") == ""
    assert keypat("conv-000123-x") == "conv-#-x"

    # a snapshot of a running Teams is always partial -- extract must survive it
    try:
        import ccl_chromium_reader  # noqa: F401
    except ImportError:
        print("ok (extract not checked: pip install -r requirements.txt)")
        return
    d = pathlib.Path(tempfile.mkdtemp(prefix="teamsexport-test-"))
    (d / "CURRENT").write_text("MANIFEST-000001")
    (d / "000005.ldb").write_bytes(b"not a real sstable")
    cmd_extract(argparse.Namespace(source=str(d), out=str(d / "m.jsonl")))
    shutil.rmtree(d, ignore_errors=True)
    print("ok")


def main() -> None:
    p = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = p.add_subparsers(dest="cmd", required=True)

    c = sub.add_parser("collect", help="copy local Teams LevelDB to a zip (stdlib only)")
    c.add_argument("-o", "--out")
    c.set_defaults(func=cmd_collect)

    e = sub.add_parser("extract", help="zip or dir -> messages.jsonl (merges into existing)")
    e.add_argument("source")
    e.add_argument("-o", "--out", default="messages.jsonl")
    e.set_defaults(func=cmd_extract)

    pr = sub.add_parser("probe", help="report what's in a snapshot: schema only, no message text")
    pr.add_argument("source")
    pr.set_defaults(func=cmd_probe)

    r = sub.add_parser("render", help="messages.jsonl -> html")
    r.add_argument("source", nargs="?", default="messages.jsonl")
    r.add_argument("-o", "--out", default="teams-export.html")
    r.set_defaults(func=cmd_render)

    w = sub.add_parser("watch", help="collect+extract+render on a loop")
    w.add_argument("--interval", type=int, default=900)
    w.add_argument("--snapshots", default="snapshots")
    w.add_argument("--store", default="messages.jsonl")
    w.add_argument("-o", "--out", default="teams-export.html")
    w.set_defaults(func=cmd_watch)

    sub.add_parser("selftest").set_defaults(func=cmd_selftest)

    args = p.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
