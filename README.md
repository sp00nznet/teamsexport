# teamsexport

Pull your Microsoft Teams chat history out of the local client cache and render it
as readable HTML. No tenant admin, no eDiscovery, no Graph app registration, no
waiting for someone to approve a compliance export.

Teams already has your messages on disk — it keeps whatever it has synced in a
Chromium IndexedDB (LevelDB) store inside its app data folder. This just reads
that and prints it nicely.

```
python teamsexport.py collect                 # on the Teams machine -> snapshot zip
python teamsexport.py extract snap.zip        # -> messages.jsonl (merges into existing)
python teamsexport.py render                  # -> teams-export.html
python teamsexport.py probe snap.zip          # what's in there: schema only, no message text
python teamsexport.py watch --interval 900    # all three, on a loop
```

## Windows executables

`build.cmd` produces two standalone exes in `dist\` — no Python, no pip, nothing
to install on the target machine:

| | |
|---|---|
| `teamsexport.exe` | console; same subcommands as above |
| `teamsexport-gui.exe` | windowed; **Export now**, **Parse a snapshot zip...**, **Open export**, **Probe**, and a **watch** toggle with an interval, over a log pane |

The GUI defaults to `%USERPROFILE%\TeamsExport` for its snapshots, `messages.jsonl`
and `teams-export.html`; change it in the box at the top. Everything runs on a
worker thread, so the window stays responsive, and a lock keeps a manual run and a
watch tick from stomping on the same jsonl.

Both bundle the LevelDB reader, so `extract` works on a machine that has never
seen pip. They're unsigned PyInstaller one-file builds (~10 MB / ~13 MB), so
SmartScreen will warn on first run and picky corporate AV may object — that's the
usual cost of an unsigned exe, not a sign of anything.

The build runs in a local `.venv` on purpose: PyInstaller refuses to run when the
obsolete `pathlib` PyPI backport is installed globally, which it is on plenty of
machines.

## Why the two-step split

`collect` is **stdlib-only** — it copies the LevelDB files into a zip and nothing
else. That's the only part that has to run on the locked-down corporate box with
Teams on it. Parsing and rendering happen wherever you like, off a zip you carry
around.

## The eviction problem

Teams is a cache, not an archive. It holds recent messages plus whatever you've
scrolled back through, and it evicts. So:

- **Scroll back** in a conversation you care about before collecting — that pulls
  older messages down into the local store.
- `extract` **merges** into `messages.jsonl` and dedupes, so repeated snapshots
  accumulate history rather than replacing it. Once a message is in your jsonl it
  stays, even after Teams drops it.
- `watch` automates that: snapshot → merge → re-render on an interval. Leave it
  running and your export only ever grows.

## Install (from source)

Not needed if you're using the exes above.

`collect` needs nothing but Python 3.11+.

`extract` needs a LevelDB/IndexedDB reader:

```
pip install -r requirements.txt
```

That pulls [`ccl_chromium_reader`](https://github.com/cclgroupltd/ccl_chromium_reader)
from GitHub (pure Python — no compiler, no `python-snappy` build). Google's
`dfindexeddb` does the same job but drags in `python-snappy`, which needs a C
toolchain on Windows, so it's out.

## Where it looks

Microsoft has moved this path twice, and the WebView2 profile directory varies,
so discovery is glob-based rather than hardcoded:

| Client | Path |
|---|---|
| New Teams (2.x) | `%LOCALAPPDATA%\Packages\MSTeams_*\LocalCache\Microsoft\MSTeams\EBWebView\*\IndexedDB\*.leveldb` |
| Teams 2.x preview | `%LOCALAPPDATA%\Packages\MicrosoftTeams_*\...` |
| Classic Teams (1.x) | `%APPDATA%\Microsoft\Teams\IndexedDB\*.leveldb` |

Fallback: recursive search under `%LOCALAPPDATA%\Packages\*Teams*` for anything
named `*.indexeddb.leveldb`. If Microsoft renames it again, that catches it.

Sibling `.blob` directories (large message values that don't fit inline) are
collected too.

**Running Teams holds some files open.** `collect` skips what it can't read and
records them in the zip's `manifest.json`. You get a more complete snapshot with
Teams closed, but it works either way — the `.ldb` files, which hold the bulk of
the history, are readable while it runs.

## Parsing approach

Every published parser for this breaks whenever Microsoft changes the object
store layout or the wrapper keys. So this one doesn't hardcode store names: it
walks *every* record in *every* object store and duck-types anything that looks
like a message (has `content`/`body` plus a message type or a timestamp) or a
conversation (has an id and a display name). Field lookups are
case-insensitive, because Teams mixes `composetime` and `composeTime` in the
same database.

That's deliberately loose. It picks up drafts and system events alongside real
messages; the HTML greys out `ThreadActivity/*` system noise rather than hiding
it.

### Conversation names

A real title (channel topic, group chat name) is used when one is cached —
including when it's nested in `threadProperties` rather than sitting at the top
level of the record, which is the common case.

Teams often hasn't cached one at all, though, and a sidebar full of
`19:...@thread.tacv2` is useless. So the fallback derives a name from the people
in the conversation, which is data we already have: **Alice Smith (chat)**,
**Bob Jones, Cara Diaz, Dan Fox +2 (channel)**. Whoever appears across the most
conversations is assumed to be you and dropped from the label, so 1:1 chats read
as the other person rather than "You, Alice". The raw thread id is still there as
the row's tooltip, and the filter box matches on it.

## The HTML

Single self-contained file. Conversation sidebar, live filter across senders and
message bodies with match highlighting, light/dark from your OS setting, no
external requests.

Message bodies are arbitrary HTML written by other people, so they go through an
allowlist sanitizer (stdlib `HTMLParser`): known-safe tags only, `href` limited to
http/https/mailto, `<script>`/`<style>` contents dropped, unclosed tags balanced.
Teams-specific tags are handled — `<emoji alt="😀">` becomes the emoji, `<at>`
becomes a styled mention, images become an `[image]` placeholder (they're
auth-gated URLs that won't load outside Teams anyway).

## Status

| | |
|---|---|
| `collect` | **working against a real Teams install**; skips locked files, limits itself to the Teams origin |
| `extract` | **working** — pulls thousands of messages off a live install. Survives a corrupt/partial LevelDB and skips unreadable records rather than aborting |
| `render` | working, verified end to end on real and synthetic data |
| `probe` | working — schema report (store names, field names, message types, record-key patterns) with no message text in the output, so it's safe to paste |
| `watch` | written, untested |
| GUI | builds and launches; buttons wired to the same functions the CLI uses |
| exes | both build clean and run — `teamsexport.exe selftest` passes, GUI window renders |
| `selftest` | `python teamsexport.py selftest` (or `teamsexport.exe selftest`) — asserts on the sanitizer, the duck-typed record walker, timestamp parsing, and that `extract` degrades gracefully on a garbage database |

Known rough edges, in order of how much they'd bother you:

- The `(no conversation id)` bucket — messages whose records carry no
  `conversationId` and no `conversationLink`. Most of these are now placed by
  falling back to the IndexedDB *record key*, which is frequently the thread id
  itself; `probe` reports on whatever is left.
- Real channel titles resolve only when Teams has cached them; everything else
  falls back to participant names.
- Attachment/card messages render as their text plus an `[image]` placeholder.

## Not doing

- **Graph API export.** Needs an app registration and, in most tenants, admin
  consent — which is the thing this exists to avoid.
- **Attachments.** Links are preserved; the files themselves live in SharePoint
  and OneDrive behind auth.
- **Decrypting anything.** This reads files your own user account can already
  read. It's your data, on your disk.

## Prior art

The forensics community got here first, and the paths and general approach come
from their work:

- [lxndrblz/forensicsim](https://github.com/lxndrblz/forensicsim) — Autopsy module for Teams IndexedDB
- [forensics.im](https://forensics.im/blog/parsing-microsoft-teams-indexeddb/) — write-up on Teams 1.x/2.x IndexedDB
- [cclgroupltd/ccl_chromium_reader](https://github.com/cclgroupltd/ccl_chromium_reader) — the LevelDB/IndexedDB/V8 reader doing the heavy lifting here
