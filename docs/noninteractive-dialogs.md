# Noninteractive add-in operations

> Release compatibility policy (owner decision, 2026-10-02): every add-in release changes its version. The supported release version defines the API contract; capability probing is not required to establish release compatibility. The spec assumes the server checks the installed add-in version and refuses unsupported releases before starting operations. The minimum supported release version must be stated when the release is assigned; do not infer it from a development rebuild. Per-call mode and policy acknowledgments still confirm the requested state and remain required. This policy supersedes earlier statements requiring capability checks instead of a version gate. It is a specification change, not evidence that version enforcement is already implemented.

Interactive use is unchanged. A caller opts in by passing a decision policy
or selecting noninteractive mode with `SetInteractionMode(2)`.

```vba
VCS.MergeBuild "block"
VCS.RunFilteredTests "prefer_source"
```

MCP passes the same strings from `vcs_import_objects` and `vcs_run_tests`,
and sets a session policy around `vcs_import_object` and `vcs_export_object`
(see Results). A policy passed to an entry point applies to that operation
only. `Operation.Finish` restores the previous interaction mode after success,
failure, or cancellation. A request that cannot begin (unknown policy, another
operation running) opens no scope and changes nothing.

A test run through automation is always headless: it shows no console and
answers prompts under its policy. MCP refuses `vcs_run_tests(noninteractive=False)`
with `interactive_tests_unsupported` before calling the add-in (X11). The
interactive test console is the ribbon's.

## Policies

| Name | Effect |
| --- | --- |
| `block` | OK-only prompts are logged and acknowledged. Any other prompt or merge conflict returns `decision_required` and is not approved. |
| `decline` | Confirmations answer No, Cancel, or Abort. Conflicts keep the database object. |
| `prefer_source` | Each conflict takes the action it asks for, and is overwritten from source when it asks for none. Other confirmations stay `decision_required`. |
| `prefer_database` | Conflicts keep the database object. Same effect as `skip`. |
| `skip` | Conflicts keep the database object and the source file is skipped. Same effect as `prefer_database`. |

Conflict policies never answer a generic confirmation such as "Overwrite?".
Only `decline` does, and it answers No, Cancel, or Abort.

`eimSilent` is unchanged: it logs `MsgBox2` and returns the caller's default.
Noninteractive mode does not use that default, because the default is
sometimes Yes.

An unresolved prompt sets `Operation.DecisionBlocked`, raises the operation
error level to critical, and makes `Finish` report `eorDecisionRequired`.
The MCP callback is an error with `decision_required: true` and a `decisions`
array. It is not `complete`.

## Results

`MergeBuild` returns a start result, not an outcome:
`{"success":true,"started":true,"operation_id":...}`. The merge continues on a
timer, and the outcome arrives through the completion callback. That callback
reports what the build recorded: a merge that rejected its target (for example
a source file name that does not match the open database) completes as an
`error` even when no critical error was logged, and a cancelled or
`decision_required` build keeps that outcome. A critical error fails a build
that recorded success, and a build that recorded no outcome completes as an
`error`, never as success.
`RunFilteredTests` is synchronous: it runs the tests before it returns, and its
return is the final results JSON. A request to either that cannot start
(unknown policy, another operation running, merge unavailable) returns
`{"success":false,"error_pattern":...,"error":...}` and posts the same payload
as an `error` callback. The patterns are `invalid_decision_policy`,
`merge_not_available`, `operation_already_running`, and `decision_required`.
A runtime error during a run that also blocked a prompt reports
`decision_required` with the decisions and the error under `runtime_error`.
The terminal callback carries the same `runtime_error` and `errorNumber`, since
an async caller never reads the return value; without a blocked prompt the
callback is an `error` whose message is the error text. A test-run callback
names `results_path` only when this run saved results. When the run finished but
the results file could not be written, the callback carries `results_error`
(the logged write error) instead.

A compile-gated test run alone reports `project_not_compiled`, `success: false`
and `cancelled: false`. If a setup prompt was blocked, `decision_required` is
primary and the compile text/pattern survive as `run_error` and
`run_error_pattern: project_not_compiled`, as defined in the shared contract
section 2 (A35). Callbacks also keep available partial results inline when a
results file cannot be saved. `results_error`, the journal and this operation's
log remain attached. Completion captures them before restoring owned scopes,
trapping and infrastructure; its retained callback emitter submits once after
restoration. A payload-construction fault adds `completion_error` and
`completion_error_number` without replacing the operation's primary diagnosis.

`BuildAs(source, output)` requires a fully qualified drive or UNC output path.
Relative, drive-relative and root-relative outputs are refused before operation
admission, leaving the open database and caller session unchanged. It builds
the source folder to the output file with no picker. With no arguments (the
ribbon) it keeps its pickers. Given only one path,
or a path it cannot use, it starts nothing and returns `invalid_build_path`
(posted as an `error` callback like any refusal): both paths are required, the
source folder must hold `vcs-options.json`, and the output must be a full path to
a file, in a folder that exists, that is not the add-in itself. Otherwise it
returns an empty string and the outcome arrives through the completion callback.
A successful full build adds `output_path`, the file it left open, to that
callback. `APICapabilities` returns
`{"success":true,"capabilities":["build_as_paths"]}`. This probe remains available
for development builds and clients that use feature discovery. Released add-ins
are identified by their release version; this probe is not required to establish
release compatibility. Call it directly with
`Application.Run "<add-in path>.APICapabilities"`, not through `API`: on an
add-in that predates it, Access refuses the call with a COM error (2517),
whereas `API` asked for a missing method stops on a modal "Run-time error 438"
dialog. `GetCapabilities` through `API` returns the same reply for callers that
already know the add-in has it.

Every call `APIAsync` launches on the timer (`Export`, `FullExport`,
`ExportVBA`, `Build`, `BuildAs`, `MergeBuild`, `RunFilteredTests`) posts exactly
one terminal callback, even when it never acquired a root. When the call
returns and nothing has posted one, the timer posts an `error` callback built
from the method's own return (`PostUnreportedOutcome`): a JSON object keeps its
fields (`error`, `errorNumber`, `error_pattern`, `decisions`) with
`success: false`, and a runtime error also carries `runtime_error`, as a root's
callback does. A test run that fails in its preflight before the root (for
example, installing `modTestAssert` raises) therefore reports the real error
instead of timing out. A Sub that returns Empty without posting, such as an
export whose `Begin` was refused, gets an `error` saying the operation did not
start. Nothing is added when a root completion or `RefusalJson` already posted
(`clsMCP.TerminalPosted`), when completion released the MCP instance, or when
the call left a staged root (a build or merge continuing on its timer), whose
continuation posts. Synchronous `API` returns are unchanged.

Async admission keeps the first caller's identity (A33). A launch made while a
timer is pending, an API is dispatching, or a root is running or staged returns
`success: false`, `error_pattern: operation_already_running`, and `api_refused:
true`, and posts one error to that incoming request's callback. Its refusal uses
a separate MCP instance, with no running journal or log attached. The admitted
method, arguments, callback, operation ID, cancellation URL, lease, interaction
state, journal and log remain with their owner. Timer dispatch repeats admission
before registering the callback, covering a timer delivered during a later
synchronous operation. Admission and dispatch errors also report to their own
caller, with `errorNumber` and `runtime_error`, and leave admission available.

`SetOperationPolicy` sets a session policy. It stays in force across
operations until `ClearOperationPolicy`, or another `SetOperationPolicy`,
replaces it; `Finish` does not close it. Each operation under it starts with
no decisions and no blocked prompt. `SetOperationPolicy` refuses with
`operation_already_running` while an operation runs. `ClearOperationPolicy`
is always safe to call, including twice. The acknowledgment is part of the
contract: `SetOperationPolicy` returns `{"success":true,"policy":"<lower-cased name>"}`
and `ClearOperationPolicy` returns `{"success":true}`. MCP requires exactly
these and reports anything else, such as the Empty an older add-in returns, as
`policy_unconfirmed` without starting the operation (M40). Like
`APICapabilities`, this is how a client checks the add-in build; the version
number does not change on a rebuild.

`API` refuses a call that arrives while another API command is still running
(a reentrant call). It returns a string starting `VCS_API_REFUSED: `, and
`APIAsync` returns `{"success":false,"error":"VCS_API_REFUSED: ..."}`. MCP
treats that as a failure on every call, with `error_pattern:
operation_already_running` and `api_refused: true`. When the refusal text says
the call arrived back in the project that sent it, the pattern is
`api_self_dispatch`, an add-in defect that a retry cannot fix (M39).

`SetInteractionMode` returns a JSON string through both the VBA function and
`VCS.API("SetInteractionMode", mode)`. Modes are 0 (normal/interactive), 1 (silent),
and 2 (noninteractive). Every response includes `success`, `requested_mode`, and
`effective_mode`; success means the requested mode is effective when the call
returns. Selecting a mode starts no operation and posts no completion callback.

An operation acquired after `SetInteractionMode(2)` opens its own interaction
scope even when no policy frame exists. Its blocked flag and decision journal
are isolated and released on completion, cancellation, or runtime error. The
terminal result retains that operation's decisions; the next operation starts
without them. Mode-only operations finish in normal mode. A caller-owned session
policy stays noninteractive until explicitly cleared or replaced, as above.

```json
{"success":true,"requested_mode":0,"effective_mode":0}
{"success":false,"requested_mode":0,"effective_mode":2,"error_pattern":"interaction_mode_refused","error":"..."}
```

An enclosing noninteractive scope prevents selecting normal or silent mode.
A running or staged root also prevents relaxing its mode. These refusals preserve
the scope, policy, decisions, blocked flag, error level, and root lease. The
scope owner must clear its session policy with `ClearOperationPolicy` (or close
its own scope token), and an active root must finish before a weaker mode can take
effect. If caller cleanup failed and left a session policy open after a failed
operation, an interactive selection returns this refusal; it never clears that
policy implicitly. Requesting an already effective mode succeeds without
changing the policy. An unknown mode returns `invalid_interaction_mode` and
leaves the state intact.

Callers that depend on confirmed interactive mode require an **A24 build or
later** and must check the structured result before starting work. Earlier builds
return VBA `Empty`, which does not establish acceptance. Require `success: true`
and `effective_mode` equal to the requested mode to confirm the requested state;
missing, empty, or malformed responses do not confirm it. Existing VBA calls
that ignore the return value continue to work. Release compatibility is based
on the supported release version. M32 owns consumption of the acknowledgment
contract on the MCP side.

`SetOperationPolicy` follows the same rule: callers dispatch only on
`{success: true, policy: <the requested policy, lower-cased>}`, and
`ClearOperationPolicy` counts as cleared only on `{success: true}`. An older
add-in returns VBA `Empty` (for example while busy), which acknowledges neither;
MCP reports that as `policy_unconfirmed` (M40). This acknowledgment confirms
operation state independently of release compatibility.

`ImportObject` and `ExportObject` take no policy argument; they run under
the session policy. The MCP sets it around each call. `ImportObject` reads
its outcome before `Finish`, which restores the error level from before the
operation, so a logged error (for example a refused add-in form merge)
returns `success: false` with the first logged error in `error`. A prompt
the policy blocked during either call returns `decision_required` with the
decisions, like the multi-object calls, and a raised error after it goes
under `runtime_error`.

Both calls close the object first, which can raise Access's native
save/discard prompt. The policy does not answer that prompt: a caller (or
`vcs_dismiss_dialog`) does. When the close is cancelled and the object is
still open, the call does not replace or export it. It returns
`success: false`, an `error` naming the object and the reason, `logPath`, and
`"cancelled": true`. `cancelled` is set only for a native Cancel (error 2501),
on import and export alike. A close that fails for any other reason and leaves
the object open returns the same `success: false` and `error`, without
`cancelled`. A normal call with no prompt is unchanged.

`ImportByType` and `ExportByType` also report a confirmed conflict-dialog
Cancel with `success: false`, `error: "Operation was canceled."`, `logPath`,
and `cancelled: true`. Both use the same scoped wrap-up. An ordinary failure
has no `cancelled` field; a blocked decision remains `decision_required`.

The same wrap-up reports a confirmed native save/discard Cancel when
`ImportByType` or `ExportByType` closes the open objects of a category (A36).
The close that raised error 2501 with the object still open stops the call
before anything is imported or exported, and the result is the cancellation
above (`error: "Operation was canceled."`, `logPath`, `cancelled: true`, and
the decision journal). A close that fails for any other reason, or raises 2501
after the object closed, is an ordinary failure with no `cancelled` field, and
a blocked decision remains `decision_required`. The native prompt itself stays
outside decision-policy dismissal. The cancellation is carried by
`Operation.NativeCloseCanceled`, which the scoped wrap-up reads beside the
conflict-dialog flag; later work on the abort-time export state (A37) shares
that wrap-up and keeps this field when it changes completion bookkeeping.

`ImportByType` and `ExportByType` share that error path. A raised error
returns the original `error` and `errorNumber`, and finishes only a root
the call itself began. An error before its `Begin` succeeded, such as an
overflowing type value, leaves another caller's running root, policy, log
and export resources alone.

Synchronous results report the decision journal the way the terminal
callback does. `ImportObject`, `ExportObject`, `ImportByType`, `ExportByType`
and the `RunFilteredTests` return carry `decisions` on every outcome
(success, failure, runtime error, cancel) when the journal is not empty: an
acknowledged OK-only prompt, an applied conflict policy, or a declined
prompt. An empty journal adds no key, so a run that met no prompt keeps its
old shape. `decision_required` stays primary when a prompt was blocked, and a
scoped `decision_required` result keeps `logPath`.

## Cancelling from MCP

`vcs_cancel_operation` records the request on the MCP server and returns
`cancel_requested: true`. The add-in finds it at its next checkpoint,
`Operation.CheckCancelRequest`, which asks `clsMCP.CheckCancelled`
(`GET /cancel-status/{operation_id}`), at most once every 500 ms. The request
takes the path closing `frmVCSMain` takes:

- A test run calls `TestRunner.Cancel`. The loop checks before each test,
  stops, and returns `cancelled: true` with the results so far.
- Export, build and merge log a critical "Canceled Operation", the level their
  loops already stop on. Export checks after each category scan and each
  object. Build checks after each fast-imported file and each component of the
  build loop. `MergeBuild` runs that same loop, so a merge is covered. The root
  then finishes `eorCanceled` and the callback is `cancelled`, not `error`. A
  blocked prompt still reports `decision_required`, and a runtime error still
  fails the root.

The entry points that run these loops can be cancelled: `Export`,
`FullExport`, `ExportVBA`, `Build`, `BuildAs`, `MergeBuild`, and
`RunFilteredTests`. These have no checkpoint and run to the end:
`ImportObject`, `ExportObject`, `ImportByType`, `ExportByType`,
`MergeAllSource`, `LoadSelected`, `RebuildAddIn`, `BuildHeadless`, and the
other synchronous single calls. The MCP server reports `cancel_not_honored`
for them.

Only a running root polls. A paused root (a user hook such as `AfterExport`,
or a form closing) does not, and neither does a private `clsOperation`. With
no MCP callback registered, `CheckCancelled` returns at once and sends no
request, so interactive use is unchanged. A failed poll reads as "not
cancelled". The first failure for a callback is logged, and later failures
are counted in `cancelPolls` and `cancelPollErrors` but not logged.

The poll uses `MSXML2.XMLHTTP` with `Cache-Control: no-cache` and an old
`If-Modified-Since`. Without those headers WinINet answers every later poll
with the first response, so a request made after the first check is never
seen. `ServerXMLHTTP` skips that cache, but it cannot connect while one of the
async log posts is in flight. That post finishes on the VBA thread the poll is
blocking, and the single-threaded callback server waits on it, so each poll
took about 2.5 s.

The test seam is `SetCancelPollForTest n`: a public procedure in `modAPI`
that a user project reaches through `Application.Run`, and
`Operation.SetCancelPollForTest` on a private instance. The Nth check from
now reports a request once, without asking MCP. 0 disarms it.

### Closing frmVCSMain over a root

`frmVCSMain` asks "Cancel Current Operation?" only for a running root that
closing it can stop (`MainFormWaitsForRoot`). A test run counts only while this
project's `TestRunner` is running it. Before this, a test-run root that nothing
in the project was running made the form refuse every close until the
heartbeat timed out, even after Yes. All copies of the add-in share one
registry key for operation state. `UpdateRegistry` now writes `RootProject`,
and a new `Operation` singleton adopts saved state only if its own add-in
file wrote it (`CanAdoptRegistryState`). This keeps a development copy loaded
beside the installed add-in from adopting the installed add-in's test run.

## Dialogs this mode prevents

- `MsgBox2` (the add-in's message boxes), including the printer-settings
  import.
- A logged error's message box before an MCP call's root begins. The
  policy scope opens with the root, so until then `Log.Error` only logs
  (`MCPCallBeforeRoot`). `MCPDebugLog` writes through `TryAppendToFile`,
  which drops a failed line without logging it.
- An options-loading error in `RunFilteredTests`, such as conflict markers in
  `vcs-options.json`. It reads `DefaultTestFilter` only after it has validated
  the policy and begun its root, so the policy acknowledges the error's box and
  the error goes to the run's log. An unknown policy or a refused root returns
  before the options load.
- `frmVCSConflict` (merge conflicts).
- The source-folder picker when the project folder is unknown.
- The Build As source-folder and save-as pickers when the caller passes both
  paths (`BuildAs(source, output)`).
- `frmVCSMain` being left visible as a results window. Every site that opens
  it goes through `ShowMainForm`, which opens it hidden for a noninteractive
  run. `Export`, `FullExport` and `ExportVBA` hide it through
  `SetMainFormVisible`, including a form that was already visible.
- A second cancel confirmation when a noninteractive run's window is closed.
  `ConfirmCancel` skips it: closing the window is the request to cancel, and no
  policy answers it. Silent mode does not skip it. `MsgBox2` shows it for a run
  a person started (it is a user-gesture prompt on an attended run). An
  unattended silent run, such as automation or a test run, gets the default
  answer, yes, and no window appears.
- The one-time offers that headless test runs skip: `modTestAssert` is
  installed silently, and the "migrate `Debug.Assert`" offer is not shown, so
  it is never recorded as `decision_required`.

## Dialogs this mode does not prevent

Use the MCP inspector (`vcs_list_dialogs`, `vcs_dismiss_dialog`,
`vcs_recover_dialogs`, `vcs_automation_status`). Those calls use Win32
messages, plus Microsoft Active Accessibility for Office NetUI boxes, and stay
available while Access COM is blocked. Details and examples are in the MCP
repository's `docs/DIALOGS.md`.

The same tools also cover a `MsgBox2` box the add-in shows on an interactive
run. Access draws that box (the `@`-separated bold `MsgBox` form) as a NetUI
`NUIDialog`, not a standard Win32 dialog. The inspector classifies it by window
class before it looks at the add-in's caption, so the box is listed as blocking
and `ready` is false (M36). An OK-only box is a `vba_msgbox`, which the `safe`
policy acknowledges unless its text is destructive. A box with two or more
buttons, such as the cancel confirmation, is `unknown` and is only reported.

- VBA `MsgBox` in the database under test.
- Microsoft Access error and warning dialogs.
- End/Debug runtime-error dialogs. End is available only when the agent asks
  for it. Debug is never clicked.
- Compile-error dialogs.
- VBA break mode, which is reported separately from a dialog.
- Trust-center prompts and "save changes?" confirms.

`ListAddinDialogs` and `DismissAddinDialog` (module `modDialogInspect`,
exposed as `VCS.ListAddinDialogs` and `VCS.DismissAddinDialog`) enumerate
open `frmVCS*` forms when Access is responsive. Refusals use the patterns
`not_addin_form`, `unsupported_action`, `operation_in_progress`, and
`not_open`.

- `action=close` closes a finished window and refuses while an operation is
  running.
- `action=cancel` sets the test-runner cancel flag, or the operation error
  level, and does not call `Finish` itself. The running operation stops at
  its next check. `interrupted` is true. That is not the same as closing a
  finished results window.

`DoCmd.SetWarnings` is not used to hide these dialogs.

## X17 session admission

Automation uses `APIExecute` / `APIExecuteAsync` with an accepted session envelope. Missing, stale or foreign sessions fail before operation, policy, mode or timer state changes. Deferred dispatch validates again before it starts. Handshake approval does not replace per-call policy/mode acknowledgments. A foreign session cannot clear another caller’s idle operation policy; disconnect clears only the owning idle policy; expiry preserves policy-owner entries. Running work retains its existing callback, cancellation and root ownership.

Legacy `API` / `APIAsync` operational calls require an explicit envelope forwarded by the admitted dispatcher; reentrant callers cannot borrow an active owner's envelope. Version/capability metadata remain diagnostic exceptions. On 2026-10-02 the developer explicitly limited X17 to MCP/CLI and API/APIAsync, preserving trusted manual/ribbon/direct-VCS VBA and admitted arbitrary VBA. These trusted routes do not independently require sessions, and admission is not a VBA sandbox. Policy-owner entries remain pinned until explicit owner cleanup, as separately selected by the developer. [Resumed X17 verification](../../verification/X17/RESUME.md) records the implementation and remaining qualification limits.
