# Noninteractive add-in operations

Interactive use is unchanged. A caller opts in by passing a decision policy.

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
timer, and the outcome arrives through the completion callback.
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

`SetOperationPolicy` sets a session policy. It stays in force across
operations until `ClearOperationPolicy`, or another `SetOperationPolicy`,
replaces it; `Finish` does not close it. Each operation under it starts with
no decisions and no blocked prompt. `SetOperationPolicy` refuses with
`operation_already_running` while an operation runs. `ClearOperationPolicy`
is always safe to call, including twice.

`SetInteractionMode` returns a JSON string through both the VBA function and
`VCS.API("SetInteractionMode", mode)`. Modes are 0 (normal/interactive), 1 (silent),
and 2 (noninteractive). Every response includes `success`, `requested_mode`, and
`effective_mode`; success means the requested mode is effective when the call
returns. Selecting a mode starts no operation and posts no completion callback.

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
return VBA `Empty`, which does not establish acceptance. This development
contract uses capability detection rather than a numeric release-version gate:
require `success: true` and `effective_mode` equal to the requested mode; missing,
empty, or malformed responses require an add-in upgrade. Existing VBA calls
that ignore the return value continue to work. Rebuilding does not increment the
add-in's version, so its current version number alone cannot identify this
capability. M32 owns consumption of this contract on the MCP side.

`ImportObject` and `ExportObject` take no policy argument; they run under
the session policy. The MCP sets it around each call. `ImportObject` reads
its outcome before `Finish`, which restores the error level from before the
operation, so a logged error (for example a refused add-in form merge)
returns `success: false` with the first logged error in `error`. A prompt
the policy blocked during either call returns `decision_required` with the
decisions, like the multi-object calls, and a raised error after it goes
under `runtime_error`.

## Dialogs this mode prevents

- `MsgBox2` (the add-in's message boxes), including the printer-settings
  import.
- A logged error's message box before an MCP call's root begins. The
  policy scope opens with the root, so until then `Log.Error` only logs
  (`MCPCallBeforeRoot`). `MCPDebugLog` writes through `TryAppendToFile`,
  which drops a failed line without logging it.
- `frmVCSConflict` (merge conflicts).
- The source-folder picker when the project folder is unknown.
- `frmVCSMain` being left visible as a results window. Every site that opens
  it goes through `ShowMainForm`, which opens it hidden for a noninteractive
  run.
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
`NUIDialog`, not a standard Win32 dialog.

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
