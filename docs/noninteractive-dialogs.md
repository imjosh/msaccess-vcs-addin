# Noninteractive add-in operations

Interactive use is unchanged. A caller opts in by passing a decision policy.

```vba
VCS.MergeBuild "block"
VCS.RunFilteredTests "prefer_source"
```

MCP passes the same strings from `vcs_import_objects` and `vcs_run_tests`.
The mode applies only to that operation. `Operation.Finish` restores the
previous interaction mode after success, failure, or cancellation. If the
operation never starts, the caller pops the scope immediately.

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

`MergeBuild` and `RunFilteredTests` return a start result, not an outcome:
`{"success":true,"started":true,"operation_id":...}`. The outcome arrives
through the completion callback. A request that cannot start (unknown policy,
another operation running, merge unavailable) returns
`{"success":false,"error_pattern":...,"error":...}` and posts the same payload
as an `error` callback. The patterns are `invalid_decision_policy`,
`merge_not_available`, `operation_already_running`, and `decision_required`.
A runtime error during a run that also blocked a prompt reports
`decision_required` with the decisions and the error under `runtime_error`.

`SetOperationPolicy` sets a session policy. It stays in force across
operations until `ClearOperationPolicy`, or another `SetOperationPolicy`,
replaces it; `Finish` does not close it. Each operation under it starts with
no decisions and no blocked prompt. `SetOperationPolicy` refuses with
`operation_already_running` while an operation runs. `ClearOperationPolicy`
is always safe to call, including twice.

`ImportObject` and `ExportObject` take no policy argument; they run under
the session policy. The MCP sets it around each call. `ImportObject` reads
its outcome before `Finish`, which restores the error level from before the
operation, so a logged error (for example a refused add-in form merge)
returns `success: false` with the first logged error in `error`.

## Dialogs this mode prevents

- `MsgBox2` (the add-in's message boxes), including the printer-settings
  import.
- `frmVCSConflict` (merge conflicts).
- The source-folder picker when the project folder is unknown.
- `frmVCSMain` being left visible as a results window. Every site that opens
  it goes through `ShowMainForm`, which opens it hidden for a noninteractive
  run.
- A second cancel confirmation when a noninteractive run's window is closed.
  Silent mode does not skip it; `MsgBox2` answers it with the caller's default.
- The one-time offers that headless test runs skip: `modTestAssert` is
  installed silently, and the "migrate `Debug.Assert`" offer is not shown, so
  it is never recorded as `decision_required`.

## Dialogs this mode does not prevent

Use the MCP inspector (`vcs_list_dialogs`, `vcs_dismiss_dialog`,
`vcs_recover_dialogs`, `vcs_automation_status`). Those calls use Win32
messages and stay available while Access COM is blocked. Details and
examples are in the MCP repository's `docs/DIALOGS.md`.

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
