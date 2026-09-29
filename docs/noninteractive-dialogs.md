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
| `prefer_source` | Conflicts are overwritten from source. Other confirmations stay `decision_required`. |
| `prefer_database` | Conflicts keep the database object. |
| `skip` | Conflicts skip the source file. |

`eimSilent` is unchanged: it logs `MsgBox2` and returns the caller's default.
Noninteractive mode does not use that default, because the default is
sometimes Yes.

An unresolved prompt sets `Operation.DecisionBlocked`, raises the operation
error level to critical, and makes `Finish` report `eorDecisionRequired`.
The MCP callback is an error with `decision_required: true` and a `decisions`
array. It is not `complete`.

## Dialogs this mode prevents

- `MsgBox2` (the add-in's message boxes).
- `frmVCSConflict` (merge conflicts).
- The source-folder picker when the project folder is unknown.
- `frmVCSMain` being left visible as a results window.
- A second cancel confirmation when a silent or noninteractive run's window
  is closed.

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

`ListAddinDialogs` and `DismissAddinDialog` enumerate open `frmVCS*` forms
when Access is responsive.

- `action=close` closes a finished window and refuses while an operation is
  running.
- `action=cancel` sets the test-runner cancel flag, or the operation error
  level, and does not call `Finish` itself. The running operation stops at
  its next check. `interrupted` is true. That is not the same as closing a
  finished results window.

`DoCmd.SetWarnings` is not used to hide these dialogs.
