# Noninteractive add-in operations

A caller opts in by passing a decision policy or selecting noninteractive mode
with `SetInteractionMode(2)`. Interactive calls retain their existing behavior.

```vba
VCS.MergeBuild "block"
VCS.RunFilteredTests "prefer_source"
```

An inline policy belongs to one operation. `SetOperationPolicy` supplies a
caller-owned policy across calls until `ClearOperationPolicy` or a replacing set.
Each admitted operation owns a fresh journal and blocked flag. Completion restores
owned scopes after success, failure, cancellation, and runtime errors.

## Policies

| Name | Effect |
| --- | --- |
| `block` | Acknowledge and log OK-only prompts. Block other prompts and conflicts. |
| `decline` | Answer No, Cancel, or Abort. Keep database objects on conflicts. |
| `prefer_source` | Apply each conflict's requested action, or overwrite from source when none is specified. |
| `prefer_database` | Keep the database object and skip its conflicting source file. |
| `skip` | Same effect as `prefer_database`. |

Conflict policies do not approve generic confirmations. An uncovered decision
returns `decision_required` with its journal. Silent mode remains distinct. It
returns the supplied prompt default, which can be Yes.

## Confirm requested state

`SetOperationPolicy` returns `{"success":true,"policy":"<lower-cased name>"}`.
`ClearOperationPolicy` returns `{"success":true}` and is idempotent.
Finishing an operation leaves a caller-owned policy in force. Setting a policy
while a root runs or is staged returns `operation_already_running` without changes.
An invalid policy also leaves existing state intact.

`SetInteractionMode` returns `success`, `requested_mode`, and `effective_mode`.
Modes are 0 for normal, 1 for silent, and 2 for noninteractive. Require success
and the requested effective mode before starting work. Missing or malformed
responses do not confirm a change. Existing statement calls remain valid.

An enclosing noninteractive scope or an active root prevents selecting a weaker
mode and returns `interaction_mode_refused`. An invalid mode returns
`invalid_interaction_mode`. These refusals preserve policy, journal, and ownership.
Mode-only noninteractive operations finish in normal mode with a clean journal
for the next operation. Caller-owned policies remain until their owner clears them.

## Operation results

`MergeBuild` returns a start marker or refusal. `{"success":true,"started":true}`
means the timer continuation was admitted. It does not establish merge completion.
An active callback adds `operation_id`; an inactive callback contributes no ID.
Read the terminal callback or operation log to establish the outcome.

`RunFilteredTests`, `ImportObject`, `ExportObject`, `ImportByType`, and
`ExportByType` return final results. The single-object and category calls use the
caller-owned policy. Nonempty journals appear as `decisions` on success, failure,
runtime errors, and conflict cancellation. Empty journals add no field.
Scoped conflict cancellation returns `success:false`, `cancelled:true`, an error,
and `logPath`. Ordinary failures have no `cancelled` field.

Refusal patterns include `invalid_decision_policy`, `merge_not_available`,
`operation_already_running`, and `decision_required`. Active callbacks receive
the same refusal. A blocked decision takes precedence over other outcomes and
retains the operation's own log path and escaped decision text.

A runtime exception retains its original text and `errorNumber`. If a decision
was blocked, the exception text also appears as `runtime_error` beside the primary
decision result. A call finishes only a root it acquired. Pre-acquisition errors
leave another caller's root, log, and policy intact.

Automation tests are headless and install a missing test helper as setup. Direct
VBA preflight offers installation under its effective policy. Its root owns any
blocked decision, including under mode-only or caller-owned policy selection.
Use the ribbon for the interactive test console.
A failed compile gate returns `project_not_compiled`, `success:false`, and
`cancelled:false`. When a decision was blocked too, the primary error is
`decision_required`; `run_error` and `run_error_pattern` retain the compile diagnosis.
Compile diagnostics are separate from runtime exceptions.

A results write failure supplies `results_error` and no nonexistent results path.
Callbacks retain available partial results inline when saving fails. Completion
captures diagnostics before teardown and makes one terminal delivery attempt
through its retained emitter. A payload construction fault can add
`completion_error` and `completion_error_number` without replacing the diagnosis.
The next operation can acquire a clean root after every completion path.

## Prompt boundaries

The policy handles `MsgBox2`, conflict choices, connection fallback prompts, and
printer-settings confirmations. Noninteractive runs keep `frmVCSMain` hidden,
including an already visible results form. Closing a noninteractive operation
window requests cancellation without a second confirmation.

Native Access save/discard, compile, trust, and error dialogs, VBA `MsgBox`, and
End/Debug dialogs remain outside this policy. The native object-close mechanics
are unchanged at this stage. Named build destinations, cancellation polling,
async admission isolation, and mutual compatibility sessions are separate work.

`ListAddinDialogs` and `DismissAddinDialog` enumerate responsive add-in forms.
`close` refuses during an operation; `cancel` requests cancellation without
finishing the root itself. These VBA calls require responsive Access COM.
A native dialog that blocks COM requires an external inspector or user action.

Rebuild and use disposable verification hosts as described in
[agentic-rebuild.md](agentic-rebuild.md) and [agent-test-runs.md](agent-test-runs.md).
The policy checks and public-operation workflow are described in
[Testing/NONINTERACTIVE-POLICIES.md](../Testing/NONINTERACTIVE-POLICIES.md).
