# Noninteractive add-in operations

Pass a decision policy on the request that performs the work. Each admitted
operation owns its policy, initiating source, public ID, callback, and journal
through preparation, execution, and queued continuations.

```vba
result = VCS.ImportObject("query", "qryOrders", "block")
result = VCS.ExportObject("query", "qryOrders", "prefer_database")
result = VCS.MergeBuild("block")
result = VCS.RunFilteredTests("block")
```

Completion, failure, cancellation, and ordinary-operation heartbeat expiry release
owned interaction state. An unresolved launched rebuild keeps its reservation
because expiry cannot prove that the helper stopped. A disconnected client can
leave accepted work running under its original policy. Subsequent requests never
inherit that policy.

## Conflict policies

Preservation skips the conflicting transfer or deletion and leaves the item intact.

| Conflict | `prefer_source` | `prefer_database` | `skip` or `decline` | `block` |
| --- | --- | --- | --- | --- |
| Import overwrites an edited database object | Import source | Preserve object | Preserve object | Require decision |
| Export overwrites an edited source file | Preserve file | Export database version | Preserve file | Require decision |
| Import deletes an edited database object with missing source | Require decision | Preserve object | Preserve object | Require decision |
| Export deletes a tracked, changed orphan source file | Preserve file | Require decision | Preserve file | Require decision |
| Export encounters a new, untracked source file | Preserve file | Preserve file | Preserve file | Preserve file |

No noninteractive policy authorizes deletion of a conflicting edited object or
tracked, changed orphan file. Review and resolve that conflict interactively.
A preserving policy permits other work to continue when preservation is acceptable.
Existing cleanup of unchanged indexed orphans remains available.

The resolver evaluates the entire conflict list before applying resolutions. A
blocked deletion prevents other conflict resolutions in that list. Scanning,
temporary exports, and unchanged-orphan cleanup can already have occurred.

OK-only prompts are acknowledged and logged. `decline` answers Cancel for
Yes/No/Cancel, No for Yes/No, and Abort for Abort/Retry/Ignore. Other policies
block generic confirmations and return `decision_required`. Preference policies
authorize the transfers in the table, rather than arbitrary confirmation prompts.

## Single-object transfers and synchronization history

An explicit policy protects both core Access-object and separate component paths
of `ImportObject` and `ExportObject`. Naming one object defines its scope.
The policy still governs independent destination changes.

The shared index compares actual source and database content. Identical content
and a missing destination permit a transfer. A destination matching its recorded
synchronization baseline permits the next source edit, import, and test cycle.
Missing, stale, or unusable history cannot authorize overwriting divergent content.
An object type without a comparable representation returns
`policy_comparison_unsupported` rather than silently bypassing protection.

Successful transfers persist the affected object's synchronization history.
Skipped and blocked calls leave that object's history unchanged and do not
synchronize unrelated objects. Index format 5 stores distinct source and database
content baselines because an import can normalize source formatting. Formats 3
and 4 remain readable. Their missing database baseline receives conservative
handling until a successful explicit transfer establishes it. Older add-ins
cannot read format 5, so a downgrade loses usable synchronization history.

A preserving result reports `success:true,skipped:true,transferred:false` and a
decision journal. A transfer reports `transferred:true,skipped:false`. Aggregate
results include transfer and skip counts. Success alone does not prove transfer.
Journal entries retain their original fields and add conflict context, including
direction, requested action, classification, tracking, reason, and recovery.
Resolution names include `import_source`, `export_database`, `keep_source`,
`keep_database`, and `blocked`. Consumers preserve unfamiliar journal fields.

After `decision_required`, inspect the affected changes and follow the user's
existing authorization. Use `prefer_source` for an authorized import overwrite or
`prefer_database` for an authorized export overwrite. Do not automatically retry
with a preference policy just to obtain success. Ask for clarification when the
user's instructions do not establish which changes to retain.

## Legacy calls and retired activation

Omitting the policy retains legacy behavior, including automatic conflict choices
for automation. Source generally wins on import and the database on export.
Legacy single-object automation retains the historical disabled-index path.
Dropping a policy also removes its protections and is not conflict recovery.

`SetOperationPolicy` returns `policy_activation_retired` without changing state.
`ClearOperationPolicy` remains an idempotent harmless success.
Standalone `SetInteractionMode(2)` returns a refusal with `requested_mode`,
`effective_mode`, and `error_pattern`. Silent mode remains supported. It can
return supplied prompt defaults, including Yes. An unrelated request cannot
change an active operation's mode, source, policy, journal, callback, or timer.

Older add-in and MCP pairs used setter/cleanup sequences and bypassed conflict
detection for single-object automation even after acknowledging a policy.
Those acknowledgments do not establish the request-owned protection described here.
Use a matched consumer that forwards policy arguments on individual requests.

## Results and ownership

`Export`, `FullExport`, `ExportVBA`, single-object and category transfers, and
`RunFilteredTests` return final JSON. `Build` and `MergeBuild` return
`started:true` only after admitting the timer continuation. That response confirms
admission. Establish completion from the terminal result or operation log.
Ordinary VBA statement calls can ignore the returned value.

Admitted results and inspection identify `operation_id`, `initiating_source`, and
`decision_policy`. An absent policy is an empty string. See the
[request identity, queue, and cancellation contract](automation-contract.md#queues-identity-and-cancellation)
for identity lifetime, MCP entry points, refusal ownership, and cancellation targets.

A blocked result retains its journal and takes precedence over ordinary errors.
Runtime exceptions retain their original text and `errorNumber`; a blocked result
also carries `runtime_error`. Compile failures retain `project_not_compiled` or
their secondary compile diagnosis. A failed results write reports `results_error`
without inventing a results path. Available partial results remain in callbacks.
Final diagnostics are captured separately from restored interaction state.

## Workflow limits

Noninteractive runs suppress the add-in's main form and optional migration offer.
Required test-helper preparation still belongs to the operation. No policy
approves the current-project full-build confirmation. Supply an explicit source
folder to avoid that confirmation. Existing external API and MCP confirmation
bypasses remain available.

`BuildHeadless`, `MergeHeadless`, and `RebuildAddIn` retain fixed unattended
workflows and refuse nonempty policy arguments. Rebuild reserves ownership and
timer capacity before helper launch. Post-launch timer failure reports that the
helper launched. Separate rebuild recovery remains outside this contract.

Native Access save/discard, compile, trust, and error dialogs, VBA `MsgBox`, and
End/Debug dialogs remain outside decision policies. Native object-close behavior,
confirmed cancellation polling, named build destinations, migration recovery, and
mutual compatibility sessions remain separately qualified work.

Use the [automation contract](automation-contract.md) for entry-point details and
[disposable test workflow](agent-test-runs.md) for add-in qualification.

## Native close cancellation

Single-object and category import/export stop if Access leaves an object open
following a close attempt. A confirmed Cancel at Access's native save/discard
prompt returns `success: false` and `cancelled: true`, with the operation's own
`logPath` and decision journal. Other close failures remain failures without a
cancellation flag. A blocked decision keeps `decision_required` precedence.
Loaded add-in forms remain protected by the existing merge refusal.

Request decision policies do not dismiss Access-owned prompts. Inspect a native
prompt and confirm its Cancel button separately before choosing it. A failed close
must leave database objects, source files and the change index untouched.
