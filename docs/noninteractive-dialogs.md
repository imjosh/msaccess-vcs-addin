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

### Named Build As

`BuildAs(source, output)` requires both arguments when either is supplied.
The source folder must contain `vcs-options.json`. The output must be a fully
qualified drive or supported UNC file path with an extension and an existing
parent folder; it cannot name the add-in itself. Relative, root-relative and
drive-relative paths return `invalid_build_path` before the root is acquired,
the current database is closed, or any picker appears. An invalid request
leaves the current operation and open database alone.

A valid pair builds from the supplied source into the named file without
pickers. On success the terminal callback's `output_path` is the file the
build actually left open. The no-argument ribbon route retains its source
folder and save-file pickers.

`APICapabilities` advertises `build_as_paths` for optional discovery.
An unknown add-in build should be probed with
`Application.Run "<add-in path>.APICapabilities"` directly. Probing a
missing method through `API` can hold Access on a run-time error 438 dialog.
The accepted MCP `template` argument does not change the add-in build.

### Terminal callbacks from the timer

Every request that `APIAsync` queues on the timer gets one terminal callback.
A root posts its own callback when it completes, and `clsMCP.TerminalPosted`
then stops the timer from posting a second one. A root left staged for a build
or merge continuation posts from that continuation. When the method returns
without either, `modTimer.HandleAPIAsyncOperation` posts the method's own JSON
return to the request's emitter, as `complete` for `success: true` and as
`error` otherwise. An empty return posts an `error` with
`error_pattern: operation_not_started`. An exception in the timer posts an
`error` with `error_pattern: dispatch_error` and is not raised into Windows.

## Cancelling from MCP

A cancel is a request, not a result. The operation stops at its next check, and
its own terminal result says whether it was cancelled. Report an operation as
cancelled only when that result is `cancelled`, which means the root finished
with `eorCanceled`.

`vcs_cancel_operation` records the request on the MCP server. The add-in finds
it at its next checkpoint. `Operation.CheckCancelRequest` asks
`clsMCP.CheckCancelled`, which sends `GET /cancel-status/{operation_id}` to the
callback server at most once every 500 ms. A request that the poll finds goes to
`Operation.RequestCancel`. Every cancel route uses that one procedure.

- A test run calls `TestRunner.Cancel`. The runner checks before each test,
  stops, and returns `cancelled: true` with the results so far.
- Export, build, and merge log a critical "Canceled Operation", the level their
  loops already stop on. The root records the request and finishes
  `eorCanceled`, so the terminal callback is `cancelled`, not `error`.

Two outcomes take precedence over a request. A blocked prompt still reports
`decision_required`, and a runtime error still fails the root. The next root
starts without the request.

### Where the add-in checks

- `clsTestRunner.RunSelected` checks before each test.
- `modExport.ExportSource` checks after each category scan and after each
  exported object.
- `modBuild.Build` checks after each fast-imported file and after each
  component of the regular loop. `MergeBuild` runs that same loop.

`Export`, `FullExport`, `ExportVBA`, `Build`, `BuildAs`, `MergeBuild`,
and `RunFilteredTests` run these loops when `APIAsync` launches them, so a request
can stop them.

### Cannot be cancelled

These calls have no checkpoint that MCP reaches, and they run to the end.

- `ImportObject` and `ExportObject`.
- `ImportByType` and `ExportByType`. Their wrap-up reports a request that their
  root accepted from VBA as cancelled, but they have no MCP poll.
- `MergeAllSource` and `LoadSelected`.
- `RebuildAddIn`.
- `BuildHeadless` and `MergeHeadless`. They run the build loop, but a call
  through `API` or `APIExecuteMCP` registers no MCP callback, so the loop never polls.
  Their result has no cancelled form.
- Every other synchronous call through `API` or `APIExecuteMCP`, for the same reason.

### Limits

- Only the session `Operation` polls, and only while its root is running. A
  paused root, inside a user hook such as `AfterExport` or while a form closes,
  does not poll. A root staged for a timer continuation polls again only after
  it resumes. A private `clsOperation` instance never polls.
- With no MCP callback registered, `CheckCancelled` returns at once and sends
  nothing, so interactive use is unchanged.
- A failed poll reads as no request and leaves the result alone. The first
  failure for a callback goes to the log without counting as an error of the
  run. `GetDiagnosticData` counts every poll in `cancelPolls` and every failure
  in `cancelPollErrors`.
- A request between checkpoints waits for the next one. A long single step,
  such as one large object, finishes first.

The poll uses `MSXML2.XMLHTTP` with `Cache-Control: no-cache` and an old
`If-Modified-Since`. Without those headers WinINet answers every later poll with
the first response, so a request made after the first check is never seen.
`ServerXMLHTTP` skips that cache, but it waits on the add-in's own asynchronous
log posts, and each poll took about 2.5 s. The throttle compares `Timer`
values. A value smaller than the last check means midnight passed, so the poll
runs.

`modAPI.SetCancelPollForTest n` is the test seam. A user project reaches it
through `Application.Run`. A test calls `Operation.SetCancelPollForTest` on a
private instance. The nth check from then on reports a request once, without
asking MCP, and 0 disarms it.

### Cancelling from VBA

`VCS.DismissAddinDialog("frmVCSMain", "cancel", operation_id)` sends the same
request to `Operation.RequestCancel`. Every caller passes the running
operation's public ID. A missing ID refuses with `operation_id_required`, and
an ID that does not name the active operation refuses with
`operation_id_mismatch`. A staged or paused root refuses with
`operation_not_running`. A refusal changes nothing. An accepted call returns `requested: true` and does not call `Finish`.
The operation stops at its next check, and its own result reports whether it was
cancelled. A synchronous API call holds the dispatcher until it returns, so the
API route reaches a root only while no API call is running, such as a build or
merge continuation on the timer.

`ImportByType` and `ExportByType` report a request that their root accepted as
`error: "Operation was canceled."` with `cancelled: true` and `logPath`. A
runtime error after the request is an ordinary failure.

`DismissAddinDialog` is the add-in's own VBA procedure, and it needs Access to
be responsive. The MCP tools `vcs_list_dialogs` and `vcs_dismiss_dialog` are a
separate inspector in the MCP server. They read and press native windows with
Win32 messages and do not call `DismissAddinDialog`.

### Closing frmVCSMain over a root

Closing `frmVCSMain` asks "Cancel Current Operation?" only for a running root
that closing it can stop (`MainFormWaitsForRoot`). A test-run root counts only
while this project's `TestRunner` is running it. A noninteractive run skips the
confirmation, because closing the window is the request. On Yes, a test run
calls `TestRunner.Cancel`, and any other root calls `Operation.RequestCancel`.
The form stays open until the owning procedure finishes the root. All copies of
the add-in share one registry key for operation state. A new `Operation`
singleton adopts saved state only if its own add-in file wrote it
(`CanAdoptRegistryState`), so a development copy does not adopt the installed
add-in's test run.
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
End/Debug dialogs remain outside decision policies. Native object-close behavior and named build destinations are described above.
Migration recovery and mutual compatibility sessions remain separate work.

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
