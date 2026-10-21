# Automation contract

What an unattended caller can rely on from the add-in. A CI pipeline is out of
scope for this project ([Project Scope](../Wiki/Project-Scope.md)); this file
defines the surface such a pipeline is built against, so it can be relied on
without reading the implementation.

The tests that guard this surface are in
`Version Control.accda.src/modules/Tests/Core/modTestHeadlessBuild.bas`.

## Entry points and request policies

| Entry point | Policy support | Result |
| --- | --- | --- |
| `Export`, `FullExport`, `ExportVBA` | Optional policy | Final JSON |
| `ImportObject`, `ExportObject` | Optional third argument | Final JSON with transfer or skip outcome |
| `ImportByType`, `ExportByType` | Optional trailing policy | Final JSON with transfer and skip counts |
| `Build`, `MergeBuild` | Optional policy | Timer admission or preflight refusal JSON |
| `RunFilteredTests` | Optional policy | Final test JSON |
| `RunTests`, `RunTestsHeadless`, `OpenTestRunner`, `RunRoundtripTests` | Legacy behavior | Existing test or UI result |
| `BuildHeadless`, `MergeHeadless` | Fixed unattended workflow, nonempty policy refused | Final build JSON |
| `RebuildAddIn` | Fixed helper workflow, nonempty policy refused | Launch or refusal JSON |
| `BuildAs`, `LoadSelected`, `MergeAllSource`, `ExportSelected` | Legacy behavior | Existing statement or UI behavior |
| Options, repair, split, localization, hook, and casing commands | Legacy behavior | Existing statement or UI behavior |

Legacy entry points remain callable through `API` when idle. Manual commands
refuse before changing shared state while automated work is active. Inspection
and targeted cancellation remain available. A request's arguments, source, and
callback stay separate from the active operation until ownership is acquired.

An explicit policy belongs to one request. An omitted policy retains legacy
behavior and never inherits an earlier request's choice. See
[noninteractive operations](noninteractive-dialogs.md) for the conflict matrix,
deletion limits, synchronization history, and result fields. `SetOperationPolicy`
is retired and refuses without changing state. `ClearOperationPolicy` succeeds
harmlessly. Standalone mode 2 is refused with mode diagnostics. Silent mode
retains its existing default-answer behavior.

`Export`, `FullExport`, `ExportVBA`, and `Build` are JSON-returning functions.
VBA statement calls can still ignore their results. `Build` and `MergeBuild`
return `started:true` only after the continuation is admitted. They return before
the build finishes, so the terminal callback or operation log establishes its
outcome. `BuildHeadless` and `MergeHeadless` return after completion.

## VBE Error Trapping

`BuildHeadless`, `MergeHeadless`, and `RunTestsHeadless` (including test runs
that API or MCP force onto the headless path) raise VBE **Error Trapping** to
**Break on Unhandled Errors** for the whole call, including preflight that runs
before an operation begins. A caller already at that mode is left alone. The
original setting is restored when the call returns, including on a refused
preflight. Attended ribbon and non-headless API/MCP operations still use
**Break in Class Module** as their floor and do not downgrade a more
permissive caller setting.

## Result shape

```json
{
  "success": true,
  "logPath": "C:\\proj\\source\\logs\\Build_20260812_101500.log",
  "errorCount": 0,
  "warningCount": 2,
  "durationMs": 48210,
  "databasePath": "C:\\proj\\MyApp.accdb"
}
```

A failure adds `error` (always) and `errorNumber` (only when a VBA runtime error
escaped, rather than the build deciding it had failed):

```json
{
  "success": false,
  "error": "Build did not complete successfully. Check the log for details.",
  "logPath": "C:\\proj\\source\\logs\\Build_20260812_101500.log",
  "errorCount": 3,
  "warningCount": 0,
  "durationMs": 12044,
  "databasePath": "C:\\proj\\MyApp.accdb"
}
```

A request refused before the build started carries `success:false`, `error`,
and an applicable `error_pattern`. It has no admitted operation log. Refusals
include building the add-in from inside itself,
merging with no database open, a source folder holding no `vcs-options.json`,
and another operation already running.

`success: false` with `errorCount: 0` is possible and is not a contradiction. It
means the build was cancelled or abandoned rather than erroring — for example a
critical condition that stopped the run before anything was logged as an error.
Treat `success` as the verdict and the counts as detail.

## Why the log path is an out-param

`modBuild.Build` finishes by calling `Operation.Finish`, which calls
`modObjects.ReleaseObjects` and drops the `Log`, `Perf`, and `Options`
singletons. By the time control returns to the caller, `Log.SavedLogFilePath` is
an empty string on a freshly constructed object.

So the values in the result are captured inside `Build`'s `CleanUp` block,
through the optional `dOutcome As Dictionary` parameter, before `Operation.Finish`
runs. Any new field in the result shape has to be collected there too, not
computed afterwards. `LoadSingleObject` solves the same problem the same way with
its `strSavedLogPath` out-param.

## Headless merges take the slow path

`Options.SkipReopenBeforeMerge` lets a merge prepare the database in place rather
than closing and reopening it, which is much faster on a large project. It works
by handing off to `SetTimer "MergeReset"` and returning — the merge then resumes
on a fresh call stack.

That is incompatible with returning a result, because the call that would return
it has already unwound. `modBuild.Build` therefore ignores the option when
`dOutcome` was supplied, and takes the reopen path instead. A headless merge is
slower than the same merge from the ribbon, by design.

## Hook ordering

Hooks are named procedures in the *target* database, run through
`RunProcInCurrentProject`. They must be `Public` in a standard module and take no
parameters.

Full build:

1. `RunBeforeBuild` — before any object is imported, in the new blank database
2. (import)
3. `RunAfterBuild`
4. `ValidateAfterBuild`

Merge build:

1. `RunBeforeMerge` — after the database is prepared, before source is merged
2. (merge)
3. `RunAfterMerge`
4. `ValidateAfterBuild`

`ValidateAfterBuild` differs from the others in one way that matters: **it gates
success**. The `RunAfter*` hooks log an error if they raise one, but the build
still reports success. Validation is for the case where a build produced every
object and still is not shippable, and a pipeline needs to hear about it before
the file reaches users.

It must be a `Function` returning `True`. Anything else fails the build: `False`,
a raised error, a missing procedure, no return value (a `Sub`), or a return that
will not coerce to a Boolean. Silence is not approval when the answer decides
whether a release ships.

```vba
Public Function ValidateBuild() As Boolean
    ' Cheap smoke checks that a compile cannot catch.
    If DCount("*", "MSysObjects", "Type = -32768") = 0 Then Exit Function
    If Not IsAppConfigLoadable() Then Exit Function
    ValidateBuild = True
End Function
```

Set the option in `vcs-options.json` (`"ValidateAfterBuild": "ValidateBuild"`),
through `VCS.Options`, or on the options form as **Validate Build With** under
**Build Hooks**.

## Silent install

`/cmd "INSTALL SILENT"` installs without prompting. A runner installing a
downloaded release should name a status file, because nothing in a fresh
environment wrote the registry value that an add-in rebuilding itself relies on:

```
MSACCESS.EXE "Version Control.accda" /cmd "INSTALL SILENT C:\runner\install-status.json"
```

The path is everything after the keyword — quotes are optional and stripped,
and interior spaces are preserved, so no inner quoting is needed. The file is
written twice: `installing` as soon as the add-in's code starts, then `complete`
or `install-failed` with an `error`.

That first write is the useful one. **No file at all** means VBA never ran, which
in practice means the `.accda` was not in an Access trusted location and Access
is sitting behind a security prompt with nobody to dismiss it. Nothing inside the
add-in can report that condition, because its own code is what did not run.
Trusting the folder is a precondition of unattended install, and an absent status
file is how a pipeline detects that it was not met.

## Operation ownership and interaction

Root work uses a root lease through `Operation.TryBeginRoot` and
`clsRootOperationLease.Complete`, or the synchronous pair `Operation.Begin` and
`Operation.Finish`. Both create the same root. Acquisition captures the request's
initiating source, policy, stable public ID, and emitter before prompt-capable
source-folder or test-helper preparation.

| Use a **lease** when | Use **Begin / Finish** when |
|---|---|
| Ownership crosses a timer continuation (`DetachForContinuation`, `ResumeRoot`) | Work begins and ends in one procedure on the same stack |
| A module holds the root across an async boundary (web test bridge) | The caller is the only code that will complete the root |
| The site may not own the root at all (conditional harness entry) | No token needs to survive a form unload or timer tick |

`Begin` wraps `TryBeginRoot` and discards the lease handle. **`Finish` supplies the
live root token to itself**, so it performs no ownership check — only call it from the
procedure that called `Begin`. A lease's `Complete` validates the token, and
`Class_Terminate` calls `AbandonRootLease` if the lease is dropped without completing,
which is load-bearing for paths such as `ExecuteTests` that have no error handler
around the run body.

Core routines such as `modBuild.Build` and `LoadSingleObject` perform work only; the
orchestration boundary that acquired the lease completes it. Timer continuations carry
an opaque **root token** through `SetTimer` and resume with `Operation.ResumeRoot` /
`RunBuildFromContinuation`.

There is no nested operation object. A routine that can run either standalone or
inside a larger operation tests for the enclosing one and, when it finds it, owns
nothing: it does not take a lease, does not clear the log or console the root is
writing to, and completes nothing on the way out. `modTestRoundtrip` and
`modTestQuerySqlBuilder` both do this, keyed on `Operation.OperationType =
eotTestRun`.

Work that hands control to foreign code — user VBA through `Application.Run`, or
closing a form whose unload handler would read a running operation as a cancel —
takes a **pause** (`Operation.TryPause` → `clsOperationPause.ResumePause`). A pause
suspends the root on the same stack that started it and issues no continuation
token; pauses nest, and only the outermost resume restores the root.

Effective interaction mode belongs to the root. An explicit decision policy opens
its noninteractive scope. Unrelated requests cannot change that scope or the
root's captured source, policy, journal, result, callback, or timer state.
`AutomationSource` and `SourceName` read the captured source while the root is
active. Completion captures final diagnostics before restoring temporary
interaction state. `Attended` remains fixed at root acquisition.

`PromptWouldDisplay` is the single rule deciding whether a prompt blocks or is
logged, and `MsgBox2` is its only caller. It honors `blnUserGesturePrompt` only when
`Attended` is true, so a user-initiated cancel confirmation can appear during an
attended silent test run without relaxing the global mode. A refused root logs its
refusal instead of prompting whenever the caller is automated or explicitly
unattended.

Interaction mode reaches only as far as the operation instance does, which is one
VBA project. A test run drives a *second* project, so the mode the driver set is
invisible there; `modTestAssert.TestRunActive`, pushed by `clsTestRunner`, carries
the guarantee across and suppresses every prompt while it is set. See
[testing-strategy.md](testing-strategy.md) for how a run spans two projects.

Everything that reaches outside the operation instance — registry state for crash
recovery, the MCP completion callback, VBE error trapping, and `ReleaseObjects` — is
gated on being the instance published by `modObjects`, and only that instance adopts
restored state from the registry. A `New clsOperation` is therefore inert with
respect to the session, which is what makes the lifecycle testable: see
`modTestOperationLifecycle`, which drives private instances rather than the root
that is running the suite.

### Known gaps (not fixed by this refactor)

Two cross-form pairs still use `Begin` in one module and `Finish` in a form unload
handler, with no lease token carried across:

- `clsVersionControl.ShowOptions` → `frmVCSOptions.Form_Close`
- `clsVersionControl.SplitFiles` → `frmVCSSplitFiles.Form_Unload`

`Finish` ends whatever root is active; these are safe only because `Begin` refuses
when another operation is running and each form is the expected owner.

### Queues, identity, and cancellation

New requests carry no existing root token. Continuations explicitly carry their
expected internal token and restore only matching work. Timer installation checks
both operation and timer availability before changing queue fields. A failed
Build, MergeReset, or MergeResume handoff completes only its matching owner
unsuccessfully. MergeResume must be queued successfully before a project reset.

`APIAsync` preserves typed method arguments and returns a synchronous refusal when
admission fails. A queued request rechecks admission when it executes. A refusal
at that point makes one terminal delivery attempt through that request's own
emitter without registering it on another operation's callback.

Public operation IDs differ from internal tokens. MCP work retains its supplied
ID. Manual and direct-API roots receive generated IDs. Admitted IDs cannot be
reused, including after completion or project-state recreation. Final results,
callbacks, logs, and inspection identify the same public ID, effective policy,
and initiating source even after callback registration is released.
Synchronous MCP work uses `APIExecuteMCP(envelope, operation_id, method, arg1, arg2, arg3)`
(`APIMCP`, which took no session/1 envelope, is removed).
The method arguments are optional. `API` keeps external-API attribution. IDs are
case-sensitive and contain 1 to 128 UTF-16 code units without null characters.

Inspection leaves unexpired work unchanged and applies ordinary-operation
heartbeat expiry already due. Automated cancellation requires the explicit public ID. Missing or
mismatched IDs refuse without changing cancellation state. Manual controls bind
to the operation they display. A matching ID identifies a target and does not
grant permission to cancel it.

Rebuild reserves operation ownership and timer capacity before helper launch and
retains that reservation through handoff waits. It arms QuitForRebuild only after
confirmed handoff. A post-launch timer failure reports `helper_launched:true`.
The result cannot establish that the helper stopped. An unresolved launched
helper keeps its owned reservation across heartbeat expiry and same-project VBA
state recreation. The outer rebuild job and the child Build use distinct public
IDs; `build_operation_id` identifies the child's progress stream. The status file
remains authoritative for the outer rebuild outcome. Separate rebuild recovery
and native close cancellation remain later work.

A cancel is a request. The root records it and stops at its next check, and its
own result reports `cancelled` only when it finished cancelled. See
[Cancelling from MCP](noninteractive-dialogs.md#cancelling-from-mcp) for the
checkpoints, the calls that cannot be cancelled, and the poll's limits.

## Compatibility sessions

A server that wants more than the 5.1.0 surface opens a
`msaccess-vcs.session/1` session. `APIIdentity` reports the add-in version,
instance and protocol. `APIHandshake(request)` admits any server version with
that protocol: the add-in checks only the protocol identity, and the server owns
version compatibility. Another protocol is refused with
`compatibility_session_invalid` and `protocol_mismatch`. The session envelope
carries `session_id`, `server_instance`, `connection_id` and `addin_instance`.
`APIValidateSession` confirms one, `APIDisconnectSession` ends it, and a session
idle for 30 minutes expires.

Admitted work goes through `APIExecute(envelope, method, ...)`,
`APIExecuteAsync(envelope, callback, method, ...)` and
`APIExecuteMCP(envelope, operation_id, method, ...)`. `API` and `APIAsync` take
the envelope as the optional trailing `strSessionEnvelope`:

| Envelope | Outcome |
| --- | --- |
| Empty | Legacy call: the methods and argument counts of upstream 5.1.0 (`9dfb6a69`) only, with no session |
| This library's live session | Admitted: the full surface, not redirected, counted as a dispatch |
| Anything else | `compatibility_session_invalid` before dispatch |

A legacy call outside the 5.1.0 surface, or with more arguments than 5.1.0
took, is refused with `compatibility_session_required` and changes nothing.
5.1.0 had no decision-policy argument, so a legacy call runs with `edpAsk`.
Session-owned state, such as interaction mode set through `API` and request
policies, belongs to admitted work only. `GetVCSVersion` and `GetCapabilities`
read no session state and answer with any envelope. A legacy call made while
another call's session is entered stays legacy and never borrows that
admission. A `Build` launched by admitted work hands its envelope to the worker,
which reconnects with its own handshake; a legacy `Build` worker stays legacy.

Option overrides belong to the add-in instance, not to a session. A legacy
`SetOption` changes what admitted work in the same instance reads, as it did in
5.1.0. A legacy `RegisterSession` or `EndSession` is the 5.1.0 option session
and does not touch a compatibility session.

## Related

- [agentic-rebuild.md](agentic-rebuild.md) — rebuilding the add-in itself, and the
  status-file protocol this borrows
- [testing-strategy.md](testing-strategy.md) — `RunTestsHeadless`, JUnit output
- [Wiki/Continuous-Integration.md](../Wiki/Continuous-Integration.md) — the
  user-facing version of this material
