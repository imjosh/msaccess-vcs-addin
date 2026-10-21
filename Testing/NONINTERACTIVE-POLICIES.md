# Verify noninteractive policies

Build this branch with the supported [source rebuild](../docs/agentic-rebuild.md).
Use Access 16 on Windows and a fresh disposable development host with its matching
repository and fixtures. Keep the installed library as the assertion receiver.
Follow [DISPOSABLE-HOSTS.md](DISPOSABLE-HOSTS.md) for compilation, fixture retention,
fresh reopen, evidence preservation, and whole-host disposal.

Run all Python verification drivers with optimization disabled. Do not use `-O`
or `-OO`. Keep `PYTHONOPTIMIZE` unset or `0`, and turn off equivalent interpreter
or launcher optimization settings. Assertions include validation and required
setup calls in these drivers. Removing assertions can omit both. The expected
interpreter state is `__debug__ == True` and `sys.flags.optimize == 0`.
Check the interpreter used for the drivers before starting native work:

```powershell
python -c "import sys; print(__debug__, sys.flags.optimize)"
```

Require `True 0`. Keep new or modified required setup outside `assert` expressions.
This prerequisite also applies to `tools/verify-caller-error-translation.py`.

Use the repaired MCP client that forwards a policy on each operation request.
Record its exact commit and the add-in commit before running the pair. Historical
MCP `232d807372477ac8547856b6b26076f6c30042c2` uses setter/cleanup activation and
cannot qualify this request-policy contract unchanged. Preserve compatibility
gates when selecting the repaired client. A successful policy-setter call from
an older pair does not establish compatibility or conflict protection here.

Before the long suite, demonstrate a successful public request-policy call through
the intended pair. Keep its arguments, result, source identities, and binary hashes.
For synchronous MCP dispatch, exercise `APIExecuteMCP` with the MCP-supplied public
operation ID. For asynchronous dispatch, exercise `APIAsync`, including typed
arguments such as an array of object types, a Boolean, and a trailing policy.
Require the policy, initiating source, and public ID to survive forwarding.
Handle synchronous admission refusals without waiting for a callback.

Run these modules through the installed runner:

```text
modTestDialogPolicy,clsTestInteractionScope,clsTestFormPolicyOverride,clsTestPromptSites,
clsTestOperationLifecycle,modTestModeOnlyScope,modTestOperationLifecycle,modTestScopedSync,
modTestExportFormPolicy,clsTestResultsPersistence,modTestBlockedCompile,
clsTestCallerErrorTranslation,modTestFileAccess,modTestHeadlessBuild,modTestRequestPolicy,
modTestCancelRequest,modTestAsyncCallerIdentity
```

Then run the unfiltered suite on a fresh host. Require passing assertions from
each selected module, no failed assertions or errored tests, and live repository
and fixture checks. Record EMPTY guards separately. Preserve existing lifecycle
and waiver dispositions with their original scope.

Run the public drivers with the repaired client's source on `PYTHONPATH`, using
that project's virtual environment. Keep evidence outside the contribution checkout.
Run commands serially with fresh labels:

```powershell
$env:PYTHONPATH = '<repaired paired client checkout>\src'
python Testing/check_request_policies.py --evidence '<evidence>' --label request-policy --client-source $env:PYTHONPATH
python Testing/check_noninteractive_policies.py --evidence '<evidence>' --label completion --client-source $env:PYTHONPATH
python Testing/check_refusal_ownership.py --evidence '<evidence>' --label refusal-owner --client-source $env:PYTHONPATH
python Testing/check_policy_workflow.py --evidence '<evidence>' --label workflow --client-source $env:PYTHONPATH
python tools/verify-caller-error-translation.py --binary '<closed source-built .accda>' --evidence '<new evidence folder>'
```

The request-policy driver checks retired calls, request isolation, targeted
cancellation, queue admission and execution refusal, stale continuations, public-ID
reuse, disconnect, expiry, and failed continuation handoffs. It also exercises
synchronous `APIExecuteMCP`, typed `APIExecuteAsync` forwarding, and statement calls that ignore
results. Inspect its receipts rather than treating this list as native evidence.

The workflow driver tests every policy through external API, asynchronous API,
and public MCP single-object tools. Its current object matrix covers queries,
database properties, and table data. It checks actual destination contents,
explicit skipped outcomes, decision journals, synchronization history, save/reopen,
ordinary edit/test cycles, history migration, deletion limits, and mixed conflict
lists. Record coverage for other object types separately. Do not infer support
for untested types from a passing query case. Use `--transports` only for focused
checks and record the restricted coverage. `--legacy-control` checks omitted-policy
behavior through the external API and is not request-policy qualification.

Check the direction-aware overwrite matrix against actual contents:

| Policy | Import into an edited database object | Export over an edited source file |
| --- | --- | --- |
| `block` | Preserve and return `decision_required` | Preserve and return `decision_required` |
| `prefer_source` | Import source | Preserve source and report `skipped:true` |
| `prefer_database` | Preserve object and report `skipped:true` | Export database version |
| `skip`, `decline` | Preserve object and report `skipped:true` | Preserve source and report `skipped:true` |

No policy authorizes conflicting deletion. Verify the deletion-specific outcomes
in [the policy contract](../docs/noninteractive-dialogs.md). New untracked files
survive all policies without blocking solely because they exist. A blocked mixed
list applies no conflict resolutions. A preserving policy can succeed without a
transfer, so inspect `skipped`, journals, and contents. Inspect a blocked conflict
using existing user instructions before choosing an overwrite policy. Obtain a
user decision when ownership remains unresolved. Removing the policy is not
recovery for a blocked deletion.

The completion driver exercises direct VBA, synchronous API, asynchronous API,
and MCP result paths. Its `cancel_poll` case polls the driver's own endpoint through
`CancelPollProbe` on a private MCP. It requires a first poll that answers false, no
new request inside the 500 ms throttle, an uncached poll that sees the switched answer,
and a poll after a last check just before midnight. Its `seam_*` cases arm
`SetCancelPollForTest` before an asynchronous run, so the Nth cancel check sees a
request. They stop the test loop, the export scan and object loops, the merge loop,
and the regular and fast build loops. Each needs one `cancelled` terminal and a
ready root. Export and test cases also have an unarmed run that completes.
The merge seam changes a saved query, then checks cancellation in the regular
file loop with no earlier critical error. A final case edits a module's
source to reproduce a critical import failure before the armed checkpoint; it
requires an error terminal. The build seams check both loops in a separate
disposable host because a cancelled build changes its host.
The refusal driver retains an active owner while other requests are refused.
Require unchanged owner state, zero refusal terminal attempts,
and exactly one owner completion. These drivers retain disposable user databases.
The translation driver edits only a disposable library copy and preserves its files.

Direct VBA test preflight offers to install a missing `modTestAssert`. `decline`
refuses it and `block` records an unanswered decision. API and MCP test runs keep
the headless helper-installation behavior. The optional TestAssert migration offer
is skipped under effective noninteractive execution. The translation probe tests
inline policies plus refusal of retired session activation and mode-only calls.

`SetOperationPolicy` must refuse with `policy_activation_retired`.
`ClearOperationPolicy` remains harmless success. Standalone `SetInteractionMode(2)`
must refuse while preserving `requested_mode`, `effective_mode`, and `error_pattern`.
A completed request leaves no policy for the next request or manual command.
Disconnect may leave admitted work running under its own policy. Supply policies
on individual requests instead of restoring a session or mode-only activation.

`ConfigureBlockedCompileFixture` is a parameterized setup-hook probe for disposable
user databases. Configure it through that host's `GlobalTestSetup`, calling the
installed library procedure. It can block a confirmation, select the compile
failure branch after the real compile command, force a results destination that
is a directory, or arm one completion construction fault. It must never run against
a shared development database. Invalid native VBA is unnecessary for these checks.

Verify synchronous JSON and callbacks for compile-only and blocked compile failure,
results write failure, construction failure, and the next successful operation.
Use one callback listener for the whole case sequence. Inspect raw JSON for duplicate
keys and require one terminal attempt per admitted root. Preserve the journal,
original exception, compile diagnostics, own log, and available partial results.
After each call require a ready root, no owned scope or visible main form, restored
logging and trapping, and no persistent policy.

Run `ScopedSyncOwnershipCheck`, `ScopedSyncCancelCheck` for import and export,
and `ExportFormPolicyCheck` outside the runner's root in a disposable user database.
Require their `OK` result. Use request-policy refusal and isolation checks in place
of the historical mode-2 activation procedure. Preserve failure artifacts before
closing the owned host.

Native dialog reachability, cancellation classification, named build output paths,
export migration recovery, compatibility sessions, and separately specified rebuild
recovery need their own evidence. Request isolation checks do not close the separate
cancellation/reentrancy submission gate. F3's rebuild reservation and post-launch
failure probes do not qualify the separate helper recovery protocol. Native
object-close mechanics remain inherited.

For UP-005's direct VBA scoped cancellation routes, run
`check_noninteractive_policies.py` with `--seam-mode=scoped` in a disposable host.
It creates a saved form Timer fixture and 300 queries. The timer reads the running
public operation ID and calls `DismissAddinDialog` during full scoped query export
and import. Require an accepted request for that same ID, a changed or processed
query before the operation returns `cancelled`, and a clean next scoped run. The
separate `ScopedSyncRequestCancelCheck` seam cannot replace this native route.
