# Verify noninteractive policies

Build this branch with the supported [source rebuild](../docs/agentic-rebuild.md).
Use Access 16 on Windows and a fresh disposable development host with its matching
repository and fixtures. Keep the installed library as the assertion receiver.
Follow [DISPOSABLE-HOSTS.md](DISPOSABLE-HOSTS.md) for compilation, fixture retention,
fresh reopen, evidence preservation, and whole-host disposal.

Run these modules through the installed runner:

```text
modTestDialogPolicy,clsTestInteractionScope,clsTestFormPolicyOverride,clsTestPromptSites,
clsTestOperationLifecycle,modTestModeOnlyScope,modTestOperationLifecycle,modTestScopedSync,
modTestExportFormPolicy,clsTestResultsPersistence,modTestBlockedCompile,
clsTestCallerErrorTranslation,modTestFileAccess,modTestHeadlessBuild
```

Then run the unfiltered suite on a fresh host. Require passing assertions from
each selected module, no failed assertions or errored tests, and live repository
and fixture checks. Record EMPTY guards separately.

Use the policy-aware legacy MCP client at commit
`232d807372477ac8547856b6b26076f6c30042c2` of
[`msaccess-vcs-mcp`](https://github.com/imjosh/msaccess-vcs-mcp) for this intermediate
API, or use an explicitly owned COM harness. This pin identifies historical test
tooling. It does not establish support in a published client. Session-enforced
clients require a later add-in stage; keep their gates intact.

In a separate disposable user database, export a query baseline and modify both
the database query and its source. Exercise all five policies. `block` must leave
the conflicting bytes intact and report `decision_required`. Verify each other
policy's action against the source and database, then run another clean operation.
Exercise single-object and category import/export results under a caller-owned
policy, plus MergeBuild start refusal and terminal completion separately.

Test a missing `modTestAssert` helper with `decline` and `block`. Neither may install
the helper without an authorized choice. The translation probe is
[`verify-caller-error-translation.py`](../tools/verify-caller-error-translation.py).

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
logging and trapping, and retention of caller policy until its explicit clear.

Run `ModeOnlyScopeProbe` for `snapshot`, `fault`, and `refused` after selecting mode
2. Run `ScopedSyncOwnershipCheck`, `ScopedSyncCancelCheck` for import and export,
and `ExportFormPolicyCheck` outside the runner's root in a disposable user database.
Require their `OK` result. Preserve failure artifacts before closing the owned host.

Native dialog dismissal, cancellation polling, async admission isolation, named
build output paths, export migration recovery, and compatibility sessions are
outside this contribution. Native object-close mechanics remain inherited.
