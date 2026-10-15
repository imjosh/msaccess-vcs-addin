# Verify disposable development test hosts

Use a Windows Access installation with trusted VBA project access and the matching
source-built add-in installed through the [supported rebuild](../docs/agentic-rebuild.md).
The Python driver requires `pywin32` and the MCP client's Python package. This
stage was tested with MCP `0.1.0` against add-in `5.1.0`. Record the actual client
commit and binary hashes with each run. Later clients may require a different API.

Keep the evidence directory outside this checkout. The driver creates a detached
Git worktree for each run, copies the closed build output and matching source
context, and opens only that disposable database. Start with no Access processes
running. The driver refuses otherwise and never terminates another process.
Older command-bar tests read menu fixtures beside the checkout. Each owned host
container includes its own matching sibling copy. The driver archives that copy
with the host and removes both after confirmed closure. Check for unexpected
SKIP assertions as well as EMPTY procedures before accepting an unfiltered run.

The module-import tests leave their fixture modules in the executing project, so
they refuse to import unless the host is explicitly designated disposable. The
designation is the database property `VCSDisposableTestHost`, whose text value is the
full path of the database carrying it. A missing property, a non-text value, or a path
naming another database records a failing assertion that names the reason, and the
test exits without importing. Repository context alone never qualifies a host. The
driver creates the designation in its owned host, before any test runs, by importing
`Testing/Fixtures/modules/vcs_disposable_host_probe.bas` and calling its
`DesignateDisposableHost`. It records the stored value as `designation` in the receipt.
Copying a designated database to another path invalidates the designation.

From this repository, run these commands with the MCP client's Python environment:

```powershell
python Testing/check_disposable_suite.py focused './Version Control.accda' HEAD --evidence ../suite-evidence --filters 'TestModuleImport_IndexesEachFileOnSharedInstance,TestModuleImportFast_IndexesEachFileOnSharedInstance,modTestClassInvocation,clsTestTargetLocalInvocation'
python Testing/check_disposable_suite.py full-1 './Version Control.accda' HEAD --evidence ../suite-evidence
```

Rebuild from the same source again, then run a second unfiltered lifecycle with a
new label such as `full-2`. Wait for each command to finish before starting another.
The driver uses the public MCP test tool. It closes its owned process, confirms
exit through the original process handle, reopens the database, reads retained
fixtures, and checks native compilation and `Application.IsCompiled` independently.
It clears session overrides and records host state and installed-library code hashes.

The focused run must execute five procedures and 14 assertions. Both import tests
must record two source-path index assertions. All four fixture modules must remain
readable after reopen, without duplicate names or replacement by the other import
test. Class tests must record real assertions. An all-EMPTY result fails qualification.
Inspect the JSON receipts and repository-dependent assertions, not only command exit.

Run these controls on separate fresh hosts:

```powershell
python Testing/check_disposable_suite.py assertion './Version Control.accda' HEAD --evidence ../suite-evidence --filters TestModuleImport_IndexesEachFileOnSharedInstance --fault assertion
python Testing/check_disposable_suite.py import-refusal './Version Control.accda' HEAD --evidence ../suite-evidence --filters TestModuleImport_IndexesEachFileOnSharedInstance --fault import
python Testing/check_disposable_suite.py class-error './Version Control.accda' HEAD --evidence ../suite-evidence --filters clsTestTargetLocalInvocation --fault class-error
python Testing/check_disposable_suite.py legacy-helper './Version Control.accda' HEAD --evidence ../suite-evidence --filters 'modTestClassInvocation,clsTestTargetLocalInvocation' --fault legacy-invoker
```

The assertion control changes a fixture's module name. The import control supplies
unresolved Git conflict markers, which the existing importer refuses before VBA
import. Both must report failures, restore the prior error level, and complete the
owned close, reopen, compilation, and disposal lifecycle. Neither uses malformed VBA.

The class-error control installs a helper in the idle disposable current project
that calls a nonexistent member. The helper catches error 438 and returns scalar
error data. The external runner must record ERROR and release the class instance.
The class writes `invocation-teardown.txt` in the disposable checkout. Require
one `terminated` line for the error control. This marker survives VBA global-state
resets after the external run, so a reset counter cannot hide successful teardown.
The legacy-helper control seeds the earlier Sub helper. Scan must upgrade it to
the scalar-returning Function, and all three class-invocation procedures must pass.
These controls never edit the loaded library.

Add `--sync` to a focused command to verify the synchronous `RunFilteredTests`
JSON API through owned COM instead of MCP callbacks. The driver sets the session
filter before dispatch, and API source forces headless execution. Require the same assertions
and reopen checks. This switch does not change the unfiltered MCP qualification.

Each receipt records errors separately from test results. The driver preserves a
verified host archive and copies logs before disposal. A failed reopen remains raw
for diagnosis. A zero exit code alone does not mean that expected test failures
passed. Check the expected outcome, assertion counts, fixture readability,
`reopen_is_compiled`, both process exits, and `disposed` in each receipt.

For a safe base comparison, build the base with only the proposed class regressions
and state-probe fixture added. Run the three class procedures. The helper-dependent
tests should fail because the base does not generate `TestClassInvoke`. Avoid
repeating the base's executing-project module deletion to obtain a red result.
Document that lifecycle difference by source inspection and candidate retention.

This lifecycle qualifies add-in development hosts. It does not establish a repair
of production module deletion or qualify a newer MCP compatibility contract.
