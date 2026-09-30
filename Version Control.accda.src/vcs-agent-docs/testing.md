# Automated Testing

The add-in includes a test runner that discovers and executes tests in whatever
database is open. Tests assert with `TestAssert`, a drop-in replacement for
`Debug.Assert`. Nothing in the project needs a compile-time reference to the add-in.

## Setup

Run `VCS.InstallTestAssertModule` in the Immediate Window to inject `modTestAssert`.
`VCS.MigrateDebugAssert` converts existing `Debug.Assert` calls in bulk.

## Writing tests

A test is a parameterless `Public Sub` in a test module:

```vba
Option Compare Database
Option Explicit
Option Private Module
'@Folder("Tests")

Public Sub TestDoubleInput()
    TestAssert MyFunction(42) = 84, "MyFunction should double input"
End Sub
```

The optional second `TestAssert` argument names the assertion, which matters
inside loops and shared helpers.

A module counts as a test module if it carries `'@Folder("...Tests...")` (in
projects using `@Folder` at all) or its name contains `Test`. Within it, only
parameterless `Public Sub` procedures are registered, and only if the module has a
`TestAssert` call. To exclude a helper, make it `Private` or give it a parameter.

**Class modules** work the same way and add per-test setup and teardown: every test
method gets a fresh instance, so `Class_Initialize` runs before it and
`Class_Terminate` after. Use parameterless `Public Sub` or `Public Function`. Name
test modules `modTest*` or `clsTest*`, and mark standard ones
`Option Private Module`.

An unhandled error in a standard-module test stops in the VBA debugger and the run
never finishes (the runner calls it via `Application.Run`, which ignores Error
Trapping). Class tests report ERROR instead; use one, or handle the error.

## Prompts during a run

A run is unattended: a `MsgBox`, `InputBox`, or modal form in code a test reaches
stalls it until someone clicks. `modTestAssert.TestRunActive` is true for the whole
run, so guard anything that waits for a person:

```vba
If Not TestRunActive Then MsgBox "Import complete.", vbInformation
```

Prefer guarding a prompt whose absence changes nothing; where a prompt decides
something, branch to the answer an unattended run should get. A `modTestAssert`
installed before this feature will not have the flag.

## Running tests

Run `?VCS.RunTests` from the Immediate Window, or **Tools > Run Tests** on the
ribbon; after a run it can re-run only the failures.

`RunTests` takes an optional `ParamArray` of filters. Each argument resolves in
priority order: exact module name, then suite or `@Folder` value (matching the full
path or its final segment, so `"SQL"` matches `"Tests.SQL"`), then procedure name or
full `Module.Procedure` key, then tag. Prefix an argument with `-` to exclude it.
Inclusions combine with OR, exclusions with AND, and a list containing only
exclusions starts from all tests.

```vba
?VCS.RunTests("-slow")                   ' Everything except slow-tagged tests
?VCS.RunTests("Reporting", "-slow")      ' The Reporting suite, skipping slow tests
?VCS.RunTests("TestInvoiceTotal")        ' One specific procedure
```

`RunTests` returns a JSON summary with per-test status, assertion detail, and tags.

## Tagging

`'@Tag("name")` annotations categorize tests (case-insensitive). A module-level tag
sits in the first ~30 lines, before any procedure, and covers every test in it. A
procedure-level tag sits at the top of the procedure body, before any executable
line including `Dim`. The two sets merge.

## Global suite hooks

`modTestAssert` may define two optional once-per-run hooks:

```vba
Public Sub GlobalTestSetup()    ' Before the first test, when at least one is selected
Public Sub GlobalTestTeardown() ' After all tests; the results JSON already exists
```

`VCS.InstallTestAssertModule` writes empty stubs; absent hooks are skipped
silently, and neither runs when no tests are selected. An error inside a hook is
non-fatal: it goes to the console and teardown still executes. Per-test
`Class_Initialize` and `Class_Terminate` nest inside these hooks.

## Where results land

| File | Contents |
|------|----------|
| `logs/TestResults_<timestamp>.json` | Per-run results with per-assertion detail |
| `logs/TestRun_<timestamp>.log` | Full console output including timing |
| `test-results/test-state.json` | Merged current state; a partial run updates only the tests it executed and flags the rest `stale` |
| `test-results/test-results.xml` | JUnit XML projection of the state file |
| `test-results/test-results.html` | Self-contained HTML dashboard |

Both folders are gitignored, so search tools miss them; see `troubleshooting.md`.
