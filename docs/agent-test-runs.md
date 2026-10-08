# Running the add-in's own tests

How an agent runs this repository's add-in suite. Use a **fresh disposable
development copy** of `Version Control.accda` with matching source/repository
context, and drive the installed-library runner through MCP.

The module-import tests retain distinct normal/batch fixtures in the disposable
host and restore their error state on success or trapped failure. Build this
exported source through the supported rebuild before running the lifecycle below.
[A43](../../issues/A43-addin-isolate-module-import-test-fixtures.md) records native
qualification; existing A42 cleanup-enabled reports remain historical evidence.

For writing a test, see [.cursor/rules/testing.mdc](../.cursor/rules/testing.mdc).
For the layers, the round-trip harness, and where results land, see
[testing-strategy.md](testing-strategy.md).

## Prepare and dispose of the entire development host

Normal developer use loads the installed add-in to operate on a separate user
database. Add-in development tests execute the add-in's own code and can modify
the project containing their test procedures. Whole-host disposal applies to this
repository's suite; it is not a requirement to discard developers' user databases.

1. Prepare an isolated checkout or equivalent repository snapshot from the current
   working source. Include the matching `.src`, fixtures, documentation and Git
   context. Resolve export/source paths inside that snapshot so tests cannot mutate
   the primary checkout or silently take missing-repository guards.
2. Use the supported source rebuild to obtain the version under test. Either build
   in the isolated checkout or copy the closed, compiled development `.accda` and
   its matching context there. Record source/binary hashes and the installed runner
   identity; never copy an open database or patch the installed library.
3. Run focused checks and the original unfiltered suite through MCP. Start each
   independent run from a fresh disposable host. Retain alpha/beta modules through
   validation; normal and batch fixtures need distinct module names/source paths
   so both tests coexist without deletion or replacement of retained fixtures.
   Preserve both indexing assertions on one shared module instance per test.
   Restore transient options/error state needed by later tests.
4. Gracefully close the identity-confirmed owned host, confirm original-handle
   process exit, then freshly reopen that same disposable database. Read every
   retained import fixture and independently check compilation/IsCompiled. Repeat
   the full lifecycle from a fresh supported source build for qualification.
5. Copy results and logs outside the temporary checkout before disposal. Preserve
   failing binaries and diagnostics too. After confirmed owned closure/exit, remove
   the disposable database, sidecars and checkout. Verify disposal on success,
   assertion failure and import failure; retain files when closure is uncertain.
   Restore temporary trust/install-source settings and caller/session state, and
   leave unrelated Access processes intact.

The installed library remains the runner and assertion receiver. Test procedures
and import assertions execute in the disposable current project. It may contain
fixture modules until disposal; the primary development database remains free of
suite mutations. A qualified disposable suite does not diagnose A42's low-level
corruption or qualify production module deletion in a separate user database.

The examples below assume the disposable checkout has already been prepared:

```
vcs_run_tests("C:\scratch\addin-suite\msaccess-vcs-addin\Version Control.accda", "clsTestInstall")
```

MCP progress is best-effort in Cursor. For live per-test output, keep this
CLI command in the foreground:

```text
msaccess-vcs run-tests "C:\scratch\addin-suite\msaccess-vcs-addin\Version Control.accda" --filter clsTestInstall
```

The stream is pytest-style: dots for fast passes, a named line after a test
that took ≥ 1s, and full FAIL/ERROR/EMPTY lines. Assertion detail stays
in the TestRun log and in the `vcs_run_tests` MCP result. The CLI prints a
compact JSON summary (no `tests` map) and a last line such as
`Tests passed. 12 subs, 40 assertions in 1.48s`.

Headless means no add-in UI (no web runner, no console form, silent dialogs),
not a hidden Access window. The host instance stays visible so a dialog or a
VBA break is on screen.

## Filtering a run

`VCS.RunTests` takes an optional `ParamArray` of filters and resolves each one in
this priority order: module name, suite or `@Folder` value (exact or final
segment), procedure or `Module.Procedure` key, then `'@Tag`. Prefix with `-` to
exclude; inclusions combine with OR and exclusions with AND. A list containing
only exclusions starts from all tests.

```vba
?VCS.RunTests                              ' Run everything
?VCS.RunTests("modTestEncoding")           ' Run one module
?VCS.RunTests("SQL", "-slow")              ' Run one suite without slow tests
?VCS.RunTests("TestParseJoinExpression")   ' Run one procedure
?VCS.RunTestsHeadless("-slow")             ' Unattended; always writes JUnit
?VCS.RunRoundtripTests                     ' Run the object round-trip corpus
```

`database_path` is the disposable development copy — the `.accda` beside
`Version Control.accda.src` in the isolated repository snapshot. The runner scans
`CurrentVBProject`, so whichever
database hosts the run is the one whose tests are found: point a run at a user
database and you get that database's tests, reported as a clean pass because
nothing you were looking for was there to fail.

**Never host a run on the installed add-in.** The copy under `%AppData%\MSAccessVCS`
exists to be loaded as an add-in and nothing else; it is not to be opened as a
database. It also has no source tree beside it — its export folder holds only
`logs`, `mcp`, `tables`, and `test-results`, no `modules` and no `forms` — and a
good part of the suite reads real exported files or asks about the enclosing Git
repository. A run hosted there produces failures that say nothing about the code
under test, which is its own reason not to believe one.

`vcs_run_tests` opens the development copy when no instance has it open, `.accda`
included. The server binds a file moniker first, which Access only honours for a
database extension, and falls back to `OpenCurrentDatabase` when that bind fails.
`AutoRun` stands down when `Application.UserControl` is False, so a COM client can
open it without the install message box closing the instance or the installer form
stranding it.

## Nothing in `Testing\` hosts this

`Testing\Testing.accdb` and `Testing\Testing.accdb.src` are a sample database
used as a build and export integration fixture. They are not where the add-in's
tests live and not a host for running them — a run pointed there searches that
sample database for tests. `Testing\Fixtures\` is the object round-trip corpus
driven by `VCS.RunRoundtripTests`, which is a different entry point again. See
[Testing/AGENTS.md](../Testing/AGENTS.md).

Agents have historically arrived at `Testing.accdb` while looking for a database
to host a call from, after `VCS_API_REFUSED` made hosting look like the problem.
If a refusal says a call was dispatched to the installed add-in and arrived back
where it started, that is a defect in `modAPI` and no choice of host will move it
— see [Where that refusal came from](#where-that-refusal-came-from).

## Run through the MCP server, not the add-in's window

The runner singleton that records assertions lives in whichever project received
the `RunTests` call, while `modTestAssert.TestAssert` always routes through
`Application.Run` to the *installed* add-in path. Invoking `VCS.RunTests` from
inside a development copy puts those two in different projects: every assertion
is silently discarded, every test reports `EMPTY`, and the run looks clean while
proving nothing.

**Treat an all-`EMPTY` result as a broken harness, not a pass.**

## Guards that skip instead of failing

`modTestRepoDocs` and `modTestAgentDocs` read the repository working tree, which
they locate from `CodeProject.Path`. Where there is no `AGENTS.md` beside that path,
`RepoIsAvailable` returns False and every check in both modules reports a passing
note instead. That guard is there for the end user whose install has no checkout,
not as a suite configuration: in the disposable development checkout these checks
must remain live. A database-only copy can make them report skips or passing notes;
inspect every guard and qualify the matching repository context explicitly.
Line budgets from [agent-docs-maintenance.md](agent-docs-maintenance.md) are worth
confirming directly either way, since a count is cheaper than a run:

```powershell
$root = Get-Content 'AGENTS.md' -Raw
$routing = [regex]::Match(
    $root, '(?ms)^## Where to read next\r?\n.*?(?=^## )'
).Value
"content chars: {0}/6000" -f (($root.Replace($routing, '') -replace '\r?\n').Length)
"routing chars: {0}/2400" -f (($routing -replace '\r?\n').Length)
"routing rows: {0}/20" -f ([regex]::Matches($routing, '(?m)^\| ').Count - 1)
Get-ChildItem '.cursor\rules\*.mdc' |
    ForEach-Object { "{0,-24} {1}/120 lines" -f $_.Name, (Get-Content $_).Count }
```

## Reaching an already-open instance

`vcs_run_vba` is the exception to the tool opening files for you. It attaches
through the Running Object Table without opening anything, so it reports
`Cannot find Access instance ... may have been closed` unless the file is already
open — for a `.accdb` as much as a `.accda`. Open it first, handing ownership to
the desktop so the instance outlives the launching script:

```powershell
$app = New-Object -ComObject Access.Application
$app.OpenCurrentDatabase("C:\path\to\msaccess-vcs-addin\Version Control.accda")
$app.Visible = $true
$app.UserControl = $true   # after opening, so AutoRun still sees automation
```

A rebuild ends with no Access process running, so reopen it before the next
iteration. See [agentic-rebuild.md](agentic-rebuild.md).

## Where that refusal came from

`modAPI.API` redirects to the installed add-in whenever the current database and
the running code are the same file, which is what carries a call made from the
development copy across to the library. Until August 2026 it did that without
checking where the call would land, so a call whose redirect target *was* the
running file re-entered the entry point while the outer call still held its
`Static IsRunning`, and came back `VCS_API_REFUSED`. The message named
`vcs_run_vba` nesting as the cause, which was wrong, and pointed at the choice of
host — which is how agents ended up hunting for some other database to call from.
`RedirectTargetIsSelf` now redirects only when the target is a different file, and
a self-dispatched refusal says so in its own words.

The lesson worth keeping: a refusal that says a call was dispatched to the
installed add-in and arrived back where it started is a defect in `modAPI`. No
choice of host database fixes one, and looking for a different host is what leads
into the two mistakes above.
