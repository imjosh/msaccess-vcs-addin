# Verify BOM-free UTF-8 file output

`WriteFileNoBom` writes the bytes copied into a binary ADODB stream. `CopyTo`
leaves that stream at EOF. Rewind it before `Read` to preserve the encoded text.
The existing SVG, POT, and PO callers use this writer.

## Run the regression

Build the proposed source in an isolated checkout using the
[source rebuild procedure](../docs/agentic-rebuild.md). Run the installed-library
test runner against a fresh disposable development `.accda` with its matching
source, repository, and fixtures. Follow the
[test-run guidance](../docs/agent-test-runs.md). Keep import fixtures through an
owned close, fresh reopen, and compilation check. Preserve results before you
dispose of the whole host.

Filter the runner to `modTestFileAccess`. These procedures cover the writer:

- `TestLayoutSvgFileWriter` checks all 15 bytes of `<svg>` followed by omega,
  `</svg>`, and CRLF. Omega is `CE A9`, and the first byte is `3C`, not a BOM.
- `TestWriteFileNoBomSkipsUnchangedContent` checks readback and preserves a
  backdated modification time when content is unchanged.
- `TestWriteFileNoBomOverwritesExistingFile` replaces a longer file with omega
  and checks its four bytes, including the writer's appended CRLF.

The procedures write into unique temporary folders and delete them after
trapped errors. The
[test source](../Version%20Control.accda.src/modules/Tests/FileIO/modTestFileAccess.bas)
contains the fixtures, so no separate data files are needed. A run with no
assertions or only `EMPTY` results does not qualify the writer.

For a before-fix control, build an isolated copy of the base with these tests
added and the rewind omitted. Run the same three procedures. Record failures
and actual output bytes. Do not edit a loaded VBA project to construct this control.

For a broader check on standalone UP-001, exclude the two existing module-import
tests that remove fixtures from their executing project. Fixture-lifecycle repair
and unfiltered verification belong to the separate UP-002 contribution. Use this
explicit filter and compare results with an equivalently built base:

```text
msaccess-vcs run-tests <fresh-disposable-dev.accda> --filter=-TestModuleImport_IndexesEachFileOnSharedInstance,-TestModuleImportFast_IndexesEachFileOnSharedInstance
```

Describe that run as filtered. Verify that both excluded procedures are absent
from the actual result map, and preserve the fresh-reopen compilation result.

## Check the callers

The Windows [caller driver](check_utf8_callers.py) requires Access and pywin32.
Run it without `-O`, `-OO`, or `PYTHONOPTIMIZE`. It calls the installed add-in as
a library and closes only its owned processes. Use the same rebuilt writer stage
in the installed library and tested development host. Use a new receipt path for
each run.

For SVG, choose a new disposable user database path with no existing `.src`
folder beside it:

```text
python Testing/check_utf8_callers.py svg <new-user.accdb> <installed-addin.accda> <new-receipt.json>
```

The driver creates a form with an omega label, enables layout SVG export, and
exports it. It checks that the SVG is nonempty UTF-8 without a BOM, parses as
XML, and contains omega.

For gettext, prepare a separate disposable development checkout. Build and close
its development `.accda` at the checkout root beside `Translation`. Prepare
`tblLanguages` with `en_TEST` and `pt_BR`. Keep all `Translation/*.pot` and
`Translation/*.po` files tracked at `HEAD` and clean in both the index and working
tree, including `en_TEST.po` and `pt_BR.po`. Remove untracked or ignored POT and
PO files from this disposable checkout before the run.

Explicitly designate that database as disposable with a UTF-8 JSON marker. Run
this PowerShell example from the repository root after replacing the paths:

```powershell
$probeDatabase = (Resolve-Path 'C:/path/to/disposable-checkout/Version Control.accda').Path
$probeMarker = Join-Path (Split-Path $probeDatabase) 'utf8-disposable-host.json'
$probeDesignation = @{
	purpose = 'utf8-gettext-caller'
	database = $probeDatabase
} | ConvertTo-Json
[IO.File]::WriteAllText($probeMarker, $probeDesignation, [Text.UTF8Encoding]::new($false))
python Testing/check_utf8_callers.py gettext $probeDatabase '<installed-addin.accda>' '<new-receipt.json>' --disposable-host-marker $probeMarker
```

The driver imports the [translation fixture](Fixtures/modules/vcs_test_utf8_translation.bas)
once into the idle development host and retains it through close, fresh reopen,
and compilation. It restores the translation contribution setting exactly.
Before the probe, it deletes the checkout's tracked POT and PO files and writes a
placeholder `MSAccessVCS.pot` for path discovery. It requires fresh, nonempty,
BOM-free UTF-8 `MSAccessVCS.pot`, `en_TEST.po`, and `pt_BR.po`, each with gettext
headers. Missing either expected language fails the check.

Keep the receipts and failure artifacts before you dispose of each owned host.
After preserving the gettext outputs, run `git restore --source=HEAD --worktree -- Translation`
in the disposable checkout to restore its tracked translation files.

## Regenerate existing output

The Forms and Reports exporter revisions increase from 3 to 4 and from 2 to 3.
On the next completed export that includes those categories, unchanged forms
and reports become eligible for re-export. With layout SVG enabled and no source
conflicts, that export regenerates their missing or empty SVG sidecars.
The inherited interrupted-export accounting can record a new category revision
before every object finishes. An object or full export remains a manual recovery
option for that case. Regenerate POT and PO files through the existing translation
commands.
