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
owned close, fresh reopen, and compilation check. Dispose of the whole host
after preserving results. Preserve unrelated Access processes.

Filter the runner to `modTestFileAccess`. These new procedures cover the writer:

- `TestLayoutSvgFileWriter` checks all 15 bytes of `<svg>` followed by omega,
  `</svg>`, and CRLF. Omega is `CE A9`, and the first byte is `3C`, not a BOM.
- `TestWriteFileNoBomSkipsUnchangedContent` checks readback and preserves a
  backdated modification time when content is unchanged.
- `TestWriteFileNoBomOverwritesExistingFile` replaces a longer file with omega
  and checks its four bytes, including the writer's appended CRLF.

The procedures create unique temporary files and clean up after trapped errors.
The [test source](../Version%20Control.accda.src/modules/Tests/FileIO/modTestFileAccess.bas)
contains the fixtures, so no separate data files are needed. A run with no
assertions or only `EMPTY` results does not qualify the writer.

For a before-fix control, build an isolated copy of the base with these tests
added and the rewind omitted. Run the same three filters. Record failures and
actual output bytes. Do not edit a loaded VBA project to construct this control.

## Check the callers

1. In a separate disposable user database, create a form with a label whose
   caption includes omega. Enable layout SVG export and export the form. Check
   that the `.svg` is nonempty, parses as XML, has no BOM, and includes the label.
2. In a disposable development checkout, use `Translation.ExportTemplate` and
   `Translation.ExportTranslations` with its local `Translation` folder. Save
   and restore any translation contribution or path settings. Check that the
   generated POT and PO files are nonempty UTF-8 without a BOM and contain their
   gettext headers and messages. Do not use a shared translation folder.
3. Run existing file-access regressions and the unfiltered suite on fresh
   disposable hosts. Compare unrelated failures with an equivalently built base.

This fix preserves bytes on new writes. It does not change the export format or
force unchanged forms to export again. To replace an existing empty SVG, export
that object again or run a full export with layout SVG enabled. Regenerate POT
and PO files through their existing translation commands.

The Windows [caller driver](check_utf8_callers.py) requires Access and pywin32.
It calls the installed add-in as a library and closes only its owned process.
Use a new user database for SVG and a disposable development checkout for
gettext. The [translation fixture](Fixtures/modules/vcs_test_utf8_translation.bas)
imports once into that idle development host. It remains through close and fresh
reopen validation. The driver restores the contribution setting exactly.

```text
python Testing/check_utf8_callers.py svg <new-user.accdb> <installed-addin.accda> <new-receipt.json>
python Testing/check_utf8_callers.py gettext <disposable-dev.accda> <installed-addin.accda> <new-receipt.json>
```

Keep the caller receipts and failure artifacts before disposing of each owned
host. Use the same rebuilt writer stage in the installed runner and tested host.
