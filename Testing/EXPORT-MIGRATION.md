# Disposable export migration fixture

`modTestExportMigration.ExportMigrationCheck(case)` exercises the public export
entry points from a fresh, empty Access `.accdb` whose filename begins with
`UP008_`. The development `Version Control.accda` is loaded as a library; the
fixture creates forms and a query only in the disposable host. Use a supported
source build of the committed add-in under test, and record its commit, source
tree, and binary SHA-256 before opening Access.

The case arguments are `blocked`, `scan_cancel`, `partial_cancel`,
`success`, `global`, `filtered`, `skip`, `warning_complete`, and
`warning_incomplete`. Call the function through
`Application.Run` with the loaded library name and one case argument. It returns
JSON with the initial public outcome, persisted index snapshots, retry outcome,
and failed assertions. Each call creates its own host, makes a baseline full
export, changes options to require a Forms SVG migration, runs the case, clears
its conflict or cancellation trigger, and uses the public fast export for the
retry. A passing return has `success: true` and an empty `failures` array.

For independent evidence, preserve the `.accdb` and its `.src` folder until
the owned Access process has closed. Record every `.svg` path, nonzero size,
SHA-256, and modification time after the initial attempt and retry. Retain the
`before`, `afterAttempt`, and `afterRetry` index snapshots from the JSON and
compare category hashes, component `ExportDate` values, and whole-export dates.
The persisted index is the authority; a log message alone is not completion.
The driver must record the selected and completed object counts for each
category and identify which categories were complete before retry. A category
that was already complete must not be exported again without a declared reason.
The warning cases arm a test seam only in a `UP008_` disposable host. The
complete case retains the SVG despite a warning; the incomplete case removes
one SVG before completion accounting. Both use the later public fast export.

Run each case in a new directory. Do not run the fixture in a user database,
the development library itself, or a sample database. Watch for Access modal
windows and bound the run. Close only the Access instance the driver opened,
then reopen the retained disposable host to compile and inspect its artifacts.
Run the relocated `TestLayoutSvgFileWriter` from `modTestFileAccess` separately.
