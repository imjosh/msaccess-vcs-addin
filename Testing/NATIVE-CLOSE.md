# Native close regression checks

Build the complete source tree using [the supported rebuild](../docs/agentic-rebuild.md).
Use the [disposable suite lifecycle](../docs/agent-test-runs.md) for focused filters
`clsTestOperationLifecycle,modTestDatabase,modTestErrorHandling,modTestScopedSync` and
an unfiltered run. Keep the installed runner and disposable development host separate.

The following parameterized helpers are additional public-driver entry points. Load
the newly compiled development add-in as a library in a disposable user database;
never run them inside the suite's active root or on the installed library as a database.

- `ScopedSyncNativeCloseCheck(False)` and `(True)` cover category import/export
  cancellation/error seams, decision-required precedence and successful controls.
  Each must return `OK`; these are fault-seam checks, not real UI evidence.
- `IndexedSingleObjectCloseCheck(sourcePath, exporting, nativeCancel)` exercises
  indexed legacy single-object import/export. First seed a saved form named
  `frmNativeCloseProbe`, export its source/index successfully, and open it in design
  view with an unsaved caption. Run all four direction/outcome combinations. Compare
  source/index SHA-256 and modification time immediately around each call. Require
  unchanged files and unsaved design, zero transfers, the owned log and cancellation
  only for the Cancel seam. Discard changes only after those checks.
- `NativeCloseLiveCall(caseName)` supports `single-import`, `single-export`,
  `category-import` and `category-export` in the same disposable fixture. It arms no
  seam. Each call must reach Access's actual save/discard prompt on an unsaved form.
  End the capture phase after recording PID, process creation identity, HWND, class,
  title, text and all buttons. In a separate bounded Win32 phase, revalidate that
  identity and the literal Cancel control before choosing it once. An unidentified
  or second prompt stops the run. Require `cancelled: true`, no object/source/index
  mutation, an owned log and released operation state. Keep real UI receipts separate.

The live helper deliberately uses the user-interface source for categories to reach
Access's save prompt; single-object calls use an explicit request-local block policy so their
expected logged refusal is acknowledged without another add-in error prompt. This does not
claim that decision policies dismiss native Access prompts.

Keep an unlocked interactive desktop awake for each bounded live run. Observe
progress and modal windows every 20 seconds, stop on unexpected dialogs, retain
failed fixtures and restore temporary settings. Manual ribbon checks remain separate.
