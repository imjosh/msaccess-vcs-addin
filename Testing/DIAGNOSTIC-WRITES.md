# Verify best-effort diagnostic writes

`MCPDebugLog` drops a failed diagnostic write. `TryAppendToFile` returns `False`
without logging the failure, changing operation error state, or retrying a locked
file. Successful writes create a UTF-8 BOM once and append CRLF-terminated lines.
The helper consumes an inbound `Err` without logging it and returns with `Err`
cleared. Capture any error details before calling it. It uses a local filesystem
object to avoid the shared getter's logging and retries.
This does not suppress other logged errors or introduce a decision-policy API.

1. Commit the source and rebuild it through the supported
   [RebuildAddIn workflow](../docs/agentic-rebuild.md). Use an isolated checkout
   in an Access trusted location. Keep the resulting closed `Version Control.accda`
   beside its matching source. Never patch the installed library.
2. Use a Windows Python environment with pywin32 and a compatible
   `msaccess_vcs_mcp` test client. These early API checks were qualified with the
   legacy 0.1.0 client and Access 16 x64. A newer client's version/session refusal
   is not a diagnostic-write regression. Do not disable a production gate.
3. Run the following commands from the checkout, with no competing native work.
   Replace `<evidence>` with a new directory outside this checkout. Each label
   must be unused. The driver owns only its newly created Access instances.

```text
python Testing/check_diagnostic_writes.py <evidence> focused-sync
python Testing/check_diagnostic_writes.py <evidence> focused-callback --transport callback
python Testing/check_diagnostic_writes.py <evidence> blocked --mode blocker
python Testing/check_diagnostic_writes.py <evidence> locked --mode locked
python Testing/check_diagnostic_writes.py <evidence> normal --mode normal
python Testing/check_diagnostic_writes.py <evidence> broad --broad
```

The focused run must execute all five `TestTryAppendToFile*` tests, with passing
assertions. The blocker and locked probes must report zero journal/error deltas,
unchanged operation level/status/mode, an active then released callback, preserved
original bytes, a failed registration after an active callback that releases
callbacks and consumes the parse error, and a successful public version query after the write attempt.
The normal probe must contain its marker with one BOM. All runs must compile
independently after fresh reopen, preserve readable components, and confirm both
owned process exits. An all-EMPTY run does not qualify.

`vcs_diagnostic_write_probe.bas` executes before any root starts, with the normal
interaction mode. Its callback URL is deliberately unused: registration and debug
writes do not send requests. No policy or session registration is needed.

The broad command excludes two upstream module-import tests that delete fixtures
from the executing project. Inspect the result keys to confirm both exclusions
apply. For an unfiltered lifecycle, combine this source with the disposable-host
fixture-retention change and follow that contribution's verification instructions.
Keep standalone and combined results separate. This change does not repair native
project deletion. Preserve failed hosts rather than repeating that trigger.

The driver archives the closed host before disposing its temporary checkout.
Receipts include exact source/binary/runner identities, test results, component
hashes, process identities, and reopen compilation. Keep failures and restoration
records outside the contribution. Restore the original installed source through
the supported rebuild, along with temporary install-source settings.

For a safe base comparison, call the unchanged public diagnostic entry in a fresh
owned base host with a file named `logs` in its source folder. Bound execution and
recover only that host's diagnostic error dialogs. Expect logged errors or prompts
from the old append path. The new helper tests cannot compile on the base because
`TryAppendToFile` does not exist there. Do not import them into a loaded library.
