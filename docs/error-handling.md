# Error handling

The add-in does not use a single top-level error handler. Instead, every
procedure that can fail opts into a structured inline pattern built on four
functions in `modErrorHandling`. This document explains the pattern and, more
importantly, how to read the log entries it produces — which are easy to
misinterpret.

---

## The pattern

```vba
Public Sub SomeOperation()
    ' Use inline error handling with debug mode check
    If DebugMode(True) Then On Error GoTo 0 Else On Error Resume Next

    ' ... operation code ...

    ' Catch and log errors inline
    CatchAny eelError, "Error description", ModuleName & ".SomeOperation", True, True

    ' For critical errors that should stop the operation
    If Operation.ErrorLevel = eelCritical Then GoTo CleanUp

CleanUp:
    ' Cleanup code
End Sub
```

| Function | Purpose |
|---|---|
| `DebugMode(True)` | Returns true if debug mode is enabled; also calls `LogUnhandledErrors` internally |
| `LogUnhandledErrors` | Call before any `On Error` directive to capture errors that would otherwise be silently discarded |
| `CatchAny()` | Log an error if one exists, optionally clearing it |
| `Catch()` | Check for specific error numbers |

Error levels come from `eErrorLevel` in `modConstants`: `eelNoError`,
`eelWarning` (logged to file), `eelError` (displayed and logged), and
`eelCritical` (cancels the current operation).

Under an MCP or external API call, `DebugMode` returns `False` and
`LogUnhandledErrors` never breaks into the debugger. See
[mcp-runvba.md](mcp-runvba.md) for why.

### Request admission exception

Request construction, identity claims, and queued refusals run before the caller
owns an operation. They must not send an error to the shared `Log`, because that
log can belong to another active request. `InitializeRequestErrors` captures
inherited `Err` diagnostics with `Debug.Print` before clearing them, then calls
`DebugMode(True)` with no pending error. The caller installs its local handler.
Parsing failures propagate to the owning dispatcher;
identity-store failures return an explicit refusal and release their mutex.
These helpers document that boundary in their source. Once a request owns the
operation, use the normal structured pattern above.

### Hosted test-run flag callback exception

The private `clsTestRunner.SetHostedTestRunFlag` callback also uses
`InitializeRequestErrors`. Acquisition can precede the hosted break guard, and
release can run after its root ends while another runner owns the active
operation. A dormant `False` call returns before any error handling. Inherited
errors stay local, so this callback cannot enter `Stop` or write to the other
runner's shared `Log`.

Callback failures return `False`. Failed acquisition becomes `RunError`, and
`RunSelected` returns before it changes error trapping or suppression. Failed
release retains ownership and the captured file-qualified target for a destructor
retry against the same project. This exception covers only that private ownership
callback. Operation code still uses the normal structured pattern above.

---

## Reading `LogUnhandledErrors` entries in a log

VBA's `On Error` statements silently clear the current `Err` object. To avoid
losing that information, `LogUnhandledErrors` is called *just before* an
`On Error` directive, capturing any leftover error before it gets wiped.
`DebugMode(True)` calls `LogUnhandledErrors` internally, so the same behavior
applies at the top of every function using the `DebugMode` pattern.

**The error did not originate where it was logged.** When you see:

```
ERROR: Unhandled error, likely before `On Error` directive
```

that entry means the exact origin is unknown. `LogUnhandledErrors` detected a
leftover error but has no information about which function raised it. The error
came from whatever code ran immediately *before* the `LogUnhandledErrors` call.

To find the real source, look at the surrounding log context — the operation in
progress, the objects being processed — then find the `LogUnhandledErrors` call
site in the source and read what executes before it.

Some call sites pass a `CallingFunction` parameter, which narrows the search to a
specific function (e.g. `Source: modBuild.Build.Unknown.LogUnhandledErrors`).
Even then, the error did not originate in that function:

```vba
Public Sub Build()
    ' ... earlier code calls helper functions ...
    SomeHelperFunction   ' <-- If this raises an error internally and doesn't handle it,
                         '     the error persists in the Err object after it returns.

    LogUnhandledErrors   ' <-- Catches the leftover error from SomeHelperFunction
    On Error Resume Next ' <-- Would have silently cleared it without the line above
    ' ... more code ...
End Sub
```

The actual source here is `SomeHelperFunction`, not `Build`.

---

## Where the logs are

Operation logs are named for the operation type, not the entry point:
`Export_*.log`, `Build_*.log`, and `Merge_*.log` — the last covering merge builds
**and** single-object imports, since `LoadSelected` / `ImportObject` begin an
`eotMerge` operation. All `logs/` directories are gitignored, so agent tools that
respect `.gitignore` will skip them; read them with shell commands. See
[testing-strategy.md](testing-strategy.md) for the full list of log locations.
