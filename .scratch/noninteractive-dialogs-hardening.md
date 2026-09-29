# Spec: Harden noninteractive decision policies

Follow-up to commit `cd22aa3` ("feat(dialogs): keep automated tests and merges
off add-in prompts"). Derived from a two-axis code review (standards and spec)
of that commit.

## Problem Statement

Agents, CI scripts, and other automated callers run merges and test suites
through the add-in with no one at the screen. Commit `cd22aa3` let them pass a
decision policy so add-in prompts don't block them. In practice the guarantee
is weaker than documented:

- A merge can still leave the main VCS window visible. One raw message box can
  still open during report import.
- If a test run hits a runtime error, the operation stays "running" and the
  logging state isn't restored. Silent mode can also stay on after the run.
  The next automated call then fails with "Another Operation Already Running",
  or it behaves differently.
- A merge that stops before it starts — for example because no source folder
  is known, merge isn't available, or the policy name is invalid — gives the
  caller no signal. The `decision_required` record is thrown away.
- Policy scopes can leak between operations. An operation's finish can pop a
  scope it didn't open, and a policy applied while another operation is
  running changes that operation's behaviour. A test module that switches to
  silent mode can quietly turn a "block" run into one that approves defaults.
- `prefer_database` and `skip` are documented as different, but they behave
  the same. `prefer_source` ignores the action each conflict asks for, such as
  deleting an orphaned object.
- A decision containing control characters produces invalid JSON.
- Almost none of this is tested. Current tests cover only the pure mapping
  from policy name to answer.

For an automated caller, each of these either hangs a run behind a hidden
dialog or reports an outcome that isn't true. That is what the feature was
meant to prevent.

## Solution

Noninteractive operations keep their promise: while a decision policy is in
effect, no add-in dialog or window appears. Every prompt is either answered by
the policy or recorded as `decision_required`, and the caller always learns
the outcome, even when the operation never starts or fails part-way through.

A policy applies to exactly one operation. It is restored on every exit path
and can't be weakened or removed by code the caller didn't write. Each policy
name has one distinct, documented meaning. The interaction-scope behaviour
lives in one small object that tests can drive directly, plus a few wiring
tests that prove the real prompt sites use it.

## User Stories

1. As an MCP agent, I want `MergeBuild` with a policy to never show the VCS
   main window, so that a hidden or visible form never steals focus or blocks
   the Access UI I'm automating.
2. As an MCP agent, I want every add-in message box reachable during a merge
   to go through the policy, so that no raw `MsgBox` can hang my session.
3. As an MCP agent, I want a merge that can't start to return a structured
   result immediately, so that I don't wait for a completion callback that
   will never come.
4. As an MCP agent, I want a missing source folder during a noninteractive
   merge reported as `decision_required` with the folder-picker detail, so
   that I know to set the export folder and retry.
5. As an MCP agent, I want "merge not available" reported as a distinct error
   pattern, so that I can run a full build instead of retrying the merge.
6. As an MCP agent, I want an unknown policy name rejected the same way by
   `MergeBuild` and `RunFilteredTests`, so that one error-handling path covers
   both.
7. As an MCP agent, I want the list of valid policy names in the error
   message, so that I can correct the call without reading source.
8. As an MCP agent, I want a test run that hits a runtime error to still
   finish its operation, so that my next call isn't refused with "Another
   Operation Already Running".
9. As an MCP agent, I want a test run's error JSON to keep
   `decision_required` and the decisions list when a prompt was blocked
   before the error, so that I see the real reason the run stopped.
10. As an MCP agent, I want the interaction mode and logging state restored
    after a failed test run, so that later operations behave the same as
    before.
11. As an MCP agent, I want the policy I pass to apply only to my operation,
    so that a policy can't leak into, or be removed by, another operation.
12. As an MCP agent, I want a policy request refused cleanly when another
    operation is running, so that I don't change that operation's behaviour.
13. As an MCP agent using `block`, I want test code that asks for silent mode
    not to weaken my policy, so that no prompt is approved without my
    consent.
14. As an MCP agent, I want every decision in the JSON to be valid JSON
    whatever characters the prompt text contains, so that my parser never
    fails on the result.
15. As an MCP agent, I want `prefer_source` to carry out each conflict's own
    action (overwrite or delete), so that "source wins" means the same thing
    as the existing agent default.
16. As an MCP agent, I want `prefer_database` and `skip` documented with their
    actual shared effect, so that I don't choose one expecting a behaviour it
    doesn't have.
17. As a CI script author, I want the final outcome of a noninteractive merge
    reported through the completion callback with the decisions list, so that
    one mechanism tells me success, failure, cancellation, or
    `decision_required`.
18. As a CI script author, I want `decision_required` never reported as
    success or as a plain cancellation, so that my pipeline fails loudly when
    a human decision is needed.
19. As a CI script author, I want a policy set for a session to be removed
    only when I clear it, not when some operation finishes, so that a
    long-running session keeps the policy I chose.
20. As a CI script author, I want clearing a session policy that is already
    gone to be harmless, so that my cleanup code is safe to run twice.
21. As an interactive user, I want merges, builds, and test runs started from
    the ribbon to behave exactly as before, so that the hardening doesn't
    change my workflow.
22. As an interactive user, I want closing a silent or noninteractive
    operation's window to cancel without a second confirmation, and this
    documented, so that the behaviour isn't surprising.
23. As an interactive user, I want an automated run that failed to leave
    Access in a normal state, so that I'm not stuck with a hidden form or a
    locked operation.
24. As an add-in maintainer, I want the interaction-scope rules in one object
    with a small interface, so that I can reason about nesting and
    restoration without reading string-parsing code.
25. As an add-in maintainer, I want scopes released by an owner token rather
    than "pop the top", so that ownership mistakes fail safely instead of
    removing someone else's scope.
26. As an add-in maintainer, I want a scope to save and restore the
    operation's error level along with mode and policy, so that a nested
    blocked prompt doesn't poison the outer operation.
27. As an add-in maintainer, I want the scope's mode and policy to survive the
    timer-based Stage and Restore cycle a merge uses, so that the second half
    of a merge still runs under the caller's policy.
28. As an add-in maintainer, I want one helper deciding whether the main form
    may be shown, so that new code paths can't forget the noninteractive
    check.
29. As an add-in maintainer, I want the new code to follow the repo's
    error-handling pattern (`DebugMode(True)`, `LogUnhandledErrors`,
    `CatchAny`), so that errors in dialog handling are logged and debug mode
    still breaks on them.
30. As an add-in maintainer, I want decision JSON built with the existing
    JSON encoder, so that there's one escaping implementation.
31. As an add-in maintainer, I want dialog enumeration and dismissal kept
    apart from policy decisions, so that each module changes for one reason.
32. As an add-in maintainer, I want the noninteractive checks to use one
    predicate everywhere, so that a later mode change isn't missed at some
    call sites.
33. As a translator, I want every new user-facing log label wrapped in
    `T()`, so that the noninteractive log output can be localized.
34. As a test author, I want to create an interaction scope in a test without
    touching the live operation, so that I can test nesting, restoration, and
    outcomes in isolation.
35. As a test author, I want wiring tests that call the real `MsgBox2` and
    conflict resolver under a scope and then restore everything, so that I
    can prove the prompt sites honour the policy without opening a dialog.
36. As a test author, I want tests for `prefer_database` and `skip` answering
    a generic confirmation, so that the "conflict policies don't approve
    unrelated prompts" rule is pinned for every policy.
37. As a reviewer, I want `docs/noninteractive-dialogs.md` to match the code
    exactly — policies, error patterns, silent-mode exceptions, and prevented
    dialogs — so that the doc can be trusted as the contract.

## Implementation Decisions

**Interaction scope object (new, primary seam)**

- A new class owns everything the scope stack currently stores as strings
  joined with vertical tabs: interaction mode, decision policy, blocked flag,
  decision journal, and the operation error level at the time the scope
  opened.
- Its interface is deliberately small:
  - open a scope with a mode and a policy, which returns an owner token;
  - close a scope by token;
  - record a decision;
  - resolve a message-box prompt for the current policy;
  - resolve the conflict action for a conflict's requested action;
  - read the current mode, policy, and blocked state;
  - read the decisions as JSON and as a structured value.
- Closing a token restores the state saved when it opened, including the
  error level, and discards any inner scopes opened after it. Closing a
  token that isn't open is a no-op. There is no "pop the top" operation.
- Each scope is marked either operation-owned or caller-owned. When an
  operation finishes, it closes only the scopes it opened. Session policies
  set through the public API are caller-owned and close only when the caller
  clears them.
- While any enclosing scope is noninteractive, a request for silent mode —
  from test modules, the public interaction-mode setter, or anywhere else —
  leaves the effective mode noninteractive. Normal and silent behaviour
  outside a noninteractive scope doesn't change.
- `clsOperation` holds one instance and delegates to it. Its existing
  interaction-mode, policy, and blocked members stay as the public surface
  existing callers read.

**Operation lifecycle**

- The policy scope for `MergeBuild` and `RunFilteredTests` opens only after
  the operation has begun, as part of the same step. If the operation can't
  begin, no scope opens and the caller's state is untouched.
- The scope's mode and policy are saved with the operation's registry-backed
  state, so they survive the Stage and Restore cycle of a timer-driven merge.
- Every exit path of a test run finishes the operation: success, failure,
  cancellation, runtime error, and early exit. Every exit path also restores
  logging state and any silent mode the run turned on. A runtime error
  finishes the operation as failed. If a prompt was blocked, it finishes as
  `decision_required`.

**Result contract**

- `MergeBuild` becomes a function that returns a JSON string. Calling it as a
  statement keeps working. The return value describes only whether the merge
  started; the final outcome still comes from the completion callback.
- Refusals that happen before the merge starts are returned synchronously.
  When MCP is active, they are also posted as a completion callback. They use
  these error patterns:
  - `invalid_decision_policy`;
  - `merge_not_available`;
  - `operation_already_running`;
  - `decision_required` (for example, an unknown source folder), with the
    decisions list.
- `RunFilteredTests` uses the same error patterns and shapes for the same
  conditions.
- If a prompt was blocked, a test run that later hits a runtime error returns
  `decision_required` as the primary result, with the decisions list and the
  runtime error both included. `decision_required` is never reported as
  success or as a plain cancellation.
- All decision JSON is built with the existing JSON encoder. The hand-written
  escaping helper is removed.

**Policy semantics**

- `prefer_source` applies each conflict's own requested action and falls
  back to overwrite when none is set. This matches the existing agent
  default.
- `prefer_database` and `skip` keep the same effect (keep the database
  object, skip the source file). Both names stay accepted. The docs and the
  enum comments state that they are equivalent.
- `decline` answers confirmations with No, Cancel, or Abort, and keeps the
  database object for conflicts.
- `block` acknowledges OK-only prompts and records every other prompt as
  `decision_required`.
- `prefer_source`, `prefer_database`, and `skip` never answer a generic
  confirmation.

**Prompt sites**

- All code paths that show the main form go through one helper that respects
  noninteractive mode. This includes the merge path's restore-main-form
  step, the shared-mode reopen, and both build launch points.
- The raw message box in printer-settings import goes through `MsgBox2`.
- Call sites use the single noninteractive predicate rather than comparing
  the mode directly.
- Closing a silent or noninteractive run's window cancels without a second
  confirmation. This stays, and the doc's statement that silent mode is
  unchanged gets this exception.

**Module boundaries and standards**

- Dialog enumeration and dismissal (`ListAddinDialogs`,
  `DismissAddinDialog`) move out of the policy module into their own
  UI-inspection module. They keep their public API methods and are not
  otherwise changed.
- The operation-state helpers that only read `Operation` move onto
  `clsOperation`.
- All new and moved code uses the documented error-handling pattern,
  including the JSON fallback inside `Finish`.
- Naming follows the repo prefixes: `lng` for Long variables, and no
  Variant prefix on typed arrays.
- New log labels are wrapped in `T()`. The button-mask number and the add-in
  form prefix length become named constants.
- Duplicated blocks are consolidated into shared helpers: the "unknown
  policy" message, the "decision not covered" message, the noninteractive
  log block in `MsgBox2`, and the failure branches in dialog dismissal.
- Comments that describe history or make incorrect claims are removed.

## Testing Decisions

**What a good test looks like**

- Tests assert only external behaviour: the prompt's return value, the
  effective mode and policy after open and close, the recorded decisions,
  the reported outcome, and the JSON shape a caller parses. They never
  assert the internal representation of the scope stack.
- Every test that changes live state restores it in teardown, even when an
  assertion fails.

**Seam 1: the interaction scope object (primary)**

- Tests create their own instance, so the live operation, registry state,
  and singletons are never touched. They cover:
  - open, close, and nesting by token;
  - closing out of order, closing twice, and closing an unknown token;
  - operation-owned versus caller-owned release;
  - silent requests not weakening a noninteractive scope;
  - error level restored on close;
  - every policy against every message-box button style;
  - conflict actions for each policy and each conflict type;
  - `decision_required` recording and the resulting outcome;
  - valid JSON for text containing quotes, backslashes, CR, LF, tab, and
    other control characters.

**Seam 2: wiring tests (thin)**

- These are class-based tests. `Class_Initialize` opens a noninteractive
  scope on the running test operation, and `Class_Terminate` closes it by
  token. Each test calls the real `MsgBox2` or the real conflict resolver
  with in-memory conflict items, and asserts that no dialog opened, the
  return value, and the recorded decision.
- After teardown, the test run's own mode, policy, blocked flag, and error
  level must be unchanged. This proves scope restoration on the live
  operation.

**End-to-end verification (not automated in the suite)**

- `MergeBuild` and `RunFilteredTests` can't run inside a test run, because
  an operation is already running. They are verified by a scripted
  `vcs_run_vba` pass against `Testing.accdb` that covers:
  - each pre-start refusal;
  - a blocked conflict;
  - `prefer_source` applied with a delete action;
  - a test run that raises a runtime error, followed by a second call that
    must succeed in starting.
- The script and its expected outputs are recorded in the PR.

**Prior art**

- Pure policy tests: the existing dialog-policy test module.
- Class-based setup and teardown: the existing options and runner-filter
  test classes.
- Saving and restoring live `Operation` state around a call: the existing
  error-handling test module.

## Out of Scope

- Dialogs the add-in doesn't own:
  - VBA `MsgBox` in the database under test;
  - Access error and warning dialogs;
  - runtime-error and compile-error dialogs;
  - trust-center prompts and save-changes confirmations;
  - VBA break mode.

  These remain the MCP-side dialog inspector's job.
- Changing how interactive or plain-silent operations behave, apart from the
  documented cancel-confirmation exception.
- New policy names or per-conflict policies (for example, "source wins for
  forms, database wins for modules").
- Making `MergeBuild` synchronous, or returning its final outcome directly.
- Redesigning the MCP callback payload beyond adding the listed error
  patterns.
- Converting `Dir()` or any unrelated message boxes elsewhere in the add-in.

## Further Notes

- The `DECISIONS.md` entry for the original feature should gain these
  rejected alternatives:
  - "pop the top of stack" instead of owner tokens;
  - opening the scope before `Operation.Begin`;
  - treating `skip` and `prefer_database` as distinct without a distinct
    behaviour.
- Consider a later guard test that scans the add-in's own VBA for bare
  `MsgBox` calls outside `MsgBox2`. That would catch regressions of the raw
  message box this spec fixes. It is left out here to keep to the two agreed
  test seams.
- Some review findings are weaker precedent and can be batched with this
  work rather than treated as blockers:
  - the missing `T()` wrappers are also missing in the older silent-mode
    branch;
  - the pure-delegate API methods on `clsVersionControl` are the documented
    facade and stay as they are.
