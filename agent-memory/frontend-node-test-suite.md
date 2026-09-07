---
name: frontend-node-test-suite
description: How to run the frontend Node test suite and which failures are pre-existing
metadata: 
  node_type: memory
  type: project
  originSessionId: 70ce39fb-ac39-4edd-a4ac-59ca01231bb8
  modified: 2026-09-07T00:00:00.000Z
---

`frontend/package.json` has no `test` script. Run the suite from `frontend/` with the
portable Node: `./node-portable/node.exe --test "tests/**/*.test.mjs"` (bare `node` is not
on PATH; passing the directory instead of the glob fails to resolve).

As of 2026-08-14 the suite at clean HEAD (a71614b) has grown to 705 tests with 14 failures;
the failing set now includes several version-stamp/modularity tests ("DFM runtime imports
never load one module under multiple version URLs", "stateful shared-grid consumers use one
cache-busted module URL", "the data-tab controller and its interaction adapter share the
grid module version"), DFM Excel freshness (x2), dfm_formula_validation, installer progress
patcher, the B&S "app-server and migration adapters retain every canonical frontend contract
value" test, shell add-tab SVG, Data-tab split/facade, Details-format sync, DSV/DFM Data
validation runtime, Result Selection apply, and bundled Codex runtime. Some are flaky
run-to-run, so always diff against a same-commit worktree baseline rather than this list.
On 2026-08-18 at HEAD b7af91d the same run reported 770 tests with 13 failures, the same
set minus installer progress patcher, Details-format sync, and bundled Codex runtime.
On 2026-08-19 at HEAD c2f4e3c it was 822 tests with 15 failures, the same cast again. The
two DFM Excel freshness failures were **not** flakes: the harness in
`tests/dfm_external_links.test.mjs` sliced source between markers containing a bare `
`
while the file on disk is CRLF, so the slice never matched. Fixed on 2026-08-19; if a
source-slicing harness reports "missing <marker>", suspect the line ending before the code.

The same trap has a second shape, found 2026-08-20 in `tests/dfm_formula_validation.test.mjs`:
a **template literal normalises its own CRLF to LF**, so any multi-line `` `import {...}` ``
passed to `.replace()` silently never matches CRLF source and the real site-absolute import
survives into the `data:` module. The symptom is not "marker missing" but
`ERR_INVALID_URL` on a `/ui/...` specifier, because a `data:` URL cannot resolve one.
Fix: normalise the source once (`(await readFile(...)).replaceAll("\r\n", "\n")`), and
point any import the harness does not stub at a real `file://` URL built with `new URL(...)`.
That made the whole `dfm_formula_validation` file load again after months of failing.

On 2026-08-20 at HEAD 4256369 a full run reported 912 tests with 9 failures: the usual
cast (B&S adapters-contract, shared-grid module URL x2, Data-tab split, RS apply, shell
add-tab SVG) plus three `project_settings_source_data.test.mjs` table-summary tests
("table summary service publishes versioned distribution data", "app server owns
date-role detection", "table summary is addressed by project") — all nine confirmed
failing at clean HEAD via a worktree baseline.

On 2026-08-29 (formula-links work, HEAD 2e50e8c plus another session's uncommitted shell/electron
edits in the same tree) a full run was 959 tests with 11 failures: the usual cast (B&S adapters,
shared-grid module URL x2, RS apply, shell add-tab SVG, 3 table-summary) plus `home_shortcuts`
(missing `ui/shared/services/local_day.js`), `dev_window_frame` "development launch", and the PI
Excel Link Manager "refresh icon" test — none of those three read files the formula work touched,
so they belong to the other session's tree state. Harnesses that build the Data-tab persistence
controller with stub factories (`dataset_draft_save.test.mjs`) must stub every link controller
factory the controller creates; a missing one surfaces as "Cannot read properties of undefined
(reading 'isDirty')" rather than a clear import error.

Separately, "changed theme and chart owners are reached through current cache-version chains"
(tests/color_theme.test.mjs) is a **flake in full-suite runs only** — it passes in isolation
and passes on a repeat full run. Earlier "every runtime frontend document bootstraps the
shared theme..." sightings were likely the same flake. Confirm any color_theme failure by
running that file alone before treating it as real. DFM Excel freshness, installer progress
patcher, and bundled Codex runtime also fail intermittently.

On 2026-09-02 (ROUND formula work, HEAD 3ada6f71) the working tree ran 1008 tests with 11
failures and the clean-HEAD worktree 995 tests with 13 (the extra two, installer progress
patcher and bundled Codex runtime, need `node-portable`/build files the worktree lacks). Two
harness lessons from that run: a `data:` module cannot `import` a `file://` URL either, so hand a
real helper to the patched module through `globalThis` (see `tests/dfm_round_formula.test.mjs`);
and `git worktree remove` fails with "Permission denied"/"in use" while the PowerShell tool's
persistent working directory is still inside the worktree — `Set-Location` back to the repo
root first.

Client PC, 2026-09-07 at HEAD 6b038382: a detached-worktree baseline ran 1028 tests with 12 failures
(the usual cast plus `home_shortcuts`, `dev_window_frame`, the PI refresh icon, and the 3 table-summary
tests). `changed theme and chart owners are reached through current cache-version chains` failed at that
baseline for a real reason this time, not the old flake: it pins
`project_settings.js?v=20260903live1`, a stamp the HTML no longer carries. `installer_progress_patcher`
fails only in the worktree (it needs build files the worktree lacks), so it is the one baseline failure
that disappears in the real tree. Copying `frontend/node-portable` into the worktree is enough to run it
there; a junction is not required.

**Why:** the suite is not green, so a failure list alone does not tell you whether a change
broke something. **How to apply:** take a baseline with `git worktree add <tmp> HEAD` and run
the suite there before attributing a failure to your edit. Many tests pin exact `?v=`
cache-busting strings, so bumping a module version per [[arcrho-dev-ui-cache-restart]]
requires updating those pins in the same change.
