---
name: persisted-json-v4-progress
description: "Where the persisted JSON contract v4 work stands: steps 1-6 and the server deploy are done and live, but the frontend release was never marked mandatory, /code-review ultra was never run, and the plan's Status line is stale"
metadata:
  node_type: memory
  type: project
  originSessionId: 54e67477-d0c1-4c16-a51e-f1b17a4e537e
  modified: 2026-09-07T01:48:20.315Z
---

Persisted JSON contract v4 (`docs/plans/persisted_json_contract_v4.md`) is functionally complete and live. Verified 2026-09-06:

- Steps 1-5 committed on `main` (c2ed598, 2de7263, cdcea68, `d432ae8`, `fcbfaeb`); step 6 converter `tools/migrate_persisted_json_v4.py` landed as `d8b4baf` and was applied to `NJ_Annual_Prod_202605_Fake` only (2,738 files, stable fixed point). The other 36 projects were deliberately left unconverted and are to be deleted and re-imported from ResQ by hand.
- Step 7: Engine, Bridge, Gateway redeployed and macros republished 2026-08-23. Frontend release **1.3.3 (2026-08-23) carries the v4 breaking-change note** and 1.4.0-1.4.2 followed.

**Still open, and easy to mistake for done:**

- **The release was never forced.** The updater treats an update as mandatory only when a GitHub release body carries `mandatory: true` (`update_checker.js`, written by `publish_github_release.ps1 -Mandatory`). None of 1.3.3 through 1.4.2 carry it (checked via the GitHub API). The "old client cannot open a converted workspace silently" check was never recorded either.
- `/code-review ultra` over the whole v4 change was deferred and never run. There is no branch to diff; it needs a path target (`python-api/src/arcrho_api/` contracts first).
- The plan's `Status:` line still says "In progress … Step 6 is next" with Last updated 2026-08-23. Two Step 7 boxes are unticked (force the release; set Status to Implemented).

**Why:** the plan header is stale, so an agent reading only the top of the doc will think Step 6 is pending. Read the checklist, not the header.

**How to apply:** if asked to finish the plan, the remaining work is (1) decide whether a mandatory marker is still wanted now that everyone is on 1.4.x, (2) run or explicitly drop the ultra review, (3) tick the two boxes and set Status to Implemented with the date.

Converter facts still worth knowing: methods convert before sidecars and the sidecar's `publication_revision` comes from the converted method; BF v2 / DFM v1 stamps are rescued (notes to sidecar) not converted; run on the Server PC because zoneless timestamps are server wall-clock; calculated caches recalculate once after conversion because their evidence lives in `.arcrho-cache-provenance/`.

Related: [[python-test-runner]], [[remote-component-deploy]], [[shared-macro-library-deploy]], [[deploy-staleness-is-mtime-based]].
