---
name: persisted-json-v4-progress
description: "The persisted JSON contract v4 plan is closed (Implemented 2026-09-06, moved to docs/plans/completed/); what shipped, what was deliberately dropped, and the converter facts that still matter"
metadata:
  node_type: memory
  type: project
  originSessionId: 54e67477-d0c1-4c16-a51e-f1b17a4e537e
  modified: 2026-09-07T01:58:11.796Z
---

Persisted JSON contract v4 is complete and live. The plan now lives at `docs/plans/completed/persisted_json_contract_v4.md` with Status "Implemented 2026-09-06".

- Steps 1-5 committed on `main` (c2ed598, 2de7263, cdcea68, `d432ae8`, `fcbfaeb`); step 6 converter `tools/migrate_persisted_json_v4.py` landed as `d8b4baf` and was applied to `NJ_Annual_Prod_202605_Fake` only (2,738 files, stable fixed point). **The other 36 projects were deliberately left unconverted** and are to be deleted and re-imported from ResQ by hand — so a share-wide scan still finds thousands of pre-v4 sidecars, which is expected, not a defect.
- Step 7: Engine, Bridge, Gateway redeployed and macros republished 2026-08-23; frontend release 1.3.3 (2026-08-23) carries the breaking-change note, and 1.4.0-1.4.2 followed.
- **Dropped by the user's decision on 2026-09-06:** the deferred `/code-review ultra` pass, and forcing the release (no `mandatory: true` marker was ever published) — see [[all-users-run-latest-app-version]].

**Why:** an agent that sees unstamped or spaced-key files on the share, or notices the release was never forced, may think the plan is unfinished. It is not; both were closed on purpose.

**How to apply:** do not reopen the plan or propose converting the remaining projects; a pre-v4 file in a project other than `NJ_Annual_Prod_202605_Fake` is handled by re-import, not by the converter. If the converter is ever run again: methods convert before sidecars and the sidecar's `publication_revision` comes from the converted method; BF v2 / DFM v1 stamps are rescued (notes to sidecar) not converted; run it on the Server PC because zoneless timestamps are server wall-clock; calculated caches recalculate once afterwards because their evidence lives in `.arcrho-cache-provenance/`.

Related: [[python-test-runner]], [[remote-component-deploy]], [[shared-macro-library-deploy]], [[deploy-staleness-is-mtime-based]].
