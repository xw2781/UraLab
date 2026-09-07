---
name: all-users-run-latest-app-version
description: "Project assumption set 2026-09-06: every ArcRho user is on the latest app version at all times, so releases need no mandatory marker, no old-client compatibility, and no forced-update step in plans"
metadata: 
  node_type: memory
  type: project
  originSessionId: f2de0503-6083-4cca-b1de-781d7696fbb4
  modified: 2026-09-07T01:57:47.951Z
---

Every ArcRho user is assumed to run the latest released app version at all times. Decided by the user on 2026-09-06 while closing the persisted JSON v4 plan.

**Why:** the user base is small and updates promptly, so the cost of keeping old-client compatibility, forcing releases, or verifying that an old client fails loudly is not worth paying. The `mandatory: true` release-body marker that `update_checker.js` honours exists but is not used in practice: none of ArcRho 1.2.6 through 1.4.2 carry it.

**How to apply:**

- Do not add "force the release" or "confirm an old client cannot open this" steps to plans, and do not treat a release published without the mandatory marker as an unfinished breaking change.
- A breaking change to a persisted file or a server contract still needs every component (Engine, Bridge, Gateway, frontend, macros) released together, as in [[remote-component-deploy]] and [[shared-macro-library-deploy]]; what it does not need is a fallback for clients that lag behind.
- Legacy readers for old file shapes may be deleted outright once the data is converted, as the v4 work did ([[persisted-json-v4-progress]]).
