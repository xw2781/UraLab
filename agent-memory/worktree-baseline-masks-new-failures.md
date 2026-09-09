---
name: worktree-baseline-masks-new-failures
description: "A fresh git worktree over-reports test failures, so diffing against it hides the failures your change actually caused; stash in the same tree instead"
metadata: 
  node_type: memory
  type: feedback
  originSessionId: 10cf579a-15e3-47d1-91f9-e76d69fe9a77
  modified: 2026-09-09T00:55:17.107Z
---

Taking a test baseline with `git worktree add <tmp> HEAD` **over-reports failures**,
and that silently hides real regressions. On 2026-09-08 the worktree baseline showed
58 python-api failures where the same commit in the working tree showed 21 — the
extra 37 were environment-dependent (lease/propagation tests, engine request queue,
anything reaching `E:\ArcRho Server` or a folder the fresh checkout lacks).

The danger is not the noise, it is the direction: `comm -13 base mine` reports
nothing new when a test fails in **both** lists for **different reasons**. Two real
regressions were invisible that way, including a published ultimate changing from
`1440.0` to `1439.9999999999998`.

**Why:** a worktree is a clean checkout without the gitignored state the suites
depend on — `test/`, `temp/`, `node-portable`, build outputs, local config.

**How to apply:** get the baseline in the *same* working tree, so the environment is
identical:

```
git diff > backup.patch          # belt and braces
git stash push -m baseline -- .
<run the suite, save the failure list>
git stash pop
```

Run the current tree with the **same cwd and PYTHONPATH shape** as the baseline —
differing invocation alone shifted the counts. Then diff. When a test appears in
both lists, open it and read the actual assertion rather than trusting the set
difference. Related: [[frontend-node-test-suite]], [[python-test-runner]].
