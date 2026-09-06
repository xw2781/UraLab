---
name: resq-com-probe
description: ResQ COM is reachable only on the Dev PC; the Client PC L-H2MQ6280FVP has no ResQ install, so in-process COM macros fail there
metadata: 
  node_type: memory
  type: project
  originSessionId: 0a5616ae-9d5a-4766-a011-883e2b0109cc
  modified: 2026-09-06T23:22:18.210Z
---

ResQ (Willis Towers Watson) is installed **only on the Dev PC**, where its COM API is directly reachable for read-only debugging: `win32com.client.Dispatch("ResQ3Automation.ResQApplication")` then `ConnectByName("JGO_CO1SQLWPV22", "", "")` (connection name from `server-components/src/arcrho_bridge/resq_client.py`). The Client PC **L-H2MQ6280FVP has no ResQ install at all** — verified 2026-08-17: no `ResQ3Automation.*` ProgID in HKLM\SOFTWARE\Classes, its WOW6432Node view, or HKCU\SOFTWARE\Classes; no vendor key, uninstall entry, install folder, or MsiInstaller event. Any in-process COM there fails with `(-2147221005, 'Invalid class string')` while `Excel.Application` dispatches fine.

**Why:** Lets an agent verify what ResQ actually returns (e.g. `dfm.CellNotes` vs method-level `dfm.Notes`) instead of guessing from bridge code — and explains why the ResQ import, sync, and (since 2026-08-27) export macros all go through the Bridge queue: the Bridge worker on the ResQ machine services them. The one macro that still Dispatches ResQ in the local macro process is `import_resq_dataset.py`, so it works only on the Server PC.

**How to apply:** Run probe scripts on the Server PC NE7SASWPN02 with `py -3.10` (`C:\Program Files\Python310`, has pywin32; the old `server-components/venvs/arcrho_bridge` venv no longer exists there as of 2026-09-06, and 3.14 has no pywin32). The shared service account is in `E:\ArcRho Server\config\config.json` under `resq` (connection name, user name, password); the probe scripts read it from there. Use `sys.stdout.reconfigure(encoding="utf-8", errors="replace")` — ResQ notes contain characters like U+25E6 that crash cp1252 console printing. Keep probes read-only; never call `Save()`. Related: [[bridge-restart-after-deploy]], [[dev-pc-and-client-pc-identity]].

**Use early binding, always.** `win32com.client.Dispatch` (late binding) silently returns *wrong* values for some ResQ properties instead of raising: on 2026-08-30 every one of the 260 `DatasetType` objects in a project reported `Calculated=False, Formula=''`, when 71 of them are calculated and carry a formula. `gencache.EnsureDispatch("ResQ3Automation.ResQApplication")` returns the real values. Instance-level properties (`Vector.Formula`) happened to survive late binding, so a probe can look plausible while its type-level answers are pure fiction. Never trust a late-bound ResQ read that returns a uniform empty/False answer across a whole collection.

**Dataset types and dataset instances are separate formula stores.** `DatasetType.Formula` is the project-wide definition; `Vector.Formula` / `Triangle.Formula` is that one reserving class's own. They disagree often, and a type can be un-calculated project-wide while every instance carries the same formula. `ReservingClass.Calculated` marks a roll-up class but does *not* mean its datasets are aggregations — roll-up classes hold genuine in-class formulas too. The reliable marker of an aggregation formula is a backslash in the formula text (it names other reserving-class paths).
