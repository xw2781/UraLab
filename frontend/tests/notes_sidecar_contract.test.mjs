import assert from "node:assert/strict";
import test from "node:test";
import { readFile } from "node:fs/promises";

const repoFile = (relativePath) => new URL(`../../${relativePath}`, import.meta.url);
const read = (relativePath) => readFile(repoFile(relativePath), "utf8");

test("dataset sidecars are the only persisted notes owner", async () => {
  const runtimeFiles = await Promise.all([
    "frontend/app_server/services/dataset_service.py",
    "frontend/app_server/services/dataset_instance_index_service.py",
    "python-api/src/arcrho_api/dataset_index_contract.py",
    "python-api/migration/resq_migration/core.py",
    "python-api/migration/resq_migration/catalog.py",
  ].map(read));
  assert.doesNotMatch(runtimeFiles.filter((_source, index) => index !== 2).join("\n"), /ArcRhoTriNotes@/u);
  assert.match(runtimeFiles[2], /migrate_legacy_notes_files/u);

  // Every producer of a method file, in both languages. Three BF files on the
  // server once carried real commentary in a `notes_tab` that nothing else
  // held, so a producer that writes the section - even as `{}` in Python,
  // which the old `\bnotes_tab\s*:` pattern could not see through the quotes
  // - would reopen that trap.
  const methodFiles = await Promise.all([
    "frontend/ui/method_pages/dfm/dfm_persistence.js",
    "frontend/ui/method_pages/result_selection/result_selection_model.js",
    "frontend/ui/method_pages/result_selection/result_selection_json_contract.js",
    "frontend/ui/method_pages/bornhuetter_ferguson/bornhuetter_ferguson_main.js",
    "frontend/ui/method_pages/bornhuetter_ferguson/bornhuetter_ferguson_json_contract.js",
    "frontend/ui/method_pages/cape_cod/cape_cod_main.js",
    "frontend/ui/method_pages/cape_cod/cape_cod_json_contract.js",
    "frontend/ui/method_pages/berquist_sherman/berquist_sherman_main.js",
    "frontend/app_server/services/result_selection_service.py",
    "frontend/app_server/services/berquist_sherman_service.py",
    "frontend/app_server/services/dfm_service.py",
    "frontend/app_server/services/bornhuetter_ferguson_service.py",
    "frontend/app_server/services/cape_cod_service.py",
    "frontend/app_server/services/bootstrap_service.py",
    "python-api/src/arcrho_api/dfm_contract.py",
    "python-api/src/arcrho_api/bornhuetter_ferguson_contract.py",
    "python-api/src/arcrho_api/cape_cod_contract.py",
    "python-api/src/arcrho_api/bootstrap_contract.py",
    "python-api/src/arcrho_api/dfm.py",
    "server-components/src/arcrho_bridge/resq_client.py",
    "python-api/migration/resq_migration/dfm.py",
    "python-api/migration/resq_migration/extractors.py",
  ].map(read));
  const methodSource = methodFiles.join("\n");
  assert.doesNotMatch(methodSource, /["']notes tab["']\s*:/u);
  assert.doesNotMatch(methodSource, /["']?notes_tab["']?\s*:/u);

  const datasetService = runtimeFiles[0];
  assert.match(datasetService, /"notes": str\(payload\.get\("notes"\) or ""\)/u);
  assert.match(datasetService, /payload\["notes"\] = str\(notes/u);

  const rpcSnapshots = await read("frontend/app_server/services/dfm_rpc_bridge_service.py");
  assert.doesNotMatch(rpcSnapshots, /^\s*["']notes["']\s*:/mu);
});

test("Project Instance DSV open uses the aggregate sidecar and CSV cache route", async () => {
  const controller = (await Promise.all([
    "frontend/ui/shared/tabs/data/data_tab_controller.js",
    "frontend/ui/shared/tabs/data/data_tab_host_controller.js",
    "frontend/ui/shared/tabs/data/data_tab_persistence_controller.js",
  ].map(read))).join("\n");
  assert.match(controller, /isProjectInstanceCachedDatasetOpen/u);
  assert.match(controller, /loadProjectInstanceCachedDataset/u);
  assert.match(controller, /loadCachedDataset\(/u);
  assert.match(controller, /sidecarData: data/u);
});

test("Project Instance method opens reuse the parent dataset-index snapshot", async () => {
  const [cache, shared, resultSelection, bf, cc, bs] = await Promise.all([
    read("frontend/ui/project_instance/project_instance_dataset_cache.js"),
    read("frontend/ui/shared/dataset/project_instance_dataset_snapshot.js"),
    read("frontend/ui/method_pages/result_selection/result_selection_data.js"),
    read("frontend/ui/method_pages/bornhuetter_ferguson/bornhuetter_ferguson_main.js"),
    read("frontend/ui/method_pages/cape_cod/cape_cod_main.js"),
    read("frontend/ui/method_pages/berquist_sherman/berquist_sherman_main.js"),
  ]);
  assert.match(cache, /publishProjectInstanceDatasetSnapshot\(projectName, normalizedPath, payload\)/u);
  assert.match(shared, /window\.sessionStorage/u);
  for (const methodSource of [resultSelection, bf, cc, bs]) {
    assert.match(methodSource, /readProjectInstanceDatasetSnapshot/u);
    assert.match(methodSource, /project_instance/u);
  }
});
