from __future__ import annotations

import sys
import json
import subprocess
import unittest
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[2]
FRONTEND_ROOT = REPO_ROOT / "frontend"
MIGRATION_ROOT = REPO_ROOT / "python-api" / "migration"
PYTHON_API_ROOT = REPO_ROOT / "python-api" / "src"
for path in (FRONTEND_ROOT, MIGRATION_ROOT, PYTHON_API_ROOT):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))

from resq_migration import extractors
from app_server.services import result_selection_service


class OutputDatasetType:
    Name = "Selected Ultimate"


class OutputVector:
    Name = "Selection"
    Modified = "2026-01-01T00:00:00Z"
    DatasetType = OutputDatasetType()


class SourceCategory:
    Name = "Loss"


class SourceDatasetType:
    Name = "Paid"
    DataFormat = 1
    Category = SourceCategory()


class SourceDataset:
    Name = "Paid"
    MethodType = 0
    DatasetType = SourceDatasetType()


class ResultSelection:
    OriginLength = 12
    OriginCount = 2
    DatasetCount = 1
    Notes = ""
    OutputVector = OutputVector()
    Name = "Selection"

    def OriginLabel(self, origin_index):
        return str(2024 + origin_index)

    def Dataset(self, _dataset_index):
        return SourceDataset()

    def DatasetValues(self, _dataset_index, origin_index, _origin_length):
        return 1.2345675 if origin_index == 1 else -1.2345675

    def Weights(self, _dataset_index, _origin_index):
        return 1

    def Ultimates(self, origin_index, _origin_length):
        return 50 + origin_index

    def UltimateOverridden(self, *args, **kwargs):
        origin_index = args[0] if args else kwargs.get("OriginIndex")
        return origin_index == 2

    def RatioBasisDataset(self, dataset_index=1, **_kwargs):
        if dataset_index != 1:
            raise IndexError(dataset_index)
        return type("RatioBasis", (), {"Name": "Premium"})()

    def RatioBasisValues(self, *args, **kwargs):
        origin_index = args[0] if args else kwargs.get("OriginIndex")
        return 1000.1234567 if origin_index == 1 else -1000.1234567


class ResultSelectionCrossProducerContractTests(unittest.TestCase):
    def test_migration_and_frontend_emit_exact_same_v2_payload(self) -> None:
        logical_method = ResultSelection()
        migration_payload = extractors.export_result_selection(logical_method)
        migration_payload.pop("_sidecar_notes", None)

        node = REPO_ROOT / "frontend" / "node-portable" / "node.exe"
        contract_uri = (
            REPO_ROOT
            / "frontend"
            / "ui"
            / "method_pages"
            / "result_selection"
            / "result_selection_json_contract.js"
        ).as_uri()
        frontend_input = {
            "details": {
                "name": "Selection",
                "outputType": "Selected Ultimate",
                "originLength": 12,
                "ratioBasis": "Premium",
                "ratioBases": ["Premium"],
                "showRatiosAsPercentages": True,
                "statisticDecimalPlaces": 1,
            },
            "originLabels": ["2025", "2026"],
            "showWeights": True,
            "sources": [{
                "name": "Paid",
                "datasetType": "Paid",
                "dataFormat": "Vector",
                "methodType": "None",
                "category": "Loss",
                "sourceKind": "input",
                "originLength": 12,
                "values": [1.2345675, -1.2345675],
                "weights": [1, 1],
            }],
            "ratioBasisValueSets": [{
                "name": "Premium",
                "values": [1000.1234567, -1000.1234567],
            }],
            "calculatedUltimate": [1.234568, -1.234568],
            "selectedUltimate": [1.234568, 52],
            "ultimateOverrides": [None, 52],
            "lastModified": "2026-01-01T00:00:00Z",
        }
        script = (
            f"import {{buildResultSelectionMethodPayload}} from {json.dumps(contract_uri)};"
            f"console.log(JSON.stringify(buildResultSelectionMethodPayload({json.dumps(frontend_input)})));"
        )
        completed = subprocess.run(
            [str(node), "--input-type=module", "-e", script],
            check=True,
            capture_output=True,
            text=True,
            timeout=30,
        )
        frontend_payload = json.loads(completed.stdout)
        backend_payload = result_selection_service.normalize_method_payload(
            frontend_payload,
            require_complete_basis=True,
        )

        self.assertEqual(frontend_payload, migration_payload)
        self.assertEqual(backend_payload, migration_payload)
        self.assertEqual(migration_payload["json_format"], "arcrho-result-selection-v4")
        self.assertEqual(
            migration_payload["method_tab"]["ratio_basis_values"],
            [{"name": "Premium", "values": [1000.123457, -1000.123457]}],
        )
        self.assertEqual(
            migration_payload["method_tab"]["loaded_datasets"][0]["values"],
            [1.234568, -1.234568],
        )


if __name__ == "__main__":
    unittest.main()
