"""Every sidecar producer emits the one shared core, with ``audit_log`` last."""

from __future__ import annotations

import json
import sys
import unittest
from pathlib import Path


sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))
sys.path.insert(0, str(Path(__file__).resolve().parent))

from arcrho_api import bootstrap_contract, bornhuetter_ferguson_contract, cape_cod_contract, dfm_contract  # noqa: E402
from arcrho_api.engine_dataset_sidecar_contract import build_engine_dataset_sidecar  # noqa: E402
from arcrho_api.sidecar_core_contract import (  # noqa: E402
    METHOD_OUTPUT_SIDECAR_FIELDS,
    SIDECAR_CORE_FIELDS,
    SIDECAR_STORED_DEVELOPMENT_FIELD,
    SIDECAR_STORED_ORIGIN_FIELD,
    SIDECAR_STORED_PERIOD_FIELD,
    SidecarContractError,
    display_lengths,
    linked_length_fields,
    linked_lengths,
    stored_length_fields,
    stored_lengths,
    validate_sidecar_core,
    with_audit_log_last,
)
import test_bootstrap_contract as bootstrap_tests  # noqa: E402
import test_bornhuetter_ferguson_contract as bf_tests  # noqa: E402
import test_cape_cod_contract as cc_tests  # noqa: E402
import test_dfm_contract as dfm_tests  # noqa: E402


_PRIOR = {
    "show_subtotal": False,
    "audit_log": [
        {"event_date": "2026-08-01T00:00:00Z", "action": "Insert", "change_info": "", "user": "Dana"},
        {"event_date": "2026-08-02T00:00:00Z", "action": "Auto Refresh", "change_info": "", "user": "Engine"},
        {"event_date": "2026-08-03T00:00:00Z", "action": "Auto Refresh", "change_info": "", "user": "Engine"},
    ],
}


def _dfm_sidecar() -> dict:
    method = dfm_contract.recalculate_dfm_method(
        dfm_tests.owned_payload(),
        input_snapshot=dfm_tests.input_snapshot(),
        ratio_basis_snapshot=dfm_tests.basis_snapshot(),
    )
    return dfm_contract.build_dfm_output_sidecar(
        method,
        project_name="Demo",
        reserving_class=r"Auto\PP",
        csv_file="Paid Selected@12.csv",
        existing=_PRIOR,
        user="tester",
        timestamp="2026-08-04T00:00:00Z",
    )


def _bf_sidecar() -> dict:
    return bornhuetter_ferguson_contract.build_bornhuetter_ferguson_output_sidecar(
        bf_tests.complete_method(),
        project_name="Demo",
        reserving_class=r"Auto\PP",
        csv_file="BF Ultimate@12.csv",
        existing=_PRIOR,
        user="tester",
        timestamp="2026-08-04T00:00:00Z",
    )


def _cc_sidecar() -> dict:
    return cape_cod_contract.build_cape_cod_output_sidecar(
        cc_tests.complete_method(),
        project_name="Demo",
        reserving_class=r"Auto\PP",
        csv_file="CC Ultimate@12.csv",
        existing=_PRIOR,
        user="tester",
        timestamp="2026-08-04T00:00:00Z",
    )


def _bootstrap_sidecar() -> dict:
    fixture = json.loads(bootstrap_tests.FIXTURE.read_text(encoding="utf-8"))
    case = fixture["methods"]["odp_single_scale"]
    reference = fixture["simulation_reference"]
    method = bootstrap_contract.recalculate_bootstrap_method(
        bootstrap_tests._seed_payload(case, reference),
        dfm_snapshot=bootstrap_tests._snapshot_from_fixture(case),
        target_snapshot=bootstrap_tests._target_snapshot(case, reference),
        timestamp="2026-08-05T00:00:00Z",
    )
    return bootstrap_contract.build_bootstrap_output_sidecar(
        method,
        project_name="Demo",
        reserving_class=r"Auto\PP",
        csv_file="F 72 A@12.csv",
        existing=_PRIOR,
        user="tester",
        timestamp="2026-08-05T00:00:00Z",
    )


def _engine_sidecar() -> dict:
    return build_engine_dataset_sidecar(
        project_name="Demo",
        reserving_class=r"Auto\PP",
        dataset_name="Paid Loss",
        dataset_type="Paid Loss",
        data_format="Triangle",
        csv_file="Paid Loss@12@12@cum@dev.csv",
        user="tester",
        created="2026-08-01T00:00:00Z",
        updated_at="2026-08-04T00:00:00Z",
        number_format="0,000",
        decimal_places=1,
        origin_length=12,
        development_length=12,
        audit_log=_PRIOR["audit_log"],
        audit_action="Update",
    )


PRODUCERS = {
    "dfm": _dfm_sidecar,
    "bornhuetter_ferguson": _bf_sidecar,
    "cape_cod": _cc_sidecar,
    "bootstrap": _bootstrap_sidecar,
    "engine": _engine_sidecar,
}


class CrossWriterSidecarCoreTests(unittest.TestCase):
    def test_every_producer_passes_the_shared_validator(self) -> None:
        for name, build in PRODUCERS.items():
            with self.subTest(producer=name):
                sidecar = build()
                self.assertEqual(validate_sidecar_core(sidecar), sidecar)
                self.assertEqual(list(sidecar)[-1], "audit_log")
                self.assertTrue(set(SIDECAR_CORE_FIELDS) <= set(sidecar))

    def test_method_outputs_add_only_the_method_fields_on_top(self) -> None:
        engine = _engine_sidecar()
        for name, build in PRODUCERS.items():
            if name == "engine":
                continue
            with self.subTest(producer=name):
                sidecar = build()
                for field in METHOD_OUTPUT_SIDECAR_FIELDS:
                    self.assertIn(field, sidecar)
                self.assertIs(sidecar["calculated"], True)
                self.assertEqual(sidecar["method_type"], sidecar["method_type"].strip())
        for field in METHOD_OUTPUT_SIDECAR_FIELDS:
            self.assertNotIn(field, engine)

    def test_every_producer_records_the_shape_its_csv_is_stored_at(self) -> None:
        # A method output is produced at its own period and an engine cache at
        # the requested one, so each writer's stored shape equals its display
        # shape for these inputs -- but every writer emits the pair.
        for name, build in PRODUCERS.items():
            with self.subTest(producer=name):
                sidecar = build()
                if sidecar["data_format"] == "Vector":
                    self.assertEqual(
                        sidecar[SIDECAR_STORED_PERIOD_FIELD], sidecar["period_length"]
                    )
                    self.assertNotIn(SIDECAR_STORED_ORIGIN_FIELD, sidecar)
                else:
                    self.assertEqual(
                        (
                            sidecar[SIDECAR_STORED_ORIGIN_FIELD],
                            sidecar[SIDECAR_STORED_DEVELOPMENT_FIELD],
                        ),
                        (sidecar["origin_length"], sidecar["development_length"]),
                    )
                    self.assertNotIn(SIDECAR_STORED_PERIOD_FIELD, sidecar)

    def test_an_engine_sidecar_keeps_a_finer_stored_shape_than_it_displays(self) -> None:
        sidecar = build_engine_dataset_sidecar(
            project_name="Demo",
            reserving_class=r"Auto\PP",
            dataset_name="Paid Loss",
            dataset_type="Paid Loss",
            data_format="Triangle",
            csv_file="Paid Loss@12@12@cum@dev.csv",
            user="tester",
            created="2026-08-01T00:00:00Z",
            updated_at="2026-08-04T00:00:00Z",
            number_format="0,000",
            decimal_places=1,
            origin_length=12,
            development_length=12,
            stored_origin_length=1,
            stored_development_length=3,
        )
        self.assertEqual(sidecar[SIDECAR_STORED_ORIGIN_FIELD], 1)
        self.assertEqual(sidecar[SIDECAR_STORED_DEVELOPMENT_FIELD], 3)

    def test_every_producer_appends_under_the_one_audit_policy(self) -> None:
        # The prior log carries two consecutive automatic entries; each writer
        # keeps the history, collapses the run, and appends its own record.
        for name, build in PRODUCERS.items():
            with self.subTest(producer=name):
                log = build()["audit_log"]
                self.assertEqual([item["action"] for item in log], ["Insert", "Auto Refresh", "Update"])
                self.assertEqual(log[1]["event_date"], "2026-08-03T00:00:00Z")
                self.assertEqual(log[-1]["user"], "tester")

    def test_a_method_writer_without_an_audit_entry_still_normalizes_the_log(self) -> None:
        method = dfm_contract.recalculate_dfm_method(
            dfm_tests.owned_payload(),
            input_snapshot=dfm_tests.input_snapshot(),
            ratio_basis_snapshot=dfm_tests.basis_snapshot(),
        )
        sidecar = dfm_contract.build_dfm_output_sidecar(
            method,
            project_name="Demo",
            reserving_class=r"Auto\PP",
            csv_file="Paid Selected@12.csv",
            existing=_PRIOR,
            append_audit=False,
        )
        self.assertEqual([item["action"] for item in sidecar["audit_log"]], ["Insert", "Auto Refresh"])


class ValidatorTests(unittest.TestCase):
    def test_a_missing_core_field_is_named(self) -> None:
        sidecar = _engine_sidecar()
        sidecar.pop("csv_file")
        with self.assertRaises(SidecarContractError) as caught:
            validate_sidecar_core(sidecar)
        self.assertIn("csv_file", str(caught.exception))

    def test_a_sidecar_without_a_stored_shape_is_refused(self) -> None:
        sidecar = dict(_engine_sidecar())
        sidecar.pop(SIDECAR_STORED_DEVELOPMENT_FIELD)
        with self.assertRaises(SidecarContractError) as caught:
            validate_sidecar_core(with_audit_log_last(sidecar))
        self.assertIn(SIDECAR_STORED_DEVELOPMENT_FIELD, str(caught.exception))

    def test_a_display_shape_that_is_not_a_multiple_of_the_stored_one_is_refused(self) -> None:
        # A view coarser than the data is a roll-up; one that does not divide
        # evenly cannot be built at all, so it must never be persisted.
        sidecar = with_audit_log_last(
            {**_engine_sidecar(), SIDECAR_STORED_ORIGIN_FIELD: 5}
        )
        with self.assertRaises(SidecarContractError) as caught:
            validate_sidecar_core(sidecar)
        self.assertIn("whole multiple", str(caught.exception))

    def test_a_stored_shape_must_be_a_positive_number_of_months(self) -> None:
        sidecar = with_audit_log_last(
            {**_engine_sidecar(), SIDECAR_STORED_ORIGIN_FIELD: 0}
        )
        with self.assertRaises(SidecarContractError):
            validate_sidecar_core(sidecar)

    def test_the_stored_fields_a_format_takes_are_written_once(self) -> None:
        self.assertEqual(stored_length_fields("Vector", 3), {SIDECAR_STORED_PERIOD_FIELD: 3})
        self.assertEqual(
            stored_length_fields("Triangle", 1, 3),
            {SIDECAR_STORED_ORIGIN_FIELD: 1, SIDECAR_STORED_DEVELOPMENT_FIELD: 3},
        )

    def test_a_reader_is_told_the_stored_shape_whatever_the_format(self) -> None:
        # The read side of the same rule: a reader that opens the CSV never has
        # to know which field its format keeps, and never sees the display one.
        self.assertEqual(
            stored_lengths({
                "data_format": "Triangle",
                "origin_length": 12,
                "development_length": 12,
                "stored_origin_length": 1,
                "stored_development_length": 3,
            }),
            (1, 3),
        )
        self.assertEqual(
            stored_lengths({"data_format": "Vector", "period_length": 12, "stored_period_length": 3}),
            (3, 3),
        )
        self.assertEqual(stored_lengths({"data_format": "Triangle", "origin_length": 12}), (0, 0))
        self.assertEqual(stored_lengths({"data_format": "Vector", "stored_period_length": "x"}), (0, 0))

    def test_a_reader_is_told_the_display_shape_the_same_way(self) -> None:
        # The display pair is read through the same door, so a window that
        # opens a dataset at the shape it was saved at never sees the stored one.
        self.assertEqual(
            display_lengths({
                "data_format": "Triangle",
                "origin_length": 12,
                "development_length": 12,
                "stored_origin_length": 12,
                "stored_development_length": 1,
            }),
            (12, 12),
        )
        self.assertEqual(
            display_lengths({"data_format": "Vector", "period_length": 12, "stored_period_length": 3}),
            (12, 12),
        )
        self.assertEqual(display_lengths({"data_format": "Triangle", "stored_origin_length": 1}), (0, 0))

    def test_the_linked_shape_is_its_own_pair_and_defaults_to_the_display(self) -> None:
        # The display a dataset's cell links were written against is recorded
        # apart from the display it is shown at now; a sidecar that states none
        # was saved with its links at the display it records.
        self.assertEqual(
            linked_length_fields("Triangle", 1, 1),
            {"linked_origin_length": 1, "linked_development_length": 1},
        )
        self.assertEqual(linked_length_fields("Vector", 3), {"linked_period_length": 3})
        self.assertEqual(
            linked_lengths({
                "data_format": "Triangle",
                "origin_length": 12,
                "development_length": 12,
                "linked_origin_length": 1,
                "linked_development_length": 1,
            }),
            (1, 1),
        )
        self.assertEqual(
            linked_lengths({"data_format": "Triangle", "origin_length": 12, "development_length": 12}),
            (12, 12),
        )
        self.assertEqual(
            linked_lengths({"data_format": "Vector", "period_length": 12, "linked_period_length": 3}),
            (3, 3),
        )
        self.assertEqual(linked_lengths({"data_format": "Vector", "period_length": 12}), (12, 12))

    def test_the_audit_log_must_be_last(self) -> None:
        sidecar = _engine_sidecar()
        sidecar["extra"] = 1
        with self.assertRaises(SidecarContractError) as caught:
            validate_sidecar_core(sidecar)
        self.assertIn("last field", str(caught.exception))
        self.assertEqual(validate_sidecar_core(with_audit_log_last(sidecar)), with_audit_log_last(sidecar))

    def test_the_audit_log_must_already_follow_the_policy(self) -> None:
        sidecar = _engine_sidecar()
        sidecar["audit_log"] = [*sidecar["audit_log"], {"event_date": "", "action": "Update"}]
        with self.assertRaises(SidecarContractError):
            validate_sidecar_core(sidecar)

    def test_a_named_method_output_is_always_calculated(self) -> None:
        sidecar = dict(_engine_sidecar())
        sidecar = with_audit_log_last({**sidecar, "method_name": "M"})
        with self.assertRaises(SidecarContractError):
            validate_sidecar_core(sidecar)  # not calculated
        sidecar = with_audit_log_last({**sidecar, "calculated": True})
        validate_sidecar_core(sidecar)

    def test_a_method_output_may_publish_no_revision(self) -> None:
        # Berquist Sherman has no contract module and computes no publication
        # fingerprint, so its output sidecars name a method and stop there.
        sidecar = with_audit_log_last({**_engine_sidecar(), "method_name": "B&S", "calculated": True})

        validate_sidecar_core(sidecar)

    def test_a_revision_without_a_method_name_is_refused(self) -> None:
        sidecar = with_audit_log_last({**_engine_sidecar(), "publication_revision": "sha256:0000000000000000"})

        with self.assertRaises(SidecarContractError):
            validate_sidecar_core(sidecar)

    def test_finalize_stamps_first_moves_the_log_last_and_drops_retired_fields(self) -> None:
        payload = {
            "audit_log": [{"event_date": "d", "action": "insert", "user": "u"}],
            "a": 1,
            "user": "u",
            "Precedents": [],
            "processing_by_csv": {},
        }
        ordered = with_audit_log_last(payload)
        self.assertEqual(list(ordered), ["json_format", "a", "audit_log"])
        self.assertEqual(ordered["json_format"], "arcrho-dataset-sidecar-v4")
        self.assertEqual(ordered["audit_log"][0]["action"], "Insert")
        self.assertEqual(with_audit_log_last({"a": 1})["audit_log"], [])

    def test_finalize_gives_both_graph_fields_the_one_entry_shape(self) -> None:
        from arcrho_api.sidecar_core_contract import finalize_sidecar

        ordered = finalize_sidecar({
            "precedents": ["Paid", "paid", {"dataset_name": "Premium"}],
            "dependents": [{"dataset_name": "Selected", "method_type": "Result Selection", "path": "x"}],
            "audit_log": [],
        })
        self.assertEqual(ordered["precedents"], [{"dataset_name": "Paid"}, {"dataset_name": "Premium"}])
        self.assertEqual(ordered["dependents"], [{"dataset_name": "Selected", "method_type": "Result Selection"}])
        self.assertEqual(finalize_sidecar({"precedents": None, "dependents": "Paid"})["precedents"], [])

    def test_dependency_entries_have_one_shape(self) -> None:
        from arcrho_api.sidecar_core_contract import dependency_entries, dependency_names

        entries = dependency_entries([
            "Paid",
            {"dataset_name": "paid"},
            {"dataset_name": "Ultimate", "method_type": "DFM", "path": r"C:\x.json", "mtime": 1},
            {"dataset_name": "Plain", "method_type": "None"},
            {"dataset_name": "Far", "reserving_class": "Other", "project": ""},
            {"name": "ignored"},
            "",
        ], method_types={"plain": "Result Selection"})
        self.assertEqual(entries, [
            {"dataset_name": "Paid"},
            {"dataset_name": "Ultimate", "method_type": "DFM"},
            {"dataset_name": "Plain", "method_type": "Result Selection"},
            {"dataset_name": "Far", "reserving_class": "Other"},
        ])
        self.assertEqual(dependency_names(entries), ["Paid", "Ultimate", "Plain", "Far"])
        self.assertEqual(dependency_names(None), [])

    def test_retired_fields_and_pathed_entries_are_refused(self) -> None:
        sidecar = _engine_sidecar()
        for field in ("method_type_code", "data_format_code", "origin_count", "user", "formula", "processing_by_csv"):
            with self.subTest(field=field):
                bad = with_audit_log_last({**sidecar, field: 1})
                bad[field] = 1
                bad["audit_log"] = bad.pop("audit_log")
                with self.assertRaises(SidecarContractError):
                    validate_sidecar_core(bad)
        pathed = dict(sidecar)
        pathed["precedents"] = [{"dataset_name": "Paid", "path": r"C:\paid.csv"}]
        with self.assertRaises(SidecarContractError):
            validate_sidecar_core(pathed)


if __name__ == "__main__":
    unittest.main()
