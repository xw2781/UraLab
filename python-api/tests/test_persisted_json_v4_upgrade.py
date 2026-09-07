"""The pre-v4 upgrade the conversion script runs (docs/plans/completed/persisted_json_contract_v4.md).

``arcrho_api.persisted_json_v4_upgrade`` is the only module that still knows
the old spellings, so it is also the only place these rules can be pinned.
"""

from __future__ import annotations

import json
import os
import sys
import unittest

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "src"))

from arcrho_api.dfm_contract import DFM_JSON_FORMAT  # noqa: E402
from arcrho_api.io import persisted_json_text  # noqa: E402
from arcrho_api.persisted_json_v4_upgrade import (  # noqa: E402
    DATASET_NUMBER_FORMATS_JSON_FORMAT,
    METHOD_FORMAT_UPGRADES,
    PROJECT_AUDIT_LOG_JSON_FORMAT,
    RUNTIME_CACHE_PROVENANCE_JSON_FORMAT,
    UNCONVERTIBLE_METHOD_FORMATS,
    PersistedJsonUpgradeError,
    UnsupportedMethodFormatError,
    sidecar_with_method_notes,
    snake_key,
    stranded_method_notes,
    upgrade_dataset_number_formats,
    upgrade_dataset_sidecar,
    upgrade_method,
    upgrade_project_audit_log,
    upgrade_runtime_cache_provenance,
    upgrade_source_import,
)
from arcrho_api.sidecar_core_contract import RETIRED_SIDECAR_FIELDS, validate_sidecar_core  # noqa: E402
from arcrho_api.source_table_contract import SOURCE_IMPORT_JSON_FORMAT  # noqa: E402


def old_dfm_method() -> dict:
    """A pre-v4 DFM file, cut down to the fields the upgrade has to act on."""

    return {
        "json format": "arcrho-dfm-method-by-tab-v2",
        "method metadata": {
            "method name": "C 12",
            "last modified": "2026-08-19T10:05:30.500000",
            "averageType": "Volume Weighted",
        },
        "data tab": {
            "origin labels": ["2024", "2025"],
            "development labels": ["12", "24"],
            "input data triangle mask": [[True, True], [True, False]],
            "source revision": "sha256:" + "a" * 64,
        },
        "ratios tab": {
            "ratio triangle": {
                "origin labels": ["2024", "2025"],
                "development labels": ["(1) 12-24"],
            },
        },
        "results tab": {
            "ratio basis origin labels": ["2024", "2025"],
            "selected ultimates": [100, 200],
        },
        "notes tab": {"notes": "  Tail based on a competitor study.  "},
        "chart tab": {},
        "audit log tab": {},
        "audit_log": [{"event_date": "2026-08-19T10:05:30", "action": "Update"}],
    }


def old_sidecar() -> dict:
    return {
        "dataset_name": "ALAE--Paid",
        "dataset_type": "ALAE--Paid",
        "reserving_class": "HOL",
        "project_name": "NJ",
        "source_kind": "engine",
        "calculated": False,
        "data_format": "Triangle",
        "method_type": "",
        "status": "Current",
        "number_format": "Comma",
        "decimal_places": 0,
        "show_subtotal": False,
        "csv_file": "ALAE--Paid.csv",
        "created": "2026-08-01T09:00:00",
        "updated_at": "2026-08-19T10:05:30.500000",
        "user": "xw2781",
        "origin_count": 2,
        "method_type_code": 0,
        "data_format_code": 1,
        "formula": "a + b",
        "processing_by_csv": {"ALAE--Paid.csv": {"config_hash": "sha256:" + "b" * 16}},
        "Precedents": [
            {"dataset_name": "Paid Loss", "method_type": "", "path": r"E:\x\Paid Loss.csv", "mtime": 1.0},
            "Case Reserves",
        ],
        "Dependents": [{"dataset_type_name": "F 41 - BF Incurred", "method_type": "Bornhuetter Ferguson"}],
        "audit_log": [{"event_date": "2026-08-19T10:05:30", "action": "Insert", "change_info": "", "user": "xw2781"}],
    }


class SnakeKeyTests(unittest.TestCase):
    def test_spaces_hyphens_and_camel_case_all_become_underscores(self) -> None:
        self.assertEqual(snake_key("ratio basis origin labels"), "ratio_basis_origin_labels")
        self.assertEqual(snake_key("averageType"), "average_type")
        self.assertEqual(snake_key("audit-log-tab"), "audit_log_tab")
        self.assertEqual(snake_key("  Precedents  "), "precedents")


class UpgradeMethodTests(unittest.TestCase):
    def test_every_old_method_stamp_maps_to_a_v4_stamp(self) -> None:
        for old, new in METHOD_FORMAT_UPGRADES.items():
            with self.subTest(old=old):
                self.assertTrue(new.endswith("-v4"), new)

    def test_keys_are_renamed_at_every_depth_and_the_stamp_comes_first(self) -> None:
        upgraded, _ = upgrade_method(old_dfm_method())

        self.assertEqual(list(upgraded)[0], "json_format")
        self.assertEqual(upgraded["json_format"], DFM_JSON_FORMAT)
        self.assertEqual(upgraded["method_metadata"]["method_name"], "C 12")
        self.assertEqual(upgraded["method_metadata"]["average_type"], "Volume Weighted")
        self.assertIn("origin_labels", upgraded["ratios_tab"]["ratio_triangle"])

    def test_a_method_keeps_no_audit_log_and_no_empty_placeholder_sections(self) -> None:
        upgraded, _ = upgrade_method(old_dfm_method())

        self.assertNotIn("audit_log", upgraded)
        self.assertNotIn("chart_tab", upgraded)
        self.assertNotIn("audit_log_tab", upgraded)

    def test_the_dfm_ratio_labels_stay_and_the_forced_copies_go(self) -> None:
        upgraded, _ = upgrade_method(old_dfm_method())

        self.assertEqual(upgraded["ratios_tab"]["ratio_triangle"]["origin_labels"], ["2024", "2025"])
        self.assertNotIn("ratio_basis_origin_labels", upgraded["results_tab"])
        self.assertNotIn("input_data_triangle_mask", upgraded["data_tab"])

    def test_notes_are_handed_back_rather_than_dropped(self) -> None:
        upgraded, notes = upgrade_method(old_dfm_method())

        self.assertNotIn("notes_tab", upgraded)
        self.assertEqual(notes, "Tail based on a competitor study.")

    def test_timestamps_become_utc_with_milliseconds(self) -> None:
        upgraded, _ = upgrade_method(old_dfm_method())

        self.assertRegex(upgraded["method_metadata"]["last_modified"], r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}Z$")

    def test_an_already_v4_payload_passes_through(self) -> None:
        upgraded, _ = upgrade_method(old_dfm_method())
        again, notes = upgrade_method(upgraded)

        self.assertEqual(again, upgraded)
        self.assertEqual(notes, "")

    def test_an_unknown_stamp_is_refused_rather_than_guessed(self) -> None:
        payload = old_dfm_method()
        payload["json format"] = "arcrho-something-nobody-wrote-v9"

        with self.assertRaises(PersistedJsonUpgradeError):
            upgrade_method(payload)


class UnconvertibleMethodTests(unittest.TestCase):
    """The four BF files holding the only copy of real commentary are exactly
    the four the app already refused to open, so the notes rescue may not
    depend on converting them (Trap 1 of the plan)."""

    def bf_v2(self) -> dict:
        return {
            "json_format": "arcrho-bornhuetter-ferguson-method-by-tab-v2",
            "details_tab": {"name": "F 41 - BF Incurred"},
            "method_tab": {},
            "chart_tab": {},
            "notes_tab": {"notes": "The loss development pattern is based on the Adjusted Incurred method."},
            "audit_log_tab": {},
            "method_metadata": {"method_type": "Bornhuetter Ferguson", "last_modified": "2025-04-22T11:44:12"},
        }

    def test_a_format_retired_before_v4_is_named_rather_than_converted(self) -> None:
        with self.assertRaises(UnsupportedMethodFormatError) as raised:
            upgrade_method(self.bf_v2())

        self.assertIn("arcrho-bornhuetter-ferguson-method-by-tab-v2", str(raised.exception))

    def test_it_is_still_a_recognized_upgrade_failure(self) -> None:
        self.assertTrue(issubclass(UnsupportedMethodFormatError, PersistedJsonUpgradeError))
        self.assertTrue(UNCONVERTIBLE_METHOD_FORMATS.isdisjoint(METHOD_FORMAT_UPGRADES))

    def test_its_notes_are_still_readable(self) -> None:
        self.assertEqual(
            stranded_method_notes(self.bf_v2()),
            "The loss development pattern is based on the Adjusted Incurred method.",
        )

    def test_notes_are_read_whatever_the_key_spelling(self) -> None:
        self.assertEqual(stranded_method_notes({"notes tab": {"notes": " Authored by Kelly. "}}), "Authored by Kelly.")
        self.assertEqual(stranded_method_notes({"notes_tab": {}}), "")
        self.assertEqual(stranded_method_notes({}), "")


class UpgradeSidecarTests(unittest.TestCase):
    def test_the_graph_keys_lose_their_capitals_and_their_machine_local_fields(self) -> None:
        upgraded = upgrade_dataset_sidecar(old_sidecar())

        self.assertEqual(
            upgraded["precedents"],
            [{"dataset_name": "Paid Loss"}, {"dataset_name": "Case Reserves"}],
        )
        self.assertEqual(
            upgraded["dependents"],
            [{"dataset_name": "F 41 - BF Incurred", "method_type": "Bornhuetter Ferguson"}],
        )

    def test_no_retired_field_survives(self) -> None:
        upgraded = upgrade_dataset_sidecar(old_sidecar())

        self.assertEqual(RETIRED_SIDECAR_FIELDS & set(upgraded), set())

    def test_the_stamp_is_first_and_the_audit_log_is_last(self) -> None:
        keys = list(upgrade_dataset_sidecar(old_sidecar()))

        self.assertEqual(keys[0], "json_format")
        self.assertEqual(keys[-1], "audit_log")

    def test_the_retired_user_field_fills_a_missing_modified_by(self) -> None:
        self.assertEqual(upgrade_dataset_sidecar(old_sidecar())["modified_by"], "xw2781")

    def test_upgrading_twice_changes_nothing(self) -> None:
        once = upgrade_dataset_sidecar(old_sidecar())

        self.assertEqual(upgrade_dataset_sidecar(once), once)

    def test_upgrading_twice_writes_the_same_text_when_a_core_field_was_missing(self) -> None:
        """Comparing the payloads is not enough: two dicts holding the same
        pairs in a different order are equal, and it is the *text* that is
        written. A field the old file left out and the two graph fields are
        all appended, so filling the core after the graph put ``status``
        behind ``precedents`` on the first pass and in front of it on the
        second. 314 of the 2,079 sidecars in ``NJ_Annual_Prod_202605_Fake``
        take this path, and none of them was a fixed point."""

        payload = old_sidecar()
        payload.pop("status")

        once = persisted_json_text(upgrade_dataset_sidecar(payload))
        twice = persisted_json_text(upgrade_dataset_sidecar(json.loads(once)))

        self.assertEqual(twice, once)

    def test_the_fields_a_builder_writes_after_the_graph_are_written_there(self) -> None:
        """Being a fixed point of the conversion is not enough -- the order has
        to be the one the app writes, or a converted sidecar changes shape the
        first time somebody saves the method that owns it. Every canonical
        builder emits these five after ``dependents`` and before ``audit_log``
        (``dfm_contract.build_dfm_output_sidecar``); a live save proved the
        converter had been putting them in front of the graph instead."""

        keys = list(upgrade_dataset_sidecar(old_sidecar(), publication_revision="sha256:" + "c" * 16))
        after_graph = ("created", "updated_at", "modified_by", "status", "publication_revision")

        self.assertEqual(keys[-1], "audit_log")
        self.assertEqual(keys[-len(after_graph) - 1 : -1], list(after_graph))
        self.assertLess(keys.index("dependents"), keys.index("created"))

    def test_every_missing_core_field_survives_a_second_conversion_in_place(self) -> None:
        for field in ("method_type", "status", "show_subtotal"):
            with self.subTest(field=field):
                payload = old_sidecar()
                payload.pop(field)

                once = persisted_json_text(upgrade_dataset_sidecar(payload))
                twice = persisted_json_text(upgrade_dataset_sidecar(json.loads(once)))

                self.assertEqual(twice, once)

    def test_a_missing_core_field_is_filled_so_the_shared_validator_accepts_it(self) -> None:
        payload = old_sidecar()
        payload.pop("show_subtotal")

        self.assertIs(upgrade_dataset_sidecar(payload)["show_subtotal"], False)

    def test_a_sidecar_that_names_a_method_is_made_calculated(self) -> None:
        payload = dict(old_sidecar(), method_name="Quarterly DFM Claim Counts", method_type="DFM", calculated=False)

        self.assertIs(upgrade_dataset_sidecar(payload)["calculated"], True)

    def test_a_plain_dataset_keeps_its_own_calculated_flag(self) -> None:
        self.assertIs(upgrade_dataset_sidecar(old_sidecar())["calculated"], False)

    def test_the_converted_sidecar_satisfies_the_shared_core(self) -> None:
        self.assertEqual(
            validate_sidecar_core(upgrade_dataset_sidecar(old_sidecar())),
            upgrade_dataset_sidecar(old_sidecar()),
        )


class FingerprintLengthTests(unittest.TestCase):
    """Rule 2a: both sides of every comparison shorten together, so a stored
    full-length digest has to shorten with them or nothing matches again."""

    def test_a_stored_full_digest_keeps_its_first_sixteen_characters(self) -> None:
        payload = old_sidecar()
        payload["processing"] = {"config_hash": "sha256:" + "0123456789abcdef" + "f" * 48}

        upgraded = upgrade_dataset_sidecar(payload)

        self.assertEqual(upgraded["processing"]["config_hash"], "sha256:0123456789abcdef")

    def test_the_same_holds_inside_a_method_file(self) -> None:
        upgraded, _ = upgrade_method(old_dfm_method())

        self.assertEqual(upgraded["data_tab"]["source_revision"], "sha256:" + "a" * 16)

    def test_a_value_already_short_is_left_alone(self) -> None:
        payload = old_sidecar()
        payload["processing"] = {"config_hash": "sha256:0123456789abcdef"}

        self.assertEqual(
            upgrade_dataset_sidecar(payload)["processing"]["config_hash"],
            "sha256:0123456789abcdef",
        )

    def test_text_that_only_looks_like_a_digest_is_left_alone(self) -> None:
        payload = old_sidecar()
        payload["notes"] = "sha256:not-hex-at-all"

        self.assertEqual(upgrade_dataset_sidecar(payload)["notes"], "sha256:not-hex-at-all")


class PublicationRevisionHandoffTests(unittest.TestCase):
    """The one fingerprint a conversion cannot shorten: the hash vocabulary
    stopped depending on the persisted key spelling in step 1, so a v4 method
    computes a different number and the sidecar must be given that one."""

    def test_the_value_from_the_converted_method_wins(self) -> None:
        payload = dict(old_sidecar(), publication_revision="sha256:" + "8" * 64)

        upgraded = upgrade_dataset_sidecar(payload, publication_revision="sha256:69781191c264d46f")

        self.assertEqual(upgraded["publication_revision"], "sha256:69781191c264d46f")

    def test_a_plain_dataset_sidecar_is_unaffected(self) -> None:
        upgraded = upgrade_dataset_sidecar(old_sidecar(), publication_revision="")

        self.assertNotIn("publication_revision", upgraded)


class MigratedNotesTests(unittest.TestCase):
    def test_notes_from_a_method_land_in_the_sidecar(self) -> None:
        sidecar = upgrade_dataset_sidecar(old_sidecar())

        merged = sidecar_with_method_notes(sidecar, "Tail based on a competitor study.")

        self.assertEqual(merged["notes"], "Tail based on a competitor study.")
        self.assertEqual(list(merged)[-1], "audit_log")

    def test_text_already_in_the_sidecar_is_kept_and_the_incoming_text_follows(self) -> None:
        sidecar = dict(upgrade_dataset_sidecar(old_sidecar()), notes="Reviewed by the actuary.")

        merged = sidecar_with_method_notes(sidecar, "Tail based on a competitor study.")

        self.assertEqual(merged["notes"], "Reviewed by the actuary.\n\nTail based on a competitor study.")

    def test_converting_the_same_workspace_twice_does_not_duplicate_a_note(self) -> None:
        sidecar = upgrade_dataset_sidecar(old_sidecar())

        once = sidecar_with_method_notes(sidecar, "Tail based on a competitor study.")
        twice = sidecar_with_method_notes(once, "Tail based on a competitor study.")

        self.assertEqual(twice["notes"], once["notes"])

    def test_a_method_with_no_notes_leaves_the_sidecar_alone(self) -> None:
        sidecar = upgrade_dataset_sidecar(old_sidecar())

        self.assertEqual(sidecar_with_method_notes(sidecar, ""), sidecar)

    def test_a_rescued_note_goes_in_front_of_the_graph_not_on_the_end(self) -> None:
        """A sidecar with no notes field yet gets one where a canonical builder
        writes it. Appended instead, it would sit behind ``precedents`` the
        first time the file is converted and in front of it the second, and
        the three sidecars that receive the only surviving copy of a retired
        method's commentary would never be fixed points."""

        merged = sidecar_with_method_notes(upgrade_dataset_sidecar(old_sidecar()), "Tail from a study.")
        keys = list(merged)

        self.assertLess(keys.index("notes"), keys.index("precedents"))

    def test_rescuing_the_same_note_twice_writes_the_same_text(self) -> None:
        note = "Tail based on a competitor study."

        def convert(payload: dict) -> str:
            return persisted_json_text(sidecar_with_method_notes(upgrade_dataset_sidecar(payload), note))

        once = convert(old_sidecar())

        self.assertEqual(convert(json.loads(once)), once)


class UpgradeProjectFilesTests(unittest.TestCase):
    def test_the_project_log_adopts_the_sidecar_record_shape(self) -> None:
        upgraded = upgrade_project_audit_log({
            "project_name": "NJ",
            "updated_at": "2026-08-19T10:05:30",
            "entries": [
                {"timestamp": "2026-08-19T10:05:30", "action": "Renamed the reserving class", "user": "xw2781"},
                {"timestamp": "2026-08-19T11:05:30", "action": "Insert", "user": "xw2781"},
            ],
        })

        self.assertEqual(upgraded["json_format"], PROJECT_AUDIT_LOG_JSON_FORMAT)
        self.assertEqual(list(upgraded)[-1], "audit_log")
        self.assertNotIn("entries", upgraded)
        free_text, known = upgraded["audit_log"]
        self.assertEqual(free_text["action"], "Update")
        self.assertEqual(free_text["change_info"], "Renamed the reserving class")
        self.assertEqual(known["action"], "Insert")
        self.assertEqual(known["change_info"], "")

    def test_cache_provenance_renames_its_version_key(self) -> None:
        upgraded = upgrade_runtime_cache_provenance({
            "format": "arcrho-runtime-cache-provenance-v1",
            "csv_fingerprint": {"mtime_ns": 17, "sha256": "c" * 64},
            "processing": {"config_hash": "sha256:" + "0123456789abcdef" + "d" * 48},
        })

        self.assertEqual(list(upgraded)[0], "json_format")
        self.assertEqual(upgraded["json_format"], RUNTIME_CACHE_PROVENANCE_JSON_FORMAT)
        self.assertNotIn("format", upgraded)
        # The fingerprint of the cached file beside it is meant to be local,
        # and is not one of the persisted fingerprints rule 2a shortens.
        self.assertEqual(upgraded["csv_fingerprint"]["mtime_ns"], 17)
        self.assertEqual(upgraded["csv_fingerprint"]["sha256"], "c" * 64)
        self.assertEqual(upgraded["processing"]["config_hash"], "sha256:0123456789abcdef")

    def test_the_number_formats_file_is_restamped(self) -> None:
        upgraded = upgrade_dataset_number_formats({"json_format": "arcrho.dataset-number-formats.v1", "datasets": {}})

        self.assertEqual(upgraded["json_format"], DATASET_NUMBER_FORMATS_JSON_FORMAT)

    def test_the_source_import_file_trades_version_for_a_stamp(self) -> None:
        upgraded = upgrade_source_import({"version": 1, "last_import": {"csv_path": r"C:\in\loss.csv"}})

        self.assertEqual(upgraded["json_format"], SOURCE_IMPORT_JSON_FORMAT)
        self.assertNotIn("version", upgraded)
        # The external source path is identity, not a workspace location.
        self.assertEqual(upgraded["last_import"]["csv_path"], r"C:\in\loss.csv")


if __name__ == "__main__":
    unittest.main()
