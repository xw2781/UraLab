from __future__ import annotations

import copy
import sys
import tempfile
import unittest
from pathlib import Path


_PYTHON_API_ROOT = Path(__file__).resolve().parents[1]
_REPOSITORY_ROOT = _PYTHON_API_ROOT.parent
_TMP_ROOT = Path(__file__).resolve().parent / "logs" / "tmp"
for _path in (_PYTHON_API_ROOT / "src", _PYTHON_API_ROOT / "migration"):
    if str(_path) not in sys.path:
        sys.path.insert(0, str(_path))

from resq_migration import sync


def _item(
    name: str = "Paid Loss",
    *,
    timestamp: float | None = 100.0,
    kind: str = "Dataset",
    data_format: str = "Triangle",
    dataset_type: str = "Paid Loss",
    method_name: str = "",
) -> dict:
    return {
        "name": name,
        "kind": kind,
        "data_format": data_format,
        "dataset_type": dataset_type,
        "method_name": method_name,
        "modified_timestamp": timestamp,
        "can_export_to_resq": True,
        "can_import_to_arcrho": True,
        "can_receive_from_arcrho": True,
    }


def _baseline(*, arcrho_timestamp: float = 100.0, resq_timestamp: float = 100.0) -> dict:
    entry = {
        "arcrho_present": True,
        "resq_present": True,
        "arcrho_timestamp": arcrho_timestamp,
        "resq_timestamp": resq_timestamp,
    }
    return {"items": {"paid loss": entry}}


class ResqSyncPlanTests(unittest.TestCase):
    def test_raw_timestamp_comparison_equal_and_unknown_are_fail_closed(self):
        cases = (
            (
                [_item(timestamp=200)],
                [_item(timestamp=100)],
                "ArcRho newer",
                sync.ACTION_ARCRHO_TO_RESQ,
                True,
            ),
            (
                [_item(timestamp=100)],
                [_item(timestamp=200)],
                "ResQ newer",
                sync.ACTION_RESQ_TO_ARCRHO,
                True,
            ),
            (
                [_item(timestamp=100)],
                [_item(timestamp=100)],
                "Same timestamp",
                "",
                False,
            ),
            (
                [_item(timestamp=None)],
                [_item(timestamp=100)],
                "Unknown timestamp",
                "",
                False,
            ),
        )

        for local, remote, status, action, selected in cases:
            with self.subTest(status=status, action=action):
                row = sync.build_sync_plan(local, remote)[0]
                self.assertEqual(row["status"], status)
                self.assertEqual(row["action"], action)
                self.assertEqual(row["selected"], selected)
                self.assertEqual(row["disabled"], not bool(action))

    def test_recorded_baseline_prevents_a_resq_save_timestamp_from_ping_ponging(self):
        state = sync.empty_sync_state("Demo", r"Auto\PP", "Demo")
        state = sync.record_synced_items(
            state,
            ["Paid Loss"],
            [_item(timestamp=100)],
            [_item(timestamp=110)],
            synced_at="2026-08-12T12:00:00+00:00",
        )

        unchanged = sync.build_sync_plan(
            [_item(timestamp=100)],
            [_item(timestamp=110)],
            state,
        )[0]
        self.assertEqual(unchanged["status"], "Synchronized")
        self.assertEqual(unchanged["action"], "")

        local_changed = sync.build_sync_plan(
            [_item(timestamp=120)],
            [_item(timestamp=110)],
            state,
        )[0]
        self.assertEqual(local_changed["status"], "ArcRho changed")
        self.assertEqual(local_changed["action"], sync.ACTION_ARCRHO_TO_RESQ)

    def _baselined_state(self):
        return sync.record_synced_items(
            sync.empty_sync_state("Demo", r"Auto\PP", "ResQ Demo"),
            ["Paid Loss"],
            [_item(timestamp=200)],
            [_item(timestamp=100)],
            synced_at="2026-08-12T12:00:00+00:00",
        )

    def test_a_batch_ripple_is_baselined_only_on_the_side_it_moved(self):
        # Saving a DFM into ResQ makes ResQ recalculate and re-stamp the Result
        # Selection downstream of it; the next review used to call that "ResQ
        # changed" and offer to pull back a copy nobody edited.
        state = self._baselined_state()
        before = sync.build_sync_plan([_item(timestamp=200)], [_item(timestamp=100)], state)
        after = sync.build_sync_plan([_item(timestamp=200)], [_item(timestamp=150)], state)
        self.assertEqual(after[0]["status"], "ResQ changed")

        updated, absorbed = sync.absorb_propagated_changes(
            state, before, after, keys=["paid loss"], synced_at="2026-08-12T12:05:00+00:00"
        )

        self.assertEqual(
            absorbed,
            [{"key": "paid loss", "name": "Paid Loss", "kind": "Dataset", "sides": ["resq"]}],
        )
        entry = updated["items"]["paid loss"]
        self.assertEqual((entry["arcrho_timestamp"], entry["resq_timestamp"]), (200.0, 150.0))
        self.assertEqual(entry["propagated_at"], "2026-08-12T12:05:00+00:00")
        self.assertEqual(entry["synced_at"], "2026-08-12T12:00:00+00:00")
        replan = sync.build_sync_plan([_item(timestamp=200)], [_item(timestamp=150)], updated)
        self.assertEqual(replan[0]["status"], "Synchronized")

    def test_a_change_pending_before_the_batch_survives_the_ripple(self):
        state = self._baselined_state()
        before = sync.build_sync_plan([_item(timestamp=300)], [_item(timestamp=100)], state)
        after = sync.build_sync_plan([_item(timestamp=300)], [_item(timestamp=150)], state)
        self.assertEqual(before[0]["status"], "ArcRho changed")

        updated, absorbed = sync.absorb_propagated_changes(state, before, after, keys=["paid loss"])

        self.assertEqual([item["sides"] for item in absorbed], [["resq"]])
        entry = updated["items"]["paid loss"]
        self.assertEqual((entry["arcrho_timestamp"], entry["resq_timestamp"]), (200.0, 150.0))
        replan = sync.build_sync_plan([_item(timestamp=300)], [_item(timestamp=150)], updated)
        self.assertEqual(replan[0]["status"], "ArcRho changed")

    def test_a_row_without_a_baseline_is_baselined_only_from_matching_timestamps(self):
        empty = sync.empty_sync_state("Demo", r"Auto\PP", "ResQ Demo")
        same_before = sync.build_sync_plan([_item(timestamp=200)], [_item(timestamp=200)])
        same_after = sync.build_sync_plan([_item(timestamp=200)], [_item(timestamp=250)])
        self.assertEqual(same_before[0]["status"], "Same timestamp")

        updated, absorbed = sync.absorb_propagated_changes(
            empty, same_before, same_after, keys=["paid loss"], synced_at="2026-08-12T12:05:00+00:00"
        )

        self.assertEqual([item["sides"] for item in absorbed], [["resq"]])
        entry = updated["items"]["paid loss"]
        self.assertEqual((entry["arcrho_timestamp"], entry["resq_timestamp"]), (200.0, 250.0))
        self.assertEqual(entry["synced_at"], "2026-08-12T12:05:00+00:00")
        replan = sync.build_sync_plan([_item(timestamp=200)], [_item(timestamp=250)], updated)
        self.assertEqual(replan[0]["status"], "Synchronized")

        pending_before = sync.build_sync_plan([_item(timestamp=300)], [_item(timestamp=200)])
        pending_after = sync.build_sync_plan([_item(timestamp=300)], [_item(timestamp=250)])
        untouched, absorbed = sync.absorb_propagated_changes(
            empty, pending_before, pending_after, keys=["paid loss"]
        )
        self.assertEqual(absorbed, [])
        self.assertEqual(untouched["items"], {})

    def test_a_row_that_held_still_is_left_alone(self):
        state = self._baselined_state()
        plan = sync.build_sync_plan([_item(timestamp=200)], [_item(timestamp=100)], state)

        updated, absorbed = sync.absorb_propagated_changes(state, plan, plan, keys=["paid loss", "missing"])

        self.assertEqual(absorbed, [])
        self.assertEqual(updated["items"], state["items"])

    def test_items_on_one_side_only_never_become_rows(self):
        plan = sync.build_sync_plan(
            [_item("Paid Loss"), _item("Only Here"), _item("Twice Here"), _item("Twice Here")],
            [_item("Paid Loss"), _item("Only There")],
        )

        self.assertEqual([row["name"] for row in plan], ["Paid Loss"])

    def test_an_interrupted_run_leaves_the_next_comparison_to_the_recorded_baseline(self):
        row = sync.build_sync_plan(
            [_item(timestamp=200)],
            [_item(timestamp=100)],
            _baseline(),
        )[0]

        self.assertEqual(row["status"], "ArcRho changed")
        self.assertEqual(row["action"], sync.ACTION_ARCRHO_TO_RESQ)
        self.assertFalse(row["review"])
        self.assertTrue(row["selected"])

    def test_both_changed_rides_with_the_reserving_class_and_is_marked_for_review(self):
        row = sync.build_sync_plan(
            [_item(timestamp=150)],
            [_item(timestamp=160)],
            _baseline(),
        )[0]

        self.assertEqual(row["status"], "Both changed")
        self.assertEqual(row["action"], sync.ACTION_RESQ_TO_ARCRHO)
        self.assertTrue(row["review"])
        self.assertIn("overwrites this ArcRho change", row["detail"])
        self.assertTrue(row["selected"])
        self.assertFalse(row["disabled"])

        equal = sync.build_sync_plan(
            [_item(timestamp=150)],
            [_item(timestamp=150)],
            _baseline(),
        )[0]
        self.assertEqual(equal["status"], "Both changed")
        self.assertEqual(equal["action"], "")
        self.assertFalse(equal["review"])
        self.assertTrue(equal["disabled"])

    def test_the_latest_timestamp_on_each_side_decides_one_direction_for_every_row(self):
        plan = sync.build_sync_plan(
            [_item("Paid Loss", timestamp=300), _item("Incurred Loss", timestamp=100)],
            [_item("Paid Loss", timestamp=100), _item("Incurred Loss", timestamp=200)],
        )

        self.assertEqual(
            sync.plan_direction(plan),
            {"direction": sync.ACTION_ARCRHO_TO_RESQ, "arcrho_timestamp": 300.0, "resq_timestamp": 200.0},
        )
        by_name = {row["name"]: row for row in plan}
        agreeing = by_name["Paid Loss"]
        self.assertEqual((agreeing["status"], agreeing["action"]), ("ArcRho newer", sync.ACTION_ARCRHO_TO_RESQ))
        self.assertFalse(agreeing["review"])
        # The row's own timestamps point the other way: it is still pushed
        # with the class, ticked, but marked so the person reads it first.
        disagreeing = by_name["Incurred Loss"]
        self.assertEqual((disagreeing["status"], disagreeing["action"]), ("ResQ newer", sync.ACTION_ARCRHO_TO_RESQ))
        self.assertTrue(disagreeing["review"])
        self.assertTrue(disagreeing["selected"])
        self.assertFalse(disagreeing["disabled"])
        self.assertIn("ArcRho copy overwrites this ResQ change", disagreeing["detail"])

    def test_matching_or_unknown_latest_timestamps_give_no_direction(self):
        matching = sync.build_sync_plan(
            [_item("Paid Loss", timestamp=300), _item("Incurred Loss", timestamp=100)],
            [_item("Paid Loss", timestamp=100), _item("Incurred Loss", timestamp=300)],
        )
        self.assertEqual(sync.plan_direction(matching)["direction"], "")
        self.assertEqual([row["action"] for row in matching], ["", ""])
        self.assertEqual([row["status"] for row in matching], ["ResQ newer", "ArcRho newer"])
        self.assertTrue(all(row["disabled"] for row in matching))

        unknown = sync.build_sync_plan([_item(timestamp=None)], [_item(timestamp=100)])
        self.assertEqual(
            sync.plan_direction(unknown),
            {"direction": "", "arcrho_timestamp": None, "resq_timestamp": 100.0},
        )

    def test_type_format_dataset_type_and_method_mismatches_do_not_offer_actions(self):
        cases = (
            (
                _item(kind="Dataset"),
                _item(kind="DFM", data_format="Vector", method_name="Selected DFM"),
                "Type mismatch",
            ),
            (
                _item(data_format="Triangle"),
                _item(data_format="Vector"),
                "Format mismatch",
            ),
            (
                _item(dataset_type="Paid Loss"),
                _item(dataset_type="Incurred Loss"),
                "Dataset Type mismatch",
            ),
            (
                _item(
                    kind="DFM",
                    data_format="Vector",
                    dataset_type="Ultimate Loss",
                    method_name="Selected DFM",
                ),
                _item(
                    kind="DFM",
                    data_format="Vector",
                    dataset_type="Ultimate Loss",
                    method_name="Alternative DFM",
                ),
                "Method mismatch",
            ),
        )

        for local, remote, status in cases:
            with self.subTest(status=status):
                row = sync.build_sync_plan([local], [remote])[0]
                self.assertEqual(row["status"], status)
                self.assertEqual(row["action"], "")
                self.assertTrue(row["disabled"])
                self.assertFalse(row["selected"])

    def test_state_round_trip_is_scoped_atomic_and_omits_transient_fields(self):
        _TMP_ROOT.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(dir=_TMP_ROOT) as temp_name:
            root = Path(temp_name)
            state_path = sync.sync_state_path(root, "Demo", r"Auto\PP", "ResQ Demo")
            state = sync.empty_sync_state("Demo", r"Auto\PP", "ResQ Demo")
            state = sync.record_synced_items(
                state,
                ["Paid Loss"],
                [_item(timestamp=101)],
                [_item(timestamp=202)],
                synced_at="2026-08-12T12:00:00+00:00",
            )
            self.assertEqual(state["_recorded_keys"], ["paid loss"])

            written = sync.write_sync_state(state_path, state)
            loaded = sync.read_sync_state(
                written,
                "Demo",
                r"Auto\PP",
                "ResQ Demo",
            )

            self.assertEqual(loaded["items"]["paid loss"]["arcrho_timestamp"], 101.0)
            self.assertEqual(loaded["items"]["paid loss"]["resq_timestamp"], 202.0)
            self.assertNotIn("_recorded_keys", loaded)
            self.assertTrue(written.read_bytes().endswith(b"\n"))
            self.assertEqual(list(written.parent.glob(f".{written.name}.*.tmp")), [])
            self.assertNotIn(r"Auto\PP", str(written))
            self.assertNotIn("ResQ Demo", str(written))

    def test_plan_signatures_detect_stale_observations(self):
        row = sync.build_sync_plan(
            [_item(timestamp=200)],
            [_item(timestamp=100)],
        )[0]
        original = sync.plan_signature(row)
        unchanged = copy.deepcopy(original)
        changed = copy.deepcopy(original)
        changed["resq"]["modified_timestamp"] = 101

        self.assertTrue(sync.signatures_equal(original, unchanged))
        self.assertFalse(sync.signatures_equal(original, changed))

        # Inside a write batch only the side being written from has to hold
        # still; the target side is re-stamped by the batch's earlier writes.
        self.assertTrue(sync.write_signatures_equal(original, changed, source_side="arcrho"))
        self.assertFalse(sync.write_signatures_equal(original, changed, source_side="resq"))
        renamed = copy.deepcopy(original)
        renamed["resq"]["dataset_type"] = "Other"
        self.assertFalse(sync.write_signatures_equal(original, renamed, source_side="arcrho"))
        with self.assertRaises(ValueError):
            sync.write_signatures_equal(original, changed, source_side="elsewhere")

        recorded_state = sync.record_synced_items(
            sync.empty_sync_state("Demo", r"Auto\PP", "ResQ Demo"),
            ["Paid Loss"],
            [_item(timestamp=200)],
            [_item(timestamp=100)],
            synced_at="2026-08-12T12:00:00+00:00",
        )
        baselined_row = sync.build_sync_plan(
            [_item(timestamp=200)],
            [_item(timestamp=100)],
            recorded_state,
        )[0]
        self.assertFalse(
            sync.signatures_equal(original, sync.plan_signature(baselined_row))
        )

    def test_newer_side_names_the_side_modified_last_or_nothing(self):
        self.assertEqual(sync.newer_side(_item(timestamp=200), _item(timestamp=100)), "arcrho")
        self.assertEqual(sync.newer_side(_item(timestamp=100), _item(timestamp=200)), "resq")
        self.assertEqual(sync.newer_side(_item(timestamp=100), _item(timestamp=100)), "")
        self.assertEqual(sync.newer_side(_item(timestamp=None), _item(timestamp=100)), "")
        self.assertEqual(sync.newer_side({}, {}), "")

    def test_changed_since_baseline_names_the_sides_edited_since_the_saved_pair(self):
        entry = _baseline(arcrho_timestamp=100, resq_timestamp=150)["items"]["paid loss"]

        self.assertEqual(
            sync.changed_since_baseline(_item(timestamp=100), _item(timestamp=150), entry),
            sync.CHANGED_NEITHER,
        )
        self.assertEqual(
            sync.changed_since_baseline(_item(timestamp=200), _item(timestamp=150), entry),
            sync.CHANGED_ARCRHO,
        )
        self.assertEqual(
            sync.changed_since_baseline(_item(timestamp=100), _item(timestamp=300), entry),
            sync.CHANGED_RESQ,
        )
        self.assertEqual(
            sync.changed_since_baseline(_item(timestamp=200), _item(timestamp=300), entry),
            sync.CHANGED_BOTH,
        )

    def test_a_pair_with_matching_timestamps_is_synchronized_whatever_the_baseline_says(self):
        # An import copies ResQ over ArcRho and stamps both with the same time
        # without recording a baseline, so the pair the last export saved
        # would otherwise report an edit on both sides.
        entry = _baseline(arcrho_timestamp=100, resq_timestamp=150)["items"]["paid loss"]

        self.assertEqual(
            sync.changed_since_baseline(_item(timestamp=300), _item(timestamp=300), entry),
            sync.CHANGED_NEITHER,
        )
        review = sync.export_review(_item(timestamp=300), _item(timestamp=300), entry)
        self.assertEqual(review["changed"], sync.CHANGED_NEITHER)
        self.assertFalse(review["overwrites_edit"])
        self.assertEqual(review["status"], "Synchronized")

    def test_changed_since_baseline_is_blank_when_no_pair_can_be_measured_against(self):
        entry = _baseline()["items"]["paid loss"]
        incomplete = dict(entry, resq_timestamp=None)

        self.assertEqual(sync.changed_since_baseline(_item(), _item(), None), "")
        self.assertEqual(sync.changed_since_baseline(_item(), _item(), {}), "")
        self.assertEqual(
            sync.changed_since_baseline(_item(), _item(), {"present": False, **entry}), ""
        )
        self.assertEqual(sync.changed_since_baseline(_item(), _item(), incomplete), "")

    def test_the_export_review_warns_only_about_a_resq_edit_made_since_the_baseline(self):
        # The stamp the last export left on ResQ is newer than ArcRho's, which
        # is exactly the case the raw comparison used to call an overwrite.
        arcrho = _item(timestamp=100)
        resq = _item(timestamp=150)
        entry = _baseline(arcrho_timestamp=100, resq_timestamp=150)["items"]["paid loss"]

        settled = sync.export_review(arcrho, resq, entry)
        self.assertEqual(settled["changed"], sync.CHANGED_NEITHER)
        self.assertFalse(settled["overwrites_edit"])
        self.assertEqual(settled["status"], "Synchronized")
        self.assertEqual(
            settled["detail"], "Neither side has changed since the two were last synchronized."
        )

        edited = sync.export_review(arcrho, _item(timestamp=300), entry)
        self.assertEqual(edited["changed"], sync.CHANGED_RESQ)
        self.assertTrue(edited["overwrites_edit"])
        self.assertIn("overwrites that change", edited["detail"])

        both = sync.export_review(_item(timestamp=200), _item(timestamp=300), entry)
        self.assertEqual(both["changed"], sync.CHANGED_BOTH)
        self.assertTrue(both["overwrites_edit"])

        pushed = sync.export_review(_item(timestamp=200), resq, entry)
        self.assertEqual(pushed["changed"], sync.CHANGED_ARCRHO)
        self.assertFalse(pushed["overwrites_edit"])

    def test_the_export_review_falls_back_to_the_timestamps_until_a_baseline_exists(self):
        resq_newer = sync.export_review(_item(timestamp=100), _item(timestamp=200), None)
        self.assertEqual(resq_newer["changed"], "")
        self.assertTrue(resq_newer["overwrites_edit"])
        self.assertEqual(resq_newer["status"], "ResQ newer")
        self.assertIn("No baseline is recorded yet", resq_newer["detail"])

        arcrho_newer = sync.export_review(_item(timestamp=200), _item(timestamp=100), None)
        self.assertEqual(arcrho_newer["status"], "ArcRho newer")
        self.assertFalse(arcrho_newer["overwrites_edit"])

        self.assertEqual(
            sync.export_review(_item(timestamp=100), _item(timestamp=100), None)["status"],
            "Same timestamp",
        )
        self.assertEqual(
            sync.export_review(_item(timestamp=None), _item(timestamp=100), None)["status"],
            "Unknown timestamp",
        )

    def test_the_export_review_carries_the_reason_an_item_cannot_be_pushed(self):
        calculated = _item()
        calculated["can_receive_from_arcrho"] = False
        calculated["receive_block_reason"] = "ResQ recalculates this dataset."

        review = sync.export_review(_item(), calculated, None)

        self.assertFalse(review["supported"])
        self.assertTrue(review["detail"].endswith("ResQ recalculates this dataset."))

    def test_export_supported_follows_the_arcrho_to_resq_support_rule(self):
        self.assertTrue(sync.export_supported(_item(), _item()))
        blocked = _item()
        blocked["can_export_to_resq"] = False
        self.assertFalse(sync.export_supported(blocked, _item()))
        calculated = _item()
        calculated["can_receive_from_arcrho"] = False
        self.assertFalse(sync.export_supported(_item(), calculated))
        self.assertFalse(sync.export_supported(None, None))


if __name__ == "__main__":
    unittest.main()
