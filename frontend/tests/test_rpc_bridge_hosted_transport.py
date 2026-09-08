"""Gateway transport for the DFM RPC bridge routes.

The ArcRho Bridge that answers these requests runs on the server host, where
the request folder and the method files are local disk; only the Client PC half
of the exchange crosses SMB. These tests pin the wiring that moves that half to
the Gateway: the registered kinds, the arguments each route sends, that a
hosted operation is never re-run locally, and that the exchange stamps the
person who asked rather than whatever account owns the process.
"""

from __future__ import annotations

import inspect
import json
import os
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

FRONTEND_ROOT = Path(__file__).resolve().parents[1]
API_SOURCE = FRONTEND_ROOT.parent / "python-api" / "src"
for path in (FRONTEND_ROOT, API_SOURCE):
    if str(path) not in sys.path:
        sys.path.insert(0, str(path))

from arcrho_workspace_mutation_contract import (
    MAX_RPC_BRIDGE_WAIT_SECONDS,
    MIN_RPC_BRIDGE_WAIT_SECONDS,
    WORKSPACE_MUTATION_KINDS,
    WorkspaceMutationContractError,
    clamp_rpc_bridge_wait,
    validate_workspace_mutation_request,
)
from arcrho_workspace_read_contract import WORKSPACE_READ_KINDS

import app_server.api  # noqa: F401  (registers the route submodules)

dfm_rpc_bridge_router = sys.modules["app_server.api.dfm_rpc_bridge_router"]

from app_server.schemas.dfm_rpc_bridge import (
    DfmRpcBridgeRequest,
    DfmRpcBridgeUpdateRemoteRequest,
)
from app_server.services import dfm_rpc_bridge_service


DFM_KINDS = {
    "dfm_rpc_bridge_sync": "hosted_send_sync_request",
    "dfm_rpc_bridge_keep_local": "hosted_keep_local",
    "dfm_rpc_bridge_cleanup": "hosted_cleanup_tmp",
    "dfm_rpc_bridge_update_remote": "hosted_update_remote",
}


def dfm_request(**overrides):
    fields = dict(
        project_name="Demo",
        reserving_class="COL",
        method_name="M",
        output_vector="Out",
        input_triangle="Paid",
        origin_length=12,
        development_length=12,
        decimal_places=4,
        timeout_sec=8.0,
    )
    fields.update(overrides)
    return DfmRpcBridgeRequest(**fields)


class RegistrationTests(unittest.TestCase):
    """The contract is the only table; a kind must name a real entry point."""

    def test_every_rpc_bridge_kind_names_its_service_function(self) -> None:
        for kind, function in DFM_KINDS.items():
            spec = WORKSPACE_MUTATION_KINDS[kind]
            self.assertEqual(spec.function, function)
            module = dfm_rpc_bridge_service
            self.assertEqual(spec.module, module.__name__.rsplit(".", 1)[-1])
            self.assertTrue(callable(getattr(module, spec.function)))

        spec = WORKSPACE_READ_KINDS["dfm_rpc_bridge_compare"]
        self.assertEqual(spec.function, "hosted_compare")
        self.assertTrue(callable(getattr(dfm_rpc_bridge_service, spec.function)))

    def test_registered_arguments_are_exactly_the_route_schema_fields(self) -> None:
        # The hosted entry point rebuilds the request model from these, so a
        # kwarg the schema does not define would only fail on the Gateway.
        for kind in list(DFM_KINDS) + ["dfm_rpc_bridge_compare"]:
            spec = WORKSPACE_MUTATION_KINDS.get(kind) or WORKSPACE_READ_KINDS[kind]
            fields = set(DfmRpcBridgeRequest.model_fields)
            if kind.endswith("update_remote"):
                fields.add("rpc_server_write_confirmed")
            self.assertEqual(spec.allowed, fields, kind)

    def test_an_omitted_optional_argument_defers_to_the_route_schema(self) -> None:
        # The hosted signature must not restate a default the schema owns: an
        # optional argument left out arrives as None and is dropped, so both
        # transports fill it in from the same place.
        for kind, function in DFM_KINDS.items():
            spec = WORKSPACE_MUTATION_KINDS[kind]
            parameters = inspect.signature(getattr(dfm_rpc_bridge_service, function)).parameters
            self.assertEqual(set(parameters), set(spec.allowed), kind)
            for name in spec.optional:
                self.assertIsNone(parameters[name].default, f"{kind}: {name}")
            for name in spec.required:
                self.assertIs(parameters[name].default, inspect.Parameter.empty, f"{kind}: {name}")

    def test_a_hosted_call_without_the_optional_arguments_uses_the_schema_defaults(self) -> None:
        required = {
            name: getattr(dfm_request(), name)
            for name in WORKSPACE_MUTATION_KINDS["dfm_rpc_bridge_sync"].required
        }
        with patch.object(dfm_rpc_bridge_service, "send_sync_request", return_value={"ok": True}) as target:
            dfm_rpc_bridge_service.hosted_send_sync_request(**required)
        passed = target.call_args.args[0]
        self.assertEqual(passed.timeout_sec, DfmRpcBridgeRequest.model_fields["timeout_sec"].default)
        self.assertEqual(passed.decimal_places, DfmRpcBridgeRequest.model_fields["decimal_places"].default)

    def test_the_confirmation_flag_reaches_the_service_rather_than_the_contract(self) -> None:
        # A false flag must produce the route's own refusal on both transports,
        # so the contract treats it as optional and the service enforces it.
        request = validate_workspace_mutation_request(
            {
                "Function": "ArcRhoWorkspaceMutation",
                "ContractVersion": 1,
                "RequestId": "a" * 32,
                "MutationKind": "dfm_rpc_bridge_update_remote",
                "Kwargs": {**dfm_request().model_dump(), "rpc_server_write_confirmed": False},
                "UserName": "xwei",
            }
        )
        self.assertFalse(request["Kwargs"]["rpc_server_write_confirmed"])

    def test_the_wait_is_clamped_into_the_hosted_range(self) -> None:
        self.assertEqual(clamp_rpc_bridge_wait(8.0), 8.0)
        self.assertEqual(clamp_rpc_bridge_wait(0), MIN_RPC_BRIDGE_WAIT_SECONDS)
        self.assertEqual(clamp_rpc_bridge_wait(10_000), MAX_RPC_BRIDGE_WAIT_SECONDS)
        with self.assertRaises(WorkspaceMutationContractError):
            clamp_rpc_bridge_wait("soon")


class _CaptureMutation:
    """Stand in for run_workspace_mutation and record what a route asked for."""

    def __init__(self, payload: dict | None = None, *, remote: bool = False) -> None:
        self.calls: list[tuple[str, dict]] = []
        self.payload = payload if payload is not None else {"ok": True}
        self.remote = remote

    def __call__(self, mutation_kind, kwargs, *, local):
        self.calls.append((mutation_kind, dict(kwargs)))
        return dict(self.payload) if self.remote else local()


class _CaptureRead(_CaptureMutation):
    def __call__(self, read_kind, kwargs, *, local, finalize=None):
        self.calls.append((read_kind, dict(kwargs)))
        if self.remote:
            payload = dict(self.payload)
            return finalize(payload) if finalize is not None else payload
        return local()


class RouteWiringTests(unittest.TestCase):
    def _assert_registered(self, capture, registry) -> None:
        for kind, kwargs in capture.calls:
            spec = registry[kind]
            self.assertTrue(set(kwargs) <= spec.allowed, f"{kind}: {sorted(kwargs)}")
            for name in spec.required:
                self.assertIn(name, kwargs)

    def test_dfm_routes_name_their_kinds_and_run_the_service_locally(self) -> None:
        mutations = _CaptureMutation()
        reads = _CaptureRead()
        request = dfm_request()
        confirmed = DfmRpcBridgeUpdateRemoteRequest(
            **{**request.model_dump(), "rpc_server_write_confirmed": True}
        )
        with (
            patch.object(dfm_rpc_bridge_router.workspace_mutation_client, "run_workspace_mutation", mutations),
            patch.object(dfm_rpc_bridge_router.workspace_read_client, "run_workspace_read", reads),
            patch.object(dfm_rpc_bridge_router.dfm_rpc_bridge_service, "send_sync_request", return_value={"ok": True}) as sync,
            patch.object(dfm_rpc_bridge_router.dfm_rpc_bridge_service, "compare", return_value={"ok": True}) as compare,
            patch.object(dfm_rpc_bridge_router.dfm_rpc_bridge_service, "keep_local", return_value={"ok": True}) as keep,
            patch.object(dfm_rpc_bridge_router.dfm_rpc_bridge_service, "cleanup_tmp", return_value={"ok": True}) as cleanup,
            patch.object(dfm_rpc_bridge_router.dfm_rpc_bridge_service, "update_remote", return_value={"ok": True}) as update,
        ):
            dfm_rpc_bridge_router.sync_dfm_rpc_bridge(request)
            dfm_rpc_bridge_router.compare_dfm_rpc_bridge(request)
            dfm_rpc_bridge_router.keep_local_dfm_rpc_bridge(request)
            dfm_rpc_bridge_router.cleanup_dfm_rpc_bridge(request)
            dfm_rpc_bridge_router.update_remote_dfm_rpc_bridge(confirmed)
        sync.assert_called_once_with(request)
        compare.assert_called_once_with(request)
        keep.assert_called_once_with(request)
        cleanup.assert_called_once_with(request)
        update.assert_called_once_with(confirmed)
        self.assertEqual(
            [kind for kind, _ in mutations.calls],
            [
                "dfm_rpc_bridge_sync",
                "dfm_rpc_bridge_keep_local",
                "dfm_rpc_bridge_cleanup",
                "dfm_rpc_bridge_update_remote",
            ],
        )
        self.assertEqual([kind for kind, _ in reads.calls], ["dfm_rpc_bridge_compare"])
        self.assertTrue(mutations.calls[-1][1]["rpc_server_write_confirmed"])
        self._assert_registered(mutations, WORKSPACE_MUTATION_KINDS)
        self._assert_registered(reads, WORKSPACE_READ_KINDS)

    def test_a_hosted_sync_is_never_also_run_locally(self) -> None:
        # The Bridge exports from ResQ for every request file it claims, so a
        # second publish after an answered one is duplicated work.
        remote = _CaptureMutation({"ok": True, "status": "compared", "via": "gateway"}, remote=True)
        with (
            patch.object(dfm_rpc_bridge_router.workspace_mutation_client, "run_workspace_mutation", remote),
            patch.object(dfm_rpc_bridge_router.dfm_rpc_bridge_service, "send_sync_request") as local,
        ):
            response = dfm_rpc_bridge_router.sync_dfm_rpc_bridge(dfm_request())
        local.assert_not_called()
        self.assertEqual(response["via"], "gateway")


class HostedEntryPointTests(unittest.TestCase):
    def test_hosted_entry_points_rebuild_the_request_and_delegate(self) -> None:
        kwargs = dfm_request().model_dump()
        for function, delegate in (
            ("hosted_compare", "compare"),
            ("hosted_send_sync_request", "send_sync_request"),
            ("hosted_keep_local", "keep_local"),
            ("hosted_cleanup_tmp", "cleanup_tmp"),
        ):
            with patch.object(dfm_rpc_bridge_service, delegate, return_value={"ok": True}) as target:
                getattr(dfm_rpc_bridge_service, function)(**kwargs)
            passed = target.call_args.args[0]
            self.assertIsInstance(passed, DfmRpcBridgeRequest)
            self.assertEqual(passed.method_name, "M")

        with patch.object(dfm_rpc_bridge_service, "update_remote", return_value={"ok": True}) as target:
            dfm_rpc_bridge_service.hosted_update_remote(
                **{**kwargs, "rpc_server_write_confirmed": True}
            )
        self.assertIsInstance(target.call_args.args[0], DfmRpcBridgeUpdateRemoteRequest)
        self.assertTrue(target.call_args.args[0].rpc_server_write_confirmed)

    def test_a_bad_argument_is_refused_by_the_route_schema(self) -> None:
        kwargs = dfm_request().model_dump()
        kwargs["origin_length"] = 0
        with self.assertRaises(Exception):
            dfm_rpc_bridge_service.hosted_compare(**kwargs)


class RequestFileTests(unittest.TestCase):
    """The published request names the person who asked, not the process."""

    def test_the_request_file_carries_the_acting_user(self) -> None:
        from app_server.services import user_identity_service

        with tempfile.TemporaryDirectory() as folder:
            with user_identity_service.acting_identity("someone_else", "Someone Else"):
                path = dfm_rpc_bridge_service._write_request_file(
                    dfm_request(), "DFM", "C:\\out.json", folder
                )
            with open(path, "r", encoding="utf-8") as handle:
                payload = json.load(handle)
        self.assertEqual(payload["UserName"], "someone_else")
        self.assertNotEqual(payload["UserName"], os.environ.get("USERNAME"))


class CompareReadCountTests(unittest.TestCase):
    """Each side of the comparison is parsed once, not two and three times."""

    def _method_payload(self) -> dict:
        return {
            "json_format": "arcrho-dfm-v4",
            "details_tab": {"name": "M", "output_dataset": "Out"},
            "ratios_tab": {"ratio_triangle": {"excluded": [[0, 1]]}, "average_formulas": {"label": ["Straight"]}},
            "method_metadata": {"last_modified": "2026-08-18T10:00:00Z"},
        }

    def _run(self, service, request) -> int:
        opened: list[str] = []
        real_open = open

        def counting_open(path, *args, **kwargs):
            opened.append(str(path))
            return real_open(path, *args, **kwargs)

        with tempfile.TemporaryDirectory() as folder:
            local_path = os.path.join(folder, "local.json")
            remote_path = os.path.join(folder, "remote.json")
            for path in (local_path, remote_path):
                with open(path, "w", encoding="utf-8") as handle:
                    json.dump(self._method_payload(), handle)
            paths = {
                "project_dir": folder,
                "data_dir": folder,
                "method_dir": folder,
                "rpc_methods_dir": folder,
                "request_dir": folder,
                "local_path": local_path,
                "remote_path": remote_path,
                "sync_status_path": os.path.join(folder, "status.json"),
            }
            with (
                patch.object(service, "build_paths", return_value=paths),
                patch("builtins.open", counting_open),
                patch.object(service, "_sidecar_method_notes_snapshot", return_value={"exists": False, "text": ""}),
            ):
                result = service.compare(request)
        self.assertTrue(result["ok"])
        self.assertEqual(result["comparison"], "same_time")
        return len([path for path in opened if path in (local_path, remote_path)])

    def test_dfm_compare_opens_each_method_json_once(self) -> None:
        self.assertEqual(self._run(dfm_rpc_bridge_service, dfm_request()), 2)


if __name__ == "__main__":
    unittest.main()
