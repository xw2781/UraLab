"""Contract for ArcRho Server-hosted workspace reads.

A Client PC that opens a reserving class, a cached dataset, a method window, or
the Project Settings table summary pays one SMB round trip per file it touches,
and a stale reserving-class index makes it open every sidecar and method JSON in
the class over the mapped drive. The app server can instead ask the machine-wide
ArcRho Gateway to run the very same ``app_server`` service function on the
server host, where the workspace is local disk, and return the service's
response verbatim.

Only the allowlisted kinds below may execute remotely. Each kind names the
canonical service function and the keyword arguments a client may pass, so the
Gateway needs no second table of its own and a new read reaches HTTP the moment
it is registered here. The read kinds are pure functions of the workspace plus
their arguments: a repeated request is safe, so this transport keeps no
idempotency receipt.

Reads run under the submitting user's identity for the same reason hosted saves
do: a load that performs a one-time on-disk upgrade stamps that user, not the
Gateway's service profile.
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Any, Mapping

from arcrho_dependent_propagation_contract import (
    DependentPropagationContractError,
    validate_project_name,
    validate_request_id,
    validate_reserving_class_path,
)


WORKSPACE_READ_FUNCTION = "ArcRhoWorkspaceRead"
WORKSPACE_READ_CONTRACT_VERSION = 1
WORKSPACE_READ_PATH = "/api/workspace-reads"
# The response header naming the workspace root the read ran against, so the
# client can rebase any machine-local path in the payload onto its own root.
WORKSPACE_ROOT_HEADER = "X-ArcRho-Workspace-Root"
# A read either serves a persisted file or, at worst, rebuilds one index or
# summary on local disk; a cached-dataset load may also wait on one Engine
# header request.
WORKSPACE_READ_TIMEOUT_SECONDS = 120.0
MAX_WORKSPACE_READ_REQUEST_BYTES = 256 * 1024


class WorkspaceReadContractError(ValueError):
    """Raised when a workspace-read payload violates this contract."""


@dataclass(frozen=True)
class WorkspaceReadKind:
    """One remotely executable ``app_server.services`` read."""

    module: str
    function: str
    required: tuple[str, ...]
    optional: tuple[str, ...] = ()

    @property
    def allowed(self) -> frozenset[str]:
        return frozenset(self.required + self.optional)


# kind -> canonical service read. The Gateway resolves reads only through this
# table; a request naming anything else, or passing an argument not listed
# here, is rejected before any import happens.
WORKSPACE_READ_KINDS: dict[str, WorkspaceReadKind] = {
    "dataset_index": WorkspaceReadKind(
        "dataset_service",
        "list_cached_dataset_names",
        ("project_name", "reserving_class"),
        ("refresh",),
    ),
    "dataset_cache_load": WorkspaceReadKind(
        "dataset_service",
        "load_cached_dataset_values",
        ("project_name", "reserving_class", "dataset_name"),
        ("csv_file", "origin_length", "development_length", "cumulative", "calendar", "at_display_shape"),
    ),
    # The id-addressed grid load that follows a dataset run. ``ds_id`` is a
    # per-process handle; the server resolves it only if it registered the
    # handle itself (a hosted run or cached load in the same Gateway
    # process) and otherwise answers with no dataset, which the client treats
    # as "resolve locally".
    "dataset_grid_load": WorkspaceReadKind(
        "dataset_service",
        "get_dataset",
        ("ds_id", "project_name", "origin_length"),
    ),
    "dfm_method_load": WorkspaceReadKind(
        "dfm_service",
        "load_dfm_method",
        ("project_name", "reserving_class", "method_name"),
        ("output_dataset",),
    ),
    "result_selection_load": WorkspaceReadKind(
        "result_selection_service",
        "load_result_selection",
        ("project_name", "reserving_class", "method_name"),
        ("include_method",),
    ),
    "bornhuetter_ferguson_load": WorkspaceReadKind(
        "bornhuetter_ferguson_service",
        "load_bornhuetter_ferguson_method",
        ("project_name", "reserving_class", "method_name"),
    ),
    "cape_cod_load": WorkspaceReadKind(
        "cape_cod_service",
        "load_cape_cod_method",
        ("project_name", "reserving_class", "method_name"),
    ),
    "bootstrap_load": WorkspaceReadKind(
        "bootstrap_service",
        "load_bootstrap_method",
        ("project_name", "reserving_class", "method_name"),
    ),
    # B&S keeps its method JSON on the host API rather than an app-server save
    # path, so this read exists to pair that file with the output sidecar in one
    # visit. ``method_type`` picks the variant's filename prefix.
    "berquist_sherman_load": WorkspaceReadKind(
        "berquist_sherman_service",
        "load_berquist_sherman_method",
        ("project_name", "reserving_class", "method_type", "method_name"),
    ),
    # The whole listing is hosted, including whether each linked workbook can
    # be opened. That answer is deliberately the server host's: a workbook a
    # Client PC can see but ArcRho Server cannot is one no retarget or refresh
    # can read, so reporting it as found would be a lie.
    "excel_link_listing": WorkspaceReadKind(
        "excel_link_service",
        "list_reserving_class_excel_links",
        ("project_name", "reserving_class"),
    ),
    "table_summary": WorkspaceReadKind(
        "table_summary_service",
        "get_table_summary",
        ("project_name",),
    ),
    # Planning a dataset-type change reads one index per reserving class of
    # the project; from a Client PC that is one round trip each, so the plan
    # the confirmation dialog shows is built on the server host when it can be.
    "dataset_types_change_plan": WorkspaceReadKind(
        "dataset_types_plan_service",
        "plan_dataset_types_change_read",
        ("project_name", "rows", "renames"),
    ),
    # Polling a source-refresh job over the mapped drive reads a file the
    # server rewrote seconds ago, and Windows' directory cache can keep serving
    # the previous copy for several seconds after a terminal status lands. The
    # hosted read answers from the file the Engine actually wrote.
    "source_refresh_status": WorkspaceReadKind(
        "source_refresh_service",
        "get_source_table_refresh_status",
        ("project_name",),
        ("job_id",),
    ),
    # The rules-save job is polled the same way, for the same reason.
    "data_processing_rules_job_status": WorkspaceReadKind(
        "data_processing_rules_job_service",
        "get_data_processing_rules_job_status",
        ("project_name",),
        ("job_id",),
    ),
    # The DFM and Result Selection sync dialogs compare the local method JSON
    # against the copy the Bridge exported from ResQ. Both files live in the
    # workspace, so from a Client PC rendering the review window costs several
    # whole-file SMB reads before anything appears. The kwargs are the route
    # schema's own fields; the service rebuilds its request model from them so
    # validation keeps one owner.
    "dfm_rpc_bridge_compare": WorkspaceReadKind(
        "dfm_rpc_bridge_service",
        "hosted_compare",
        (
            "project_name",
            "reserving_class",
            "method_name",
            "output_vector",
            "input_triangle",
            "origin_length",
            "development_length",
        ),
        ("decimal_places", "timeout_sec"),
    ),
    "result_selection_rpc_bridge_compare": WorkspaceReadKind(
        "result_selection_rpc_bridge_service",
        "hosted_compare",
        ("project_name", "reserving_class", "method_name", "origin_length"),
        ("output_type", "timeout_sec"),
    ),
    # Resolving a Dataset window's internal cell links reads one cached
    # dataset per unique referenced name; on the server host those reads are
    # local disk, so a Client PC pays one HTTP round trip instead of one SMB
    # visit per referenced dataset.
    "dataset_internal_links_resolve": WorkspaceReadKind(
        "dataset_internal_link_service",
        "resolve_dataset_internal_links",
        ("project_name", "reserving_class", "references"),
    ),
    # The ResQ import and sync macros poll the Bridge worker's heartbeat, and
    # the status file of the request it is running, while they wait. Over the
    # mapped drive Windows serves those timestamps from a cache that can lag a
    # heartbeat written every second by ten seconds, so the look is taken on
    # the server host, where it is exact.
    "bridge_worker_liveness": WorkspaceReadKind(
        "bridge_liveness_service",
        "get_bridge_worker_liveness",
        (),
        ("queue", "request_id"),
    ),
}

HTTP_WORKSPACE_READ_KINDS: tuple[str, ...] = tuple(sorted(WORKSPACE_READ_KINDS))


def build_workspace_read_request(
    *,
    request_id: str,
    read_kind: str,
    kwargs: Mapping[str, Any],
    user_name: str,
    user_display_name: str = "",
) -> dict[str, Any]:
    return validate_workspace_read_request(
        {
            "Function": WORKSPACE_READ_FUNCTION,
            "ContractVersion": WORKSPACE_READ_CONTRACT_VERSION,
            "RequestId": request_id,
            "ReadKind": read_kind,
            "Kwargs": dict(kwargs),
            "UserName": user_name,
            "UserDisplayName": user_display_name,
        }
    )


def validate_workspace_read_request(payload: Mapping[str, Any]) -> dict[str, Any]:
    if not isinstance(payload, Mapping):
        raise WorkspaceReadContractError("A workspace-read request must be a JSON object.")
    if str(payload.get("Function") or "") != WORKSPACE_READ_FUNCTION:
        raise WorkspaceReadContractError("Not a workspace-read request.")
    version = payload.get("ContractVersion")
    if version != WORKSPACE_READ_CONTRACT_VERSION:
        raise WorkspaceReadContractError(
            f"Unsupported workspace-read contract version: {version!r}"
        )
    kind = str(payload.get("ReadKind") or "").strip()
    spec = WORKSPACE_READ_KINDS.get(kind)
    if spec is None:
        raise WorkspaceReadContractError(f"Unknown workspace-read kind: {kind!r}")
    kwargs = payload.get("Kwargs")
    if not isinstance(kwargs, Mapping):
        raise WorkspaceReadContractError("Workspace-read Kwargs must be an object.")
    unexpected = sorted(set(kwargs) - spec.allowed)
    if unexpected:
        raise WorkspaceReadContractError(
            f"Workspace read {kind!r} does not accept: {', '.join(unexpected)}."
        )
    missing = [name for name in spec.required if not str(kwargs.get(name) or "").strip()]
    if missing:
        raise WorkspaceReadContractError(
            f"Workspace read {kind!r} requires: {', '.join(missing)}."
        )
    try:
        request_id = validate_request_id(payload.get("RequestId"))
        # Only logical identifiers travel; a machine-local project folder or a
        # drive-letter reserving-class path is refused before any lookup. A kind
        # that names no project at all — the Bridge-worker liveness look reads
        # only the runtime folder — has nothing to check here, and the required
        # -field check above already refused an empty name where one is needed.
        if "project_name" in spec.allowed and kwargs.get("project_name"):
            validate_project_name(kwargs["project_name"], "project_name")
        if "reserving_class" in spec.allowed and kwargs.get("reserving_class"):
            validate_reserving_class_path(kwargs["reserving_class"])
    except DependentPropagationContractError as exc:
        raise WorkspaceReadContractError(str(exc)) from exc
    return {
        "Function": WORKSPACE_READ_FUNCTION,
        "ContractVersion": WORKSPACE_READ_CONTRACT_VERSION,
        "RequestId": request_id,
        "ReadKind": kind,
        "Kwargs": dict(kwargs),
        "UserName": str(payload.get("UserName") or "").strip(),
        "UserDisplayName": str(payload.get("UserDisplayName") or "").strip(),
    }
