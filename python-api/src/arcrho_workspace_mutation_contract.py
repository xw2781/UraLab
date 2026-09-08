"""Contract for ArcRho Server-hosted workspace mutations.

``arcrho_workspace_read_contract`` hosts reads, whose defining property is that
they are pure functions of the workspace: an uncertain answer may simply be
asked again, or answered locally instead. A mutation cannot borrow that
reasoning, so it gets its own registry and its own route rather than being
smuggled into the read table.

Only allowlisted kinds may execute remotely, and every registered kind must be
**idempotent**: running it twice against the same workspace must leave the same
end state as running it once. That is what lets this transport, like the read
one, keep no durable receipt — an answer the client never saw is either already
applied or safe to ask for again. A mutation that cannot make that promise
belongs on the hosted-save path (``arcrho_hosted_save_http_contract``), which
buys durability with a request file, an Engine claim, and a receipt.

What the client may *not* do is fall back to its own mapped drive after the
server may already have acted. Reads fall back freely; a mutation whose outcome
is unknown must be reported, not repeated somewhere else, because the second
run would answer about a workspace the first one already changed.

Mutations run under the submitting user's identity so the audit trail, sidecar
stamps, and log lines name the person who asked rather than the Gateway's
service profile.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Mapping

from arcrho_dependent_propagation_contract import (
    DependentPropagationContractError,
    validate_project_name,
    validate_request_id,
    validate_reserving_class_path,
)


WORKSPACE_MUTATION_FUNCTION = "ArcRhoWorkspaceMutation"
WORKSPACE_MUTATION_CONTRACT_VERSION = 1
WORKSPACE_MUTATION_PATH = "/api/workspace-mutations"
WORKSPACE_MUTATION_CAPABILITY_FIELD = "workspace_mutation_kinds"
# A mutation removes or rewrites files in one reserving class and rebuilds that
# class's index; it never waits on the Engine, so it needs no save-sized budget.
WORKSPACE_MUTATION_TIMEOUT_SECONDS = 180.0
MAX_WORKSPACE_MUTATION_REQUEST_BYTES = 256 * 1024


class WorkspaceMutationContractError(ValueError):
    """Raised when a workspace-mutation payload violates this contract."""


@dataclass(frozen=True)
class WorkspaceMutationKind:
    """One remotely executable, idempotent ``app_server.services`` mutation."""

    module: str
    function: str
    required: tuple[str, ...]
    optional: tuple[str, ...] = ()
    # Arguments whose value is a list of names rather than a single string.
    # Named here so validation checks the right shape instead of coercing a
    # list to ``str`` and silently accepting "['a', 'b']" as one dataset.
    list_args: tuple[str, ...] = field(default_factory=tuple)

    @property
    def allowed(self) -> frozenset[str]:
        return frozenset(self.required + self.optional)


# The RPC-bridge kinds take their route schema's own fields, so the service can
# rebuild its request model from them and pydantic stays the single validator.
_DFM_RPC_BRIDGE_REQUIRED: tuple[str, ...] = (
    "project_name",
    "reserving_class",
    "method_name",
    "output_vector",
    "input_triangle",
    "origin_length",
    "development_length",
)
_DFM_RPC_BRIDGE_OPTIONAL: tuple[str, ...] = ("decimal_places", "timeout_sec")

# A hosted mutation that waits for the Bridge holds a Gateway worker thread for
# the duration, so the caller's budget is clamped into this range rather than
# trusted. The frontend asks for 8 s today; the ceiling exists so a client
# cannot pin a thread for minutes.
MIN_RPC_BRIDGE_WAIT_SECONDS = 0.1
MAX_RPC_BRIDGE_WAIT_SECONDS = 60.0


def clamp_rpc_bridge_wait(timeout_sec: Any) -> float:
    """Return the wait a hosted RPC-bridge exchange may hold a thread for."""

    try:
        wait = float(timeout_sec)
    except (TypeError, ValueError) as exc:
        raise WorkspaceMutationContractError("timeout_sec must be a number.") from exc
    if wait != wait:  # NaN
        raise WorkspaceMutationContractError("timeout_sec must be a number.")
    return min(MAX_RPC_BRIDGE_WAIT_SECONDS, max(MIN_RPC_BRIDGE_WAIT_SECONDS, wait))


# kind -> canonical service mutation. The Gateway resolves mutations only
# through this table; a request naming anything else, or passing an argument
# not listed here, is rejected before any import happens.
WORKSPACE_MUTATION_KINDS: dict[str, WorkspaceMutationKind] = {
    # Deleting a cached dataset removes its files and rebuilds the reserving
    # class index. It is idempotent because a file that is already gone is
    # skipped rather than failed, and the rebuild derives the index from
    # whatever survives; a repeat therefore reports "nothing matched" against
    # an end state identical to the first run's.
    "cached_dataset_delete": WorkspaceMutationKind(
        "dataset_service",
        "delete_cached_datasets",
        ("project_name", "reserving_class", "dataset_names"),
        list_args=("dataset_names",),
    ),
    # Submitting a source-table refresh publishes two small files into the
    # Engine's queue. It is idempotent because the client owns the request id:
    # an id that already has a published status is returned as-is rather than
    # queued a second time, so a lost response can never start a second import.
    # ``reserving_class_types`` is a list of ``{Name, Level}`` objects rather
    # than names, so it is not a list arg; the source-refresh contract
    # validates its shape when the request is built.
    "source_table_refresh_submit": WorkspaceMutationKind(
        "source_refresh_service",
        "submit_source_table_refresh_job",
        ("project_name", "request_id"),
        (
            "import_source",
            "force",
            "refresh_dependents",
            "dataset_types",
            "reserving_class_types",
        ),
        list_args=("dataset_types",),
    ),
    # Submitting a data-processing-rules save publishes two small files into
    # the Engine's queue, idempotent by the client-owned request id exactly as
    # the source refresh above. ``expected_revision`` may be 0 and ``rules``
    # may be an empty list (the user removed every rule), and the required
    # check reads both as absent, so they are listed as optional here and the
    # rules-job contract enforces their shape when the request is built.
    "data_processing_rules_save_submit": WorkspaceMutationKind(
        "data_processing_rules_job_service",
        "submit_data_processing_rules_job",
        ("project_name", "request_id"),
        ("expected_revision", "rules"),
    ),
    # The DFM sync dialog publishes a request file the ArcRho Bridge claims,
    # then waits for the JSON the Bridge exports from ResQ. Both halves are
    # server-local for the Bridge and both cross SMB for a Client PC, where the
    # publish is several round trips and every wait tick writes and deletes a
    # probe file so the redirector cannot serve a cached "not found". Hosted,
    # the request lands on local disk and the wait is a file-system event.
    #
    # Idempotent: the stale response and status files are deleted first, and a
    # repeat regenerates the same export from the same ResQ method. A second
    # run costs a duplicate ResQ export, never a divergent workspace.
    "dfm_rpc_bridge_sync": WorkspaceMutationKind(
        "dfm_rpc_bridge_service",
        "hosted_send_sync_request",
        _DFM_RPC_BRIDGE_REQUIRED,
        _DFM_RPC_BRIDGE_OPTIONAL,
    ),
    # Deleting the temporary response and status JSON. Idempotent because a
    # file that is already gone is skipped rather than failed.
    "dfm_rpc_bridge_cleanup": WorkspaceMutationKind(
        "dfm_rpc_bridge_service",
        "hosted_cleanup_tmp",
        _DFM_RPC_BRIDGE_REQUIRED,
        _DFM_RPC_BRIDGE_OPTIONAL,
    ),
    # Keeping the local method and discarding the remote export: the same
    # delete, with the message the dialog reports.
    "dfm_rpc_bridge_keep_local": WorkspaceMutationKind(
        "dfm_rpc_bridge_service",
        "hosted_keep_local",
        _DFM_RPC_BRIDGE_REQUIRED,
        _DFM_RPC_BRIDGE_OPTIONAL,
    ),
    # Writing the local method's owned settings back into the RPC server. This
    # is the one kind whose effect lands outside the workspace, so it carries
    # the transport rule most strictly: once the Gateway has accepted the
    # request, an ambiguous outcome is reported, never retried over SMB.
    # Idempotent in the sense the contract requires: a repeat writes the same
    # values from the same local method and saves again. The confirmation flag
    # is optional here on purpose: a false value must reach the service so both
    # transports refuse it with the same message.
    "dfm_rpc_bridge_update_remote": WorkspaceMutationKind(
        "dfm_rpc_bridge_service",
        "hosted_update_remote",
        _DFM_RPC_BRIDGE_REQUIRED,
        _DFM_RPC_BRIDGE_OPTIONAL + ("rpc_server_write_confirmed",),
    ),
    # The Sync and Export Reserving Class with ResQ macros publish one request
    # file into the Bridge's sync queue. Hosted, that write lands on the
    # server's local disk instead of crossing the share from a Client PC.
    # Idempotent because the client owns the request id: an id that already
    # has a request or a status file is returned as-is rather than published
    # again, so a lost response can never queue a second run. A reviewed
    # synchronization sends back the rows exactly as the preview reported them;
    # a whole-class transfer sends the direction it is reviewing, or the names
    # that review ticked.
    "resq_sync_request_publish": WorkspaceMutationKind(
        "resq_sync_queue_service",
        "publish_resq_sync_request",
        ("project_name", "reserving_class", "request_id", "phase"),
        ("selected_rows", "selected_names", "direction"),
    ),
    # Both ResQ import macros copy the reserving class they are about to
    # rewrite into the server's pre-import backups. That copy is one file per
    # method, sidecar and data file, so from a Client PC it is a round trip
    # each; hosted, the whole copy is local disk and the macro pays one
    # request.
    #
    # Idempotent because the client owns the backup id: an id whose copy this
    # host already finished -- which its manifest records -- is reported as it
    # stands rather than copied again under a second folder. A copy that died
    # part way leaves no manifest and is never presented as a restore point.
    "resq_import_backup": WorkspaceMutationKind(
        "resq_import_backup_service",
        "back_up_reserving_class_for_import",
        ("project_name", "reserving_class", "backup_id"),
        ("import_policy",),
    ),
}

HTTP_WORKSPACE_MUTATION_KINDS: tuple[str, ...] = tuple(sorted(WORKSPACE_MUTATION_KINDS))


def build_workspace_mutation_request(
    *,
    request_id: str,
    mutation_kind: str,
    kwargs: Mapping[str, Any],
    user_name: str,
    user_display_name: str = "",
) -> dict[str, Any]:
    return validate_workspace_mutation_request(
        {
            "Function": WORKSPACE_MUTATION_FUNCTION,
            "ContractVersion": WORKSPACE_MUTATION_CONTRACT_VERSION,
            "RequestId": request_id,
            "MutationKind": mutation_kind,
            "Kwargs": dict(kwargs),
            "UserName": user_name,
            "UserDisplayName": user_display_name,
        }
    )


def _validate_list_arg(kind: str, name: str, value: Any) -> list[str]:
    if not isinstance(value, (list, tuple)):
        raise WorkspaceMutationContractError(
            f"Workspace mutation {kind!r} expects {name!r} to be a list of names."
        )
    names = [str(item or "").strip() for item in value]
    names = [item for item in names if item]
    if not names:
        raise WorkspaceMutationContractError(
            f"Workspace mutation {kind!r} requires at least one {name!r} entry."
        )
    return names


def validate_workspace_mutation_request(payload: Mapping[str, Any]) -> dict[str, Any]:
    if not isinstance(payload, Mapping):
        raise WorkspaceMutationContractError("A workspace-mutation request must be a JSON object.")
    if str(payload.get("Function") or "") != WORKSPACE_MUTATION_FUNCTION:
        raise WorkspaceMutationContractError("Not a workspace-mutation request.")
    version = payload.get("ContractVersion")
    if version != WORKSPACE_MUTATION_CONTRACT_VERSION:
        raise WorkspaceMutationContractError(
            f"Unsupported workspace-mutation contract version: {version!r}"
        )
    kind = str(payload.get("MutationKind") or "").strip()
    spec = WORKSPACE_MUTATION_KINDS.get(kind)
    if spec is None:
        raise WorkspaceMutationContractError(f"Unknown workspace-mutation kind: {kind!r}")
    kwargs = payload.get("Kwargs")
    if not isinstance(kwargs, Mapping):
        raise WorkspaceMutationContractError("Workspace-mutation Kwargs must be an object.")
    unexpected = sorted(set(kwargs) - spec.allowed)
    if unexpected:
        raise WorkspaceMutationContractError(
            f"Workspace mutation {kind!r} does not accept: {', '.join(unexpected)}."
        )

    normalized_kwargs = dict(kwargs)
    missing = [
        name
        for name in spec.required
        if name not in spec.list_args and not str(kwargs.get(name) or "").strip()
    ]
    if missing:
        raise WorkspaceMutationContractError(
            f"Workspace mutation {kind!r} requires: {', '.join(missing)}."
        )
    for name in spec.list_args:
        if name in spec.required or name in kwargs:
            normalized_kwargs[name] = _validate_list_arg(kind, name, kwargs.get(name))

    try:
        request_id = validate_request_id(payload.get("RequestId"))
        # Only logical identifiers travel; a machine-local project folder or a
        # drive-letter reserving-class path is refused before any lookup.
        validate_project_name(normalized_kwargs["project_name"], "project_name")
        if "reserving_class" in spec.allowed and normalized_kwargs.get("reserving_class"):
            validate_reserving_class_path(normalized_kwargs["reserving_class"])
    except DependentPropagationContractError as exc:
        raise WorkspaceMutationContractError(str(exc)) from exc
    return {
        "Function": WORKSPACE_MUTATION_FUNCTION,
        "ContractVersion": WORKSPACE_MUTATION_CONTRACT_VERSION,
        "RequestId": request_id,
        "MutationKind": kind,
        "Kwargs": normalized_kwargs,
        "UserName": str(payload.get("UserName") or "").strip(),
        "UserDisplayName": str(payload.get("UserDisplayName") or "").strip(),
    }
