from __future__ import annotations

from typing import Any, Dict, List

from fastapi import APIRouter, HTTPException

from app_server.schemas.dataset import (
    CachedDatasetDeleteRequest,
    DatasetCacheLoadRequest,
    DatasetCalculatedPreviewRequest,
    DatasetInternalLinksResolveRequest,
    DatasetNotesSaveRequest,
    DatasetNumberFormatsSaveRequest,
    DatasetSidecarLoadRequest,
    DatasetSidecarSaveRequest,
    EmptyDatasetCacheCreateRequest,
    PatchRequest,
)
from app_server.services import dataset_service, engine_hosted_save_service
from app_server.services import calculated_dataset_service
from app_server.services import dataset_internal_link_service
from app_server.services import dataset_number_format_service
from app_server.services import workspace_mutation_client
from app_server.services import workspace_read_client

router = APIRouter()


@router.get("/dataset/number-format-defaults")
def get_dataset_number_format_defaults(
    dataset_type_name: str = "",
) -> Dict[str, Any]:
    return dataset_number_format_service.get_preferences(
        dataset_type_name=dataset_type_name,
    )


@router.put("/dataset/number-format-defaults")
def save_dataset_number_format_defaults(req: DatasetNumberFormatsSaveRequest) -> Dict[str, Any]:
    return dataset_number_format_service.save_preferences(
        expected_revision=req.expected_revision,
        default_number_format=req.default_number_format,
        overrides=[item.model_dump() for item in req.overrides],
    )


@router.get("/datasets")
def list_datasets() -> List[Dict[str, Any]]:
    return dataset_service.list_datasets()


@router.get("/datasets/cached")
def list_cached_dataset_names(project_name: str, reserving_class: str, refresh: bool = False) -> Dict[str, Any]:
    return workspace_read_client.run_workspace_read(
        "dataset_index",
        {"project_name": project_name, "reserving_class": reserving_class, "refresh": bool(refresh)},
        local=lambda: dataset_service.list_cached_dataset_names(
            project_name, reserving_class, refresh=refresh
        ),
    )


@router.get("/datasets/cached/index-signature")
def get_cached_dataset_index_signature(project_name: str, reserving_class: str) -> Dict[str, Any]:
    # Deliberately local even when the index itself is read on the Gateway: this
    # is a single stat that clients poll on a timer, so a hosted round trip
    # would cost more than the call saves.
    return dataset_service.get_cached_dataset_index_signature(project_name, reserving_class)


@router.post("/datasets/cached/delete")
def delete_cached_datasets(req: CachedDatasetDeleteRequest) -> Dict[str, Any]:
    # Hosted on the Server when the Gateway offers it: the dependency check
    # reads one sidecar per selected dataset, each unlink is its own round trip,
    # and the index rebuild reads the folder again -- all local disk there.
    return workspace_mutation_client.run_workspace_mutation(
        "cached_dataset_delete",
        {
            "project_name": req.project_name,
            "reserving_class": req.reserving_class,
            "dataset_names": list(req.dataset_names or []),
        },
        local=lambda: dataset_service.delete_cached_datasets(
            req.project_name,
            req.reserving_class,
            req.dataset_names,
        ),
    )


@router.post("/datasets/cached/empty")
def create_empty_cached_dataset(req: EmptyDatasetCacheCreateRequest) -> Dict[str, Any]:
    return dataset_service.create_empty_cached_dataset(
        req.project_name,
        req.reserving_class,
        req.dataset_type,
        instance_name=req.instance_name,
        data_format=req.data_format,
        origin_length=req.origin_length,
        development_length=req.development_length,
        cumulative=req.cumulative,
        calendar=req.calendar,
    )


@router.get("/dataset/{ds_id}")
def get_dataset(ds_id: str, project_name: str, origin_length: int) -> Dict[str, Any]:
    def load_locally() -> Dict[str, Any] | None:
        return dataset_service.get_dataset(ds_id, project_name=project_name, origin_length=origin_length)

    # The dataset handle is per process: the Gateway resolves it only when it
    # registered the id itself (a hosted dataset run or cached load in that
    # process). An answer without a dataset therefore means "not known
    # there", not "unknown", and this process — which registered the same id
    # from the rebased response — resolves it locally.
    result = workspace_read_client.run_workspace_read(
        "dataset_grid_load",
        {"ds_id": ds_id, "project_name": project_name, "origin_length": origin_length},
        local=load_locally,
    )
    if isinstance(result, dict) and result.get("id") is None:
        result = load_locally()
    if result is None:
        raise HTTPException(404, f"Unknown dataset: {ds_id}")
    return result


@router.get("/dataset/{ds_id}/diagonal")
def get_diagonal(ds_id: str, project_name: str, origin_length: int, k: int = 0) -> Dict[str, Any]:
    result = dataset_service.get_diagonal(
        ds_id,
        project_name=project_name,
        origin_length=origin_length,
        k=k,
    )
    if result is None:
        raise HTTPException(404, f"Unknown dataset: {ds_id}")
    return result


@router.post("/dataset/{ds_id}/patch")
def patch_dataset(ds_id: str, req: PatchRequest) -> Dict[str, Any]:
    result = dataset_service.patch_dataset(ds_id, req.items, file_mtime=req.file_mtime)
    if result is None:
        raise HTTPException(404, f"Unknown dataset: {ds_id}")
    if result.get("conflict"):
        raise HTTPException(409, "File changed on disk. Reload and retry.")
    return result


@router.post("/dataset/sidecar/load")
def load_dataset_sidecar(req: DatasetSidecarLoadRequest) -> Dict[str, Any]:
    return dataset_service.load_dataset_sidecar(
        req.project_name,
        req.reserving_class,
        req.dataset_name,
    )


@router.post("/dataset/notes/save")
def save_dataset_notes(req: DatasetNotesSaveRequest) -> Dict[str, Any]:
    return dataset_service.save_dataset_notes(
        req.project_name,
        req.reserving_class,
        req.dataset_name,
        req.notes,
    )


@router.post("/dataset/cache/load")
def load_dataset_cache(req: DatasetCacheLoadRequest) -> Dict[str, Any]:
    def _adopt(payload: Dict[str, Any]) -> Dict[str, Any]:
        dataset_service.register_dataset_handle(payload.get("id"), payload.get("path"))
        return payload

    return workspace_read_client.run_workspace_read(
        "dataset_cache_load",
        {
            "project_name": req.project_name,
            "reserving_class": req.reserving_class,
            "dataset_name": req.dataset_name,
            "csv_file": req.csv_file,
            "origin_length": req.origin_length,
            "development_length": req.development_length,
            "cumulative": req.cumulative,
            "calendar": req.calendar,
        },
        local=lambda: dataset_service.load_cached_dataset_values(
            req.project_name,
            req.reserving_class,
            req.dataset_name,
            csv_file=req.csv_file,
            origin_length=req.origin_length,
            development_length=req.development_length,
            cumulative=req.cumulative,
            calendar=req.calendar,
        ),
        finalize=_adopt,
    )


@router.post("/dataset/internal_links/resolve")
def resolve_dataset_internal_links(req: DatasetInternalLinksResolveRequest) -> Dict[str, Any]:
    # One cached-dataset read per unique referenced name; hosted on the
    # Gateway when it offers the kind so a Client PC pays one HTTP round trip
    # instead of one SMB visit per referenced dataset.
    return workspace_read_client.run_workspace_read(
        "dataset_internal_links_resolve",
        {
            "project_name": req.project_name,
            "reserving_class": req.reserving_class,
            "references": list(req.references),
        },
        local=lambda: dataset_internal_link_service.resolve_dataset_internal_links(
            req.project_name,
            req.reserving_class,
            req.references,
        ),
    )


@router.post("/dataset/calculated/preview")
def preview_calculated_dataset_dependents(req: DatasetCalculatedPreviewRequest) -> Dict[str, Any]:
    return calculated_dataset_service.preview_dependents(
        req.project_name,
        req.reserving_class,
        req.changed_dataset_name,
        changed_dataset_type_name=req.changed_dataset_type_name,
        values=req.values,
        mask=req.mask,
        origin_labels=req.origin_labels,
        development_labels=req.development_labels,
    )


def _dataset_sidecar_save_call(req: DatasetSidecarSaveRequest) -> Dict[str, Any]:
    """The one argument projection the plan and the save both run against."""

    return {
        "args": [req.project_name, req.reserving_class, req.dataset_name],
        "kwargs": {
            "dataset_type": req.dataset_type,
            "instance_name": req.instance_name,
            "source_kind": req.source_kind,
            "data_format": req.data_format,
            "origin_length": req.origin_length,
            "development_length": req.development_length,
            "stored_development_length": req.stored_development_length,
            "cumulative": req.cumulative,
            "transposed": req.transposed,
            "calendar": req.calendar,
            "show_subtotal": req.show_subtotal,
            "number_format": req.number_format,
            "decimal_places": req.decimal_places,
            "origin_labels": req.origin_labels,
            "csv_file": req.csv_file,
            "method_type": req.method_type,
            "status": req.status,
            "notes": req.notes,
            "precedents": req.precedents,
            "external_links": req.external_links,
            "internal_links": req.internal_links,
            "formula_links": req.formula_links,
            "values": req.values,
            "mask": req.mask,
        },
    }


@router.post("/dataset/sidecar/save/plan")
def plan_dataset_sidecar_save(req: DatasetSidecarSaveRequest) -> Dict[str, Any]:
    # Step one of the two-step save: name the dependent objects this save
    # would refresh. Nothing is written and no lease is taken.
    return engine_hosted_save_service.run_hosted_save_plan(
        "dataset_sidecar",
        req.project_name,
        req.reserving_class,
        **_dataset_sidecar_save_call(req),
    )


@router.post("/dataset/sidecar/save")
def save_dataset_sidecar(req: DatasetSidecarSaveRequest) -> Dict[str, Any]:
    # The save runs on ArcRho Engine next to the data; this endpoint keeps
    # its exact response shape and error codes.
    return engine_hosted_save_service.run_hosted_save(
        "dataset_sidecar",
        req.project_name,
        req.reserving_class,
        plan_fingerprint=req.plan_fingerprint,
        **_dataset_sidecar_save_call(req),
    )
