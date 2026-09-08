"""ArcRho Web UI — FastAPI application.

This module creates the FastAPI ``app`` instance, includes all API routers,
and mounts the static frontend.  All business logic lives in
``app_server.services.*`` and route handlers in ``app_server.api.*``.
"""
from __future__ import annotations

import asyncio
import logging
import os
from contextlib import asynccontextmanager
from pathlib import Path

from fastapi import FastAPI
from fastapi.responses import RedirectResponse

from arcrho_log_retention_contract import prune_aged_log_files

from app_server import config
from app_server.ui_static import RevalidatedStaticFiles
from app_server.services import hosted_save_enrollment_service
from app_server.api import (
    workflow_router,
    app_control_router,
    workspace_paths_router,
    audit_log_router,
    dataset_router,
    book_router,
    excel_router,
    excel_link_router,
    arcrho_router,
    project_settings_router,
    table_summary_router,
    source_table_router,
    field_mapping_router,
    dataset_types_router,
    dependent_propagation_router,
    object_change_watch_router,
    reserving_class_router,
    scripting_router,
    dfm_rpc_bridge_router,
    dfm_method_router,
    result_selection_router,
    bornhuetter_ferguson_router,
    berquist_sherman_router,
    cape_cod_router,
    bootstrap_router,
    dfm_method_index_router,
    project_user_preferences_router,
    ui_automation_router,
    snowflake_router,
    sql_server_router,
    sql_formatting_router,
    data_processing_rules_router,
    user_identity_router,
)

# ---------------------------------------------------------------------------
# App
# ---------------------------------------------------------------------------

LOGGER = logging.getLogger(__name__)


@asynccontextmanager
async def lifespan(_app: FastAPI):
    await asyncio.to_thread(
        prune_aged_log_files,
        Path(config.get_client_save_latency_log_path()).parent,
    )
    enrollment = await asyncio.to_thread(
        hosted_save_enrollment_service.auto_enroll_current_user
    )
    LOGGER.info("Gateway startup enrollment status: %s", enrollment["status"])
    yield


app = FastAPI(title="Triangle Demo API", version="0.1", lifespan=lifespan)

# --- Include routers (API routes BEFORE static mount) ---
app.include_router(workflow_router)
app.include_router(app_control_router)
app.include_router(workspace_paths_router)
app.include_router(audit_log_router)
app.include_router(dataset_router)
app.include_router(book_router)
app.include_router(excel_router)
app.include_router(excel_link_router)
app.include_router(arcrho_router)
app.include_router(project_settings_router)
app.include_router(table_summary_router)
app.include_router(source_table_router)
app.include_router(field_mapping_router)
app.include_router(dataset_types_router)
app.include_router(dependent_propagation_router)
app.include_router(object_change_watch_router)
app.include_router(reserving_class_router)
app.include_router(scripting_router)
app.include_router(dfm_rpc_bridge_router)
app.include_router(dfm_method_router)
app.include_router(result_selection_router)
app.include_router(bornhuetter_ferguson_router)
app.include_router(berquist_sherman_router)
app.include_router(cape_cod_router)
app.include_router(bootstrap_router)
app.include_router(dfm_method_index_router)
app.include_router(project_user_preferences_router)
app.include_router(ui_automation_router)
app.include_router(snowflake_router)
app.include_router(sql_server_router)
app.include_router(sql_formatting_router)
app.include_router(data_processing_rules_router)
app.include_router(user_identity_router)

# --- Frontend assets (served from ./ui and ./icons, no /static) ---
# Mount AFTER API routes to avoid conflicts

app.mount("/ui", RevalidatedStaticFiles(directory=str(config.PROJECT_ROOT / "ui"), html=True), name="ui")
app.mount("/icons", RevalidatedStaticFiles(directory=str(config.PROJECT_ROOT / "icons")), name="icons")


@app.get("/")
def home():
    return RedirectResponse(url="/ui/")


@app.get("/app/health")
def app_health():
    return {
        "ok": True,
        "app": "arcrho",
        "token": os.environ.get("ARCRHO_BACKEND_TOKEN", ""),
        "backend_artifact_id": os.environ.get("ARCRHO_BACKEND_ARTIFACT_ID", ""),
        "project_root": str(config.PROJECT_ROOT),
    }
