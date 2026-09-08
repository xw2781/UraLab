"""Application service package with explicit lazy module loading.

Service modules are independent units.  Eagerly importing every one of them
made a targeted service import load unrelated integrations (including the web
application's optional runtime dependencies).  Keep the public module names
unchanged while importing a service only when a consumer asks for it.
"""
from __future__ import annotations

from importlib import import_module
from typing import Any


__all__ = [
    "workflow_service",
    "audit_service",
    "book_service",
    "dataset_instance_index_service",
    "dataset_number_format_service",
    "dataset_sidecar_status_service",
    "dataset_service",
    "excel_service",
    "arcrho_runtime_service",
    "project_settings_service",
    "table_summary_service",
    "dataset_types_service",
    "dataset_types_change_service",
    "dataset_types_plan_service",
    "calculated_dataset_service",
    "reserving_class_service",
    "field_mapping_service",
    "dfm_rpc_bridge_service",
    "dfm_service",
    "result_selection_service",
    "project_user_preferences_service",
    "ui_automation_service",
    "snowflake_service",
    "mssql_odbc",
    "sql_console_results",
    "sql_server_service",
    "data_processing_rules_service",
    "user_identity_service",
]


def __getattr__(name: str) -> Any:
    """Resolve a documented service module on first access."""

    if name not in __all__:
        raise AttributeError(f"module {__name__!r} has no attribute {name!r}")
    module = import_module(f"{__name__}.{name}")
    globals()[name] = module
    return module
