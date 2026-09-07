import re
from typing import List, Optional

from pydantic import BaseModel, Field, field_validator


class PatchItem(BaseModel):
    r: int = Field(..., ge=0)
    c: int = Field(..., ge=0)
    value: Optional[float] = None


class PatchRequest(BaseModel):
    items: List[PatchItem]
    file_mtime: Optional[float] = None


class DatasetNotesSaveRequest(BaseModel):
    project_name: str
    reserving_class: str
    dataset_name: str
    notes: str = ""


class DatasetSidecarLoadRequest(BaseModel):
    project_name: str
    reserving_class: str
    dataset_name: str


class DatasetNumberFormatOverride(BaseModel):
    reserving_class: str = Field(..., min_length=1, max_length=512)
    dataset_type_name: str = Field(..., min_length=1, max_length=256)
    number_format: str = Field(..., min_length=1, max_length=64)

    @field_validator("reserving_class", "dataset_type_name", "number_format")
    @classmethod
    def normalize_text(cls, value: str) -> str:
        normalized = re.sub(r"\s+", " ", value.replace("\r", " ").replace("\n", " ").replace("\t", " ")).strip()
        if not normalized:
            raise ValueError("value must not be blank")
        return normalized


class DatasetNumberFormatsSaveRequest(BaseModel):
    expected_revision: int = Field(..., ge=0)
    default_number_format: str = Field(..., min_length=1, max_length=64)
    overrides: List[DatasetNumberFormatOverride] = Field(default_factory=list, max_length=5000)

    @field_validator("default_number_format")
    @classmethod
    def normalize_default_number_format(cls, value: str) -> str:
        normalized = value.replace("\r", " ").replace("\n", " ").replace("\t", " ").strip()
        if not normalized:
            raise ValueError("default_number_format must not be blank")
        return normalized


class DatasetCacheLoadRequest(BaseModel):
    project_name: str
    reserving_class: str
    dataset_name: str
    csv_file: str = ""
    origin_length: Optional[int] = Field(None, ge=1)
    development_length: Optional[int] = Field(None, ge=1)
    cumulative: bool = True
    calendar: bool = False


class DatasetCalculatedPreviewRequest(BaseModel):
    project_name: str
    reserving_class: str
    changed_dataset_name: str
    changed_dataset_type_name: str = ""
    values: List[List[Optional[float]]] = Field(default_factory=list)
    mask: Optional[List[List[bool]]] = None
    origin_labels: Optional[List[str]] = None
    development_labels: Optional[List[str]] = None


class DatasetExternalLinkTargetCell(BaseModel):
    row: int = Field(..., ge=0, strict=True)
    column: int = Field(..., ge=0, strict=True)
    source_cell: Optional[str] = None

    @field_validator("source_cell")
    @classmethod
    def normalize_source_cell(cls, value: Optional[str]) -> Optional[str]:
        if value is None:
            return None
        normalized = value.strip().replace("$", "").upper()
        if not re.fullmatch(r"[A-Z]+[1-9][0-9]*", normalized):
            raise ValueError("source_cell must be a valid Excel cell address")
        return normalized


class DatasetExternalLink(BaseModel):
    reference: str = Field(..., min_length=1, strict=True)
    target_cells: List[DatasetExternalLinkTargetCell] = Field(..., min_length=1)

    @field_validator("reference")
    @classmethod
    def normalize_reference(cls, value: str) -> str:
        normalized = value.strip()
        if not normalized:
            raise ValueError("reference must not be blank")
        return normalized


class DatasetInternalLinkTargetCell(BaseModel):
    row: int = Field(..., ge=0, strict=True)
    column: int = Field(..., ge=0, strict=True)
    source_row: int = Field(..., ge=0, strict=True)
    source_column: int = Field(..., ge=0, strict=True)


class DatasetInternalLink(BaseModel):
    reference: str = Field(..., min_length=1, strict=True)
    target_cells: List[DatasetInternalLinkTargetCell] = Field(..., min_length=1)

    @field_validator("reference")
    @classmethod
    def normalize_reference(cls, value: str) -> str:
        normalized = value.strip()
        if not normalized:
            raise ValueError("reference must not be blank")
        return normalized


class DatasetFormulaLinkTargetCell(BaseModel):
    row: int = Field(..., ge=0, strict=True)
    column: int = Field(..., ge=0, strict=True)
    result_row: int = Field(..., ge=0, strict=True)
    result_column: int = Field(..., ge=0, strict=True)


class DatasetFormulaLink(BaseModel):
    formula: str = Field(..., min_length=1, strict=True)
    target_cells: List[DatasetFormulaLinkTargetCell] = Field(..., min_length=1)

    @field_validator("formula")
    @classmethod
    def normalize_formula(cls, value: str) -> str:
        normalized = value.strip()
        if not normalized:
            raise ValueError("formula must not be blank")
        return normalized


class DatasetInternalLinksResolveRequest(BaseModel):
    project_name: str
    reserving_class: str
    references: List[str] = Field(..., min_length=1)


class DatasetSidecarSaveRequest(BaseModel):
    project_name: str
    reserving_class: str
    dataset_name: str
    dataset_type: str = ""
    instance_name: str = ""
    source_kind: str = ""
    data_format: str = ""
    origin_length: int = Field(..., ge=1)
    development_length: int = Field(..., ge=1)
    # The months per period the CSV is written at, when the caller asks for a
    # store finer than the display. Omitted, the store follows the display.
    stored_development_length: Optional[int] = Field(None, ge=1)
    cumulative: bool = True
    transposed: bool = False
    calendar: bool = False
    show_subtotal: Optional[bool] = None
    number_format: str = ""
    decimal_places: Optional[int] = Field(None, ge=0, le=6)
    origin_labels: Optional[List[str]] = None
    csv_file: str = ""
    method_type: str = ""
    status: Optional[int] = None
    notes: Optional[str] = None
    precedents: Optional[List[str]] = None
    external_links: Optional[List[DatasetExternalLink]] = None
    internal_links: Optional[List[DatasetInternalLink]] = None
    formula_links: Optional[List[DatasetFormulaLink]] = None
    values: Optional[List[List[Optional[float]]]] = None
    mask: Optional[List[List[bool]]] = None
    # Fingerprint of the dependent-update plan the user confirmed. The Engine
    # rechecks it under the reserving-class lease and refuses with 409 if the
    # class changed while the plan was on screen.
    plan_fingerprint: str = ""


class EmptyDatasetCacheCreateRequest(BaseModel):
    project_name: str
    reserving_class: str
    dataset_type: str
    instance_name: str = ""
    data_format: str = "Triangle"
    origin_length: int = Field(12, ge=1)
    development_length: int = Field(12, ge=1)
    cumulative: bool = True
    calendar: bool = False


class CachedDatasetDeleteRequest(BaseModel):
    project_name: str
    reserving_class: str
    dataset_names: List[str] = Field(default_factory=list)
