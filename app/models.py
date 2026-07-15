from pydantic import BaseModel
from typing import Optional


class ImportResponse(BaseModel):
    project_id: int
    project_name: str
    phases_count: int
    nodes_count: int
    tasks_count: int


class NarrativeUpdate(BaseModel):
    content: str


class NarrativeBatchItem(BaseModel):
    section: str
    content: str


class NarrativeBatchUpdate(BaseModel):
    items: list[NarrativeBatchItem]


class NodeDetail(BaseModel):
    id: int
    date: str
    node_type: str


class PhaseDetail(BaseModel):
    id: int
    wbs: str
    name: str
    start_date: Optional[str]
    end_date: Optional[str]
    nodes: list[NodeDetail]


class ProjectItem(BaseModel):
    id: int
    name: str
    code: str
    created_at: str
    phases_count: int


class NarrativeSection(BaseModel):
    section: str
    label: str
    content: str
