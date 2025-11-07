"""Models for background task responses."""

from typing import Optional, Any, Dict
from pydantic import BaseModel, Field


class TaskCreatedResponse(BaseModel):
    """Response when a background task is created."""

    task_id: str = Field(..., description="Unique task identifier")
    status: str = Field(..., description="Current task status (pending, processing, completed, failed)")
    message: str = Field(..., description="Human-readable message")
    task_type: str = Field(..., description="Type of task being executed")
    created_at: str = Field(..., description="ISO timestamp when task was created")
    status_url: str = Field(..., description="URL to check task status")

    class Config:
        json_schema_extra = {
            "example": {
                "task_id": "550e8400-e29b-41d4-a716-446655440000",
                "status": "pending",
                "message": "Presentation creation task started",
                "task_type": "create_presentation",
                "created_at": "2025-01-15T10:30:00.123456",
                "status_url": "/api/presentations/tasks/550e8400-e29b-41d4-a716-446655440000"
            }
        }


class TaskStatusResponse(BaseModel):
    """Response for task status queries."""

    task_id: str = Field(..., description="Unique task identifier")
    task_type: str = Field(..., description="Type of task")
    status: str = Field(..., description="Current task status")
    description: str = Field(..., description="Task description")
    progress: int = Field(..., ge=0, le=100, description="Progress percentage (0-100)")
    result: Optional[Any] = Field(None, description="Task result (available when completed)")
    error: Optional[str] = Field(None, description="Error message (if failed)")
    metadata: Dict[str, Any] = Field(default_factory=dict, description="Additional task metadata")
    created_at: str = Field(..., description="ISO timestamp when task was created")
    started_at: Optional[str] = Field(None, description="ISO timestamp when task started processing")
    completed_at: Optional[str] = Field(None, description="ISO timestamp when task completed")

    class Config:
        json_schema_extra = {
            "example": {
                "task_id": "550e8400-e29b-41d4-a716-446655440000",
                "task_type": "create_presentation",
                "status": "completed",
                "description": "Creating presentation for project TestProject",
                "progress": 100,
                "result": {
                    "presentation_id": 12345,
                    "total_slides": 24,
                    "message": "Presentation created successfully"
                },
                "error": None,
                "metadata": {
                    "project": "TestProject",
                    "display_name": "Test_2025"
                },
                "created_at": "2025-01-15T10:30:00.123456",
                "started_at": "2025-01-15T10:30:01.234567",
                "completed_at": "2025-01-15T10:32:45.678901"
            }
        }


class TaskListResponse(BaseModel):
    """Response for listing tasks."""

    tasks: list[TaskStatusResponse] = Field(..., description="List of tasks")
    total: int = Field(..., description="Total number of tasks")

    class Config:
        json_schema_extra = {
            "example": {
                "tasks": [
                    {
                        "task_id": "550e8400-e29b-41d4-a716-446655440000",
                        "task_type": "create_presentation",
                        "status": "completed",
                        "description": "Creating presentation",
                        "progress": 100,
                        "result": {"presentation_id": 12345},
                        "error": None,
                        "metadata": {},
                        "created_at": "2025-01-15T10:30:00",
                        "started_at": "2025-01-15T10:30:01",
                        "completed_at": "2025-01-15T10:32:45"
                    }
                ],
                "total": 1
            }
        }
