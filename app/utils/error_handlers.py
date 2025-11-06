"""Centralized error handling utilities to improve consistency and reduce duplication."""

from __future__ import annotations

from functools import wraps
from typing import Any, Callable, Optional, TypeVar, cast

from fastapi import HTTPException, status

from app.utils.logging_utils import get_logger

logger = get_logger(__name__)

T = TypeVar('T')


class ServiceError(Exception):
    """Base exception for service-layer errors.

    OPTIMIZATION: Provides structured error handling across all services.
    """

    def __init__(self, message: str, status_code: int = 500, details: Optional[dict] = None):
        self.message = message
        self.status_code = status_code
        self.details = details or {}
        super().__init__(self.message)


class DatabaseError(ServiceError):
    """Exception for database-related errors."""

    def __init__(self, message: str, details: Optional[dict] = None):
        super().__init__(message, status_code=500, details=details)


class ValidationError(ServiceError):
    """Exception for validation errors."""

    def __init__(self, message: str, details: Optional[dict] = None):
        super().__init__(message, status_code=400, details=details)


class NotFoundError(ServiceError):
    """Exception for resource not found errors."""

    def __init__(self, message: str, details: Optional[dict] = None):
        super().__init__(message, status_code=404, details=details)


class ConflictError(ServiceError):
    """Exception for resource conflict errors."""

    def __init__(self, message: str, details: Optional[dict] = None):
        super().__init__(message, status_code=409, details=details)


def handle_service_errors(func: Callable[..., T]) -> Callable[..., T]:
    """Decorator to handle service errors and convert them to HTTPException.

    OPTIMIZATION: Centralizes error handling for FastAPI endpoints, eliminating
    ~50+ duplicate try/except blocks across route handlers.

    Args:
        func: The function to decorate (typically a FastAPI route handler)

    Returns:
        Decorated function with error handling

    Example:
        >>> @handle_service_errors
        ... async def get_presentation(presentation_id: int):
        ...     if not presentation_id:
        ...         raise ValidationError("Invalid presentation ID")
        ...     return {"id": presentation_id}
    """
    @wraps(func)
    async def wrapper(*args: Any, **kwargs: Any) -> T:
        try:
            return await func(*args, **kwargs)
        except ServiceError as e:
            logger.error(
                "Service error in %s: %s (status=%d, details=%s)",
                func.__name__,
                e.message,
                e.status_code,
                e.details
            )
            raise HTTPException(
                status_code=e.status_code,
                detail={"message": e.message, "details": e.details}
            )
        except HTTPException:
            # Re-raise HTTP exceptions as-is
            raise
        except Exception as e:
            logger.error(
                "Unexpected error in %s: %s",
                func.__name__,
                str(e),
                exc_info=True
            )
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail={"message": "An unexpected error occurred", "error": str(e)}
            )

    return cast(Callable[..., T], wrapper)


def handle_service_errors_sync(func: Callable[..., T]) -> Callable[..., T]:
    """Decorator to handle service errors in synchronous functions.

    OPTIMIZATION: Same as handle_service_errors but for non-async functions.

    Args:
        func: The synchronous function to decorate

    Returns:
        Decorated function with error handling

    Example:
        >>> @handle_service_errors_sync
        ... def get_data(data_id: int):
        ...     if not data_id:
        ...         raise ValidationError("Invalid data ID")
        ...     return {"id": data_id}
    """
    @wraps(func)
    def wrapper(*args: Any, **kwargs: Any) -> T:
        try:
            return func(*args, **kwargs)
        except ServiceError as e:
            logger.error(
                "Service error in %s: %s (status=%d, details=%s)",
                func.__name__,
                e.message,
                e.status_code,
                e.details
            )
            raise HTTPException(
                status_code=e.status_code,
                detail={"message": e.message, "details": e.details}
            )
        except HTTPException:
            # Re-raise HTTP exceptions as-is
            raise
        except Exception as e:
            logger.error(
                "Unexpected error in %s: %s",
                func.__name__,
                str(e),
                exc_info=True
            )
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail={"message": "An unexpected error occurred", "error": str(e)}
            )

    return wrapper


def safe_execute(
    func: Callable[..., T],
    default: T,
    error_message: str = "Error executing function",
    log_errors: bool = True,
    *args: Any,
    **kwargs: Any
) -> T:
    """Safely execute a function and return a default value on error.

    OPTIMIZATION: Eliminates repetitive try/except patterns where we want to
    continue execution with a fallback value instead of propagating the error.

    Args:
        func: Function to execute
        default: Default value to return on error
        error_message: Custom error message for logging
        log_errors: If True, log errors (default: True)
        *args: Positional arguments for func
        **kwargs: Keyword arguments for func

    Returns:
        Function result or default value on error

    Example:
        >>> result = safe_execute(
        ...     some_risky_function,
        ...     default=[],
        ...     error_message="Failed to fetch data",
        ...     param1="value1"
        ... )
    """
    try:
        return func(*args, **kwargs)
    except Exception as e:
        if log_errors:
            logger.error("%s: %s", error_message, str(e), exc_info=True)
        return default


def validate_required_fields(data: dict, required_fields: list[str]) -> None:
    """Validate that all required fields are present and non-empty in a dictionary.

    OPTIMIZATION: Centralized validation to replace manual if/else checks.

    Args:
        data: Dictionary to validate
        required_fields: List of required field names

    Raises:
        ValidationError: If any required field is missing or empty

    Example:
        >>> request_data = {"name": "Test", "email": "test@example.com"}
        >>> validate_required_fields(request_data, ["name", "email", "phone"])
        >>> # Raises ValidationError: Missing required fields: phone
    """
    missing_fields = []
    empty_fields = []

    for field in required_fields:
        if field not in data:
            missing_fields.append(field)
        elif not data[field]:
            empty_fields.append(field)

    errors = []
    if missing_fields:
        errors.append(f"Missing required fields: {', '.join(missing_fields)}")
    if empty_fields:
        errors.append(f"Empty required fields: {', '.join(empty_fields)}")

    if errors:
        raise ValidationError(
            "; ".join(errors),
            details={"missing": missing_fields, "empty": empty_fields}
        )


def validate_positive_integer(value: Any, field_name: str) -> int:
    """Validate that a value is a positive integer.

    OPTIMIZATION: Common validation pattern used across services.

    Args:
        value: Value to validate
        field_name: Name of the field (for error messages)

    Returns:
        The validated integer value

    Raises:
        ValidationError: If value is not a positive integer

    Example:
        >>> presentation_id = validate_positive_integer(request.presentation_id, "presentation_id")
    """
    try:
        int_value = int(value)
        if int_value <= 0:
            raise ValidationError(
                f"{field_name} must be a positive integer, got {value}",
                details={"field": field_name, "value": value}
            )
        return int_value
    except (ValueError, TypeError):
        raise ValidationError(
            f"{field_name} must be a valid integer, got {value}",
            details={"field": field_name, "value": value}
        )


def log_and_raise_http_exception(
    status_code: int,
    message: str,
    details: Optional[dict] = None,
    log_level: str = "error"
) -> None:
    """Log an error and raise an HTTPException.

    OPTIMIZATION: Combines logging and exception raising into a single call.

    Args:
        status_code: HTTP status code
        message: Error message
        details: Optional additional details
        log_level: Logging level (default: "error")

    Raises:
        HTTPException: Always raises

    Example:
        >>> if not presentation:
        ...     log_and_raise_http_exception(
        ...         404,
        ...         f"Presentation {presentation_id} not found",
        ...         details={"presentation_id": presentation_id}
        ...     )
    """
    # Log the error
    log_func = getattr(logger, log_level, logger.error)
    if details:
        log_func("%s - Details: %s", message, details)
    else:
        log_func("%s", message)

    # Raise HTTPException
    raise HTTPException(
        status_code=status_code,
        detail={"message": message, "details": details or {}}
    )
