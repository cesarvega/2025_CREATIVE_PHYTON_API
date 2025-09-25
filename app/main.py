"""
FastAPI main application file
"""

import logging

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from scalar_fastapi import get_scalar_api_reference

from app.api.routes import (
    bi_guidelines,
    excel_processing,
    pptx_conversion,
    presentation_routes,
)
from app.config.settings import settings
from app.utils.logging_utils import setup_logging

# Ensure directories exist before mounting static files
settings.ensure_directories()

# Setup logging
setup_logging()
logger = logging.getLogger(__name__)

# Create FastAPI application
app = FastAPI(
    title=settings.app_name,
    description="API for generating professional converting PowerPoint files",
    version=settings.app_version,
    docs_url="/docs",
    root_path=(
        "/CreativePythonAPI" if settings.environment.lower() == "production" else ""
    ),
)

# CORS configuration
app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.cors_origins_list,
    allow_credentials=True,
    allow_methods=settings.cors_methods_list,
    allow_headers=settings.cors_headers_list,
)

# Include API routes
app.include_router(pptx_conversion.router, prefix="/api", tags=["PPTX Conversion"])
app.include_router(excel_processing.router, prefix="/api", tags=["Excel Processing"])
app.include_router(bi_guidelines.router, prefix="/api", tags=["BI Guidelines"])
app.include_router(
    presentation_routes.router, prefix="/api", tags=["Presentation Creation"]
)


@app.get("/", summary="Root endpoint")
async def root():
    """Root endpoint - returns basic API information"""
    return {
        "message": settings.app_name,
        "version": settings.app_version,
        "docs": "/docs",
        "scalar_docs": "/scalar",
    }


@app.get("/scalar", include_in_schema=False)
async def scalar_html():
    """
    Modern API documentation using Scalar
    """
    return get_scalar_api_reference(
        openapi_url=app.openapi_url,
        title=f"{settings.app_name} - API Documentation",
    )


@app.get("/health", summary="Health check endpoint")
async def health_check():
    """Health check endpoint"""
    return {"status": "healthy", "message": f"{settings.app_name} is running"}
