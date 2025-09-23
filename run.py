"""
Script to run the Report Generator API application
"""

import uvicorn

from app.config.settings import settings
from app.utils.logging_utils import print_blue

def main():
    """Main function to start the FastAPI application"""
    print_blue(f"Starting {settings.app_name} – FastAPI service for PPTX conversion")
    print_blue(f"Documentation available at: http://{settings.host}:{settings.port}/docs")
    print_blue(f"Server starting on: http://{settings.host}:{settings.port}")

    try:
        uvicorn.run(
            "app.main:app",
            host=settings.host,
            port=settings.port,
            reload=True,
            log_level=settings.log_level.lower(),
            reload_dirs=["app"],
        )
    except KeyboardInterrupt:
        print_blue("Server stopped by user")
    except Exception as e:
        print_blue(f"Error starting server: {e}")
        raise


if __name__ == "__main__":
    main()
