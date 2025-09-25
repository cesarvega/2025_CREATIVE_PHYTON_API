"""
Script to run the Report Generator API application
"""

import uvicorn

from app.config.settings import settings
from app.utils.logging_utils import print_blue


def main():
    """Main function to start the FastAPI application"""
    # Display environment configuration
    env_info = settings.get_environment_info()
    print_blue(f"🚀 Starting {settings.app_name} – FastAPI service for PPTX conversion")
    print_blue(f"🌍 Environment: {env_info['environment'].upper()}")
    print_blue("📁 Base directories:")
    print_blue(f"   - NW Projects: {env_info['base_dir_nw']}")
    print_blue(f"   - BI Presents: {env_info['base_dir_bipresents']}")

    # Set URLs based on environment
    base_url = f"http://localhost:{settings.port}"
    if settings.environment.lower() == "production":
        base_url += "/CreativePythonAPI"

    print_blue(f"📖 Documentation available at: {base_url}/docs")
    print_blue(f"✨ Modern docs (Scalar): {base_url}/scalar")
    print_blue(f"🌐 Server starting on: {base_url}")

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
