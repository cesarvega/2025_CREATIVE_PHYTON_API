"""
Generate thumbnails for all existing background templates.

This script:
1. Connects to BI_GUIDELINES database
2. Retrieves all templates from nw_Templates table
3. For each template, generates a thumbnail (300x169px) if it doesn't exist
4. Saves thumbnails in images/BackGrounds/thumbnails/ directory
"""

import sys
import logging
from pathlib import Path
from typing import List, Dict
import pyodbc

# Add project root to path
project_root = Path(__file__).parent
sys.path.insert(0, str(project_root))

from app.config.settings import settings
from app.config.db import get_connection_scope
from app.utils.image_processor import create_thumbnail, get_thumbnail_filename

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


def get_all_templates() -> List[Dict]:
    """Retrieve all templates from database."""
    with get_connection_scope(timeout=30) as cursor:
        query = """
            SELECT TemplateId, TemplateName, TemplateGroup, TemplateFileName
            FROM [BI_GUIDELINES].[dbo].[nw_Templates]
            WHERE TemplateFileName IS NOT NULL
            ORDER BY TemplateId
        """
        
        cursor.execute(query)
        rows = cursor.fetchall()
        
        templates = []
        for row in rows:
            templates.append({
                'id': row.TemplateId,
                'name': row.TemplateName,
                'group': row.TemplateGroup,
                'file_name': row.TemplateFileName
            })
        
        return templates


def generate_thumbnails():
    """Generate thumbnails for all existing templates."""
    logger.info("=" * 80)
    logger.info("Generating thumbnails for existing background templates")
    logger.info("=" * 80)
    
    # Get all templates from database
    try:
        templates = get_all_templates()
        logger.info(f"\nFound {len(templates)} templates in database")
    except Exception as e:
        logger.error(f"Failed to retrieve templates from database: {e}")
        return
    
    # Determine base directory
    if settings.environment == "production":
        base_dir = Path("C:/inetpub/wwwroot/nw2/assets")
    else:
        base_dir = Path(settings.backgrounds_dir).parent.parent  # Go up from backgrounds to assets
    
    logger.info(f"Base directory: {base_dir}")
    
    # Statistics
    created = 0
    skipped = 0
    errors = 0
    
    # Process each template
    logger.info("\n" + "-" * 80)
    for template in templates:
        file_name = template['file_name']
        
        if not file_name:
            skipped += 1
            continue
        
        # Build full path to original image
        # file_name format: "images/BackGrounds/Backgrounds2019/BMW_1.jpg"
        source_path = base_dir / file_name
        
        if not source_path.exists():
            logger.warning(f"⚠️  Source not found: {source_path}")
            errors += 1
            continue
        
        # Build thumbnail path maintaining subdirectory structure
        # Example: images/BackGrounds/Backgrounds2019/BMW_1.jpg
        #       -> images/BackGrounds/Backgrounds2019/thumbnails/BMW_1.jpg
        path_parts = Path(file_name)
        thumbnail_filename = get_thumbnail_filename(source_path.name)
        
        # If file is in subdirectory, maintain structure
        if path_parts.parent and str(path_parts.parent) != '.':
            # Create thumbnails directory inside the same subdirectory
            thumbnails_dir = base_dir / path_parts.parent / "thumbnails"
        else:
            thumbnails_dir = base_dir / "thumbnails"
        
        thumbnails_dir.mkdir(parents=True, exist_ok=True)
        thumbnail_path = thumbnails_dir / thumbnail_filename
        
        # Check if thumbnail already exists
        if thumbnail_path.exists():
            logger.info(f"✓ Thumbnail exists: {thumbnail_filename}")
            skipped += 1
            continue
        
        # Generate thumbnail
        try:
            create_thumbnail(source_path, thumbnail_path)
            logger.info(f"✓ Created: {thumbnail_filename} (from {source_path.name})")
            created += 1
        except Exception as e:
            logger.error(f"✗ Failed to create thumbnail for {source_path.name}: {e}")
            errors += 1
    
    # Summary
    logger.info("\n" + "=" * 80)
    logger.info("SUMMARY")
    logger.info("=" * 80)
    logger.info(f"Total templates: {len(templates)}")
    logger.info(f"✓ Thumbnails created: {created}")
    logger.info(f"- Already existed: {skipped}")
    logger.info(f"✗ Errors: {errors}")
    logger.info("=" * 80)
    
    if created > 0:
        logger.info("\n✓ Thumbnail generation completed successfully!")
        logger.info(f"\nThumbnails saved maintaining subdirectory structure:")
        logger.info(f"  - {base_dir}/images/BackGrounds/thumbnails/")
        logger.info(f"  - {base_dir}/images/BackGrounds/Backgrounds2019/thumbnails/")
    elif skipped == len(templates):
        logger.info("\n✓ All thumbnails already exist!")
    else:
        logger.warning("\n⚠️  Some thumbnails could not be generated. Check errors above.")


if __name__ == "__main__":
    try:
        generate_thumbnails()
    except KeyboardInterrupt:
        logger.info("\n\nOperation cancelled by user")
    except Exception as e:
        logger.error(f"\n✗ Unexpected error: {e}", exc_info=True)
        sys.exit(1)
