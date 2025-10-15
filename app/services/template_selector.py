"""
Background image selection and rotation logic for presentation generation.

IMPORTANT: This module handles BACKGROUND IMAGES, not PowerPoint template files.

The logic works as follows:
1. The frontend sends background_name with template names separated by "|" (e.g., "BMW_1|BrandDNA|Kitchen2")
2. This service queries the nw_Templates table to get the image paths for each template name
3. The service provides methods to rotate through backgrounds for each slide
4. The backend saves both TemplateName and TemplateFileName to the database
5. The web viewer (frontend) is responsible for composing the slide image over the background image
"""

from __future__ import annotations

from typing import Dict, List, Optional

from app.models.presentation_models import BackgroundType, CreatePresentationMetadata
from app.services.bi_guidelines_service import bi_guidelines_service
from app.utils.logging_utils import get_logger

logger = get_logger(__name__)


class BackgroundImageSelector:
    """Service for selecting and rotating background images for presentations.

    This service does NOT handle PowerPoint template files (.pptx).
    It only handles background image paths from the nw_Templates table.
    """

    def __init__(self):
        """Initialize the background image selector."""
        self._background_cache: Dict[str, str] = {}

    def parse_background_names(self, background_name: str) -> List[str]:
        """Parse the background_name string to extract individual template names.

        Args:
            background_name: String with template names separated by "|"
                           Example: "BMW_1|BrandDNA|Kitchen2"

        Returns:
            List of template names
        """
        if not background_name:
            return []

        # Split by pipe delimiter
        names = [name.strip() for name in background_name.split('|') if name.strip()]

        logger.debug("Parsed background names: %s", names)
        return names

    def get_background_image_paths(self, background_name: str) -> Dict[str, str]:
        """Query the database to get background image paths for the given template names.

        Args:
            background_name: String with template names separated by "|"
                           Example: "BMW_1|BrandDNA|Kitchen2"

        Returns:
            Dictionary mapping template_name to template_file_name (image path)
            Example: {
                'BMW_1': 'images/BackGrounds/Backgrounds2019/BMW_1.jpg',
                'BrandDNA': 'images/BackGrounds/Backgrounds2019/BrandDNA.jpg'
            }
        """
        if not background_name:
            logger.warning("background_name is empty, returning empty dict")
            return {}

        # Parse template names
        template_names = self.parse_background_names(background_name)

        if not template_names:
            logger.warning("No template names found after parsing")
            return {}

        logger.info(
            "Querying nw_Templates for %d background(s): %s",
            len(template_names),
            template_names
        )

        # Query database for template paths
        template_map = bi_guidelines_service.get_background_templates_by_names(template_names)

        return template_map

    def create_background_rotator(self, background_name: str) -> 'BackgroundRotator':
        """Create a BackgroundRotator instance for rotating through backgrounds.

        Args:
            background_name: String with template names separated by "|"

        Returns:
            BackgroundRotator instance
        """
        template_map = self.get_background_image_paths(background_name)
        template_names = self.parse_background_names(background_name)

        return BackgroundRotator(template_names, template_map)


class BackgroundRotator:
    """Helper class to rotate through a list of background images for each slide.

    This class maintains state for rotating through backgrounds in order.
    """

    def __init__(self, template_names: List[str], template_map: Dict[str, str]):
        """Initialize the rotator.

        Args:
            template_names: Ordered list of template names (preserves user selection order)
            template_map: Dictionary mapping template_name to template_file_name
        """
        self.template_names = template_names
        self.template_map = template_map
        self.current_index = 0

        logger.info(
            "BackgroundRotator initialized with %d template(s)",
            len(template_names)
        )

    def get_next_background(self) -> tuple[Optional[str], Optional[str]]:
        """Get the next background in rotation.

        Returns:
            Tuple of (template_name, template_file_name)
            Returns (None, None) if no backgrounds available
        """
        if not self.template_names:
            return None, None

        # Get current template name
        template_name = self.template_names[self.current_index]

        # Get corresponding file path
        template_file_name = self.template_map.get(template_name)

        # Move to next index (circular rotation)
        self.current_index = (self.current_index + 1) % len(self.template_names)

        logger.debug(
            "Rotated to background: template_name=%s, path=%s",
            template_name,
            template_file_name
        )

        return template_name, template_file_name

    def reset(self):
        """Reset rotation to start from the beginning."""
        self.current_index = 0


# Global service instance
background_image_selector = BackgroundImageSelector()
