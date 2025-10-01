"""Unit tests for shared path utility helpers."""

from pathlib import Path
import unittest

from app.config.settings import settings
from app.utils import path_utils


class TestSanitizeFolderName(unittest.TestCase):
    """Tests covering folder name sanitization logic."""

    def test_removes_invalid_characters_and_normalizes(self) -> None:
        raw_name = " Client: Demo<> /Project? *2025 "
        sanitized = path_utils.sanitize_folder_name(raw_name)
        self.assertEqual(sanitized, "CLIENT_DEMO_PROJECT_2025")

    def test_returns_default_when_result_empty(self) -> None:
        raw_name = "<>:/\\|?*"
        sanitized = path_utils.sanitize_folder_name(raw_name)
        self.assertEqual(sanitized, "UNNAMED_PROJECT")


class TestResolveProjectOutput(unittest.TestCase):
    """Tests for filesystem path resolution helpers."""

    def test_uses_project_base_dir_when_type_valid(self) -> None:
        output_dir, used_fallback = path_utils.resolve_project_output(
            "Launch Readout",
            settings.PROJECT_TYPE_NW,
        )

        expected_dir = settings.base_dir_nw / "LAUNCH_READOUT"
        self.assertEqual(output_dir, expected_dir)
        self.assertFalse(used_fallback)

    def test_uses_fallback_when_type_invalid(self) -> None:
        output_dir, used_fallback = path_utils.resolve_project_output(
            "Mystery",
            "unknown",
        )

        expected_dir = settings.nw_files_dir / "temp_slides" / "MYSTERY"
        self.assertEqual(output_dir, expected_dir)
        self.assertTrue(used_fallback)

    def test_uses_fallback_when_type_missing(self) -> None:
        output_dir, used_fallback = path_utils.resolve_project_output(
            "No Type",
            None,
        )

        expected_dir = settings.nw_files_dir / "temp_slides" / "NO_TYPE"
        self.assertEqual(output_dir, expected_dir)
        self.assertTrue(used_fallback)


class TestBuildRelativeSlidePath(unittest.TestCase):
    """Tests for relative slide path construction."""

    def test_builds_path_for_known_project_type(self) -> None:
        relative = path_utils.build_relative_slide_path(
            settings.PROJECT_TYPE_BIPRESENTS,
            "Warm Up",
            "slide_1.jpg",
        )
        self.assertEqual(relative, "bsr_slides/WARM_UP/slide_1.jpg")

    def test_builds_path_with_subdir_and_fallback_root(self) -> None:
        relative = path_utils.build_relative_slide_path(
            None,
            "Fallback Test",
            "slide.png",
            subdir="thumbs",
            fallback_root="temp_slides",
        )
        self.assertEqual(relative, "temp_slides/FALLBACK_TEST/thumbs/slide.png")

    def test_sanitizes_display_name_in_relative_path(self) -> None:
        relative = path_utils.build_relative_slide_path(
            settings.PROJECT_TYPE_NW,
            " Client<> Showcase ",
            "slide.png",
        )
        self.assertEqual(relative, "nw_slides/CLIENT_SHOWCASE/slide.png")


if __name__ == "__main__":
    unittest.main()
