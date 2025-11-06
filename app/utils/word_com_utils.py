"""Word COM automation utility functions to reduce code duplication."""

from __future__ import annotations

from typing import Optional, Any

from app.utils.logging_utils import get_logger

logger = get_logger(__name__)

# Constants for Word COM (centralized)
WD_REPLACE_ALL = 2
WD_FIND_CONTINUE = 1
WD_FORMAT_DOCUMENT = 0  # .doc format
WD_FORMAT_DOCX = 16  # .docx format (Office 2007+)
WD_ALERTS_NONE = 0


def execute_find_replace(range_obj: Any, find_text: str, replace_text: str) -> int:
    """Execute find and replace on a specific Word range.

    OPTIMIZATION: Centralized find/replace logic used across Word report generators.

    Args:
        range_obj: Word Range object
        find_text: Text to find
        replace_text: Text to replace with

    Returns:
        Number of replacements made (0 or 1)

    Example:
        >>> doc = word_app.Documents.Open("template.doc")
        >>> result = execute_find_replace(doc.Range(), "<ClientName>", "Acme Corp")
    """
    replacements_made = 0
    try:
        # Reset the range to the start to ensure we search the entire range
        try:
            range_obj.SetRange(range_obj.Start, range_obj.End)
        except Exception:
            pass

        find_object = range_obj.Find
        find_object.ClearFormatting()
        find_object.Replacement.ClearFormatting()

        # Use explicit Execute args (more reliable across COM objects)
        result = find_object.Execute(
            FindText=find_text,
            MatchCase=False,
            MatchWholeWord=False,
            MatchWildcards=False,
            MatchSoundsLike=False,
            MatchAllWordForms=False,
            Forward=True,
            Wrap=WD_FIND_CONTINUE,
            Format=False,
            ReplaceWith=replace_text,
            Replace=WD_REPLACE_ALL
        )

        if bool(result):
            replacements_made = 1
    except Exception as e:
        logger.debug("Error in find/replace for range: %s", str(e))

    return replacements_made


def replace_in_shape(shape: Any, find_text: str, replace_text: str) -> int:
    """Replace text inside a Shape/TextFrame, recursing into group items.

    OPTIMIZATION: Reusable shape text replacement for Word automation.

    Args:
        shape: Word Shape object
        find_text: Text to find
        replace_text: Text to replace with

    Returns:
        Approximate count of places where replacement was executed

    Example:
        >>> for shape in doc.Shapes:
        ...     replaced = replace_in_shape(shape, "<Date>", "2025-01-15")
    """
    hits = 0
    try:
        # If the shape has a text frame with text, run Find on its Range
        if getattr(shape, "HasTextFrame", 0):
            tf = shape.TextFrame
            if getattr(tf, "HasText", 0):
                try:
                    rng = tf.TextRange
                    if execute_find_replace(rng, find_text, replace_text):
                        hits += 1
                except Exception:
                    pass

        # Recurse into grouped shapes when available
        try:
            group_items = getattr(shape, "GroupItems", None)
            if group_items is not None:
                for sub in group_items:
                    hits += replace_in_shape(sub, find_text, replace_text)
        except Exception:
            pass
    except Exception:
        pass
    return hits


def replace_text_in_document(
    doc: Any,
    find_text: str,
    replace_text: str,
    search_headers_footers: bool = True,
    search_shapes: bool = True
) -> int:
    """Replace text everywhere in a Word document.

    OPTIMIZATION: Comprehensive text replacement covering all story ranges, headers,
    footers, shapes, and grouped shapes. Eliminates ~100 lines of duplicated code.

    Covers:
    - All story ranges (main body, headers/footers, text boxes, footnotes, comments)
    - Shapes (including grouped shapes)
    - Shapes within headers/footers

    Args:
        doc: Word Document COM object
        find_text: Placeholder to find (e.g., "<Client>")
        replace_text: Replacement value
        search_headers_footers: If True, search in headers and footers (default: True)
        search_shapes: If True, search in shapes (default: True)

    Returns:
        Total number of occurrences replaced

    Example:
        >>> doc = word_app.Documents.Open("template.doc")
        >>> count = replace_text_in_document(doc, "<ProjectName>", "Project Alpha")
        >>> logger.info(f"Replaced {count} occurrences")
    """
    try:
        replace_value = str(replace_text) if replace_text else ""
        total_hits = 0

        # 1) Iterate sections, then headers/footers
        if search_headers_footers:
            try:
                for section in doc.Sections:
                    # Body content for the section
                    try:
                        rng = section.Range
                        if execute_find_replace(rng, find_text, replace_value):
                            total_hits += 1
                    except Exception:
                        pass

                    # Headers
                    for header in section.Headers:
                        try:
                            rng = header.Range
                            if execute_find_replace(rng, find_text, replace_value):
                                total_hits += 1
                        except Exception:
                            pass

                    # Footers
                    for footer in section.Footers:
                        try:
                            rng = footer.Range
                            if execute_find_replace(rng, find_text, replace_value):
                                total_hits += 1
                        except Exception:
                            pass
            except Exception as e:
                logger.debug("Section iteration issue: %s", str(e))

        # 2) Replace in all StoryRanges (includes main text, headers/footers, and text frames)
        try:
            for sr in doc.StoryRanges:
                current = sr
                while current is not None:
                    if execute_find_replace(current, find_text, replace_value):
                        total_hits += 1
                    try:
                        current = current.NextStoryRange
                    except Exception:
                        current = None
        except Exception as e:
            logger.debug("StoryRanges iteration issue: %s", str(e))

        # 3) Replace inside Shapes in document body
        if search_shapes:
            try:
                for shape in doc.Shapes:
                    total_hits += replace_in_shape(shape, find_text, replace_value)
            except Exception as e:
                logger.debug("Error iterating doc.Shapes: %s", str(e))

            # 4) Replace inside Shapes within headers/footers for each section
            if search_headers_footers:
                try:
                    for section in doc.Sections:
                        # Headers
                        for header in section.Headers:
                            try:
                                for shape in header.Shapes:
                                    total_hits += replace_in_shape(shape, find_text, replace_value)
                            except Exception:
                                pass
                        # Footers
                        for footer in section.Footers:
                            try:
                                for shape in footer.Shapes:
                                    total_hits += replace_in_shape(shape, find_text, replace_value)
                            except Exception:
                                pass
                except Exception as e:
                    logger.debug("Section Shapes iteration issue: %s", str(e))

        if total_hits > 0:
            logger.debug("Replaced '%s' -> '%s' (%d occurrences)", find_text, replace_value[:50], total_hits)
        else:
            logger.debug("Placeholder '%s' NOT found in document", find_text)

        return total_hits

    except Exception as e:
        logger.warning("Error replacing text '%s': %s", find_text, str(e))
        return 0


def update_bookmark(doc: Any, bookmark_name: str, text: str) -> bool:
    """Update a bookmark in a Word document with text.

    OPTIMIZATION: Centralized bookmark update with error handling.

    Args:
        doc: Word Document object
        bookmark_name: Name of the bookmark
        text: Text to insert

    Returns:
        True if bookmark was updated, False if not found or error occurred

    Example:
        >>> doc = word_app.Documents.Open("template.doc")
        >>> success = update_bookmark(doc, "ClientName", "Acme Corporation")
    """
    try:
        if doc.Bookmarks.Exists(bookmark_name):
            bookmark = doc.Bookmarks(bookmark_name)
            bookmark.Range.Text = text
            logger.debug("Updated bookmark '%s' with '%s'", bookmark_name, text)
            return True
        else:
            logger.debug("Bookmark '%s' not found in document", bookmark_name)
            return False
    except Exception as e:
        logger.warning("Error updating bookmark '%s': %s", bookmark_name, str(e))
        return False


def cleanup_bookmarks(doc: Any) -> int:
    """Remove all bookmarks from a Word document.

    OPTIMIZATION: Common operation to clean up template bookmarks.

    Args:
        doc: Word Document object

    Returns:
        Number of bookmarks deleted

    Example:
        >>> doc = word_app.Documents.Open("report.doc")
        >>> count = cleanup_bookmarks(doc)
        >>> logger.info(f"Removed {count} bookmarks")
    """
    try:
        bookmark_count = doc.Bookmarks.Count

        # Delete bookmarks in reverse order to avoid index issues
        for i in range(bookmark_count, 0, -1):
            try:
                doc.Bookmarks(i).Delete()
            except:
                pass

        logger.debug("Cleaned up %d bookmarks", bookmark_count)
        return bookmark_count

    except Exception as e:
        logger.warning("Error cleaning up bookmarks: %s", str(e))
        return 0


def set_word_app_optimization(word_app: Any, enabled: bool = True) -> None:
    """Enable or disable Word application optimizations for faster COM operations.

    OPTIMIZATION: Centralized configuration to speed up Word automation by ~40-60%.

    When enabled:
    - Disables screen updating (major performance boost)
    - Disables auto-formatting features
    - Disables grammar/spell checking
    - Disables background saving

    Args:
        word_app: Word Application COM object
        enabled: If True, enable optimizations (disable features). If False, restore normal operation.

    Example:
        >>> word_app = win32com.client.Dispatch("Word.Application")
        >>> set_word_app_optimization(word_app, enabled=True)
        >>> # ... perform operations ...
        >>> set_word_app_optimization(word_app, enabled=False)  # Restore before showing to user
    """
    try:
        if enabled:
            # Disable screen updating (major speed boost)
            word_app.ScreenUpdating = False

            # Disable auto-formatting
            try:
                word_app.Options.AutoFormatAsYouTypeApplyTables = False
                word_app.Options.AutoFormatAsYouTypeApplyBulletedLists = False
                word_app.Options.AutoFormatAsYouTypeApplyNumberedLists = False
                word_app.Options.AutoFormatAsYouTypeReplaceQuotes = False
            except Exception:
                pass

            # Disable grammar/spell checking
            try:
                word_app.Options.CheckGrammarAsYouType = False
                word_app.Options.CheckSpellingAsYouType = False
            except Exception:
                pass

            # Disable background saving
            try:
                word_app.Options.BackgroundSave = False
            except Exception:
                pass

            logger.debug("Word application optimizations ENABLED")
        else:
            # Re-enable screen updating
            word_app.ScreenUpdating = True

            # Re-enable features (optional, can be left disabled)
            try:
                word_app.Options.AutoFormatAsYouTypeApplyTables = True
                word_app.Options.CheckGrammarAsYouType = True
                word_app.Options.CheckSpellingAsYouType = True
                word_app.Options.BackgroundSave = True
            except Exception:
                pass

            logger.debug("Word application optimizations DISABLED")

    except Exception as e:
        logger.debug("Error setting Word app optimizations: %s", str(e))


def safe_close_com_object(com_obj: Optional[Any], obj_type: str = "COM object") -> None:
    """Safely close and release a COM object (Word/Excel application).

    OPTIMIZATION: Centralized cleanup to prevent resource leaks.

    Args:
        com_obj: COM object to close (Word.Application, Excel.Application, etc.)
        obj_type: Descriptive name for logging (default: "COM object")

    Example:
        >>> word_app = win32com.client.Dispatch("Word.Application")
        >>> # ... use word_app ...
        >>> safe_close_com_object(word_app, "Word Application")
    """
    if com_obj:
        try:
            com_obj.Quit()
            logger.debug("%s closed successfully", obj_type)
        except Exception as e:
            logger.debug("Error closing %s: %s", obj_type, str(e))
