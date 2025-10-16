"""Utility functions for NW data transformation and processing."""

from __future__ import annotations

import re
from typing import List, Tuple

from app.models.nw_reports_models import VoteValue
from app.utils.logging_utils import get_logger

logger = get_logger(__name__)

# Delimiters for grouped names
NAME_GROUP_DELIMITERS = ["##", "$$"]


def convert_vote_to_text(vote_value: int) -> str:
    """Convert numeric vote value to text representation.

    Args:
        vote_value: Numeric vote (-1, 0, 1)

    Returns:
        String representation ("Negative", "Neutral", "Positive")
    """
    if vote_value == -1:
        return VoteValue.NEGATIVE.value
    elif vote_value == 0:
        return VoteValue.NEUTRAL.value
    elif vote_value == 1:
        return VoteValue.POSITIVE.value
    else:
        logger.warning("Unknown vote value: %s, defaulting to Neutral", vote_value)
        return VoteValue.NEUTRAL.value


def split_grouped_names(name_string: str) -> List[str]:
    """Split a delimited name string into individual names.

    Looks for ## or $$ delimiters to identify grouped names.

    Args:
        name_string: String potentially containing delimited names

    Returns:
        List of individual names
    """
    if not name_string:
        return []

    # Check which delimiter is present
    delimiter = None
    for delim in NAME_GROUP_DELIMITERS:
        if delim in name_string:
            delimiter = delim
            break

    if delimiter:
        # Split and filter out empty strings
        names = [n.strip() for n in name_string.split(delimiter) if n.strip()]
        logger.debug(
            "Split grouped name '%s' into %d names using delimiter '%s'",
            name_string[:50],
            len(names),
            delimiter,
        )
        return names
    else:
        # No delimiter found, return as single name
        return [name_string.strip()] if name_string.strip() else []


def is_grouped_name(name_string: str) -> bool:
    """Check if a name string contains group delimiters.

    Args:
        name_string: String to check

    Returns:
        True if string contains ## or $$ delimiters
    """
    if not name_string:
        return False

    return any(delim in name_string for delim in NAME_GROUP_DELIMITERS)


def convert_unicode_to_text(text: str) -> str:
    """Convert Unicode characters to readable text representation.

    This is a simplified version. The original VB.NET function (fnConvertUnicodeToText)
    had more complex logic for specific character mappings.

    Args:
        text: Text potentially containing Unicode characters

    Returns:
        Converted text with readable characters
    """
    if not text:
        return ""

    # Basic Unicode normalization
    # You can extend this with specific character mappings as needed
    try:
        # Normalize Unicode to closest ASCII equivalent
        import unicodedata
        normalized = unicodedata.normalize('NFKD', text)
        # Try to encode to ASCII, replacing non-ASCII characters
        ascii_text = normalized.encode('ascii', 'ignore').decode('ascii')

        if ascii_text != text:
            logger.debug(
                "Converted Unicode text (length: %d chars changed)",
                len(text) - len(ascii_text)
            )

        return ascii_text if ascii_text else text
    except Exception as e:
        logger.warning("Failed to convert Unicode text: %s", e)
        return text


def split_name_rationale_column(rationale: str, max_part1_length: int = 100) -> Tuple[str, str]:
    """Split a rationale column into two parts for Word document layout.

    The original VB.NET code splits the NameRationale column into two parts
    for better presentation in Word tables.

    Args:
        rationale: The rationale text to split
        max_part1_length: Maximum length for the first part

    Returns:
        Tuple of (part1, part2)
    """
    if not rationale:
        return ("", "")

    if len(rationale) <= max_part1_length:
        return (rationale, "")

    # Try to split at a sentence boundary
    # Look for '. ' or '! ' or '? ' near the middle
    split_points = [
        rationale.find('. ', max_part1_length // 2, max_part1_length),
        rationale.find('! ', max_part1_length // 2, max_part1_length),
        rationale.find('? ', max_part1_length // 2, max_part1_length),
    ]

    valid_split_points = [p for p in split_points if p > 0]

    if valid_split_points:
        split_index = min(valid_split_points) + 2  # Include the punctuation and space
        return (rationale[:split_index].strip(), rationale[split_index:].strip())

    # No sentence boundary found, split at word boundary near max_length
    if len(rationale) > max_part1_length:
        # Find last space before max_part1_length
        last_space = rationale.rfind(' ', 0, max_part1_length)
        if last_space > 0:
            return (rationale[:last_space].strip(), rationale[last_space:].strip())

    # Fallback: hard split
    return (rationale[:max_part1_length].strip(), rationale[max_part1_length:].strip())


def clean_html_for_word(html_content: str) -> str:
    """Clean and prepare HTML content for insertion into Word documents.

    Args:
        html_content: HTML string

    Returns:
        Cleaned HTML suitable for Word
    """
    if not html_content:
        return ""

    # Remove problematic HTML tags/attributes
    # This is a simplified version - extend as needed
    cleaned = html_content

    # Remove style attributes that might conflict
    cleaned = re.sub(r'\sstyle="[^"]*"', '', cleaned)

    # Remove script tags
    cleaned = re.sub(r'<script[^>]*>.*?</script>', '', cleaned, flags=re.DOTALL)

    # Normalize whitespace
    cleaned = re.sub(r'\s+', ' ', cleaned)

    return cleaned.strip()


def format_percentage(value: float, decimals: int = 1) -> str:
    """Format a float value as a percentage string.

    Args:
        value: Float value (0.0 to 1.0 or 0 to 100)
        decimals: Number of decimal places

    Returns:
        Formatted percentage string (e.g., "85.5%")
    """
    if value is None:
        return "0.0%"

    # If value is already in percentage form (> 1), use as is
    if value > 1:
        percentage = value
    else:
        percentage = value * 100

    return f"{percentage:.{decimals}f}%"


def sanitize_filename(filename: str) -> str:
    """Sanitize a filename by removing invalid characters.

    Args:
        filename: Original filename

    Returns:
        Sanitized filename safe for file system
    """
    if not filename:
        return "unnamed"

    # Remove invalid filename characters
    invalid_chars = r'[<>:"/\\|?*]'
    sanitized = re.sub(invalid_chars, '_', filename)

    # Remove leading/trailing spaces and dots
    sanitized = sanitized.strip('. ')

    # Limit length
    max_length = 200
    if len(sanitized) > max_length:
        sanitized = sanitized[:max_length]

    return sanitized if sanitized else "unnamed"


def extract_grouped_rationales(rationale_string: str) -> List[str]:
    """Extract individual rationales from a delimited rationale string.

    Similar to split_grouped_names but for rationales.

    Args:
        rationale_string: String potentially containing delimited rationales

    Returns:
        List of individual rationales
    """
    if not rationale_string:
        return []

    # Check which delimiter is present
    delimiter = None
    for delim in NAME_GROUP_DELIMITERS:
        if delim in rationale_string:
            delimiter = delim
            break

    if delimiter:
        # Split and filter out empty strings
        rationales = [r.strip() for r in rationale_string.split(delimiter) if r.strip()]
        return rationales
    else:
        # No delimiter found, return as single rationale
        return [rationale_string.strip()] if rationale_string.strip() else []


def merge_names_with_delimiter(names: List[str], delimiter: str = "##") -> str:
    """Merge a list of names with the specified delimiter.

    Args:
        names: List of names to merge
        delimiter: Delimiter to use (default: ##)

    Returns:
        Merged string with trailing delimiter
    """
    if not names:
        return ""

    # Filter out empty names
    valid_names = [n.strip() for n in names if n and n.strip()]

    if not valid_names:
        return ""

    # Join with delimiter and add trailing delimiter
    return delimiter.join(valid_names) + delimiter


def parse_project_display_name(display_name: str) -> Tuple[str, str]:
    """Parse a display name to extract project name and identifier.

    Many display names follow patterns like "ProjectName_Identifier" or
    "ProjectName (Identifier)".

    Args:
        display_name: The display name to parse

    Returns:
        Tuple of (project_name, identifier)
    """
    if not display_name:
        return ("", "")

    # Try pattern: "Name_Identifier"
    if '_' in display_name:
        parts = display_name.split('_', 1)
        return (parts[0].strip(), parts[1].strip() if len(parts) > 1 else "")

    # Try pattern: "Name (Identifier)"
    match = re.match(r'^(.+?)\s*\((.+)\)$', display_name)
    if match:
        return (match.group(1).strip(), match.group(2).strip())

    # No pattern matched, return full name as project name
    return (display_name.strip(), "")


def calculate_vote_statistics(
    positive_count: int,
    neutral_count: int,
    negative_count: int
) -> dict:
    """Calculate vote statistics from counts.

    Args:
        positive_count: Number of positive votes
        neutral_count: Number of neutral votes
        negative_count: Number of negative votes

    Returns:
        Dictionary with statistics including percentages
    """
    total = positive_count + neutral_count + negative_count

    if total == 0:
        return {
            "total": 0,
            "positive": 0,
            "neutral": 0,
            "negative": 0,
            "positive_pct": "0.0%",
            "neutral_pct": "0.0%",
            "negative_pct": "0.0%",
        }

    positive_pct = (positive_count / total) * 100
    neutral_pct = (neutral_count / total) * 100
    negative_pct = (negative_count / total) * 100

    return {
        "total": total,
        "positive": positive_count,
        "neutral": neutral_count,
        "negative": negative_count,
        "positive_pct": f"{positive_pct:.1f}%",
        "neutral_pct": f"{neutral_pct:.1f}%",
        "negative_pct": f"{negative_pct:.1f}%",
    }
