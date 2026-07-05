"""Centralized string parsing utilities for invoice data extraction.

All string parsing logic used by rules must live here.
Rules must never parse strings themselves — call these functions instead.
"""

import re


def after_colon(text: str) -> str:
    """Return text after the first colon, stripped.
    If no colon found, return the original text stripped."""
    if not text:
        return ""
    if ":" in text:
        return text.split(":", 1)[1].strip()
    return text.strip()


def before_comma_after_colon(text: str) -> str:
    """Return text after colon, then take text before the first comma."""
    after = after_colon(text)
    if "," in after:
        return after.split(",", 1)[0].strip()
    return after


def after_comma(text: str) -> str:
    """Return text after the first comma.
    Assumes the caller has already extracted the relevant portion
    (typically after colon)."""
    if not text:
        return ""
    if "," in text:
        return text.split(",", 1)[1].strip()
    return text.strip()


def before_v(text: str) -> str:
    """Return text before 'V' or 'v' (case-insensitive), stripped."""
    if not text:
        return ""
    parts = re.split(r'[Vv]', text, maxsplit=1)
    return parts[0].strip()


def replace_line_break(text: str, replacement: str = "/") -> str:
    """Replace line breaks with the given replacement string."""
    if not text:
        return ""
    return text.replace("\n", replacement).replace("\r", "")
