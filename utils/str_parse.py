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
    If no comma found, return empty string (per spec)."""
    if not text:
        return ""
    if "," in text:
        return text.split(",", 1)[1].strip()
    return ""


def before_v(text: str) -> str:
    """Return text before 'V.' (V followed by dot), case-insensitive.
    If no 'V.' found, return the original text stripped."""
    if not text:
        return ""
    # Look for 'V.' pattern (case-insensitive) — this is the vessel version marker
    match = re.search(r'[Vv]\.', text)
    if match:
        return text[:match.start()].strip()
    return text.strip()


def replace_line_break(text: str, replacement: str = "/") -> str:
    """Replace line breaks with the given replacement string."""
    if not text:
        return ""
    return text.replace("\n", replacement).replace("\r", "")
