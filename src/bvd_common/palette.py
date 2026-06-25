"""Shared colour palette for BVD data-analysis decks.

Base neutrals + brand colours reused across every skill. Skill-specific
*semantic* colours (e.g. "child red", "adult teal") live in the skill that
needs them, not here — promote a colour to this module only once a second
skill uses it.
"""
from pptx.dml.color import RGBColor


def rgb(h):
    """Hex string -> RGBColor, e.g. rgb("0B2D4E")."""
    return RGBColor(int(h[:2], 16), int(h[2:4], 16), int(h[4:], 16))


# ── Base neutrals / brand ─────────────────────────────────────────────────────
NAVY   = rgb("0B2D4E")
NAVY_D = rgb("0D3A5E")
SLATE  = rgb("F4F6F8")
DARK   = rgb("1E293B")
LGRAY  = rgb("94A3B8")
WHITE  = rgb("FFFFFF")
TEAL   = rgb("0D6B7A")
CORAL  = rgb("C0392B")
AMBER  = rgb("D97706")
GREEN  = rgb("1A7A4A")
ICE    = rgb("CADCFC")
