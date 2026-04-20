"""UI palette + typography + button-variant helpers.

Self-contained so the project never depends on the external ``design_kit``
reference folder. Dark graphite background, elevated panels, cyan primary,
rimmed buttons, Segoe-based typography.
"""

from __future__ import annotations

from typing import Any

PALETTE_DARK: dict[str, str] = {
    "bg": "#0c0c12",
    "panel": "#14141c",
    "panel_elev": "#1a1a26",
    "border": "#2a2a3a",
    "text": "#f0f0f8",
    "muted": "#7a7a92",
    "cyan": "#00c8ff",
    "cyan_dim": "#006080",
    "cyan_hover": "#33d6ff",
    "gold": "#c9a227",
    "gold_dim": "#6b5a12",
    "stop": "#e53935",
    "btn_rim": "#050508",
    "primary_border": "#021a22",
}

PALETTE_LIGHT: dict[str, str] = {
    "bg": "#e4e9f2",
    "panel": "#f5f7fb",
    "panel_elev": "#ffffff",
    "border": "#b4bccf",
    "text": "#10141c",
    "muted": "#4a5568",
    "cyan": "#007ea3",
    "cyan_dim": "#0096c7",
    "cyan_hover": "#23b2e0",
    "gold": "#a67c00",
    "gold_dim": "#c9a227",
    "stop": "#c62828",
    "btn_rim": "#1e2430",
    "primary_border": "#005a75",
}

# Geometry.
BTN_RADIUS = 10
BTN_H = 36

# Typography.
FONT_APP_TITLE = ("Segoe UI Black", 18)
FONT_SECTION = ("Segoe UI Semibold", 15)
FONT_UI = ("Segoe UI", 14)
FONT_UI_SM = ("Segoe UI", 12)
FONT_HINT = ("Segoe UI", 11)
FONT_BTN = ("Segoe UI Black", 10)
FONT_BTN_PRIMARY = ("Segoe UI Black", 11)
FONT_BTN_NAV = ("Segoe UI Semibold", 10)


def button_kwargs(
    palette: dict[str, str],
    variant: str = "ghost",
    *,
    height: int = BTN_H,
    font: tuple | None = None,
    width: int | None = None,
) -> dict[str, Any]:
    """Return a ``CTkButton(**kwargs)`` dict for the requested variant.

    Central helper (do not pass any of these keys a second time at the
    call-site - CustomTkinter raises ``TypeError`` on duplicate kwargs).
    """
    kw: dict[str, Any] = dict(
        corner_radius=BTN_RADIUS,
        font=font or FONT_BTN,
        height=height,
        border_width=2,
        border_color=palette["btn_rim"],
    )
    if width is not None:
        kw["width"] = width

    if variant == "ghost":
        kw.update(
            fg_color=palette["panel_elev"],
            hover_color=palette["border"],
            text_color=palette["text"],
        )
    elif variant == "primary":
        kw.update(
            fg_color=palette["cyan_dim"],
            hover_color=palette["cyan"],
            text_color=palette["text"],
            border_color=palette["primary_border"],
        )
    elif variant == "primary_emphasis":
        kw.update(
            fg_color=palette["cyan_dim"],
            hover_color=palette["cyan"],
            text_color=palette["text"],
            border_color=palette["primary_border"],
            font=FONT_BTN_PRIMARY,
        )
    elif variant == "gold":
        kw.update(
            fg_color=palette["gold_dim"],
            hover_color=palette["gold"],
            text_color=palette["text"],
        )
    elif variant == "danger_soft":
        kw.update(
            fg_color=palette["panel_elev"],
            hover_color=palette["stop"],
            text_color=palette["text"],
        )
    elif variant == "nav_idle":
        kw.update(
            fg_color=palette["panel_elev"],
            hover_color=palette["border"],
            text_color=palette["muted"],
            font=FONT_BTN_NAV,
        )
    elif variant == "nav_active":
        kw.update(
            fg_color=palette["cyan_dim"],
            hover_color=palette["cyan"],
            text_color=palette["text"],
            border_color=palette["cyan"],
            font=FONT_BTN,
        )
    return kw
