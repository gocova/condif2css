# Copyright (c) 2025 Jose Gonzalo Covarrubias M <gocova.dev@gmail.com>
#
# Part of: batch_xlsx2html (bxx2html)
#
# Aknowledge:
#   - Inspiration from xlsx2html, extended with theme support and aRGB colors

from openpyxl.styles.colors import COLOR_INDEX, Color, aRGB_REGEX
from .color import (
    aRGB_to_ms_hls,
    rgb_to_hex,
    ms_hls_to_rgb,
    tint_luminance,
)


def create_themed_get_css_color(theme_aRGBs_list: list[str]):
    """Create a get_css_color based on provided theme"""
    if theme_aRGBs_list is None or (
        isinstance(theme_aRGBs_list, list) and len(theme_aRGBs_list) < 2
    ):
        theme_aRGBs_list = ["FFFFFF", "000000"]
    theme_len = len(theme_aRGBs_list)

    def cova__get_css_color(color: Color):
        rgb = None

        if color.type == "theme":
            if (
                isinstance(color.value, int)
                and color.value >= 0
                and color.value < theme_len
            ):
                rgb_base: str = theme_aRGBs_list[color.value]
                if color.tint == 0.0:
                    rgb = f"00{rgb_base}"
                else:
                    h_part, l_part, s_part = aRGB_to_ms_hls(rgb_base)
                    rgb = f"00{rgb_to_hex(*ms_hls_to_rgb(h_part, tint_luminance(color.tint, l_part), s_part))}"

            else:
                rgb = "00000000"

        elif color.type == "rgb":
            rgb = color.rgb

        if color.type == "indexed":
            # Reference: https://openpyxl.readthedocs.io/en/stable/styles.html#indexed-colours
            if color.indexed > 0:
                if color.indexed < 64:
                    rgb = COLOR_INDEX[color.indexed]

                # The indices 64 and 65 are reserved for the system
                # foreground and background colours respectively

                elif color.indexed == 64:
                    rgb = theme_aRGBs_list[1]  # 'dk1' | windowText
                elif color.indexed == 65:
                    rgb = theme_aRGBs_list[0]  # 'lt1' | window
            rgb = "00000000" if not rgb or not aRGB_REGEX.match(rgb) else rgb

        return rgb if isinstance(rgb, str) else None

    return cova__get_css_color
