# Copyright (c) 2026 Jose Gonzalo Covarrubias M <gocova.dev@gmail.com>
#
# Part of: batch_xlsx2html (bxx2html)
#
#
# Acknowledge:
#   - Inspiration from xlsx2html, extended with theme support and aRGB colors

from openpyxl.styles.colors import COLOR_INDEX, Color, aRGB_REGEX
from .color import (
    argb_to_ms_hls,
    rgb_to_hex,
    ms_hls_to_rgb,
    tint_luminance,
)


def create_themed_css_color_resolver(theme_argbs_list: list[str]):
    """
    Creates a function that returns the CSS color string representation of the given color.

    If the color is a theme color, it will be resolved to its corresponding RGB value.
    If the color is an indexed color, it will be resolved to its corresponding RGB value.
    If the color is an RGB color, it will be returned as is.

    :param theme_argbs_list: A list of aRGB values for the theme colors.
    :return: A function that takes a Color and returns its CSS color string representation, or None if the color is not valid
    """
    if theme_argbs_list is None or (
        isinstance(theme_argbs_list, list) and len(theme_argbs_list) < 2
    ):
        theme_argbs_list = ["FFFFFF", "000000"]
    theme_len = len(theme_argbs_list)

    def get_css_color(color: Color):
        """
        Returns the CSS color string representation of the given color.

        If the color is a theme color, it will be resolved to its corresponding RGB value.
        If the color is an indexed color, it will be resolved to its corresponding RGB value.
        If the color is an RGB color, it will be returned as is.

        :param color: The color to be resolved
        :return: The CSS color string representation of the given color, or None if the color is not valid
        """
        rgb = None

        if color.type == "theme":
            if (
                isinstance(color.value, int)
                and color.value >= 0
                and color.value < theme_len
            ):
                rgb_base: str = theme_argbs_list[color.value]
                if color.tint == 0.0:
                    rgb = f"00{rgb_base}"
                else:
                    h_part, l_part, s_part = argb_to_ms_hls(rgb_base)
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
                    rgb = theme_argbs_list[1]  # 'dk1' | windowText
                elif color.indexed == 65:
                    rgb = theme_argbs_list[0]  # 'lt1' | window
            rgb = "00000000" if not rgb or not aRGB_REGEX.match(rgb) else rgb

        return rgb if isinstance(rgb, str) else None

    return get_css_color
