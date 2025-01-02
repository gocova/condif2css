# Copyright (c) 2025 Jose Gonzalo Covarrubias M <gocova.dev@gmail.com>
#
# Part of: batch_xlsx2html (bxx2html)
#
#
# Acknowledge:
#   Syed Bashar Milton
from colorsys import rgb_to_hls, hls_to_rgb

from openpyxl.styles.colors import aRGB_REGEX
# https://bitbucket.org/openpyxl/openpyxl/issues/987/add-utility-functions-for-colors-to-help

RGBMAX = 0xFF  # Corresponds to 255
HLSMAX = 240  # MS excel's tint function expects that HLS is base 240. see:
# https://social.msdn.microsoft.com/Forums/en-US/e9d8c136-6d62-4098-9b1b-dac786149f43/excel-color-tint-algorithm-incorrect?forum=os_binaryfile#d3c2ac95-52e0-476b-86f1-e2a697f24969


def aRGB_to_ms_hls(aRGB: str) -> tuple[int, int, int]:
    """Converts a hex string of the form '[aa]rrggbb' to HLSMAX based HLS, (alpha values are ignored)"""
    if isinstance(aRGB, str):
        m = aRGB_REGEX.match(aRGB)
        if m is None:
            raise ValueError("Colors must be aRGB hex values")
        if len(aRGB) > 6:
            color_str = aRGB[-6:]  # Ignore alpha values
        else:
            color_str = aRGB
        blue = int(color_str[4:], 16) / RGBMAX
        green = int(color_str[2:4], 16) / RGBMAX
        red = int(color_str[0:2], 16) / RGBMAX
        return rgb_to_ms_hls(red, green, blue)

    else:
        raise TypeError("aRGB arg shoud be an str")


def aRGB_to_css(aRGB: str) -> str:
    """Converts a hex string of the form [aa]rrggbb to CSS color string"""
    if isinstance(aRGB, str):
        m = aRGB_REGEX.match(aRGB)
        if m is None:
            raise ValueError("Colors must be aRGB hex values")
        if len(aRGB) == 6:
            return f"#{aRGB}"
        if aRGB.startswith("00"):
            return f"#{aRGB[-6:]}"
        blue = int(aRGB[6:], 16)
        green = int(aRGB[4:6], 16)
        red = int(aRGB[2:4], 16)
        alpha = int(aRGB[0:2], 16)
        return f"rgba({red}, {green}, {blue}, {alpha})"

    else:
        raise TypeError("aRGB arg shoud be an str")


def rgb_to_ms_hls(red: float, green: float, blue: float) -> tuple[int, int, int]:
    """Converts rgb values in range (0,1) to HLSMAX based HLS"""
    h, l, s = rgb_to_hls(red, green, blue)
    return (int(round(h * HLSMAX)), int(round(l * HLSMAX)), int(round(s * HLSMAX)))


# def ms_hls_to_rgb(hue, lightness=None, saturation=None):
def ms_hls_to_rgb(
    hue: int, lightness: int, saturation: int
) -> tuple[float, float, float]:
    """Converts HLSMAX based HLS values to rgb values in the range (0,1)"""
    # if lightness is None:
    #     hue, lightness, saturation = hue
    return hls_to_rgb(hue / HLSMAX, lightness / HLSMAX, saturation / HLSMAX)


# def rgb_to_hex(red, green=None, blue=None):
#     """Converts (0,1) based RGB values to a hex string 'rrggbb'"""
#     if green is None:
#         red, green, blue = red
def rgb_to_hex(red: float, green: float, blue: float) -> str:
    """Converts (0,1) based RGB values to a hex string 'rrggbb'"""
    return (
        "%02x%02x%02x"
        % (
            int(round(red * RGBMAX)),
            int(round(green * RGBMAX)),
            int(round(blue * RGBMAX)),
        )
    ).upper()


def tint_luminance(tint: float | None, lum: float) -> int:
    """Tints a HLSMAX based luminance"""
    # See: http://ciintelligence.blogspot.co.uk/2012/02/converting-excel-theme-color-and-tint.html

    # int() is not required due to round implementation
    return round(
        (
            (lum * (1.0 + tint))
            if tint < 0
            else (lum * (1.0 - tint) + (HLSMAX - HLSMAX * (1.0 - tint)))
        )
        if tint is not None
        else (lum)
    )
