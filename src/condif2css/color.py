# Copyright (c) 2026 Jose Gonzalo Covarrubias M <gocova.dev@gmail.com>
#
# Part of: batch_xlsx2html (bxx2html)
#
#
# Acknowledge:
#   Syed Bashar Milton
#
from colorsys import rgb_to_hls, hls_to_rgb

from openpyxl.styles.colors import aRGB_REGEX
# https://bitbucket.org/openpyxl/openpyxl/issues/987/add-utility-functions-for-colors-to-help

RGBMAX = 0xFF  # Corresponds to 255
HLSMAX = 240  # MS excel's tint function expects that HLS is base 240. see:
# https://social.msdn.microsoft.com/Forums/en-US/e9d8c136-6d62-4098-9b1b-dac786149f43/excel-color-tint-algorithm-incorrect?forum=os_binaryfile#d3c2ac95-52e0-476b-86f1-e2a697f24969


def argb_to_ms_hls(argb: str) -> tuple[int, int, int]:
    """
    Converts a hex string of the form [aa]rrggbb to HLSMAX based HLS values. (alpha values are ignored)

    :param argb: A hex string of the form [aa]rrggbb
    :return: A tuple containing the hue, lightness, and saturation of the color in the range (0, HLSMAX)
    :raises ValueError: If the color is not a valid aRGB hex value
    :raises TypeError: If the argb arg is not an str
    """
    if isinstance(argb, str):
        m = aRGB_REGEX.match(argb)
        if m is None:
            raise ValueError("Colors must be aRGB hex values")
        if len(argb) > 6:
            color_str = argb[-6:]  # Ignore alpha values
        else:
            color_str = argb
        blue = int(color_str[4:], 16) / RGBMAX
        green = int(color_str[2:4], 16) / RGBMAX
        red = int(color_str[0:2], 16) / RGBMAX
        return rgb_to_ms_hls(red, green, blue)

    else:
        raise TypeError("argb arg shoud be an str")


def argb_to_css(argb: str) -> str:
    """
    Converts a hex string of the form [aa]rrggbb to CSS color string

    :param argb: A hex string of the form [aa]rrggbb
    :return: A CSS color string representation of the given color
    :raises ValueError: If the color is not a valid aRGB hex value
    :raises TypeError: If the argb arg is not an str
    """
    if isinstance(argb, str):
        m = aRGB_REGEX.match(argb)
        if m is None:
            raise ValueError("Colors must be aRGB hex values")
        if len(argb) == 6:
            return f"#{argb}"
        if argb.startswith("00"):
            return f"#{argb[-6:]}"
        blue = int(argb[6:], 16)
        green = int(argb[4:6], 16)
        red = int(argb[2:4], 16)
        alpha = int(argb[0:2], 16)
        return f"rgba({red}, {green}, {blue}, {alpha})"

    else:
        raise TypeError("argb arg shoud be an str")


def rgb_to_ms_hls(red: float, green: float, blue: float) -> tuple[int, int, int]:
    """
    Converts RGB values in the range (0,1) to HLSMAX based HLS values.

    :param red: The red component of the color in the range (0,1)
    :param green: The green component of the color in the range (0,1)
    :param blue: The blue component of the color in the range (0,1)
    :return: A tuple containing the hue, lightness, and saturation of the color in the range (0, HLSMAX)
    """
    h, l, s = rgb_to_hls(red, green, blue)
    return (int(round(h * HLSMAX)), int(round(l * HLSMAX)), int(round(s * HLSMAX)))


# def ms_hls_to_rgb(hue, lightness=None, saturation=None):
def ms_hls_to_rgb(
    hue: int, lightness: int, saturation: int
) -> tuple[float, float, float]:
    """
    Converts HLSMAX based HLS values to RGB values in the range (0,1)

    :param hue: The hue component of the color in the range (0, HLSMAX)
    :param lightness: The lightness component of the color in the range (0, HLSMAX)
    :param saturation: The saturation component of the color in the range (0, HLSMAX)
    :return: A tuple containing the red, green, and blue components of the color in the range (0,1)
    """
    # if lightness is None:
    #     hue, lightness, saturation = hue

    return hls_to_rgb(hue / HLSMAX, lightness / HLSMAX, saturation / HLSMAX)


# def rgb_to_hex(red, green=None, blue=None):
#     """Converts (0,1) based RGB values to a hex string 'rrggbb'"""
#     if green is None:
#         red, green, blue = red
def rgb_to_hex(red: float, green: float, blue: float) -> str:
    """
    Converts (0,1) based RGB values to a hex string 'rrggbb'

    :param red: The red component of the color in the range (0,1)
    :param green: The green component of the color in the range (0,1)
    :param blue: The blue component of the color in the range (0,1)
    :return: A hex string representation of the given color
    :raises ValueError: If the color is not a valid RGB hex value
    :raises TypeError: If the argb arg is not an str
    """
    return (
        "%02x%02x%02x"
        % (
            int(round(red * RGBMAX)),
            int(round(green * RGBMAX)),
            int(round(blue * RGBMAX)),
        )
    ).upper()


def tint_luminance(tint: float | None, lum: float) -> int:

    """
    Tints the given luminance (HLSMAX based) value by the given tint value.
    
    The tint value is a float value in the range (-1.0, 1.0).
    If tint is None, the luminance value will be returned as is.
    
    :param tint: The tint value to apply to the luminance.
    :param lum: The luminance value to tint.
    :return: The tinted luminance value as an integer in the range (0, HLSMAX)

    :seealso: http://ciintelligence.blogspot.co.uk/2012/02/converting-excel-theme-color-and-tint.html
    """

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
