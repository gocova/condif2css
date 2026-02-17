import sys

sys.path.append("src")

import pytest
from openpyxl import Workbook
from openpyxl.writer.theme import theme_xml

from condif2css.themes import ThemeColorsError, get_theme_colors


def test_get_theme_colors_from_openpyxl_theme_xml():
    wb = Workbook()
    wb.loaded_theme = theme_xml

    colors = get_theme_colors(wb)

    assert len(colors) == 12
    assert colors[0] == "FFFFFF"
    assert colors[1] == "000000"
    assert colors[-1] == "800080"


def test_get_theme_colors_returns_empty_without_loaded_theme_by_default():
    wb = Workbook()
    assert get_theme_colors(wb, strict=False) == []


def test_get_theme_colors_raises_custom_exception_when_strict():
    wb = Workbook()
    with pytest.raises(ThemeColorsError):
        get_theme_colors(wb, strict=True)


def test_get_theme_colors_invalid_theme_returns_empty_or_raises():
    wb = Workbook()
    wb.loaded_theme = "<not valid xml>"

    assert get_theme_colors(wb, strict=False) == []
    with pytest.raises(ThemeColorsError):
        get_theme_colors(wb, strict=True)
