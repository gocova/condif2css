import sys

sys.path.append("src")

import pytest
from openpyxl import Workbook
from openpyxl.writer.theme import theme_xml

from condif2css.themes import get_theme_colors


def test_get_theme_colors_from_openpyxl_theme_xml():
    wb = Workbook()
    wb.loaded_theme = theme_xml

    colors = get_theme_colors(wb)

    assert len(colors) == 12
    assert colors[0] == "FFFFFF"
    assert colors[1] == "000000"
    assert colors[-1] == "800080"


def test_get_theme_colors_raises_without_loaded_theme():
    wb = Workbook()
    with pytest.raises(ValueError):
        get_theme_colors(wb)

