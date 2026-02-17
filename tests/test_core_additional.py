import sys

sys.path.append("src")

from openpyxl.styles import Color

from condif2css import create_themed_get_css_color


def test_create_themed_get_css_color_falls_back_for_short_theme_list():
    normalize = create_themed_get_css_color([])

    assert normalize(Color(theme=0)) == "00FFFFFF"


def test_create_themed_get_css_color_handles_system_indexed_colors():
    normalize = create_themed_get_css_color(["112233", "AABBCC"])

    assert normalize(Color(indexed=64)) == "AABBCC"
    assert normalize(Color(indexed=65)) == "112233"

