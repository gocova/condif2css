import sys

sys.path.append("src")

from openpyxl.styles import Color

from condif2css import create_themed_css_color_resolver


def test_create_themed_css_color_resolver_falls_back_for_short_theme_list():
    normalize = create_themed_css_color_resolver([])

    assert normalize(Color(theme=0)) == "00FFFFFF"


def test_create_themed_css_color_resolver_handles_system_indexed_colors():
    normalize = create_themed_css_color_resolver(["112233", "AABBCC"])

    assert normalize(Color(indexed=64)) == "AABBCC"
    assert normalize(Color(indexed=65)) == "112233"

