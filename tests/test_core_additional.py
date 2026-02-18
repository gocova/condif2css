from openpyxl.styles import Color
from openpyxl.styles.colors import COLOR_INDEX

from condif2css import create_themed_css_color_resolver


def test_create_themed_css_color_resolver_falls_back_for_short_theme_list():
    normalize = create_themed_css_color_resolver([])

    assert normalize(Color(theme=0)) == "00FFFFFF"


def test_create_themed_css_color_resolver_handles_system_indexed_colors():
    normalize = create_themed_css_color_resolver(["112233", "AABBCC"])

    assert normalize(Color(indexed=64)) == "AABBCC"
    assert normalize(Color(indexed=65)) == "112233"


def test_create_themed_css_color_resolver_handles_none_and_zero_index():
    normalize = create_themed_css_color_resolver(["112233", "AABBCC"])

    assert normalize(None) is None
    assert normalize(Color(indexed=0)) == COLOR_INDEX[0]
