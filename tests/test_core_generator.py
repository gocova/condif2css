import sys
sys.path.append('src')

from condif2css import create_themed_get_css_color
from openpyxl.styles import Color
from condif2css.color import aRGB_to_CSS

testing_theme = ['FFFFFF',
 '000000',
 'E8E8E8',
 '0E2841',
 '156082',
 'E97132',
 '196B24',
 '0F9ED5',
 'A02B93',
 '4EA72E',
 '467886',
 '96607D']


def test_normalize_rgb_color():
    normalize = create_themed_get_css_color(testing_theme)
    testing_color = Color(rgb="00AABBCC")
    result = normalize(testing_color)
    assert result == "00AABBCC"
    assert aRGB_to_CSS(result) == '#AABBCC'

def test_normalize_argb_color():
    normalize = create_themed_get_css_color(testing_theme)
    testing_color = Color(rgb="AAAABBCC")
    result = normalize(testing_color)
    assert result == "AAAABBCC"
    assert aRGB_to_CSS(result) == "rgba(170, 187, 204, 170)"

def test_normalize_theme_no_tint():
    normalize = create_themed_get_css_color(testing_theme)
    testing_color = Color(theme=5)
    assert testing_color.type == 'theme'
    assert testing_color.value == 5
    assert testing_color.tint == 0.0
    result = normalize(testing_color)
    assert result == '00E97132'
    assert aRGB_to_CSS(result)  == "#E97132"

def test_normalize_theme_tint():
    normalize = create_themed_get_css_color(testing_theme)
    testing_color = Color(theme=4, tint=0.5)
    assert testing_color.type == 'theme'
    assert testing_color.value == 4
    assert testing_color.tint == 0.5
    result = normalize(testing_color)
    assert result == '0065BFE6'
    assert aRGB_to_CSS(result)  == "#65BFE6"

def test_normalize_theme_invalid_value():
    normalize = create_themed_get_css_color(testing_theme)
    testing_color = Color(theme=32)
    assert normalize(testing_color) == "00000000"

def test_normalize_indexed_color():
    normalize = create_themed_get_css_color(testing_theme)
    testing_color = Color(indexed=24)
    assert normalize(testing_color) == '009999FF'
