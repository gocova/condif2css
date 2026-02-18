from condif2css.color import argb_to_css, ms_hls_to_rgb, rgb_to_hex, rgb_to_ms_hls, tint_luminance, argb_to_ms_hls
import pytest
from mvin import TokenBool
import condif2css.processor as processor

def test_to_css_rgb():
    assert argb_to_css('AABBDD') == '#AABBDD'

def test_to_css_rgba():
    assert argb_to_css('AAAABBDD') == 'rgba(170, 187, 221, 0.667)'

def test_to_css_rgba_zero_alpha():
    assert argb_to_css('00AABBDD') == '#AABBDD'

def test_to_css_no_aRGB():
    with pytest.raises(ValueError):
        argb_to_css('no aRGB color')

def test_to_css_no_str_arg():
    with pytest.raises(TypeError):
        argb_to_css(None) # type: ignore

def test_to_css_no_args():
    with pytest.raises(TypeError):
        argb_to_css() # type: ignore

def test_to_token_bool_is_boolean_token():
    assert isinstance(processor._to_token(True), TokenBool)

def test_rgb_to_hex():
    assert rgb_to_hex(1.0,0.5,0.0) == 'FF8000'

def test_tint_luminance_zero_zero():
    assert tint_luminance(0, 0) == 0

def test_tint_luminance_half_zero():
    assert tint_luminance(0.5, 0) == 120

def test_tint_luminance_minus_one():
    assert tint_luminance(-0.5, 1) == 0

def test_rgb_to_ms_hls():
    assert rgb_to_ms_hls(0.054901960784313725, 0.1568627450980392, 0.2549019607843137) == (140,37,155)

def test_rgb_to_ms_hls_2():
    assert rgb_to_ms_hls(0.08235294117647059, 0.3764705882352941, 0.5098039215686274) == (132,71,173)

def test_rgb_to_ms_hls_3():
    assert rgb_to_ms_hls(0.9137254901960784, 0.44313725490196076, 0.19607843137254902) == (14,133,193)

def test_0():
    # print(ms_hls_to_rgb(140, 189, 155))
    assert ms_hls_to_rgb(140, 189, 155) == (0.6502604166666667, 0.7874999999999999, 0.9247395833333333)

def test_rgb_to_hex_multi():
    assert rgb_to_hex(0.6502604166666667, 0.7874999999999999, 0.9247395833333333) == 'A6C9EC'
    assert rgb_to_hex(0.5124305555555556, 0.7983611111111109, 0.9209027777777777) == '83CCEB'
    assert rgb_to_hex(0.9649131944444445, 0.7776093749999999, 0.6767534722222222) == 'F6C6AD'

def test_mult():
    assert tint_luminance(0.749992370372631, 37) == 189
    assert tint_luminance(0.5999938962981048, 71) == 172
    assert tint_luminance(0.5999938962981048, 133) == 197
    assert tint_luminance(0.0999786370433668, 37) == 57
    assert tint_luminance(0.249977111117893, 37) == 88
    assert tint_luminance(-0.249977111117893, 133) == 100

def test_complex_0():
    rgb_base = '0E2841'
    assert argb_to_ms_hls(rgb_base) == (140, 37, 155)
    h_part, l_part, s_part = argb_to_ms_hls(rgb_base)
    tint:float = 0.749992370372631
    pre = ms_hls_to_rgb(h_part, tint_luminance(tint, l_part), s_part)
    assert pre == (0.6502604166666667, 0.7874999999999999, 0.9247395833333333)
    assert f'00{rgb_to_hex(*pre)}'== '00A6C9EC'

def test_complex_1():
    rgb_base = 'E97132'
    assert argb_to_ms_hls(rgb_base) == (14, 133, 193)
    h_part, l_part, s_part = argb_to_ms_hls(rgb_base)
    tint = 0.5999938962981048
    pre = ms_hls_to_rgb(h_part, tint_luminance(tint, l_part), s_part)
    assert pre == (0.9649131944444445, 0.7776093749999999, 0.6767534722222222)
    assert f'00{rgb_to_hex(*pre)}'== '00F6C6AD'
