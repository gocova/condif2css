# Copyright (c) 2026 Jose Gonzalo Covarrubias M <gocova.dev@gmail.com>
#
# Part of: batch_xlsx2html (bxx2html)
#

from openpyxl.styles.differential import DifferentialStyle
from openpyxl.workbook import Workbook


def get_differential_style(workbook: Workbook, dxf_id: int) -> DifferentialStyle | None:
    """
    Safely returns a differential style from a workbook by `dxf_id`.

    Returns `None` when the workbook has no differential styles or when `dxf_id`
    is invalid/out of range.
    """
    if not isinstance(dxf_id, int) or dxf_id < 0:
        return None

    differential_styles = getattr(workbook, "_differential_styles", None)
    if differential_styles is None:
        return None

    styles = getattr(differential_styles, "styles", None)
    if not isinstance(styles, list):
        return None

    if dxf_id >= len(styles):
        return None

    style = styles[dxf_id]
    return style if isinstance(style, DifferentialStyle) else None
