# Copyright (c) 2026 Jose Gonzalo Covarrubias M <gocova.dev@gmail.com>
#
# Part of: batch_xlsx2html (bxx2html)
#

import logging
from collections.abc import Iterable
from typing import Dict, Tuple

from mvin import TokenEmpty, TokenNumber, TokenString
from mvin.interpreter import get_interpreter
from openpyxl.cell import Cell, MergedCell
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.worksheet.worksheet import Worksheet

StyleDetails = Tuple[str, str, int, int, bool]


def _get_offsets_for(cell_coord: str, row_offset: int, column_offset: int) -> Tuple[int, int]:
    """
    Returns the row and column offsets for a cell coordinate based on the presence of `$`.
    """
    if isinstance(cell_coord, str) and len(cell_coord) >= 2:
        if cell_coord.startswith("$"):
            offset_col = 0
            inner_coord = cell_coord[1:]
        else:
            offset_col = column_offset
            inner_coord = cell_coord

        offset_row = row_offset if len(inner_coord.split("$")) == 1 else 0
        return offset_row, offset_col

    return 0, 0


def _iter_cells(range_or_cell):
    if isinstance(range_or_cell, (Cell, MergedCell)):
        yield range_or_cell
        return

    if isinstance(range_or_cell, Iterable):
        for row in range_or_cell:
            if isinstance(row, (Cell, MergedCell)):
                yield row
            elif isinstance(row, Iterable):
                for cell in row:
                    if isinstance(cell, (Cell, MergedCell)):
                        yield cell


def _extract_anchor_cell(sheet: Worksheet, first_range: str) -> Cell | None:
    first_range_obj = sheet[first_range]
    if isinstance(first_range_obj, Cell):
        return first_range_obj

    if isinstance(first_range_obj, Tuple) and len(first_range_obj) > 0:
        first_row = first_range_obj[0]
        if isinstance(first_row, Tuple) and len(first_row) > 0:
            first_cell = first_row[0]
            if isinstance(first_cell, Cell):
                return first_cell

    return None


def _to_token(value):
    if isinstance(value, str):
        return TokenString(value)
    if isinstance(value, (int, float)):
        return TokenNumber(value)
    if value is None:
        return TokenEmpty()
    return None


def _build_ref_values(
    sheet: Worksheet,
    formula_inputs,
    delta_row: int,
    delta_col: int,
) -> Tuple[dict, bool]:
    ref_values = {}
    if not isinstance(formula_inputs, set):
        return ref_values, True

    for ref in formula_inputs:
        ref_cell = sheet[ref]
        if not isinstance(ref_cell, Cell):
            logging.error(
                f"process: Unsupported reference '{ref}' for formula argument"
            )
            return {}, False

        try:
            offset_row, offset_col = _get_offsets_for(ref, delta_row, delta_col)
            offset_cell = ref_cell.offset(row=offset_row, column=offset_col)
        except Exception as exc:
            logging.error(
                f"process: Exception while getting offset for reference '{ref}' -> {str(exc)}"
            )
            return {}, False

        if not isinstance(offset_cell, (Cell, MergedCell)):
            logging.error(
                f"process: Unable to apply '{ref}'.offset(row={offset_row}, column={offset_col})"
            )
            return {}, False

        curr_ref_value = getattr(offset_cell, "value", None)
        curr_token = _to_token(curr_ref_value)
        if curr_token is not None:
            ref_values[ref] = curr_token

    return ref_values, True


def _save_result(
    results: Dict[str, StyleDetails],
    sheet: Worksheet,
    cell: Cell | MergedCell,
    cf_priority: int,
    dxf_id: int,
    cf_stop_if_true: bool | None,
):
    code = f"{sheet.title}\\!{cell.coordinate}"
    proposed_style: StyleDetails = (
        sheet.title,
        cell.coordinate,
        cf_priority,
        dxf_id,
        cf_stop_if_true if cf_stop_if_true is not None else False,
    )

    if code in results:
        _, _, old_priority, _, _ = results[code]
        if old_priority <= cf_priority:
            return

    results[code] = proposed_style


def process_conditional_formatting(
    sheet: Worksheet,  # required for styles? and reference
    fail_ok: bool = True,
) -> Dict[str, StyleDetails]:
    """
    Process a worksheet conditional formatting rules into a dictionary mapping cell references to their corresponding respective conditional formatting styles.

    The resulting dictionary will contain the following information for each cell:
        - The worksheet title
        - The cell coordinate
        - The priority of the conditional formatting rule
        - The differential style index
        - A boolean indicating whether the style should be applied if the formula evaluates to True

    :param sheet: The worksheet to process
    :param fail_ok: If True, then exceptions will be caught and logged. Otherwise, exceptions will be raised.

    :return: A dictionary mapping cell references to their respective conditional formatting styles
    """
    results: Dict[str, StyleDetails] = {}
    if sheet.conditional_formatting is None:
        logging.debug("process: worksheet has no conditional formatting")
        return results

    for cf in sheet.conditional_formatting:
        cf_range = str(cf.cells)
        cf_ranges_list = cf_range.split(" ")
        logging.debug(f"process: cf -> range: {cf_range}")
        for rule in cf.rules:
            dxf_id = rule.dxfId
            if dxf_id is None:
                continue

            cf_priority = rule.priority
            formulas = rule.formula
            cf_stop_if_true = rule.stopIfTrue
            if len(formulas) != 1:
                logging.warning(
                    f"process: Only 1 formula per rule is currently supporter! Skipping rule: {rule}"
                )
                continue

            curr_formula_str = formulas[0]
            curr_formula_str = (
                curr_formula_str
                if curr_formula_str.startswith("=")
                else f"={curr_formula_str}"
            )
            logging.debug(
                f"process: cf formula[p: {cf_priority}] -> {curr_formula_str}"
            )

            curr_tokenizer = Tokenizer(curr_formula_str)
            if not curr_tokenizer or not curr_tokenizer.items:
                logging.warning(
                    f"process: Unable to parse formula: '{curr_formula_str}'"
                )
                continue

            curr_formula = get_interpreter(
                [
                    item
                    if item.subtype != "TEXT"
                    else TokenString(item.value.strip('"'))
                    for item in curr_tokenizer.items
                ]
            )
            if not curr_formula:
                logging.warning(
                    f"process: Unable to get callable formula from: '{curr_formula_str}'"
                )
                continue

            anchor_cell = _extract_anchor_cell(sheet, cf_ranges_list[0])
            if anchor_cell is None:
                logging.warning(
                    f"process: Unable to get anchor cell from range '{cf_ranges_list[0]}' to apply conditional formatting formula!"
                )
                continue

            curr_formula_inputs = getattr(curr_formula, "inputs", None)
            logging.debug(f"process: Using formula inputs: {curr_formula_inputs}")

            for specific_range in cf_ranges_list:
                possible_range = sheet[specific_range]
                for cell in _iter_cells(possible_range):
                    delta_col = ( cell.column if cell.column else 0 ) - anchor_cell.column
                    delta_row = ( cell.row if cell.row else 0 ) - anchor_cell.row

                    ref_values, should_apply_func = _build_ref_values(
                        sheet,
                        curr_formula_inputs,
                        delta_row,
                        delta_col,
                    )
                    if not should_apply_func:
                        continue

                    try:
                        formula_result = curr_formula(ref_values)
                    except Exception as exc:
                        logging.error(
                            f"process: Exception found during formula '{curr_formula_str}' evaluation for reference '{cell.coordinate}' -> {str(exc)}"
                        )
                        if not fail_ok:
                            raise exc
                        continue

                    logging.debug(f"process: Formula result -> {formula_result}")
                    if not isinstance(formula_result, bool):
                        logging.warning(
                            f"process: Expected bool for result, but '{formula_result}' was found!"
                        )
                        continue

                    if formula_result:
                        logging.debug(
                            f"process: Applying differential style with index: {dxf_id} for cell['{cell.coordinate}']"
                        )
                        _save_result(
                            results,
                            sheet,
                            cell,
                            cf_priority,
                            dxf_id,
                            cf_stop_if_true,
                        )

    return results
