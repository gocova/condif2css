from collections.abc import Iterable
from mvin import TokenEmpty, TokenString, TokenNumber
from openpyxl.cell import Cell, MergedCell
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.formula.tokenizer import Tokenizer
from mvin.interpreter import get_interpreter
from typing import Dict, Tuple
import logging


def process(
    sheet: Worksheet,  # required for styles? and reference
) -> Dict[str, Tuple[str, str, int, int, bool]]:
    """
    Returns:
        Dict[
            key
            , Tuple[
                sheetname
                cell_ref
                priority
                dxf_id
                stop_if_true
    """

    def get_offsets_for(
        cell_coord: str, row_offset: int, column_offset: int
    ) -> Tuple[int, int]:
        if isinstance(cell_coord, str) and len(cell_coord) >= 2:
            if cell_coord.startswith("$"):
                offset_col = 0
                inner_coord = cell_coord[1:]
            else:
                offset_col = column_offset
                inner_coord = cell_coord

            offset_row = row_offset if len(inner_coord.split("$")) == 1 else 0

            return offset_row, offset_col
        else:
            return 0, 0

    results: Dict[str, Tuple[str, str, int, int, bool]] = {}
    if sheet.conditional_formatting is not None:
        group_id = 0
        row_id = 0
        for cf in sheet.conditional_formatting:
            cf_range = str(cf.cells)
            logging.debug(f"process: cf -> range: {cf_range}")
            for rule in cf.rules:
                dxf_id = rule.dxfId
                if dxf_id is not None:
                    cf_priority = rule.priority
                    formulas = rule.formula
                    cf_stop_if_true = rule.stopIfTrue

                    if len(formulas) == 1:
                        curr_formula_str = formulas[0]
                        curr_formula_str = (
                            curr_formula_str
                            if curr_formula_str.startswith("=")
                            else f"={curr_formula_str}"
                        )
                        logging.debug(f"process: cf formula[p: {cf_priority}] -> {curr_formula_str}")
                        curr_tokenizer = Tokenizer(curr_formula_str)
                        if curr_tokenizer and curr_tokenizer.items:

                            curr_formula = get_interpreter(
                                [
                                    x
                                    if x.subtype != "TEXT"
                                    else TokenString(x.value.strip('"'))
                                    for x in curr_tokenizer.items
                                ]
                            )
                            if curr_formula:
                                cf_ranges_list = cf_range.split(" ")
                                first_range = sheet[cf_ranges_list[0]]
                                anchor_column = 1
                                anchor_row = 1
                                anchor_error = False
                                if isinstance(first_range, Cell):
                                    anchor_column = first_range.column
                                    anchor_row = first_range.row
                                elif isinstance(first_range, Tuple):
                                    if len(first_range) > 0:
                                        first_range_first_row = first_range[0]
                                        if (
                                            isinstance(first_range_first_row, Tuple)
                                            and len(first_range_first_row) > 0
                                        ):
                                            first_range_first_cell = (
                                                first_range_first_row[0]
                                            )
                                            if isinstance(first_range_first_cell, Cell):
                                                anchor_column = (
                                                    first_range_first_cell.column
                                                )
                                                anchor_row = first_range_first_cell.row
                                            else:
                                                anchor_error = True
                                        else:
                                            anchor_error = True
                                    else:
                                        anchor_error = True
                                else:
                                    anchor_error = True

                                if not anchor_error:
                                    curr_formula_inputs = getattr(
                                        curr_formula, "inputs", None
                                    )
                                    logging.debug(
                                        f"process: Using formula inputs: {curr_formula_inputs}"
                                    )
                                    for specific_range in cf_range.split(" "):
                                        possible_range = sheet[specific_range]
                                        for row in (
                                            possible_range
                                            if isinstance(possible_range, Iterable)
                                            else ((possible_range,),)
                                        ):
                                            for cell in row:
                                                delta_col = cell.column - anchor_column
                                                delta_row = cell.row - anchor_row

                                                ref_values = {}
                                                should_apply_func = True

                                                if isinstance(curr_formula_inputs, set):
                                                    for ref in curr_formula_inputs:
                                                        ref_cell = sheet[ref]

                                                        if isinstance(ref_cell, Cell):
                                                            offset_row, offset_col = (
                                                                get_offsets_for(
                                                                    ref,
                                                                    delta_row,
                                                                    delta_col,
                                                                )
                                                            )
                                                            offset_cell = (
                                                                ref_cell.offset(
                                                                    row=offset_row,
                                                                    column=offset_col,
                                                                )
                                                            )
                                                            if isinstance(
                                                                offset_cell,
                                                                Cell | MergedCell,
                                                            ):
                                                                curr_ref_value = (
                                                                    getattr(
                                                                        offset_cell,
                                                                        "value",
                                                                        None,
                                                                    )
                                                                )
                                                                if isinstance(
                                                                    curr_ref_value, str
                                                                ):
                                                                    ref_values[ref] = (
                                                                        TokenString(
                                                                            curr_ref_value
                                                                        )
                                                                    )
                                                                elif isinstance(
                                                                    curr_ref_value, int
                                                                ) or isinstance(
                                                                    curr_ref_value,
                                                                    float,
                                                                ):
                                                                    ref_values[ref] = (
                                                                        TokenNumber(
                                                                            curr_ref_value
                                                                        )
                                                                    )
                                                                elif (
                                                                    curr_ref_value
                                                                    is None
                                                                ):
                                                                    ref_values[ref] = (
                                                                        TokenEmpty()
                                                                    )
                                                            else:
                                                                logging.error(
                                                                    f"process: Unable to apply '{ref}'.offset(row={offset_row}, column={offset_col})"
                                                                )
                                                                should_apply_func = (
                                                                    False
                                                                )
                                                                break
                                                        else:
                                                            logging.error(
                                                                f"process: Unsupported reference '{ref}' for formula argument"
                                                            )
                                                            should_apply_func = False
                                                            break
                                                if should_apply_func:
                                                    formula_result = curr_formula(
                                                        ref_values
                                                    )

                                                    logging.debug(
                                                        f"process: Formula result -> {formula_result}"
                                                    )

                                                    if isinstance(formula_result, bool):
                                                        if formula_result:
                                                            logging.debug(
                                                                f"process: Applying differential style with index: {dxf_id} for cell['{cell.coordinate}']"
                                                            )
                                                            code = f"{sheet.title}\\!{cell.coordinate}"
                                                            proposed_style = (
                                                                sheet.title,
                                                                cell.coordinate,
                                                                cf_priority,
                                                                dxf_id,
                                                                cf_stop_if_true
                                                                if cf_stop_if_true
                                                                is not None
                                                                else False,
                                                            )
                                                            if code in results:
                                                                (
                                                                    _,
                                                                    _,
                                                                    old_priority,
                                                                    _,
                                                                    _,
                                                                ) = results[code]
                                                                if (
                                                                    old_priority
                                                                    <= cf_priority
                                                                ):
                                                                    continue
                                                            results[code] = (
                                                                proposed_style
                                                            )
                                                    else:
                                                        logging.warning(
                                                            f"process: Expected bool for result, but '{formula_result}' was found!"
                                                        )
                                else:
                                    logging.warning(
                                        f"process: Unable to get anchor cell from range '{cf_ranges_list[0]}' to apply conditional formatting formula!"
                                    )

                            else:
                                logging.warning(
                                    f"process: Unable to get callable formula from: '{curr_formula_str}'"
                                )
                        else:
                            logging.warning(
                                f"process: Unable to parse formula: '{curr_formula_str}'"
                            )
                    else:
                        logging.warning(
                            f"process: Only 1 formula per rule is currently supporter! Skipping rule: {rule}"
                        )

                row_id += 1
            group_id += 1

    else:
        logging.debug("process: ")
    return results
