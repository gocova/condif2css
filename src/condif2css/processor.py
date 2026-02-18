# Copyright (c) 2026 Jose Gonzalo Covarrubias M <gocova.dev@gmail.com>
#
# Part of: batch_xlsx2html (bxx2html)
#

import logging
from collections.abc import Callable, Iterable
from typing import Any, Dict, Tuple

from mvin import TokenBool, TokenEmpty, TokenNumber, TokenString
from mvin.interpreter import get_interpreter
from openpyxl.cell import Cell, MergedCell
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.worksheet.worksheet import Worksheet

StyleDetails = Tuple[str, str, int, int, bool]
CompiledFormula = Tuple[str, Callable[[dict], Any], object]


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
    if isinstance(value, bool):
        return TokenBool(value)
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
    if formula_inputs is None:
        return ref_values, True

    if isinstance(formula_inputs, str):
        refs = [formula_inputs]
    elif isinstance(formula_inputs, Iterable):
        refs = formula_inputs
    else:
        return ref_values, True

    for ref in refs:
        if not isinstance(ref, str):
            logging.error(
                f"process: Unsupported formula input type '{type(ref)}'"
            )
            return {}, False
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


def _cell_code(sheet: Worksheet, cell: Cell | MergedCell) -> str:
    return f"{sheet.title}\\!{cell.coordinate}"


def _compile_formula(
    formula: str,
    fail_ok: bool,
) -> CompiledFormula | None:
    curr_formula_str = formula if formula.startswith("=") else f"={formula}"
    try:
        curr_tokenizer = Tokenizer(curr_formula_str)
    except Exception as exc:
        logging.error(
            f"process: Exception while parsing formula '{curr_formula_str}' -> {str(exc)}"
        )
        if not fail_ok:
            raise exc
        return None

    if not curr_tokenizer or not curr_tokenizer.items:
        logging.warning(
            f"process: Unable to parse formula: '{curr_formula_str}'"
        )
        return None

    try:
        curr_formula = get_interpreter(
            [
                item
                if item.subtype != "TEXT"
                else TokenString(item.value.strip('"'))
                for item in curr_tokenizer.items
            ]
        )
    except Exception as exc:
        logging.error(
            f"process: Exception while compiling formula '{curr_formula_str}' -> {str(exc)}"
        )
        if not fail_ok:
            raise exc
        return None

    if not curr_formula:
        logging.warning(
            f"process: Unable to get callable formula from: '{curr_formula_str}'"
        )
        return None

    return curr_formula_str, curr_formula, getattr(curr_formula, "inputs", None)


def _evaluate_cell_is_rule(
    operator: str | None,
    cell_value,
    operands: list,
) -> bool | None:
    try:
        if operator in (None, "equal"):
            return len(operands) == 1 and cell_value == operands[0]
        if operator == "notEqual":
            return len(operands) == 1 and cell_value != operands[0]
        if operator == "greaterThan":
            return len(operands) == 1 and cell_value > operands[0]
        if operator == "greaterThanOrEqual":
            return len(operands) == 1 and cell_value >= operands[0]
        if operator == "lessThan":
            return len(operands) == 1 and cell_value < operands[0]
        if operator == "lessThanOrEqual":
            return len(operands) == 1 and cell_value <= operands[0]
        if operator == "between":
            return len(operands) == 2 and operands[0] <= cell_value <= operands[1]
        if operator == "notBetween":
            return len(operands) == 2 and not (operands[0] <= cell_value <= operands[1])
    except Exception:
        return None
    return None


def _evaluate_text_rule(rule_type: str, text: str, cell_value) -> bool:
    left = "" if cell_value is None else str(cell_value)
    right = text
    left = left.lower()
    right = right.lower()

    if rule_type == "containsText":
        return right in left
    if rule_type == "notContainsText":
        return right not in left
    if rule_type == "beginsWith":
        return left.startswith(right)
    if rule_type == "endsWith":
        return left.endswith(right)
    return False


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

    flattened_rules = []
    for cf_order, cf in enumerate(sheet.conditional_formatting):
        cf_range = str(cf.cells)
        cf_ranges_list = cf_range.split(" ")
        logging.debug(f"process: cf -> range: {cf_range}")
        for rule_order, rule in enumerate(cf.rules):
            cf_priority = getattr(rule, "priority", None)
            normalized_priority = (
                cf_priority if isinstance(cf_priority, int) else 999_999
            )
            flattened_rules.append(
                (
                    normalized_priority,
                    cf_order,
                    rule_order,
                    cf_ranges_list,
                    rule,
                )
            )

    flattened_rules.sort(key=lambda item: (item[0], item[1], item[2]))
    stopped_cells: set[str] = set()

    for cf_priority, _, _, cf_ranges_list, rule in flattened_rules:
        dxf_id = rule.dxfId
        formulas = list(rule.formula or [])
        cf_stop_if_true = rule.stopIfTrue
        rule_type = getattr(rule, "type", "expression")

        main_formula = None
        cellis_operands = []
        text_rule_text: str | None = None

        if rule_type == "expression":
            if len(formulas) != 1:
                logging.warning(
                    f"process: Only 1 formula per rule is currently supported! Skipping rule: {rule}"
                )
                continue

            main_formula = _compile_formula(formulas[0], fail_ok=fail_ok)
            if main_formula is None:
                continue
        elif rule_type == "cellIs":
            operator = getattr(rule, "operator", None)
            expected_formulas = 2 if operator in {"between", "notBetween"} else 1
            if len(formulas) != expected_formulas:
                logging.warning(
                    f"process: Invalid 'cellIs' formula count ({len(formulas)}) for operator '{operator}'. Skipping rule: {rule}"
                )
                continue

            cellis_operands = []
            invalid_formula = False
            for formula in formulas:
                compiled = _compile_formula(formula, fail_ok=fail_ok)
                if compiled is None:
                    invalid_formula = True
                    break
                cellis_operands.append(compiled)
            if invalid_formula:
                continue
        elif rule_type in {"containsText", "notContainsText", "beginsWith", "endsWith"}:
            maybe_text = getattr(rule, "text", None)
            if isinstance(maybe_text, str):
                text_rule_text = maybe_text
            elif len(formulas) > 0 and isinstance(formulas[0], str):
                text_rule_text = formulas[0].strip('"')
            if not isinstance(text_rule_text, str):
                logging.warning(
                    f"process: Rule type '{rule_type}' does not provide text payload. Skipping rule: {rule}"
                )
                continue
        else:
            logging.warning(
                f"process: Unsupported rule type '{rule_type}'. Skipping rule: {rule}"
            )
            continue

        anchor_cell = _extract_anchor_cell(sheet, cf_ranges_list[0])
        if anchor_cell is None:
            logging.warning(
                f"process: Unable to get anchor cell from range '{cf_ranges_list[0]}' to apply conditional formatting formula!"
            )
            continue

        if main_formula is not None:
            curr_formula_str, _, curr_formula_inputs = main_formula
            logging.debug(f"process: cf formula[p: {cf_priority}] -> {curr_formula_str}")
            logging.debug(f"process: Using formula inputs: {curr_formula_inputs}")

        for specific_range in cf_ranges_list:
            possible_range = sheet[specific_range]
            for cell in _iter_cells(possible_range):
                code = _cell_code(sheet, cell)
                if code in stopped_cells:
                    continue

                delta_col = (
                    (cell.column if cell.column else 0) - anchor_cell.column
                )
                delta_row = (cell.row if cell.row else 0) - anchor_cell.row

                formula_result = None
                if rule_type == "expression":
                    curr_formula_str, curr_formula, curr_formula_inputs = main_formula  # type: ignore[misc]
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
                    if not isinstance(formula_result, bool):
                        logging.warning(
                            f"process: Expected bool for result, but '{formula_result}' was found!"
                        )
                        continue
                elif rule_type == "cellIs":
                    operator = getattr(rule, "operator", None)
                    operand_values = []
                    is_valid = True
                    for operand_formula_str, operand_formula, operand_inputs in cellis_operands:
                        operand_ref_values, can_eval_operand = _build_ref_values(
                            sheet,
                            operand_inputs,
                            delta_row,
                            delta_col,
                        )
                        if not can_eval_operand:
                            is_valid = False
                            break
                        try:
                            operand_values.append(operand_formula(operand_ref_values))
                        except Exception as exc:
                            logging.error(
                                f"process: Exception found during formula '{operand_formula_str}' evaluation for reference '{cell.coordinate}' -> {str(exc)}"
                            )
                            if not fail_ok:
                                raise exc
                            is_valid = False
                            break

                    if not is_valid:
                        continue

                    formula_result = _evaluate_cell_is_rule(
                        operator,
                        getattr(cell, "value", None),
                        operand_values,
                    )
                    if formula_result is None:
                        logging.warning(
                            f"process: Unable to evaluate 'cellIs' operator '{operator}' for cell '{cell.coordinate}'."
                        )
                        continue
                else:
                    formula_result = _evaluate_text_rule(
                        rule_type,
                        text_rule_text if text_rule_text is not None else "",
                        getattr(cell, "value", None),
                    )

                if not formula_result:
                    continue

                if isinstance(dxf_id, int) and dxf_id >= 0:
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

                if cf_stop_if_true:
                    stopped_cells.add(code)

    return results
