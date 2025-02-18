from collections.abc import Iterable
from mvin import TokenString
from mvin.functions.excel_lib import TokenNumber
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.formula.tokenizer import Tokenizer
from mvin.interpreter import get_interpreter
from typing import Dict, Tuple
import logging
import re

CONTAINS_TEXT_REGEXP = re.compile(
    r'NOT\(ISERROR\(SEARCH\("(?P<text>.+)"\s*,\s*(?P<reference>\$?[A-Z]+\$?[1-9]+\d*)\)\)\)'  # r-string
)

NOT_CONTAINS_TEXT_REGEXP = re.compile(
    r'ISERROR\(SEARCH\("(?P<text>.+)"\s*,\s*(?P<reference>\$?[A-Z]+\$?[1-9]+\d*)\)\)'  # r-string
)

TEXT_REGEXP_ENDSWITH = re.compile(
    r'RIGHT\((?P<reference>\$?[A-Z]+\s?[1-9]+\d?)\s*,\s*LEN\("(?P<text>.*)"\)\)="(?P<text_1>.*)"'  # r-string
)


def _text_func_endswith(cell_value: str, formula_text: str) -> bool:
    return cell_value.endswith(formula_text)


TEXT_REGEXP_BEGINSWITH = re.compile(
    r'LEFT\((?P<reference>\$?[A-Z]+\s?[1-9]+\d?)\s*,\s*LEN\("(?P<text>.*)"\)\)="(?P<text_1>.*)"'  # r-string
)


def _text_func_beginswith(cell_value: str, formula_text: str) -> bool:
    return cell_value.startswith(formula_text)


PRECOMPILED_TEXT_EXPRS = {
    "containsText": (
        CONTAINS_TEXT_REGEXP,
        lambda cell_value, formula_text: formula_text in cell_value,
    ),
    "notContainsText": (
        NOT_CONTAINS_TEXT_REGEXP,
        lambda cell_value, formula_text: formula_text not in cell_value,
    ),
    "endsWith": (TEXT_REGEXP_ENDSWITH, _text_func_endswith),
    "beginsWith": (TEXT_REGEXP_BEGINSWITH, _text_func_beginswith),
}


def process(
    sheet: Worksheet,  # required for styles? and reference
    # differential_styles,
    # ) -> List[Tuple[int, int, str, int, int, bool]]:
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
    results: Dict[str, Tuple[str, str, int, int, bool]] = {}
    if sheet.conditional_formatting is not None:
        group_id = 0
        row_id = 0
        for cf in sheet.conditional_formatting:
            cf_range = str(cf.cells)
            for rule in cf.rules:
                dxf_id = rule.dxfId
                if dxf_id is not None:
                    cf_type = rule.type
                    cf_priority = rule.priority
                    cf_text = rule.text
                    formulas = rule.formula
                    cf_stop_if_true = rule.stopIfTrue
                    # print(formulas)

                    possible_text_expr = PRECOMPILED_TEXT_EXPRS.get(cf_type)
                    if possible_text_expr is not None:
                        precompiled_regexp, check_func = possible_text_expr
                        if len(formulas) == 1:
                            matched = precompiled_regexp.search(formulas[0])
                            if matched:
                                if matched.group("text") == cf_text:
                                    cell_ref = matched.group("reference")
                                    if isinstance(cell_ref, str):
                                        cell = sheet[cell_ref]
                                        cell_value = (
                                            cell.value if cell is not None else ""
                                        )
                                        if (
                                            cell_value is not None
                                            and isinstance(cell_value, str)
                                            and check_func(cell_value, cf_text)
                                        ):
                                            # results.append(
                                            #     (
                                            #         row_id,
                                            #         group_id,
                                            #         cf_range,
                                            #         cf_priority,
                                            #         dxf_id,
                                            #         cf_stop_if_true
                                            #         if cf_stop_if_true is not None
                                            #         else False,
                                            #     )
                                            # )
                                            code = f"{sheet.title}\\!{cell.coordinate}"
                                            proposed_style = (
                                                sheet.title,
                                                cell.coordinate,
                                                cf_priority,
                                                dxf_id,
                                                cf_stop_if_true
                                                if cf_stop_if_true is not None
                                                else False,
                                            )
                                            if code in results:
                                                _, _, old_priority, _, _ = results[code]
                                                if old_priority >= cf_priority:
                                                    continue
                                            results[code] = proposed_style
                                    else:
                                        logging.error(
                                            f"process: Invalid cell reference found in formula '{formulas[0]}'"
                                        )
                                else:
                                    logging.warning(
                                        f"process: Inconsistent state for rule '{cf_type}' -> Expected text '{cf_text}' but '{matched.group('text')}' was found!"
                                    )
                        else:
                            logging.warning(
                                f"process: Only 1 formula per rule is currently supported! Skipping rule: {rule}"
                            )
                    else:
                        if len(formulas) == 1:
                            curr_formula_str = formulas[0]
                            curr_formula_str = (
                                curr_formula_str
                                if curr_formula_str.startswith("=")
                                else f"={curr_formula_str}"
                            )
                            curr_tokenizer = Tokenizer(curr_formula_str)
                            if curr_tokenizer and curr_tokenizer.items:
                                curr_formula = get_interpreter(curr_tokenizer.items)
                                if curr_formula:
                                    ref_values = {}
                                    curr_formula_inputs = getattr(
                                        curr_formula, "inputs", None
                                    )
                                    if isinstance(curr_formula_inputs, set):
                                        # curr_formula_inputs = cast(
                                        #     Set[str], curr_formula.inputs
                                        # )
                                        for ref in curr_formula_inputs:
                                            curr_ref_value = getattr(
                                                sheet[ref], "value", None
                                            )
                                            if isinstance(curr_ref_value, str):
                                                ref_values[ref] = TokenString(
                                                    curr_ref_value
                                                )
                                            elif isinstance(
                                                curr_ref_value, int
                                            ) or isinstance(curr_ref_value, float):
                                                ref_values[ref] = TokenNumber(
                                                    curr_ref_value
                                                )
                                    formula_result = curr_formula(ref_values)
                                    if isinstance(formula_result, bool):
                                        if formula_result:
                                            for specific_range in cf_range.split(" "):
                                                print(
                                                    f" > Saving specific_range for each cell of '{specific_range}'"
                                                )
                                                print(
                                                    f" > CellRange: {sheet[specific_range]}"
                                                )
                                                possible_range = sheet[specific_range]
                                                for row in (
                                                    possible_range
                                                    if isinstance(
                                                        possible_range, Iterable
                                                    )
                                                    else ((possible_range,),)
                                                ):
                                                    for cell in row:
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
                                                            _, _, old_priority, _, _ = (
                                                                results[code]
                                                            )
                                                            if (
                                                                old_priority
                                                                >= cf_priority
                                                            ):
                                                                continue
                                                        results[code] = proposed_style
                                    else:
                                        logging.warning(
                                            f"process: Expected bool for result, but '{formula_result}' was found!"
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

                        # parse_result = parse(formulas, sheet)
                        # if parse_result is not None:
                        #     results.append(
                        #         (
                        #             row_id,
                        #             group_id,
                        #             cf_range,
                        #             dxf_id,
                        #             cf_stop_if_true
                        #             if cf_stop_if_true is not None
                        #             else False,
                        #         )
                        #     )
                        # else:
                        #     logging.error(f"process: Unsupported rule: {rule}")
                row_id += 1
            group_id += 1

    else:
        logging.debug("process: ")
    return results
