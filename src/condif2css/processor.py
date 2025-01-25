# from openpyxl.worksheet._read_only import ReadOnlyWorksheet
from openpyxl.worksheet.worksheet import Worksheet
from typing import List, Tuple
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
) -> List[Tuple[int, int, str, int, int, bool]]:
    results = []
    if sheet.conditional_formatting is not None:
        group_id = 0
        row_id = 0
        for cf in sheet.conditional_formatting:
            cf_range = str(cf.cells)
            for rule in cf.rules:
                dxfId = rule.dxfId
                if dxfId is not None:
                    cf_type = rule.type
                    priority = rule.priority
                    cf_text = rule.text
                    formulas = rule.formula
                    cf_stop_if_true = rule.stopIfTrue

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
                                            results.append(
                                                (
                                                    row_id,
                                                    group_id,
                                                    cf_range,
                                                    priority,
                                                    dxfId,
                                                    cf_stop_if_true
                                                    if cf_stop_if_true is not None
                                                    else False,
                                                )
                                            )
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
                        print(rule)
                row_id += 1
            group_id += 1

    else:
        logging.debug("process: ")
    return results
