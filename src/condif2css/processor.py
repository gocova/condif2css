from collections.abc import Iterable
from mvin import TokenEmpty, TokenString, TokenNumber
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
    results: Dict[str, Tuple[str, str, int, int, bool]] = {}
    if sheet.conditional_formatting is not None:
        group_id = 0
        row_id = 0
        for cf in sheet.conditional_formatting:
            cf_range = str(cf.cells)
            print(f"\ncf range:{cf_range}")
            for rule in cf.rules:
                print(rule)
                dxf_id = rule.dxfId
                if dxf_id is not None:
                    # cf_type = rule.type
                    cf_priority = rule.priority
                    # cf_text = rule.text
                    formulas = rule.formula
                    cf_stop_if_true = rule.stopIfTrue

                    if len(formulas) == 1:
                        curr_formula_str = formulas[0]
                        curr_formula_str = (
                            curr_formula_str
                            if curr_formula_str.startswith("=")
                            else f"={curr_formula_str}"
                        )
                        print(f"formula[p: {cf_priority}] -> {curr_formula_str}")
                        curr_tokenizer = Tokenizer(curr_formula_str)
                        if curr_tokenizer and curr_tokenizer.items:
                            logger = logging.getLogger()
                            # logger.setLevel(logging.DEBUG)
                            curr_formula = get_interpreter(
                                [
                                    x
                                    if x.subtype != "TEXT"
                                    else TokenString(x.value.strip('"'))
                                    for x in curr_tokenizer.items
                                ]
                            )
                            if curr_formula:
                                ref_values = {}
                                curr_formula_inputs = getattr(
                                    curr_formula, "inputs", None
                                )
                                print(curr_formula_inputs)
                                if isinstance(curr_formula_inputs, set):
                                    # curr_formula_inputs = cast(
                                    #     Set[str], curr_formula.inputs
                                    # )
                                    for ref in curr_formula_inputs:
                                        print(ref)
                                        curr_ref_value = getattr(
                                            sheet[ref], "value", None
                                        )
                                        print(curr_ref_value)
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
                                        elif curr_ref_value is None:
                                            ref_values[ref] = TokenEmpty()
                                formula_result = curr_formula(ref_values)
                                logger.setLevel(logging.INFO)
                                print(f"formula_result: {formula_result}")
                                if isinstance(formula_result, bool):
                                    if formula_result:
                                        print(
                                            f"process: Applying differential style with index: {dxf_id}"
                                        )
                                        for specific_range in cf_range.split(" "):
                                            logging.debug(
                                                f"process: -->> Saving specific_range for each cell of '{specific_range}'"
                                            )
                                            logging.debug(
                                                f"process: -->> CellRange: {sheet[specific_range]}"
                                            )
                                            possible_range = sheet[specific_range]
                                            for row in (
                                                possible_range
                                                if isinstance(possible_range, Iterable)
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
                                                        if cf_stop_if_true is not None
                                                        else False,
                                                    )
                                                    if code in results:
                                                        _, _, old_priority, _, _ = (
                                                            results[code]
                                                        )
                                                        if old_priority <= cf_priority:
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

                row_id += 1
            group_id += 1

    else:
        logging.debug("process: ")
    return results
