import logging
from typing import Dict, List, Tuple

from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.formula.tokenizer import Tokenizer


class Interpreter:
    def __init__(self, formula_items) -> None:
        self.formula_items = formula_items if isinstance(formula_items, list) else []
        print(self.formula_items)
        self.operands_jagged_stack = [[]]
        self.solution = None
        self.error = None
        self.current_operands_stack_index = 0
        self.operators_stack = []
        self.reference_operands: Dict[str, Tuple[int, int]] = dict()

        current_operands_stack = self.operands_jagged_stack[
            self.current_operands_stack_index
        ]
        for item in self.formula_items:
            # print(item.type)
            if item.type == "OPERAND":
                # print(item.subtype)
                if item.subtype == "RANGE":
                    cell_reference = item.value
                    index = len(current_operands_stack)
                    self.reference_operands[cell_reference] = (
                        self.current_operands_stack_index,
                        index,
                    )
                    current_operands_stack.append("<range!>")
                else:
                    current_operands_stack.append(item.value)
            elif item.type == "OPERATOR-INFIX":
                print(str(item))
                print(item.value)
            else:
                logging.error(f"Interpreter: Found unsupported formula item: {item}")
                self.error = "Unsupported"

    def resolve_references(self, sheet: Worksheet):
        if self.error is not None:
            return

        for cell_reference, address in self.reference_operands.items():
            cell = sheet[cell_reference]
            if cell is not None:
                cell_value = cell.value

                stack_index, operand_index = address

                stack: List = self.operands_jagged_stack[stack_index]
                stack[operand_index] = (
                    cell_value if not isinstance(cell_value, str) else f'"{cell_value}"'
                )
        self.reference_operands = dict()

    def inspect(self):
        print("--> Start values")
        print(self.error)
        print(self.current_operands_stack_index)
        print(self.operands_jagged_stack)
        print(self.reference_operands)
        print(self.operators_stack)
        print("--> End values")


def parse(formulas: List[str], sheet: Worksheet) -> bool | None:
    if isinstance(formulas, list):
        if len(formulas) == 1:
            current_formula = formulas[0]
            current_formula = (
                f"={current_formula}"
                if not current_formula.startswith("=")
                else current_formula
            )

            formula_items = Tokenizer(current_formula).items
            # print(formula_items)
            interpreter = Interpreter(formula_items)
            interpreter.resolve_references(sheet)
            interpreter.inspect()
        else:
            logging.error(
                f"parse: The interpreter only supports 1 formula, but {len(formulas)} formulas were found!"
            )

    return None
