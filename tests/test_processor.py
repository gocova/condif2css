import sys

sys.path.append("src")

import pytest
from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import PatternFill

import condif2css.processor as processor


def _make_rule(formulas, dxf_id, priority, stop_if_true=None):
    fill = PatternFill(patternType="solid", fgColor="00FF0000")
    rule = FormulaRule(formula=formulas, fill=fill, stopIfTrue=stop_if_true)
    rule.dxfId = dxf_id
    rule.priority = priority
    return rule


def test_process_applies_expression_with_relative_offsets():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "No"
    ws["A2"] = "Yes"
    ws["A3"] = "Yes"

    ws.conditional_formatting.add(
        "A1:A3",
        _make_rule(['A1="Yes"'], dxf_id=7, priority=3, stop_if_true=True),
    )

    result = processor.process(ws)
    assert set(result.keys()) == {"Sheet1\\!A2", "Sheet1\\!A3"}
    assert result["Sheet1\\!A2"] == ("Sheet1", "A2", 3, 7, True)
    assert result["Sheet1\\!A3"] == ("Sheet1", "A3", 3, 7, True)


def test_process_keeps_lowest_priority_value_for_same_cell():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "Yes"
    ws["F1"] = "ON"

    ws.conditional_formatting.add(
        "A1",
        _make_rule(['$F$1="ON"'], dxf_id=2, priority=2),
    )
    ws.conditional_formatting.add(
        "A1",
        _make_rule(['$F$1="ON"'], dxf_id=1, priority=1),
    )

    result = processor.process(ws)
    assert result == {"Sheet1\\!A1": ("Sheet1", "A1", 1, 1, False)}


def test_process_skips_rules_without_dxf_or_with_multiple_formulas():
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "Yes"

    ws.conditional_formatting.add(
        "A1",
        _make_rule(['A1="Yes"'], dxf_id=None, priority=1),
    )
    ws.conditional_formatting.add(
        "A1",
        _make_rule(['A1="Yes"', 'A1="No"'], dxf_id=9, priority=1),
    )

    assert processor.process(ws) == {}


def test_process_can_raise_when_fail_ok_is_false(monkeypatch):
    class ExplodingFormula:
        inputs = set()

        def __call__(self, _):
            raise RuntimeError("boom")

    monkeypatch.setattr(processor, "get_interpreter", lambda _: ExplodingFormula())

    wb = Workbook()
    ws = wb.active
    ws["A1"] = "X"
    ws.conditional_formatting.add("A1", _make_rule(['A1="X"'], dxf_id=1, priority=1))

    with pytest.raises(RuntimeError, match="boom"):
        processor.process(ws, fail_ok=False)

    assert processor.process(ws, fail_ok=True) == {}


def test_process_returns_empty_when_conditional_formatting_is_missing():
    class DummySheet:
        conditional_formatting = None

    assert processor.process(DummySheet()) == {}

