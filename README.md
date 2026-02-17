# condif2css

`condif2css` helps translate Excel conditional formatting into reusable CSS classes.

The library is designed to work with `openpyxl` workbooks and is currently focused on:

- Evaluating conditional-format formulas per cell (`processor.process`)
- Resolving workbook/theme/indexed colors to aRGB (`create_themed_get_css_color`)
- Building deduplicated CSS rules from `Cell` / `DifferentialStyle` objects (`css.py`)

## What It Does

1. Reads conditional-format rules from an `openpyxl` worksheet.
2. Evaluates each rule formula against each target cell (with relative/absolute ref offsets).
3. Returns matching `(sheet, cell, priority, dxf_id, stop_if_true)` metadata.
4. Converts style objects into CSS declarations and stable class names.

## Installation

Install from PyPI:

```bash
pip install condif2css
```

If the latest published version is a pre-release, use:

```bash
pip install --pre condif2css
```

Local development install:

```bash
pip install -e .
```

Or with `uv` (including dev dependencies):

```bash
uv sync --group dev
```

## Quick Example

```python
from openpyxl import load_workbook

from condif2css import create_themed_get_css_color, get_theme_colors
from condif2css.color import aRGB_to_css
from condif2css.css import CssBuilder, CssRulesRegistry, create_get_css_from_cell
from condif2css.processor import process

wb = load_workbook("input.xlsx", data_only=True)
ws = wb["Sheet1"]

theme_colors = get_theme_colors(wb)
get_argb = create_themed_get_css_color(theme_colors)

def get_css_color(color):
    argb = get_argb(color)
    return aRGB_to_css(argb) if isinstance(argb, str) else None

css_builder = CssBuilder(get_css_color)
css_registry = CssRulesRegistry(prefix="cf")
get_css_from_cell = create_get_css_from_cell(css_registry, css_builder)

# { "Sheet1\\!A1": ("Sheet1", "A1", priority, dxf_id, stop_if_true), ... }
matched_rules = process(ws)

# Apply each matched differential style and collect class names per cell
cell_classes = {}
for code, (_, _, _, dxf_id, _) in matched_rules.items():
    dxf = wb._differential_styles[dxf_id]  # openpyxl differential style
    cell_classes[code] = sorted(get_css_from_cell(dxf, is_important=True))

# Final CSS text
css_text = "\n".join(css_registry.get_rules())
```

## Public API (Current)

- `condif2css.create_themed_get_css_color(theme_colors)`
- `condif2css.get_theme_colors(workbook)`
- `condif2css.processor.process(sheet, fail_ok=True)`
- `condif2css.css.CssBuilder`
- `condif2css.css.CssRulesRegistry`
- `condif2css.css.create_get_css_from_cell(...)`
- `condif2css.color.aRGB_to_css(...)` and color/tint utility helpers

## Current Scope and Limitations

- Conditional-format evaluation currently expects rules with exactly one formula.
- Formula evaluation expects references that resolve to single cells.
- Non-boolean formula results are ignored.
- Style extraction supports:
  - Borders: top/right/bottom/left
  - Alignment: horizontal/vertical
  - Fill: solid fills (or `DifferentialStyle.fill.bgColor`)
  - Font: size, color, bold, italic, underline
- The package provides building blocks; no CLI is included.

## Development

Run tests:

```bash
pytest
```

## Acknowledgements

- Theme color extraction is based on public-domain work by Mike-honey:
  <https://gist.github.com/Mike-Honey/b36e651e9a7f1d2e1d60ce1c63b9b633>

## License

Dual-licensed under:

- Apache-2.0 (`LICENSE_APACHE`)
- MIT (`LICENSE_MIT`)

You may choose either license.
