# condif2css

[![PyPI Version](https://img.shields.io/pypi/v/condif2css.svg)](https://pypi.org/project/condif2css/)
[![License](https://img.shields.io/badge/License-MIT%20%2F%20Apache%202.0-green.svg)](https://opensource.org/licenses/)
[![Buy Me a Coffee](https://img.shields.io/badge/Buy%20Me%20a%20Coffee-Support-orange?logo=buy-me-a-coffee&style=flat-square)](https://buymeacoffee.com/gocova)

`condif2css` helps translate Excel conditional formatting into reusable CSS classes.

The library is designed to work with `openpyxl` workbooks and is currently focused on:

- Evaluating conditional-format formulas per cell (`processor.process`)
- Resolving workbook/theme/indexed colors to aRGB (`create_themed_css_color_resolver`)
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

from condif2css import (
    CssBuilder,
    CssRulesRegistry,
    argb_to_css,
    create_get_css_from_cell,
    create_themed_css_color_resolver,
    get_differential_style,
    get_theme_colors,
    process_conditional_formatting,
)

wb = load_workbook("input.xlsx", data_only=True)
ws = wb["Sheet1"]

theme_colors = get_theme_colors(wb)
get_argb = create_themed_css_color_resolver(theme_colors)

def get_css_color(color):
    argb = get_argb(color)
    return argb_to_css(argb) if isinstance(argb, str) else None

css_builder = CssBuilder(get_css_color)
css_registry = CssRulesRegistry(prefix="cf")
get_css_from_cell = create_get_css_from_cell(css_registry, css_builder)

# { "Sheet1\\!A1": ("Sheet1", "A1", priority, dxf_id, stop_if_true), ... }
matched_rules = process_conditional_formatting(ws)

# Apply each matched differential style and collect class names per cell
cell_classes = {}
for code, (_, _, _, dxf_id, _) in matched_rules.items():
    dxf = get_differential_style(wb, dxf_id)
    if dxf is None:
        continue
    cell_classes[code] = sorted(get_css_from_cell(dxf, is_important=True))

# Final CSS text
css_text = "\n".join(css_registry.get_rules())
```

## Public API (Current)

- `condif2css.create_themed_css_color_resolver(theme_colors)`
- `condif2css.get_theme_colors(workbook, strict=False)`
- `condif2css.process_conditional_formatting(sheet, fail_ok=True)`
- `condif2css.get_differential_style(workbook, dxf_id)`
- `condif2css.ThemeColorsError`
- `condif2css.CssBuilder`
- `condif2css.CssRulesRegistry`
- `condif2css.create_get_css_from_cell(...)`
- `condif2css.color.argb_to_css(...)` and color/tint utility helpers

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

Licensed under either of:

- Apache License, Version 2.0 ([`LICENSE_APACHE`](LICENSE_APACHE) or <https://www.apache.org/licenses/LICENSE-2.0>)
- MIT license ([`LICENSE_MIT`](LICENSE_MIT) or <https://opensource.org/licenses/MIT>)

at your option.
