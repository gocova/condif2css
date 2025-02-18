import logging
from hashlib import blake2b
from typing import Any, Callable, Dict, Iterable, List, Tuple, Literal
from openpyxl.cell import Cell
from openpyxl.styles.colors import Color
from openpyxl.styles.differential import DifferentialStyle


DEFAULT_BORDER_STYLE = [
    ("border-{direction}-style", "solid"),
    ("border-{direction}-width", "1px"),
]

BORDER_STYLES = {
    "dashDot": None,
    "dashDotDot": None,
    "dashed": [("border-{direction}-style", "dashed")],
    "dotted": [("border-{direction}-style", "dotted")],
    "double": [("border-{direction}-style", "double")],
    "hair": None,
    "medium": [
        ("border-{direction}-style", "solid"),
        ("border-{direction}-width", "2px"),
    ],
    "mediumDashDot": [
        ("border-{direction}-style", "solid"),
        ("border-{direction}-width", "2px"),
    ],
    "mediumDashDotDot": [
        ("border-{direction}-style", "solid"),
        ("border-{direction}-width", "2px"),
    ],
    "mediumDashed": [
        ("border-{direction}-style", "dashed"),
        ("border-{direction}-width", "2px"),
    ],
    "slantDashDot": None,
    "thick": [
        ("border-{direction}-style", "solid"),
        ("border-{direction}-width", "1px;"),
    ],
    "thin": [
        ("border-{direction}-style", "solid"),
        ("border-{direction}-width", "1px"),
    ],
}


class CssBuilder:
    def __init__(self, get_css_color: Callable[[Color], str | None]) -> None:
        self.get_css_color = get_css_color

    def font_size(self, size: int, is_important: bool = False) -> Tuple[str, str]:
        is_important_label = " !important" if is_important else ""
        return "font-size", f"{size}px{is_important_label}"

    def height(self, size: int, is_important: bool = False) -> Tuple[str, str]:
        is_important_label = " !important" if is_important else ""
        return "height", f"{size}px{is_important_label}"

    def font_color(
        self, color: Color, is_important: bool = False
    ) -> Tuple[str, str] | None:
        css_color = self.get_css_color(color)
        if css_color is None:
            return None
        is_important_label = " !important" if is_important else ""
        return "color", f"{css_color}{is_important_label}"

    def background_color(
        self, color: Color, is_important: bool = False
    ) -> Tuple[str, str] | None:
        css_color = self.get_css_color(color)
        if css_color is None:
            return None
        is_important_label = " !important" if is_important else ""
        return "background-color", f"{css_color}{is_important_label}"

    def background_transparent(self, is_important: bool = False) -> Tuple[str, str]:
        is_important_label = " !important" if is_important else ""
        return "background-color", f"transparent{is_important_label}"

    def font_underline(self, is_important: bool = False) -> Tuple[str, str]:
        is_important_label = " !important" if is_important else ""
        return "font-decoration", f"underline{is_important_label}"

    def font_bold(self, is_important: bool = False) -> Tuple[str, str]:
        is_important_label = " !important" if is_important else ""
        return "font-weight", f"bold{is_important_label}"

    def font_italic(self, is_important: bool = False) -> Tuple[str, str]:
        is_important_label = " !important" if is_important else ""
        return "font-style", f"italic{is_important_label}"

    def text_align_horizontal(
        self, horizontal, is_important: bool = False
    ) -> Tuple[str, Any] | None:
        if not isinstance(horizontal, str):
            return None
        is_important_label = " !important" if is_important else ""
        return "text-align", f"{horizontal}{is_important_label}"

    def text_align_vertical(
        self, vertical, is_important: bool = False
    ) -> Tuple[str, Any] | None:
        if not isinstance(vertical, str):
            return None
        is_important_label = " !important" if is_important else ""
        return "vertical-align", f"{vertical}{is_important_label}"

    def border(
        self,
        style: str | None,
        direction: Literal["right", "left", "top", "bottom"],
        color: Color,
        is_important: bool = False,
    ) -> List[Tuple[str, str]] | None:
        if style is None:
            return None

        is_important_label = " !important" if is_important else ""

        border_style = BORDER_STYLES.get(style)
        border_style = [
            (x[0].format(direction=direction), f"{x[1]}{is_important_label}")
            for x in (DEFAULT_BORDER_STYLE if border_style is None else border_style)
        ]

        css_color = self.get_css_color(color)

        if css_color is not None:
            border_style.append(
                (f"border-{direction}-color", f"{css_color}{is_important_label}")
            )

        return border_style


class CssRulesRegistry:
    def __init__(self, prefix: str = "xx2h") -> None:
        self._prefix = prefix
        self._rules: List[str] = []
        self._rules_source: List[Dict[str, str]] = []
        self._classnames: List[str] = []
        self._hash_rels: Dict[str, int] = {}

    def register(self, items: Iterable) -> str:
        new_rule = dict(items)
        css_rule_contents = "\n\t".join([f"{x[0]}: {x[1]};" for x in new_rule.items()])
        css_rule_contents = f"{{\n\t{css_rule_contents}\n}}"

        hash_object = blake2b(digest_size=12)

        hash_object.update(f"{len(css_rule_contents)}|{css_rule_contents}".encode())

        css_rule_hash = hash_object.hexdigest()

        if css_rule_hash in self._hash_rels:
            rule_index = self._hash_rels[css_rule_hash]
            print(f"rule[{rule_index}]: {self._classnames[rule_index]}")
            return self._classnames[rule_index]

        rule_index = len(self._rules)
        classname = f"{self._prefix}_x{hex(rule_index)[2:].zfill(4)}"
        self._rules.append(css_rule_contents)
        self._rules_source.append(new_rule)
        self._classnames.append(classname)
        self._hash_rels[css_rule_hash] = rule_index

        print(f"rule[{rule_index}]: {classname}")

        return classname

    def get_rules(self) -> List[str]:
        return [
            f"{self._classnames[index]} {css_rule}"
            for index, css_rule in enumerate(self._rules)
        ]


def get_border_styles_from_cell(
    cell: Cell | DifferentialStyle,
    css_builder: CssBuilder,
    is_important: bool = False,
) -> List[Tuple[str, str]]:
    border_styles = []

    cell_border = getattr(cell, "border")
    if cell_border is None:
        return border_styles
    for border_direction in ["right", "left", "top", "bottom"]:
        border_style = getattr(cell_border, border_direction)
        # print(border_style)
        if not border_style:
            continue

        border_css = css_builder.border(
            border_style.style,
            direction=border_direction,  # type: ignore
            color=border_style.color,
            is_important=is_important,
        )
        if border_css is not None:
            border_styles = border_styles + border_css

    return border_styles


def create_get_css_from_cell(css_registry: CssRulesRegistry, css_builder: CssBuilder):
    def get_css_from_cell(
        cell: Cell | DifferentialStyle,
        merged_cell_map=None,
        is_important: bool = False,
    ):
        nonlocal css_builder

        # print(cell)
        cell_classes = set()

        css_borders = get_border_styles_from_cell(
            cell, css_builder, is_important=is_important
        )

        merged_cell_map = merged_cell_map or {}
        if merged_cell_map:
            # TODO edged_cells
            for m_cell in merged_cell_map["cells"]:
                css_borders = css_borders + get_border_styles_from_cell(
                    m_cell, css_builder, is_important=is_important
                )

        if len(css_borders) > 0:
            cell_classes.update([css_registry.register(css_borders)])

        css_contents = []
        cell_alignment = getattr(cell, "alignment")
        if cell_alignment:
            horizontal_alignment = css_builder.text_align_horizontal(
                getattr(cell_alignment, "horizontal"), is_important=is_important
            )
            if horizontal_alignment is not None:
                css_contents.append(horizontal_alignment)

            vertical_alignment = css_builder.text_align_vertical(
                getattr(cell_alignment, "vertical"), is_important=is_important
            )
            if vertical_alignment is not None:
                css_contents.append(vertical_alignment)

        if len(css_contents) > 0:
            cell_classes.update([css_registry.register(css_contents)])

        css_color = []
        cell_fill = getattr(cell, "fill")
        print(f"--> cell.fill: {cell_fill}")
        if cell_fill is not None:
            print(f"-->> {isinstance(cell, DifferentialStyle)}")

            if not isinstance(cell, DifferentialStyle):
                cell_fill_pattern_type = getattr(cell_fill, "patternType")
                print(f"--> --> patternType: {cell_fill_pattern_type}")
                if cell_fill_pattern_type is not None:
                    if cell_fill_pattern_type == "solid":
                        background_color = css_builder.background_color(
                            getattr(cell_fill, "fgColor"), is_important=is_important
                        )
                        if background_color is not None:
                            css_color.append(background_color)
                    elif cell_fill_pattern_type is not None:
                        # TODO patternType != 'solid'
                        logging.warning(
                            f"css (components): Pattern type is not supported: {cell_fill_pattern_type}"
                        )
            else:
                cell_bgcolor = getattr(cell_fill, "bgColor")

                if cell_bgcolor is not None:
                    background_color = css_builder.background_color(
                        cell_bgcolor, is_important=is_important
                    )
                    if background_color is not None:
                        css_color.append(background_color)

        if len(css_color) > 0:
            cell_classes.update([css_registry.register(css_color)])

        css_font = []
        cell_font = getattr(cell, "font")
        print(f"--> cell.font: {cell_font}")
        if cell_font is not None:
            cell_font_size = getattr(cell_font, "sz")
            if cell_font_size:
                cell_font_size = int(cell_font_size)

                css_font.append(
                    css_builder.font_size(cell_font_size, is_important=is_important)
                )

            cell_font_color = getattr(cell_font, "color")
            if cell_font_color is not None:
                css_font_color = css_builder.font_color(
                    cell_font_color, is_important=is_important
                )
                if css_font_color is not None:
                    css_font.append(css_font_color)

            cell_font_b = getattr(cell_font, "b")
            if cell_font_b is not None:
                css_font.append(css_builder.font_bold(is_important=is_important))

            cell_font_i = getattr(cell_font, "i")
            if cell_font_i is not None:
                css_font.append(css_builder.font_italic(is_important=is_important))

            cell_font_u = getattr(cell_font, "u")
            if cell_font_u is not None:
                css_font.append(css_builder.font_underline(is_important=is_important))

        if len(css_font) > 0:
            cell_classes.update([css_registry.register(css_font)])

        return cell_classes

    return get_css_from_cell
