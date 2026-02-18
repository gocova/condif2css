# Copyright (c) 2026 Jose Gonzalo Covarrubias M <gocova.dev@gmail.com>
#
# Part of: batch_xlsx2html (bxx2html)
#

import logging
import blake3
from typing import Any, Callable, Dict, Iterable, List, Set, Tuple, Literal
from openpyxl.cell import Cell, MergedCell
from openpyxl.styles.colors import Color
from openpyxl.styles.differential import DifferentialStyle


DEFAULT_BORDER_STYLE = [
    ("border-{direction}-style", "solid"),
    ("border-{direction}-width", "1px"),
]

BORDER_STYLES = {
    "dashDot": [("border-{direction}-style", "dashed")],
    "dashDotDot": [("border-{direction}-style", "dashed")],
    "dashed": [("border-{direction}-style", "dashed")],
    "dotted": [("border-{direction}-style", "dotted")],
    "double": [("border-{direction}-style", "double")],
    "hair": [
        ("border-{direction}-style", "solid"),
        ("border-{direction}-width", "1px"),
    ],
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
    "slantDashDot": [("border-{direction}-style", "dashed")],
    "thick": [
        ("border-{direction}-style", "solid"),
        ("border-{direction}-width", "3px"),
    ],
    "thin": [
        ("border-{direction}-style", "solid"),
        ("border-{direction}-width", "1px"),
    ],
}


class CssBuilder:
    def __init__(self, get_css_color: Callable[[Color | None], str | None]) -> None:
        """
        Initializes a CssBuilder instance.

        Args:
            get_css_color (Callable[[Color], str | None]): A function that takes
                a Color and returns its CSS representation as a string, or None.
        """
        self.get_css_color = get_css_color

    def font_size(self, size: int, is_important: bool = False) -> Tuple[str, str]:
        """
        Returns a tuple containing the CSS property "font-size" and its value.

        Args:
            size (int): The font size in pixels.
            is_important (bool, optional): Whether to include "!important" in the CSS declaration. Defaults to False.

        Returns:
            Tuple[str, str]: A tuple containing the CSS property "font-size" and its value.
        """
        is_important_label = " !important" if is_important else ""
        return "font-size", f"{size}px{is_important_label}"

    def height(self, size: int, is_important: bool = False) -> Tuple[str, str]:
        """
        Returns a tuple containing the CSS property "height" and its value.

        Args:
            size (int): The height in pixels.
            is_important (bool, optional): Whether to include "!important" in the CSS declaration. Defaults to False.

        Returns:
            Tuple[str, str]: A tuple containing the CSS property "height" and its value.
        """
        is_important_label = " !important" if is_important else ""
        return "height", f"{size}px{is_important_label}"

    def font_color(
        self, color: Color, is_important: bool = False
    ) -> Tuple[str, str] | None:
        """
        Returns a tuple containing the CSS property "color" and its value.

        Args:
            color (Color): The color to use.
            is_important (bool, optional): Whether to include "!important" in the CSS declaration. Defaults to False.

        Returns:
            Tuple[str, str] | None: A tuple containing the CSS property "color" and its value, or None if the color is invalid.
        """
        css_color = self.get_css_color(color)
        if css_color is None:
            return None
        is_important_label = " !important" if is_important else ""
        return "color", f"{css_color}{is_important_label}"

    def background_color(
        self, color: Color, is_important: bool = False
    ) -> Tuple[str, str] | None:
        """
        Returns a tuple containing the CSS property "background-color" and its value.

        Args:
            color (Color): The color to use.
            is_important (bool, optional): Whether to include "!important" in the CSS declaration. Defaults to False.

        Returns:
            Tuple[str, str] | None: A tuple containing the CSS property "background-color" and its value, or None if the color is invalid.
        """
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
        return "text-decoration", f"underline{is_important_label}"

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
        color: Color | None,
        is_important: bool = False,
    ) -> List[Tuple[str, str]] | None:
        """
        Returns a list of tuples containing the CSS property "border-{direction}" and its value, for the given style, direction and color.

        Args:
            style (str | None): The border style.
            direction (Literal["right", "left", "top", "bottom"]): The direction of the border.
            color (Color): The color of the border.
            is_important (bool, optional): Whether to include "!important" in the CSS declaration. Defaults to False.

        Returns:
            List[Tuple[str, str]] | None: A list of tuples containing the CSS property "border-{direction}" and its value, or None if the style is invalid.
        """
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
    def __init__(self, prefix: str = "xx2h", digest_size: int = 28) -> None:
        """
        Initializes a CssRulesRegistry instance.

        Args:
            prefix (str, optional): The prefix to use for generating classnames. Defaults to "xx2h".
            digest_size (int, optional): The size of the hash to generate for each rule. Defaults to 28.

        The CssRulesRegistry instance will store a mapping of hash values to their corresponding CSS rules, as well as the classnames and the source items for each rule.

        The instance will be initialized with an empty mapping of hash values to rules, and an empty list of classnames.

        The digest size is used to generate a stable hash for each rule, and can be adjusted to trade off between hash quality and performance.

        :param prefix: The prefix to use for generating classnames.
        :param digest_size: The size of the hash to generate for each rule.
        :type prefix: str
        :type digest_size: int
        """
        self._prefix = prefix
        # self._rules: List[str] = []
        # self._rules_source: List[Dict[str, str]] = []
        # self._classnames: List[str] = []
        # self._hash_rels: Dict[str, int] = {}
        self._existing_rules: Dict[
            str,
            Tuple[
                str,  # classname
                str,  # css_rule_contents
                Dict[str, str],  # rule_source
            ],
        ] = {}
        self._digest_size = digest_size

    def register(self, items: Iterable) -> str:
        """
        Registers a new CSS rule based on the given items.

        The items will be sorted to ensure consistent rule generation and hashing.
        The CSS rule string will be built by joining the items with a colon and a newline.
        A stable hash will be generated for the rule using the blake3 algorithm.
        If a rule with the same hash already exists, the existing classname will be returned.
        Otherwise, a new classname will be generated and the rule will be registered.

        Parameters
        ----------
        items : Iterable
            A collection of key-value pairs to build the CSS rule from.

        Returns
        -------
        str
            The classname associated with the registered CSS rule.
        """

        # Sort the input to ensure consistent rule generation and hashing
        sorted_items = sorted(items)

        # Build CSS rule string
        css_properties = "\n\t".join(f"{k}: {v};" for k, v in sorted_items)
        css_rule_contents = f"{{\n\t{css_properties}\n}}"

        # Generate a stable hash for the rule
        hash_input = f"{len(css_rule_contents)}|{css_rule_contents}".encode()
        css_rule_hash = blake3.blake3(hash_input).hexdigest(length=self._digest_size)

        # Check if this rule already exists
        if css_rule_hash in self._existing_rules:
            classname, _, _ = self._existing_rules[css_rule_hash]
            logging.debug(f"register: rule[{css_rule_hash}] --> {classname}")
            return classname

        # Register new rule
        rule_count = len(self._existing_rules)
        classname = f"{self._prefix}_x{hex(rule_count)[2:].zfill(4)}"
        new_rule = dict(sorted_items)

        self._existing_rules[css_rule_hash] = (classname, css_rule_contents, new_rule)

        logging.debug(f"register: rule[{css_rule_hash}] --> {classname}")

        return classname

    def get_rules(self) -> List[str]:
        return [f".{t[0]} {t[1]}" for _, t in self._existing_rules.items()]


def get_border_styles_from_cell(
    cell: Cell | MergedCell | DifferentialStyle,
    css_builder: CssBuilder,
    is_important: bool = False,
) -> List[Tuple[str, str]]:
    """
    Returns a list of tuples, where each tuple contains a CSS property and its value, representing the border styles of a cell.

    Args:
        cell (Cell | MergedCell | DifferentialStyle): The cell from which to extract the border styles.
        css_builder (CssBuilder): The builder used to construct the CSS rules.
        is_important (bool, optional): Whether to include "!important" in the CSS declaration. Defaults to False.

    Returns:
        List[Tuple[str, str]]: A list of tuples, where each tuple contains a CSS property and its value, representing the border styles of the cell.
    """
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
    """
    Creates a function that returns a set of CSS classes representing the styles of a cell.

    The returned function takes a cell, a map of merged cells, and a boolean indicating whether
    the styles should be marked as important.

    The function will return a set of CSS classes, each representing a CSS rule or ruleset.
    The CSS classes will be registered in the provided CssRulesRegistry.

    The function will extract the following styles from the cell:

    - Border styles
    - Alignment (horizontal and vertical)
    - Fill color (if the cell has a solid fill pattern)
    - Font styles (font size, color, bold, italic, underline)

    :param css_registry: The registry in which the CSS classes will be registered
    :param css_builder: The builder used to construct the CSS rules
    :return: A function that takes a cell, a map of merged cells, and a boolean indicating whether the styles should be marked as important, and returns a set of CSS classes
    """
    def get_css_from_cell(
        cell: Cell | MergedCell | DifferentialStyle,
        merged_cell_map=None,
        is_important: bool = False,
    ) -> Set[str]:
        """
        Returns a set of CSS classes representing the styles of a cell.

        The returned function takes a cell, a map of merged cells, and a boolean indicating whether
        the styles should be marked as important.

        The function will return a set of CSS classes, each representing a CSS rule or ruleset.
        The CSS classes will be registered in the provided CssRulesRegistry.

        The function will extract the following styles from the cell:

        - Border styles
        - Alignment (horizontal and vertical)
        - Fill color (if the cell has a solid fill pattern)
        - Font styles (font size, color, bold, italic, underline)

        :param cell: The cell from which to extract the styles
        :param merged_cell_map: A map of merged cells, for which the styles will also be extracted
        :param is_important: A boolean indicating whether the styles should be marked as important
        :return: A set of CSS classes representing the styles of the cell
        """
        nonlocal css_builder

        # print(cell)
        cell_classes = set()

        css_borders = get_border_styles_from_cell(
            cell, css_builder, is_important=is_important
        )

        merged_cells = []
        if isinstance(merged_cell_map, dict):
            merged_cells = merged_cell_map.get("cells") or []
        if merged_cells:
            # TODO edged_cells
            for m_cell in merged_cells:
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
        logging.debug(f"get_css_from_cell: Processing --> cell.fill: {cell_fill}")
        if cell_fill is not None:
            logging.debug(
                f"get_css_from_cell: got DifferentialStyle -->> {isinstance(cell, DifferentialStyle)}"
            )

            cell_fill_pattern_type = getattr(cell_fill, "patternType", None)
            is_differential_style = isinstance(cell, DifferentialStyle)
            primary_fill_color = (
                getattr(cell_fill, "fgColor", None)
                if is_differential_style
                else getattr(cell_fill, "fgColor", None)
            )
            secondary_fill_color = getattr(cell_fill, "bgColor", None)

            if cell_fill_pattern_type == "solid":
                background_color = css_builder.background_color(
                    primary_fill_color or secondary_fill_color,
                    is_important=is_important,
                )
                if background_color is not None:
                    css_color.append(background_color)
            elif cell_fill_pattern_type == "none":
                css_color.append(
                    css_builder.background_transparent(is_important=is_important)
                )
            elif cell_fill_pattern_type is not None:
                # Excel pattern fills do not map 1:1 to CSS; approximate with a flat color.
                background_color = css_builder.background_color(
                    primary_fill_color or secondary_fill_color,
                    is_important=is_important,
                )
                if background_color is not None:
                    css_color.append(background_color)
                logging.warning(
                    f"css (components): Pattern type is approximated as flat color: {cell_fill_pattern_type}"
                )

        if len(css_color) > 0:
            cell_classes.update([css_registry.register(css_color)])

        css_font = []
        cell_font = getattr(cell, "font")
        logging.debug(f"get_css_from_cell: Processing --> cell.font: {cell_font}")
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
            if cell_font_b:
                css_font.append(css_builder.font_bold(is_important=is_important))

            cell_font_i = getattr(cell_font, "i")
            if cell_font_i:
                css_font.append(css_builder.font_italic(is_important=is_important))

            cell_font_u = getattr(cell_font, "u")
            if cell_font_u:
                css_font.append(css_builder.font_underline(is_important=is_important))

        if len(css_font) > 0:
            cell_classes.update([css_registry.register(css_font)])

        return cell_classes

    return get_css_from_cell
