# Using source from: https://gist.github.com/Mike-Honey/b36e651e9a7f1d2e1d60ce1c63b9b633
#
# Author: Mike-honey (https://gist.github.com/Mike-Honey)
# License: Public domain (https://gist.github.com/Mike-Honey/c24d979c6d626e0be6b543be01671e34)
from openpyxl.workbook import Workbook


class ThemeColorsError(Exception):
    """Raised when workbook theme colors cannot be extracted."""


def get_theme_colors(wb: Workbook, strict: bool = True) -> list[str]:

    """
    Retrieves the colors of a theme from a workbook.

    :param wb: The workbook
    :param strict: If True, raise ThemeColorsError on parsing/structure errors. If False, return an empty list.
    :return: A list of colors in the theme, in the order of light1, dark1, light2, dark2, accent1, accent2, accent3, accent4, accent5, accent6, hyperlink, followedhyperlink.

    :seealso: https://groups.google.com/forum/#!topic/openpyxl-users/I0k3TfqNLrc
    """

    # see: https://groups.google.com/forum/#!topic/openpyxl-users/I0k3TfqNLrc

    from openpyxl.xml.functions import QName, fromstring

    try:
        xlmns = "http://schemas.openxmlformats.org/drawingml/2006/main"
        root = fromstring(wb.loaded_theme)
        themeEl = root.find(QName(xlmns, "themeElements").text)
        if themeEl is None:
            raise ThemeColorsError("Missing 'themeElements' node in workbook theme.")

        colorSchemes = themeEl.findall(QName(xlmns, "clrScheme").text)
        if len(colorSchemes) == 0:
            raise ThemeColorsError("Missing 'clrScheme' node in workbook theme.")
        firstColorScheme = colorSchemes[0]

        colors = []

        for c in [
            "lt1",
            "dk1",
            "lt2",
            "dk2",
            "accent1",
            "accent2",
            "accent3",
            "accent4",
            "accent5",
            "accent6",
            "hlink",
            "folHlink",
        ]:
            accent = firstColorScheme.find(QName(xlmns, c).text)
            if accent is None:
                raise ThemeColorsError(f"Missing '{c}' color node in workbook theme.")

            accent_values = list(accent)
            if len(accent_values) == 0:
                raise ThemeColorsError(f"Color node '{c}' does not contain values.")
            accent_value = accent_values[0].attrib
            val = accent_value.get("val")
            if val is None:
                raise ThemeColorsError(f"Color node '{c}' is missing 'val' attribute.")
            if "window" in val:
                last_clr = accent_value.get("lastClr")
                if last_clr is None:
                    raise ThemeColorsError(
                        f"Color node '{c}' is missing 'lastClr' attribute."
                    )
                colors.append(last_clr)
            else:
                colors.append(val)

        return colors
    except ThemeColorsError:
        if strict:
            raise
        return []
    except Exception as exc:
        if strict:
            raise ThemeColorsError(
                "Unable to parse workbook theme colors."
            ) from exc
        return []
