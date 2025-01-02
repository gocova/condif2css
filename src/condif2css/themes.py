# Using source from: https://gist.github.com/Mike-Honey/b36e651e9a7f1d2e1d60ce1c63b9b633
#
# Author: Mike-honey (https://gist.github.com/Mike-Honey)
# License: Public domain (https://gist.github.com/Mike-Honey/c24d979c6d626e0be6b543be01671e34)
from openpyxl.workbook import Workbook


def get_theme_colors(wb: Workbook):
    """Gets theme colors from the workbook"""

    # see: https://groups.google.com/forum/#!topic/openpyxl-users/I0k3TfqNLrc
    from openpyxl.xml.functions import QName, fromstring

    xlmns = "http://schemas.openxmlformats.org/drawingml/2006/main"
    root = fromstring(wb.loaded_theme)
    themeEl = root.find(QName(xlmns, "themeElements").text)
    colorSchemes = themeEl.findall(QName(xlmns, "clrScheme").text)
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

        if "window" in accent.getchildren()[0].attrib["val"]:
            colors.append(accent.getchildren()[0].attrib["lastClr"])
        else:
            colors.append(accent.getchildren()[0].attrib["val"])

    return colors
