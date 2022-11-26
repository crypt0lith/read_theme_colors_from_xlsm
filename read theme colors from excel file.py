from openpyxl.xml.functions import fromstring, QName
from openpyxl import load_workbook

data_file = 'data/exampleFile.xlsm'

wb = load_workbook(data_file)
xmlns = 'http://schemas.openxmlformats.org/drawingml/2006/main'
root = fromstring(wb.loaded_theme)
themeEl = root.find(QName(xmlns, 'themeElements').text)
colorSchemes = themeEl.findall(QName(xmlns, 'clrScheme').text)
firstColorScheme = colorSchemes[0]

colors = []
for c in ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6']:
    accent = firstColorScheme.find(QName(xmlns, c).text)

    if 'window' in accent.getchildren()[0].attrib['val']:
        colors.append(accent.getchildren()[0].attrib['lastClr'])
    else:
        colors.append(accent.getchildren()[0].attrib['val'])
print(colors)
