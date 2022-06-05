from FDM1 import module1
# import xml.etree.ElementTree as xml
import xlwings as xw


def test_module1():
    wb = xw.Book("FDTest1.xlsx")
    sheets_1 = wb.sheets[0]
    jur_per_1 = str(sheets_1.range('B12').options(index=False).value)
    # tree = ET.parse('FDTest1.xml')
    # root = tree.getroot()
    assert module1("FDTest1.xlsx") == jur_per_1
