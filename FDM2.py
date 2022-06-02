# import xml.etree.ElementTree as xml
import xlwings as xw
import time
import pytest
import pandas as pd

start_time = time.time()

def module2(filename):
    wb = xw.Book(filename)
    Sheets = wb.sheets[5]
    # dano = Sheets.range('A18:G22').options(index=False).value
    # table = pd.DataFrame(dano)
    # print(table.head())
    # print(table.iloc[0, 4])
    # print(list(table[2]))
    x1 = int(Sheets.range('F18').options(index=False).value)
    x2 = int(Sheets.range('F19').options(index=False).value)
    x3 = int(Sheets.range('F20').options(index=False).value)
    x4 = int(Sheets.range('F21').options(index=False).value)
    x5 = int(Sheets.range('F22').options(index=False).value)
    y1 = int(Sheets.range('G18').options(index=False).value)
    y2 = int(Sheets.range('G19').options(index=False).value)
    y3 = int(Sheets.range('G20').options(index=False).value)
    y4 = int(Sheets.range('G21').options(index=False).value)
    y5 = int(Sheets.range('G22').options(index=False).value)
    dlina_1_2 = ((x2-x1)**2+(y2-y1)**2)**0.5
    dlina_2_3 = ((x3 - x2) ** 2 + (y3 - y2) ** 2) ** 0.5
    dlina_3_4 = ((x4 - x3) ** 2 + (y4 - y3) ** 2) ** 0.5
    dlina_4_5 = ((x5 - x4) ** 2 + (y5 - y4) ** 2) ** 0.5
    dlina_5_1 = ((x1 - x5) ** 2 + (y1 - y5) ** 2) ** 0.5
    # print(int(Sheets.range('E18').options(index=False).value)==int(dlina_1_2),
    #       int(Sheets.range('E19').options(index=False).value)==int(dlina_2_3),
    #       int(Sheets.range('E20').options(index=False).value)==int(dlina_3_4),
    #       int(Sheets.range('E21').options(index=False).value)==int(dlina_4_5),
    #       int(Sheets.range('E22').options(index=False).value)==int(dlina_5_1))
    return int(dlina_1_2)#, int(dlina_2_3), int(dlina_3_4), int(dlina_4_5), int(dlina_5_1)


def test_module2():
    #import xlwings
    wb = xlwings.Book("FDTest2.xlsx")
    Sheets = wb.sheets[5]
    d1 = int(Sheets.range('E18').options(index=False).value)
    d2 = int(Sheets.range('E19').options(index=False).value)
    d3 = int(Sheets.range('E20').options(index=False).value)
    d4 = int(Sheets.range('E21').options(index=False).value)
    d5 = int(Sheets.range('E22').options(index=False).value)
    assert module2("FDTest2.xlsx") == d1#, d2, d3, d4, d5


if __name__ == "__main__":
    module2("FDTest2.xlsx")
    print(time.time() - start_time, "seconds")
    pytest.main(['-v'])




