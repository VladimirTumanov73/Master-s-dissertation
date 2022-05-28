import xml.etree.ElementTree as xml
import xlwings as xw
import time

start_time = time.time()

wb = xw.Book('FDTest1.xlsx')
Sheets1 = wb.sheets[0]
Sheets2 = wb.sheets[1]
Sheets3 = wb.sheets[2]
Sheets4 = wb.sheets[3]
Sheets5 = wb.sheets[4]
Sheets6 = wb.sheets[5]

def module1(filename):
    # Start with the root element
    root = xml.Element("forestDeclaration")
    xml.register_namespace("d2p1", "http://rosleshoz.gov.ru/xmlns/cTypes")
    number = xml.SubElement(root, "number")
    number.text = str(Sheets1.range('C7').options(index=False).value)
    date = xml.SubElement(root, "date")
    date.text = str(Sheets1.range('C8').options(index=False).value)
    # Header
    header = xml.Element("header")
    root.append(header)
    ex_Aut = xml.SubElement(header, "executiveAuthority")
    ex_Aut.text = str(Sheets1.range('B11').options(index=False).value)
    period = xml.Element("period")
    header.append(period)
    d3p1_begin = xml.SubElement(header, "d3p1:begin")
    d3p1_begin.text = str(Sheets1.range('C9').options(index=False).value)
    d3p1_end = xml.SubElement(header, "d3p1:end")
    d3p1_end.text = str(Sheets1.range('C10').options(index=False).value)
    partner = xml.Element("partner")
    header.append(partner)
    juridicalPerson = xml.Element("juridicalPerson")
    partner.append(juridicalPerson)
    d4p1_name = xml.SubElement(juridicalPerson, "d4p1:name")
    d4p1_name.text = str(Sheets1.range('B12').options(index=False).value)
    d4p1_inn = xml.SubElement(juridicalPerson, "d4p1:inn")
    d4p1_inn.text = str(Sheets1.range('B13').options(index=False).value)
    signerData = xml.Element("signerData")
    header.append(signerData)
    d3p1_employee = xml.Element("d3p1:employee")
    signerData.append(d3p1_employee)
    d3p1_first_name = xml.SubElement(d3p1_employee, "d3p1:first_name")
    d3p1_first_name.text = str(Sheets1.range('C14').options(index=False).value)
    d3p1_post = xml.SubElement(d3p1_employee, "d3p1:post")
    d3p1_post.text = str(Sheets1.range('C15').options(index=False).value)
    d3p1_basisAuthority = xml.SubElement(d3p1_employee, "d3p1:basisAuthority")
    d3p1_basisAuthority.text = str(Sheets1.range('C16').options(index=False).value)
    contract = xml.Element("contract")
    header.append(contract)
    d3p1_type = xml.SubElement(contract, "d3p1:type")
    d3p1_type.text = str(Sheets1.range('C19').options(index=False).value)
    d3p1_number = xml.SubElement(contract, "d3p1:number")
    d3p1_number.text = str(Sheets1.range('C20').options(index=False).value)
    d3p1_date = xml.SubElement(contract, "d3p1:date")
    d3p1_date.text = str(Sheets1.range('C21').options(index=False).value)
    # HarvestingWood
    hW = xml.Element("harvestingWood")
    root.append(hW)
    row_hW = xml.Element("row")
    hW.append(row_hW)
    specialPurpose = xml.SubElement(row_hW, "specialPurpose")
    specialPurpose.text = " "
    protectionCategory = xml.SubElement(row_hW, "protectionCategory")
    protectionCategory.text = " "
    location = xml.Element("location")
    location.append(row_hW)
    forestry = xml.SubElement(row_hW, "forestry")
    forestry.text = str(Sheets2.range('I12').options(index=False).value)
    subforestry = xml.SubElement(row_hW, "subforestry")
    subforestry.text = str(Sheets2.range('W12').options(index=False).value)
    quarter = xml.SubElement(row_hW, "quarter")
    quarter.text = str(Sheets2.range('AW12').options(index=False).value)
    taxationUnit = xml.SubElement(row_hW, "taxationUnit")
    taxationUnit.text = str(Sheets2.range('BD12').options(index=False).value)
    cuttingArea = xml.SubElement(row_hW, "cuttingArea")
    cuttingArea.text = str(Sheets2.range('BK12').options(index=False).value)
    area = xml.SubElement(row_hW, "area")
    area.text = str(Sheets2.range('BK12').options(index=False).value)
    formCutting = xml.SubElement(row_hW, "formCutting")
    formCutting.text = str(Sheets2.range('BU12').options(index=False).value)
    typeCutting = xml.SubElement(row_hW, "typeCutting")
    typeCutting.text = str(Sheets2.range('CB12').options(index=False).value)
    farm = xml.SubElement(row_hW, "farm")
    farm.text = str(Sheets2.range('CI12').options(index=False).value)
    tree = xml.SubElement(row_hW, "tree")
    tree.text = str(Sheets2.range('CP12').options(index=False).value)
    unitType = xml.SubElement(row_hW, "unitType")
    unitType.text = str(Sheets2.range('CW12').options(index=False).value)
    volume = xml.SubElement(row_hW, "volume")
    volume.text = str(Sheets2.range('DD12').options(index=False).value)


    # Output
    xml.indent(root, space=' ', level=0)

    tree1 = xml.ElementTree(root)

    with open("FD.xml", 'wb') as f:
        tree1.write(f, encoding='utf-8', xml_declaration=True)

    return tree1




if __name__ == "__main__":
    module1("FD.xml")


print(time.time() - start_time, "seconds")
