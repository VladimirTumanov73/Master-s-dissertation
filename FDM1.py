import xml.etree.ElementTree as xml
import xlwings as xw
import time

start_time = time.time()

wb = xw.Book('FDTest1.xlsx')
Sheets1 = wb.sheets[0]


def createXML(filename):
    # Start with the root element
    root = xml.Element("forestDeclaration")
    number = xml.SubElement(root, "number")
    number.text = Sheets1.range('C7').options(index=False).value
    date = xml.SubElement(root, "date")
    date.text = Sheets1.range('C8').options(index=False).value

    header = xml.Element("header")
    root.append(header)

    ex_Aut = xml.SubElement(header, "executiveAuthority")
    ex_Aut.text = Sheets1.range('B11').options(index=False).value
    period = xml.Element("period")
    header.append(period)
    d3p1_begin = xml.SubElement(header, "d3p1:begin")
    d3p1_begin.text = Sheets1.range('C9').options(index=False).value
    d3p1_end = xml.SubElement(header, "d3p1:end")
    d3p1_end.text = Sheets1.range('C10').options(index=False).value
    partner = xml.Element("partner")
    header.append(partner)

    juridicalPerson = xml.Element("juridicalPerson")
    partner.append(juridicalPerson)
    d4p1_name = xml.SubElement(juridicalPerson, "d4p1:name")
    d4p1_name.text = Sheets1.range('B12').options(index=False).value
    d4p1_inn = xml.SubElement(juridicalPerson, "d4p1:inn")
    d4p1_inn.text = Sheets1.range('B13').options(index=False).value

    signerData = xml.Element("signerData")
    header.append(signerData)
    d3p1_employee = xml.Element("d3p1:employee")
    signerData.append(d3p1_employee)
    d3p1_first_name = xml.SubElement(d3p1_employee, "d3p1:first_name")
    d3p1_first_name.text = Sheets1.range('C14').options(index=False).value

    d3p1_post = xml.SubElement(d3p1_employee, "d3p1:post")
    d3p1_post.text = Sheets1.range('C15').options(index=False).value
    d3p1_basisAuthority = xml.SubElement(d3p1_employee, "d3p1:basisAuthority")
    d3p1_basisAuthority.text = Sheets1.range('C16').options(index=False).value
    contract = xml.Element("contract")
    header.append(contract)
    d3p1_type = xml.SubElement(contract, "d3p1:type")
    d3p1_type.text = Sheets1.range('C19').options(index=False).value
    d3p1_number = xml.SubElement(contract, "d3p1:number")
    d3p1_number.text = str(Sheets1.range('C20').options(index=False).value)
    d3p1_date = xml.SubElement(contract, "d3p1:date")
    d3p1_date.text = Sheets1.range('C21').options(index=False).value

    xml.indent(root, space=' ', level=0)

    tree = xml.ElementTree(root)

    with open("FD.xml", 'wb') as f:
        tree.write(f, encoding='utf-8')



if __name__ == "__main__":
    createXML("test.xml")

print(time.time() - start_time, "seconds")
