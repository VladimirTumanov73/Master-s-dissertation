import xlwings as xw
# import pandas as pd
import xml.etree.ElementTree as xml
import re
import os
import time

# FD_dict = dict.fromkeys(['number', 'date', 'd3p1:begin', 'd3p1:end',
#                          'phone', 'd4p1:name', 'd4p1:inn', 'd4p1:ogrn', 'd4p1:address',
#                         'd3p1:first_name', 'd3p1:last_name', 'd3p1:patronimic_name', 'd3p1:post', 'd3p1:basisAuthority',
#                          'd3p1:type', 'd3p1:number', 'd3p1:date', 'd3p1:registrationNumber',
#                          ])
start_time = time.time()
wb = xw.Book('FD3.xlsx')
Sheets1 = wb.sheets['титульник']
# FD_dict['number'] = Sheets1.range('AE6').options(index=False).value
# FD_dict['date'] = Sheets1.range('N7').options(index=False).value
# FD_dict['d3p1:begin'] = Sheets1.range('I22').options(index=False).value
# FD_dict['d3p1:end'] = Sheets1.range('Z22').options(index=False).value
# FD_dict['phone'] = Sheets1.range('AT12').options(index=False).value
# FD_dict['d4p1:name'] = Sheets1.range('A11').options(index=False).value
# FD_dict['d4p1:inn'] = Sheets1.range('U12').options(index=False).value
# FD_dict['d4p1:ogrn'] = Sheets1.range('AF12').options(index=False).value
# FD_dict['d4p1:address'] = Sheets1.range('A12').options(index=False).value
# FD_dict['d3p1:first_name'] = Sheets1.range('J19').options(index=False).value
# FD_dict['d3p1:last_name'] = Sheets1.range('A19').options(index=False).value
# FD_dict['d3p1:patronimic_name'] = Sheets1.range('AE19').options(index=False).value
# FD_dict['d3p1:post'] = Sheets1.range('O17').options(index=False).value
# FD_dict['d3p1:basisAuthority'] = Sheets1.range('R21').options(index=False).value
# FD_dict['d3p1:type'] = Sheets1.range('AB14').options(index=False).value
# FD_dict['d3p1:number'] = Sheets1.range('R16').options(index=False).value
# FD_dict['d3p1:date'] = Sheets1.range('C16').options(index=False).value
# FD_dict['d3p1:registrationNumber'] = Sheets1.range('A17').options(index=False).value

root = xml.Element("forestDeclaration")
number = xml.SubElement(root, "number")
number.text = Sheets1.range('AE6').options(index=False).value
date = xml.SubElement(root, "date")
date.text = Sheets1.range('N7').options(index=False).value

header = xml.Element("header")
root.append(header)
subject = xml.SubElement(header, "subject")
subject.text = Sheets1.range('A9').options(index=False).value
ex_Aut = xml.SubElement(header, "executiveAuthority")
ex_Aut.text = Sheets1.range('AA9').options(index=False).value
period = xml.Element("period")
header.append(period)
d3p1_begin = xml.SubElement(header, "d3p1:begin")
d3p1_begin.text = Sheets1.range('I23').options(index=False).value
d3p1_end = xml.SubElement(header, "d3p1:end")
d3p1_end.text = Sheets1.range('Z23').options(index=False).value
partner = xml.Element("partner")
header.append(partner)
phone = xml.SubElement(partner, "phone")
phone.text = Sheets1.range('AT12').options(index=False).value
juridicalPerson = xml.Element("juridicalPerson")
partner.append(juridicalPerson)
d4p1_name = xml.SubElement(juridicalPerson, "d4p1:name")
d4p1_name.text = Sheets1.range('A11').options(index=False).value
d4p1_inn = xml.SubElement(juridicalPerson, "d4p1:inn")
d4p1_inn.text = Sheets1.range('U12').options(index=False).value
d4p1_ogrn = xml.SubElement(juridicalPerson, "d4p1:ogrn")
d4p1_ogrn.text = Sheets1.range('AF12').options(index=False).value
d4p1_address = xml.SubElement(juridicalPerson, "d4p1:address")
d4p1_address.text = Sheets1.range('A12').options(index=False).value
signerData = xml.Element("signerData")
header.append(signerData)
d3p1_employee = xml.Element("d3p1:employee")
signerData.append(d3p1_employee)
d3p1_first_name = xml.SubElement(d3p1_employee, "d3p1:first_name")
d3p1_first_name.text = Sheets1.range('J19').options(index=False).value
d3p1_last_name = xml.SubElement(d3p1_employee, "d3p1:last_name")
d3p1_last_name.text = Sheets1.range('A19').options(index=False).value
d3p1_pat_name = xml.SubElement(d3p1_employee, "d3p1:patronimic_name")
d3p1_pat_name.text = Sheets1.range('AE19').options(index=False).value
d3p1_post = xml.SubElement(d3p1_employee, "d3p1:post")
d3p1_post.text = Sheets1.range('O17').options(index=False).value
d3p1_basisAuthority = xml.SubElement(d3p1_employee, "d3p1:basisAuthority")
d3p1_basisAuthority.text = Sheets1.range('R21').options(index=False).value
contract = xml.Element("contract")
header.append(contract)
d3p1_type = xml.SubElement(contract, "d3p1:type")
d3p1_type.text = Sheets1.range('AB14').options(index=False).value
d3p1_number = xml.SubElement(contract, "d3p1:number")
d3p1_number.text = Sheets1.range('R16').options(index=False).value
d3p1_date = xml.SubElement(contract, "d3p1:date")
d3p1_date.text = Sheets1.range('C16').options(index=False).value
d3p1_RN = xml.SubElement(contract, "d3p1:registrationNumber")
d3p1_RN.text = Sheets1.range('A17').options(index=False).value

dirname = "/FDapp"
files = os.listdir(dirname)


print(files)


# tree = xml.ElementTree(root)
# xml.dump(tree)

data = xml.tostring(root, encoding='utf8', method='xml')

# new_text = re.sub(r'(</.+?>)', r'\1\n', root)
# data = xml.fromstring(root, encoding='utf8', method='xml')
print(data)

# groups = itertools.groupby(xml.fromstring(root), lambda element: element.tag)
# groups = [list(group[1]) for group in groups]

file = open("FD.xml",  "w", encoding='utf8')
file.write(data.decode('utf-8'))
print(time.time() - start_time, "seconds")
