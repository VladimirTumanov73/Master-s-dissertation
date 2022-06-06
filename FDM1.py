import xml.etree.ElementTree as xml
import xlwings as xw
import time
import os
# import threading
# import multiprocessing
# import shutil
# import zipfile


def Creat_xml(filename):
    start_time = time.time()
    wb = xw.Book(filename)
    Sheets1 = wb.sheets[0]
    Sheets2 = wb.sheets[1]
    Sheets3 = wb.sheets[2]
    Sheets4 = wb.sheets[3]
    Sheets5 = wb.sheets[4]

    full_name = os.path.basename(filename)
    name = os.path.splitext(full_name)[0]
    namexml = name+".xml"
    control = []
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
    control.append(str(Sheets1.range('C9').options(index=False).value))
    d3p1_end = xml.SubElement(header, "d3p1:end")
    d3p1_end.text = str(Sheets1.range('C10').options(index=False).value)
    control.append(str(Sheets1.range('C10').options(index=False).value))
    partner = xml.Element("partner")
    header.append(partner)
    juridicalPerson = xml.Element("juridicalPerson")
    partner.append(juridicalPerson)
    d4p1_name = xml.SubElement(juridicalPerson, "d4p1:name")
    d4p1_name.text = str(Sheets1.range('B12').options(index=False).value)
    control.append(str(Sheets1.range('B12').options(index=False).value))
    d4p1_inn = xml.SubElement(juridicalPerson, "d4p1:inn")
    d4p1_inn.text = str(Sheets1.range('B13').options(index=False).value)
    control.append(str(Sheets1.range('B13').options(index=False).value))
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
    control.append(str(Sheets1.range('C19').options(index=False).value))
    d3p1_number = xml.SubElement(contract, "d3p1:number")
    d3p1_number.text = str(Sheets1.range('C20').options(index=False).value)
    control.append(str(Sheets1.range('C20').options(index=False).value))
    d3p1_date = xml.SubElement(contract, "d3p1:date")
    d3p1_date.text = str(Sheets1.range('C21').options(index=False).value)
    control.append(str(Sheets1.range('C21').options(index=False).value))

    # HarvestingWood
    hW = xml.Element("harvestingWood")
    root.append(hW)
    row_hW = xml.Element("row")
    hW.append(row_hW)
    specialPurpose = xml.SubElement(row_hW, "specialPurpose")
    specialPurpose.text = " "
    protectionCategory = xml.SubElement(row_hW, "protectionCategory")
    protectionCategory.text = " "
    locationHW = xml.Element("location")
    row_hW.append(locationHW)
    forestryHW = xml.SubElement(locationHW, "forestry")
    forestryHW.text = str(Sheets2.range('I12').options(index=False).value)
    subforestryHW = xml.SubElement(locationHW, "subforestry")
    subforestryHW.text = str(Sheets2.range('W12').options(index=False).value)
    quarterHW = xml.SubElement(locationHW, "quarter")
    quarterHW.text = str(Sheets2.range('AW12').options(index=False).value)
    taxationUnitHW = xml.SubElement(locationHW, "taxationUnit")
    taxationUnitHW.text = str(Sheets2.range('BD12').options(index=False).value)
    cuttingAreaHW = xml.SubElement(locationHW, "cuttingArea")
    cuttingAreaHW.text = str(Sheets2.range('BK12').options(index=False).value)
    areaHW = xml.SubElement(row_hW, "area")
    areaHW.text = str(Sheets2.range('BK12').options(index=False).value)
    formCuttingHW = xml.SubElement(row_hW, "formCutting")
    formCuttingHW.text = str(Sheets2.range('BU12').options(index=False).value)
    typeCuttingHW = xml.SubElement(row_hW, "typeCutting")
    typeCuttingHW.text = str(Sheets2.range('CB12').options(index=False).value)
    farmHW = xml.SubElement(row_hW, "farm")
    farmHW.text = str(Sheets2.range('CI12').options(index=False).value)
    treeHW = xml.SubElement(row_hW, "tree")
    treeHW.text = str(Sheets2.range('CP12').options(index=False).value)
    unitTypeHW = xml.SubElement(row_hW, "unitType")
    unitTypeHW.text = str(Sheets2.range('CW12').options(index=False).value)
    volumeHW = xml.SubElement(row_hW, "volume")
    volumeHW.text = str(Sheets2.range('DD12').options(index=False).value)
    usageTypeHW = xml.SubElement(hW, "usageType")
    usageTypeHW.text = str(Sheets2.range('AP6').options(index=False).value)

    # HarvestingObject
    hO = xml.Element("harvestingObject")
    root.append(hO)
    row_hO = xml.Element("row")
    hO.append(row_hO)
    object = xml.SubElement(row_hO, "object")
    object.text = str(Sheets3.range('I7').options(index=False).value)
    objectNumber = xml.SubElement(row_hO, "objectNumber")
    objectNumber.text = str(Sheets3.range('A7').options(index=False).value)
    measure = xml.SubElement(row_hO, "measure")
    measure.text = str(Sheets3.range('S7').options(index=False).value)
    locationHO = xml.Element("location")
    row_hO.append(locationHO)
    forestryHO = xml.SubElement(locationHO, "forestry")
    forestryHO.text = str(Sheets3.range('AC7').options(index=False).value)
    subforestryHO = xml.SubElement(locationHO, "subforestry")
    subforestryHO.text = str(Sheets3.range('AM7').options(index=False).value)
    quarterHO = xml.SubElement(locationHO, "quarter")
    quarterHO.text = str(Sheets3.range('BG7').options(index=False).value)
    taxationUnitHO = xml.SubElement(locationHO, "taxationUnit")
    taxationUnitHO.text = str(Sheets3.range('BN7').options(index=False).value)
    areaHO = xml.SubElement(row_hO, "area")
    areaHO.text = str(Sheets3.range('BU7').options(index=False).value)
    formCuttingHO = xml.SubElement(row_hO, "formCutting")
    formCuttingHO.text = str(Sheets3.range('CH7').options(index=False).value)
    typeCuttingHO = xml.SubElement(row_hO, "typeCutting")
    typeCuttingHO.text = str(Sheets3.range('CM7').options(index=False).value)
    farmHO = xml.SubElement(row_hO, "farm")
    farmHO.text = str(Sheets3.range('CS7').options(index=False).value)
    treeHO = xml.SubElement(row_hO, "tree")
    treeHO.text = str(Sheets3.range('CY7').options(index=False).value)
    unitTypeHO = xml.SubElement(row_hO, "unitType")
    unitTypeHO.text = "куб.м"
    volumeHO = xml.SubElement(row_hO, "volume")
    volumeHO.text = str(Sheets3.range('DE7').options(index=False).value)

    # OtherUsageTypes
    oT = xml.Element("OtherUsageTypes")
    root.append(oT)
    row_oT = xml.Element("row")
    oT.append(row_oT)
    specialPurpose = xml.SubElement(row_oT, "specialPurpose")
    specialPurpose.text = " "
    locationoT = xml.Element("location")
    row_oT.append(locationoT)
    forestryoT = xml.SubElement(locationoT, "forestry")
    forestryoT.text = str(Sheets4.range('A13').options(index=False).value)
    subforestryoT = xml.SubElement(locationoT, "subforestry")
    subforestryoT.text = str(Sheets4.range('K13').options(index=False).value)
    quarteroT = xml.SubElement(locationoT, "quarter")
    quarteroT.text = str(Sheets4.range('AG13').options(index=False).value)
    taxationUnitoT = xml.SubElement(locationoT, "taxationUnit")
    taxationUnitoT.text = str(Sheets4.range('AN13').options(index=False).value)
    areaoT = xml.SubElement(row_oT, "area")
    areaoT.text = str(Sheets4.range('AU13').options(index=False).value)
    resource = xml.SubElement(row_oT, "resource")
    resource.text = str(Sheets4.range('BE13').options(index=False).value)
    unitType = xml.SubElement(row_oT, "unitType")
    unitType.text = str(Sheets4.range('BN13').options(index=False).value)
    volumeoT1 = xml.SubElement(row_oT, "volume")
    volumeoT1.text = str(Sheets4.range('BU13').options(index=False).value)
    formCuttingoT = xml.SubElement(row_oT, "formCutting")
    formCuttingoT.text = str(Sheets4.range('CH13').options(index=False).value)
    typeCuttingoT = xml.SubElement(row_oT, "typeCutting")
    typeCuttingoT.text = str(Sheets4.range('CM13').options(index=False).value)
    treeoT = xml.SubElement(row_oT, "tree")
    treeoT.text = str(Sheets4.range('CY13').options(index=False).value)
    volumeoT2 = xml.SubElement(row_oT, "volume")
    volumeoT2.text = str(Sheets4.range('DE13').options(index=False).value)
    usageType_oT = xml.SubElement(oT, "usageType")
    usageType_oT.text = str(Sheets4.range('AN6').options(index=False).value)

    # OtherUsageObjects
    OUO = xml.Element("otherUsageObjects")
    root.append(OUO)
    row_OUO = xml.Element("row")
    OUO.append(row_OUO)
    objectOUO = xml.SubElement(row_OUO, "object")
    objectOUO.text = str(Sheets5.range('I8').options(index=False).value)
    objectNumberOUO = xml.SubElement(row_OUO, "objectNumber")
    objectNumberOUO.text = str(Sheets5.range('A8').options(index=False).value)
    measureOUO = xml.SubElement(row_OUO, "measure")
    measureOUO.text = str(Sheets5.range('S8').options(index=False).value)
    locationOUO = xml.Element("location")
    row_OUO.append(locationOUO)
    forestryOUO = xml.SubElement(locationOUO, "forestry")
    forestryOUO.text = str(Sheets5.range('AC8').options(index=False).value)
    subforestryOUO = xml.SubElement(locationOUO, "subforestry")
    subforestryOUO.text = str(Sheets5.range('AM8').options(index=False).value)
    quarterOUO = xml.SubElement(locationOUO, "quarter")
    quarterOUO.text = str(Sheets5.range('BG8').options(index=False).value)
    taxationUnitOUO = xml.SubElement(locationOUO, "taxationUnit")
    taxationUnitOUO.text = str(Sheets5.range('BN8').options(index=False).value)
    unitTypeOUO = xml.SubElement(row_OUO, "unitType")
    unitTypeOUO.text = "га"
    volumeOUO = xml.SubElement(row_OUO, "volume")
    volumeOUO.text = str(Sheets5.range('BU8').options(index=False).value)
    areaOUO = xml.SubElement(row_OUO, "area")
    areaOUO.text = str(Sheets5.range('CB8').options(index=False).value)
    formCuttingOUO = xml.SubElement(row_OUO, "formCutting")
    formCuttingOUO.text = str(Sheets5.range('CH8').options(index=False).value)
    typeCuttingOUO = xml.SubElement(row_OUO, "typeCutting")
    typeCuttingOUO.text = str(Sheets5.range('CM8').options(index=False).value)
    treeOUO = xml.SubElement(row_OUO, "tree")
    treeOUO.text = str(Sheets5.range('CY8').options(index=False).value)
    treeVolumeOUO = xml.SubElement(row_OUO, "treeVolume")
    treeVolumeOUO.text = str(Sheets5.range('DE8').options(index=False).value)

    # Output
    xml.indent(root, space=' ', level=0)

    tree = xml.ElementTree(root)

    with open(namexml, 'wb') as f:
        tree.write(f, encoding='utf-8', xml_declaration=True)

    print(time.time() - start_time, "seconds")

    return control
    # src = os.path.realpath(namexml)
    # root_dir, tail = os.path.split(src)
    # shutil.make_archive("test", 'zip', "files")
    # # print(os.path.dirname(namexml))
    #
    # with zipfile.ZipFile("test.zip", "a") as newzip:
    #     newzip.write(namexml)

if __name__ == "__main__":
    file_name = "FDTest1.xlsx"
    # print("Default")
    Creat_xml(file_name)

    # for f in range(1, 6):
    #     print("Flow {}".format(f))
    #     flow = threading.Thread(target=module1(file_name), args=(f, ''))
    #     flow.start()
    #
    # processes = []
    # for i in range(1, 7):
    #     print("Proces {}".format(i))
    #     p = multiprocessing.Process(target=module1(file_name), args=[i])
    #     p.start()
    #     processes.append(p)
    #
    # for p in processes:
    #     p.join()
