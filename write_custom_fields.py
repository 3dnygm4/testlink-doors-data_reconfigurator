from color_func import retColors
import xlrd
from lxml import etree as ET
from xml.dom import minidom as minid

book = xlrd.open_workbook("data\\EOS_testlink_Req.xls", formatting_info = True)
sheets = book.sheet_names()
ws = book.sheet_by_name(sheets[0])

colors = retColors(ws)

root = ET.Element("custom_fields")

fields = ["name", "label", "type", "possible_values", "default_value", "valid_regex", "length_min", "length_max", "show_on_design", "enable_on_design", "show_on_execution", "enable_on_execution", "show_on_testplan", "enable_on_testplan", "node_type_id"]

for i in range(ws.ncols):
    cf = ET.SubElement(root, "custom_field")
    for j in range(len(fields)):
        sub = ET.SubElement(cf, fields[j])
        if fields[j] == "name" or fields[j] == "label":
            cell_content = ws.cell(0, i).value
            sub.text = ET.CDATA(str((cell_content)))
        elif j == 2:
            sub.text =ET.CDATA(str((0)))
        elif j in range(3,6):
            sub.text = ET.CDATA(str(()))
        elif j in range(6, 8):            
            sub.text = ET.CDATA(str((0)))
        elif j in range(8, 10):
            sub.text = ET.CDATA(str((1)))
        elif j in range(10, 14):
            sub.text = ET.CDATA(str((0)))
        else:
           sub.text = ET.CDATA(str((7)))



xmlstr = minid.parseString(ET.tostring(root)).toprettyxml(indent="   ", encoding='UTF-8')
with open('./data/test_customFields.xml', 'w') as f:
    f.write(str(xmlstr.decode('UTF-8')))
        





