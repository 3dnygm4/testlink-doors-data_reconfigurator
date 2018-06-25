import xlrd
from lxml import etree as ET
import re

book = xlrd.open_workbook("data\\EOS_testlink_Req.xlsx")
sheets = book.sheet_names()
ws = book.sheet_by_name(sheets[0])

# print((ws.cell(143, 3).value))

pat = re.compile(r"^((\d+)|(\d+\.\d+)|(\d+\.\d+\.\d+)|(\d+\.\d+\.\d+\.\d+))\s(.*)")

root = ET.Element('requirement-specification')

for i in range(ws.nrows):
    temp = str(ws.cell(i, 3).value)
    mat = pat.match(temp)

    if mat:
        j = i+1
        cont_buff = ""
        cont_type = str(ws.cell(i+1, 1).value)
        while ws.cell(j,3).value and not pat.match(str(ws.cell(j, 3).value)):
        
            cont_buff += "[ACC Software Requirement Specification]\n\n"
            cont_buff += str(ws.cell(j, 3).value)
            cont_buff += "\n\n"
            if ws.cell(j, 4).value:
                cont_buff += "[OriginalD_BT]\n\n"
                cont_buff += str(ws.cell(j, 4).value)
                cont_buff += "\n\n"
            if ws.cell(j, 5).value:
                cont_buff += "[Interface]\n\n"
                cont_buff += str(ws.cell(j,5).value)
                cont_buff += "\n\n"
            if ws.cell(j, 6).value:
                cont_buff += "[InternalComments_BT]\n\n"
                cont_buff += str(ws.cell(j, 6).value)
                cont_buff += "\n\n"
            if ws.cell(j, 7).value:
                cont_buff += "[Comments]\n\n"
                cont_buff += str(ws.cell(j, 7).value)
                cont_buff += "\n\n"
            if ws.cell(j, 8).value:
                cont_buff += "[RBS]\n\n"
                cont_buff += str(ws.cell(j, 8).value)
                cont_buff += "\n\n"
            if ws.cell(j, 9).value:
                cont_buff += "[Acceptance Status (All modules)]\n\n"
                cont_buff += str(ws.cell(j,9).value)
                cont_buff += "\n\n"
            if ws.cell(j, 10).value:
                cont_buff += "[V&V Method]\n\n"
                cont_buff += str(ws.cell(j, 10).value)
                cont_buff += "\n\n"
            if ws.cell(j, 11).value:
                cont_buff += "[Application]\n\n"
                cont_buff += str(ws.cell(j, 11).value)
                cont_buff += "\n\n"
            if ws.cell(j, 12).value:
                cont_buff += "[PowerRange]\n\n"
                cont_buff += str(ws.cell(j, 12).value)
                cont_buff += "\n\n"

            j += 1
        
        # print(cont_buff)

        if  mat.group(2):
            
            sub1 = ET.SubElement(root, "req_spec", title = str(mat.group(6)), doc_id = str(mat.group(2)))
            rev = ET.SubElement(sub1, "revision")
            rev.text = ET.CDATA(str(1))
            tipe = ET.SubElement(sub1, "type")
            tipe.text = ET.CDATA(str(1))
            node_order = ET.SubElement(sub1, "node_order")
            node_order.text = ET.CDATA(str(1))
            total_req = ET.SubElement(sub1, "total_req")
            total_req.text = ET.CDATA(str(0))
            scope = ET.SubElement(sub1, "scope")
            if cont_buff:
                scope.text = ET.CDATA("<pre>" + cont_buff + "</pre>")
            else:
                scope.text = ET.CDATA(str(None))

        elif mat.group(3):
            sub2 = ET.SubElement(sub1, "req_spec", title = str(mat.group(6)), doc_id = str(mat.group(3)))
            rev = ET.SubElement(sub2, "revision")
            rev.text = ET.CDATA(str(1))
            tipe = ET.SubElement(sub2, "type")
            tipe.text = ET.CDATA(str(1))
            node_order = ET.SubElement(sub2, "node_order")
            node_order.text = ET.CDATA(str(1))
            total_req = ET.SubElement(sub2, "total_req")
            total_req.text = ET.CDATA(str(0))
            scope = ET.SubElement(sub2, "scope")
            if cont_buff:
                scope.text = ET.CDATA("<pre>" + cont_buff + "</pre>")
            else:
                scope.text = ET.CDATA(str(None))

        elif mat.group(4):
            if cont_type == "Information":
                sub3 = ET.SubElement(sub2, "req_spec", title = str(mat.group(6)), doc_id = str(mat.group(4)))
                rev = ET.SubElement(sub3, "revision")
                rev.text = ET.CDATA(str(1))
                tipe = ET.SubElement(sub3, "type")
                tipe.text = ET.CDATA(str(1))
                node_order = ET.SubElement(sub3, "node_order")
                node_order.text = ET.CDATA(str(1))
                total_req = ET.SubElement(sub3, "total_req")
                total_req.text = ET.CDATA(str(0))
                scope = ET.SubElement(sub3, "scope")
                if cont_buff:
                    scope.text = ET.CDATA("<pre>" + cont_buff + "</pre>")
                else:
                    scope.text = ET.CDATA(str(None))
            elif cont_type == "Requirement":
                sub3 = ET.SubElement(sub2, "requirement")
                docid = ET.SubElement(sub3, "docid")
                docid.text = ET.CDATA(str(mat.group(4)))
                title = ET.SubElement(sub3, "title")
                title.text = ET.CDATA(str(mat.group(6)))
                version = ET.SubElement(sub3, "version")
                version.text = str(1)
                revision = ET.SubElement(sub3, "revision")
                revision.text = str(1)
                node_order = ET.SubElement(sub3, "node_order")
                node_order.text = str(1)
                description = ET.SubElement(sub3, "description")
                description.text = ET.CDATA("<pre>" + cont_buff + "</pre>")
                status = ET.SubElement(sub3, "status")
                status.text = ET.CDATA("D")
                tipe = ET.SubElement(sub3, "type")
                tipe.text = ET.CDATA(str(1))
                expected_coverage = ET.SubElement(sub3, "expected_coverage")
                expected_coverage.text = ET.CDATA(str(0))


        elif mat.group(5):
            if int(mat.group(5).split('.')[0]) < 3:
                sub4 = ET.SubElement(sub3, "req_spec", title = str(mat.group(6)), doc_id = str(mat.group(2)))
                rev = ET.SubElement(sub4, "revision")
                rev.text = ET.CDATA(str(1))
                tipe = ET.SubElement(sub4, "type")
                tipe.text = ET.CDATA(str(1))
                node_order = ET.SubElement(sub4, "node_order")
                node_order.text = ET.CDATA(str(1))
                total_req = ET.SubElement(sub4, "total_req")
                total_req.text = ET.CDATA(str(0))
                scope = ET.SubElement(sub4, "scope")
                if cont_buff:
                    scope.text = ET.CDATA("<pre>" + cont_buff + "</pre>")
                else:
                    scope.text = ET.CDATA(str(None))
            
            else:

                sub4 = ET.SubElement(sub3, "requirement")
                docid = ET.SubElement(sub4, "docid")
                docid.text = ET.CDATA(str(mat.group(5)))
                title = ET.SubElement(sub4, "title")
                title.text = ET.CDATA(str(mat.group(6)))
                version = ET.SubElement(sub4, "version")
                version.text = str(1)
                revision = ET.SubElement(sub4, "revision")
                revision.text = str(1)
                node_order = ET.SubElement(sub4, "node_order")
                node_order.text = str(1)
                description = ET.SubElement(sub4, "description")
                description.text = ET.CDATA("<pre>" + cont_buff + "</pre>")
                status = ET.SubElement(sub4, "status")
                status.text = ET.CDATA("D")
                tipe = ET.SubElement(sub4, "type")
                tipe.text = ET.CDATA(str(1))
                expected_coverage = ET.SubElement(sub4, "expected_coverage")
                expected_coverage.text = ET.CDATA(str(0))

tree = ET.ElementTree(root)
tree.write("./data/test_new_3.xml", pretty_print=True, xml_declaration=True, encoding="UTF-8")


