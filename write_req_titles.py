from color_func import retColors as rc
import xlrd
from lxml import etree as ET
from xml.dom import minidom as md
import re
import sys

book = xlrd.open_workbook("data\\EOS_testlink_Req.xls", formatting_info = True)
sheets = book.sheet_names()
ws = book.sheet_by_name(sheets[0])

# titles_buffer = ws.row(0)
# titles = []

# pat = re.compile('.*\'(.*)\'.*')

# for i in j:
#     mat = pat.match(str(i))

# req = titles[3]



# pat1 = re.compile(r"^(\d+)\s(.*)")
# pat2 = re.compile(r"^(\d+\.\d+)\s(.*)")
# pat3 = re.compile(r"^(\d+\.\d+\.\d+)\s(.*)")
# pat4 = re.compile(r"^(\d+\.\d+\.\d+\.\d+)\s(.*)")
pat = re.compile(r"^((\d+)|(\d+\.\d+)|(\d+\.\d+\.\d+)|(\d+\.\d+\.\d+\.\d+))\s(.*)")

root = ET.Element('requirement-specification')


for i in range(ws.nrows):
    temp = str(ws.cell(i, 3).value)
    # mat1 = pat1.match(str(temp))
    # mat2 = pat2.match(str(temp))
    # mat3 = pat3.match(str(temp))
    # mat4 = pat4.match(str(temp))
    mat = pat.match(temp)
    
    count1 = 1
    count2 = 1
    count3 = 1
    count4 = 1
    if mat:
        if mat.group(2):
            sub1 = ET.SubElement(root, "req_spec", title = str(mat.group(6)), doc_id = str(mat.group(2)))
            rev = ET.SubElement(sub1, "revision")
            rev.text = ET.CDATA(str(1))
            tipe = ET.SubElement(sub1, "type")
            tipe.text = ET.CDATA(str(1))
            node_order = ET.SubElement(sub1, "node_order")
            node_order.text = ET.CDATA(str(count1))
            count1 += 1
            total_req = ET.SubElement(sub1, "total_req")
            total_req.text = ET.CDATA(str(0))
            scope = ET.SubElement(sub1, "scope")
            xt_cell_cont = str(ws.cell(i+1, 3).value)
            mat_buff = pat.match(xt_cell_cont)
            if not mat_buff:
                scope.text = ET.CDATA(xt_cell_cont)
            else:
                scope.text = ET.CDATA(str(None))
                count2 = 0
        

        elif mat.group(3):
            # sub2 = ET.SubElement(sub1, "req_spec", title = str(mat2.group(2)), id = str(mat2.group(1)))
            # rev = ET.SubElement(sub2, "revision")
            # rev.text = ET.CDATA(str(1))
            # type = ET.SubElement(sub2, "type")
            # if str(ws.cell(i, 1).value) == "Information":
            #     type.text = ET.CDATA(str(1))
            # elif str(ws.cell(i, 1).value) == "Requirement":
            #     type.text = ET.CDATA(str(7))
            # else:
            #     type.text = ET.CDATA("")
            # node_order = ET.SubElement(sub2, "node_order")
            # node_order.text = ET.CDATA(str(count2))
            # count2 += 1
            # total_req = ET.SubElement(sub1, "total_req")
            # total_req.text = ET.CDATA(str(0))
            # cell_cont = str(ws.cell(i+1, 3).value)
            # if pat

            j = i+1
            cont_buff = ""
            while not pat.match(str(ws.cell(j, 3).value)):
                cont_buff += str(ws.cell(j, 3).value)
                j += 1

            # mat_buff = pat.match(str(ws.cell(j, 3).value))
            # if mat_buff.group(4):
            sub2 = ET.SubElement(sub1, "req_spec", title = str(mat.group(6)), doc_id = str(mat.group(3)))
            rev = ET.SubElement(sub2, "revision")
            rev.text = ET.CDATA(str(1))
            type = ET.SubElement(sub2, "type")
            type.text = ET.CDATA(str(1))
            node_order = ET.SubElement(sub2, "node_order")
            node_order.text = ET.CDATA(str(count2))
            count2 += 1
            total_req = ET.SubElement(sub2, "total_req")
            total_req.text = ET.CDATA(str(0))
            scope = ET.SubElement(sub2, "scope")
            scope.text = ET.CDATA(cont_buff)

            # else:
            #     sub2 = ET.SubElement(sub1, "requirement")
            #     docid = ET.SubElement(sub2, "docid")
            #     docid.text = ET.CDATA(str(mat.group(3)))
            #     title = ET.SubElement(sub2, "title")
            #     title.text = ET.CDATA(str(mat.group(6)))
            #     version = ET.SubElement(sub2, "version")
            #     version.text = str(1)
            #     revision = ET.SubElement(sub2, "revision")
            #     revision.text = str(1)
            #     node_order = ET.SubElement(sub2, "node_order")
            #     node_order.text = str(count2)
            #     count2 += 1
            #     description = ET.SubElement(sub2, "description")
            #     description.text = ET.CDATA(cont_buff)
            #     status = ET.SubElement(sub2, "status")
            #     status.text = ET.CDATA("D")
            #     tipe = ET.SubElement(sub2, "type")
            #     tipe.text = ET.CDATA(str(1))
            #     expected_coverage = ET.SubElement(sub2, "expected_coverage")
            #     expected_coverage.text = ET.CDATA(str(0))
            #     count3 = 0



        elif mat.group(4):
            # sub3 = ET.SubElement(sub2, "req_spec", title = str(mat3.group(2)), id = str(mat3.group(1)))
            # rev = ET.SubElement(sub3, "revision")
            # rev.text = ET.CDATA(str(1))
            # type = ET.SubElement(sub3, "type")
            # if str(ws.cell(i, 1).value) == "Information":
            #     type.text = ET.CDATA(str(1))
            # elif str(ws.cell(i, 1).value) == "Requirement":
            #     type.text = ET.CDATA(str(7))
            # else:
            #     type.text = ET.CDATA("")
            # node_order = ET.SubElement(sub3, "node_order")
            # node_order.text = ET.CDATA(str(count3))
            # count3 += 1
            # total_req = ET.SubElement(sub1, "total_req")
            # total_req.text = ET.CDATA(str(0))
            # scope = ET.SubElement(sub3, "scope")
            # if not(mat4):
            #     scope.text = ET.CDATA(str(ws.cell(i+1, 3).value))
            # else:
            #     scope.text = ET.CDATA(str(None))
            # count4 = 0

            j = i+1
            cont_buff = ""
            while not pat.match(str(ws.cell(j, 3).value)):
                cont_buff += str(ws.cell(j, 3).value)
                j += 1

            mat_buff = pat.match(str(ws.cell(j, 3).value))
            if mat_buff.group(5):
                sub3 = ET.SubElement(sub2, "req_spec", title = str(mat.group(6)), doc_id = str(mat.group(4)))
                rev = ET.SubElement(sub3, "revision")
                rev.text = ET.CDATA(str(1))
                type = ET.SubElement(sub3, "type")
                type.text = ET.CDATA(str(1))
                node_order = ET.SubElement(sub3, "node_order")
                node_order.text = ET.CDATA(str(count3))
                count3 += 1
                total_req = ET.SubElement(sub3, "total_req")
                total_req.text = ET.CDATA(str(0))
                scope = ET.SubElement(sub3, "scope")
                scope.text = ET.CDATA(cont_buff)

            else:
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
                node_order.text = str(count3)
                count3 += 1
                description = ET.SubElement(sub3, "description")
                description.text = ET.CDATA(cont_buff)
                status = ET.SubElement(sub3, "status")
                status.text = ET.CDATA("D")
                tipe = ET.SubElement(sub3, "type")
                tipe.text = ET.CDATA(str(1))
                expected_coverage = ET.SubElement(sub3, "expected_coverage")
                expected_coverage.text = ET.CDATA(str(0))
                count4 = 0


        elif mat.group(5):
            j = i+1
            cont_buff = ""
            while not pat.match(str(ws.cell(j, 3).value)):
                cont_buff += str(ws.cell(j, 3).value)
                j += 1

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
            node_order.text = str(count4)
            count4 += 1
            description = ET.SubElement(sub4, "description")
            description.text = ET.CDATA(cont_buff)
            status = ET.SubElement(sub4, "status")
            status.text = ET.CDATA("D")
            tipe = ET.SubElement(sub4, "type")
            tipe.text = ET.CDATA(str(1))
            expected_coverage = ET.SubElement(sub4, "expected_coverage")
            expected_coverage.text = ET.CDATA(str(0))
            

            


xmlstr = md.parseString(ET.tostring(root)).toprettyxml(indent="\t", encoding="UTF-8")

st = ET.tostring(root, encoding = "UTF-8", pretty_print=True)
#print(st.decode("UTF-8"))

tree = ET.ElementTree(root)
tree.write("./data/test_new_2.xml", pretty_print=True, xml_declaration=True, encoding="UTF-8")


import fileinput

with fileinput.FileInput("./data/test_new_1.xml", inplace=True, backup='.bak') as file:
    for line in file:
        print(line.replace("None", ""), end='')
