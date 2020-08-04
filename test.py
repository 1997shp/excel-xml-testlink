# -*- coding: utf-8 -*-
from xml.dom.minidom import parse
import xlrd
import xml.dom.minidom

if __name__ == '__main__':

    wb = xlrd.open_workbook('TestCase1.xls', formatting_info=True)
    table = wb.sheet_by_name('Sheet1')
    nrows = table.nrows  # 获取总行数
    i = 1
    n = 1
    impl = xml.dom.minidom.getDOMImplementation()
    domTree = impl.createDocument(None, 'testsuite', None)
    while(i < nrows):

        rows = table.row_values(i)
        n = i + 1
        if (n < nrows):
            rows1 = table.row_values(n)
            if (len(rows1[0]) == 0):
                while (len(rows1[0]) == 0):
                    n = n + 1
                    rows = rows + rows1

                    rows1 = table.row_values(n)

                # print(len(rows))
                # print(' ')
                # domTree = dom.documentElement
                rootNode = domTree.documentElement
                # 新建一个customer节点
                customer_node = domTree.createElement("testsuite")
                customer_node.setAttribute("name", rows[0])


                customer_node1 = domTree.createElement("testcase")
                customer_node1.setAttribute("name", rows[1])
                customer_node.appendChild(customer_node1)


                comments_node = domTree.createElement("summary")
                cdata_text_value = domTree.createCDATASection(rows[2])
                comments_node.appendChild(cdata_text_value)
                customer_node1.appendChild(comments_node)

                comments_node = domTree.createElement("preconditions")
                cdata_text_value = domTree.createCDATASection(rows[3])
                comments_node.appendChild(cdata_text_value)
                customer_node1.appendChild(comments_node)

                comments_node = domTree.createElement("execution_type")
                cdata_text_value = domTree.createCDATASection(str(int(rows[4])))
                comments_node.appendChild(cdata_text_value)
                customer_node1.appendChild(comments_node)

                comments_node = domTree.createElement("importance")
                cdata_text_value = domTree.createCDATASection(str(int(rows[5])))
                comments_node.appendChild(cdata_text_value)
                customer_node1.appendChild(comments_node)

                customer_node2 = domTree.createElement("steps")
                # customer_node2.setAttribute("name", "")
                customer_node1.appendChild(customer_node2)

                q = 0   # 行数
                w = len(rows)
                e = w / 8   #行数
                r = 6   # 操作步骤内容
                t = 7   # 期望结果内容

                while (q < e):
                    q = q + 1
                    customer_node3 = domTree.createElement("step")
                    # customer_node2.setAttribute("name", "")
                    customer_node2.appendChild(customer_node3)

                    comments_node = domTree.createElement("step_number")
                    cdata_text_value = domTree.createCDATASection(str(q))
                    comments_node.appendChild(cdata_text_value)
                    customer_node3.appendChild(comments_node)

                    comments_node = domTree.createElement("actions")
                    cdata_text_value = domTree.createCDATASection(rows[r])
                    comments_node.appendChild(cdata_text_value)
                    customer_node3.appendChild(comments_node)
                    # print(rows[t])
                    comments_node = domTree.createElement("expectedresults")
                    cdata_text_value = domTree.createCDATASection(rows[t])
                    comments_node.appendChild(cdata_text_value)
                    customer_node3.appendChild(comments_node)


                    r = r + 8
                    t = t + 8
                rootNode.appendChild(customer_node)
                i = n

            else:
                # print(rows)
                # print(' ')

                # domTree = dom.documentElement
                rootNode = domTree.documentElement
                # 新建一个customer节点
                customer_node = domTree.createElement("testsuite")  # 用例集名称
                customer_node.setAttribute("name", rows[0])


                customer_node1 = domTree.createElement("testcase")  #用例名称
                customer_node1.setAttribute("name", rows[1])
                customer_node.appendChild(customer_node1)

                comments_node = domTree.createElement("summary")    #用例摘要
                cdata_text_value = domTree.createCDATASection(rows[2])
                comments_node.appendChild(cdata_text_value)
                customer_node1.appendChild(comments_node)

                comments_node = domTree.createElement("preconditions")  #前提条件
                cdata_text_value = domTree.createCDATASection(rows[3])
                comments_node.appendChild(cdata_text_value)
                customer_node1.appendChild(comments_node)

                comments_node = domTree.createElement("execution_type") #执行方式  0代表手工  2代表自动的
                cdata_text_value = domTree.createCDATASection(str(int(rows[4])))
                comments_node.appendChild(cdata_text_value)
                customer_node1.appendChild(comments_node)

                comments_node = domTree.createElement("importance")     #重要性 0代表高 1代表低 2代表中
                cdata_text_value = domTree.createCDATASection(str(int(rows[5])))
                comments_node.appendChild(cdata_text_value)
                customer_node1.appendChild(comments_node)

                customer_node2 = domTree.createElement("steps")     # 操作步骤和期望结果集
                # customer_node2.setAttribute("name", "")
                customer_node1.appendChild(customer_node2)

                customer_node3 = domTree.createElement("step")
                # customer_node2.setAttribute("name", "")
                customer_node2.appendChild(customer_node3)

                comments_node = domTree.createElement("step_number")    # 操作步骤数
                cdata_text_value = domTree.createCDATASection("1")
                comments_node.appendChild(cdata_text_value)
                customer_node3.appendChild(comments_node)

                comments_node = domTree.createElement("actions")    # 操作步骤
                cdata_text_value = domTree.createCDATASection(rows[6])
                comments_node.appendChild(cdata_text_value)
                customer_node3.appendChild(comments_node)

                comments_node = domTree.createElement("expectedresults")    #期望结果
                cdata_text_value = domTree.createCDATASection(rows[7])
                comments_node.appendChild(cdata_text_value)
                customer_node3.appendChild(comments_node)

                rootNode.appendChild(customer_node)
                i = i + 1

        else:
            break

    with open('added_customer.xml', 'w') as f:
        domTree.writexml(f, addindent=' ', encoding='gb2312')

    print("写入成功 文件名为added_customer.xml")


