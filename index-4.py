

import os
from os import path
import shutil
import xlrd     # module to deal excel
# from xml.dom.minidom import parse,parseString

import xml.etree.ElementTree as ET



from tkinter import *
from tkinter import messagebox
tk = Tk()
tk.withdraw()

vinNumber = 0
xmlNumbers = 0

rootdir = './'
condition_xml_directory = 'condition_xml'
result_xml_directory = 'xml'

list = os.listdir(os.getcwd())


for line in list:
    filepath = os.path.join(rootdir, line)
    if os.path.isfile(filepath):
        if '.xlsx' in filepath:
            root_vin_excel_path = filepath
        if '.xml' in filepath:
            root_xml_path = filepath
# print(root_xml_path,root_vin_excel_path)

wb = xlrd.open_workbook(root_vin_excel_path)
ws = wb.sheet_by_name("Missing_Data_VIN")
allVinUpperList = []
allVinLowerList = []
for index in range(len(ws.col_values(0))):
    if index !=0:
        allVinUpperList.append(str(ws.col(0)[index])[6:-1])
        allVinLowerList.append(str.lower(str(ws.col(0)[index])[6:-1]))
vinNumber = len(allVinUpperList)     # 在excel中vincode总个数


shutil.copyfile('Template.xml','Temp.xml')

tree = ET.parse('Temp.xml')
root = tree.getroot()
os.chdir(condition_xml_directory)     #进入condition_xml

for index in range(len(allVinLowerList)):
    for node in root.iter('GeneralInfo'):
        for subNode in node.iter('VehicleIdentNumber'):
            subNode.text = allVinUpperList[index]

    condition_xml = 'vin_' + str(allVinLowerList[index]) + '_coding.xml'
    c_tree = ET.parse(condition_xml)
    c_root = c_tree.getroot()

    for node in root.iter('Feature'):
        # print(node.attrib['name'])             # ACM,BCM,CCU,接下来确认Template.xml中该feature下有cs,ci,es
        for subNode in node:
            for subsubNode in subNode:
                if (subsubNode.attrib['name'] == 'ES'):      #修改模板中的ES值
                    for c_node in c_root.iter('vehicle'):
                        for c_subNode in c_node.iter('esk'):
                            # print(c_subNode.text)
                            # print(len(c_subNode.text))
                            format_text = ''
                            for i in range(len(c_subNode.text)):
                                if i > 0 and i %2 == 0:
                                    format_text = format_text + ' ' + c_subNode.text[i:i+1]
                                else:
                                    format_text = format_text + c_subNode.text[i:i+1]
                            # subsubNode.text = c_subNode.text
                            subsubNode.text = format_text


                if(subsubNode.attrib['name'] == 'CS' or 'CI'):
                    for c_node in c_root.iter('ecus'):
                        for c_subNode in c_node.iter('ecu'):
                            # print(c_subNode.attrib['shortName'])
                            if (c_subNode.attrib['shortName'] == node.attrib['name']):
                                for c_subsubNode in c_subNode.iter('coding'):
                                    for c_item in c_subsubNode:
                                        # print(c_item.tag)
                                        if(subsubNode.attrib['name'] == 'CS' and c_item.tag == 'cs_data'):   #修改模板中的CS值
                                            # print(c_item.text)
                                            subsubNode.text = c_item.text
                                        if(subsubNode.attrib['name'] == 'CI' and c_item.tag == 'ci_data'):    #修改模板中的CI值
                                            # print(c_item.text)
                                            subsubNode.text = c_item.text

    os.chdir(os.pardir)    #回到根目录并准备进入xml目录

    target_xml_name = str(allVinUpperList[index]) + '.xml'
    if not os.path.exists(result_xml_directory):   # 生成xml的文件夹,位置
        os.mkdir(result_xml_directory)
        os.chdir(result_xml_directory)
    else:
        os.chdir(result_xml_directory)

    tree.write(target_xml_name)
    with open(target_xml_name, 'r+') as f:
        content = f.read()
        f.seek(0, 0)
        f.write('<?xml version="1.0" encoding="UTF-8" ?>\n' + content)

    xmlNumbers = xmlNumbers + 1
    os.chdir(os.pardir)  # 回到根目录并进入condition_xml目录
    shutil.copyfile('Template.xml', 'Temp.xml')
    tree = ET.parse('Temp.xml')
    root = tree.getroot()
    os.chdir(condition_xml_directory)


os.chdir(os.pardir)                 #回到根目录准备删除临时文件Temp.xml
try:
    os.remove('Temp.xml')
except OSError as e:
    print(e)


if(vinNumber == xmlNumbers):      # tell user the success result
    resultMessage = 'Excel中共有 ' + str(vinNumber) + ' 个VIN码,成功转换生成 ' + str(xmlNumbers) + " 个XML文件!!"
else:   # tell user the failure resul
    resultMessage = 'Excel中共有 ' + str(vinNumber) + ' 个VIN码,成功转换生成 ' + str(xmlNumbers) + " 个XML文件!!" + '缺失了 ' + str(vinNumber-xmlNumbers) +' 个vin_***_coding.xml文件'
txt = messagebox.showinfo("Result",resultMessage)


if txt == "ok":
    tk.destroy()
    tk.mainloop()