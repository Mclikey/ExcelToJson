#!/usr/bin/python3
### create by tishoy
### 2023-12-7
### Copyright (c) 2023-2033 Banding 
import sys
import os
import openpyxl 
import json


import re
 
def is_chinese(text):
    pattern = r'[\u4e00-\u9fa5]+' # 定义中文范围的Unicode编码区间
    result = re.findall(pattern, text) # 查找符合条件的内容
    
    if len(result) > 0:
        return True
    else:
        return False

if (len(sys.argv) == 1):
    ExcelPath = "./"
    JsonPath = "./"
elif (len(sys.argv)  > 1):
    print("第一个传入的参数为:", sys.argv[1])
    ExcelPath = sys.argv[1]
    if os.path.exists(ExcelPath) == False:
        print("Excel文件夹不存在")
        ExcelPath = "./"
    JsonPath = "./"
    if len(sys.argv) == 3:
        print("第二个传入的参数为:", sys.argv[2])
        JsonPath = sys.argv[2]
        if os.path.exists(JsonPath) == False:
            print("Json文件夹不存在")
            JsonPath = "./"

amount = 0

for file in os.listdir(ExcelPath):
    if file.startswith("~$") == False and file.endswith(".xlsx"):
        print("Excel文件为：",ExcelPath + "/" + file)
        workbook = openpyxl.load_workbook(ExcelPath + "/" + file) 
        print(workbook.sheetnames)
        result = {}
        for  worksheet in workbook.sheetnames:
            if is_chinese(worksheet): 
                continue;
            data = [] 
            for row in workbook[worksheet].iter_rows(values_only=True): 
                data.append(row) 
            
                lines = []
                for row in data: 
                    lines.append(list(row))
     
            sheetData = []
            types = lines.pop(0) #第一行出栈，作为变量类型名
            lines.pop(0) #第二行为注释
            keys = lines.pop(0) #第三行出栈，作为键值

            
            for i in range(len(lines)):#在剩余行中依次建立字典
                itemData = lines[i]                
                if itemData[0] == None:
                    break
                item = {}
                print(itemData)
                for j in range(len(keys)):
                    if keys[j] == None:
                        continue
                    if itemData[j] == "null":
                        item[keys[j]] = None
                    elif types[j] == "STRING":
                        item[keys[j]] = itemData[j] 
                    elif types[j] == "NUMBER":
                        if type(itemData[j]) == int:
                            item[keys[j]] = int(itemData[j])
                        elif type(itemData[j]) == float:
                            item[keys[j]] = float(itemData[j])
                        else:
                            print("数字类型错误")
                    elif types[j] == "ARRAY":
                        item[keys[j]] = itemData[j].split(",")
                sheetData.append(item)
            result[worksheet] = sheetData

        with open(JsonPath +"/"+ file.split(".")[0] + ".json", "w", encoding="utf-8") as jsonFile:
            json.dump(result, jsonFile, ensure_ascii=False)
            jsonFile.close()
        workbook.close()
        amount = amount + 1
        print("转换完成", JsonPath +"/"+ file.split(".")[0] + ".json")
            
print("全部完成" + str(amount)+ "个文件转换")