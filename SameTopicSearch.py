# -*- coding: utf-8 -*-
import os
import sys
import time
import xlrd
from xlutils.copy import copy
import requests
import json
from datetime import datetime

# 查询包含关键字的文件个数
def searchApi(keyword, filterDir, filterExt):
    url = "http://127.0.0.1:9920"
    data = {
        "id": 123,
        "jsonrpc": "2.0",
        "method": "ATRpcServer.Searcher.V1.Search",
        "params": {
            "input": {
                "pattern": keyword,
                "filterDir": filterDir,
                "filterExt": filterExt
            }
        }
    }
    headers = {"Content-Type": "application/json", "Accept": "application/json"}
    # 将数据转换为JSON字符串
    data_json = json.dumps(data, ensure_ascii=False)
    # 发送POST请求
    response = requests.post(url, data=data_json, headers=headers)
    # 获取响应结果
    response_text = response.text
    # 将结果转换成dict
    return json.loads(response_text)

# 题干检查
def topicMatch(matchDir, isSplit, sheet, copySheet):
    # 试题内容列，因为导出的模板不会变化，所以直接写死列值，注意索引从0开始
    topicColNum = 9
    # 检索字段写入列
    valueColNum = 17
    # 比较结果写入列
    compareResultColNum = 18
    # 从第4行开始读取（索引为3）
    for row_idx in range(3, sheet.nrows):
        # 获取试题内容
        cellValue = sheet.cell(row_idx, topicColNum).value.strip()
        cellValue = cellValue.replace("【图或公式丢失】","").replace("【图】","")
        # cell内没有数据则跳过
        if cellValue is None or cellValue == "":
            continue
        # 试题内容按照（    ）拆分，获取字符最长的那部分进行检查
        value = max(cellValue.split("（    ）"), key=len)
        # 精准查询，要搜索的内容增加上双引号则不会被拆词,则就变成了精确查找
        if isSplit == "0":
            value = '"' + value + '"'
        # 试题内容比较
        resDict = searchApi(value, matchDir, "*")
        count = resDict.get("result").get("data").get("output").get("count")
        # 检查结果写入目标sheet
        copySheet.write(row_idx, valueColNum, value)
        copySheet.write(row_idx, compareResultColNum, count)
        print("第" + str(row_idx + 1) + "行，检索内容：" + value + "，查重个数：" + str(count))
        # 100ms执行一次
        time.sleep(0.1)

# 按试题选项进行匹配
def optionMatch(matchDir, isSplit, sheet, copySheet):
    # 试题内容列，因为导出的模板不会变化，所以直接写死列值，注意索引从0开始
    topicColNum = 10
    # 比较结果写入列
    compareResultColNum = 17
    # 从第4行开始读取（索引为3）
    for row_idx in range(3, sheet.nrows):
        # 获取试题内容
        cellValue = sheet.cell(row_idx, topicColNum).value.strip()
        cellValue = cellValue.replace("【图或公式丢失】","").replace("【图】","")
        # cell内没有数据则跳过
        if cellValue is None or cellValue == "":
            continue
        # 精准查询，要搜索的内容增加上双引号则不会被拆词,则就变成了精确查找
        if isSplit == "0":
            cellValue = '"' + cellValue + '"'
        # 试题内容比较
        resDict = searchApi(cellValue, matchDir, "*")
        count = resDict.get("result").get("data").get("output").get("count")
        # 检查结果写入目标sheet
        copySheet.write(row_idx, compareResultColNum, count)
        # 控制台打印
        print("第" + str(row_idx + 1) + "行，检索内容：" + cellValue + "，查重个数：" + str(count))
        # 100ms执行一次
        time.sleep(0.1)

# 主函数
if __name__ == '__main__':
    # 第一个参数：用来校验的目录地址,需注意：1、windows盘符需要大写；2、路径中如果有空格需要用引号包裹起来
    # 第二个参数：要校验的excel文件名称，需要将校验文件放在执行程序同级目录
    # 第三个参数：要校验的字段（列）：[topic:题干；option:选项A]；
    # 第四个参数: 是否拆词校验：[0:不拆（精准搜索）；1:拆]
    # 参数个数检查
    if len(sys.argv) != 4 and len(sys.argv) != 5:
        print("参数不对，请使用: SameTopicSearch matchDir fileName matchType isSplit")
        sys.exit(1)
    # 用来校验的目录地址检查
    matchDir = sys.argv[1]
    if not (os.path.exists(matchDir) and os.path.isdir(matchDir)):
        print("用来校验的目录不存在")
        sys.exit(1)
    # 文件校验
    fileName = sys.argv[2]
    filePath = os.getcwd() + "/" + fileName
    if not (os.path.exists(filePath) and os.path.isfile(filePath)):
        print("要校验的excel文件不存在")
        sys.exit(1)
    # 后缀检查
    fileNameList = os.path.splitext(fileName)
    extension = fileNameList[1].lower()
    if not extension in ['.xls', '.xlsx']:
        print("要校验的excel文件格式不对")
        sys.exit(1)
    # 要校验的字段检查
    matchType = sys.argv[3]
    if not matchType in ("topic", "option"):
        print("匹配格式不对")
        sys.exit(1)
    # 是否拆词，默认为0(精准搜索)
    isSplit = "0"
    if len(sys.argv) == 4:
        isSplit = "0"
    else:
        isSplit = sys.argv[4]
        #是否拆词标记
        if isSplit not in ("0", "1"):
            print("是否拆词标记格式输入有误")
            sys.exit(1)

    # ------------------ 读取excel文件------------------
    wb = xlrd.open_workbook(filePath)
    # 通过索引获取第一个工作表
    sheet = wb.sheet_by_index(0)
    if not sheet.name == "正式题目":
        print("被校验文件sheet名称错误")
        sys.exit(1)
    # 最后一行
    lastRowNum = sheet.nrows
    # 有要检查的数据才执行
    if not lastRowNum > 3:
        print("sheet内无要检查的数据，执行结束")
        sys.exit(1)
    # 因为出题模板使用的excel版本较早，openpyxl不支持，所以采用了xlrd。
    # xlrd不支持边读边写，为了能够写入核验结果，需复制原workbook做个新文件
    copyWb = copy(wb)
    copySheet = copyWb.get_sheet(0)
    # 按条件核验结果
    if matchType == "topic":
        topicMatch(matchDir, isSplit, sheet, copySheet)
    else:
        optionMatch(matchDir, isSplit, sheet, copySheet)
    # 检查结果存储为新文件
    # 格式化日期和时间为年月日时分秒字符串
    datetime_str = datetime.now().strftime("%Y%m%d%H%M%S")
    copyWb.save(fileNameList[0] + "-checked-" + datetime_str + fileNameList[1])



