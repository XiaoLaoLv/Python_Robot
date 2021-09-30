import xlrd
import xlwt
import os
import json
import datetime
import random
import time


def readExcel(readExcelPath = '', readSheets = None, contentType = 'json'):
    """
    读取Excel，返回Dict的list
    """
    readExcelPath = str(readExcelPath)
    contentType = str(contentType).lower()
    readSheetsContent = {}  # key为sheet表名称，value为对应sheet的内容list
    readSheetsInfoList = []

    try:
        if not os.path.isfile(os.path.abspath(readExcelPath)):
            readSheetsContent['文件错误'] = '文件名称或路径设置错误[%s]' % (readExcelPath)
            readSheetsInfoList['规则'] = '只设置文件名称, 需要将文件放到当前目录下；如果设置路径，必须为文件的绝对路径，如%s' % (r"E:\test\aa.xls", )
            if contentType != 'dict':
                readSheetsContent = json.dumps(readSheetsContent, ensure_ascii=False)
            return readSheetsContent
        if not (os.path.abspath(readExcelPath).endswith('.xls') or os.path.abspath(readExcelPath).endswith('.xlsx')):
            readSheetsContent['文件错误'] = ' 文件格式错误，必须为excel文件，后缀名只能为.xls和.xlsx'
            if contentType != 'dict':
                readSheetsContent = json.dumps(readSheetsContent, ensure_ascii=False)
            return readSheetsContent

        # 打开文件
        excelInfo = xlrd.open_workbook(readExcelPath)

        # 需要读取的sheet表参数传参时默认读取第一个
        if not isinstance(readSheets, list) or not readSheets:
            temSheet = excelInfo.sheet_by_index(0)
            readSheetsInfoList.append(temSheet)
        else:
            # 如果设置为all， 处理全部Sheet表格
            if len(readSheets) == 1 and str(readSheets[0]).lower() == 'all':
                for sheetIndex in range(0, excelInfo.nsheets):
                    temSheet = excelInfo.sheet_by_index(sheetIndex)
                    readSheetsInfoList.append(temSheet)

            # 按设置的sheet表名称进行处理
            else:
                for sheetName in readSheets:
                    sheetName = str(sheetName)
                    try:
                        temSheet = excelInfo.sheet_by_name(sheetName)
                        readSheetsInfoList.append(temSheet)
                    except:
                        # sheet表名称找不到时返回提示信息，注意：区分大小写
                        readSheetsContent[sheetName] = '该sheet表不存在'

        # 开始处理sheet表内容
        for sheetInfo in readSheetsInfoList:
            dataRows = sheetInfo.nrows
            dataCols = sheetInfo.ncols
            if dataRows > 0:
                temDict = []
                for row in range(0, dataRows):
                    for col in range(0, dataCols):
                        temD = {}
                        temD['row'] = row + 1
                        temD['col'] = col + 1
                        temD['content'] = sheetInfo.cell_value(row, col)
                        temDict.append(temD)
                readSheetsContent[sheetInfo.name] = temDict
            else:
                readSheetsContent[sheetInfo.name] = []

        # 处理返回结果的数据类型，字典或json
        if contentType != 'dict':
            readSheetsContent = json.dumps(readSheetsContent, ensure_ascii=False)
        return readSheetsContent
    except Exception as e:
        return e


def writeExcel(writeExcelPath='', writeExcelName='', sheetHeaders=None, content=None):
    """
    向Excel里写入内容
    """

    # 参数说明
    # writeExcelPath  写入文件的绝对路径，如果不指定，默认在当前目录下生成文件，如r"E:\test"
    # writeExcelName   生成的文件名称（aa.xls aa.xlsx），如果不指定或者与路径下的文件名称重复，名称默认为当前时间加随机数字

    # excelHeaders  sheet表的表头信息，必须时列表格式，否则表头为空，顺序为列表中元素的顺序，格式：
    # [[sheet1_h1, sheet1_h2, ..., sheet1_hn], [sheet2_h1, sheet2_h2, ..., sheet2_hn], ...]

    # content   要写入的内容，从第二行开始写入，默认为空，格式支持字典格式和json格式。格式：
    # {sheetname1:[第一行[第一列content1, 第二列content2, ...], 第二行[第一列content3, 第二列content4, ...], ...], ...}

    # 注意1：如果路径中出现转义字符，如\t，\n等，路径前面加r，如 r"E:\test\aa.txt"
    # 注意2：不能往已有文件中写入数据，只能新增一个文件
    # 注意3：如果len(sheetHeaders)大于len(content)会生成部分只有表头的sheet，如果两个相等则生成的sheet都带表头，否则会有部分不带表头

    writeExcelPath = str(writeExcelPath)
    writeExcelName = str(writeExcelName)

    try:
        # 获取当前路径
        currentPath = os.path.dirname(os.path.abspath(__file__))

        # 如果sheetHeaders不是列表，转为列表，并将表头中的数字转为字符串
        if not isinstance(sheetHeaders, list):
            sheetHeaders = [[sheetHeaders]]
        if sheetHeaders == []:
            sheetHeaders = [sheetHeaders]
        sheetHeadersTem = []
        for x in range(0, len(sheetHeaders)):
            tem = sheetHeaders[x]
            if not isinstance(tem, list):
                aa = []
                aa.append(str(tem))
                sheetHeadersTem.append(aa)
            else:
                bb = []
                for temItem in tem:
                    bb.append(str(temItem))
                sheetHeadersTem.append(bb)
        sheetHeaders = sheetHeadersTem

        # 如果content是json格式，转换为字典，否则将提示信息写入到excel的第一个sheet中
        if not isinstance(content, dict):
            try:
                content = json.loads(content)
            except:
                # return 'content格式不正确，请按格式设置。'
                content = {
                    '': [
                        ["写入的内容格式不正确，写入失败。"],
                        ["格式：{sheetname1:[第一行[第一列content1, 第二列content2, ...], 第二行[第一列content3, 第二列content4, ...], ...], ...}"],
                        [
                            '''示例：{
                                   '第1个sheet表':[[第2行第1列内容, 第2行第2列内容, ..., 第2行第n列内容], [第3行第1列内容, 第3行第2列内容, ..., 第3行第n列内容], ..., [第m行第1列内容, 第m行第2列内容, ..., 第m行第n列内容]],  
                                   '第2个sheet表':[[第2行第1列内容, 第2行第2列内容, ..., 第2行第n列内容], [第3行第1列内容, 第3行第2列内容, ..., 第3行第n列内容], ..., [第m行第1列内容, 第m行第2列内容, ..., 第m行第n列内容]], 
                                   ...
                                   '第n个sheet表':[[第2行第1列内容, 第2行第2列内容, ..., 第2行第n列内容], [第3行第1列内容, 第3行第2列内容, ..., 第3行第n列内容], ..., [第m行第1列内容, 第m行第2列内容, ..., 第m行第n列内容]]
                                }'''
                        ]
                    ]
                }

        # 处理写入文件的目录和文件名称
        if writeExcelPath == '' or not os.path.isdir(os.path.abspath(writeExcelPath)):
            writeExcelPath = currentPath
        writeExcelPath = writeExcelPath.rstrip(r'\\') + r'\\'

        isFileExist = os.path.isfile(writeExcelPath + writeExcelName)
        isNoName = writeExcelName.split('.')[0]
        isRightExtension = writeExcelName.endswith('.xls') or writeExcelName.endswith('.xlsx')
        if isFileExist or not isNoName or not isRightExtension:
            writeExcelName = datetime.datetime.now().strftime('%Y%m%d%H%M%S') + ''.join(random.sample('04512346789567012389', 4)) + ".xls"

        filePath = writeExcelPath + writeExcelName

        # 设置sheet表头样式
        wb = xlwt.Workbook(encoding='utf-8')
        font = xlwt.Font()
        font.bold = True
        style = xlwt.easyxf()
        style.font = font

        # 处理非字符串类型的sheetName（其他类型的sheetName会报错）
        sheetNameList = list(content.keys())

        # 处理非字符串的sheetName，否则会报错
        for k in range(0, len(sheetNameList)):
            if not isinstance(sheetNameList[k], str):
                content[str(sheetNameList[k])] = content.pop(sheetNameList[k])

        # 按内容中的sheetName的个数和sheet表头的个数来确定sheet表的最大数量
        contentSheets, headerSheets = len(content), len(sheetHeaders)
        if contentSheets > headerSheets:
            for j in range(0, contentSheets - headerSheets):
                sheetHeaders.append([])
        elif contentSheets < headerSheets:
            for jj in range(0, headerSheets - contentSheets):
                sheetNameRand = 'randomSheet' + str(jj + 1)
                if sheetNameRand in sheetNameList:
                    sheetNameRand = 'randomSheet' + ''.join(
                        random.sample('019280190192837465', random.randint(1, 3))) + ''.join(
                        random.sample('019280190192837465', random.randint(1, 3)))
                content[sheetNameRand] = []
                sheetNameList.append(sheetNameRand)

        # 准备写入excel
        for i in range(0, len(content)):
            sheetName = str(sheetNameList[i])
            contentList = content[sheetName]
            # sheetName为空时自动生成一个，如果sheetName已存在，则在后面添加加下划线和随机数
            if not sheetName:
                sheetName1 = 'randomSheet' + ''.join(random.sample('045123467804512346704512346789567012389895670123899567012389', random.randint(1,6)))
                ws = wb.add_sheet(sheetName1)
            else:
                ws = wb.add_sheet(sheetName)

            # 设置sheet表头信息
            if i <= headerSheets:
                for ii in range(0, len(sheetHeaders[i])):
                    ws.write(0, ii, sheetHeaders[i][ii], style)

            # 处理写入的行数据，如果行数据类型不是列表则默认写入第二行第一列的单元格
            if not isinstance(content, list):
                ws.write(1, 0, contentList)
                continue
            for row in range(0, range(contentList)):
                # 处理写入的列数据，如果列数据类型不是列表则默认写入当前行第一列的单元格
                if not isinstance(contentList[row], list):
                    ws.write(row + 1, 0, contentList[row])
                    continue
                # 写入数据
                for col in range(0, len(contentList[row])):
                    ws.write(row+1, col, contentList[row][col])

        # 生成excel文件
        try:
            f = open(filePath, 'r')
            f.close()
        except IOError:
            f = open(filePath, 'w')
        wb.save(filePath)
        return {'结果': '写文件成功', '文件地址': filePath}

    except Exception as e:
        return {'结果': '写文件失败', '错误信息': e}





