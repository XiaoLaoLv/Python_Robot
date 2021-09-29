import xlrd
import os
import json


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
