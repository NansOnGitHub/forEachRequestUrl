import requests
import xlrd
import xlwt


def myEecelToDIC(filename):
    # 转换Excel到dictionary
    d = {}
    wb = xlrd.open_workbook(filename)
    sh = wb.sheet_by_index(0)
    col_len = len(sh.col_values(0))
    for i in range(col_len):
        cell_value_class = sh.cell(i, 4).value
        cell_value_id = sh.cell(i, 1).value
        d[cell_value_id] = cell_value_class
    return d


def requestTest(imgurl):
    response = requests.get(imgurl)
    return response


def writeIntoFile(d):
    wb = xlwt.Workbook()
    sheet = wb.add_sheet("fail")
    a = 0
    b = 0
    for i, j in d.items():
        sheet.write(a, b, i)
        sheet.write(a, b + 1, j)
        a = a +1
        wb.save("fail.xls")


def processToFile():
    d = {}
    excelFile = "nobgcolor-material-object.xls"
    dic = myEecelToDIC(excelFile)
    for i, j in dic.items():
        try:
            # TODO 需要拓展为并发，否则几百个数据也要几十秒才可以完成
            sendResp = requestTest(j)
            code = sendResp.status_code
        except:
            code = 404
        if code == 200:
            print("pk=" + str(i) + "-url正常\n")
        else:
            print("pk=" + str(i) + "-url【异常】\n")
            d[i] = j
    writeIntoFile(d)
    print("DONE!")


if __name__ == '__main__':
    processToFile()
