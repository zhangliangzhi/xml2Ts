#!/usr/bin/python3.7.8
# -*- coding: utf-8 -*- 
# 自动生成表
import codecs
import xdrlib,sys
import xlrd
import time
import os
import importlib
# imp.reload(sys)

# 1.导出文件名字
outFileName             = "gPropData"
outFileNameCN           = u"d道具表"
# 2.excel表格名字
excelName               = u"d道具表.xlsx"
# 3.导出该excel的第几个表格
sheetIndex              = 0
# 4.定义每个字段的名字
# 第一行说明
arrDesc = []
# 第二行字段名
arrTitle = []
importlib.reload(sys)
# sys.setdefaultencoding('utf8')

# print(len(arrTitle), arrTitle)

outputjsfile   = codecs.open('D:/'+outFileName+".ts", 'w', 'utf-8')
outputjsfile.write(u"// author:\t项目-自动生成\n// name:\t" + outFileNameCN + ' ' + outFileName + "\n// genTime:\t"+ time.strftime("%Y-%m", time.localtime())  +"\n\n")

# 开始写输出文件
outputjsfile.write("import { D } from \"./AIndex\";\n")
outputjsfile.write("D." + outFileName + " = [\n")


def is_float(s):
    try:
        float(s)
        return True
    except ValueError:
        pass
    try:
        import unicodedata
        unicodedata.numeric(s)
        return True
    except (TypeError, ValueError):
        pass
    return False

def is_int(s):
    try:
        int(s)
        return s % 1 == 0
    except ValueError:
        return False

def open_excel(file):
    try:
        data = xlrd.open_workbook(file)
        return data
    except Exception as e:
        # print str(e)
        return
# 根据索引获取Excel表格中的数据   参数:file：Excel文件路径     colnameindex：表头列名所在行的索引  ，by_index：表的索引
def excel_table_byindex(file= excelName, by_index=sheetIndex, colnameindex=0):
    data = open_excel(file)
    table = data.sheets()[by_index]
    nrows = table.nrows #行数
    ncols = table.ncols #列数

    # 空一列，就截止，不读下去了，后面的都为备注
    for icol in range(0,ncols):
        oneColData = []
        oneColData = table.col_values(icol)
        value = oneColData[0]
        if not value:
            ncols = icol-1
            break

    print( u'【'+ excelName + u'】表格共有：' + str(nrows) + u'行 ' + str(ncols) + u'列')
    # print(nrows, ncols)

    colnames = table.row_values(colnameindex) #某一行数据


    # 第0行为备注, 从第2行开始
    for irow in range(0,nrows):
        # outputjsfile.write("\t")
        oneRowData = []
        oneRowData = table.row_values(irow)

        # 拆分字段
        for icol in range(0,ncols):
            value = oneRowData[icol]
            if irow == 0:
                arrDesc.append(value)
                continue
            if irow == 1:
                arrTitle.append(value)
                if icol == 0:
                    outputjsfile.write("// ")
                print(arrDesc)
                outputjsfile.write(str(icol+1) + arrDesc[icol] + '(' + value+"), ")
                continue
            
            # excel里面为空内容时，填0
            if not value:
                value = 0
            # 数字用:
            # outputjsfile.write("\""+arrTitle[icol]+"\":"+str(int(value))+", ")
            # 字符串用:
            # outputjsfile.write("\""+arrTitle[icol]+"\":\""+((value))+"\", ")
            # 数字字符串用:
            # outputjsfile.write("\""+arrTitle[icol]+"\":\""+str(int(value))+"\", ")

            if icol == 0:
                # 开始： id字段, float->int->str
                # outputjsfile.write(""+str(int(value))+": { ")
                outputjsfile.write("{\""+arrTitle[icol]+"\":"+str(int(value))+", ")
                pass
                pass
            else:
                strEnd = ', '
                # 最后: 不加逗号
                if icol == (ncols-1):
                    strEnd = ' '

                if is_int(value):
                    outputjsfile.write("\""+arrTitle[icol]+"\":"+str(int(value)) + strEnd)
                elif is_float(value):
                    outputjsfile.write("\""+arrTitle[icol]+"\":"+str(float(value)) + strEnd)
                else:
                    outputjsfile.write("\""+arrTitle[icol]+"\":\""+((value))+"\"" + strEnd)
                pass
            # print(oneRowData[icol])
            pass
        # print(u"第"+str(irow)+u"行解析成功")
        # print(oneRowData)

        # 最后一个数字不要加，号
        if irow == (nrows-1) :
            outputjsfile.write("}\n")
        elif irow == 0:
            pass
        elif irow == 1:
            outputjsfile.write("\n\t")
        else:
            outputjsfile.write("},\n\t")
       
        pass

        pass

       

       
        # outputjsfile.write("\t\t")



def main():
    excel_table_byindex()

main()
outputjsfile.write("]\n")
outputjsfile.close()
# 按任意键继续
os.system('pause')