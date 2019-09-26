#-*-coding:utf-8-*-

#功能:汇总多份格式相同的excel文件的内容，输出一份文件
#环境:Python3.5, PyCharm2017, Win10
#需要的库:pandas、xlrd、openyxl（read/write Excel 2010 xlsx/xlsm files）

#注意事项：
#1）输出文件必须提前存在，并且存在索引值。否则会更新失败。
#2）输出文件被更新后，样式会变化，需手动复原样式。更新操作不会修改内容。

import sys
import importlib
import lib.reload(sys)

import numpy as np
import pandas as pd

def main():
    xlsfile1 = r'T:\workshop_for_python\Se3560\G1.xlsx'

    xlsfile2 = r'T:\workshop_for_python\Se3560\2-VPDN.xlsx'

    xlsfile3 = r'T:\workshop_for_python\Se3560\2-3G-4G.xlsx'

    xlsfile0 = r'T:\workshop_for_python\Se3560\output2.xlsx'
    
	#读取各excel文件的每一个sheet到不同DataFrame，命名格式：df[文件序号][sheet序号]
	#使用ip地址作为索引
    df10 = pd.read_excel(xlsfile1,sheet_name=0,index_col='ip')
    df11 = pd.read_excel(xlsfile1,sheet_name=1,index_col='ip')
    df12 = pd.read_excel(xlsfile1,sheet_name=2,index_col=u'IP地址')

    df20 = pd.read_excel(xlsfile2,sheet_name=0,index_col='ip')
    df21 = pd.read_excel(xlsfile2,sheet_name=1,index_col='ip')
    df22 = pd.read_excel(xlsfile2,sheet_name=2,index_col=u'IP地址')

    df30 = pd.read_excel(xlsfile3,sheet_name=0,index_col='ip')
    df31 = pd.read_excel(xlsfile3,sheet_name=1,index_col='ip')
    df32 = pd.read_excel(xlsfile3,sheet_name=2,index_col=u'IP地址')

    df00 = pd.read_excel(xlsfile0,sheet_name=0,index_col='ip')
    df01 = pd.read_excel(xlsfile0,sheet_name=1,index_col='ip')
    df02 = pd.read_excel(xlsfile0,sheet_name=2,index_col=u'IP地址')

	#利用各组数据更新汇总DataFrame
    df00.update(df10)
    df01.update(df11)
    df02.update(df12)

    df00.update(df20)
    df01.update(df21)
    df02.update(df22)

    df00.update(df30)
    df01.update(df31)
    df02.update(df32)
    
	#输出汇总DataFrame到excel文件
    writer = pd.ExcelWriter('T:\workshop_for_python\Se3560\output2.xlsx')

    df00.to_excel(writer,u'疑似主机')
    df01.to_excel(writer,u'疑似网络设备')
    df02.to_excel(writer,u'未知设备类型')

    writer.save()

if __name__ == "__main__":
    main();


