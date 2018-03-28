#!/usr/bin/env python
# -*- coding: utf-8 -*-
#
# Convert excel file for ISE
#
# alexeykr@gmail.com
# coding=utf-8
import codecs
import argparse
import pandas as pd
import sys
reload(sys)
sys.setdefaultencoding('utf-8')
# import random

description = "Yantar: Convert file, v1.0"
epilog = "https://github.com/alexeykr65/yantar"

flagDebug = int()
fileName = ""
fileNameOut = ""
headerCSV = [
    'Имя:',
    'Фамилия:',
    'Адрес электронной почты:',
    'Номер телефона:',
    'Компания:',
    'Посещаемое лицо (адрес эл. почты):',
    'Причина посещения:',
    'Ваучер №',
    'Дата заезда',
    'Дата отъезда',
    'Номер в АС ОК',
    'Номер №',
    'Путёвка №'
]


def cmdArgsParser():
    global fileName, flagDebug, fileNameOut
    if flagDebug > 0: print "Analyze options ... "
    parser = argparse.ArgumentParser(description=description, epilog=epilog)
    parser.add_argument('-f', '--file', help='File name input', dest="fileName", default='янтарь.xls'.decode('utf-8'))
    parser.add_argument('-o', '--fileout', help='File name output, default=yantar.csv', dest="fileNameOut", default='yantar.csv')
    parser.add_argument('-d', '--debug', help='Debug information view(default =1, 2- more verbose)', dest="flagDebug", default=0)

    arg = parser.parse_args()
    fileName = arg.fileName
    fileNameOut = arg.fileNameOut
    flagDebug = int(arg.flagDebug)


def convertFileToFormatISE():
    df = pd.read_excel(fileName, sheet_name=0, header=4)
    df.columns = ['Num', 'Name', '', '', 'Family', 'Id', 'Num_Putevka', 'Room', '', 'DataIn', 'DataOut']
    fw = codecs.open(fileNameOut, mode="w", encoding="utf-8")
    sRow = ','.join(headerCSV)
    fw.write(sRow.decode('utf-8'))
    fw.write("\r\n")
    for i in df.index:
        listRow = [df['Name'][i], df['Family'][i], '', '', '', '', '', '', str(df['DataIn'][i]).replace('.','/'), str(df['DataOut'][i]).replace('.','/'), str(df['Id'][i]), str(df['Room'][i]), str(df['Num_Putevka'][i])]
        sRow = ','.join(listRow)
        if flagDebug > 0: print sRow
        fw.write(sRow)
        fw.write("\r\n")

    fw.close()


if __name__ == '__main__':
    cmdArgsParser()
    if flagDebug > 0: print "File Name input: " + fileName
    if flagDebug > 0: print "File Name output: " + fileNameOut
    convertFileToFormatISE()
    if flagDebug > 0: print "Script complete successful!!! "
    sys.exit()
