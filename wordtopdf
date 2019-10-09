#!/usr/bin/env python
# -*- coding: utf-8 -*-
# Auth:Gryph0n
from comtypes.client import CreateObject
import os

fPath = input('请输入文件夹路径：')
'''
def list_nohidden(path):
    filesLieBiao = os.listdir(path)
    fileSingleList = [f for f in filesLieBiao]
    fileSingle = str(fileSingleList)
    print(fileSingle)
    judge = str.startswith(fileSingle,0,1)
    if judge == '~':
        pass
    else :
        yield fileSingle
'''
try:
    class pdfConverter:
        def __init__(self):
            self.wdFormatPDF = 17
            self.wdToPDF = CreateObject("Word.Application")


        def wd_to_pdf(self, folder):
            files = os.listdir(folder)
            # files = list_nohidden(folder)
            print(files)
            wdfiles = [f for f in files if f.endswith((".doc", ".docx"))]
            wdfiles2 = [f for f in wdfiles if not f.startswith('~') ]
            for wdfile in wdfiles2:
                wdPath = os.path.join(folder, wdfile)

                if wdPath.split(".")[-1] == 'docx':
                    a = wdPath.rstrip(".docx")
                    pdfPath = a + '.pdf'
                elif wdPath.split(".")[-1] == 'doc':
                    a = wdPath.rstrip(".doc")
                    pdfPath = a + '.pdf'

                #if pdfPath[-3:] != 'pdf':
                 #   pdfPath = pdfPath + '.pdf'
                pdfCreate = self.wdToPDF.Documents.Open(wdPath)
                pdfCreate.SaveAs(pdfPath, self.wdFormatPDF)
                pdfCreate.Close()

    if __name__ == "__main__":
        converter = pdfConverter()
        converter.wd_to_pdf(fPath)
except Exception as e:
    print(e)
input("—————————————————转换完成！Successful—————————————————")
