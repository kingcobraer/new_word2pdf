from win32com.client import Dispatch
from os import walk
import streamlit as st
import pythoncom

directory = st.text_input('请输入文件夹地址')
if directory is None:
    st.stop

wdFormatPDF = 17
def doc2pdf(input_file):
    pythoncom.CoInitialize()
    word = Dispatch('Word.Application')
    doc = word.Documents.Open(input_file)
    doc.SaveAs(input_file.replace(".docx", ".pdf"), FileFormat=wdFormatPDF)
    doc.Close()
    word.Quit()
    #释放资源
    pythoncom.CoUninitialize()

if __name__ == "__main__":
    doc_files = []
    for root, dirs, filenames in walk(directory):
        for file in filenames:
            if file.endswith(".doc") or file.endswith(".docx"):
                doc2pdf(str(root + "\\" + file))
'# 1 级 标 题'
