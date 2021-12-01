# -*- coding: utf-8 -*-
'''
Created on Sat Jun 19 14:06:21 2021

@author: yoshi@nagaokauniv.ac.jp
'''
version = 'ver.0.94'
#ver.0.94   20211130    コニカミノルタ複合機　ページ位置対応

import sys
import os
import glob
import datetime
import fitz
from pathlib import Path
import tkinter as tk
from tkinter import *
from tkinter import ttk
from tkinter import messagebox
from tkinter import filedialog
from configobj import ConfigObj
import subprocess
from natsort import natsorted

#global val
inputPath = os.path.abspath(os.path.dirname(__file__))
outputPath = os.path.abspath(os.path.dirname(__file__))
outputFileName = ''
headerText = ''
pageNumber = False
headerX = 0
headerY = 10
fontSize = 8

def readConfigFile():
    global inputPath
    global outputPath
    global headerText
    global pageNumber
    global headerX
    global headerY
    global fontSize

    config = ConfigObj('MergePdf.ini', encoding='utf-8')
    if config:
        inputPath = config['inputPath']
        outputPath = config['outputPath']
        headerText = config['headerText']
        pageNumber = eval(config['pageNumber'])
        headerX = config['headerX']
        headerY = config['headerY']
        fontSize = config['fontSize']
    
def writeConfigFile():
    config = ConfigObj('MergePdf.ini', encoding='utf-8')
    config['inputPath'] = inputPath
    config['outputPath'] = outputPath
    config['outputFileName'] = outputFileName
    config['headerText'] = headerText
    config['pageNumber'] = pageNumber
    config['headerX'] = headerX
    config['headerY'] = headerY
    config['fontSize'] = fontSize
    config.write()

# ページ表示位置修正用座標計算
def adjustPoint(p, hPage, hImage):
    r = hImage / hPage
    x = float(p[0]) * r
    y = float(p[1]) * r + hPage - hImage
    return (x, y)

def UpdateListBox():
    listbox.delete(0, END)
    if inputPath:
        pdf_files = natsorted(glob.glob(inputPath + '/*.pdf'))
        listbox.delete(0, END)
        for file in pdf_files:
            listbox.insert(END, os.path.basename(file))

def UpdateDialog():
    # フォルダ内のPDFファイル一覧
    global inputPath
    global outputPath
    global headerText
    global pageNumber
    global headerX
    global headerY
    global fontSize
    inputPath = eInputPath.get()
    outputPath = eOutputPath.get()
    headerText = eHeaderText.get()
    pageNumber = bPageNumber.get()
    headerX = eHeaderX.get()
    headerY = eHeaderY.get()
    fontSize = eFontSize.get()
    UpdateListBox()

# フォルダ指定の関数
def input_dirdialog_clicked():
    global inputPath
    inputPath = eInputPath.get()
    if inputPath:
        iDir = inputPath
    else:
        iDir = os.path.abspath(os.path.dirname(__file__))

    fTyp = [('PDFファイル', '*.pdf')]
    iDirPath = filedialog.askopenfilename(filetype = fTyp, initialdir = iDir)
    if iDirPath:
        inputPath = os.path.abspath(os.path.dirname(iDirPath))
        eInputPath.set(inputPath)
        UpdateListBox()

# ファイル指定の関数
def output_dirdialog_clicked():
    global outputPath
    outputPath = eOutputPath.get()
    if outputPath:
        iFile = outputPath    
    else:
        iFile = os.path.abspath(os.path.dirname(__file__))
    iFilePath = filedialog.askdirectory(initialdir = iFile)
    if iFilePath:
        outputPath = os.path.abspath(iFilePath)
        eOutputPath.set(outputPath)

# 実行ボタン押下時の実行関数
def conductMain():
    global outputFileName
    
    UpdateDialog()

    if inputPath is None:
        messagebox.showerror('error', 'フォルダの指定がありません。')
        return 0

#    documents_path = os.getenv('HOMEDRIVE') + os.getenv('HOMEPATH') + '\\My Documents'
    # フォルダ内のPDFファイル一覧
    pdf_files = natsorted(glob.glob(inputPath + '/*.pdf'))
    if 0 == len(pdf_files):
        messagebox.showerror('error', 'PDFファイルが見つかりません')
        return 0

    # １つのPDFファイルにまとめる
    pdf_writer = fitz.open()
    page_info = []
    for file in pdf_files:
        pdf_reader = fitz.open(str(file))
        pdf_writer.insertPDF(pdf_reader)
        # 座標系修正のためページごとにのmetadata，画像情報記録
        for n in range(0, pdf_reader.page_count):
            pInfo = []
            page = pdf_reader[n]
            pInfo.append(pdf_reader.metadata)
            for j, img in enumerate(page.getImageList()):
                pInfo.append(pdf_reader.extractImage(img[0]))
            page_info.append(pInfo)

    # 保存ファイル名（先頭と末尾のファイル名で作成）
    now = datetime.datetime.now()
    outputFileName = outputPath + '\\' + Path(pdf_files[0]).stem + '-' + now.strftime('%Y%m%d_%H%M%S') + '.pdf'

   #ページ番号
    if pageNumber or 0 < len(headerText):
        pageCount = pdf_writer.page_count 
        for num in range(0, pdf_writer.page_count):
            page = pdf_writer[num]
            pInfo = page_info[num]   # metadata，画像情報
            text = headerText + ' [{:3}/{:3}]'.format(num + 1, pageCount)
            tw = fitz.TextWriter(page.rect)
            if(1 < len(pInfo) and pInfo[0]['format'] == 'PDF 1.3'): # コニカミノルタ複合機ページ位置補正用
                hImage = pInfo[1]['height']
                hPage = page.rect.y1
                tw.append(adjustPoint((headerX, headerY), hPage, hImage), text, fontsize = float(fontSize) * hImage / hPage)
            else:
                tw.append((headerX, headerY), text, fontsize = float(fontSize))
            tw.write_text(page)
    # 保存
    pdf_writer.save(outputFileName)
    writeConfigFile()
    subprocess.Popen(['start', '', outputFileName], shell=True)
    
if __name__ == '__main__':
    readConfigFile()
    
    # rootの作成
    root = Tk()
    root.title('PDFファイル結合' + ' ' + version)

    # 参照フォルダ
    frame1 = ttk.Frame(root, padding=10)
#    frame1.grid(row=0, column=1, sticky=W)
    frame1.pack()
    ILabel = ttk.Label(frame1, text='参照フォルダ:')
    ILabel.pack(side=LEFT)
    eInputPath = StringVar()
    eInputPath.set(inputPath)
    IEntry = ttk.Entry(frame1, textvariable=eInputPath, width=60)
    IEntry.pack(side=LEFT)
    # 「ファイル参照」ボタンの作成
    IButton = ttk.Button(frame1, text='参照', command=input_dirdialog_clicked)
    IButton.pack(side=LEFT)

    # 出力ファイル
    frame2 = ttk.Frame(root)
    frame2.pack()
    ILabel = ttk.Label(frame2, text='出力フォルダ:', padding=(0, 0))
    ILabel.pack(side=LEFT)
    eOutputPath = StringVar()
    eOutputPath.set(outputPath)
    IEntry = ttk.Entry(frame2, textvariable=eOutputPath, width=60)
    IEntry.pack(side=LEFT)
    # 「ファイル参照」ボタンの作成
    IButton = ttk.Button(frame2, text='参照', command=output_dirdialog_clicked)
    IButton.pack(side=LEFT)

    # ヘッダ文字列
    frame3 = ttk.Frame(root, padding=10)
    frame3.pack()
    ILabel = ttk.Label(frame3, text='ヘッダ文字列:', padding=(5, 2))
    ILabel.pack(side=LEFT)
    eHeaderText = StringVar()
    eHeaderText.set(headerText)
    IEntry = ttk.Entry(frame3, textvariable=eHeaderText, width=64)
    IEntry.pack(side=LEFT)

    frame4 = ttk.Frame(root, padding=10)
    frame4.pack()
    bPageNumber = BooleanVar()
    bPageNumber.set(pageNumber)
    chk = ttk.Checkbutton(frame4, variable=bPageNumber, text='ページ番号を付ける')
    chk.pack(side=LEFT)

    ILabel = ttk.Label(frame4, text='横位置:', padding=(5, 2))
    ILabel.pack(side=LEFT)
    eHeaderX = StringVar()
    eHeaderX.set(headerX)
    IEntry = ttk.Entry(frame4, textvariable=eHeaderX, width=6)
    IEntry.pack(side=LEFT)
    ILabel = ttk.Label(frame4, text='縦位置:', padding=(5, 2))
    ILabel.pack(side=LEFT)
    eHeaderY = StringVar()
    eHeaderY.set(headerY)
    IEntry = ttk.Entry(frame4, textvariable=eHeaderY, width=6)
    IEntry.pack(side=LEFT)
    ILabel = ttk.Label(frame4, text='フォントサイズ:', padding=(5, 2))
    ILabel.pack(side=LEFT)
    eFontSize = StringVar()
    eFontSize.set(fontSize)
    IEntry = ttk.Entry(frame4, textvariable=eFontSize, width=6)
    IEntry.pack(side=LEFT)

    #リストボックスの作成
    frame5 = ttk.Frame(root, padding=10)
    frame5.pack()
    list_value=tk.StringVar()
    listbox=tk.Listbox(frame5,height=20,width=64,listvariable=list_value)
    listbox.pack()

  # ボタンの設置
    frame9 = ttk.Frame(root, padding=10)
    frame9.pack()
    button1 = ttk.Button(frame9, text='実行', command=conductMain)
    button1.pack(fill = 'x', padx=30, side = 'left')
    button2 = ttk.Button(frame9, text=('閉じる'), command=quit)
    button2.pack(fill = 'x', padx=30, side = 'left')

    UpdateDialog()

    root.mainloop()
