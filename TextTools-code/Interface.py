# -*- coding: utf-8 -*-
# @Time : 2021/6/6 20:55
# @Author : DrMelt
# @File : TextTools
# @Software : PyCharm

import sys
from tkinter import filedialog
import time
import epub
from bs4 import BeautifulSoup
import re
import os
from Functions.functions import functions
import tkinter as tk
import tkinter.messagebox
from PIL import Image, ImageTk
import shutil
import chardet
import zipfile

subFiles = []
softName = 'TextTools v1.4'
versions = 'v1.4'
RBforTXTset_save = ''
Functions = functions()
setupParameters = {'Passworld': 'ForRuBi', 'login': 'ForRuBi', 'RBmode': 'hiragana', 'saveImages': 'True',
                   'ConvertImage': 'False', 'Imagescale': '2',
                   'w:rFonts': '"Yu Mincho"', 'w:hps': '10', 'w:hpsRaise': '18', 'w:hpsBaseText': '21'}
setting_default = '''
login = ForRuBi
Passworld = ForRuBi
RBmode = hiragana
# hiragana:0/pykakasi:1


saveImages = False
#DOCX sets
ConvertImage = False
#图片缩放比 9525//Imagescale
Imagescale = 2
#字体
w:rFonts = "Yu Mincho"
#注音大小
w:hps = 10
#注音高度
w:hpsRaise = 18
w:hpsBaseText = 21
'''

handle = {"setWindow": 0}


def EPUBtoTXT():
    readSet()
    subFiles.clear()
    # 文件夹选择
    root = tk.Tk()
    root.withdraw()
    Filepath = filedialog.askopenfilename(filetypes=[('EPUB', '*.epub')])

    print(Filepath)
    os.chdir(os.path.split(Filepath)[0])  # 改正工作目录
    Filename = (os.path.split(Filepath)[1])[0:len(os.path.split(Filepath)[1]) - 5]
    print(Filename)

    File = epub.open_epub(Filepath)
    File.extractall(os.path.split(Filepath)[0] + os.sep + Filename)
    File.close()

    fileRoot = os.path.abspath(".")
    opfpath = Functions.findSubFiles(dir2=os.path.split(Filepath)[0] + os.sep + Filename,
                                     rootSite=os.path.split(Filepath)[0] + os.sep + Filename, subFiles=subFiles)
    os.chdir(os.path.split(Filepath)[0])  # 改正输出目录

    # 合并到m
    if os.path.exists("%s.txt" % Filename):
        m = open("%s.txt" % Filename, encoding='utf-8', mode='w')
    else:
        m = open("%s.txt" % Filename, encoding='utf-8', mode='x')

    for i in range(0, len(subFiles)):  # 前闭后开
        f = open("%s" % subFiles[i], encoding='utf-8', mode='r')
        content = f.readlines()  # 列表，全部读完
        for temp in content:
            if setupParameters['saveImages'] == 'False':
                temp = re.sub(r'<(img|image).*?(src|href)="(.*?)".*?/>',
                              '<p><br/></p>\n<p><br/></p>\n<p><br/></p>\n<p><br/></p>\n<p><br/></p>', temp)
            elif setupParameters['saveImages'] == 'True':
                htmlFilesPath = re.sub('//', '/',
                                       re.sub(r'\\', '/', os.path.split(subFiles[i])[0])[
                                       len(os.path.split(Filepath)[0]):] + '/')
                temp = re.sub(r'<(img|image).*?(src|href)="(.*?)".*?/>', r'<p>[image]%s\3[/image]</p>' % htmlFilesPath,
                              temp)  # <img class="fit" src="../image/cover.jpg" alt=""/>
                # <image height="2048" width="1440" xlink:href="../Images/cover00256.jpeg"/>
                temp = re.sub(r'%s/\.\./' % os.path.split(htmlFilesPath[0:len(htmlFilesPath) - 1])[1], r'', temp)
            m.write("%s" % temp)
        f.close
    m.close()

    file = open("%s.txt" % Filename, mode='rb')
    m = file.read().decode("utf-8")
    file.close()
    result = open("%s.txt" % Filename, encoding='utf-8', mode='w')
    xm = BeautifulSoup(m, "html.parser")
    mp = ''
    for item in xm.find_all("p"):
        if re.search(r'<p.*?>(.*?)</p>', str(item)) is None:
            continue
        m0 = (re.search(r'<p.*?>(.*?)</p>', str(item))).group(1)
        # m0 = re.sub(r'\(', '（', m0)
        # m0 = re.sub(r'\)', '）', m0)
        m0 = re.sub(r'<rt>.*?</rt>', r'', m0)
        m0 = re.sub(r'<.*?>', '', m0)
        if m0[:7] != '[image]' or mp != m0 and m0[:7] == '[image]':
            result.write(str(m0) + "\n")  # 写入result
        mp = m0
    result.close()
    tk.messagebox.showinfo(title='message', message='Finished')


def EPUBtoDOCX():
    readSet()
    subFiles.clear()
    # 文件夹选择
    root = tk.Tk()
    root.withdraw()
    Filepath = filedialog.askopenfilename(filetypes=[('EPUB', '*.epub')])

    Filename = (os.path.split(Filepath)[1])[0:len(os.path.split(Filepath)[1]) - 5]
    dirName = Filename + '-DOCX'
    dirPath = os.path.split(Filepath)[0] + os.sep + dirName
    # 复制docx样本
    os.chdir(operatingPath)
    if os.path.exists(dirPath):
        shutil.rmtree(dirPath)
        time.sleep(1)
        shutil.copytree("docxSample", dirPath)
    else:
        shutil.copytree("docxSample", dirPath)

    print(Filepath)
    os.chdir(os.path.split(Filepath)[0])  # 改正工作目录
    Filename = (os.path.split(Filepath)[1])[0:len(os.path.split(Filepath)[1]) - 5]
    print(Filename)

    File = epub.open_epub(Filepath)
    File.extractall(os.path.split(Filepath)[0] + os.sep + Filename)
    File.close()

    fileRoot = os.path.abspath(".")
    opf = Functions.findSubFiles(dir2=os.path.split(Filepath)[0] + os.sep + Filename,
                                 rootSite=os.path.split(Filepath)[0] + os.sep + Filename, subFiles=subFiles)
    os.chdir(os.path.split(Filepath)[0])  # 改正输出目录

    # 合并到m
    if os.path.exists("%s.buffer" % Filename):
        m = open("%s.buffer" % Filename, encoding='utf-8', mode='w')
    else:
        m = open("%s.buffer" % Filename, encoding='utf-8', mode='x')

    for i in range(0, len(subFiles)):  # 前闭后开
        f = open("%s" % subFiles[i], encoding='utf-8', mode='r')
        content = f.readlines()  # 列表，全部读完
        for temp in content:
            if setupParameters['saveImages'] == 'False':
                temp = re.sub(r'<(img|image).*?(src|href)="(.*?)".*?/>',
                              '<p><br/></p>\n<p><br/></p>\n<p><br/></p>\n<p><br/></p>\n<p><br/></p>', temp)
            elif setupParameters['saveImages'] == 'True':
                htmlFilesPath = re.sub('//', '/',
                                       re.sub(r'\\', '/', os.path.split(subFiles[i])[0])[
                                       len(os.path.split(Filepath)[0]):] + '/')
                temp = re.sub(r'<(img|image).*?(src|href)="(.*?)".*?/>', r'<p>[image]%s\3[/image]</p>' % htmlFilesPath,
                              temp)  # <img class="fit" src="../image/cover.jpg" alt=""/>
                # <image height="2048" width="1440" xlink:href="../Images/cover00256.jpeg"/>
                temp = re.sub(r'%s/\.\./' % os.path.split(htmlFilesPath[0:len(htmlFilesPath) - 1])[1], r'', temp)
            m.write("%s" % temp)
        f.close
    m.close()

    file = open("%s.buffer" % Filename, mode='rb')
    m = file.read().decode("utf-8")
    file.close()
    xm = BeautifulSoup(m, "html.parser")
    Text = ''
    mp = ''
    for item in xm.find_all("p"):
        if re.search(r'<p.*?>(.*?)</p>', str(item)) is None:
            continue
        m0 = (re.search(r'<p.*?>(.*?)</p>', str(item))).group(1)
        # m0 = re.sub(r'(<span.*?>|</span>|<p>)', r'', m0)
        m0 = re.sub(r'(<a.*?>|</a>|<a>|<sup>|<p>|<br.*?>)', r'', m0)
        m0 = re.sub(r'(<span.*?>|</span>)', r'', m0)
        m0 = re.sub('(<br/>|<br />)', '', m0)
        m0 = re.sub('\n', '<br>', m0)
        # m0 = re.sub(r'\(', '（', m0)
        # m0 = re.sub(r'\)', '）', m0)
        mb = ''
        while mb != m0:
            mb = m0
            m0 = re.sub(r'<ruby.*?>(.*?)<rt.*?>(.*?)</rt>(.*?)</ruby>',
                        r'<Ruby><Rb>\1</Rb><Rp>(</Rp><Rt>\2</Rt><Rp>)</Rp></Ruby><ruby>\3</ruby>',
                        m0)  # <ruby>素<rt>す</rt>晴<rt>ば</rt></ruby>
            m0 = re.sub('(<ruby></ruby>)', '', m0)
        m0 = re.sub('(<rb.*?>|</rb>)', '', m0)
        if m0[:7] != '[image]' or mp != m0 and m0[:7] == '[image]':
            Text = Text + m0 + '<br>'
        mp = m0
    # 文字处理完成，写入docx
    Functions.makeDocx(Text, Filepath, 0, setupParameters)
    os.unlink(os.path.split(Filepath)[0] + os.sep + "%s.buffer" % Filename)
    tk.messagebox.showinfo(title='message', message='Finished')


def HTMLtoDOCX():
    readSet()
    subFiles.clear()
    # 文件选择
    root = tk.Tk()
    root.withdraw()
    Filepath = filedialog.askopenfilename()

    File = os.path.split(Filepath)[1]
    for i in range(len(File) - 1, 0, -1):
        if File[i] == '.':
            Filename = File[0:i]
            break
    dirName = Filename + '-DOCX'
    dirPath = os.path.split(Filepath)[0] + os.sep + dirName
    # 复制docx样本
    os.chdir(operatingPath)
    if os.path.exists(dirPath):
        shutil.rmtree(dirPath)
        time.sleep(1)
        shutil.copytree("docxSample", dirPath)
    else:
        shutil.copytree("docxSample", dirPath)

    print(Filepath)
    os.chdir(os.path.split(Filepath)[0])  # 改正工作目录
    print(Filename)

    # 添加处理文件目录
    subFiles.append(Filepath)
    m = u''
    for i in range(0, len(subFiles)):  # 前闭后开
        site = "%s" % subFiles[i]
        fileData = open(site, mode='rb').read()
        encoded = chardet.detect(fileData)['encoding']
        print(encoded)
        f = open(site, encoding=encoded, mode='r')
        content = f.readlines()  # 列表，全部读完
        f.close()
        for item in content:
            m = m + item

    xm = BeautifulSoup(m, "html.parser")
    Text = u''
    for temp in xm.find_all("body"):
        item = str(temp)
        item = re.sub(
            r'(<a.*?>|</a>|<a>|<sup>|<p>|</p>|<br.*?>|</head>|<head>|</ul>|<ul>|</li>|<li>|</body>|<body>|</tr>|<tr>|<td>|</td>)',
            r'', item)
        item = re.sub(r'(<span.*?>|</span>|<div.*?>|</div>|<h.*?>|</h.>|<table.*?>|</table>|<script.*?>|</script>)',
                      r'', item)
        item = re.sub('\n', '<br>', item)
        # mb = ''
        # while mb != item:
        #     mb = item
        #     item = re.sub(r'<ruby.*?>(.*?)<rt.*?>(.*?)</rt>(.*?)</ruby>',
        #                 r'<Ruby><Rb>\1</Rb><Rp>(</Rp><Rt>\2</Rt><Rp>)</Rp></Ruby><ruby>\3</ruby>',
        #                 item)  # <ruby>素<rt>す</rt>晴<rt>ば</rt></ruby>
        #     item = re.sub('(<ruby></ruby>)', '', item)
        # item = re.sub('(<rb>|</rb>)', '', item)
        Text = Text + item + '<br>'
    # 文字处理完成，写入docx
    Text = re.sub('<!--.*?-->', '', Text)
    Functions.makeDocx(Text, Filepath, 0, setupParameters)
    tk.messagebox.showinfo(title='message', message='Finished')


def HTMLtoTXT():
    readSet()
    subFiles.clear()
    # 文件选择
    root = tk.Tk()
    root.withdraw()
    Filepath = filedialog.askopenfilename()

    File = os.path.split(Filepath)[1]
    for i in range(len(File) - 1, 0, -1):
        if File[i] == '.':
            Filename = File[0:i]
            break
    print(Filepath)
    os.chdir(os.path.split(Filepath)[0])  # 改正工作目录
    print(Filename)

    # 添加处理文件目录
    subFiles.append(Filepath)
    m = u''
    for i in range(0, len(subFiles)):  # 前闭后开
        site = "%s" % subFiles[i]
        fileData = open(site, mode='rb').read()
        encoded = chardet.detect(fileData)['encoding']
        print(encoded)
        f = open(site, encoding=encoded, mode='r')
        content = f.readlines()  # 列表，全部读完
        for item in content:
            m = m + item

    xm = BeautifulSoup(m, "html.parser")
    Text = u''
    for temp in xm.find_all("body"):
        Text = Text + str(temp) + '\n'
    # 文字处理完成，写入
    Text = re.sub(r'<rt>.*?</rt>', r'', Text)
    Text = re.sub(r'<rp>.*?</rp>', r'', Text)
    Text = re.sub(r'<.*?>', '', Text, flags=re.S)
    Text = re.sub(r'<|>', '', Text)
    txt = open('%s.txt' % Filename, encoding='utf-8', mode='w')
    txt.write(Text)
    txt.close()
    tk.messagebox.showinfo(title='message', message='Finished')


def RBforTXT():
    readSet()
    value = listbox.get(listbox.curselection())  # 获取当前选中的文本
    RBforTXTsave = value  # 为label设置值
    if RBforTXTsave == '':
        tk.messagebox.showwarning(title='Warning', message='chose one!')
        return 'Warning'

    # 文件选择
    root = tk.Tk()
    root.withdraw()
    Filepath = filedialog.askopenfilename(filetypes=[('TXT', '*.txt')])

    print('\n')
    print(Filepath)
    print('\n')
    os.chdir(os.path.split(Filepath)[0])  # 改正工作目录
    Filename = (os.path.split(Filepath)[1])[0:len(os.path.split(Filepath)[1]) - 4]
    print(Filename)
    print('\n')

    fromText = functions().readFile(Filepath)
    if RBforTXTsave == 'saveAsTXT':
        m = open('%s(RuBi).txt' % Filename, encoding='utf-8', mode='w')
        if setupParameters['RBmode'] == 'hiragana':
            text = functions().getRB(fromText, operatingPath, setupParameters)
        elif setupParameters['RBmode'] == 'pykakasi':
            text = functions().getRB_withoutInternet(fromText)
        text = re.sub(r'<.*?>', '', text)
        m.write(text)  # 写入
        m.close()
        tk.messagebox.showinfo(title='message', message='Finish')
    elif RBforTXTsave == 'saveAsHTML':
        m = open('%s(RuBi).html' % Filename, encoding='utf-8', mode='w')
        text = u''
        if setupParameters['RBmode'] == 'hiragana':
            text = functions().getRB(fromText, operatingPath, setupParameters)
        elif setupParameters['RBmode'] == 'pykakasi':
            text = functions().getRB_withoutInternet(fromText)
        m.write(text)  # 写入
        m.close()
        tk.messagebox.showinfo(title='message', message='Finish')
    elif RBforTXTsave == 'saveAsDocx':
        Filename = (os.path.split(Filepath)[1])[0:len(os.path.split(Filepath)[1]) - 4]
        dirName = Filename + '(RuBi)-DOCX'
        dirPath = os.path.split(Filepath)[0] + os.sep + dirName
        # 复制docx样本
        os.chdir(operatingPath)
        if os.path.exists(dirPath):
            shutil.rmtree(dirPath)
            time.sleep(1)
            shutil.copytree("docxSample", dirPath)
        else:
            shutil.copytree("docxSample", dirPath)

        if setupParameters['RBmode'] == 'hiragana':
            text = functions().getRB(fromText, operatingPath, setupParameters)
        elif setupParameters['RBmode'] == 'pykakasi':
            text = functions().getRB_withoutInternet(fromText)
        Functions.makeDocx(text, Filepath, 1, setupParameters)
        tk.messagebox.showinfo(title='message', message='Finished')


def RBforDOCX():
    readSet()

    # 文件选择
    root = tk.Tk()
    root.withdraw()
    Filepath = filedialog.askopenfilename(filetypes=[('DOCX', '*.docx')])

    print('\n')
    print(Filepath)
    os.chdir(os.path.split(Filepath)[0])  # 改正工作目录
    Filename = (os.path.split(Filepath)[1])[0:len(os.path.split(Filepath)[1]) - 5]
    print(Filename)

    # 解压DOCX
    zfile = zipfile.ZipFile(Filepath, "r")
    zfile.extractall(os.path.split(Filepath)[0]+os.sep+Filename+'-RuBy')
    zfile.close()

    # 获取文本
    testPath = os.path.split(Filepath)[0]+os.sep+Filename+'-RuBy'+os.sep+'word'+os.sep+'document.xml'
    testFile = open(testPath, encoding='utf-8', mode='r')
    Otext = testFile.read()
    testFile.close()

    # 获取注音
    if setupParameters['RBmode'] == 'hiragana':
        text = functions().getRB(Otext, operatingPath, setupParameters, RBforDOCX=True)
    elif setupParameters['RBmode'] == 'pykakasi':
        text = functions().getRB_withoutInternet(Otext, RBforDOCX=True)

    # 整理格式
    text = re.sub('<br>', '', text)
    text = functions().RuByforDocxConformityToXML(text)
    open(testPath, encoding='utf-8', mode='w').write(text)

    # 打包DOCX
    file_list = []
    functions().ListAllFiles(dir=os.path.split(Filepath)[0]+os.sep+Filename+'-RuBy', file_list=file_list)
    os.chdir(os.path.split(Filepath)[0])
    with zipfile.ZipFile('%s-RuBy.docx' % Filename, 'w') as zipobj:
        os.chdir(os.path.split(Filepath)[0] + os.sep + Filename+'-RuBy')
        for file in file_list:
            zipobj.write(file)
        zipobj.close()

    tk.messagebox.showinfo(title='message', message='Finished')


def windowColse():
    window.destroy()
    sys.exit()


def main():
    print('\n' + os.path.abspath('.'))
    global operatingPath
    operatingPath = os.path.abspath(".")  # 记录程序路径
    readSet()
    global window
    window = tk.Tk()
    window.title(softName)

    if os.path.exists('backg.jpg'):  # 背景处理
        image1 = ImageTk.PhotoImage(Image.open('backg.jpg'))
        # get the image size
        w = image1.width()
        h = image1.height()
        # make the root window the size of the image
        window.geometry("%dx%d" % (w + 3, h + 17))
        lableB = tk.Label(window, image=image1, compound='top', text='Made by DrMelt', font=('Yu Gothic', 7))
        lableB.grid(column=1, row=0, columnspan=2, rowspan=1000)
        lable = tk.Label(window, text=softName, bg='#D6C0CC', font=('Yu Gothic', 12), width=28 * w // 282, height=1)
        # 说明： bg为背景，font为字体，width为长，height为高，这里的长和高是字符的长和高，比如height=2,就是标签有2个字符这么高
        lable.grid(row=0, column=1, columnspan=2)
    else:
        window.geometry("%dx%d" % (282 + 3, 400 + 17))
        lableMark = tk.Label(window, compound='top', text='Made by DrMelt', font=('Yu Gothic', 7), height=68)
        # lableB = tk.Label(window, text='Made by DrMelt', font=('Yu Gothic', 7), height=1)
        lableMark.grid(column=1, row=0, columnspan=2, rowspan=1000)
        # lableB.grid(column=1, row=999, columnspan=2)
        lable = tk.Label(window, text=softName, bg='#D6C0CC', font=('Yu Gothic', 12), width=28, height=1)
        # 说明： bg为背景，font为字体，width为长，height为高，这里的长和高是字符的长和高，比如height=2,就是标签有2个字符这么高
        lable.grid(row=0, column=1, columnspan=2)
    # 第一行
    setButton = tk.Button(window, bg='#D6C0CC', compound='top', text='Set', font=('Yu Gothic UI', 8), height=1, width=3,
                          command=setWindow)
    setButton.grid(sticky=tk.E, row=0, column=1, columnspan=2)

    window.resizable(0, 0)  # 冻结窗口大小
    # 第二行
    b1 = tk.Button(window, text='HTMLtoTXT', font=('Yu Gothic', 8), width=15, height=1,
                   command=HTMLtoTXT)  # lambda:  传参
    b1.grid(sticky=tk.E, row=1, column=1)

    b1 = tk.Button(window, text='HTMLtoDOCX', font=('Yu Gothic', 8), width=15, height=1,
                   command=HTMLtoDOCX)  # lambda:  传参
    b1.grid(sticky=tk.W, row=1, column=2)
    # 第三行
    b1_1 = tk.Button(window, text='EPUBtoDOCX', font=('Yu Gothic', 8), width=15, height=1,
                     command=EPUBtoDOCX)  # lambda:  传参
    b1_1.grid(sticky=tk.E, row=2, column=1)

    b1 = tk.Button(window, text='EPUBtoTXT', font=('Yu Gothic', 8), width=15, height=1,
                   command=EPUBtoTXT)  # lambda:  传参
    b1.grid(sticky=tk.W, row=2, column=2)
    # 第四行
    b2 = tk.Button(window, text='RBforTXT', font=('Yu Gothic', 8), width=15, height=1, command=RBforTXT)
    b2.grid(sticky=tk.E, row=3, column=1)

    global listbox
    listbox = tk.Listbox(window, bg='#C7F8FC', height=1, width=14)
    listbox.grid(sticky=tk.W, row=3, column=2)
    listbox.insert('end', 'saveAsDocx')
    listbox.insert('end', 'saveAsTXT')
    listbox.insert('end', 'saveAsHTML')

    # 第五行
    b1 = tk.Button(window, text='RBforDOCX', font=('Yu Gothic', 8), width=32, height=1,
                   command=RBforDOCX)  # lambda:  传参
    b1.grid(row=4, column=1, columnspan=2)

    # lable3 = tk.Label(window, text='Need Chrome', bg='#C7F8FC', fg='#A02A3F', font=('Yu Gothic', 8), width=13, height=1)
    # lable3.grid(row=3, column=1, columnspan=2)

    # lable4 = tk.Label(window, text='Made by DrMelt', bg='#C7F8FC', font=('Yu Gothic', 7), width=14, height=1)
    # lable4 = tk.Label(window, text='Made by DrMelt', compound='top', image=imageA, font=('Yu Gothic', 7))
    # lable4.grid(row=999, column=1, columnspan=2)
    window.protocol("WM_DELETE_WINDOW", windowColse)
    window.mainloop()
    # 注意，loop因为是循环的意思，window.mainloop就会让window不断的刷新，如果没有mainloop,就是一个静态的window,传入进去的值就不会有循环，mainloop就相当于一个很大的while循环，有个while，每点击一次就会更新一次，所以我们必须要有循环
    # 所有的窗口文件都必须有类似的mainloop函数，mainloop是窗口文件的关键的关键
    # window.destroy()


def setWindowClose():
    handle['setWindow'] = 0
    setwindow.destroy()


def setWindow():
    if handle["setWindow"] == 0:
        handle["setWindow"] = 1
        readSet()
        global setwindow
        setwindow = tk.Tk()  # 导入tkinter中的tk模块
        setwindow.title('Set')
        setwindow.geometry('200x50')
        setwindow.resizable(0, 0)
        # 窗口的宽x高
        # 当geometry函数的参数是上面这种两个加号风格的时候，就是调整窗口在屏幕上的位置，
        # 第1个加号是距离屏幕左边的宽，第2个加号是距离屏幕顶部的高。注意加号后面可以跟负数，这是一种隐藏窗口的方式
        RBmodeV = tk.StringVar()
        # RBmodeV.set(value=setupParameters['RBmode'])
        RBmode0 = tk.Radiobutton(setwindow, text='RBwith hiragana.jp', value='hiragana',
                                 variable=RBmodeV, command=lambda: writeSet('RBmode', 'hiragana'))
        RBmode0.pack()
        RBmode1 = tk.Radiobutton(setwindow, text='RBwith Pykakasi', value='pykakasi',
                                 variable=RBmodeV, command=lambda: writeSet('RBmode', 'pykakasi'))
        RBmode1.pack()
        if setupParameters['RBmode'] == 'hiragana':
            RBmode0.select()
        elif setupParameters['RBmode'] == 'pykakasi':
            RBmode1.select()
        setwindow.protocol("WM_DELETE_WINDOW", setWindowClose)
        setwindow.update()
        setwindow.mainloop()


def readSet():
    os.chdir(operatingPath)
    if os.path.exists('configuration'):
        configuration = open('configuration', encoding='utf-8', mode='r')
        for item in configuration.readlines():
            line = re.search(r'(.*?) = (.*)', item)
            if line is not None:
                try:
                    setupParameters[line.group(1)] = line.group(2)
                except:
                    print('设置错误')
        configuration.close()
    else:
        configuration = open('configuration', encoding='utf-8', mode='w')
        configuration.write(setting_default)
        configuration.close()
        readSet()


def writeSet(property, value):  # 设置 属性，值
    lines = []
    if os.path.exists('configuration'):
        configuration = open('configuration', encoding='utf-8', mode='r')
        for item in configuration.readlines():
            line = re.sub(r'(%s = )(.*)' % property, r'\1%s' % value, item)
            lines.append(line)
        configuration.close()
        configuration = open('configuration', encoding='utf-8', mode='w')
        for item in lines:
            configuration.write(item)
        configuration.close()
        readSet()
    else:
        configuration = open('configuration', encoding='utf-8', mode='w')
        configuration.write(setting_default)
        configuration.close()
        writeSet(property, value)


if __name__ == "__main__":  # 当程序执行时 调用-->入口
    main()
