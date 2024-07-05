# -*- coding: utf-8 -*-
# @Time : 2022/3/25 11:52
# @Author : DrMelt
# @File : TextTools

import time
import shutil
import chardet
from selenium import webdriver
import re
import os
import pykakasi  # 注音库
import zipfile
from PIL import Image as Im
import sys
from tkinter import filedialog
import time
import epub
from bs4 import BeautifulSoup
from Functions.functions import functions
import tkinter as tk
import tkinter.messagebox
from PIL import Image, ImageTk


def main():

    root = tk.Tk()
    root.withdraw()
    Filepath = filedialog.askopenfilename()
    fileData = open(Filepath, mode='rb').read()
    encoded = chardet.detect(fileData)['encoding']
    print(encoded)
    text = open(Filepath, encoding=encoded, mode='r')
    Filename = (os.path.split(Filepath)[1])[0:len(os.path.split(Filepath)[1]) - 4]
    tFileName = Filename + '-DeRuby.txt'
    tFilepath = os.path.split(Filepath)[0] + os.sep + tFileName
    print(tFilepath)
    target = open(tFilepath, encoding='CP932', mode='w')
    Text = ''
    for i in text.readlines():
        Text = Text + i
    Text = re.sub('《(.*?)》', r'', Text)
    target.write(Text)
    target.close()
    print('Finished')


if __name__ == "__main__":  # 当程序执行时 调用-->入口
    main()
