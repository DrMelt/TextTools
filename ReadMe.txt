功能简介：
EPUBtoTXT：epub格式转换为txt。解决calibre转换中注音下移问题（直接把注音删了）
EPUBtoDOCX：epub格式转换为docx（不删注音）
RBforTXT：为txt格式日语文本注假名（记得选择保存方式（滚轮可以上下划））
要想注音，请使用RBforTXT。先使用EPUBtoTXT，再使用RBforTXT（可以保存图片）


注意事项：
保存图片转换出错请在configuration中关掉（True-->False）

提示 拒绝访问 和 目录不是空的 报错只要再来一次就好
Permission denied 报错请检查是否打开文件

所有文件（除了这个txt）需要在同级目录
程序读取backg.jpg（282x400 etc.）作为背景

可能出现有一些进程可能无法关闭，请用任务管理器关闭


现在不推荐使用hiragana.jp注音
RBforTXT使用hiragana.jp注音需要Chrome浏览器
！！！！！如果chromedriver.exe有问题请到
http://chromedriver.storage.googleapis.com/index.html
下载对应Chrome版本的chromedriver.exe并替换

对于Chrome版本 115 及更高版本 目前没有支持 请参考
https://developer.chrome.com/docs/chromedriver/downloads/version-selection?hl=zh-cn
https://googlechromelabs.github.io/chrome-for-testing/

肯定存在大量bug……


Author:DrMelt

--2024.7.5 v1.5--
修复pykakasi注音数据库，仍然只能用pykakasi注音，速度较慢

--2022.3.27 v1.4--
docx注音，还有很多问题，只能用pykakasi注音

--2021.10.16 v1.3--
我也不知道之前更新了什么
还是使用pyinstaller打包吧

--2021.7.23--
.docx保存图像（可在configuration中关掉）

--2021.7.17--
优化
用Nuitka打包以降低运行速度

---2021.7.11 v1.2---
增加RBforTXT保存为.docx格式功能
增加RBforTXT使用pykakasi注音功能
增加设置功能（configuration为配置文件）


---2021.6.21 v1.1---
修复了根本无法使用的bug



