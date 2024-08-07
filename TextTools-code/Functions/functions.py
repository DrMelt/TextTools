# -*- coding: utf-8 -*-
# @Time : 2021/6/7 14:48
# @Author : DrMelt
# @File : TextTools
# @Software : PyCharm

import time
import shutil
import chardet
from selenium import webdriver
import re
import os
import pykakasi  # 注音库
import zipfile
from PIL import Image as Im


class functions(object):
    def setWord(self, setupParameters={}):
        self.documentDefault0 = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
    xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"
    xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex"
    xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex"
    xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex"
    xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex"
    xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex"
    xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex"
    xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex"
    xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink"
    xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d"
    xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
    xmlns:v="urn:schemas-microsoft-com:vml"
    xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:w10="urn:schemas-microsoft-com:office:word"
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
    xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
    xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"
    xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"
    xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml"
    xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash"
    xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex"
    xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
    xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
    xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
    xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14">
    <w:body>'''
        self.documentDefault1 = '''
        <w:sectPr w:rsidR="002F5773" w:rsidRPr="002F5773">
            <w:pgSz w:w="11906" w:h="16838"/>
            <w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/>
            <w:cols w:space="425"/>
            <w:docGrid w:type="lines" w:linePitch="312"/>
        </w:sectPr>
    </w:body>
</w:document>
'''
        self.DOCXmark_TXT0 = '''
            <w:r>
                <w:rPr>
                    <w:rFonts w:eastAsia=%s w:hint="eastAsia"/>
                    <w:lang w:eastAsia="ja-JP"/>
                </w:rPr>
                <w:t>''' % setupParameters['w:rFonts']
        self.DOCXmark_TXT1 = '''</w:t>
            </w:r>'''
        # 0 + RB + 1 + Word + 2
        self.DOCXmark_RB0 = '''
            <w:pPr>
                <w:rPr>
                    <w:rFonts w:eastAsia=%s/>
                    <w:lang w:eastAsia="ja-JP"/>
                </w:rPr>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:rFonts w:eastAsia=%s/>
                    <w:lang w:eastAsia="ja-JP"/>
                </w:rPr>
                <w:ruby>
                    <w:rubyPr>
                        <w:rubyAlign w:val="distributeSpace"/>
                        <w:hps w:val="%s"/>
                        <w:hpsRaise w:val="%s"/>
                        <w:hpsBaseText w:val="%s"/>
                        <w:lid w:val="ja-JP"/>
                    </w:rubyPr>
                    <w:rt>
                        <w:r>
                            <w:rPr>
                                <w:rFonts w:ascii=%s w:eastAsia=%s w:hAnsi=%s/>
                                <w:sz w:val="10"/>
                                <w:lang w:eastAsia="ja-JP"/>
                            </w:rPr>
                            <w:t>''' % (
            setupParameters['w:rFonts'], setupParameters['w:rFonts'], setupParameters['w:hps'],
            setupParameters['w:hpsRaise'], setupParameters['w:hpsBaseText'], setupParameters['w:rFonts'],
            setupParameters['w:rFonts'], setupParameters['w:rFonts'])
        self.DOCXmark_RB1 = '''</w:t>
                        </w:r>
                    </w:rt>
                    <w:rubyBase>
                        <w:r>
                            <w:rPr>
                                <w:rFonts w:eastAsia=%s/>
                                <w:lang w:eastAsia="ja-JP"/>
                            </w:rPr>
                            <w:t>''' % setupParameters['w:rFonts']
        self.DOCXmark_RB2 = '''</w:t>
                        </w:r>
                    </w:rubyBase>
                </w:ruby>
            </w:r>'''
        self.PerPx = int(9525 / float(setupParameters['Imagescale']))

    PerPx = 9525 // 2
    documentxml0 = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
    <Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
    <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" Target="webSettings.xml"/>
    <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" Target="fontTable.xml"/>
    <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/>
    <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
'''
    documentxml1 = '''</Relationships>'''
    DOCXmark_Img0 = '''
            <w:p>
                <w:r>
                <w:rPr>
                    <w:noProof/>
                </w:rPr>
                <w:drawing>
                    <wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">
                        <wp:simplePos x="0" y="0"/>
                        <wp:positionH relativeFrom="column">
                            <wp:posOffset>0</wp:posOffset>
                        </wp:positionH>
                        <wp:positionV relativeFrom="paragraph">
                            <wp:posOffset>0</wp:posOffset>
                        </wp:positionV>
                        <wp:extent cx="'''
    DOCXmark_Img1 = '''"/>
                        <wp:effectExtent l="0" t="0" r="0" b="0"/>
                        <wp:wrapTopAndBottom/>
                        <wp:docPr id="1" name="图片 1"/>
                        <wp:cNvGraphicFramePr>
                            <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
                        </wp:cNvGraphicFramePr>
                        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                    <pic:nvPicPr>
                                        <pic:cNvPr id="1" name="Picture 1"/>
                                        <pic:cNvPicPr>
                                            <a:picLocks noChangeAspect="1" noChangeArrowheads="1"/>
                                        </pic:cNvPicPr>
                                    </pic:nvPicPr>
                                    <pic:blipFill>
                                        <a:blip r:embed="IrId'''
    DOCXmark_Img2 = '''" cstate="print">
                                            <a:extLst>
                                                <a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
                                                    <a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>
                                                </a:ext>
                                            </a:extLst>
                                        </a:blip>
                                        <a:srcRect/>
                                        <a:stretch>
                                            <a:fillRect/>
                                        </a:stretch>
                                    </pic:blipFill>
                                    <pic:spPr bwMode="auto">
                                        <a:xfrm>
                                            <a:off x="0" y="0"/>
                                            <a:ext cx="'''
    DOCXmark_Img3 = '''"/>
                                        </a:xfrm>
                                        <a:prstGeom prst="rect">
                                            <a:avLst/>
                                        </a:prstGeom>
                                        <a:noFill/>
                                        <a:ln>
                                            <a:noFill/>
                                        </a:ln>
                                    </pic:spPr>
                                </pic:pic>
                            </a:graphicData>
                        </a:graphic>
                    </wp:anchor>
                </w:drawing>
            </w:r>
        </w:p>'''
    documentDefault0 = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
    xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"
    xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex"
    xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex"
    xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex"
    xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex"
    xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex"
    xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex"
    xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex"
    xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink"
    xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d"
    xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
    xmlns:v="urn:schemas-microsoft-com:vml"
    xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:w10="urn:schemas-microsoft-com:office:word"
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
    xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
    xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"
    xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"
    xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml"
    xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash"
    xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex"
    xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
    xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
    xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
    xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14">
    <w:body>'''
    documentDefault1 = '''
        <w:sectPr w:rsidR="002F5773" w:rsidRPr="002F5773">
            <w:pgSz w:w="11906" w:h="16838"/>
            <w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/>
            <w:cols w:space="425"/>
            <w:docGrid w:type="lines" w:linePitch="312"/>
        </w:sectPr>
    </w:body>
</w:document>
'''
    DOCXmark_TXT0 = '''
            <w:r>
                <w:t>'''
    DOCXmark_TXT1 = '''</w:t>
            </w:r>'''
    # 0 + RB + 1 + Word + 2
    DOCXmark_RB0 = '''
            <w:pPr>
                <w:rPr>
                    <w:rFonts w:eastAsia="Yu Mincho"/>
                    <w:lang w:eastAsia="ja-JP"/>
                </w:rPr>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:rFonts w:eastAsia="Yu Mincho"/>
                    <w:lang w:eastAsia="ja-JP"/>
                </w:rPr>
                <w:ruby>
                    <w:rubyPr>
                        <w:rubyAlign w:val="distributeSpace"/>
                        <w:hps w:val="10"/>
                        <w:hpsRaise w:val="18"/>
                        <w:hpsBaseText w:val="21"/>
                        <w:lid w:val="ja-JP"/>
                    </w:rubyPr>
                    <w:rt>
                        <w:r>
                            <w:rPr>
                                <w:rFonts w:ascii="Yu Mincho" w:eastAsia="Yu Mincho" w:hAnsi="Yu Mincho"/>
                                <w:sz w:val="10"/>
                                <w:lang w:eastAsia="ja-JP"/>
                            </w:rPr>
                            <w:t>'''
    DOCXmark_RB1 = '''</w:t>
                        </w:r>
                    </w:rt>
                    <w:rubyBase>
                        <w:r>
                            <w:rPr>
                                <w:rFonts w:eastAsia="Yu Mincho"/>
                                <w:lang w:eastAsia="ja-JP"/>
                            </w:rPr>
                            <w:t>'''
    DOCXmark_RB2 = '''</w:t>
                        </w:r>
                    </w:rubyBase>
                </w:ruby>
            </w:r>'''
    DOCXConformity_RB0 = '''<w:r>
                <w:rPr><w:rFonts w:eastAsia="Yu Mincho"/><w:lang w:eastAsia="ja-JP"/></w:rPr>
                <w:ruby>
                    <w:rubyPr>
                        <w:rubyAlign w:val="distributeSpace"/>
                        <w:hps w:val="10"/>
                        <w:hpsRaise w:val="18"/>
                        <w:hpsBaseText w:val="21"/>
                        <w:lid w:val="ja-JP"/>
                    </w:rubyPr>
                    <w:rt>
                        <w:r>
                            <w:rPr>
                                <w:rFonts w:ascii="Yu Mincho" w:eastAsia="Yu Mincho" w:hAnsi="Yu Mincho"/>
                                <w:sz w:val="10"/>
                                <w:lang w:eastAsia="ja-JP"/>
                            </w:rPr>
                            <w:t>'''
    DOCXConformity_RB1 = '''</w:t></w:r></w:rt><w:rubyBase><w:r><w:rPr><w:rFonts w:eastAsia="Yu Mincho"/><w:lang w:eastAsia="ja-JP"/></w:rPr><w:t>'''

    def __init__(self):
        pass

    def DecToHex(self, dec, wei=8):
        return self.buwei(str(hex(dec))[2:].upper(), wei)

    def buwei(self, xuhao, wei):
        if len(xuhao) >= wei:
            return xuhao
        else:
            xuhao = "0" + xuhao
            return self.buwei(xuhao, wei)

    def makeDocx(self, Text, Filepath, mode=0, setupParameters={}):
        self.setWord(setupParameters)  # 应用设置
        if mode == 1:
            Filename = (os.path.split(Filepath)[1])[
                0:len(os.path.split(Filepath)[1]) - 4]
            dirName = Filename + '(RuBi)-DOCX'
        elif mode == 0:
            Filename = (os.path.split(Filepath)[1])[
                0:len(os.path.split(Filepath)[1]) - 5]
            dirName = Filename + '-DOCX'
        dirPath = os.path.split(Filepath)[0] + os.sep + dirName

        if setupParameters['ConvertImage'] == 'True':
            # 读取图像地址
            Images = {}
            ImagesZipin = []
            ImagesRelsin = {}
            count = 1
            for i in re.finditer(r'\[image](.*?)\[/image]', Text, re.M):
                if Images.get(i.group(1), -1) == -1:  # 不存在这个key时添加，防止重复
                    Images[i.group(1)] = count
                    count = count + 1
            # 复制图像
            os.chdir(os.path.split(Filepath)[0])
            for i in list(Images.keys()):
                ImagePath = 'media/image' + \
                    str(Images[i]) + re.search(r'\..*',
                                               os.path.split(i)[1]).group()
                ImagesZipin.append('word/' + ImagePath)
                ImagesRelsin[ImagePath] = Images[i]
                i = re.sub(
                    r'<ruby><rb>(.*?)</rb><rp>\(</rp><rt>(.*?)</rt><rp>\)</rp></ruby>', r'\1', i)
                shutil.copyfile(os.path.abspath(
                    '.' + i), os.path.abspath('./' + dirName + '/word/' + ImagePath))
            # 注册图像
            os.chdir(dirPath + '/word/_rels')
            rels = open('document.xml.rels', mode='w', encoding='utf-8')
            rels.write(self.documentxml0)
            for item in list(ImagesRelsin.keys()):
                rels.write('    <Relationship Id="IrId' + str(ImagesRelsin[
                    item]) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="' + item + '"/>\n')
            rels.write(self.documentxml1)
            rels.close()

        # 写入内容预备
        os.chdir(dirPath + '/word')
        docx = open('document.xml', encoding='utf-8', mode='w')
        Text = '<br>' + Text + '<br>'
        Text = re.sub('\n', '', Text)  # 删除回车
        Text = re.sub(r'［＃挿絵（(.*?)）入る］',
                      r'[image]/\1[/image]', Text, re.M)  # 转换青空文库标记
        Text = re.sub(r'〝', r'『', Text)  # 转换青空文库『』
        Text = re.sub(r'〟', r'』', Text)  # 转换青空文库『』

        if setupParameters['ConvertImage'] == 'True':
            # 图像标记处理
            os.chdir(os.path.split(Filepath)[0])

            ImgS = re.search(r'\[image](.*?)\[/image]<br>', Text, re.M)
            while ImgS is not None:
                imgPath = ImgS.group(1)
                img = Im.open('.' + imgPath)
                w = img.width  # 图片的宽
                h = img.height  # 图片的高
                Text = Text[:ImgS.start()] + '%s%s" cy="%s%s%s%s%s" cy="%s%s' % (
                    self.DOCXmark_Img0, w * self.PerPx, h * self.PerPx, self.DOCXmark_Img1,
                    Images[Text[ImgS.start(1):ImgS.end(
                        1)]], self.DOCXmark_Img2, w * self.PerPx, h * self.PerPx,
                    self.DOCXmark_Img3) + Text[ImgS.end():]
                ImgS = re.search(r'\[image](.*?)\[/image]<br>', Text, re.M)

            # ImgS = re.search(r'［＃.*?（(.*?)）.*?］', Text, re.M)
            # while ImgS is not None:
            #     imgPath = ImgS.group(1)
            #     imgPath = '/' + imgPath
            #     img = Im.open('.' + imgPath)
            #     w = img.width  # 图片的宽
            #     h = img.height  # 图片的高
            #     Text = Text[:ImgS.start()] + '%s%s" cy="%s%s%s%s%s" cy="%s%s' % (
            #         self.DOCXmark_Img0, w * self.PerPx, h * self.PerPx, self.DOCXmark_Img1,
            #         Images[Text[ImgS.start(1):ImgS.end(1)]], self.DOCXmark_Img2, w * self.PerPx, h * self.PerPx,
            #         self.DOCXmark_Img3) + Text[ImgS.end():]
            #     ImgS = re.search(r'［＃.*?（(.*?)）.*?］', Text, re.M)

        Text = self.documentDefault0 + Text + self.documentDefault1  # 拼接头尾
        # Text = re.sub(r'<Ruby><Rb>(.*?)</Rb><Rp>[(（]</Rp><Rt>(.*?)</Rt><Rp>[)）]</Rp></Ruby>',
        #               r'%s\2%s\1%s' % (self.DOCXmark_RB0, self.DOCXmark_RB1, self.DOCXmark_RB2), Text)  # 注音处理1
        Text = re.sub(r'<ruby><rb>(.*?)</rb><rp>[(（]</rp><rt>(.*?)</rt><rp>[)）]</rp></ruby>',
                      r'%s\2%s\1%s' % (
                          self.DOCXmark_RB0, self.DOCXmark_RB1, self.DOCXmark_RB2), Text,
                      flags=re.I)  # 注音处理2
        Text = re.sub(r'(</w:r>)(.*?)<br>', r'\1%s\2%s%s' % (self.DOCXmark_TXT0, self.DOCXmark_TXT1, '\n       </w:p>'),
                      Text)  # </w:r><br>  行结尾\n处理
        Text = re.sub('<br>', '\n<w:p/>', Text)
        Text = re.sub('(</w:r>)(.*?)(\n            <w:pPr>)', r'\1%s\2%s\3' % (self.DOCXmark_TXT0, self.DOCXmark_TXT1),
                      Text)  # </w:r>の\n<w:pPr>  行内部处理
        # Text = re.sub('</w:p>(.+)', r'</w:p><w:p>%s\1%s</w:p>' % (self.DOCXmark_TXT0, self.DOCXmark_TXT1), Text)  # 无注音单行处理1         </w:p>「はあ……、そうなんですか」
        Text = re.sub('</w:p>(.+)', r'</w:p><w:p>%s\1%s' % (self.DOCXmark_TXT0, self.DOCXmark_TXT1),
                      Text)  # 无注音单行处理1         </w:p>「はあ……、そうなんですか」
        Text = re.sub('<w:p/>\n            <w:pPr>',
                      '<w:p/>\n\n        <w:p>\n            <w:pPr>', Text)  # 行开头处理1
        Text = re.sub('(<w:p/>)(.*?)(\n            <w:pPr>)',
                      r'\1<w:p>%s\2%s\3' % (self.DOCXmark_TXT0, self.DOCXmark_TXT1), Text)  # 行开头处理2
        Text = re.sub('</w:p>\n            <w:pPr>',
                      '</w:p>\n\n        <w:p>\n            <w:pPr>', Text)  # 行开头处理3
        Text = re.sub('<w:p/>(.+)\n<w:p/>',
                      r'<w:p/><w:p>%s\1%s%s' % (self.DOCXmark_TXT0,
                                                self.DOCXmark_TXT1, '\n        </w:p>'),
                      Text)  # 无注音单行处理2         <w:p/>○\n<w:p/>
        Text = re.sub('</w:p>(.+)\n<w:p/>',
                      r'</w:p><w:p>%s\1%s%s' % (self.DOCXmark_TXT0,
                                                self.DOCXmark_TXT1, '\n        </w:p>'),
                      Text)  # 无注音单行处理3
        Text = re.sub('<w:p/>(.+)\n            <w:p>',
                      r'<w:p/><w:p>%s\1%s%s<w:p>' % (
                          self.DOCXmark_TXT0, self.DOCXmark_TXT1, '\n        </w:p>'),
                      Text)  # 无注音单行处理4
        if setupParameters['ConvertImage'] == 'True':
            Text = re.sub('</w:p>(.+)\n            <w:p>',
                          r'</w:p><w:p>%s\1%s%s<w:p>' % (
                              self.DOCXmark_TXT0, self.DOCXmark_TXT1, '\n        </w:p>'),
                          Text)  # 无注音单行处理5
            Text = re.sub('</w:r>(.*?)\n            <w:p>',
                          r'</w:r>%s\1%s%s<w:p>' % (self.DOCXmark_TXT0,
                                                    self.DOCXmark_TXT1, '\n        </w:p>'),
                          Text)  # 无注音单行处理6

        Text = re.sub('</w:r>\n<w:p/>', r'</w:r></w:p>',
                      Text)  # 末处理 </w:r>\n<w:p/>
        Text = re.sub('<w:body>\n<w:p/>', '<w:body>', Text)  # 修剪开头空格

        docx.write(Text)
        docx.close()

        # 压缩为docx
        os.chdir(os.path.split(Filepath)[0])
        file_list = ['[Content_Types].xml', 'word/document.xml', 'word/fontTable.xml', 'word/settings.xml',
                     'word/styles.xml', 'word/webSettings.xml', 'word/_rels/document.xml.rels', 'word/theme/theme1.xml',
                     'docProps/core.xml', 'docProps/app.xml', '_rels/.rels']
        if mode == 1:
            with zipfile.ZipFile('%s(RuBi).docx' % Filename, 'w') as zipobj:
                os.chdir(os.path.split(Filepath)[0] + os.sep + dirName)
                for file in file_list:
                    zipobj.write(file)
                if setupParameters['ConvertImage'] == 'True':
                    for Image in ImagesZipin:
                        zipobj.write(Image)
            zipobj.close()
        elif mode == 0:
            with zipfile.ZipFile('%s.docx' % Filename, 'w') as zipobj:
                os.chdir(os.path.split(Filepath)[0] + os.sep + dirName)
                for file in file_list:
                    zipobj.write(file)
                if setupParameters['ConvertImage'] == 'True':
                    for Image in ImagesZipin:
                        zipobj.write(Image)
            zipobj.close()

    def RuByforDocxConformityToXML(self, text):
        self.DOCXConformity_RB0 = re.sub(
            '> *?<', '><', re.sub('\n', '', self.DOCXConformity_RB0))
        RB0 = self.DOCXConformity_RB0
        RB1 = self.DOCXConformity_RB1
        RB2 = re.sub('> *?<', '><', re.sub('\n', '', self.DOCXmark_RB2))
        TXT0 = re.sub(' *?<', '<', re.sub('> *?<', '><',
                      re.sub('\n', '', self.DOCXmark_TXT0)))
        TXT1 = '</w:t></w:r>'

        text = re.sub(r'<w:body>', '<w:body>\n', text)
        textLines = self.segmentationLine(text)

        txt = ''
        for line in textLines:
            temp = re.search(
                r'(<w:r[ >].*?<w:t.*?>)(.*?)(</w:t></w:r>)', line, flags=re.M)
            if temp is not None:
                TXT0 = temp.group(1)  # 获取字形
                # print(TXT0)
                # font = re.search(r'w:eastAsia="(.*?)"', TXT0)
                # if font is not None:
                #     font1 = font.group(1)
                #     RB0 = re.sub('Yu Mincho', font1, self.DOCXConformity_RB0)
                #     RB1 = re.sub('Yu Mincho', font1, self.DOCXConformity_RB1)
                font = re.search(r'(<w:rPr>.*?</w:rPr>)', TXT0)
                if font is not None:
                    font2 = font.group(1)
                    RB0 = re.sub('<w:rPr><w:rFonts w:eastAsia="Yu Mincho"/><w:lang w:eastAsia="ja-JP"/></w:rPr>', font2,
                                 self.DOCXConformity_RB0)
                    RB1 = re.sub('<w:rPr><w:rFonts w:eastAsia="Yu Mincho"/><w:lang w:eastAsia="ja-JP"/></w:rPr>', font2,
                                 self.DOCXConformity_RB1)

                line = re.sub(r'(<w:t[^<>]*?>)([^<>]+?)<ruby><rb>',
                              r'\1\2%s<ruby><rb>' % TXT1, line, flags=re.M)
                line = re.sub(r'</rp></ruby>([^<>]+?)<ruby><rb>', r'</rp></ruby>%s\1%s<ruby><rb>' % (TXT0, TXT1), line,
                              flags=re.M)
                line = re.sub(r'</rp></ruby>([^<>]+?)</w:t></w:r>', r'</rp></ruby>%s\1%s' % (TXT0, TXT1), line,
                              flags=re.M)
                line = re.sub(r'</rp></ruby></w:t></w:r>',
                              r'</rp></ruby>', line, flags=re.M)
                line = re.sub(
                    r'(<w:r[ >].*?<w:t[^<>]*?>)<ruby><rb>', r'<ruby><rb>', line, flags=re.M)

                line = re.sub(r'<ruby><rb>(.*?)</rb><rp>\(</rp><rt>(.*?)</rt><rp>\)</rp></ruby>',
                              RB0 + r'\2' + RB1 + r'\1' + RB2, line, flags=re.M)

            # line = re.sub(r'(<w:r[ >].*?<w:t.*?>)(.*?<ruby><rb>.*?)</w:t></w:r>', r'\2', line, flags=re.M)
            # line = re.sub(r'<w:r[ >].*?<w:t.*?>(.*?)</w:t></w:r>(.*?</rp></ruby>)', r'\1\2', line, flags=re.M)
            # line = re.sub(r'</w:t></w:r>.*?<w:t>(.*?<ruby><rb>)', r'\1', line, flags=re.M)
            # line = re.sub(r'</w:pPr>([^<>]+?)<ruby><rb>', r'</w:pPr>%s\1%s<ruby><rb>' % (TXT0, TXT1), line, flags=re.M)
            # line = re.sub(r'</rp></ruby>([^<>]+?)<ruby><rb>', r'</rp></ruby>%s\1%s<ruby><rb>' % (TXT0, TXT1), line,
            #               flags=re.M)
            # line = re.sub(r'</rp></ruby>([^<>]+?)</w:p>', r'</rp></ruby>%s\1%s</w:p>' % (TXT0, TXT1), line, flags=re.M)
            # line = re.sub(r'(^[^<>]+?)<ruby><rb>', r'%s\1%s<ruby><rb>' % (TXT0, TXT1), line, flags=re.M)

            # line = re.sub(r'<ruby><rb>(.*?)</rb><rp>\(</rp><rt>(.*?)</rt><rp>\)</rp></ruby>', RB0 + r'\2' + RB1 + r'\1' + RB2, line, flags=re.M)
            txt = txt + line
        text = txt
        text = re.sub('\n\n', '\n', text)
        return text

    def readFile(self, site):
        fileData = open(site, mode='rb').read()
        encoded = chardet.detect(fileData)['encoding']
        print(encoded)

        f = open(site, encoding=encoded, mode='r')
        content = f.readlines()  # 列表，全部读完
        sumT = ''
        for temp in content:
            sumT = sumT + temp
        f.close
        return sumT

    def ListAllFiles(self, dir, filePath='', file_list=[]):
        for i in os.listdir(dir):
            os.chdir(dir)
            if os.path.isdir(dir + os.sep + i):
                if filePath is '':
                    self.ListAllFiles(dir=dir + os.sep + i,
                                      filePath=i, file_list=file_list)
                else:
                    self.ListAllFiles(
                        dir=dir + os.sep + i, filePath=filePath + os.sep + i, file_list=file_list)
            else:
                if filePath is '':
                    file_list.append(i)
                else:
                    file_list.append(filePath + os.sep + i)

    def findSubFiles(self, dir2, rootSite, subFiles=[], content='', flag=False):
        if flag is not True:  # 查找"container.xml"
            for i in os.listdir(dir2):
                os.chdir(dir2)
                if os.path.isdir(dir2 + os.sep + i):
                    self.findSubFiles(dir2=dir2 + os.sep + i, rootSite=rootSite, subFiles=subFiles, content=content,
                                      flag=flag)
                else:
                    if str(i).lower() == "container.xml":
                        m = open('container.xml', mode='r', encoding='utf-8')
                        for item in m.readlines():
                            item2 = (re.search('full-path="(.*?)"', item))
                            if item2 is not None:
                                content = rootSite + os.sep + \
                                    ((re.search('full-path="(.*?)"', item)
                                      ).groups())[0]
                                # 切换地址到full-path=目录
                                os.chdir(os.path.split(content)[0])
                        flag = True  # 已找到
                        break

        if flag is True:  # 查找html文件地址并写入subFiles数组
            contentM = open(content, encoding='utf-8', mode='r')
            contentT = contentM.readlines()
            contentM.close()
            for item in contentT:
                i = re.search('<item(.*?)/>', str(item))
                if i is not None:  # <item href="toc.ncx" id="ncx" media-type="application/x-dtbncx+xml" />
                    i = re.search('href="(.*?)"', str((i.groups())[0]))
                    if i is not None:
                        i = ((re.search('href="(.*?)"', item)).groups())[0]
                        if str(i[-6:]).lower() == ".xhtml" or str(i[-5:]).lower() == ".html":
                            subFiles.append(os.path.abspath(
                                '.') + os.sep + str(i))
        return content

    def findXhtml(self, dir1, subFiles=[]):
        for i in os.listdir(dir1):
            os.chdir(dir1)
            if os.path.isdir(dir1 + os.sep + i):
                self.findXhtml(dir1 + os.sep + i, subFiles)  # self 实例地址不变
            else:
                if str(i[-6:]).lower() == ".xhtml" or str(i[-5:]).lower() == ".html":
                    subFiles.append(os.path.abspath(i))
        return subFiles

    def getRB_with_pykakasi(self, Text, RBforDOCX=False):
        kks = pykakasi.kakasi()
        Text = re.sub('〇', '◯', Text)
        segmented = functions().segmentationLine(Text)  # 分割

        progressSum = len(segmented)
        oneline_percent = 1.0 / progressSum
        processed_percent = 0
        process_print_target = 0.1

        text_out = u''
        for text in segmented:
            result = kks.convert(text)
            # <ruby><rb>著者</rb><rp>(</rp><rt>ちょしゃ</rt><rp>)</rp></ruby>
            for item in result:
                if not (item['hira'] == item['orig']) and item['hira'] != '':
                    # hira = re.sub('\n', '', item['hira'])
                    hira = item['hira']
                    text_out = text_out + \
                        ('<ruby><rb>' + item['orig'] +
                         '</rb><rp>(</rp><rt>' + hira + '</rt><rp>)</rp></ruby>')
                else:
                    text_out = text_out + item['orig']

            processed_percent += oneline_percent
            if processed_percent >= process_print_target:
                print(str(int(processed_percent * 100)) + r'%')
                process_print_target += 0.1

        text_out = self.RBrepair(text_out, RBforDOCX)

        return text_out

    def getRB(self, Text, operatingPath, setupParameters, RBforDOCX=False):
        Text = re.sub('〇', '◯', Text)
        segmented = functions().segmentation(Text)  # 分割
        progressSum = len(segmented)
        progress = 0
        txt = u''
        os.chdir(operatingPath)  # 改正工作目录
        wd = webdriver.Chrome("chromedriver.exe")  # 创建对象
        wd.implicitly_wait(0.5)
        wd.get("https://hiragana.jp/reading/")  # 打开网址
        # 登入
        element = wd.find_element_by_name('id')
        element.send_keys('%s' % setupParameters['Passworld'])
        element = wd.find_element_by_name('pass')
        element.send_keys('%s' % setupParameters['login'])
        element = wd.find_element_by_name('login')
        element.click()

        for item in segmented:
            wd.switch_to.default_content()
            wd.switch_to.frame('query')
            # element = wd.find_element_by_css_selector('textarea')
            # element.send_keys(item)

            # 使用JS输入
            # js = r'document.getElementsByTagName("textarea")[0].value="%s"' % item
            # wd.execute_script(r'document.getElementsByTagName("textarea")[0].value="%s"' % item)
            wd.execute_script('document.getElementsByTagName("textarea")[0].value=arguments[0]',
                              item)  # arguments[0]解决换行问题

            # element = wd.find_element_by_css_selector('[type="checkbox"]')
            # element.click()
            element = wd.find_element_by_css_selector('[type="submit"]')
            element.click()
            time.sleep(1)
            wd.switch_to.default_content()  # 返回主html
            wd.switch_to.frame('result')
            element = wd.find_element_by_tag_name('body')
            # for i in element:
            #     print(i.get_attribute("innerHTML")+'\n')
            txt = txt + element.get_attribute("innerHTML")
            wd.switch_to.default_content()
            wd.switch_to.frame('query')
            element = wd.find_element_by_css_selector('[type="reset"]')
            element.click()
            progress = progress + 1
            print(str(int(progress / progressSum * 100)) + r'%')
        wd.close()
        os.system("taskkill /f /im chromedriver.exe")
        txt = re.sub('&nbsp;', ' ', txt)

        txt = self.RBrepair(txt, RBforDOCX)

        return txt

    def RBrepair(self, txt, RBforDOCX):
        txt = re.sub('\n', '<br>\n', txt)
        if not RBforDOCX:
            txt = re.sub(r'(［＃.*?］)', r'\1' + '\n', txt)  # 防止两个标志之间的注音被消除

            testLines = self.segmentationLine(txt)
            test = ''
            for line in testLines:
                temp = re.search(
                    r'［＃.*?(</rb><rp>\(</rp><rt>.*?</rt><rp>\)</rp></ruby>).*?］', line, re.M)
                while temp is not None:
                    line = line[:temp.start(1)] + line[temp.end(1):]
                    temp = re.search(
                        r'［＃.*?(</rb><rp>\(</rp><rt>.*?</rt><rp>\)</rp></ruby>).*?］', line, re.M)
                temp = re.search(r'［＃.*?(<ruby><rb>).*?］', line, re.M)
                while temp is not None:
                    line = line[:temp.start(1)] + line[temp.end(1):]
                    temp = re.search(r'［＃.*?(<ruby><rb>).*?］', line, re.M)

                temp = re.search(
                    r'\[image].*?(</rb><rp>\(</rp><rt>.*?</rt><rp>\)</rp></ruby>).*?\[/image]', line, re.M)
                while temp is not None:
                    line = line[:temp.start(1)] + line[temp.end(1):]
                    temp = re.search(r'\[image].*?(</rb><rp>\(</rp><rt>.*?</rt><rp>\)</rp></ruby>).*?\[/image]', line,
                                     re.M)
                temp = re.search(
                    r'\[image].*?(<ruby><rb>).*?\[/image]', line, re.M)
                while temp is not None:
                    line = line[:temp.start(1)] + line[temp.end(1):]
                    temp = re.search(
                        r'\[image].*?(<ruby><rb>).*?\[/image]', line, re.M)
                test = test + line
            txt = test
        if RBforDOCX:
            txt = re.sub(
                r'(<[^<>]*?="[^<>]*?)<ruby><rb>(.*?)</rb><rp>\(</rp><rt>.*?</rt><rp>\)</rp></ruby>([^<>]*?"[^<>]*?/>)',
                r'\1\2\3', txt, flags=re.M)

        return txt

    def segmentation(self, text, maxSet=50000):  # 分割成转换字符数上限
        segmented = []
        index = 0
        while index + maxSet <= len(text):
            indexIn = index + maxSet
            # text[index:index + maxSet:-1]:
            for i in range(index + maxSet, index, -1):
                if text[i] == "\n" or text[i - 6:i] == '</w:r>':
                    segmented.append(text[index:indexIn])
                    index = indexIn
                    break
                indexIn += -1
        segmented.append(text[index:])

        return segmented

    def segmentationLine(self, text: str):  # 分割str为行
        line = []
        index = 0
        indexIn = 0
        while index < len(text):
            if text[index] == "\n" or text[index - 6:index] == '</w:r>':
                line.append(text[indexIn:index])
                # line.append('\n')
                indexIn = index
            index += 1
        line.append(text[indexIn:])
        return line
