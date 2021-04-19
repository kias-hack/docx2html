from zipfile import ZipFile
import sys
import os
from lxml import etree
from os import path

docx = open("document.xml", "r", encoding="utf-8").read(-1)

root = etree.parse("document.xml").getroot()

nsmap = root.nsmap

body = root.find("w:body", root.nsmap)
#  [{'xml':'http://www.w3.org/XML/1998/namespace'}]

# [print(e.attrib.get("{http://www.w3.org/XML/1998/namespace}space")) for e in body.findall("w:p/w:r/w:t", root.nsmap)]

# exit()

class DocxFactory:
    @staticmethod
    def build(root, nsmap : list):
        if None is root:
            return StubTag(root, nsmap)

        prefix = "{" + nsmap['w'] + "}"
        
        if prefix + "p" == root.tag:
            return PTag(root, nsmap)
        if prefix + "r" == root.tag:
            return RTag(root, nsmap)
        if prefix + "t" == root.tag:
            return TTag(root, nsmap)
        if prefix + "drawing" == root.tag:
            return DrawingTag(root, nsmap)
        if prefix + "pPr" == root.tag:
            return Properties(root, nsmap)
        if prefix + "rPr" == root.tag:
            return Properties(root, nsmap)
        if prefix + "br" == root.tag:
            return BrTag(root, nsmap)

        return StubTag(root, nsmap)
        

    @staticmethod
    def createAll(root, nsmap = []):
        for tag in root:
            yield DocxFactory.build(tag, nsmap)

class BrTag:
    def __init__(self, root, nsmap = []):
        pass

    def html(self, styles_settings = None):
        return "<br/>"

class StubTag:
    def __init__(self, root, nsmap = []):
        pass
    def html(self, styles_settings = None):
        return ""

class DrawingTag:
    def __init__(self, root, nsmap = []):
        pass

    def html(self, styles_settings = None):
        return "<img />"

class TTag:
    def __init__(self, root, nsmap):
        self.__tag = root
        self.__text = "" if None is root.text else root.text

    def html(self, styles_settings = None):
        font_size = styles_settings.getFontSize()
        color = styles_settings.getColor()
        
        style = ";".join([
            "" if None is font_size else "font-size: {}px".format(str(font_size)),
            "" if None is color else "color: {}".format(str(color))
             ])

        res = '<span style="{}"'.format(style) + '>{}</span>'.format(self.__text)

        if styles_settings.isUnderline():
            res = "<u>{}</u>".format(res)
        if styles_settings.isItalic():
            res = "<em>{}</em>".format(res)
        if styles_settings.isBold():
            res = "<strong>{}</strong>".format(res)

        return res

class Properties:
    def __init__(self, root, nsmap=[]):
        self.__isBold = False
        self.__isItalic = False
        self.__isUnderline = False
        self.__fontSize = None
        self.__header = None
        self.__color = None

        if root is None:
            return

        if None is not root.find("w:color", nsmap):
            self.__color = root.find("w:color", nsmap).attrib.get("{" + nsmap["w"] + "}val")

        if None is not root.find("w:pStyle", nsmap):
            header = root.find("w:pStyle", nsmap).attrib.get("{" + nsmap["w"] + "}val")

            if "a3" == header:
                self.__header = "h1"

        if None is not root.find("w:b", nsmap):
            self.__isBold = True

        if None is not root.find("w:sz", nsmap):
            self.__fontSize = root.find("w:sz", nsmap).attrib.get("{" + nsmap["w"] + "}val")

        if None is not root.find("w:i", nsmap):
            self.__isItalic = True

        if None is not root.find("w:u", nsmap):
            self.__isUnderline = True
    
    def isBold(self):
        return self.__isBold

    def isItalic(self):
        return self.__isItalic

    def isUnderline(self):
        return self.__isUnderline

    def getHeader(self):
        return self.__header

    def getFontSize(self):
        return self.__fontSize
    
    def getColor(self):
        return self.__color

    def html(self, styles_settings = None):
        return ""

class RTag:
    def __init__(self, root, nsmap=[]):
        self.__rPr = Properties(root.find("w:rPr", nsmap), nsmap)

        self.__content = DocxFactory.createAll(root, nsmap)

    def html(self, styles_settings = None):
        return "".join([str(i.html(self.__rPr)) for i in self.__content])
        
class PTag:
    def __init__(self, root, nsmap = []):
        self.__pPr = Properties(root.find("w:pPr", nsmap), nsmap)

        self.__r = [DocxFactory.build(tag, nsmap) for tag in root.findall("w:r", nsmap)]

    def html(self, styles_settings = None):
        inner = "".join([str(i.html(self.__pPr)) for i in self.__r])

        if self.__pPr.getHeader():
            h = self.__pPr.getHeader()
            return "<" + h + ">{}</".format(inner) + h + ">"
        return "<p>{}</p>".format(inner)


for i in DocxFactory.createAll(body, nsmap):
    print(i.html())