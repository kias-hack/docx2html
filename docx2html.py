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

"""
TODO сделать отработку пробелов, нету пробелов в одном из примере. В xml не знаю где они деваются возможно какое то свойство (но везде стоит xml:space="preserve", стоит подумать)
"""

"""
TODO возможно вынести построение тегов в статические методы, чтобы можно было делать конвертацию в две стороны
"""

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

    def html(self):
        return self.__text

class Properties:
    def __init__(self, root, nsmap=[]):
        self.__nsmap = nsmap
        self.__styles = {}
        self.__header = None

        if root is None:
            return

        if None is not self.getValue(root, "color"):
            self.__styles["color"] = self.getValue(root, "color")

        if None is not self.getValue(root, "pStyle"):
            header = self.getValue(root, "pStyle")

            if "a3" == header:
                self.__header = "h1"

        if None is not root.find("w:b", nsmap):
            self.__styles["font-weight"] = "bold"

        if None is not self.getValue(root, "sz"):
            self.__styles["font-size"] = self.getValue(root, "sz")

        if None is not root.find("w:i", nsmap):
            self.__styles["font-style"] = "italic"

        if None is not root.find("w:u", nsmap):
            self.__styles["text-decoration"] = "underline"

        if None is not self.getValue(root, "shd", "fill"):
            self.__styles["background-color"] = "#{}".format(self.getValue(root, "shd", "fill"))
    
    def getValue(self, root, tag, attr = "val"):
        tag = root.find("w:{}".format(tag), self.__nsmap)

        if None is tag:
            return None

        attrname = "{" + (self.__nsmap["w"]) + "}" + "{}".format(attr)
        return tag.attrib.get(attrname)

    def getHeader(self):
        return self.__header

    def html(self):
        style_html = ";".join(["{}:{}".format(key, self.__styles[key]) for key in self.__styles.keys()])

        if len(style_html) == 0:
            return ""

        return ' style="{}"'.format(style_html)

class RTag:
    def __init__(self, root, nsmap=[]):
        self.__rPr = Properties(root.find("w:rPr", nsmap), nsmap)

        self.__content = DocxFactory.createAll(root, nsmap)

    def html(self):
        inner = "".join([str(i.html()) for i in self.__content])

        return "<span{}>{}</span>".format(self.__rPr.html(), inner)
        
class PTag:
    def __init__(self, root, nsmap = []):
        self.__pPr = Properties(root.find("w:pPr", nsmap), nsmap)

        self.__r = [DocxFactory.build(tag, nsmap) for tag in root.findall("w:r", nsmap)]

    def html(self):
        inner = "".join([str(i.html()) for i in self.__r])

        if self.__pPr.getHeader():
            h = self.__pPr.getHeader()
            return "<" + h + "{}>{}</".format(self.__pPr.html(), inner) + h + ">"
        
        return "<p{}>{}</p>".format(self.__pPr.html(), inner)


for i in DocxFactory.createAll(body, nsmap):
    print(i.html())
