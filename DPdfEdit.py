from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.lib.colors import black, red
import datetime
#import addpage
#from pdfformfiller import PdfFormFiller
#import DPdfEdit

class Pdf:
    PAGE_SIZE = {'x': 595, 'y': 842}       # A4サイズ（x,y） x,yは、ポイント  辞書型です
    PAGE_MARGIN = {'top': 72, 'bottom': 72, 'left': 53, 'right': 53}
    FONT = {'name': 'HeiseiMin-W3', 'size': 10}
    ROW_HEIGHT = 18
    LINE_WIDTH = 0.5
    TEXT_COL_MARGIN = 3
    TEXT_ROW_MARGIN = 5
    HEADER_Y = 40
    FOOTER_Y = 820

    LinePos = [0, 0]
    VERTICAL_LINE_1 = [140]             # 印刷範囲内での左からの位置
    VERTICAL_LINE_2 = [70, 140, 210, 270, 390]
    VERTICAL_LINE_3 = [140, -210, -420]       # 負の値は線を引かない

    textPos1 = []
    textPos2 = []
    textPos3 = []

    pdfFile = 0     # 実際にはクラスである
    fileName = ''
    pageNo = 0
    textColor = 0

    def __init__(self, fileName):
        self.pageWidth = self.PAGE_SIZE['x'] - self.PAGE_MARGIN['left'] - self.PAGE_MARGIN['right']
        self.pageHeight = self.PAGE_SIZE['y'] - self.PAGE_MARGIN['top'] - self.PAGE_MARGIN['bottom']
        self.pageCenterX = self.pageWidth // 2 + self.PAGE_MARGIN['left']
        self.maxRowsCount = self.pageHeight // self.ROW_HEIGHT

        self.LinePos[0] = self.PAGE_MARGIN['left']
        self.LinePos[1] = self.LinePos[0] + self.pageWidth

        self.textPos1.clear()
        self.textPos1.append(self.LinePos[0] + self.TEXT_COL_MARGIN)
        for i in self.VERTICAL_LINE_1:
            self.textPos1.append(i + self.PAGE_MARGIN['left'] + self.TEXT_COL_MARGIN)

        self.textPos2.clear()
        self.textPos2.append(self.LinePos[0] + self.TEXT_COL_MARGIN)
        for i in self.VERTICAL_LINE_2:
            self.textPos2.append(i + self.PAGE_MARGIN['left'] + self.TEXT_COL_MARGIN)

        self.textPos3.clear()
        self.textPos3.append(self.LinePos[0] + self.TEXT_COL_MARGIN)
        for i in self.VERTICAL_LINE_3:
            if i >= 0:
                self.textPos3.append(i + self.PAGE_MARGIN['left'] + self.TEXT_COL_MARGIN)
            else:
                self.textPos3.append(-i + self.PAGE_MARGIN['left'] + self.TEXT_COL_MARGIN)

        self.pdfFile = canvas.Canvas(fileName, bottomup=False)     # 左上を原点にする
        self.pdfFile.saveState()

        self.pdfFile.setAuthor('DLitePrinter')
        self.pdfFile.setTitle('アマチュア無線局免許申請書')
        self.pdfFile.setSubject('サンプル')

#        self.pdfFile.setP.setPageSize(self.PAGE_SIZE['x'], self.PAGE_SIZE['y'])     # A4サイズ
#        self.pdfFile.setPageSize(self.PAGE_SIZE['x'], self.PAGE_SIZE['y'])     # A4サイズ
        pdfmetrics.registerFont(UnicodeCIDFont(self.FONT['name']))
        self.pdfFile.setFont(self.FONT['name'], self.FONT['size'])
        self.pdfFile.setLineWidth(self.LINE_WIDTH)

        self.textColor = black


    def NewPage(self):
        self.PrintFooter()
        #        self.pdfFile.restoreState()
        self.pdfFile.showPage()
        self.pdfFile.setFont(self.FONT['name'], self.FONT['size'])
        self.pdfFile.setLineWidth(self.LINE_WIDTH)
        self.PrintHeader()

    def Finalize(self):
        self.PrintFooter()
        # self.pdfFile.restoreState()
        self.pdfFile.saveState()
        self.pdfFile.save()

    def GetTextWidth(self, text):
        return self.pdfFile.stringWidth(text, self.FONT['name'], self.FONT['size'])

    def PosY(self, rowNo=0):
        return rowNo * self.ROW_HEIGHT + self.PAGE_MARGIN['top'] - self.TEXT_ROW_MARGIN

    def ChangeColor(self):
        self.pdfFile.setFillColor(self.textColor)
        self.pdfFile.setStrokeColor(self.textColor)
        self.textColor = black

    def PrintText0(self, rowNo=0, left=0, text=''):
        x = self.PAGE_MARGIN['left'] + left
        y = self.PosY(rowNo)
        self.pdfFile.drawCentredString(self.pageCenterX, y, text)

    def PrintText1(self, rowNo=0, text1='', text2='', upper=False, lower=False):
        y = self.PosY(rowNo)
        if text1 != '':
            self.ChangeColor()  # 指定した色にする
        self.pdfFile.drawString(self.textPos1[0], y, text1)
        self.pdfFile.drawString(self.textPos1[1], y, text2)
        self.ChangeColor()  # 黒に戻す
        self.DrowHorizontalLine(rowNo, upper, lower)

    def PrintText2(self, rowNo=0, text1='', text2='', text3='', upper=False, lower=False, printing=True):
        if printing:
            y = self.PosY(rowNo)
            self.pdfFile.drawString(self.textPos2[0], y, '装置の区別')
            self.pdfFile.drawString(self.textPos2[1], y, text1)
            self.pdfFile.drawString(self.textPos2[2], y, '変更の種別')
            if text2 != '':
                self.textColor = red
                self.ChangeColor()      # 指定した色にする
            self.pdfFile.drawString(self.textPos2[3], y, text2)
            self.ChangeColor()      # 黒に戻す
            self.pdfFile.setFillColor(black)
            self.pdfFile.setStrokeColor(black)
            self.pdfFile.drawString(self.textPos2[4], y, '技術基準適合証明書番号')
            self.pdfFile.drawString(self.textPos2[5], y, text3)
            self.DrowHorizontalLine(rowNo, upper, lower)

    def PrintText3(self, rowNo=0, text1='', text2='', text3='',  text4='', upper=False, lower=False, printing=True):
        if printing:
            y = self.PosY(rowNo)
            self.pdfFile.drawString(self.textPos3[0], y, text1)
            self.pdfFile.drawString(self.textPos3[1], y, text2)
            self.pdfFile.drawString(self.textPos3[2], y, text3)
            self.pdfFile.drawString(self.textPos3[3], y, text4)
            self.DrowHorizontalLine(rowNo, upper, lower)

    def PrintText4(self, rowNo=0, text1='', text2='', printing=True):
        if printing:
            x = self.LinePos[1]
            y = self.PosY(rowNo)
            s = text1
            if text2 != '':
                s += '(' + text2 + ')'
            self.pdfFile.drawRightString(x, y, s)

    def PrintHeader(self):
        y = self.HEADER_Y
        x = self.textPos1[0]
        if len(self.fileName) > 70:
            s = self.fileName[0:13] + '....' + self.fileName[-50:]
        else:
            s = self.fileName
        self.pdfFile.drawString(x, y, s)
        x = self.LinePos[1]
        now = datetime.datetime.now()
        s = now.strftime("%Y-%m-%d %H:%M:%S")
        self.pdfFile.drawRightString(x, y, s)

    def PrintFooter(self):
        y = self.FOOTER_Y
        x = self.PAGE_SIZE['x'] // 2
        self.pageNo += 1
        s = "-  Page {}  -".format(str(self.pageNo))
        self.pdfFile.drawCentredString(x, y, s)

    def DrowHorizontalLine(self, rowNo=0, upper=False, lower=False):
        x1 = self.LinePos[0]
        x2 = self.LinePos[1]
        y1 = (rowNo - 1) * self.ROW_HEIGHT + self.PAGE_MARGIN['top']
        y2 = rowNo * self.ROW_HEIGHT + self.PAGE_MARGIN['top']
        if upper:
            self.pdfFile.line(x1, y1, x2, y1)
        if lower:
            self.pdfFile.line(x1, y2, x2, y2)

    def DrowRectangle(self, y1, y2):
        x = self.LinePos[0]
        y = y1

        width = self.LinePos[1] - x
        height = y2 - y1
        self.pdfFile.setLineWidth(self.LINE_WIDTH)
        self.pdfFile.rect(x, y, width, height)

    def DrowVerticalLine(self, linePos, y1, y2):
        self.pdfFile.setLineWidth(self.LINE_WIDTH)
        for li in linePos:
            x = li
            if x >= 0:
                x = x + self.PAGE_MARGIN['left']
                self.pdfFile.line(x, y1, x, y2)
        self.DrowRectangle(y1, y2)

    def DrowVerticalLine1(self, rowNo=0, rowLines=1):
        y1 = (rowNo - 1) * self.ROW_HEIGHT + self.PAGE_MARGIN['top']
        y2 = y1 + rowLines * self.ROW_HEIGHT
        self.DrowVerticalLine(self.VERTICAL_LINE_1, y1, y2)

    def DrowVerticalLine2(self, rowNo=0, rowLines=1):
        y1 = (rowNo - 1)* self.ROW_HEIGHT + self.PAGE_MARGIN['top']
        y2 = y1 + self.ROW_HEIGHT
        self.DrowVerticalLine(self.VERTICAL_LINE_2, y1, y2)

    def DrowVerticalLine3(self, rowNo=0, rowLines=1):
        y1 = (rowNo - 1) * self.ROW_HEIGHT + self.PAGE_MARGIN['top']
        y2 = y1 + self.ROW_HEIGHT
        self.DrowVerticalLine(self.VERTICAL_LINE_3, y1, y2)


# なぜか「EOF marker not found」のエラーになる。
# ファイル内には「%%EOF」が存在する
    def AddPage(self, infile, outfile):
        addpage.addPage(infile, outFile=outfile, fontName='HeiseiMin-W3', start=1, skip=0,  marginX=10, marginY=50, pageFormat='-  Page %d / %d  -', alignment='CENTER', fontSize=10)



if __name__ == "__main__":
    # pFile = Pdf('./python.pdf')
    # pFile.DrowHorizontalLine(5, True, True)
    # pFile.DrowHorizontalLine(15, True, False)
    # pFile.DrowHorizontalLine(25, False, True)
    # pFile.DrowVerticalLine(5, 1, True, True)
    # pFile.DrowVerticalLine(15, 2, True, False)
    # pFile.DrowVerticalLine(25, 10, False, True)
    # pFile.PrintText(5, 0, 'string1 あいうえお 無線局申請書')
    # pFile.PrintText(15, 10, 'string2 かきくけこ')
    # pFile.PrintText(25, 20, 'string3 さしすせそ')

    # pFile.Finalize()
    #
    # print(pFile.maxRows)
    pass


