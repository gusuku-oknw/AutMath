import os
import wx
import random
import datetime
import math
from docx import Document
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Inches, RGBColor
from docx.enum.style import WD_STYLE_TYPE
from docx.api import Document
from docx.opc.oxml import qn

from docx2pdf import convert

Name_Text = '＃__    名前_____________'
MathL = ''
MathL_answer = ''


class Configuration:
    def __init__(self):
        percent = 0

    def probability(self, percent):
        coin = random.random()
        # %で確率調整
        if coin * 100 < percent:
            return True
        else:
            return False

    def now_time(self):
        dt_now = datetime.datetime.now()
        dt_name = str(dt_now.strftime('%Y.%m%d-%M%S'))
        return dt_name

    @staticmethod
    def mypath(file):
        cwd = os.path.abspath(file)
        # cwd += '/'
        print(cwd)
        return cwd


class ThreeButtonEvent(wx.PyCommandEvent):
    def __init__(self, evtType, id):
        wx.PyCommandEvent.__init__(self, evtType, id)
        self.index = None

    def set_selected_index(self, index):
        self.index = index

    def get_selected_index(self):
        return self.index


myEVT_THREE_BUTTON = wx.NewEventType()
EVT_THREE_BUTTON = wx.PyEventBinder(myEVT_THREE_BUTTON, 1)


class ThreeButtonPanel(wx.Panel):
    def __init__(self, parent, id=-1):
        wx.Panel.__init__(self, parent, id)
        # Create widgets.
        self.button1 = wx.ToggleButton(self, label='生成')
        self.button2 = wx.ToggleButton(self, label='2')
        self.button3 = wx.ToggleButton(self, label='出力')
        self.button1.index = 0
        self.button2.index = 1
        self.button3.index = 2
        # Set event handlers.
        self.Bind(wx.EVT_TOGGLEBUTTON, self.on_toggle_button)
        # Set sizer.
        sizer = wx.BoxSizer(wx.HORIZONTAL)
        sizer.Add(self.button1)
        sizer.Add(self.button2)
        sizer.Add(self.button3)
        self.SetSizer(sizer)

    def on_toggle_button(self, evt):
        button = evt.GetEventObject()
        index = button.index
        if button.GetValue():
            if button != self.button1:
                self.button1.SetValue(False)
            if button != self.button2:
                self.button2.SetValue(False)
            if button != self.button3:
                self.button3.SetValue(False)
        else:
            index = -1
        # Raise event.
        evt = ThreeButtonEvent(myEVT_THREE_BUTTON, self.GetId())
        evt.set_selected_index(index)
        self.GetEventHandler().ProcessEvent(evt)


class AutMath:
    # 問題docx生成のメソッド
    @staticmethod
    def add_docx(name):
        global MathL

        print(MathL)
        doc = Document()

        styles = doc.styles
        style = styles.add_style('Original-Style', WD_STYLE_TYPE.PARAGRAPH)
        font = style.font
        font.size = Pt(15)
        font.name = 'BIZ UDゴシック'
        paragraph_format = style.paragraph_format

        p = doc.add_paragraph(MathL, style=style)

        p.alignment = WD_ALIGN_PARAGRAPH.LEFT

        doc.save(name + ".docx")

        convert(name + ".docx", name + ".pdf")

    # 答えdocx生成のメソッド
    @staticmethod
    def add_docx_answer(name):
        global MathL_answer

        doc_answer = Document()

        styles = doc_answer.styles
        style = styles.add_style('Original-Style', WD_STYLE_TYPE.PARAGRAPH)
        font = style.font
        font.size = Pt(15)
        font.name = 'BIZ UDゴシック'
        paragraph_format = style.paragraph_format

        p_answer = doc_answer.add_paragraph(MathL_answer, style=style)

        p_answer.alignment = WD_ALIGN_PARAGRAPH.LEFT

        doc_answer.save(name + ".docx")

        convert(name + ".docx", name + ".pdf")

        # convert()

    # 足し算のメソッド
    def addition(self):
        ad1 = random.randint(1, 99)
        ad2 = random.randint(1, 99)
        coin = Configuration()
        if coin.probability(30):
            ad1 *= 10

        if coin.probability(30):
            ad1 *= -1

        if coin.probability(30):
            ad2 *= 10

        sadd = (str(ad1) + '+' + str(ad2) + '=  \t')
        ad3 = ad1 + ad2
        sadd_answer = ('=' + str(ad3) + '\t' + '\t')
        return sadd, sadd_answer

    # 引き算のメソッド
    @staticmethod
    def subtraction():
        su1 = random.randint(1, 99)
        su2 = random.randint(1, 99)
        coin = Configuration()
        if coin.probability(30):
            su1 *= 10

        if coin.probability(30):
            su1 *= -1

        if coin.probability(30):
            su2 *= 10

        ssub = (str(su1) + '-' + str(su2) + '=  \t')
        su3 = su1 - su2
        ssub_answer = ('=' + str(su3) + '\t' + '\t')
        return ssub, ssub_answer

    # 掛け算のメソッド
    @staticmethod
    def multiplication():
        mu1 = random.randint(1, 9)
        mu2 = random.randint(1, 9)
        coin = Configuration()
        if coin.probability(60):
            mu1 *= random.randint(1, 10)

        if coin.probability(30):
            mu1 *= -1

        if coin.probability(50):
            mu2 *= random.randint(1, 10)

        if coin.probability(30):
            mu2 *= -1

        smul = (str(mu1) + '×' + str(mu2) + '=  \t')
        smul_answer = ('=' + str(mu1 * mu2) + '\t' + '\t')
        return smul, smul_answer

    # 割り算のメソッド
    @staticmethod
    def division():
        di2 = random.randint(2, 9)
        di1 = di2 * random.randint(2, 10)
        coin = Configuration()
        if coin.probability(30):
            di1 *= random.randint(1, 10)
            if coin.probability(30):
                di2 *= random.randint(1, 10)
        if coin.probability(30):
            di1 *= -1
        if coin.probability(30):
            di2 *= -1

        sdiv = (str(di1) + '÷' + str(di2) + '=')
        sdiv_answer = ('=' + str(math.floor(di1 / di2)))
        return sdiv, sdiv_answer


class MyFrame(wx.Frame):
    def __init__(self):
        wx.Frame.__init__(self, None, -1, "Title", size=(300, 300))
        panel = ThreeButtonPanel(self)
        panel.Bind(EVT_THREE_BUTTON, self.OnThreeButton)

    @staticmethod
    def OnThreeButton(evt):
        global Name_Text
        global MathL
        global MathL_answer

        if evt.get_selected_index() == 0:
            max = 25
            MathL += Name_Text + '\n'
            MathL_answer += Name_Text + '\n'
            for i in range(max):
                math = AutMath()
                add_math = math.addition()

                MathL += '(' + str(1 + i).zfill(2) + ')' + add_math[0]
                MathL_answer += '(' + str(1 + i).zfill(2) + ')' + add_math[1]

                math = AutMath()
                sub_math = math.subtraction()
                MathL += '(' + str(26 + i) + ')' + sub_math[0]
                MathL_answer += '(' + str(26 + i) + ')' + sub_math[1]

                math = AutMath()
                mul_math = math.multiplication()
                MathL += '(' + str(51 + i) + ')' + mul_math[0]
                MathL_answer += '(' + str(51 + i) + ')' + mul_math[1]

                math = AutMath()
                div_math = math.division()
                MathL += '(' + str(76 + i) + ')' + div_math[0]
                MathL_answer += '(' + str(76 + i) + ')' + div_math[1]

                MathL += '\n'
                MathL_answer += '\n'

        if evt.get_selected_index() == 2:
            file_data = Configuration()
            math = AutMath()

            # 問題のファイル名設定
            file_name = file_data.now_time()
            file_name += "_Q"
            file_name = file_data.mypath(file_name)

            # 答えのファイル名設定
            file_name_answer = file_data.now_time()
            file_name_answer += "_A"
            file_name_answer = file_data.mypath(file_name_answer)

            # 問題のdocxの生成
            math.add_docx(file_name)
            # 答えのdocxの生成
            math.add_docx_answer(file_name_answer)
            # 内容のリセット
            MathL = None
            MathL_answer = None

        print('Selected index =', evt.get_selected_index())


if __name__ == '__main__':
    app = wx.PySimpleApp()
    MyFrame().Show()
    app.MainLoop()
