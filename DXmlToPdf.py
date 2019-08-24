import os
import argparse
import zipfile

import xml.etree.ElementTree as ET
import configparser                 # 設定ファイルのパーサー
import codecs
import DPdfEdit
from reportlab.lib.colors import black, red
import webbrowser

errMessage = {1: "Input file name isn't designated.",
              2: "Input file extension isn't 'zip' and 'xml'.",
              3: "Input file doesn't exist.",
              4: "Other application use this output file.",
              5: "The form of input file is different.",
              6: "Input file doesn't exist in zip file.",
              7: "Cannot Open DXmlToPdf.ini file."}

def OutErrMsg(ErrNo, Option=''):
    ConfigSetInt('ERROR', 'Error', '{0:04}'.format(ErrNo))
    s = errMessage[ErrNo]
    if Option != '':
        s += (' "' + Option + '"')
    print('Error={0:04} ---{1}'.format(ErrNo, s))
    ConfigSet('ERROR', 'Message', s)

iniFileName = './DXmlToPdf.ini'


isBrowse = True
config = configparser.ConfigParser()
jichitai = dict()
#dict = {}

ns = {'m': 'http://www.denpa.soumu.go.jp/sinsei/kousei',
      'ds': 'http://www.w3.org/2000/09/xmldsig#',
      'p': 'http://www.denpa.soumu.go.jp/sinsei'}

class Count:
    c = 0
    def __init__(self, intVal=0):
        if type(intVal) is int:
            self.c = intVal
    def Inc(self):
        self.c += 1
        return self.c
    def Get(self):
        return self.c
    def Set(self, intVal):
        self.__init__(intVal)

def SplitText(pFile, text, length):
    lst = text.split(',')
    r = ''
    t = []
    for s in lst:
        if (pFile.GetTextWidth(r) + pFile.GetTextWidth(s)) <= length:
            r += (',' + s)
        else:
            t.append(r[1:])
            r = (',' + s)
    if r != '':
        t.append(r[1:])
    return t


def ConfigGet(section, option, default=''):
    r = default
    if section in config:
        if config.has_option(section, option):
            r = config.get(section, option)
    return r

def ConfigGetInt(section, option, default=0):
    r = default
    s = ConfigGet(section, option)
    if s != '':
        r = int(s)
    return r

def ConfigSet(section, option, value):
    if not (section in config):
        config.add_section(section)
    config.set(section, option, value)

def ConfigSetInt(section, option, value):
    ConfigSet(section, option, str(value))

def ConfigSetBool(section, option, value):
    ConfigSet(section, option, str(value))

def LoadJichitaiCode():
    with open('JichitaiCode.txt', encoding='utf-8') as f:
        for line in f:
            s = line.split()
            s1 = s[0]
            s2 = s[1]
            jichitai[s1] = s2

def GetJichitai(code):
    if code == '':
        return ''
    if len(jichitai) == 0:      # 自治体コードがロードされてないなら
        LoadJichitaiCode()
    s = jichitai[code]
    return s

def XmlTetsuzuki(node):
    n = node.find('m:構成管理情報/m:管理情報/m:手続情報/m:手続ID', ns)
    s1 = n.text or ''
    s2 = ConfigGet('TETUZUKI', s1)
    return [s1, s2]

def XmlVersion(node):
    n = node.find('m:構成管理情報/m:申請書属性情報/m:申請書様式バージョン', ns)
    return n.text or ''

def XmlDate(node):
    s1 = ConfigGet('GENGO', node.find('p:年号', ns).text or '')
    s2 = node.find('p:年', ns).text or ''
    s3 = node.find('p:月', ns).text or ''
    s4 = node.find('p:日', ns).text or ''
    if s1 == '':
        return ''
    else:
        return s1 + s2 + '年' + s3 + '月' + s4 + '日'

def XmlMenkyo(node):
    s1 = node.find('p:免許の番号地方部', ns).text or ''
    s2 = node.find('p:免許の番号区分', ns).text or ''
    s3 = node.find('p:免許の番号_連番部', ns).text or ''
    if s1 == '':
        return ''
    else:
        return s1 + s2 + '第' + s3 + '号'

def XmlMode(node):
    s = ''
    for n in node.findall('p:電波の型式等情報', ns):
        s1 = n.find('p:電波の型式', ns).text or ''
        if s1 == '':
            break
        s2 = n.find('p:占有周波数帯幅', ns).text or ''
        s = s + ',' + s1
        if s2 != '':
            s = s + '(' + s2 + ')'
    if s != '':
        s = s[1:]
    return s

def XmlPower(power):
    f = float('0' + power)
    s = ''
    un = ['kW', 'W', 'mW', 'nW', 'uW']
    for x, u in enumerate([1000, 1, 0.001, 0.000001, 0.000000001]):
        f1 = f // u
        if f1 != 0:
            s = str(int(f / u)) + un[x]
            break
    return s

def XmlGiteki(node, version):
    if version <= '0008':
        s = node.find('p:届出番号', ns) or ''
        if s == '':
            s = node.find('p:技適_番号/p:技適_番号_記号部', ns).text or ''
            s2 = node.find('p:技適_番号/p:技適_番号_番号部', ns).text or ''
            if s2 != '':
                s = s + '-' + s2
    else:
        s = node.text or ''
    return s

def XmlUmu(node):
    return "有" if (node.text or '') == '1' else "無"


def XmlCommon(root):
    dict = {}
    dict.clear()
    dict['ERROR'] = False
    d = XmlTetsuzuki(root)
    dict['手続き'] = d[0]
    dict['手続き名'] = d[1]
    if dict['手続き名'] == '':
        dict['ERROR'] = True
        return
    dict['申請書バージョン'] = XmlVersion(root)

# 0.申請者先に関する項目
    dict['宛先'] = ConfigGet('ATESAKI', root.find('p:申請書/p:申請事項/p:宛先', ns).text or '')
    dict['申請区分'] = ConfigGet('SHINSEI', root.find('p:申請書/p:申請事項/p:申請区分', ns).text or '')

# 1.申請者に関する項目
    n = root.find('p:申請書/p:申請者等情報', ns)
    dict['個人_社団の別'] = ConfigGet('RADIO', n.find('p:団体_個人の別', ns).text or '')
    if dict['個人_社団の別'] == '個人':    # 個人
        dict['社団名'] = ''
        dict['社団名フリガナ'] = ''
        dict['氏名'] = n.find('p:社団名_クラブ名_又は氏名', ns).text or ''
        dict['氏名フリガナ'] = n.find('p:社団名_クラブ名_又は氏名フリガナ', ns).text or ''
    else:
        dict['社団名'] = n.find('p:社団名_クラブ名_又は氏名', ns).text or ''
        dict['社団名フリガナ'] = n.find('p:社団名_クラブ名_又は氏名フリガナ', ns).text or ''
        dict['氏名'] = n.find('p:代表者名', ns).text or ''
        dict['氏名フリガナ'] = n.find('p:代表者名フリガナ', ns).text or ''
    s = n.find('p:郵便番号', ns).text or ''
    dict['郵便番号'] = s[0:3] + '-' + s[3:7]
    dict['都道府県'] = GetJichitai(n.find('p:住所/p:都道府県_市区町村/p:都道府県', ns).text or '')
    dict['市区町村'] = GetJichitai(n.find('p:住所/p:都道府県_市区町村/p:市区町村', ns).text or '')
    dict['町名'] = n.find('p:住所/p:町_丁目', ns).text or ''
    dict['住所'] = dict['都道府県'] + dict['市区町村'] + dict['町名']
    dict['電話番号'] = n.find('p:電話番号', ns).text or ''
    dict['国籍'] = GetJichitai(n.find('p:国籍', ns).text or '')

# 2.欠格事由に関する項目
    if dict['申請書バージョン'] <= '0008':
        n = root.find('p:申請書/p:再免許情報/p:無線局事項書及び工事設計書の内容/p:欠格事由の有無', ns)
    else:
        n = root.find('p:申請書/p:欠格事由', ns)
    dict['欠格事由の有無'] = XmlUmu(n)

# 3.免許に関する項目
    if dict['手続き'] == 'D055':       # 廃局の時
        n = root.find('p:申請書/p:廃止情報', ns)
        dict['免許の番号'] = XmlMenkyo(n.find('p:免許の番号', ns))
        dict['呼出符号'] = n.find('p:呼出符号', ns).text or ''
        dict['廃止年月日'] = XmlDate(n.find('p:廃止する年月日', ns))
        dict['備考'] = n.find('p:廃止_備考', ns) or ''
    elif dict['手続き'] == 'D053':  # 変更の時
        n = root.find('p:申請書/p:事項書_工事設計書情報/p:無線局の種別等', ns)
        dict['免許の番号'] = XmlMenkyo(n.find('p:免許の番号', ns))
        dict['呼出符号'] = n.find('p:呼出符号', ns).text or ''
        dict['備考'] = ''
    else:
        if dict['申請書バージョン'] <= '0008':
            n = root.find('p:申請書/p:再免許情報', ns)
            dict['免許の番号'] = XmlMenkyo(n.find('p:免許の番号', ns))
            dict['呼出符号'] = n.find('p:呼出符号', ns).text or ''
            dict['免許の年月日'] = XmlDate(n.find('p:免許の年月日', ns))
            dict['備考'] = n.find('p:備考/備考_入力欄', ns) or ''
        else:
            n = root.find('p:申請書', ns)
            dict['免許の番号'] = XmlMenkyo(n.find('p:免許の番号', ns))
            dict['呼出符号'] = n.find('p:呼出符号', ns).text or ''
            dict['免許の年月日'] = XmlDate(n.find('p:免許の年月日', ns))
            dict['備考'] = n.find('p:備考/備考_入力欄', ns) or ''

# 4.電波利用料に関する項目
    n = root.find('p:申請書/p:電波利用料の前納の申出', ns)
    dict['前納_有無'] = XmlUmu(n.find('p:前納_有無', ns))
    if dict['前納_有無'] == '有':
        s = n.find('p:前納に係る期間/p:前納に係る期間区分', ns).text or ''
        if s  == '1':
            dict['前納期間'] = '無線局の免許の有効期間まで前納します。'
        else:
            dict['前納期間'] = 'その他（' + (n.find('p:前納に係る期間/p:年', ns).text or '') + '年）'
    else:
        dict['前納期間'] = ''

# 5.連絡先に関する項目
    if dict['申請書バージョン'] <= '0008':
        dict['連絡先氏名'] = ''
        dict['連絡先氏名フリガナ'] = ''
        dict['連絡先電話番号'] = ''
        dict['電子メールアドレス'] = ''
    else:
        n = root.find('p:申請書/p:申請者等情報/p:申請に関する連絡責任者', ns)
        dict['連絡先氏名'] = n.find('p:氏名', ns).text or ''
        dict['連絡先氏名フリガナ'] = n.find('p:氏名フリガナ', ns).text or ''
        dict['連絡先電話番号'] = n.find('p:電話番号', ns).text or ''
        dict['電子メールアドレス'] = n.find('p:電子メールアドレス', ns).text or ''

# 6.申請手数料等に関する項目
    if dict['申請書バージョン'] >= '0009':
        n = root.find('p:申請書/p:申請手数料', ns)
        dict['手数料額'] =  int('0' + (n.find('p:手数料額', ns).text or ''))
        n = root.find('p:申請書/p:免許状受取方法', ns)
        d = ['', '返信用封筒別送', '窓口受領', '送料受取人払いによる受領(料金:500円)']
        s = '0' + (n.find('p:免許状受取区分', ns).text or '')
        dict['免許状受取方法'] = d[int(s)]

    return dict



def Saimen(root, pFile, dict):
    if dict['ERROR']:
        OutErrMsg(5)      # ファイル形式不正
        return
    pFile.PrintHeader()
    i = Count(0)
    pFile.PrintText0(i.Get(), 200, dict['手続き名'])

    i.Inc()

    pFile.PrintText1(i.Inc(), '宛先', dict['宛先'], True, True)
    j = i.Get()
    pFile.DrowVerticalLine1(j, i.Get() - j + 1)

    i.Inc()

    pFile.PrintText1(i.Inc(), '1. 申請者', '')
    pFile.PrintText1(i.Inc(),'個人_社団の別', dict['個人_社団の別'], True, False)
    j = i.Get()
    pFile.PrintText1(i.Inc(), '郵便番号', dict['郵便番号'], True, False)
    pFile.PrintText1(i.Inc(), '住所', dict['住所'], True, False)
    pFile.PrintText1(i.Inc(), '電話番号', dict['電話番号'], True, False)
    if dict['国籍'] != '':
        pFile.PrintText1(i.Inc(), '国籍', dict['国籍'], True, False)
    if dict['個人_社団の別'] == '個人':    # 個人
        pFile.PrintText1(i.Inc(), '(代表者)氏名', dict['氏名'], True, False)
        pFile.PrintText1(i.Inc(), '上記フリガナ', dict['氏名フリガナ'], True, True)
    else:
        pFile.PrintText1(i.Inc(), '社団名', dict['社団名'], True, False)
        pFile.PrintText1(i.Inc(), '上記フリガナ', dict['社団名フリガナ'], True, False)
        pFile.PrintText1(i.Inc(), '(代表者)氏名', dict['氏名'], True, False)
        pFile.PrintText1(i.Inc(), '上記フリガナ', dict['氏名フリガナ'], True, True)
    pFile.DrowVerticalLine1(j, i.Get() - j + 1)

    i.Inc()

#    n = root.find('p:申請書/p:再免許情報', ns)         # 20190824 Delete
    pFile.PrintText1(i.Inc(), '2. 欠格事項', '')
    pFile.PrintText1(i.Inc(), '欠格事由の有無', dict['欠格事由の有無'], True, True)
    j = i.Get()
    pFile.DrowVerticalLine1(j, i.Get() - j + 1)

    i.Inc()

    pFile.PrintText1(i.Inc(), '3. 免許/再免許に関する事項', '')
    pFile.PrintText1(i.Inc(), '免許の番号', dict['免許の番号'], True, False)
    j = i.Get()
    if dict['呼出符号'] != '':
        pFile.PrintText1(i.Inc(), '呼出符号', dict['呼出符号'], True, False)        # 2019-08-25 修正
    pFile.PrintText1(i.Inc(), '免許の年月日', dict['免許の年月日'], True, False)
    pFile.PrintText1(i.Inc(), '備考', dict['備考'], True, True)
    pFile.DrowVerticalLine1(j, i.Get() - j + 1)

    i.Inc()

    pFile.PrintText1(i.Inc(), '4. 電波利用料', '')
    pFile.PrintText1(i.Inc(), '前納_有無', dict['前納_有無'], True, False)
    j = i.Get()
    pFile.PrintText1(i.Inc(), '前納にかかる期間', dict['前納期間'], True, True)
    pFile.DrowVerticalLine1(j, i.Get() - j + 1)

    i.Inc()

    n = root.find('p:申請書/p:申請者等情報/p:申請に関する連絡責任者', ns)
    pFile.PrintText1(i.Inc(), '5. 連絡先', '')
    pFile.PrintText1(i.Inc(), '氏名', dict['連絡先氏名'], True, False)
    j = i.Get()
    pFile.PrintText1(i.Inc(), '上記フリガナ', dict['連絡先氏名フリガナ'], True, False)
    pFile.PrintText1(i.Inc(), '電話番号', dict['連絡先電話番号'], True, False)
    pFile.PrintText1(i.Inc(), 'E-Mailアドレス', dict['電子メールアドレス'], True, True)
    pFile.DrowVerticalLine1(j, i.Get() - j + 1)

    if dict['申請書バージョン'] >= '0009':
        i.Inc()

        n = root.find('p:申請書/p:申請手数料', ns)
        pFile.PrintText1(i.Inc(), '申請手数料等', '')
        k = int('0' + (n.find('p:手数料額', ns).text or ''))
        pFile.PrintText1(i.Inc(), '申請手数料', "{:,}".format(dict['手数料額']) + ' 円　　電子納付', True, False)
        j = i.Get()
        pFile.PrintText1(i.Inc(), '免許状受取方法', dict['免許状受取方法'], True, True)
        pFile.DrowVerticalLine1(j, i.Get() - j + 1)

    pFile.Finalize()



def Haikyoku(root, pFile, dict):
#    dict = XmlCommon(root)
    if dict['ERROR']:
        OutErrMsg(5)      # ファイル形式不正
        return
    pFile.PrintHeader()
    i = Count(0)
    pFile.PrintText0(i.Get(), 200, dict['手続き名'])

    i.Inc()

    pFile.PrintText1(i.Inc(), '宛先', dict['宛先'], False, False)
    j = i.Get()
    pFile.DrowVerticalLine1(j, i.Get() - j + 1)

    i.Inc()

    pFile.PrintText1(i.Inc(), '1. 届出者', '')
    pFile.PrintText1(i.Inc(),'個人_社団の別', dict['個人_社団の別'], False, False)
    j = i.Get()
    pFile.PrintText1(i.Inc(), '郵便番号', dict['郵便番号'], True, False)
    pFile.PrintText1(i.Inc(), '住所', dict['住所'], True, False)
    pFile.PrintText1(i.Inc(), '電話番号', dict['電話番号'], True, False)
    if dict['国籍'] != '':
        pFile.PrintText1(i.Inc(), '国籍', dict['国籍'], True, False)
    if dict['個人_社団の別'] == '個人':    # 個人
        pFile.PrintText1(i.Inc(), '氏名', dict['氏名'], True, False)
        pFile.PrintText1(i.Inc(), '上記フリガナ', dict['氏名フリガナ'], True, True)
    else:
        pFile.PrintText1(i.Inc(), '社団名', dict['社団名'], True, False)
        pFile.PrintText1(i.Inc(), '上記フリガナ', dict['社団名フリガナ'], True, False)
        pFile.PrintText1(i.Inc(), '代表者氏名', dict['氏名'], True, False)
        pFile.PrintText1(i.Inc(), '上記フリガナ', dict['氏名フリガナ'], True, False)
    pFile.DrowVerticalLine1(j, i.Get() - j + 1)

    i.Inc()

    pFile.PrintText1(i.Inc(), '2. 廃止に係る事項', '')
    pFile.PrintText1(i.Inc(), '呼出符号', dict['呼出符号'], False, False)
    j = i.Get()
    pFile.PrintText1(i.Inc(), '免許の番号', dict['免許の番号'], True, False)
    pFile.PrintText1(i.Inc(), '廃止年月日', dict['廃止年月日'], True, False)
    pFile.PrintText1(i.Inc(), '備考', dict['備考'], True, False)
    pFile.DrowVerticalLine1(j, i.Get() - j + 1)

    i.Inc()

    pFile.PrintText1(i.Inc(), '3. 連絡先', '')
    pFile.PrintText1(i.Inc(), '氏名', dict['連絡先氏名'], False, False)
    j = i.Get()
    pFile.PrintText1(i.Inc(), '上記フリガナ', dict['連絡先氏名フリガナ'], True, False)
    pFile.PrintText1(i.Inc(), '電話番号', dict['連絡先電話番号'], True, False)
    pFile.PrintText1(i.Inc(), 'E-Mailアドレス', dict['電子メールアドレス'], True, False)
    pFile.DrowVerticalLine1(j, i.Get() - j + 1)

    pFile.Finalize()


def KaikyokuHenkou(root, pFile, dict):
#    dict = XmlCommon(root)
    if dict['ERROR']:
        OutErrMsg(5)      # ファイル形式不正
        return
    pFile.PrintHeader()
    i = Count(0)
    pFile.PrintText0(i.Get(), 200, dict['手続き名'])

    i.Inc()

    pFile.PrintText1(i.Inc(), '宛先', dict['宛先'], False, False)
    j = i.Get()
    pFile.DrowVerticalLine1(j, i.Get() - j + 1)

    i.Inc()

    pFile.PrintText1(i.Inc(), '1. 申請者', '')
    pFile.PrintText1(i.Inc(),'個人_社団の別', dict['個人_社団の別'], True, False)
    j = i.Get()
    pFile.PrintText1(i.Inc(), '郵便番号', dict['郵便番号'], True, False)
    pFile.PrintText1(i.Inc(), '住所', dict['住所'], True, False)
    pFile.PrintText1(i.Inc(), '電話番号', dict['電話番号'], True, False)
    if dict['国籍'] != '':
        pFile.PrintText1(i.Inc(), '国籍', dict['国籍'], True, False)
    if dict['個人_社団の別'] == '個人':    # 個人
        pFile.PrintText1(i.Inc(), '氏名', dict['氏名'], True, False)
        pFile.PrintText1(i.Inc(), '上記フリガナ', dict['氏名フリガナ'], True, True)
    else:
        pFile.PrintText1(i.Inc(), '社団名', dict['社団名'], True, False)
        pFile.PrintText1(i.Inc(), '上記フリガナ', dict['社団名フリガナ'], True, False)
        pFile.PrintText1(i.Inc(), '(代表者)氏名', dict['氏名'], True, False)
        pFile.PrintText1(i.Inc(), '上記フリガナ', dict['氏名フリガナ'], True, True)
    pFile.DrowVerticalLine1(j, i.Get() - j + 1)

    if dict['手続き'] == 'D051':       # 開局の時のみ
        i.Inc()

        n = root.find('p:申請書/p:再免許情報', ns)
        pFile.PrintText1(i.Inc(), '2. 欠格事項', '')
        pFile.PrintText1(i.Inc(),'欠格事由の有無', dict['欠格事由の有無'], True, True)
        j = i.Get()
        pFile.DrowVerticalLine1(j, i.Get() - j + 1)

    if dict['手続き'] == 'D053':       # 変更の時のみ
        i.Inc()

        pFile.PrintText1(i.Inc(), '2. 無線局に関する事項', '')
        pFile.PrintText1(i.Inc(), '呼出符号', dict['呼出符号'], True, False)
        j = i.Get()
        pFile.PrintText1(i.Inc(), '免許の番号', dict['免許の番号'], True, False)
        pFile.PrintText1(i.Inc(), '備考', dict['備考'], True, True)
        pFile.DrowVerticalLine1(j, i.Get() - j + 1)

    if dict['手続き'] == 'D051':       # 開局の時のみ
        i.Inc()
        n = root.find('p:申請書/p:電波利用料の前納の申出', ns)
        pFile.PrintText1(i.Inc(), '4. 電波利用料', '')
        pFile.PrintText1(i.Inc(), '申出の有無', dict['前納_有無'], True, False)
        j = i.Get()
        if dict['前納_有無'] != '無':
            pFile.PrintText1(i.Inc(), '前納期間', dict['前納期間'], True, True)
        pFile.DrowVerticalLine1(j, i.Get() - j + 1)

    i.Inc()
    if dict['申請書バージョン'] <= '0008':
        pass
    else:
        if dict['手続き'] == 'D051':
            z = '5'
        else:
            z = '3'
        n = root.find('p:申請書/p:申請者等情報/p:申請に関する連絡責任者', ns)
        pFile.PrintText1(i.Inc(), z + '. 連絡先', '')
        pFile.PrintText1(i.Inc(), '氏名', dict['連絡先氏名'], True, False)
        j = i.Get()
        pFile.PrintText1(i.Inc(), '上記フリガナ', dict['連絡先氏名フリガナ'], True, False)
        pFile.PrintText1(i.Inc(), '電話番号', dict['連絡先電話番号'], True, False)
        pFile.PrintText1(i.Inc(), 'E-Mailアドレス', dict['電子メールアドレス'], True, True)
        pFile.DrowVerticalLine1(j, i.Get() - j + 1)

#  本当はここに事項書

    if dict['手続き'] == 'D051':       # 開局の時のみ
        if dict['申請書バージョン'] >= '0009':
            i.Inc()
            n = root.find('p:申請書/p:申請手数料', ns)
            pFile.PrintText1(i.Inc(), '申請手数料/免許状受け取り方法', '')
            k = int('0' + (n.find('p:手数料_空中線電力', ns).text or ''))
            pFile.PrintText1(i.Inc(), '空中線電力', "{:,}".format(k) + ' W', True, False)
            j = i.Get()
            pFile.PrintText1(i.Inc(), '申請手数料', "{:,}".format(k) + ' 円　　電子納付', True, False)
            n = root.find('p:申請書/p:免許状受取方法', ns)
            d = ['', '返信用封筒別送', '窓口受領', '送料受取人払いによる受領(料金:500円)']
            k = int('0' + (n.find('p:免許状受取区分', ns).text or ''))
            pFile.PrintText1(i.Inc(), '免許状受取方法', d[k], True, True)
            pFile.DrowVerticalLine1(j, i.Get() - j + 1)

    pFile.NewPage()
    i = Count(0)
    Jikousho(root, pFile, dict)

    pFile.NewPage()
    i = Count(0)
    z = Sekkeisho1(root, pFile, dict, i)
    i.Set(z)
    Sekkeisho2(root, pFile, dict, i)

    pFile.Finalize()

def Jikousho(root, pFile, dict):
    n = root.find('p:申請書/p:事項書_工事設計書情報', ns)

    isChange10 = False
    isChange03 = False
    isChange07 = False
    isChange11 = False
    isChange12 = False
    isChange13 = False
    isChange16 = False
    changeItem = []
    if dict['手続き'] == 'D053':
        ｎ1 = n.find('p:変更項目', ns)
        if (n1.find('p:c.呼出符号', ns).text or '') == '1':
            isChange10 = True
            changeItem.append('3. 呼出符号')
        if (n1.find('p:c.申請_届出_者名等', ns).text or '') == '1':
            isChange03 = True
            changeItem.append('5. 申請者名等')
        if (n1.find('p:c.無線従事者免許証の番号', ns).text or '') == '1':
            isChange07 = True
            changeItem.append('8. 無線従事者免許証の番号')
        if (n1.find('p:c.無線設備の設置場所又は常置場所', ns).text or '') == '1':
            isChange11 = True
            changeItem.append('11. 設置/常置場所')
        if (n1.find('p:c.移動範囲', ns).text or '') == '1':
            isChange12 = True
            changeItem.append('12. 常置場所')
        if (n1.find('p:c.電波の型式並びに希望する周波数及び空中線電力', ns).text or '') == '1':
            isChange13 = True
            changeItem.append('13. 電波の型式/周波数/空中線電力')
        if (n1.find('p:c.工事設計書', ns).text or '') == '1':
            isChange16 = True
            changeItem.append('16. 工事設計書')

    i = Count(0)
#    pFile.PrintLine(i.Get(), '事項書 (申請書と同じ項目は表示していません)', '')
    pFile.PrintText1(i.Get(), '事項書', '')

    n1 = n.find('p:無線局の種別等', ns)
    if dict['手続き'] != 'D051':
        s = XmlMenkyo(n1.find('p:免許の番号', ns))
        pFile.PrintText1(i.Inc(), '1. 免許の番号', s, True, False)
        j = i.Get()
        pFile.PrintText1(i.Inc(), '2. 申請(届出)の区分', dict['申請区分'], True, False)
    else:
        pFile.PrintText1(i.Inc(), '2. 申請(届出)の区分', dict['申請区分'], True, False)
        j = i.Get()

    n1 = n.find('p:申請_届出_者名等', ns)
    s9 = n1.find('p:社団_個人の別', ns).text or ''
    if isChange03:
        pFile.textColor = red
    pFile.PrintText1(i.Inc(), '3. 社団/個人の別', ConfigGet('RADIO', s9), True, False)
    s = n1.find('p:郵便番号', ns).text or ''
    if isChange03:
        pFile.textColor = red
    pFile.PrintText1(i.Inc(), '4. 住所', '郵便番号 ' + s[0:3] + '-' + s[3:], True, False)
    s = n1.find('p:住所/p:都道府県_市区町村/p:都道府県', ns).text or ''
    s1 = GetJichitai(s)
    s = n1.find('p:住所/p:都道府県_市区町村/p:市区町村', ns).text or ''
    s2 = GetJichitai(s)
    s3 = n1.find('p:住所/p:町_丁目', ns).text or ''
    pFile.PrintText1(i.Inc(), '', s1 + s2 + s3, False, False)
    pFile.PrintText1(i.Inc(), '', '電話番号 ' + n1.find('p:電話番号', ns).text or '', False, False)
    s = n1.find('p:国籍', ns).text or ''
    if s != '':
        pFile.PrintText1(i.Inc(), '', '国籍 ' + GetJichitai(s), False, True)

    if s9 == '3':  # 個人
        s1 = n1.find('p:氏名_姓名/p:姓', ns).text or ''
        s2 = n1.find('p:氏名_姓名/p:名', ns).text or ''
        if isChange03:
            pFile.textColor = red
        pFile.PrintText1(i.Inc(), '5. 氏名', s1 + '' + s2, True, False)
        s1 = n1.find('p:氏名フリガナ_姓名/p:姓フリガナ', ns).text or ''
        s2 = n1.find('p:氏名フリガナ_姓名/p:名フリガナ', ns).text or ''
        pFile.PrintText1(i.Inc(), '', s1 + '' + s2, False, False)
    else:
        if isChange03:
            pFile.textColor = red
        pFile.PrintText1(i.Inc(), '5. 社団名', n1.find('p:社団_クラブ_名', ns).text or '', True, False)
        pFile.PrintText1(i.Inc(), '', n1.find('p:社団_クラブ_名フリガナ', ns).text or '', False, False)
        s1 = n1.find('p:代表者名_姓名/p:姓', ns).text or ''
        s2 = n1.find('p:代表者名_姓名/p:名', ns).text or ''
        pFile.PrintText1(i.Inc(), '  代表者', s1 + '' + s2, False, False)
        s1 = n1.find('p:代表者名フリガナ_姓名/p:姓フリガナ', ns).text or ''
        s2 = n1.find('p:代表者名フリガナ_姓名/p:名フリガナ', ns).text or ''
        pFile.PrintText1(i.Inc(), '', s1 + '' + s2, False, False)

    if dict['手続き'] != 'D051':
        n1 = n.find('p:無線局の種別等/p:工事落成の予定期日', ns)
        s = n1.find('p:工事落成の予定期日_区分', ns).text or ''
        if s != '':
            if s == '1':
                s = '日付指定  ' + XmlDate(n1.find('p:日付指定', ns)) or ''
            elif s == '2':
                s = '予備免許の日から  ' + (n1.find('p:月目', ns).text or '') + ' 月目の日'
            elif s == '3':
                s = '予備免許の日から  ' + (n1.find('p:日目', ns).text or '') + ' 日目の日'
            pFile.PrintText1(i.Inc(), '6. 工事落成の予定日', s, True, False)

    n1 = n.find('p:目的等/p:無線従事者免許証の番号', ns)
    s1 = n1.find('p:番号/p:上位1', ns).text or ''
    s2 = n1.find('p:番号/p:下位1', ns).text or ''
    if dict['申請書バージョン'] >= '0009':
        s3 = n1.find('p:c.施行規則第34条の8に規定する_外国政府の証明書', ns).text or ''
    else:
        s3 = ''
    if s1 != '':
        s = s1
        if s2 != '':
            s = s + '-' + s2
    else:
        s = s3
    if isChange07:
        pFile.textColor = red
    pFile.PrintText1(i.Inc(), '7. 無線従事者免許証の番号', s, True, False)

    if dict['手続き'] != 'D051':
        n1 = n.find('p:無線局の種別等', ns)
        if isChange10:
            pFile.textColor = red
        pFile.PrintText1(i.Inc(), '10. 呼出符号', n1.find('p:呼出符号', ns).text or '', True, False)

    ｎ1 = n.find('p:設置場所等/p:無線設備の設置場所又は常置場所/p:住所', ns)
    s1 = GetJichitai(n1.find('p:都道府県_市区町村/p:都道府県', ns).text or '')
    s2 = GetJichitai(n1.find('p:都道府県_市区町村/p:市区町村', ns).text or '')
    s3 = n1.find('p:設置場所_町_丁目', ns).text or ''
    if isChange11:
        pFile.textColor = red
    pFile.PrintText1(i.Inc(), '11. 設置/常置場所', s1 + s2 + s3, True, False)

    ｎ1 = n.find('p:設置場所等', ns)
    s = n1.find('p:移動範囲', ns).text or ''
    if isChange12:
        pFile.textColor = red
    pFile.PrintText1(i.Inc(), '12. 移動範囲', ConfigGet('IDO', s), True, False)

    if isChange13:
        pFile.textColor = red
    pFile.PrintText1(i.Inc(), '13. 電波型式/周波数/空中線電力', '', True, False)
    ｎ1 = n.find('p:周波数', ns)
    s2 = ''
    for n2 in n1.findall('p:周波数情報', ns):
        s1 = ConfigGet('BAND', n2.find('p:周波数帯', ns).text or '')
        s21 = n2.find('p:記号', ns).text or ''    # 3MA,2HC等の記号

        if s21 != '':
            s2 += (',' + s21)
        for n3 in n2.findall('p:電波の型式等情報', ns):     # 3MA,2HC等の記号に含まれない電波型式
            s21 = n3.find('p:電波の型式', ns).text or ''
            if s21 == '':
                break
            s2 += (',' + s21)
            s22 = n3.find('p:占有周波数帯幅', ns).text or ''
            if s22 != '':
                s2 += ('(' + s22 + ')')
        s2 = s2[1:]                                         # 電波の型式を全部まとめたもの

        f = float('0' + (n2.find('p:空中線電力', ns).text or ''))
        un = ['kW', 'W', 'mW', 'nW', 'uW']
        s3 = ''
        for x, u in enumerate([1000, 1, 0.001, 0.000001, 0.000000001]):
            f1 = f // u
            if f1 != 0:
                s3 = str(int(f / u)) + un[x]                # 空中電電力
                break

        z = SplitText(pFile, s2, abs(pFile.LinePos[1]) - abs(pFile.textPos3[2]) - 5)    # 最初行の最後まで長さで分けてみる
        for j in range(len(z)):
            if j == len(z) - 1:
                y = SplitText(pFile, z[j], abs(pFile.textPos3[3]) - abs(pFile.textPos3[2]) - 5)  # 空中線電力が行内で印刷できるか？
                if len(y) == 1:
                    pFile.PrintText3(i.Inc(), '', s1, z[j], s3, False, False, True)
                elif len(y) == 2:
                    pFile.PrintText3(i.Inc(), '', s1, z[j], '', False, False, True)
                    pFile.PrintText3(i.Inc(), '', '', '', s3, False, False, True)
        #       else:
        #            pass    # このケースは無いはず
            else:
                pFile.PrintText3(i.Inc(), '', s1, z[j], '', False, False, True)
                s1 = ''
        s2 = ''

    pFile.DrowHorizontalLine(i.Get(), False, True)

    if dict['手続き'] == 'D053':
        s = ", ".join(changeItem)
        pFile.PrintText1(i.Inc(), '14. 変更する項目', s, False, True)

    pFile.PrintText1(i.Inc(), '15. 備考', '', True, True)
    j = j + 1
    pFile.DrowVerticalLine1(j, i.Get() - j + 1)


def Sekkeisho1(root, pFile, dict, i):
    def Sekkeisho1_1(root, pFile, printing=True):
        ver = dict['申請書バージョン']
        n = root  # rootは工事設計1

        n1 = n.find('p:装置の区別等', ns)
        s1 = n1.find('p:装置の区別', ns).text or ''

        a = dict['免許の番号']
        c = ConfigGet(a, 'unit' + s1)
        e = ConfigGet(a, 'comment' + s1)
        pFile.PrintText4(i.Get(), c, e, printing)       # 機種名等をプリント

        s2 = ConfigGet('HENKO', n1.find('p:変更の種別', ns).text or '')
        s3 = XmlGiteki(n1.find('p:技術基準適合証明番号', ns), ver)
        pFile.PrintText2(i.Inc(), s1, s2, s3, False, False, printing)
        j = i.Get()
        if printing:
            pFile.DrowVerticalLine2(j, i.Get() - j + 1)

        noneGiteki = False
        for n2 in n1.findall('p:発射可能な電波の型式及び周波数の範囲情報', ns):
            s2 = ConfigGet('BAND', n2.find('p:周波数帯', ns).text or '')  # 最初の周波数帯だけ判断すればいいはずだが？
            if s2 == '':
                noneGiteki = False
            else:
                noneGiteki = True
            break

        if noneGiteki:       # 技適以外の時
            s1 = '電波型式/周波数'
            u = True
            for n2 in n1.findall('p:発射可能な電波の型式及び周波数の範囲情報', ns):
                s2 = ConfigGet('BAND', n2.find('p:周波数帯', ns).text or '')
                s3 = ''
                for n3 in n2.findall('p:電波の型式等情報', ns):
                    s3 = s3 + ',' + (n3.find('p:電波の型式', ns).text or '')
                    s4 = n3.find('p:占有周波数帯幅', ns).text or ''
                    if s4 != '':
                        s3 = s3 + '(' + s4 + ')'
                s3 = s3[1:]
                z = SplitText(pFile, s3, abs(pFile.LinePos[1]) - abs(pFile.textPos3[2] - 5))
                for s3 in z:
                    pFile.PrintText3(i.Inc(), s1, s2, s3, '', u, False, printing)
                    s1 = ''
                    s2 = ''
                    u = False
                s1 = ''

            s1 = '変調方式'
            u = True
            for n2 in n.findall('p:変調方式', ns):
                s2 = n2.find('p:電波の型式', ns).text or ''
                s3 = ConfigGet('MODULATION', n2.find('p:変調方式_コード', ns).text or '')
                s4 = n2.find('p:変調方式_備考', ns).text or ''
                if s4 != '':
                    s3 = s3 + '(' + s4 + ')'
                pFile.PrintText3(i.Inc(), s1, s2, s3, '', u, False, printing)
                s1 = ''
                u = False

            n1 = n.find('p:終段管', ns)
            s1 = '終段管/電圧'
            u = True
            t1 = []
            t2 = []
            for n2 in n1.findall('p:名称個数情報', ns):
                s2 = (n2.find('p:終段管_名称', ns).text or '') + ' × ' + (n2.find('p:終段管_個数', ns).text or '')
                t1.append(s2)
            for n2 in n1.findall('p:電圧情報', ns):
                t2.append((n2.find('p:電圧', ns).text or '') + ' V')
            for k in range(len(t1)):
                if t1[k] != '':
                    if k < len(t2):
                        s9 = t2[k]
                    else:
                        s9 = ''
                    pFile.PrintText3(i.Inc(), s1, t1[k], '', s9, u, False, printing)
                    u = False
                    s1 = ''

            s1 = '定格出力'
            s2 = ''
            for n2 in n.findall('p:定格出力', ns):
                s2 = s2 + ',' + XmlPower((n2.find('p:定格出力値', ns).text or ''))
            if s2 != '':
                s2 = s2[1:]
            pFile.PrintText3(i.Inc(), s1, s2, '', '', True, True, printing)

            n2 = n.find('p:添付書類_工事設計書', ns)
            s2 = n2.find('p:書類種別', ns) or ''
            s3 = n2.find('p:添付ファイル名', ns) or ''
            if s3 != '':  # 添付書類があるときのみ印刷する
                if s2 == '205':
                    s3 = '送信機系統図  ' + s3
                else:
                    s3 = 'その他　　　  ' + s3
                pFile.PrintText3(i.Inc(), '添付書類', s3, '', '', True, False, printing)
                s2 = n2.find('p:通信欄', ns) or ''
                pFile.PrintText3(i.Inc(), '通信欄', s2, '', '', True, True, printing)

            if printing:
                pFile.DrowVerticalLine1(j, i.Get() - j + 1)
        return i.Get() - j  # 印刷した行数を返す
#    def Sekkeisho1_1 の終わり

    n = root.find('p:申請書/p:事項書_工事設計書情報', ns)
    pFile.PrintText1(i.Get(), '16. 工事設計書(1)', '')

    u = True
    for n1 in n.findall('p:工事設計1', ns):
        k = i.Get()
        z = Sekkeisho1_1(n1, pFile,  False)
        i.Set(k)
        if z > pFile.maxRowsCount - i.Get():
            pFile.NewPage()
            i = Count(0)
            u = True
        if not u:
            i.Inc()
        Sekkeisho1_1(n1, pFile, True)
        u = False
    return i.Get()


def Sekkeisho2(root, pFile, dict, i):
    def Sekkeisho2_1(root, pFile, printing=True):
        i.Inc()
        n = root  # 工事設計2
        pFile.PrintText3(i.Inc(), '16. 工事設計書(2)', '', '', '', False, False, printing)

        s1 = '空中線の型式'
        s2 = ''
        for n1 in n.findall('p:送信空中線の型式情報', ns):
            s = ConfigGet('ANTENNA', n1.find('p:送信空中線の型式', ns).text or '')
            s2 += (',' + s)
        s2 = s2[1:]
        pFile.PrintText3(i.Inc(), s1, s2, '', '', True, False, printing)
        j = i.Get()

        s2 = ConfigGet('COUNTER', n.find('p:周波数測定装置の有無', ns).text or '')
        pFile.PrintText3(i.Inc(), '周波数測定装置の有無', s2, '', '', True, False, printing)

        s = n.find('p:その他の工事設計', ns).text or ''
        s2 = ''
        if s == '1':
            s2 = '電波法第３章に規定する条件に合致する。'
        pFile.PrintText3(i.Inc(), 'その他の工事', s2, '', '', True, False, printing)

        if printing:
            pFile.DrowVerticalLine1(j, i.Get() - j + 1)

        return i.Get() - j      # 印刷した行数を返す
#   def Sekkeisho2_1(root, pFile, printing=True):の終わり

    n = root.find('p:申請書/p:事項書_工事設計書情報/p:工事設計2', ns)

    k = i.Get()
    z = Sekkeisho2_1(n, pFile, False)
    i.Set(k)
    if z > pFile.maxRowsCount - i.Get():
        pFile.NewPage()
        i = Count(0)
    Sekkeisho2_1(n, pFile, True)



def DXmlToPdf(fileName):
    g = os.path.splitext(os.path.basename(fileName))        # ファイル名と拡張子を取得
    if (g[1] != '.zip') and (g[1] != '.xml'):
        OutErrMsg(2, fileName)      # 拡張子不正
        return False

    if not os.path.isfile(fileName):
        OutErrMsg(3, fileName)      # ファイルの存在
        return False

    if g[1] == '.zip':
        with zipfile.ZipFile(fileName, "r") as zips:
            f = False
            for fn in zips.namelist():
                fnU = fn.upper()
                if (fnU == "SINSEI.XML") or (fnU == "SHINSEI.XML"):
                    f = True
                    xmlString = zips.read(fn)
            if not f:
                OutErrMsg(6, "SHINSEI.XML")  # ファイルの存在
                return False
    else:
        f = open(fileName, encoding='utf-8')
        xmlString = f.read()
        f.close()
    root = ET.fromstring(xmlString)  # rootの設定

# pdfファイルが他プロセス瀬使われていないことを確認する必要がある
    g = os.path.splitext(fileName)        # パス名と拡張子を取得
    fn = g[0] + '.pdf'
    try:
        with open(fn, 'w') as fo:       # 試しに出力で開いてみる
            fo.close
    except Exception as e:
        OutErrMsg(4, fn)  # ファイルが使われている
        return False

    pFile = DPdfEdit.Pdf(fn)
    pFile.fileName = fileName
    dict = XmlCommon(root)

    tetsuzuki = XmlTetsuzuki(root)[0]
    if tetsuzuki == 'D052':     # 再免許
        Saimen(root, pFile, dict)
    elif tetsuzuki == 'D055':    # 廃局
        Haikyoku(root, pFile, dict)
    else:                       # 開局・変更
        KaikyokuHenkou(root, pFile, dict)

    ConfigSet('main', 'fileName', fileName)
    ConfigSet('main', 'Application', dict['手続き名'])
    ConfigSet('main', 'license', dict['免許の番号'])
    ConfigSet('main', 'callsign', dict['呼出符号'])
#    pFile.AddPage(fileName, 'test.pdf')    # 通算ページNo表示のため、　まだ正常じゃない
    if isBrowse:
        webbrowser.open(fn)

# argparse.ArgumentParserのWrapperクラス
# 実行時にファイル名が指定されているかどうかを検出するため
class MyArgumentParser(argparse.ArgumentParser):
    def _print_message(self, message, file=None):
        pass

if __name__ == "__main__":
    isSlave = False
    config.remove_section('ERROR')
    try:
        config = configparser.ConfigParser()
        config.read(iniFileName, 'utf-8-sig')
    except configparser.MissingSectionHeaderError as e:
        config.read(iniFileName, 'utf-8')
    except UnicodeDecodeError as e:
        config.read(iniFileName, 'ansi')
    except SystemExit as e:
        OutErrMsg(7, iniFileName)

    ConfigSet('Args', 'TEST', 'args.FileName')
    try:
        parser = argparse.ArgumentParser()
        parser.add_argument('FileName', help='Input file name')
        parser.add_argument('-b', '--browse', help='Browse rusult file', action='store_true')
        args = parser.parse_args()
        FileName = args.FileName
        isBrowse = args.browse

#        inFileName = r"C:\Users\Kunio\Python Projects\DXmlToPdf\shinsei_E19-0000120282-D.zip"    # for DEBUG
#        isBrowse = True                                                                          # for DEBUG

# Iniファイルにファイル名等を書き込む
        ConfigSet('Args', 'FileName', inFileName)
        ConfigSetBool('Args', 'Browse', isBrowse)
        DXmlToPdf(inFileName)

    except SystemExit as e:
        OutErrMsg(1, inFileName)        # エラー理由をINIファイルに書き込み

    with open(iniFileName, 'w', encoding='utf-8-sig') as configfile:
        config.write(configfile)

