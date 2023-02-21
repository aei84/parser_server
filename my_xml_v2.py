#import xml.etree.ElementTree as ET
import lxml.etree as ET
from zipfile import ZipFile
import socket
import struct
import sys
#import os
from pathlib import *
import shutil

import openpyxl
from openpyxl.styles import PatternFill, Border, Side

from copy import copy

IED = "{http://www.iec.ch/61850/2003/SCL}IED"
PRIVATE = "{http://www.iec.ch/61850/2003/SCL}Private"
ACCESSPOINT = "{http://www.iec.ch/61850/2003/SCL}AccessPoint"
SERVER = "{http://www.iec.ch/61850/2003/SCL}Server"
LDEVICE = "{http://www.iec.ch/61850/2003/SCL}LDevice"
LN0 = "{http://www.iec.ch/61850/2003/SCL}LN0"
DATASET = "{http://www.iec.ch/61850/2003/SCL}DataSet"
HEADER = "{http://www.iec.ch/61850/2003/SCL}Header"

LN = "{http://www.iec.ch/61850/2003/SCL}LN"
DOI = "{http://www.iec.ch/61850/2003/SCL}DOI"
DAI = "{http://www.iec.ch/61850/2003/SCL}DAI"

COMMUNICATION = '{http://www.iec.ch/61850/2003/SCL}Communication'
SUBNETWORK = "{http://www.iec.ch/61850/2003/SCL}SubNetwork"
CONNECTEDAP = "{http://www.iec.ch/61850/2003/SCL}ConnectedAP"
ADDRESS = "{http://www.iec.ch/61850/2003/SCL}Address"
GSE = "{http://www.iec.ch/61850/2003/SCL}GSE"
P = "{http://www.iec.ch/61850/2003/SCL}P"

GSECONTROL = "{http://www.iec.ch/61850/2003/SCL}GSEControl"
FCDA = "{http://www.iec.ch/61850/2003/SCL}FCDA"

INPUTS = "{http://www.iec.ch/61850/2003/SCL}Inputs"
EXTREF = "{http://www.iec.ch/61850/2003/SCL}ExtRef"

CIPA = "{http://www.iec.ch/61850/2003/SCL}CIPA"

HORIZ = 3 # координаты начала заполнения EXCEL
VERT = 3 # координаты начала заполнения EXCEL

class Report_cipa():
    def __init__(self):
        self.err = set()
        self.ok_term = set()
    def make_report(self, path_report: Path):
        #print(len(err))
        self.err.discard(None)
        with open(path_report/"report_cipa.txt", 'w', encoding='utf-8') as file:
            file.write("Ненайденных терминалов: " + str(len(self.err)) + "\n")
            for i in self.err:
                file.write("\t" + i + "\n")
            file.write("Обработанных терминалов :" + str(len(self.ok_term)) + "\n")
            for i in self.ok_term:
                file.write("\t" + i + "\n")

def whose_cid(cid: str):
    print("whose_cid")
    #print("\n!!!", cid)
    try:
        tree = ET.parse(cid)
    except:
        p = cid.rsplit("/", 2)
        print("Не удалось открыть", cid)
        err = open (p[0] + '/' + "no_read.txt", 'a')
        err.write(p[-1])
        err.write("\n")
        err.close()
        return "Не удалось открыть"
    #print("Определяем чей cid ",cid)
    root = tree.getroot()
    #print(root.nsmap)
    try:
        manufacturer = root.find(f"{IED}").attrib["manufacturer"] == "EKRA"
    except:
        #print("Неизвестное устройство")
        return "Неизвестное устройство"
    if manufacturer:
        #print("This is EKRA")
        return "IEC61850_TERMINAL"
        return "EKRA"

    else:
        #print("IEC61850_TERMINAL")
        return "IEC61850_TERMINAL"


    tree = ET.parse(file_name)
    print("\n Читаем ",file_name)
    root = tree.getroot()
    print(type(root))
    #root_ = tree.getroot()

    #header_id = root.find(f"{HEADER}").attrib["id"]
    ied = root.find(f"{IED}") # считаю что нужный мне IED первый. В АВВ иногда не так, потом разберемся
    ied_name = ied.attrib["name"]
    #header_id = ied.attrib["name"]
    for ln in root:
        #print(ln.tag)
        if ln.tag == COMMUNICATION:
            # print("Communication")
            continue
        if ln.tag == IED and ln.get("name") == ied_name:
            # print("ied_name")
            continue
        root.remove(ln)
    #print(root.find(f".//{CONNECTEDAP}/{GSE}").attrib)
    communication = root.find(f"{COMMUNICATION}")
    print(type(communication))
    test = communication.find(f".//{GSE}")
    for i in test.iterancestors():
        print(i.tag)
    #node = communication.xpath(f"//{GSE}")[0]
    node = root.xpath(f"//{GSE}")[0]
    ancestors = list(node.iterancestors())
    for elem in root.iter():
        if elem not in ancestors and elem.tag != GSE:
            elem.getparent().remove(elem)

    # r = reversed(list(communication.find(f".//{CONNECTEDAP}[@iedName='{ied_name}']").iterancestors())[:-1])
    # for parent in r:    
    #     print(parent.tag, parent.attrib)    
    #     # if parent is not communication:
    #     #     print("\t", parent.tag)
    #     for child in parent:
    #         flag = True
    #         print("\t", child.tag, child.attrib)
    #         if child.find(f".//{GSE}"):
    #             print("\t\t", child.tag, child.attrib)
    #             if child.tag == CONNECTEDAP and child.attrib["iedName"] == ied_name:                        
    #                 flag = False
    #                 print("\t\t\tFalse", child.tag, child.attrib)
    #             if flag:
    #                 print("\t\t\tttdel", child.tag, child.attrib)
    #                 parent.remove(child)
    #             else:
    #                 print("\t\t\tno del", child.tag, child.attrib)
    #                 flag = True
    #         else:
    #             r.remove(parent)
    # # #         #parent.getparent().remove(parent)

    return root
            # terminal.communication["IP_GOOSE"] = ip.text
            # print(terminal.communication["IP_GOOSE"])
    # бежим по исходящим гусям терминала
    for gse in root.findall(f"{COMMUNICATION}/{SUBNETWORK}/{CONNECTEDAP}[@iedName='{header_id}']/{GSE}"):
        goose = Goose()
        # набиваем гуся параметрами
        for p in gse.findall(f"{ADDRESS}/{P}"):
            goose.goose_out_param[p.attrib["type"]] = p.text
        ldInst = gse.attrib["ldInst"] # <GSE ldInst="UD1" cbName="Control_DataSet">
        cdName = gse.attrib["cbName"] # <GSE ldInst="UD1" cbName="Control_DataSet">
        datSet = root.find(f"{IED}[@name='{header_id}']/{ACCESSPOINT}/{SERVER}/{LDEVICE}[@inst='{ldInst}']/{LN0}/{GSECONTROL}[@name='{cdName}']").attrib["datSet"] # <GSEControl datSet="DataSet" confRev="20001" type="GOOSE" appID="1_W6_PA_S1_21A_QT" name="Control_DataSet">
        # набиваем гуся датасетами
        for fcda in root.findall(f"{IED}[@name='{header_id}']/{ACCESSPOINT}/{SERVER}/{LDEVICE}[@inst='{ldInst}']/{LN0}/{DATASET}[@name='{datSet}']/{FCDA}"):
            ldInst = fcda.attrib["ldInst"]
            prefix = fcda.attrib["prefix"]
            lnClass = fcda.attrib["lnClass"]
            inst = fcda.attrib["lnInst"]
            doName = fcda.attrib["doName"]
            ln = root.find(f"{IED}[@name='{header_id}']/{ACCESSPOINT}/{SERVER}/{LDEVICE}[@inst='{ldInst}']/{LN}[@prefix='{prefix}'][@lnClass='{lnClass}'][@inst='{inst}']").get("desc") # у ЭКРЫ тут нет атрибута desc
            # print("cdName =", cdName)
            # print("ldInst =", ldInst)
            # print("prefix =", prefix)
            # print("lnClass =", lnClass)
            # print("inst =", inst)
            # print("doName =", doName)
            # print("ln =", ln)
            try: # В сиде ДЗШ первый гусь нормально описан, второй нет, надо разобраться
                doi = root.find(f"{IED}[@name='{header_id}']/{ACCESSPOINT}/{SERVER}/{LDEVICE}[@inst='{ldInst}']/{LN}[@prefix='{prefix}'][@lnClass='{lnClass}'][@inst='{inst}']/{DOI}[@name='{doName}']").get("desc")
                if not doi: # У ЭКРЫ глубже название сигнала
                    doi = root.find(f"{IED}[@name='{header_id}']/{ACCESSPOINT}/{SERVER}/{LDEVICE}[@inst='{ldInst}']/{LN}[@prefix='{prefix}'][@lnClass='{lnClass}'][@inst='{inst}']/{DOI}[@name='{doName}']/{DAI}").get("desc")
            except:
                doi = None
            # print("doi =", doi)

            goose.goose_out_data.append({"ldInst": fcda.attrib["ldInst"], 
                                        "prefix": fcda.attrib["prefix"], 
                                        "lnClass": fcda.attrib["lnClass"], 
                                        "lnInst": fcda.attrib["lnInst"], 
                                        "doName": fcda.attrib["doName"], 
                                        "daName": fcda.attrib["daName"], 
                                        "fc": fcda.attrib["fc"],
                                        "ln": ln,
                                        "doi": doi})
        # добовляем исходящего гуся в терминал, ключ пара inst логического устройства и имя датасета
        terminal.goose_out[(ldInst, cdName)] = goose
    # пробегаем по подпискам
    #first_ied = root.find()
    if "sip4" in root.nsmap:
        p = ied.findall(f".//{INPUTS}/{EXTREF}") # если сипротек 4
    else:
        p = ied.findall(f".//{INPUTS}/{EXTREF}[@serviceType='GOOSE']") # любой другой терминал
    for inputs in p:
        subs = {}
        subs["iedName"] = inputs.get("iedName")
        if not subs["iedName"]:
            #subs["iedName"] = terminal.communication["iedName"] # У Экры в Inputs нет iedName, но его мы уже знаем, просто забираем из нашей структуры
            continue # У ЭКРЫ прописаны пусте исходящие гуси
        subs["srcCBName"] = inputs.attrib["srcCBName"]
        subs["ldInst"] = inputs.attrib["ldInst"]
        subs["prefix"] = inputs.get("prefix")
        subs["lnClass"] = inputs.attrib["lnClass"]
        subs["lnInst"] = inputs.attrib["lnInst"]
        subs["doName"] = inputs.attrib["doName"]
        subs["daName"] = inputs.attrib["daName"]
        subs["desc"] = inputs.get("desc") # в 4м сипротеке деска не будет. Поэтому идем по ифу ниже
        if not subs["desc"]:         
            if inputs.get("intAddr"):
                xxx = inputs.get("intAddr").split('/')[-2]
                # print(inputs.getparent().getparent().find(f"{DOI}[@name='{intAddr}']"))
                subs["desc"] = inputs.getparent().getparent().find(f"{DOI}[@name='{xxx}']").attrib["desc"]
        if ied.get("manufacturer") == "EKRA":
            xxx = inputs.get("intAddr").split('.')
            for doi_ in inputs.getparent().getparent().getparent():
                if "" + doi_.get("prefix", "") + doi_.get("lnClass", "") + doi_.get("inst", "") == xxx[0]:
                    subs["desc"] = doi_.find(f"{DOI}[@name='{xxx[1]}']/{DAI}").attrib["desc"]

        terminal.goose_in.append(subs)
        # terminal.goose_in.append({"ldInst": inputs.attrib["ldInst"],
        #                         "prefix": inputs.attrib["prefix"], 
        #                         "lnClass": inputs.attrib["lnClass"], 
        #                         "lnInst": inputs.attrib["lnInst"], 
        #                         "doName": inputs.attrib["doName"], 
        #                         "daName": inputs.attrib["daName"], 
        #                         #"fc": inputs.attrib["fc"],   
        #                         "iedName": inputs.attrib["iedName"]
        #                         # "desc": inputs.attrib["desc"]
        #                                                 })
    # else:
    #     for inputs in root.findall(f".//{INPUTS}/{EXTREF}[@serviceType='GOOSE']"):
    #         terminal.goose_in.append({"ldInst": inputs.attrib["ldInst"],
    #                                 "prefix": inputs.attrib["prefix"], 
    #                                 "lnClass": inputs.attrib["lnClass"], 
    #                                 "lnInst": inputs.attrib["lnInst"], 
    #                                 "doName": inputs.attrib["doName"], 
    #                                 "daName": inputs.attrib["daName"], 
    #                                 #"fc": inputs.attrib["fc"],   
    #                                 "iedName": inputs.attrib["iedName"],
    #                                 "desc": inputs.attrib["desc"]})
    print_terminal(terminal)
    
    # for private in root.findall(f"{IED}/{PRIVATE}"): # пробегаю по параметрам исходящего GOOSE
    #     if "EKRA-GOOSEOutParam-" in private.attrib["type"]:
    #         is_goose = True
    #         if private.attrib["type"] == "EKRA-GOOSEOutParam-all": # В ПАДСах отличается
    #             for p in private.text.split(";"):
    #                 p = p.split(":")
    #                 if len(p) == 2:
    #                     terminal.goose_out_param[int(p[0][1:])] = p[1]
    #         else:
    #             terminal.goose_out_param[int(private.attrib["type"][19:])] = private.text    
    #     elif "EKRA-GOOSEIn-" in private.attrib["type"]  and "-Param-" in private.attrib["type"]:
    #         tt = private.attrib["type"][13:].split("-Param-")
    #         tt[0], tt[1] = int(tt[0]), int(tt[1])
    #         terminal.goose_in.setdefault(tt[0], [None] * 11)[tt[1]] = private.text
    # for private in root.findall(f"{IED}/{ACCESSPOINT}/{SERVER}/{LDEVICE}/{LN0}/{DATASET}/{PRIVATE}"):
    #     if "type" in private.attrib:
    #         if private.attrib["type"] == "EKRA-DSNum-all": # В ПАДСах отличается
    #             for p in private.text.split(";"):
    #                 p = p.split(":")
    #                 if len(p) == 2:
    #                     terminal.goose_out_data[int(p[0][1:])] = p[1]
    #         else:
    #             terminal.goose_out_data[int(private.attrib["type"][11:])] = private.text
    # for g in terminal.goose_out_data:
    #     terminal.subscribers[g] = []

    # for dai in root.findall(f"{IED}/{ACCESSPOINT}/{SERVER}/{LDEVICE}/{LN}/{DOI}/{DAI}"):
    #     if "desc" in dai.attrib:
    #         try:
    #             k, v = dai.attrib["desc"].split(" - ", 1)
    #             terminal.signal_names[int(k)] = v
    #         except:
    #             pass
    # if is_goose:
    #     return terminal
    # else:
    #     return None
    return ied

def make_xl(cids, path):
    print("make_xl")
    #print(substation[0].signal_names)
    wb = openpyxl.Workbook()
    ws = wb.active
    horiz = HORIZ + 1 # первый столбец для заполнения горизонтали терминалов
    vert = VERT + 1 # первая строка для заполнения вертикали терминалов
    for t in cids:
        print("t =", t)
        terminal = ET.parse(path.parent/f"{t}.cid").getroot()

        #print("заполняем  xl", t.goose_out_param[5])
        ws.cell(VERT, horiz).value = t
        ws.cell(VERT, horiz).alignment = openpyxl.styles.Alignment(textRotation=90)
        #ws.cell(VERT - 2, horiz).value = t.communication["IP"] 
        #ws.cell(VERT - 2, horiz).alignment = openpyxl.styles.Alignment(textRotation=90)
        letter = ws.cell(VERT, horiz).column_letter # получаем букву текущей ячейки чтобы изменить ее ширину
        ws.column_dimensions[letter].width = 3 # меняем ширину текущего столбца
        horiz += 1
        start_group = ws.cell(VERT, horiz).column_letter # получаем букву текущей для группировки ячеек
        end_group = None # если у терминала не будет входящих гусей, то конец гусей не определиться и колонки не будут группироваться
        for extref in terminal.findall(f".//{IED}[@name='{t}']//{INPUTS}/{EXTREF}"): # бежим по всем входящим нашего терминала
        #for ind, goose in terminal.findall()
            ied_name = extref.get("iedName") # если iedName нет, значит гусь не используется, пропускаем
            if ied_name == None:
                continue
            serviceType = extref.get("serviceType") # если serviceType "SMV" (SV-поток), пропускаем
            if serviceType == "SMV":
                continue

            ws.cell(VERT-2, horiz).value = t #f"{extref.attrib['iedName']}"
            ws.cell(VERT-2, horiz).alignment = openpyxl.styles.Alignment(textRotation=90)
            
            ws.cell(VERT-1, horiz).value = f"{extref.attrib['iedName']}/{extref.attrib['ldInst']}/{extref.attrib['prefix']}{extref.attrib['lnClass']}{extref.attrib['lnInst']}/{extref.attrib['doName']}/{extref.attrib['daName']}"
            ws.cell(VERT-1, horiz).alignment = openpyxl.styles.Alignment(textRotation=90)
            ws.cell(VERT, horiz).value = f"{extref.attrib['intAddr']}"
            ws.cell(VERT, horiz).alignment = openpyxl.styles.Alignment(textRotation=90)
            letter = ws.cell(VERT, horiz).column_letter # получаем букву текущей ячейки чтобы изменить ее ширину
            ws.column_dimensions[letter].width = 3 # меняем ширину текущего столбца
            end_group = ws.cell(VERT, horiz).column_letter
            horiz += 1
        if end_group:
            ws.column_dimensions.group(start_group, end_group, hidden=True)
        ws.cell(vert, HORIZ).value = t
        #ws.cell(vert, HORIZ - 2).value = t.communication["IP"] 
        vert += 1
        start_group = vert
        end_group = None
        for gse in terminal.findall(f".//{COMMUNICATION}//{CONNECTEDAP}[@iedName='{t}']/{GSE}"): # прохожу по всем исходящим гусям
            cbName = gse.get("cbName") # получаю cbName чтобы потом найти нужный datSet
            ldInst = gse.get("ldInst") # получаю ldInst чтобы потом найти нужный datSet
            print("cbName =", cbName)
            datSet = terminal.find(f".//{IED}[@name='{t}']//{ACCESSPOINT}//{SERVER}//{LDEVICE}[@inst='{ldInst}']/{LN0}/{GSECONTROL}[@name='{cbName}']").get("datSet") # получил имя DataSet с набором данных гуся
            print("datSet =", datSet)
            for fcda in terminal.findall(f".//{IED}[@name='{t}']/{ACCESSPOINT}/{SERVER}/{LDEVICE}[@inst='{ldInst}']/{LN0}/{DATASET}[@name='{datSet}']/{FCDA}"): # бегу по исходящим сигналам гуся
                print(fcda.attrib)
                
                ws.cell(vert, HORIZ-2).value = t
                prefix = fcda.get("prefix", "")
                lnInst = fcda.get("lnInst", "")
                ws.cell(vert, HORIZ-1).value = f"{t}/{fcda.attrib['ldInst']}/{prefix}{fcda.attrib['lnClass']}{lnInst}/{fcda.attrib['doName']}/{fcda.attrib['daName']}"
                
                doName = fcda.attrib["doName"]
                cipa_name_signal = terminal.find(f".//{IED}[@name='{t}']//{LDEVICE}[@inst='{fcda.get('ldInst')}']/{LN}[@prefix='{fcda.get('prefix')}'][@lnClass='{fcda.get('lnClass')}'][@inst='{fcda.get('lnInst')}']/{DOI}[@name='{doName}']") # нахожу узел на который подписан [@doName='{doName}']
                #print(test1, test2, cipa_name_signal.attrib)
                #cipa = ET.Element(CIPA)
                #cipa.attrib["IED"] = ied_name
                #cipa.attrib["name"] = cipa_name_signal.get("name")
                try:
                    desc_in_DOI = cipa_name_signal.get("desc")
                except:
                    desc_in_DOI = "Рухнули на строчке 351"
                if desc_in_DOI:
                    ws.cell(vert, HORIZ).value = desc_in_DOI
                else:
                    try:
                        desc_in_DAI = cipa_name_signal.find(DAI).get("desc")
                        ws.cell(vert, HORIZ).value = desc_in_DAI
                    except:
                        ws.cell(vert, HORIZ).value = "Не удалось получить название сигнала"
                
                
                
                # try:
                #     ws.cell(vert, HORIZ).value = fcda.find(f"./{CIPA}").attrib["desc"]
                # except:
                #     ws.cell(vert, HORIZ).value = 'Проблема в 343  ws.cell(vert, HORIZ).value = fcda.find(f"./CIPA").attrib["desc"]'
                # ws.cell(vert, HORIZ - 1).value = signal
                # ws.cell(vert, HORIZ - 2).value = t.signal_names.get(int(signal), "В ЭКРАвских сидах не найти, бери мануал(((")
                end_group = vert
                vert += 1
                
        # for ind, signal in t.goose_out_data.items():
        #     ws.cell(vert, HORIZ).value = "ind_" + str(ind)
        #     ws.cell(vert, HORIZ - 1).value = signal
        #     ws.cell(vert, HORIZ - 2).value = t.signal_names.get(int(signal), "В ЭКРАвских сидах не найти, бери мануал(((")
        #     end_group = vert
        #     vert += 1
        if end_group:
            ws.row_dimensions.group(start_group, end_group, hidden=True) 
    ws.column_dimensions.group('A', ws.cell(VERT, HORIZ - 1).column_letter, hidden=True)
    ws.row_dimensions.group(1, VERT - 1, hidden=True)
    ws.column_dimensions[ws.cell(VERT, HORIZ - 2).column_letter].width = 22
    ws.column_dimensions[ws.cell(VERT, HORIZ - 1).column_letter].width = 5
    ws.column_dimensions[ws.cell(VERT, HORIZ).column_letter].width = 20
    ws.row_dimensions[VERT].height = 100
    ws.row_dimensions[VERT - 2].height = 110
    ws.freeze_panes = ws.cell(VERT + 1, HORIZ + 1)
    wb.save(path.parent/"For_Sancho.xlsx")
    fill_xl(path.parent/"For_Sancho.xlsx")
    #to_vasiliy_xl(path.parent/"For_Sancho.xlsx", "vas.xlsx")

def fill_xl(file_xl):
    print("fill_xl")
    wb = openpyxl.load_workbook(file_xl)
    ws = wb.active
    empty_vertical = [True] * (ws.max_row + 1)
    empty_horizontal = [True] * (ws.max_column + 1)
    horiz = HORIZ + 1 # первый столбец для заполнения пересечений
    vert = VERT + 1# первая строка для заполнения пересечений
    terminal_suppression = {} # ключи - названия терминалов записанные друг за друном, значения - адрес ячейки где терминалы пересекаются    
    border_color = 'C0C0C0' # Задаем цвет границ    
    cell_color = 'D3D3D3' # Задаем цвет границ    
    border_side = Side(style='thin', color=border_color)  # Создаем объект Side с нужным цветом    
    border = Border(left=border_side, right=border_side, top=border_side, bottom=border_side) # Создаем объект Border с нужными границами
    for x in range(horiz, ws.max_column + 1):
        for y in range(vert, ws.max_row + 1):
            ws.cell(y, x).border = border
            if ws.cell(y, HORIZ - 1).value == None:
                ws.cell(y, x).fill = PatternFill('solid', fgColor=cell_color)
            if ws.cell(VERT - 1, x).value == None:
                ws.cell(y, x).fill = PatternFill('solid', fgColor=cell_color)
                if ws.cell(y, HORIZ - 1).value == None:
                    # ws.cell(y, x).value = 'X'
                    terminal_suppression[ws.cell(VERT, x).value + ws.cell(y, HORIZ).value] = ws.cell(y, x)
                    continue
            if ws.cell(y, HORIZ - 1).value == ws.cell(VERT - 1, x).value:
                ws.cell(y, x).value = 'x'
                row = terminal_suppression[ws.cell(VERT - 2, x).value + ws.cell(y, HORIZ - 2).value].row
                column = terminal_suppression[ws.cell(VERT - 2, x).value + ws.cell(y, HORIZ - 2).value].column
                #terminal_suppression[ws.cell(VERT - 2, x).value + ws.cell(y, HORIZ - 2).value].value = 'X'
                ws.cell(row, column).value = 'X'
                ws.cell(row, x).value = '->'
                ws.cell(row, x).alignment = openpyxl.styles.Alignment(textRotation=180)
                ws.cell(y, column).value = '->'

                ws.cell(row, HORIZ).fill = PatternFill('solid', fgColor=cell_color)
                ws.cell(VERT, column).fill = PatternFill('solid', fgColor=cell_color)
                
                continue
    wb.save(file_xl)
    #return empty_vertical, empty_horizontal

def paint_xl(file_xl, empty_vertical, empty_horizontal):    
    print("paint_xl")
    wb = openpyxl.load_workbook(file_xl)
    ws = wb.active
    horiz = HORIZ + 1 # первый столбец для заполнения пересечений
    vert = VERT + 1# первая строка для заполнения пересечений
    print("красим входящие, которые не подписаны")
    for x in range(horiz, ws.max_column + 1):        
        if empty_horizontal[x]:
            ws.cell(VERT, x).fill = PatternFill('solid', fgColor="8A3324")
    print("красим исходящие, на которые подписаны")
    for y in range(vert, ws.max_row + 1):        
        if empty_vertical[y]:
            ws.cell(y, HORIZ).fill = PatternFill('solid', fgColor="8A3324")
    wb.save(file_xl)

def to_vasiliy_xl (sourсe: Path, dest: Path):
    print("to_vasiliy_xl")
    # загружаем файлы
    source_wb = openpyxl.load_workbook(sourсe) # открываем файл-источник
    dest_wb = openpyxl.load_workbook(dest) # открываем файл назначения
    # выбираем листы для работы
    source_ws = source_wb.active # выбираем активный лист в файле-источнике
    dest_ws = dest_wb.worksheets[0] # выбираем первый лист в файле назначения
    # копируем значения и форматирование
    for row in source_ws.iter_rows(min_row=1, max_row=source_ws.max_row,
                                    min_col=1, max_col=source_ws.max_column):  # проходимся по всем строкам в листе-источнике
        for cell in row:  # проходимся по всем ячейкам в текущей строке
            dest_cell = dest_ws.cell(row=cell.row, column=cell.column)  # выбираем соответствующую ячейку в листе-назначении
            dest_cell.value = cell.value  # копируем значение ячейки
            dest_cell.data_type = cell.data_type  # копируем тип данных ячейки
            if cell.has_style:  # если ячейка имела стиль, копируем его в ячейку-назначение
                dest_cell.font = copy(cell.font)
                dest_cell.border = copy(cell.border)
                dest_cell.fill = copy(cell.fill)
                dest_cell.number_format = copy(cell.number_format)
                dest_cell.protection = copy(cell.protection)
                dest_cell.alignment = copy(cell.alignment)
    # сохраняем изменения в перезаписывая исходный файл
    dest_wb.save(sourсe)

def make_substation(cids: Path):
    print("make_substation")
    my_file_xml = cids.parent/f"{cids.stem}.xml" # прописываю путь и имя копии
    substation = {}
    report_cipa = Report_cipa()
    with ZipFile(cids, mode='r') as zip_file:
        info = zip_file.namelist() 
        zip_file.extractall(cids.parent)
        for file in info:
            manufacturer = whose_cid(cids.parent/file)
            if manufacturer == "IEC61850_TERMINAL":
                tree = ET.parse(cids.parent/file)
                print("\n Читаем ",file)
                root = tree.getroot()
                ied_name = root.find(IED).attrib["name"] # получаю IED файла с которым работаю
                # на этом месте нужно описась входящие и исходящие сигналы
                substation[ied_name] = tree
            else:
                print("\n !!!!!!!!Пропускаем ", file)
    test1 = 0
    #new_substation = substation.copy()

    for ied_name, tree in substation.items():
        test2 = 0
        test1 += 1
        root = tree.getroot()
        # print(type(tree))
        # print(type(root))
        report_cipa.ok_term.add(ied_name) # отправили в отчет, что этот терминал обработан VV_T1
        for extref in root.findall(f".//{IED}[@name='{ied_name}']//{INPUTS}/{EXTREF}"): # бежим по всем входящим нашего терминала
            
            test2 += 1
            source_ied_name = extref.get("iedName") # выдергиваем ИЕД на который он подписан VV_T2
            if source_ied_name in substation: # если терминал с этим ИЕДом у нас есть, заходим в него и говорим что подписаны на него
                root_source = substation[source_ied_name].getroot() # заходим в корень терминала на котрый подписаны
                doName = extref.attrib["doName"] # "SPCSO40"
                print(test1, test2, doName)
                cipa_name_signal = root_source.find(f".//{IED}[@name='{source_ied_name}']//{LDEVICE}[@inst='{extref.get('ldInst')}']/{LN}[@prefix='{extref.get('prefix')}'][@lnClass='{extref.get('lnClass')}'][@inst='{extref.get('lnInst')}']/{DOI}[@name='{doName}']") # нахожу узел на который подписан [@doName='{doName}']
                print(test1, test2, cipa_name_signal.attrib)
                cipa = ET.Element(CIPA)
                cipa.attrib["IED"] = ied_name
                cipa.attrib["name"] = cipa_name_signal.get("name")
                desc_in_DOI = cipa_name_signal.get("desc")
                if desc_in_DOI:
                    cipa.attrib["desc"] = desc_in_DOI
                else:
                    try:
                        desc_in_DAI = cipa_name_signal.find(DAI).get("desc")
                        cipa.attrib["desc"] = desc_in_DAI
                    except:
                        cipa.attrib["desc"] = f"!!!!!!!!!!!!{source_ied_name}"
                    # if desc_in_DAI:
                    #     cipa.attrib["desc"] = desc_in_DAI
                    # else:
                    #     cipa.attrib["desc"] = f"!!!!!!!!!!!!{source_ied_name}"
                print(test1, test2, cipa.attrib)
                cipa_signal = root_source.find(f".//{IED}[@name='{source_ied_name}']//{FCDA}[@ldInst='{extref.get('ldInst')}'][@prefix='{extref.get('prefix')}'][@lnClass='{extref.get('lnClass')}'][@lnInst='{extref.get('lnInst')}'][@doName='{extref.get('doName')}'][@daName='{extref.get('daName')}']")
                cipa_signal.append(cipa)                
            else:
                print("НЕ нашли", source_ied_name)
                report_cipa.err.add(source_ied_name)
    
    for k, v in substation.items():
        v.write(cids.parent/f"{k}.cid", encoding="UTF-8")
        #zip_file.write(cids.parent/("c_" + k), (k + ".cid"))

    make_xl(substation.keys(), cids)

    with ZipFile(cids.parent/"cipa.zip", mode='w') as zip_file:
        # for k, v in substation.items():
        #     v.write(cids.parent/("c_" + k), encoding="UTF-8")
        #     zip_file.write(cids.parent/("c_" + k), (k + ".cid"))
        report_cipa.make_report(cids.parent)
        zip_file.write(cids.parent/("report_cipa.txt"), ("report_cipa.txt"))
        zip_file.write(cids.parent/"For_Sancho.xlsx", "For_Sancho.xlsx")


class SuperSocket():
    def __init__(self, sock):
        self._sock = sock
    def send_msg(self, msg):
        # Каждое сообщение будет иметь префикс в 4 байта блинной(network byte order)
        msg = struct.pack('>I', len(msg)) + msg
        self._sock.send(msg)
    def recv_msg(self):
        # Получение длины сообщения и распаковка в integer
        raw_msglen = self.recvall(4)
        if not raw_msglen:
            return None
        msglen = struct.unpack('>I', raw_msglen)[0]
        # Получение данных
        return self.recvall(msglen)
    def recvall(self, n):
        # Функция для получения n байт или возврата None если получен EOF
        data = b''
        while len(data) < n:
            packet = self._sock.recv(n - len(data))
            if not packet:
                return None
            data += packet
        return data

sock = socket.socket()
# IP = "213.178.155.215"
IP = ""
PORT = 51085
sock.bind((IP, PORT))
sock.listen(5)

while True:
    print("start")
    # начинаем принимать соединения
    conn, addr = sock.accept()
    # выводим информацию о подключении
    print('connected:', addr)
    super_sock = SuperSocket(conn)
     # получаем название файла
    name_f = super_sock.recv_msg().decode('UTF-8')
    # создаем название каталога клиента
    path_f = Path.cwd()/"clients"/name_f.removesuffix(".zip")
    # создаем каталог клиента если его нет
    if not Path.is_dir(path_f):
        Path.mkdir(path_f)
    # открываем файл в режиме байтовой записи в папке клиента и записываем его из сокета
    (path_f/name_f).write_bytes(super_sock.recv_msg())
    #f.close()

    print("Обрабытываем сиды")
    # make_substation("clients/" + name_f.strip(".zip") + "/" + name_f)
    make_substation(path_f/name_f)
    
    # name_f = "clients/" + name_f.strip(".zip") + "/" + name_f.strip(".zip") + ".txt"
    #name_f_client = path + name + ".xlsx"
    print("Отправляем", "test.xlsx")
    #super_sock.send_msg((f"{name_f.removesuffix('.zip')}.xml").encode('UTF-8'))
    #f = open (path_f/f"{name_f.removesuffix('.zip')}.xml", "rb")
    super_sock.send_msg(("cipa.zip").encode('UTF-8'))
    f = open (path_f/"cipa.zip", "rb")
    super_sock.send_msg(f.read())
    f.close()

    conn.close()
    print("end")
    print("убери break после отладки")
    break

