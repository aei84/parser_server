#import xml.etree.ElementTree as ET
import lxml.etree as ET
from zipfile import ZipFile
import socket
import struct
import sys
import os

import openpyxl
from openpyxl.styles import PatternFill

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

COMMUNICATION = "{http://www.iec.ch/61850/2003/SCL}Communication"
SUBNETWORK = "{http://www.iec.ch/61850/2003/SCL}SubNetwork"
CONNECTEDAP = "{http://www.iec.ch/61850/2003/SCL}ConnectedAP"
ADDRESS = "{http://www.iec.ch/61850/2003/SCL}Address"
GSE = "{http://www.iec.ch/61850/2003/SCL}GSE"
P = "{http://www.iec.ch/61850/2003/SCL}P"

GSECONTROL = "{http://www.iec.ch/61850/2003/SCL}GSEControl"
FCDA = "{http://www.iec.ch/61850/2003/SCL}FCDA"

INPUTS = "{http://www.iec.ch/61850/2003/SCL}Inputs"
EXTREF = "{http://www.iec.ch/61850/2003/SCL}ExtRef"

class Goose():
    def __init__(self):
        self.goose_out_param ={} # {0: '1', 1: '010CCD010014', 2: '4', 3: '3', 4: '3114', 5: 'QC1G_F6_A2_PA', 6: '1', 7: '2.4', 8: '0'}
        self.goose_out_data = [0] # {1: '443', 2: '444', 3: '363', 4: '356', 5: '357', 6: '360', 7: '329', 8: '223', 9: '154'}
        # 0 - разрешение на передачу
        # 1 - MAC-Address
        # 2 - VLAN-PRIORITY
        # 3 - VLAN-ID
        # 4 - APPID
        # 5 - строковый иденцификатор goID (iedname)
        # 6 - confRef
        # 7 - интервал передачи при отсутствии изменений 
        # 8 - знак качества
        # 9 - ???        

class Terminal():
    def __init__(self):
        self.communication = {}
        self.goose_out = {}
        self.goose_in = []
        self.subscribers = {} # {1: [('QC1G_A1_DD', 5), .... 7: [], 8: [], 9: []}
        self.errors = []
        self.signal_names = {} # {214: 'Готовность LAN1', 215: 'Готовность LAN2'...}
        self.ied = {}


def whose_cid(cid: str):
    print("\n!!!", cid)
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
    print("Определяем чей cid ",cid)
    root = tree.getroot()
    print(root.nsmap)
    try:
        manufacturer = root.find(f"{IED}").attrib["manufacturer"] == "EKRA"
    except:
        print("Неизвестное устройство")
        return "Неизвестное устройство"
    if manufacturer:
        print("This is EKRA")
        return "EKRA"
    else:
        print("IEC61850_TERMINAL")
        return "IEC61850_TERMINAL"

def print_terminal(t: Terminal):
    print("IP_GOOSE", t.communication["IP_GOOSE"])
    for k, v in t.goose_out.items():
        print(k)
        for kk, vv in v.goose_out_param.items():
            print("\t", kk, vv)
        for d in v.goose_out_data:
            print("\t", d)
    print("подписки")
    for goose_in in t.goose_in:
        print("\t", goose_in)

def make_terminal_siemens(file_name: str):
    terminal = Terminal()
    tree = ET.parse(file_name)
    print("\n Читаем ",file_name)
    terminal.file_name = file_name
    root = tree.getroot()
    is_goose = False
    is_ekra = False
    #header_id = root.find(f"{HEADER}").attrib["id"]
    ied = root.find(f"{IED}")
    header_id = ied.attrib["name"]
    for k, v in ied.items():
        terminal.ied[k] = v
    # if terminal.ied["manufacturer"] != "SIEMENS":
    #     return None
    # получаем ip терминала
    for sub_net in root.findall(f"{COMMUNICATION}/{SUBNETWORK}"):
        ip = sub_net.find(f"{CONNECTEDAP}[@iedName='{header_id}']/{ADDRESS}/{P}[@type='IP']")
        if sub_net.find(f"{CONNECTEDAP}[@iedName='{header_id}']/{GSE}"):
            terminal.communication["IP_GOOSE"] = ip.text
            print(terminal.communication["IP_GOOSE"])
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
            ln = root.find(f"{IED}[@name='{header_id}']/{ACCESSPOINT}/{SERVER}/{LDEVICE}[@inst='{ldInst}']/{LN}[@prefix='{prefix}'][@lnClass='{lnClass}'][@inst='{inst}']").attrib["desc"]
            doi = root.find(f"{IED}[@name='{header_id}']/{ACCESSPOINT}/{SERVER}/{LDEVICE}[@inst='{ldInst}']/{LN}[@prefix='{prefix}'][@lnClass='{lnClass}'][@inst='{inst}']/{DOI}[@name='{doName}']").attrib["desc"]
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
        subs["iedName"] = inputs.attrib["iedName"]
        subs["ldInst"] = inputs.attrib["ldInst"]
        # if inputs.hasAttribute("prefix"):
        #     subs["prefix"] = inputs.attrib["prefix"]
        # else:
        #     subs["prefix"] = ""
        subs["prefix"] = inputs.get("lnClass")
        subs["lnClass"] = inputs.attrib["lnClass"]
        subs["lnInst"] = inputs.attrib["lnInst"]
        subs["doName"] = inputs.attrib["doName"]
        subs["daName"] = inputs.attrib["daName"]
        subs["intAddr"] = inputs.attrib["intAddr"]
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


def make_substation(cids: str):
    substation: Terminal = []
    with ZipFile(cids, mode='r') as zip_file:
        info = zip_file.namelist() 
        zip_file.extractall(cids.removesuffix(".zip"))
        for file in info:
            #print(file)
            #t = make_terminal(file) # сооздаю терминал
            manufacturer = whose_cid(cids.removesuffix(".zip") + "/" + file)
            if manufacturer == "EKRA":
                pass
                # ter = make_terminal_ekra(cids.removesuffix(".zip") + "/" + file)
                # if ter:
                #     substation.append(ter) # добовляю терминал в подстанцию
            #elif manufacturer == "SIEMENS-5":
            elif manufacturer == "IEC61850_TERMINAL":
                ter = make_terminal_siemens(cids.removesuffix(".zip") + "/" + file)
                if ter:
                    substation.append(ter) # добовляю терминал в подстанцию

    # print(len(substation)) # убеждаюсь что на подстанции 2 терминала
    #[print_terminal(t) for t in substation] # распечатываю обатерминала. Почему то они одинаковые, как второй файл а архиве, оба как t  на второй итерации цикла

    for ter_out in substation: # бегу по терминалам подстанции
        # print(ter_out.file_name, ter_out.goose_out_param)
        # print(ter_out.file_name, "Исходящий гусь", ter_out.goose_out_param[4])
        # print(ter_out.communication["IP"])
        appid_goose = ter_out.goose_out_param[4] # уникальный номер изходящего гуся
        quality = int(ter_out.goose_out_param[8]) # качество 0 - нет, 1 - до, 2 - после
        #quality = 0 # качество 0 - нет, 1 - до, 2 - после
        for ter_in in substation: # бегу по терминалам подстанции
            for goose_in_x in ter_in.goose_in.items(): # 
                if goose_in_x[1][3] == appid_goose: # 
                    if quality == 0:
                        index = int(goose_in_x[1][6])
                    elif quality == 1:
                        index = int(goose_in_x[1][6]) / 2
                    elif quality == 2:
                        index = (int(goose_in_x[1][6]) + 1) / 2
                    try:
                        ter_out.subscribers[index].append((ter_in.goose_out_param[5], int(goose_in_x[0]) + 1))
                    except:
                        ter_out.errors.append((index, ter_in.goose_out_param[5], int(goose_in_x[0]) + 1))

    substation.sort(key=lambda x: x.communication["IP"])                
    # [print_terminal(t) for t in substation]
    # [print_terminal_in_file(t, cids.rstrip(".zip") + ".txt") for t in substation] 
    # make_xl(substation, cids)
    # empty_vertical, empty_horizontal = fill_xl(cids.removesuffix(".zip") + ".xlsx", substation)
    # #paint_xl(cids.removesuffix(".zip") + ".xlsx")
    # paint_xl(cids.removesuffix(".zip") + ".xlsx", empty_vertical, empty_horizontal)


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
    print("Имя архива (str)", name_f)
    path = "clients/" + name_f.removesuffix(".zip") + "/"
    name = name_f.removesuffix(".zip")
    # создаем каталог для файлов клиента
    # print(type("clients/" + name_f.strip(".zip")))
    # print("clients/" + name_f.strip(".zip"))
    if not os.path.isdir("clients/" + name_f.removesuffix(".zip")):
        os.mkdir("clients/" + name_f.removesuffix(".zip"))
    # открываем файл в режиме байтовой записи в папке клиента
    f = open("clients/" + name_f.removesuffix(".zip") + "/" + name_f, 'wb')
    f.write(super_sock.recv_msg())
    f.close()

    print("Обрабытываем сиды")
    # make_substation("clients/" + name_f.strip(".zip") + "/" + name_f)
    make_substation(path + name_f)
    
    # name_f = "clients/" + name_f.strip(".zip") + "/" + name_f.strip(".zip") + ".txt"
    name_f = path + name + ".xlsx"
    print("Отправляем", name_f)
    super_sock.send_msg((name + ".xlsx").encode('UTF-8'))
    f = open (name_f, "rb")
    super_sock.send_msg(f.read())
    f.close()

    conn.close()
    print("end")
    print("убери break после отладки")
    break

