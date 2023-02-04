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


def make_terminal(file_name: str):
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

def make_substation(cids: Path):
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
                ied_name = root.find(IED).attrib["name"] # получаю IED вайла с которым работаю
                # на этом месте нужно описась входящие и исходящие сигналы
                substation[ied_name] = tree
            else:
                print("\n !!!!!!!!Пропускаем ", file)
    
    for ied_name, tree in substation.items():
        root = tree.getroot()
        #for extref in root.findall(f".//{IED}[@name='{ied_name}'//{INPUTS}/{EXTREF}"):
        report_cipa.ok_term.add(ied_name)
        for extref in root.findall(f".//{INPUTS}/{EXTREF}"):
            source_ied_name = extref.get("iedName")
            if source_ied_name in substation:
                print("нашли", source_ied_name)
                
                pass
            else:
                #print("НЕ нашли", source_ied_name)
                report_cipa.err.add(source_ied_name)
    
    print(len(report_cipa.err))
    for i in report_cipa.err:
        print(i)
    print()
    print(len(report_cipa.ok_term))
    for i in report_cipa.ok_term:
        print(i)
    with ZipFile(cids.parent/"cipa.zip", mode='w') as zip_file:
        for k, v in substation.items():
            v.write(cids.parent/("c_" + k))
            zip_file.write(cids.parent/("c_" + k), (k + ".cid"))
        report_cipa.make_report(cids.parent)
        zip_file.write(cids.parent/("report_cipa.txt"), ("report_cipa.txt"))


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

