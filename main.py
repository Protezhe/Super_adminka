import time
from PyQt6 import uic
from PyQt6.QtWidgets import QApplication
from wakeonlan import send_magic_packet
import telnetlib
import socket
import pandas as pd
import requests

Form, Window = uic.loadUiType("admin.ui")

app = QApplication([])
window = Window()
form = Form()
form.setupUi(window)
window.show()

# Проекторы


def projector(ip, command):

    timeout = 1

    try:
        session = telnetlib.Telnet(ip, 23, timeout)
    except socket.timeout:
        print("Нет соединения, иди с пультом " + ip)
    else:

        session.write(command.encode('ascii') + b"\r")

        session.read_until(b"OK", timeout=timeout)

        session.close()

# Скрипты на компьютерах


def sendsript(ip, message):

    host = ip
    port = 8748

    u_serv_sock = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    u_serv_sock.setsockopt(socket.SOL_SOCKET, socket.SO_BROADCAST, 1)
    u_serv_sock.settimeout(1.0)

    u_serv_sock.sendto(message, (host, port))

    u_serv_sock.close()

    print('Скрипт на ' + ip)



# Включение компов


def wake_on(ip, mac):
    send_magic_packet(mac, ip_address=ip)
    print('Включаем ' + ip)


# Реле Rodos


def reley(http, ip, text):
    print(text + ip)
    try:
        requests.get(http)

    except requests.ConnectionError:
        print('Нет подключения к реле ' + ip)


# Sharp

def tv_sharp(ip, command):

    timeout = 1

    try:
        session = telnetlib.Telnet(ip, 10008, timeout)
    except socket.timeout:
        print("Нет соединения, иди с пультом " + ip)


    else:

        session.read_until(b"Login:", timeout=timeout)

        session.write(b'\r')
        session.read_until(b"Password:", timeout=timeout)

        session.write(b'\r')
        session.read_until(b"OK", timeout=timeout)

        session.write(command.encode('ascii') + b"\r")
        session.read_until(b"OK", timeout=timeout)


        session.close()

def tv_iiyama_off(ip):

    timeout = 1

    try:
        session = telnetlib.Telnet(ip, 5000, timeout)
    except socket.timeout:
        print("Нет соединения, иди с пультом " + ip)


    else:

        session.write(b'\xA6\x01\x00\x00\x00\x04\x01\x18\x01\xBB')

        session.close()

# Включение iiyama


def tv_iiyama_on(ip, mac):
    send_magic_packet(mac, ip_address=ip)
    print('Включаем ' + ip)
    

def otchet_1():
    form.line.setText('Парк включен')


def otchet_0():
    form.line.setText('Парк выключен')


# Функции кнопок

# Весь парк


def park1():

    reley1()

    projectors_vkl()

    tv_sharp1()

    tv_iiyama1()

    pc_on()

    time.sleep(10)

    pc_script1()

    otchet_1()


def park0():

    projectors_vykl()

    pc_off()

    pc_script0()

    reley0()

    tv_sharp0()

    tv_iiyama0()

    otchet_0()


# Проекторы на мультики:


def mult_teatr_vkl():
    number = 0
    for number in range(nums_multteatr):
        projector(ip_mult_teatr[number], '~0000 1')

    if number + 1 == nums_multteatr:
        form.line.setText('Проекторы на мокап включены')


def mult_teatr_vykl():
    number = 0
    for number in range(nums_multteatr):
        projector(ip_mult_teatr[number], '~0000 0')

    if number + 1 == nums_multteatr:
        form.line.setText('Проекторы на мультики включены')

# Все проекторы:


def projectors_vkl():
    number = 0

    for number in range(nums_projector):
        projector(ip_projector[number], '~0000 1')

    if number + 1 == nums_projector:
        form.line.setText('Проекторы включены')


def projectors_vykl():

    number = 0

    for number in range(nums_projector):
        projector(ip_projector[number], '~0000 0')

    if number + 1 == nums_projector:
        form.line.setText('Проекторы выключены')


def pc_on():
    # number = 0

    for number in range(num_pc_on):

        wake_on('192.168.3.255', mac_pc[number])


def pc_off():

    # number = 0

    for number in range(num_pc_on):

        sendsript(ip_pc[number], b'script10')

# Скрипты:


def pc_script1():
    # number = 0

    for number in range(num_pc_script):

        sendsript(ip_script[number], b'script1')


def pc_script0():
    # number = 0

    for number in range(num_pc_script):
        sendsript(ip_script[number], b'script0')

# Реле


def reley1():
    for number in range(nums_ip_reley):
        reley('http://admin:admin@' + ip_reley[number] + '/protect/rb0f.cgi', ip_reley[number], 'Включаем реле ')


def reley0():
    for number in range(nums_ip_reley):
        reley('http://admin:admin@' + ip_reley[number] + '/protect/rb0n.cgi', ip_reley[number], 'Выключаем реле ')



# Sharp

def tv_sharp1():
    for number in range(nums_sharp_ip):
        tv_sharp(sharp_ip[number], 'POWR0001')

def tv_sharp0():
    for number in range(nums_sharp_ip):
        tv_sharp(sharp_ip[number], 'POWR0000')

# Iiyama
def tv_iiyama1():
    for number in range(nums_iiyama_ip):
        tv_iiyama_on('192.168.3.255', iiyama_mac[number])

def tv_iiyama0():
    for number in range(nums_iiyama_ip):
        tv_iiyama_off(iiyama_ip[number])

# ip адреса зон

ip_mult_teatr = ['192.168.2.141', '192.168.2.142', '192.168.2.140']
# Театр проекторы, не нужные для мультфильмов

# Читаем эксель
df = pd.read_excel('list.xlsx')

# Список проекторов и длина списка

ip_projector = df['ip_projector'].tolist()
ip_projector = [item for item in ip_projector if not pd.isnull(item)]
nums_projector = (len(ip_projector))

# Список скриптов и длина списка

ip_script = df['ip_script'].tolist()
ip_script = [item for item in ip_script if not pd.isnull(item)]
num_pc_script = (len(ip_script))

# Ip для включения

ip_pc = df['ip_pc'].tolist()
ip_pc = [item for item in ip_pc if not pd.isnull(item)]
num_pc_on = (len(ip_pc))


# Список мак адресов

mac_pc = df['mac_pc'].tolist()

# Список реле и длина списка

ip_reley = df['ip_reley'].tolist()
ip_reley = [item for item in ip_reley if not pd.isnull(item)]
nums_ip_reley = (len(ip_reley))

# Список sharp и длина списка

sharp_ip = df['sharp_ip'].tolist()
sharp_ip = [item for item in sharp_ip if not pd.isnull(item)]
nums_sharp_ip = (len(sharp_ip))

# Список iiyama и длина списка

iiyama_ip = df['iiyama_ip'].tolist()
iiyama_ip = [item for item in iiyama_ip if not pd.isnull(item)]
nums_iiyama_ip = (len(iiyama_ip))

# Список мак адресов

iiyama_mac = df['iiyama_mac'].tolist()


# Длина списков (Смотреть в экселе)
nums_multteatr = (len(ip_mult_teatr))




# Кнопки:

# Включение и выключение парка
form.park1.clicked.connect(park1)
form.park0.clicked.connect(park0)

# Проекторы на мультики:
form.teatrButton1.clicked.connect(mult_teatr_vkl)
form.teatrButton0.clicked.connect(mult_teatr_vykl)

# Все проекторы:
form.prButton1.clicked.connect(projectors_vkl)
form.prButton0.clicked.connect(projectors_vykl)

# Включение и выключение компов:
form.pc1.clicked.connect(pc_on)
form.pc0.clicked.connect(pc_off)

# Скрипты:
form.scriptButton1.clicked.connect(pc_script1)
form.scriptButton0.clicked.connect(pc_script0)

# Включение и выключение реле
form.rele1.clicked.connect(reley1)
form.rele0.clicked.connect(reley0)

# Включение и выключение tv sharp
form.tv_sharp1.clicked.connect(tv_sharp1)
form.tv_sharp0.clicked.connect(tv_sharp0)

# Включение и выключение tv iiyama
form.tv_iiyama1.clicked.connect(tv_iiyama1)
form.tv_iiyama0.clicked.connect(tv_iiyama0)


app.exec()
