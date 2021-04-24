from typing import Callable
import subprocess
import csv
import os
import pathlib
import PySimpleGUI as sg
import socket
import ipaddress
import sys
import tempfile
from pefile import PE
import threading
import time


class Property:
    def __init__(self, name: str, func: Callable, description: str='', enabled: bool=True) -> None:
        self.name = name
        self.func = func
        self.description = description
        self.enabled = enabled

class Func:
    @staticmethod
    def get_computer_data(ip: str) -> dict:
        data = dict()
        try:
            r = Func.runPS(f'gwmi win32_computersystem -comp {ip}')
            # for key in r.keys():
            #    print(key, ':', r[key])
            data['User Name'] = r['UserName']
            data['Name'] = r['Name']
            data['System Family'] = r['SystemFamily']
            data['Logical Processors'] = r['NumberOfLogicalProcessors']
            data['Memory'] = f"{float(r['TotalPhysicalMemory']) / (2 ** 30):2.2f} GB"
            return data
        except:
            pass
        return data

    @staticmethod
    def get_os_version(ip: str) -> dict:
        OS_VERSIONS = {
            '10.0.19042': 'Windows 10 (20H2)',
            '10.0.19041': 'Windows 10 (2004)',
            '10.0.18363': 'Windows 10 (1909)',
            '10.0.18362': 'Windows 10 (1903)',
            '10.0.17763': 'Windows 10 (1809)',
            '10.0.17134': 'Windows 10 (1803)',
            '10.0.16299': 'Windows 10 (1709)',
            '10.0.15063': 'Windows 10 (1703)',
            '10.0.14393': 'Windows 10 (1607)',
            '10.0.10586': 'Windows 10 (1511)',
            '10.0.10240': 'Windows 10',
            '6.3.96008': 'Windows 8.1 (Update 1)',
            '6.3.9200': 'Windows 8.1',
            '6.2.9200': 'Windows 8',
            '6.1.7601': 'Windows 7 SP1',
            '6.1.7600': 'Windows 7'
        }
        data = dict()
        try:
            r = Func.runPS(f'Get-WmiObject Win32_OperatingSystem -ComputerName {ip}')
            # for key in r.keys():
            #    print(key, ':', r[key])
            data['OS'] = r['Caption']
            data['OS Version'] = OS_VERSIONS.get(r['Version'], 'build '+ r['Version'])
            # data['SerialNumber'] = r['SerialNumber']
            return data
        except:
            pass
        return data

    @staticmethod
    def get_office_version(ip: str) -> dict:
        data=dict()
        try:
            r = Func.runPS(f'Get-WmiObject win32_product -ComputerName {ip} |  ' + 'where{$_.Name -like "Microsoft Office Standard*" -or $_.Name -like "Microsoft Office Professional*"} | select Name,Version')
            # for key in r.keys():
            #    print(key, ':', r[key])
            data['Office'] = r['Name'] # + ' (' + r['Version'] +')'
            return data
        except:
            pass
        return data


    @staticmethod
    def runPS(cmd: str) -> dict:
        '''
        ;param cmd: komenda Powershell do wykonania
        :return: dict zawierająca dane lub pustą dict gdy wystąpił błąd wykonania
        komenda powinna zwracać obiekt, który da się zamienić na dict
        jesli komenda zwraca listę obiektów, to zwracany jest tylko pierwszy obiekt
        '''
        try:
            ps = subprocess.Popen(["powershell", "-Command", cmd + ' | ConvertTo-Csv'],
                                  stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                  creationflags=subprocess.CREATE_NO_WINDOW)
            data = ps.communicate()[0]
            data = data.decode(encoding='cp852')
            d = data.split(sep='\n')
            d.pop(0)
            csv_reader = csv.DictReader(d, delimiter=',')
            return list(csv_reader)[0]

        except:
            return dict()

    @staticmethod
    def get_time_source(ip: str) -> dict:
        data = dict()
        try:
            r = Func.runPS(f'w32tm /query /status /computer:{ip} | Select-String -Pattern "^Source:.*"')
            data['Time Source'] = r['Line'].replace('Source: ', '')
            return data
        except:
            pass
        return data

    @staticmethod
    def get_last_user(ip: str) -> dict:
        try:
            last = (None, 0)
            for user in os.scandir(r'\\' + ip + r'\c$\Users'):
                if user.name == 'All Users':
                    continue
                moje_konto = pathlib.Path('\\\\' + ip + '\\c$\\Users\\' + user.name + '\\Desktop\\Moje konto.url')
                if moje_konto.exists():
                    if moje_konto.stat().st_mtime > last[1]:
                        last = user.name, moje_konto.stat().st_mtime
                    #logs.append((user.name, datetime.datetime.fromtimestamp(moje_konto.stat().st_mtime)))
                    #print(user, datetime.datetime.fromtimestamp(moje_konto.stat().st_mtime).isoformat())
            return {'Last Logged User': last[0]}
        except PermissionError:
            return {'Last Logged User': ''}

    @staticmethod
    def detect_on(ip: str) -> bool:
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.settimeout(0.2)
        try:
            s.connect((ip, 445))
            return True
        except socket.error as e:
            return False
        s.close()

    @staticmethod
    def get_version(path):
        '''podaje wersję pliku podanego w path'''
        pe = PE(path)
        if not 'VS_FIXEDFILEINFO' in pe.__dict__:
            return
        if not pe.VS_FIXEDFILEINFO:
            return
        verinfo = pe.VS_FIXEDFILEINFO[0]
        wersja_pliku = (verinfo.FileVersionMS >> 16, verinfo.FileVersionMS & 0xFFFF,
                             verinfo.FileVersionLS >> 16, verinfo.FileVersionLS & 0xFFFF)
        wersja_produktu = (verinfo.ProductVersionMS >> 16, verinfo.ProductVersionMS & 0xFFFF,
                                verinfo.ProductVersionLS >> 16, verinfo.ProductVersionLS & 0xFFFF)
        if wersja_produktu == wersja_pliku:
            wersja = ".".join(str(i) for i in wersja_pliku)
        else:
            wersja = str(wersja_produktu[0]) + '.' + ".".join(str(i) for i in wersja_pliku)
        return wersja

    @staticmethod
    def get_chrome_version(ip):
        wersja = ''
        try:
            if not isinstance(ip, str):
                ip=ip.__str__()
            filename = f'\\\\{ip}\\c$\\Program Files\\Google\\Chrome\\Application\\chrome.exe'
            if os.path.isfile(filename):
                wersja = Func.get_version(filename)
            filename = f'\\\\{ip}\\c$\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe'
            if os.path.isfile(filename):
                wersja = Func.get_version(filename)
        except:
            pass
        return {"Chrome Version": wersja}

    @staticmethod
    def exists_printconfig(ip):
        if not isinstance(ip, str):
            ip=ip.__str__()
        if os.path.isfile(r'\\' + ip + r'\c$\Windows\System32\spool\drivers\x64\3\Printconfig.dll' ):
            return {'PrintConfig.dll': 'OK'}
        else:
            return {'PrintConfig.dll': '-'}

properties = (
    Property('Name', Func.get_computer_data, 'nazwa komputera', False),
    Property('Last Logged User', Func.get_last_user, 'ostatni zalogowany użytkownik', True),
    Property('User Name', Func.get_computer_data, 'zalogowany użytkownik', False),
    Property('OS', Func.get_os_version, 'system operacyjny', False),
    Property('OS Version', Func.get_os_version, 'wersja (wydanie) systemu operacyjnego', False),
    Property('System Family', Func.get_computer_data, 'model komputera', False),
    Property('Logical Processors', Func.get_computer_data, 'liczba procesorów logicznych', False),
    Property('Memory', Func.get_computer_data, 'całkowita pamięć fizyczna komputera', False),
    Property('Time Source', Func.get_time_source, 'nazwa serwera synchronizacji czasu', False),
    Property('Chrome Version', Func.get_chrome_version, 'wersja przeglądarki chrome', False),
    Property('PrintConfig.dll', Func.exists_printconfig, 'obecnośc pliku PrintConfig.dll potrzebnego do drukowania z aplikacji UPW', False),
    Property('Office', Func.get_office_version, 'wersja programu Office', False),
)

def get_params():
    global window
    #sg.preview_all_look_and_feel_themes()
    my_ip = get_my_ip()
    sg.theme('LightBrown1')
    layout = [[sg.Text("Parametry do wyświetlenia:")]]
    for pr in properties:
        layout.append([sg.Checkbox(text=pr.name, tooltip=pr.description, default=pr.enabled, key='property-' + pr.name)])
    layout.append([sg.Text('IP addr range 10.128.', pad=(1,3)), sg.InputText(size=(3,1), default_text=my_ip[2], key='ips3', pad=(1,3)), sg.InputText(size=(3,1), default_text='1', key='ips4', pad=(1,3)),\
        sg.Text(' - 10.128.', pad=(1,3)), sg.InputText(size=(3,1), default_text=my_ip[2], key='ipe3', pad=(1,3)), sg.InputText(size=(3,1), default_text='170', key='ipe4', pad=(1,3))])
    layout.append([sg.ProgressBar(1, orientation='h', size=(20, 20), key='progress'), sg.Submit("OK", pad=((20, 10), 3))],)
    #layout.append([sg.ProgressBar(1, orientation='h', size=(20, 20), key='progress')],)
    window = sg.Window("ipscanner 1.0", layout)
    while True:
        event, values = window.read()
        if event == "OK":
            # TODO disable controls of the form
            window.Element('OK').update(disabled=True)
            break
        if event == sg.WIN_CLOSED:
            sys.exit(0)
            # TODO close running threads
    #window.close()
    prop_list = [k.replace('property-', '') for k,v in values.items() if v and k.startswith('property-')]
    fun_list = {p.func for p in properties if p.name in prop_list}  #lista funkcji do wykonania
    prop_list = ['No', 'IP'] + prop_list   #lista kolumn do wyświetlenia
    return {'property_list': prop_list, 'function_list': fun_list,
            'ips3': values['ips3'], 'ips4': values['ips4'], 'ipe3': values['ipe3'],'ipe4': values['ipe4'],}

def get_my_ip() -> list:
    return socket.gethostbyname(socket.gethostname()).split('.')

def display_table(table, header):
    fn = tempfile.TemporaryFile().name
    with open(fn, 'w', newline='\n') as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=header)
        writer.writeheader()
        for data in table:
            writer.writerow(data)
    ps = subprocess.Popen(["powershell", "-Command", f"Import-Csv -path {fn} | Out-GridView -Wait"],
                          creationflags=subprocess.CREATE_NO_WINDOW)

def clear_data(data: dict, header: list) -> dict:
    d = dict()
    for key in header:
        try:
            d[key] = data[key]
        except KeyError:
            d[key] = 'n/a'
    return d

def check_computer(ip, no, cfg, table):
    if Func.detect_on(ip.__str__()):
        data = {'No': f'{no:03}', 'IP': ip.__str__()}
        for f in cfg['function_list']:
            data.update(f(ip.__str__()))
        lock.acquire()
        table.append(clear_data(data, cfg['property_list']))
        lock.release()
        #print(ip)

if __name__ == '__main__':
    cfg = get_params()
    try:
        ip_start = ipaddress.ip_address('10.128.' + cfg['ips3'] + '.' + cfg['ips4'])
        ip_end = ipaddress.ip_address('10.128.' + cfg['ipe3'] + '.' +cfg['ipe4'])
        if ip_start > ip_end:
            raise ValueError
    except ValueError:
        sg.popup_ok('Błędny adres IP.', auto_close=True, auto_close_duration=3)
        sys.exit(1)
    table = []
    lock = threading.Lock()
    ip = ip_start
    no = 1
    threads = []
    while ip <= ip_end:
        t = threading.Thread(target=check_computer, args=(ip, no, cfg, table,))
        threads.append(t)
        t.start()
        no = no + 1
        ip = ip + 1
    while threading.active_count() > 1:
        time.sleep(0.2)
        window.FindElement('progress').UpdateBar(no-threading.active_count(), no-1)
    window.FindElement('progress').UpdateBar(no - threading.active_count(), no - 1)
    time.sleep(0.2)
    for t in threads:
        t.join()
    window.close()
    table = sorted(table, key=lambda i:i['No'])
    for r in range(len(table)):
        table[r]['No'] = f'{r+1:03}'
    display_table(table, cfg['property_list'])
