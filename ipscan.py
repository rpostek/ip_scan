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
import json
import openpyxl

class Property:
    def __init__(self, name: str, func: Callable, description: str='', enabled: bool=True) -> None:
        self.name = name
        self.func = func
        self.description = description
        self.enabled = enabled

class Func:
    @staticmethod
    def get_processor_data(ip: str) -> dict:
        data = dict()
        try:
            r = Func.runPSjson(f'gwmi win32_Processor -comp {ip}')
            data['Processor'] = r['Name']
            return data
        except:
            pass
        return data

    @staticmethod
    def get_network_data(ip: str) -> dict:
        data = dict()
        try:
            #r = Func.runPSjson(f"gwmi -ClassName Win32_NetworkAdapterConfiguration -Filter \"IPEnabled='True'\" -comp {ip} |  Where-Object {{$_.PNPDeviceID -like 'PCI*'}} | select *")
            #r = Func.runPSjson(f"gwmi -ClassName Win32_NetworkAdapter -comp {ip} |  Where-Object {{($_.PNPDeviceID -like 'PCI*') -and ($_.NetEnabled )}} | select *")
            r = Func.runPSjson(f"Get-NetNeighbor {ip}")

            data['MAC Address'] = r['LinkLayerAddress']
            return data
        except:
            pass
        return data

    '''@staticmethod
    def get_WoL(ip: str) -> dict:
        data = dict()
        try:
            r = Func.runPSjson(f"$wol = gwmi -Class Win32_NetworkAdapter -Filter \"netenabled = 'true'\" -ComputerName {ip} |  Where-Object {{$_.PNPDeviceID -like 'PCI*'}}; gwmi MSPower_DeviceWakeEnable -Namespace root\wmi -ComputerName {ip}| where " + "{$_.instancename -match [regex]::escape($wol.PNPDeviceID) } | select -Property Enable")
            data['WakeOnLan'] = r['Enable']
            return data
        except:
            pass
        return data
    '''

    @staticmethod
    def get_bios_data(ip: str) -> dict:
        data = dict()
        try:
            r = Func.runPSjson(f'gwmi Win32_BIOS -comp {ip} | select *')
            data['BIOS'] = r['SMBIOSBIOSVersion']
            data['BIOS date'] = r['ReleaseDate'][:4] + '-' + r['ReleaseDate'][4:6] + '-' + r['ReleaseDate'][6:8]
            return data
        except:
            pass
        return data

    @staticmethod
    def get_computer_data(ip: str) -> dict:
        data = dict()
        try:
            r = Func.runPSjson(f'gwmi win32_computersystem -comp {ip}')
            data['User Name'] = r['UserName']
            data['Name'] = r['Name']
            data['System Family'] = r['SystemFamily']
            data['Logical Processors'] = r['NumberOfLogicalProcessors']
            data['Memory'] = f"{float(r['TotalPhysicalMemory']) / (2 ** 30):2.2f} GB"
            data['Model'] = r['Model']
            data['Manufacturer'] = r['Manufacturer']
            return data
        except:
            pass
        return data

    @staticmethod
    def get_monitor_data(ip: str) -> dict:
        data = dict()
        try:
            psqry = f'$dt = gwmi -Namespace root\\wmi -Class wmiMonitorID -comp {ip};'\
                    + r" $name = '';foreach ($i in $dt.UserFriendlyName) {$name += [char[]]$i};"\
                    + r" $sn = '';foreach ($i in $dt.SerialNumberID) {$sn += [char[]]$i};"\
                    + r" @{'Monitor YearOfManufacture'=$dt.YearOfManufacture;'Monitor Name'=$name;'Monitor SN'=$sn}"
            data = Func.runPSjson(psqry)
            return data
        except:
            pass
        return data

    @staticmethod
    def get_os_version(ip: str) -> dict:
        OS_VERSIONS = {
            '10.0.19043': 'Windows 10 (21H1)',
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
            '6.3.9600': 'Windows 8.1 (Update 1)',
            '6.3.9200': 'Windows 8.1',
            '6.2.9200': 'Windows 8',
            '6.1.7601': 'Windows 7 SP1',
            '6.1.7600': 'Windows 7'
        }
        data = dict()
        try:
            r = Func.runPSjson(f'Get-WmiObject Win32_OperatingSystem -ComputerName {ip}')
            # for key in r.keys():
            #    print(key, ':', r[key])
            data['OS'] = r['Caption'] + ' ' + r['OSArchitecture']
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
            r = Func.runPSjson(f'Get-WmiObject win32_product -ComputerName {ip} |  ' + 'where{$_.Name -like "Microsoft Office Standard*" -or $_.Name -like "Microsoft Office Professional*"} | select Name,Version')
            # for key in r.keys():
            #    print(key, ':', r[key])
            data['Office'] = r['Name'] # + ' (' + r['Version'] +')'
            return data
        except:
            pass
        return data

    @staticmethod
    def runPSjson(cmd: str) -> str:
        try:
            ps = subprocess.Popen(["powershell", "-Command", cmd + ' | ConvertTo-json'],
                                  stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                  creationflags=subprocess.CREATE_NO_WINDOW)
            data = ps.communicate()[0]
            data = data.decode(encoding='cp852')
            return json.loads(data)
        except:
            return dict()

    @staticmethod
    def get_time_source(ip: str) -> dict:
        data = dict()
        try:
            #r = Func.runPS(f'w32tm /query /status /computer:{ip} | Select-String -Pattern "^Source:.*"')
            #data['Time Source'] = r['Line'].replace('Source: ', '')
            r = Func.runPSjson(f'w32tm /query /status /computer:{ip}')
            data_temp=dict()
            for parm in r:
                try:
                    pair = parm.split(':', 1)
                    data_temp[pair[0]] = pair[1].strip()
                except:
                    pass
            data['Time Source'] = data_temp['Source'] + ' / ' + data_temp['Last Successful Sync Time']
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
    Property('MAC Address', Func.get_network_data, 'adres MAC aktywnej karty sieciowej', False),
    #Property('WakeOnLan', Func.get_WoL, 'czy włączona jest funkcja Wake on LAN', False),
    Property('Last Logged User', Func.get_last_user, 'ostatni zalogowany użytkownik', True),
    Property('User Name', Func.get_computer_data, 'zalogowany użytkownik', False),
    Property('OS', Func.get_os_version, 'system operacyjny', False),
    Property('OS Version', Func.get_os_version, 'wersja (wydanie) systemu operacyjnego', False),
    Property('Manufacturer', Func.get_computer_data, 'producent komputera', False),
    Property('System Family', Func.get_computer_data, 'model komputera', False),
    Property('Model', Func.get_computer_data, 'model komputera', False),
    Property('BIOS', Func.get_bios_data, 'wersja BIOSu', False),
    Property('BIOS date', Func.get_bios_data, 'data wydania BIOSu', False),
    Property('Processor', Func.get_processor_data, 'typ procesora', False),
    Property('Logical Processors', Func.get_computer_data, 'liczba procesorów logicznych', False),
    Property('Memory', Func.get_computer_data, 'całkowita pamięć fizyczna komputera', False),
    Property('Time Source', Func.get_time_source, 'nazwa serwera synchronizacji czasu i czas ostatniej synchronizacji', False),
    Property('Chrome Version', Func.get_chrome_version, 'wersja przeglądarki chrome', False),
    Property('PrintConfig.dll', Func.exists_printconfig, 'obecnośc pliku PrintConfig.dll potrzebnego do drukowania z aplikacji UPW', False),
    Property('Office', Func.get_office_version, 'wersja programu Office', False),
    Property('Monitor YearOfManufacture', Func.get_monitor_data, 'rok produkcji monitora', False),
    Property('Monitor Name', Func.get_monitor_data, 'model monitora', False),
    Property('Monitor SN', Func.get_monitor_data, 'numer seryjny monitora', False),
)

def get_params():
    global window
    my_ip = get_my_ip()
    sg.theme('LightBrown1')
    #layout = [[sg.Text("Parametry do wyświetlenia:")]]
    layout=[]
    frame_layout = []
    for pr in properties:
        frame_layout.append([sg.Checkbox(text=pr.name, tooltip=pr.description, default=pr.enabled, key='property-' + pr.name)])
    layout.append([sg.Frame('Parameter list', frame_layout, title_color='black')])
    layout.append([sg.Text(f'IP addr range ', pad=(1,3)),
                   sg.InputText(size=(3, 1), default_text=my_ip[0], key='ips1', pad=(1, 3)),
                   sg.InputText(size=(3, 1), default_text=my_ip[1], key='ips2', pad=(1, 3)),
                   sg.InputText(size=(3, 1), default_text=my_ip[2], key='ips3', pad=(1, 3)),
                   sg.InputText(size=(3, 1), default_text='1', key='ips4', pad=(1, 3)),
                   sg.Text(f' - ', pad=(1,3)),
                   sg.InputText(size=(3,1), default_text=my_ip[0], key='ipe1', pad=(1,3)),
                   sg.InputText(size=(3,1), default_text=my_ip[1], key='ipe2', pad=(1,3)),
                   sg.InputText(size=(3,1), default_text=my_ip[2], key='ipe3', pad=(1,3)),
                   sg.InputText(size=(3,1), default_text='170', key='ipe4', pad=(1,3))])
    layout.append([sg.Checkbox(text='export to Excel', default=False, key='save-xlsx')],)
    layout.append([sg.ProgressBar(1, orientation='h', size=(20, 20), key='progress'), sg.Submit("OK", pad=((20, 10), 3))],)
    #layout.append([sg.ProgressBar(1, orientation='h', size=(20, 20), key='progress')],)
    window = sg.Window("ipscanner 1.3", layout)
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
            'ips1': values['ips1'], 'ips2': values['ips2'], 'ips3': values['ips3'], 'ips4': values['ips4'],
            'ipe1': values['ipe1'], 'ipe2': values['ipe2'], 'ipe3': values['ipe3'], 'ipe4': values['ipe4'],
            'save-xlsx': values['save-xlsx'],}

def get_my_ip():
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

def check_computer(ip: ipaddress, no: int, cfg: dict, table: list) -> None:
    if Func.detect_on(ip.__str__()):
        data = {'No': f'{no:03}', 'IP': ip.__str__()}
        for f in cfg['function_list']:
            data.update(f(ip.__str__()))
        lock.acquire()
        table.append(clear_data(data, cfg['property_list']))
        lock.release()

def save_xls(table, header):
    filename = sg.popup_get_file('Save as',
    title = None,
    default_path = "",
    default_extension = ".xlsx",
    save_as = True,
    multiple_files = False,
    file_types = (('Excel files', '*.xlsx'), ('ALL Files', '*.*')),
    no_window = True,
    no_titlebar = False,
    grab_anywhere = False,
    keep_on_top = True,
    location = (None, None),
    initial_folder = None,
    image = None,
    files_delimiter = ";",
    modal = True,
    history = False,
    history_setting_filename = None)
    if filename:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(header)
        for data in table:
            row=[]
            for field in header:
                row.append(data[field])
            ws.append(row)
        wb.save(filename)

if __name__ == '__main__':
    cfg = get_params()
    try:
        ip_start = ipaddress.ip_address(cfg['ips1'] + '.' + cfg['ips2'] + '.' + cfg['ips3'] + '.' + cfg['ips4'])
        ip_end = ipaddress.ip_address(cfg['ipe1'] + '.'  + cfg['ipe2'] + '.'  + cfg['ipe3'] + '.' +cfg['ipe4'])
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
    if cfg['save-xlsx']:
        save_xls(table, cfg['property_list'])
    display_table(table, cfg['property_list'])
