import sqlite3
import PySimpleGUI as sg
from wakeonlan import send_magic_packet

def get_computers():
    con = sqlite3.connect('computers.sqlite')
    cur = con.cursor()
    data = cur.execute('''
        WITH cte as (
            SELECT *, ROW_NUMBER() OVER (
                PARTITION BY name ORDER BY time desc) as rn
            from computers as c
            )
            select name, user, ip, mac, time from cte where rn=1
            order by name
    ''')
    data = list(data)
    con.close()
    return data

if __name__ == '__main__':
    computers = get_computers()
    column = []
    for c in computers:
        column.append([sg.Text(c[0],size=(14,1)), sg.Text(c[1],size=(10,1)), sg.Text(c[2],size=(11,1)), sg.Text(c[3],size=(14,1)), sg.Button('Wake', key="BUTTON-" + c[3])])
    layout = [[sg.Column(column, scrollable=True, size=(530, 800))]]
    window = sg.Window("WakeOnLAN", layout)
    while True:
        event, values = window.read()
        if event == sg.WIN_CLOSED:
            break
        if event[:7] == "BUTTON-":
            mac = event[7:]
            send_magic_packet(mac)
