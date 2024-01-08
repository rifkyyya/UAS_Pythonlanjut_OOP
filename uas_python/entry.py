import PySimpleGUI as sg
import pandas as pd
import openpyxl as df

sg.theme('DarkBlue12')

EXCEL_FILE = 'laundry.xlsx'

# Periksa apakah file excel sudah ada atau belum
try:
    df = pd.read_excel(EXCEL_FILE)
except FileNotFoundError:
    df = pd.DataFrame(columns=['Nama Pelanggan', 'No Telp', 'Alamat', 'Tgl Masuk', 'Jenis Layanan', 'Berat'])
    df.to_excel(EXCEL_FILE, index=False)

# Tentukan path untuk gambar logo Anda
LOGO_IMAGE_PATH = 'logo1.png'

layout = [
    [sg.Text('', size=(30, 1)), sg.Image(filename=LOGO_IMAGE_PATH, size=(150, 150), pad=((0, 0), (0, 0)))],
    [sg.Text('Masukkan Data Laundry:')],
    [sg.Text('Nama Pelanggan', size=(15, 1)), sg.InputText(key='Nama')],
    [sg.Text('No Telp', size=(15, 1)), sg.InputText(key='Tlp')],
    [sg.Text('Alamat', size=(15, 1)), sg.Multiline(key='Alamat')],
    [sg.Text('Tgl Masuk', size=(15, 1)), sg.InputText(key='Tgl Masuk'),
     sg.CalendarButton('Kalender', target='Tgl Masuk', format=('%d-%M-%Y'))],
    [sg.Text('Jenis Layanan', size=(15, 1)), sg.Combo(['Cuci Kering', 'Setrika', 'Dry Cleaning'], key='Jenis Layanan')],
    [sg.Text('Berat (kg)', size=(15, 1)), sg.InputText(key='Berat')],
    [sg.Submit(), sg.Button('clear'), sg.Exit()]
]

window = sg.Window('Form Pendaftaran Laundry', layout)

def clear_input():
    for key in values:
        window[key]('')
    return None

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'EXIT':
        break
    if event == 'clear':
        clear_input()
    if event == 'Submit':
        df = df._append(values, ignore_index=True)
        df.to_excel(EXCEL_FILE, index=False)
        sg.popup('Data Berhasil Disimpan')
        clear_input()

window.close()
