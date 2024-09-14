import ctypes
import sys
import win32con
import win32gui
import time
import os
import pywinauto
import csv
from datetime import datetime

window_title_prefix = "CNC V4.03.64 / SIMULATION"
base_dir = r'C:\Users\Marco\Desktop\Fabrica Programação-alterado'
logging_dir = r'C:\CNC4.03\logging'
search_phrase = "ProcessRequests:cnccommand.cpp"
job_started_phrase = "Job started"
csv_file_path = os.path.join(base_dir, 'resultados.csv')

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def run_as_admin():
    if not is_admin():
        print("Reiniciando o script com privilégios de administrador...")
        ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, ' '.join(sys.argv), None, 1)
        sys.exit()

def send_key(key_code):
    ctypes.windll.user32.keybd_event(key_code, 0, 0, 0)
    time.sleep(0.05)
    ctypes.windll.user32.keybd_event(key_code, 0, win32con.KEYEVENTF_KEYUP, 0)

def focus_application(title_prefix):
    hwnd = find_window_by_title_prefix(title_prefix)
    if hwnd:
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.SetForegroundWindow(hwnd)
        time.sleep(2)
        if win32gui.GetForegroundWindow() == hwnd:
            print(f"Focado na janela: {win32gui.GetWindowText(hwnd)}")
            return hwnd
        else:
            print(f"Falha ao focar na janela: {win32gui.GetWindowText(hwnd)}")
            return None
    else:
        print(f"Janela com título que começa com '{title_prefix}' não encontrada!")
        return None

def find_window_by_title_prefix(prefix):
    def enum_windows_callback(hwnd, titles):
        title = win32gui.GetWindowText(hwnd)
        if title.startswith(prefix):
            titles.append(hwnd)
    
    titles = []
    win32gui.EnumWindows(enum_windows_callback, titles)
    return titles[0] if titles else None

def remove_parentheses_from_filename(file_path):
    dir_name, file_name = os.path.split(file_path)
    new_file_name = file_name.replace('(', '').replace(')', '')
    new_file_path = os.path.join(dir_name, new_file_name)
    
    if new_file_path != file_path:
        os.rename(file_path, new_file_path)
    
    return new_file_path

def ler_arquivos_em_pastas(base_dir):
    arquivos = []

    for root, dirs, files in os.walk(base_dir):
        for file in files:
            if file.endswith('.nc'):
                caminho_completo = os.path.join(root, file)
                novo_caminho = remove_parentheses_from_filename(caminho_completo)
                arquivos.append(novo_caminho)

    return arquivos

def encontrar_arquivo_especifico(diretorio, nome_arquivo):
    for root, dirs, files in os.walk(diretorio):
        if nome_arquivo in files:
            return os.path.join(root, nome_arquivo)
    return None

def ler_arquivo(caminho_arquivo):
    try:
        with open(caminho_arquivo, 'r') as file:
            conteudo = file.read()
            print("Conteúdo do arquivo:")
            print(conteudo)
            return conteudo
    except Exception as e:
        print(f"Erro ao ler o arquivo: {e}")
        return ""

def identificar_e_ler_arquivo(diretorio):
    arquivos = [os.path.join(diretorio, f) for f in os.listdir(diretorio) if os.path.isfile(os.path.join(diretorio, f))]
    
    if len(arquivos) != 1:
        print("O diretório deve conter exatamente um arquivo. Arquivos encontrados:")
        for arquivo in arquivos:
            print(arquivo)
        return None, None
    caminho_arquivo = arquivos[0]
    return ler_arquivo(caminho_arquivo), caminho_arquivo

def extrair_horario(linha):
    try:
        time_str = linha.split('->')[0]
        return datetime.strptime(time_str, '%d-%m %H:%M:%S')
    except Exception as e:
        print(f"Erro ao extrair horário: {e}")
        return None

def verificar_frase_no_arquivo(conteudo, search_phrase, job_started_phrase):
    linhas = conteudo.splitlines()
    job_start_time = None
    end_time = None

    for linha in linhas:
        if job_started_phrase in linha:
            job_start_time = extrair_horario(linha)
        elif search_phrase in linha:
            end_time = extrair_horario(linha)
            if job_start_time and end_time:
                break

    if job_start_time and end_time:
        diferenca_tempo = end_time - job_start_time
        print(f"Diferença de tempo: {diferenca_tempo}")
        return diferenca_tempo

    return None

run_as_admin()

# Cria o arquivo CSV e escreve o cabeçalho
with open(csv_file_path, 'w', newline='') as csvfile:
    fieldnames = ['Arquivo', 'Diferença de Tempo']
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
    writer.writeheader()

hwnd = focus_application(window_title_prefix)
if hwnd is None:
    sys.exit()

arquivos_encontrados = ler_arquivos_em_pastas(base_dir)

for arquivo in arquivos_encontrados:
    time.sleep(2)
    send_key(win32con.VK_F4)
    time.sleep(2)
    send_key(win32con.VK_F2)
    pywinauto.keyboard.send_keys(arquivo, with_spaces=True)
    time.sleep(2)
    pywinauto.keyboard.send_keys('{ENTER}')
    time.sleep(2)
    send_key(win32con.VK_F4)
    time.sleep(2)
    
    while True:
        conteudo, caminho_arquivo = identificar_e_ler_arquivo(logging_dir)
        if conteudo is not None:
            diferenca_tempo = verificar_frase_no_arquivo(conteudo, search_phrase, job_started_phrase)
            if diferenca_tempo is not None:
                print(f"Frase '{search_phrase}' encontrada.")
                with open(csv_file_path, 'a', newline='') as csvfile:
                    writer = csv.DictWriter(csvfile, fieldnames=['Arquivo', 'Diferença de Tempo'])
                    writer.writerow({'Arquivo': arquivo, 'Diferença de Tempo': diferenca_tempo})
                break
        print(f"Frase '{search_phrase}' não encontrada. Tentando novamente em 90 segundos...")
        time.sleep(90) 

    time.sleep(10)
    send_key(win32con.VK_F12)

    if caminho_arquivo:
        with open(caminho_arquivo, 'w') as file:
            file.write('')

time.sleep(100)
