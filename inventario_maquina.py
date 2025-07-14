import platform
import socket
import psutil
import os
import getpass
import subprocess
import requests
from openpyxl import Workbook, load_workbook

# 🔁 Pasta onde os arquivos serão salvos
DATA_DIR = "data"
os.makedirs(DATA_DIR, exist_ok=True)  # garante que a pasta exista
FILENAME = os.path.join(DATA_DIR, "inventario_maquinas.xlsx")

def get_wmic_value(command):
    try:
        result = subprocess.check_output(command, shell=True)
        lines = result.decode(errors="ignore").strip().split('\n')
        return lines[1].strip() if len(lines) > 1 else "Não encontrado"
    except Exception:
        return "Erro"

def get_windows_name():
    try:
        result = subprocess.check_output(
            ["powershell", "-Command", "(Get-CimInstance Win32_OperatingSystem).Caption"],
            shell=True
        )
        return result.decode(errors="ignore").strip()
    except Exception:
        return "Erro"

def get_windows_license_status():
    try:
        result = subprocess.check_output(
            ['cscript', '//Nologo', 'C:\\Windows\\System32\\slmgr.vbs', '/xpr'],
            shell=True
        )
        output = result.decode(errors="ignore").lower()
        return "Sim" if "permanente" in output or "permanently" in output else "Não"
    except Exception:
        return "Erro"

def get_memory_type():
    try:
        cmd = [
            "powershell",
            "-Command",
            "Get-CimInstance Win32_PhysicalMemory | Select-Object -First 1 -ExpandProperty SMBIOSMemoryType"
        ]
        result = subprocess.check_output(cmd, shell=True)
        code = int(result.decode(errors="ignore").strip())
        memory_types = {
            20: "DDR",
            21: "DDR2",
            22: "DDR2 FB-DIMM",
            24: "DDR3",
            26: "DDR4",
            30: "DDR5"
        }
        return memory_types.get(code, f"Desconhecido (código {code})")
    except Exception:
        return "Erro"

def get_pc_type():
    try:
        result = subprocess.check_output('wmic computersystem get pcSystemType', shell=True)
        lines = result.decode(errors="ignore").strip().split('\n')
        if len(lines) < 2:
            return "Desconhecido"
        code = lines[1].strip()
        pc_types = {
            '1': 'Desktop',
            '2': 'Notebook'
        }
        return pc_types.get(code, 'Outro')
    except Exception:
        return "Erro"

def get_city_from_ip():
    try:
        response = requests.get("https://ipinfo.io/json")
        data = response.json()
        return data.get("city", "Cidade não encontrada")
    except Exception:
        return "Erro ao obter cidade"

def get_machine_info():
    info = {}
    info["Nome da máquina"] = socket.gethostname()
    info["Proprietário"] = getpass.getuser()
    info["Etiqueta"] = ""  # manual
    info["Cidade"] = get_city_from_ip()  # Detecta cidade automaticamente
    info["Departamento"] = ""  # manual
    info["Unidade Residente"] = ""  # manual
    info["Marca"] = get_wmic_value("wmic computersystem get manufacturer")
    info["Número de Série"] = get_wmic_value("wmic bios get serialnumber")
    info["Tipo"] = get_pc_type()
    info["Modelo"] = get_wmic_value("wmic computersystem get model")
    info["Licença"] = get_windows_name()
    info["Processador"] = get_wmic_value("wmic cpu get name")
    info["Troca de máquina"] = ""  # manual
    info["Tipo de memória"] = get_memory_type()
    info["Pentes"] = "1"  # estimado
    info["Tamanho"] = round(psutil.virtual_memory().total / (1024 ** 3), 2)
    disk = psutil.disk_usage('/')
    info["Armazenamento"] = round(disk.total / (1024 ** 3), 2)
    info["Tipo de armazenamento"] = "SSD ou HDD"
    info["Licença Windows"] = get_windows_license_status()
    info["Troca ou Upgrade"] = ""
    info["Prioridade"] = ""
    info["Antivírus"] = ""
    info["Upgrade?"] = ""
    info["Em uso?"] = "Sim"
    info["Está no AD?"] = os.environ.get('USERDOMAIN', "")
    info["Observações"] = ""
    return info

def save_to_excel(info, filename=FILENAME):
    try:
        wb = load_workbook(filename)
        ws = wb.active
    except FileNotFoundError:
        wb = Workbook()
        ws = wb.active
        ws.append(list(info.keys()))

    ws.append(list(info.values()))
    wb.save(filename)
    print(f"✅ Planilha '{filename}' salva com sucesso!")

def send_api(filepath):
    try:
        url = "http://192.168.0.138:5000//upload_excel"  # Substitua pelo IP ou domínio da sua API
        with open(filepath, "rb") as f:
            files = {'file': (os.path.basename(filepath), f, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')}
            response = requests.post(url, files=files)

        if response.status_code == 200:
            print("✅ Arquivo enviado para a API com sucesso.")
        else:
            print(f"⚠️ Erro ao enviar para a API: {response.status_code} - {response.text}")
    except Exception as e:
        print(f"⚠️ Falha ao conectar com a API: {e}")

if __name__ == "__main__":
    info = get_machine_info()
    save_to_excel(info)
    send_api(FILENAME)
