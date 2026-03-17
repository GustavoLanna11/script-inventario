import platform
import socket
import psutil
import os
import getpass
import subprocess
import requests

API_URL = "https://api-inventario-wudx.onrender.com/upload_excel"

def run_command_with_fallback(ps_command=None, wmic_command=None, fallback=None):
    if ps_command:
        try:
            result = subprocess.check_output(
                ["powershell", "-Command", ps_command],
                shell=True
            )
            value = result.decode(errors="ignore").strip()
            if value:
                return value
        except:
            pass

    if wmic_command:
        try:
            result = subprocess.check_output(wmic_command, shell=True)
            lines = result.decode(errors="ignore").strip().split('\n')
            if len(lines) > 1:
                return lines[1].strip()
        except:
            pass

    if fallback:
        try:
            return fallback()
        except:
            pass

    return "Não encontrado"


def get_windows_name():
    return run_command_with_fallback(
        ps_command="(Get-CimInstance Win32_OperatingSystem).Caption",
        fallback=lambda: platform.system()
    )


def get_windows_license_status():
    try:
        result = subprocess.check_output(
            ['cscript', '//Nologo', 'C:\\Windows\\System32\\slmgr.vbs', '/xpr'],
            shell=True
        )
        output = result.decode(errors="ignore").lower()
        return "Sim" if "permanente" in output or "permanently" in output else "Não"
    except:
        return "Erro"


def get_memory_type():
    try:
        result = subprocess.check_output(
            ["powershell", "-Command",
             "Get-CimInstance Win32_PhysicalMemory | Select-Object -First 1 -ExpandProperty SMBIOSMemoryType"],
            shell=True
        )
        code = int(result.decode(errors="ignore").strip())
        memory_types = {
            20: "DDR", 21: "DDR2", 24: "DDR3", 26: "DDR4", 30: "DDR5"
        }
        return memory_types.get(code, f"Desconhecido ({code})")
    except:
        return "Erro"

def get_pc_type():
    result = run_command_with_fallback(
        ps_command="(Get-CimInstance Win32_ComputerSystem).PCSystemType",
        wmic_command="wmic computersystem get pcSystemType",
        fallback=lambda: "Desconhecido"
    )

    clean = "".join(filter(str.isdigit, str(result)))

    pc_types = {
        "1": "Desktop",
        "2": "Notebook"
    }

    return pc_types.get(clean, "Outro")


def get_disk_type():
    try:
        result = subprocess.check_output(
            ["powershell", "-Command",
             "Get-PhysicalDisk | Select-Object -First 1 -ExpandProperty MediaType"],
            shell=True
        )
        media = result.decode(errors="ignore").strip().lower()
        return media.upper() if media in ["ssd", "hdd"] else "Desconhecido"
    except:
        return "Desconhecido"


def get_city_from_ip():
    try:
        response = requests.get("https://ipinfo.io/json", timeout=5)
        return response.json().get("city", "") if response.status_code == 200 else ""
    except:
        return ""


def has_kaspersky():
    try:
        result = subprocess.check_output(
            ["powershell", "-Command",
             "Get-CimInstance -Namespace root/SecurityCenter2 -ClassName AntivirusProduct | Select-Object -ExpandProperty displayName"],
            shell=True
        )
        return "Sim" if "kaspersky" in result.decode(errors="ignore").lower() else "Não"
    except:
        return "Erro"


def get_machine_info():
    info = {}

    processor_name = run_command_with_fallback(
        ps_command="(Get-CimInstance Win32_Processor).Name",
        wmic_command="wmic cpu get name",
        fallback=lambda: platform.processor()
    )

    weak_cpus = [
        "intel core i3-2120", "intel core i3-3220", "intel core i3-4130",
        "intel core i5-2430m", "amd a4-6300", "intel core i3-4005u",
        "intel core i3-5015u", "intel celeron j1800", "intel celeron g460",
        "intel core i3-3217u"
    ]

    normalized_processor = processor_name.lower().split("@")[0].split("cpu")[0].strip()

    info["Nome da máquina"] = socket.gethostname()
    info["Proprietário"] = getpass.getuser()
    info["Etiqueta"] = ""
    info["Cidade"] = get_city_from_ip()
    info["Departamento"] = ""
    info["Unidade Residente"] = ""

    info["Marca"] = run_command_with_fallback(
        ps_command="(Get-CimInstance Win32_ComputerSystem).Manufacturer",
        wmic_command="wmic computersystem get manufacturer"
    )

    info["Número de Série"] = run_command_with_fallback(
        ps_command="(Get-CimInstance Win32_BIOS).SerialNumber",
        wmic_command="wmic bios get serialnumber"
    )

    info["Tipo"] = get_pc_type()

    info["Modelo"] = run_command_with_fallback(
        ps_command="(Get-CimInstance Win32_ComputerSystem).Model",
        wmic_command="wmic computersystem get model"
    )

    info["Licença"] = get_windows_name()
    info["Processador"] = processor_name
    info["Tipo de memória"] = get_memory_type()
    info["Pentes"] = "1"

    ram_gb = round(psutil.virtual_memory().total / (1024 ** 3), 2)
    info["Tamanho"] = ram_gb

    disk = psutil.disk_usage('/')
    info["Armazenamento"] = round(disk.total / (1024 ** 3), 2)
    info["Tipo de armazenamento"] = get_disk_type()

    info["Antivírus"] = has_kaspersky()
    info["Em uso?"] = "Sim"
    info["Está no AD?"] = os.environ.get('USERDOMAIN', "")
    info["Observações"] = ""

    is_weak_cpu = any(cpu in normalized_processor for cpu in weak_cpus)

    if is_weak_cpu:
        info["Troca de máquina"] = "Sim"
        info["Upgrade?"] = "Não"
        info["Troca ou Upgrade"] = "Troca"
        info["Prioridade"] = "Alta"
        info["Licença Windows"] = "Máquina para troca"
    else:
        info["Troca de máquina"] = "Não"

        if ram_gb < 4 or info["Tipo de armazenamento"].lower() == "hdd":
            info["Upgrade?"] = "Sim"
            info["Troca ou Upgrade"] = "Upgrade"
            info["Prioridade"] = "Não será trocada"
            info["Licença Windows"] = get_windows_license_status()
        else:
            info["Upgrade?"] = "Não"
            info["Licença Windows"] = get_windows_license_status()
            info["Troca ou Upgrade"] = "Nenhum"
            info["Prioridade"] = "Não será trocada"

    return info


def normalize_data(info):
    string_fields = [
        "Departamento",
        "Cidade",
        "Etiqueta",
        "Observações",
        "Unidade Residente"
    ]

    for field in string_fields:
        value = info.get(field)
        info[field] = None if value in ["", None] else str(value)

    try:
        info["Pentes"] = int(info.get("Pentes", 0))
    except:
        info["Pentes"] = 0

    return info


def send_api(info):
    try:
        response = requests.post(API_URL, json=info, timeout=10)

        if response.status_code == 200:
            print("🌐✅ Inventário enviado com sucesso!")
        else:
            print(f"⚠️ Erro API: {response.status_code} - {response.text}")

    except Exception as e:
        print(f"⚠️ Falha conexão API: {e}")


if __name__ == "__main__":
    info = get_machine_info()
    info = normalize_data(info)
    send_api(info)