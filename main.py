import os
import platform
import subprocess
import winreg
import wmi
import psutil
import socket
import time
import sounddevice as sd
import soundfile as sf
import cv2
import numpy as np
import datetime
import json
from collections import defaultdict
from colorama import Fore, Back, Style, init
import sys
import io


# para el color
init()


if sys.stdout.encoding != 'UTF-8':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
if sys.stderr.encoding != 'UTF-8':
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')

def safe_wmi_query(wmi_class, attributes):
    """Realiza consultas WMI de manera segura"""
    c = wmi.WMI()
    try:
        items = []
        for item in c.query(f"SELECT * FROM {wmi_class}"):
            item_info = {}
            for attr in attributes:
                try:
                    item_info[attr] = getattr(item, attr, "N/A")
                except:
                    item_info[attr] = "Error"
            items.append(item_info)
        return items
    except Exception as e:
        return [{'Error': f"Consulta WMI fallida: {str(e)}"}]

def get_system_info():
    """Obtiene información general del sistema"""
    print("\n=== INFORMACIÓN DEL SISTEMA ===")
    
    info = {}
    
    # Información básica
    info['Sistema Operativo'] = f"{platform.system()} {platform.release()} {platform.version()}"
    info['Arquitectura'] = platform.machine()
    info['Nombre del Host'] = socket.gethostname()
    
    # Información del procesador
    try:
        cpu_info = safe_wmi_query("Win32_Processor", ["Name", "NumberOfCores", "NumberOfLogicalProcessors", "MaxClockSpeed"])[0]
        info['Procesador'] = {
            'Modelo': cpu_info.get('Name', 'N/A').strip(),
            'Núcleos Físicos': cpu_info.get('NumberOfCores', 'N/A'),
            'Núcleos Lógicos': cpu_info.get('NumberOfLogicalProcessors', 'N/A'),
            'Frecuencia Máxima': f"{cpu_info.get('MaxClockSpeed', 'N/A')} MHz" if cpu_info.get('MaxClockSpeed') != 'N/A' else 'N/A'
        }
    except Exception as e:
        info['Procesador'] = {'Error': f"No se pudo obtener información: {str(e)}"}
    
    # Información de memoria RAM
    try:
        mem = psutil.virtual_memory()
        info['Memoria RAM'] = {
            'Total': f"{round(mem.total / (1024**3), 2)} GB",
            'Disponible': f"{round(mem.available / (1024**3), 2)} GB",
            'En uso': f"{round(mem.used / (1024**3), 2)} GB",
            'Porcentaje en uso': f"{mem.percent}%"
        }
    except Exception as e:
        info['Memoria RAM'] = {'Error': f"No se pudo obtener información: {str(e)}"}
    
    # Información de discos
    try:
        disks = []
        for partition in psutil.disk_partitions():
            if 'fixed' in partition.opts or 'remote' in partition.opts:
                try:
                    usage = psutil.disk_usage(partition.mountpoint)
                    disks.append({
                        'Dispositivo': partition.device,
                        'Punto de Montaje': partition.mountpoint,
                        'Sistema de Archivos': partition.fstype,
                        'Espacio Total': f"{round(usage.total / (1024**3), 2)} GB",
                        'Espacio Usado': f"{round(usage.used / (1024**3), 2)} GB",
                        'Espacio Libre': f"{round(usage.free / (1024**3), 2)} GB",
                        'Porcentaje Usado': f"{usage.percent}%"
                    })
                except:
                    continue
        info['Discos'] = disks if disks else [{'Error': 'No se encontraron discos'}]
    except Exception as e:
        info['Discos'] = {'Error': f"No se pudo obtener información: {str(e)}"}
    
    # Información de batería
    try:
        battery = psutil.sensors_battery()
        if battery:
            battery_info = {
                'Porcentaje': f"{battery.percent}%",
                'Estado': "Cargando" if battery.power_plugged else "Descargando",
                'Tiempo estimado': "Calculando..." if battery.power_plugged else f"{round(battery.secsleft/3600, 2)} horas" if battery.secsleft else "Desconocido"
            }
            
            # Estimación de consumo
            if not battery.power_plugged:
                consumption = "Moderado"
                cpu_usage = psutil.cpu_percent(interval=1)
                if cpu_usage > 70: consumption = "Alto"
                elif cpu_usage < 30: consumption = "Bajo"
                battery_info['Consumo estimado'] = consumption
                
            info['Batería'] = battery_info
        else:
            info['Batería'] = {'Estado': 'No detectada'}
    except Exception as e:
        info['Batería'] = {'Error': f"No se pudo obtener información: {str(e)}"}
    
    # Información de red
    try:
        net_info = []
        for interface, addrs in psutil.net_if_addrs().items():
            for addr in addrs:
                if addr.family == socket.AF_INET:
                    net_info.append({
                        'Interfaz': interface,
                        'Dirección IP': addr.address,
                        'Máscara de Red': addr.netmask,
                        'Estado': "Conectado" if interface in psutil.net_if_stats() and psutil.net_if_stats()[interface].isup else "Desconectado"
                    })
        info['Red'] = net_info if net_info else [{'Error': 'No se encontraron interfaces de red'}]
    except Exception as e:
        info['Red'] = {'Error': f"No se pudo obtener información: {str(e)}"}
    
    # Información de GPU
    try:
        gpus = safe_wmi_query("Win32_VideoController", ["Name", "AdapterRAM", "CurrentHorizontalResolution", "CurrentVerticalResolution", "DriverVersion"])
        for gpu in gpus:
            if gpu.get('AdapterRAM') not in ['N/A', 'Error']:
                gpu['AdapterRAM'] = f"{round(int(gpu['AdapterRAM']) / (1024**3), 2)} GB"
        info['GPU'] = gpus if gpus else [{'Error': 'No se detectaron GPUs'}]
    except Exception as e:
        info['GPU'] = {'Error': f"No se pudo obtener información: {str(e)}"}
    
    return info

def check_headphone_jack():
    jack_info = {'Estado': 'No verificado'}
    
    try:
        import winsound
        try:
            winsound.Beep(1000, 100)
            jack_info['Estado'] = 'Puerto Audio Existente'
            jack_info['Prueba'] = 'Conecte auriculares para verificar si sale sonido'
        except:
            jack_info['Estado'] = 'Posible problema de audio'
            
        # Verificación básica con WMI
        try:
            c = wmi.WMI()
            jacks = [dev for dev in c.Win32_SoundDevice() 
                    if 'jack' in dev.Name.lower()]
            jack_info['Puerto_Detectado'] = 'Sí' if jacks else 'No'
        except:
            pass
            
    except Exception as e:
        jack_info['Error'] = str(e)
    
    return jack_info

def check_ports():
    """Verifica los puertos disponibles y su estado"""
    print("\n=== INFORMACIÓN DE PUERTOS ===")
    
    ports_info = {}
    c = wmi.WMI()
    
    # Puertos USB
    try:
        usb_devices = safe_wmi_query("Win32_USBHub", ["Name", "Status", "PNPDeviceID"])
        ports_info['Puertos USB'] = {
            'Total detectados': len(usb_devices),
            'Dispositivos': usb_devices,
            'Nota': 'Verificar físicamente conectando dispositivos. Algunos pueden ser controladores internos.'
        }
    except Exception as e:
        ports_info['Puertos USB'] = {'Error': f"No se pudo obtener información: {str(e)}"}
    
    # Puerto AUX
    ports_info['Jack_Auriculares'] = check_headphone_jack()  


    # HDMI
    try:
        hdmi_devices = []
        for display in safe_wmi_query("Win32_DesktopMonitor", ["Name", "Status", "PNPDeviceID"]):
            if 'HDMI' in str(display.get('Name', '')).upper() or 'HDMI' in str(display.get('PNPDeviceID', '')).upper():
                hdmi_devices.append(display)
        ports_info['HDMI'] = {
            'Dispositivos': hdmi_devices if hdmi_devices else [{'Estado': 'No se detectaron salidas HDMI'}],
            'Nota': 'Verificar físicamente conectando un monitor externo.'
        }
    except Exception as e:
        ports_info['HDMI'] = {'Error': f"No se pudo obtener información: {str(e)}"}
    
    # Bluetooth
    try:
        bt_devices = []
        for device in safe_wmi_query("Win32_PnPEntity", ["Name", "Status"]):
            if device.get('Name') and 'bluetooth' in str(device.get('Name')).lower():
                bt_devices.append(device)
        ports_info['Bluetooth'] = {
            'Dispositivos': bt_devices if bt_devices else [{'Estado': 'No se detectaron dispositivos Bluetooth'}],
            'Estado': 'Disponible' if bt_devices else 'No disponible'
        }
    except Exception as e:
        ports_info['Bluetooth'] = {'Error': f"No se pudo obtener información: {str(e)}"}
    
    # WiFi
    try:
        wifi_adapters = []
        for adapter in safe_wmi_query("Win32_NetworkAdapter", ["Name", "NetConnectionStatus", "NetEnabled"]):
            if adapter.get('Name') and ('wireless' in str(adapter.get('Name')).lower() or 'wi-fi' in str(adapter.get('Name')).lower()):
                wifi_adapters.append({
                    'Nombre': adapter.get('Name'),
                    'Estado': 'Habilitado' if adapter.get('NetEnabled') == 1 else 'Deshabilitado',
                    'Conexión': 'Conectado' if adapter.get('NetConnectionStatus') == 2 else 'Desconectado'
                })
        ports_info['WiFi'] = {
            'Adaptadores': wifi_adapters if wifi_adapters else [{'Estado': 'No se detectaron adaptadores WiFi'}],
            'Estado': 'Disponible' if wifi_adapters else 'No disponible'
        }
    except Exception as e:
        ports_info['WiFi'] = {'Error': f"No se pudo obtener información: {str(e)}"}
    
    return ports_info

def get_installed_software():
    """Obtiene una lista de software instalado"""
    print("\n=== SOFTWARE INSTALADO ===")
    
    software_list = []
    try:
        # Software de 32 bits
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
        for i in range(0, winreg.QueryInfoKey(key)[0]):
            try:
                subkey_name = winreg.EnumKey(key, i)
                subkey = winreg.OpenKey(key, subkey_name)
                name = winreg.QueryValueEx(subkey, "DisplayName")[0]
                version = winreg.QueryValueEx(subkey, "DisplayVersion")[0] if winreg.QueryValueEx(subkey, "DisplayVersion") else "N/A"
                publisher = winreg.QueryValueEx(subkey, "Publisher")[0] if winreg.QueryValueEx(subkey, "Publisher") else "N/A"
                install_date = winreg.QueryValueEx(subkey, "InstallDate")[0] if winreg.QueryValueEx(subkey, "InstallDate") else "N/A"
                software_list.append({
                    'Nombre': name,
                    'Versión': version,
                    'Publicador': publisher,
                    'Fecha de Instalación': install_date
                })
            except:
                continue
        
        # Software de 64 bits
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall")
        for i in range(0, winreg.QueryInfoKey(key)[0]):
            try:
                subkey_name = winreg.EnumKey(key, i)
                subkey = winreg.OpenKey(key, subkey_name)
                name = winreg.QueryValueEx(subkey, "DisplayName")[0]
                version = winreg.QueryValueEx(subkey, "DisplayVersion")[0] if winreg.QueryValueEx(subkey, "DisplayVersion") else "N/A"
                publisher = winreg.QueryValueEx(subkey, "Publisher")[0] if winreg.QueryValueEx(subkey, "Publisher") else "N/A"
                install_date = winreg.QueryValueEx(subkey, "InstallDate")[0] if winreg.QueryValueEx(subkey, "InstallDate") else "N/A"
                software_list.append({
                    'Nombre': name,
                    'Versión': version,
                    'Publicador': publisher,
                    'Fecha de Instalación': install_date
                })
            except:
                continue
    except Exception as e:
        return [{'Error': f"No se pudo obtener software instalado: {str(e)}"}]
    
    # Ordenar y limitar a 20 elementos
    return sorted([s for s in software_list if s.get('Nombre')], key=lambda x: x['Nombre'])[:20]

def get_boot_time():
    """Obtiene información del tiempo de arranque"""
    try:
        boot_time = psutil.boot_time()
        return {
            'Último arranque': datetime.datetime.fromtimestamp(boot_time).strftime("%Y-%m-%d %H:%M:%S"),
            'Tiempo activo': str(datetime.datetime.now() - datetime.datetime.fromtimestamp(boot_time))
        }
    except Exception as e:
        return {'Error': f"No se pudo obtener: {str(e)}"}

def get_windows_activation_status():
    """Verifica el estado de activación de Windows"""
    try:
        result = subprocess.run(['cscript', os.path.join(os.environ['SYSTEMROOT'], 'System32', 'slmgr.vbs'), '/dli'], 
                               capture_output=True, text=True, check=True)
        activation_info = {}
        for line in result.stdout.split('\n'):
            if 'Nombre:' in line:
                activation_info['Nombre'] = line.split(':')[1].strip()
            elif 'Estado de la licencia:' in line:
                activation_info['Estado'] = line.split(':')[1].strip()
        return activation_info
    except Exception as e:
        return {'Error': f"No se pudo verificar: {str(e)}"}


def get_bios_info():
    """Obtiene información detallada de la BIOS"""
    print("\n=== INFORMACIÓN DE LA BIOS ===")
    
    bios_info = {}
    c = wmi.WMI()
    
    try:
        # Información básica de la BIOS
        bios_data = safe_wmi_query("Win32_BIOS", [
            "Manufacturer", 
            "Name", 
            "Version", 
            "ReleaseDate", 
            "SerialNumber",
            "SMBIOSBIOSVersion",
            "SMBIOSMajorVersion",
            "SMBIOSMinorVersion",
            "Status"
        ])[0]
        
        bios_info = {
            'Fabricante': bios_data.get('Manufacturer', 'N/A'),
            'Nombre': bios_data.get('Name', 'N/A'),
            'Versión': bios_data.get('Version', 'N/A'),
            'Fecha de Lanzamiento': bios_data.get('ReleaseDate', 'N/A'),
            'Número de Serie': bios_data.get('SerialNumber', 'N/A'),
            'Versión SMBIOS': bios_data.get('SMBIOSBIOSVersion', 'N/A'),
            'Estado': bios_data.get('Status', 'N/A')
        }
        
        # Verificación de estado
        if bios_info['Estado'] == 'OK':
            bios_info['Estado_Verificacion'] = 'Funcional'
        else:
            bios_info['Estado_Verificacion'] = 'Requiere atención'
            
        # Detección de modo seguro
        try:
            secure_boot = subprocess.check_output(
                "powershell Confirm-SecureBootUEFI",
                shell=True,
                stderr=subprocess.PIPE
            ).decode().strip()
            bios_info['SecureBoot'] = 'Activado' if 'True' in secure_boot else 'Desactivado'
        except:
            bios_info['SecureBoot'] = 'No soportado o error'
            
    except Exception as e:
        bios_info['Error'] = f"No se pudo obtener información: {str(e)}"
    
    return bios_info


def check_health():
    """Verifica el estado de salud del sistema"""
    print("\n=== ESTADO DE SALUD DEL SISTEMA ===")
    
    health_info = {}
    
    # Temperatura
    try:
        temps = psutil.sensors_temperatures()
        if temps:
            temp_info = {}
            for name, entries in temps.items():
                temp_info[name] = [{'Sensor': entry.label, 'Actual': entry.current, 'Máxima': entry.high} for entry in entries]
            health_info['Temperaturas'] = temp_info
        else:
            health_info['Temperaturas'] = {'Estado': 'No hay sensores disponibles'}
    except Exception as e:
        health_info['Temperaturas'] = {'Error': f"No disponible: {str(e)}"}
    
    # CPU
    try:
        cpu_percent = psutil.cpu_percent(interval=1, percpu=True)
        health_info['CPU'] = {
            'Uso total': f"{sum(cpu_percent)/len(cpu_percent):.1f}%",
            'Uso por núcleo': [f"{p}%" for p in cpu_percent],
            'Frecuencia actual': f"{psutil.cpu_freq().current:.1f} MHz" if hasattr(psutil, 'cpu_freq') and psutil.cpu_freq() else 'N/A'
        }
    except Exception as e:
        health_info['CPU'] = {'Error': f"No se pudo obtener: {str(e)}"}
    
    # Discos
    try:
        disk_health = []
        for disk in psutil.disk_partitions():
            try:
                usage = psutil.disk_usage(disk.mountpoint)
                disk_health.append({
                    'Disco': disk.device,
                    'Uso': f"{usage.percent}%",
                    'Estado': 'OK' if usage.percent < 90 else 'ALTO USO'
                })
            except:
                continue
        health_info['Discos'] = disk_health if disk_health else [{'Estado': 'No se pudieron verificar discos'}]
    except Exception as e:
        health_info['Discos'] = {'Error': f"No se pudo obtener: {str(e)}"}
    
    return health_info

def prueba_microfono_sonido():
    """Función para probar micrófono y sistema de sonido"""
    print("\n=== PRUEBA DE MICRÓFONO Y SONIDO ===")
    print("1. Se grabará audio por 3 segundos")
    print("2. Se reproducirá lo grabado")
    print("Preparado? Presione Enter para comenzar...")
    input()
    
    try:
        # Configuración de audio
        fs = 44100  # Frecuencia de muestreo
        duration = 3  # Duración en segundos
        
        print(f"\nGrabando audio ({duration} segundos)...")
        recording = sd.rec(int(duration * fs), samplerate=fs, channels=2)
        sd.wait()  # Espera hasta que termine la grabación
        print("Grabación completada")
        
        print("Reproduciendo audio grabado...")
        sd.play(recording, fs)
        sd.wait()
        print("Prueba completada con éxito!")
        return True
    except Exception as e:
        print(f"\nError en prueba de audio: {str(e)}")
        print("El micrófono o sistema de sonido no funciona correctamente")
        return False

def prueba_camara():
    """Función para probar la cámara web"""
    print("\n=== PRUEBA DE CÁMARA ===")
    print("1. Se abrirá la cámara por 5 segundos")
    print("2. Mire directamente a la cámara")
    print("Preparado? Presione Enter para comenzar...")
    input()
    
    cap = None
    try:
        cap = cv2.VideoCapture(0)
        if not cap.isOpened():
            print("No se pudo abrir la cámara")
            return False
            
        print("\nCámara activada - Sonría! (5 segundos)")
        start_time = time.time()
        
        while (time.time() - start_time) < 5:
            ret, frame = cap.read()
            if not ret:
                break
            cv2.imshow('Prueba de Cámara (Presione Q para salir)', frame)
            if cv2.waitKey(1) & 0xFF == ord('q'):
                break
                
        print("Prueba de cámara completada")
        return True
        
    except Exception as e:
        print(f"\nError en prueba de cámara: {str(e)}")
        return False
    finally:
        if cap is not None:
            cap.release()
        cv2.destroyAllWindows()

def save_to_file(data, filename="notebook_report.json"):
    """Guarda el reporte en un archivo JSON"""
    try:
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
        print(f"\nReporte guardado en {filename}")
    except Exception as e:
        print(f"\nError al guardar el reporte: {str(e)}")

def clear_screen():
    os.system('cls' if os.name == 'nt' else 'clear')

def main():
    clear_screen()
    print("===🔧 INICIANDO REVISIÓN DEL DISPOSITIVO 🔧===")
    
    report = {
        'Información General': get_system_info(),
        'Puertos': check_ports(),
        'BIOS': get_bios_info(), 
        'Estado de Salud': check_health(),
        'Tiempo de Arranque': get_boot_time(),
        'Activación de Windows': get_windows_activation_status(),
        'Software Instalado (primeros 20)': get_installed_software(),
        'Fecha de Revisión': datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    

    # Mostrar resumen en pantalla
    print(f"\n{Fore.YELLOW}=== RESUMEN ==={Style.RESET_ALL}")
    # BIOS
    print(f"\n{Fore.CYAN}BIOS 🧬:{Style.RESET_ALL}")
    bios_info = report['BIOS']
    print(f"  {Fore.CYAN}Fabricante:{Style.RESET_ALL} {bios_info.get('Fabricante', 'N/A')}")
    print(f"  {Fore.CYAN}Versión:{Style.RESET_ALL} {bios_info.get('Versión', 'N/A')}")
    print(f"  {Fore.CYAN}Estado:{Style.RESET_ALL} {bios_info.get('Estado_Verificacion', 'N/A')}")
    print(f"  {Fore.CYAN}Secure Boot:{Style.RESET_ALL} {bios_info.get('SecureBoot', 'N/A')}")

    # Sistema operativo
    print ("")
    print(f"{Fore.CYAN}Sistema Operativo 🖥️ :{Style.RESET_ALL} {report['Información General'].get('Sistema Operativo', 'N/A')}")
    
    # Procesador
    cpu_info = report['Información General'].get('Procesador', {})
    print(f"\n{Fore.CYAN}Procesador 🧠 :{Style.RESET_ALL} {cpu_info.get('Modelo', 'N/A')}")
    print(f"{Fore.CYAN}Núcleos:{Style.RESET_ALL} {cpu_info.get('Núcleos Físicos', 'N/A')} físicos, {cpu_info.get('Núcleos Lógicos', 'N/A')} lógicos")
    
    # Memoria RAM
    ram_info = report['Información General'].get('Memoria RAM', {})
    print(f"\n{Fore.CYAN}Memoria RAM 💾:{Style.RESET_ALL}")
    print(f"  {Fore.CYAN}Total:{Style.RESET_ALL} {ram_info.get('Total', 'N/A')}")
    print(f"  {Fore.CYAN}Disponible:{Style.RESET_ALL} {ram_info.get('Disponible', 'N/A')}")
    print(f"  {Fore.CYAN}En uso:{Style.RESET_ALL} {ram_info.get('En uso', 'N/A')} ({ram_info.get('Porcentaje en uso', 'N/A')})")
    
    # Batería
    battery_info = report['Información General'].get('Batería', {})
    if battery_info:
        print(f"\n{Fore.CYAN}Batería 🗃️:{Style.RESET_ALL}")
        print(f"  {Fore.CYAN}Nivel:{Style.RESET_ALL} {battery_info.get('Porcentaje', 'N/A')}")
        print(f"  {Fore.CYAN}Estado:{Style.RESET_ALL} {battery_info.get('Estado', 'N/A')}")
        if 'Tiempo estimado' in battery_info:
            print(f"  {Fore.CYAN}Tiempo restante:{Style.RESET_ALL} {battery_info['Tiempo estimado']}")
        if 'Consumo estimado' in battery_info:
            print(f"  {Fore.CYAN}Consumo:{Style.RESET_ALL} {battery_info['Consumo estimado']}")
    
    # Puertos USB
    usb_info = report['Puertos'].get('Puertos USB', {})
    print(f"\n{Fore.CYAN}Puertos USB 🔌:{Style.RESET_ALL}")
    print(f"  {Fore.CYAN}Total detectados:{Style.RESET_ALL} {usb_info.get('Total detectados', 'N/A')}")
    for i, device in enumerate(usb_info.get('Dispositivos', [])[:3], 1):
        print(f"  {Fore.CYAN}Dispositivo {i}:{Style.RESET_ALL} {device.get('Name', 'N/A')}")
        print(f"    {Fore.CYAN}Estado:{Style.RESET_ALL} {device.get('Status', 'N/A')}")

    # Puertos audio 
    jack_info = report['Puertos'].get('Jack_Auriculares', {})
    print(f"\n{Fore.CYAN}Puerto Auriculares 🎧:{Style.RESET_ALL}")
    print(f"  {Fore.CYAN}Estado:{Style.RESET_ALL} {jack_info.get('Estado', 'N/A')}")
    print(f"  {Fore.CYAN}Detectado:{Style.RESET_ALL} {jack_info.get('Detectado_Logicamente', 'N/A')}")
    if jack_info.get('Prueba_Sonido'):
        print(f"  {Fore.CYAN}Prueba de sonido:{Style.RESET_ALL} {jack_info['Prueba_Sonido']}")
    
    # Bluetooth y WiFi
    print(f"\n{Fore.CYAN}Bluetooth 🛰️:{Style.RESET_ALL} {report['Puertos'].get('Bluetooth', {}).get('Estado', 'N/A')}")
    print(f"{Fore.CYAN}WiFi 🌐:{Style.RESET_ALL} {report['Puertos'].get('WiFi', {}).get('Estado', 'N/A')}")
    
    # Estado de salud
    health_info = report['Estado de Salud']
    print(f"\n{Fore.CYAN}Estado de salud 🧪:{Style.RESET_ALL}")
    print(f"  {Fore.CYAN}Uso de CPU:{Style.RESET_ALL} {health_info.get('CPU', {}).get('Uso total', 'N/A')}")
    if 'Discos' in health_info and health_info['Discos']:
        print(f"  {Fore.CYAN}Uso de disco principal:{Style.RESET_ALL} {health_info['Discos'][0].get('Uso', 'N/A')}")
    
    # Guardar reporte completo
    save_to_file(report)
    
    print(f"\n{Fore.YELLOW}=== RECOMENDACIONES ==={Style.RESET_ALL}")
    print(f"{Fore.GREEN}1. Verificar físicamente todos los puertos USB conectando dispositivos{Style.RESET_ALL}")
    print(f"{Fore.GREEN}2. Probar el puerto de audio con auriculares y micrófono{Style.RESET_ALL}")
    print(f"{Fore.GREEN}3. Conectar un monitor externo para probar HDMI{Style.RESET_ALL}")
    print(f"{Fore.GREEN}4. Verificar el estado de la batería (tiempo de duración real){Style.RESET_ALL}")
    print(f"{Fore.GREEN}5. Probar conexiones Bluetooth y WiFi{Style.RESET_ALL}")
    print(f"{Fore.GREEN}6. Revisar el estado físico del notebook (teclado, pantalla, bisagras){Style.RESET_ALL}")
    print(f"{Fore.GREEN}7. Comprobar que no hay sectores dañados en los discos duros{Style.RESET_ALL}")
    print(f"{Fore.GREEN}8. Verificar que todos los núcleos del procesador funcionan correctamente{Style.RESET_ALL}")

    # Después de mostrar el resumen inicial:
    print(f"\n{Fore.YELLOW}=== PRUEBAS ADICIONALES ==={Style.RESET_ALL}")
    
    # Prueba de audio
    input(f"\n{Fore.MAGENTA}Presione Enter para iniciar prueba de micrófono y sonido...{Style.RESET_ALL}")
    audio_ok = prueba_microfono_sonido()
    
    # Prueba de cámara
    input(f"\n{Fore.MAGENTA}Presione Enter para iniciar prueba de cámara...{Style.RESET_ALL}")
    camara_ok = prueba_camara()
    
    # Resultados finales
    print(f"\n{Fore.YELLOW}=== RESULTADOS PRUEBAS ==={Style.RESET_ALL}")
    print(f"Micrófono/Sonido: {Fore.GREEN if audio_ok else Fore.RED}{'OK' if audio_ok else 'FALLÓ'}{Style.RESET_ALL}")
    print(f"Cámara: {Fore.GREEN if camara_ok else Fore.RED}{'OK' if camara_ok else 'FALLÓ'}{Style.RESET_ALL}")

if __name__ == "__main__":
    try:
        if sys.stdout.encoding != 'UTF-8':
            sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        if sys.stderr.encoding != 'UTF-8':
            sys.stderr.reconfigure(encoding='utf-8', errors='replace')
        
        clear_screen()
        main()
        input(f"\n{Fore.GREEN}Presione Enter para salir...{Style.RESET_ALL}")
        
    except KeyboardInterrupt:
        print(f"\n{Fore.RED}Ejecución cancelada por el usuario{Style.RESET_ALL}")
        sys.exit(0)
    except Exception as e:
        print(f"\n{Fore.RED}Error crítico:{Style.RESET_ALL} {str(e)}", file=sys.stderr)
        input("\nPresione Enter para salir...")
        sys.exit(1)