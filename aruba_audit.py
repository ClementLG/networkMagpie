# aruba_audit.py
# !/usr/bin/env python3

import json
import csv
import os
import datetime
import re
from netmiko import ConnectHandler
from netmiko.exceptions import NetmikoTimeoutException, NetmikoAuthenticationException, SSHException
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import traceback

# --- Copier les constantes Excel et les fonctions d'aide Excel de cisco_audit.py ---
GREEN_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
ORANGE_FILL = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
BLUE_FILL = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
BOLD_FONT = Font(bold=True)


def apply_header_style(ws, row_num=1):  # Identique à cisco_audit.py
    for cell in ws[row_num]:
        cell.font = BOLD_FONT;
        cell.fill = BLUE_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")


def auto_fit_columns(ws):  # Identique à cisco_audit.py
    for col in ws.columns:
        max_length = 0;
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length: max_length = len(str(cell.value))
            except:
                pass
        ws.column_dimensions[column].width = (max_length + 2)


def set_cell_status_color(cell, level):  # Identique à cisco_audit.py
    if level == "good":
        cell.fill = GREEN_FILL
    elif level == "warning":
        cell.fill = ORANGE_FILL
    elif level == "error" or level == "bad":
        cell.fill = RED_FILL


# --- Fonctions de collecte Aruba OS-CX ---
def get_aruba_device_info(net_connect):
    info = {}
    try:
        # OS-CX 'show version' est souvent parsable par TextFSM avec le template 'hp_procurve_show_version' ou un générique.
        # Alternative: 'show system'
        version_out = net_connect.send_command("show version", use_textfsm=True, expect_string=r"#")
        system_out = net_connect.send_command("show system", use_textfsm=True,
                                              expect_string=r"#")  # Plus d'infos sur le modèle/série

        if version_out and isinstance(version_out, list):
            ver_data = version_out[0]
            info['hostname'] = ver_data.get('hostname', net_connect.base_prompt.strip("#> "))
            info['ios_version'] = ver_data.get('version',
                                               ver_data.get('software_version', 'N/A'))  # Clés peuvent varier
        else:  # Fallback si TextFSM échoue
            raw_ver = net_connect.send_command("show version", expect_string=r"#")
            # Regex basique pour hostname et version (à adapter)
            match_hostname = re.search(r"Hostname\s*:\s*(\S+)", raw_ver, re.IGNORECASE)
            if match_hostname:
                info['hostname'] = match_hostname.group(1)
            else:
                info['hostname'] = net_connect.base_prompt.strip("#> ")
            match_version = re.search(r"Software version\s*:\s*(\S+)", raw_ver, re.IGNORECASE)
            if match_version:
                info['ios_version'] = match_version.group(1)
            else:
                info['ios_version'] = "N/A"

        if system_out and isinstance(system_out, list):
            sys_data = system_out[0]
            info['model'] = sys_data.get('product_name', sys_data.get('model', 'N/A'))
            info['serial_number'] = sys_data.get('serial_number', 'N/A')
        else:  # Fallback
            raw_sys = net_connect.send_command("show system", expect_string=r"#")
            match_model = re.search(r"Product Name\s*:\s*(\S+)", raw_sys, re.IGNORECASE)
            if match_model:
                info['model'] = match_model.group(1)
            else:
                info['model'] = "N/A"
            match_serial = re.search(r"Serial Number\s*:\s*(\S+)", raw_sys, re.IGNORECASE)
            if match_serial:
                info['serial_number'] = match_serial.group(1)
            else:
                info['serial_number'] = "N/A"

        # Uptime - 'show system' peut le contenir, ou 'show uptime'
        uptime_out = net_connect.send_command("show uptime", expect_string=r"#")  # Pas de TextFSM pour ça souvent
        match_uptime = re.search(r"System Uptime\s*:\s*(.*)", uptime_out)
        if match_uptime:
            info['uptime'] = match_uptime.group(1).strip()
        else:
            info['uptime'] = "N/A"

        return info
    except Exception as e:
        print(f"  Erreur get_aruba_device_info: {e}")
        return {'hostname': 'Error', 'ios_version': 'Error', 'model': 'Error', 'uptime': 'Error',
                'serial_number': 'Error'}


def get_aruba_interfaces(net_connect):
    interfaces = []
    try:
        # OS-CX 'show interface brief'
        # La sortie est souvent: Port | Admin | Link | Speed | Duplex | Type | Vlans
        # Ou pour les IPs: Interface | IP Address      | Status      | Protocol
        # Il faut peut-être deux commandes ou un parsing intelligent

        # Pour les IPs et statut L3
        ip_brief_out = net_connect.send_command("show ip interface brief", use_textfsm=True, expect_string=r"#")
        ip_brief_list = ip_brief_out if isinstance(ip_brief_out, list) else []

        # Pour le statut L1/L2 et détails
        int_status_out = net_connect.send_command("show interface status", use_textfsm=True,
                                                  expect_string=r"#")  # Pourrait être 'show interface all'
        int_status_list = int_status_out if isinstance(int_status_out, list) else []

        # Descriptions
        # descriptions_out = net_connect.send_command("show interface description", use_textfsm=True, expect_string=r"#")
        # desc_list = descriptions_out if isinstance(descriptions_out, list) else []
        # desc_map = {item.get('port', item.get('interface')): item.get('description') for item in desc_list if item.get('port',item.get('interface')) and item.get('description')}
        # Alternative pour descriptions à partir de la config
        running_config = net_connect.send_command("show running-config", read_timeout=120, expect_string=r"#")
        desc_map = {}
        for line in running_config.splitlines():
            if line.strip().startswith("interface "):
                current_if = line.strip().split()[-1]
            elif "description " in line.strip() and current_if:
                desc_map[current_if] = line.strip().split("description ", 1)[1]

        # Créer un map pour les statuts L1/L2
        status_l2_map = {item.get('port', item.get('interface')): item for item in int_status_list}

        for iface_l3 in ip_brief_list:
            name = iface_l3.get('interface', iface_l3.get('intf'))
            if not name: continue

            status_l2_detail = status_l2_map.get(name, {})

            interfaces.append({
                "name": name,
                "ip_address": iface_l3.get('ip_address', iface_l3.get('ipaddr', 'N/A')),
                "status_link": status_l2_detail.get('link_status', iface_l3.get('status', 'N/A')).lower(),
                # 'status' de ip_brief est L3
                "status_protocol": iface_l3.get('protocol', 'N/A').lower(),  # 'protocol' de ip_brief est L3
                "description": desc_map.get(name, "N/A"),
                "type": "Virtual" if name.lower().startswith(("vlan", "loopback", "lag")) else "Physical",
                "vlan": status_l2_detail.get('vlan', 'N/A'),
                # Pour les access ports, 'show vlan port detail' serait mieux
                "duplex": status_l2_detail.get('duplex', 'N/A'),
                "speed": status_l2_detail.get('speed', 'N/A'),
            })
        return interfaces
    except Exception as e:
        print(f"  Erreur get_aruba_interfaces: {e}");
        traceback.print_exc()
        return []


def get_aruba_vlans(net_connect):
    vlans_list = []
    try:
        # 'show vlan brief' ou 'show vlan'
        vlan_out = net_connect.send_command("show vlan brief", use_textfsm=True,
                                            expect_string=r"#")  # Vérifier le template
        vlan_data = vlan_out if isinstance(vlan_out, list) else []

        if not vlan_data:  # Fallback si TextFSM échoue ou si la commande est 'show vlan'
            raw_vlan_out = net_connect.send_command("show vlan", use_textfsm=False, expect_string=r"#")
            # Parsing Regex ou ligne par ligne (complexe)
            # Exemple de ligne 'show vlan': VLAN ID Name                 Status    Ports
            # Exemple de ligne 'show vlan brief': 1       DEFAULT_VLAN         Port-up   1/1/1-1/1/2, lag1
            current_vlan_id, current_vlan_name, current_vlan_status = None, None, None
            for line in raw_vlan_out.splitlines():
                match_vlan = re.match(r"^\s*(\d+)\s+(\S+)\s+(\S+)\s*(.*)", line)  # Simplifié
                if match_vlan:
                    vid, vname, vstatus, vports_str = match_vlan.groups()
                    vports = [p.strip() for p in vports_str.split(',') if p.strip()]
                    vlans_list.append({
                        "id": vid, "name": vname, "status": vstatus,
                        "ports": ", ".join(vports) if vports else "N/A"
                    })

        else:  # Traitement de la sortie TextFSM si elle a fonctionné
            for v in vlan_data:
                vlans_list.append({
                    "id": v.get('vlan_id', 'N/A'),
                    "name": v.get('name', 'N/A'),
                    "status": v.get('status', 'N/A'),
                    "ports": ", ".join(v.get('ports', [])) if v.get('ports') else "N/A"
                })
        return vlans_list
    except Exception as e:
        print(f"  Erreur get_aruba_vlans: {e}");
        traceback.print_exc()
        return []


def get_aruba_arp_table(net_connect):
    arp_table = []
    try:
        # 'show arp'
        arp_out = net_connect.send_command("show arp", use_textfsm=True, expect_string=r"#")  # Vérifier template
        arp_data = arp_out if isinstance(arp_out, list) else []

        if not arp_data:  # Fallback
            raw_arp_out = net_connect.send_command("show arp", use_textfsm=False, expect_string=r"#")
            # Exemple: 192.168.1.1   00:01:02:03:04:05   vlan10    1/1/1    dynamic
            for line in raw_arp_out.splitlines():
                # Regex simplifié, à adapter selon format exact
                match_arp = re.match(r"^\s*([\d\.]+)\s+([0-9a-f\.\:]+)\s+(\S+)\s+\S+\s+\S+", line, re.IGNORECASE)
                if match_arp:
                    ip, mac, intf = match_arp.groups()
                    arp_table.append({
                        "protocol": "Internet", "address": ip, "age": "N/A",  # Age non dispo dans ce format simple
                        "mac_address": mac, "type": "ARPA", "interface": intf
                    })
        else:
            for entry in arp_data:
                arp_table.append({
                    "protocol": entry.get('protocol', 'Internet'),
                    "address": entry.get('address', entry.get('ip_address', 'N/A')),
                    "age": entry.get('age', 'N/A'),
                    "mac_address": entry.get('mac', entry.get('mac_address', 'N/A')).replace(':', ''),
                    # Standardiser format MAC
                    "type": entry.get('type', 'ARPA'),
                    "interface": entry.get('interface', 'N/A')
                })
        return arp_table
    except Exception as e:
        print(f"  Erreur get_aruba_arp_table: {e}");
        traceback.print_exc()
        return []


def check_aruba_security_features(net_connect, running_config):
    security_audit = {}
    # Ceci est un SQUELETTE. Les commandes et la logique doivent être adaptées pour Aruba OS-CX.
    # Exemple:
    if "aaa authentication port-access" in running_config:  # Très simplifié
        security_audit["aaa_port_access"] = {"status": True, "level": "good",
                                             "details": "AAA pour l'accès port est configuré."}
    else:
        security_audit["aaa_port_access"] = {"status": False, "level": "warning",
                                             "details": "AAA pour l'accès port non configuré."}

    # Mot de passe local et politique
    if "password minimum-length" in running_config:
        match_pass_len = re.search(r"password minimum-length\s+(\d+)", running_config)
        min_len = int(match_pass_len.group(1)) if match_pass_len else 0
        level = "good" if min_len >= 12 else "warning" if min_len >= 8 else "bad"
        security_audit["password_min_length"] = {"status": f"Longueur min: {min_len}", "level": level,
                                                 "details": f"Longueur minimale des mots de passe: {min_len}. Recommandé: >=12."}
    else:
        security_audit["password_min_length"] = {"status": "Non configuré (défaut)", "level": "bad",
                                                 "details": "Politique de longueur minimale des mots de passe non configurée."}

    # SSH
    try:
        ssh_status = net_connect.send_command("show ip ssh", expect_string=r"#")
        if "SSH Protocol Version : SSHv2" in ssh_status and "SSHv1 Ciphers" not in ssh_status:  # Simplifié
            security_audit["ssh_v2_only"] = {"status": True, "level": "good",
                                             "details": "SSHv2 semble être la seule version activée."}
        elif "SSHv1 Ciphers" in ssh_status:
            security_audit["ssh_v2_only"] = {"status": False, "level": "bad",
                                             "details": "SSHv1 est activé. À désactiver."}
        else:
            security_audit["ssh_v2_only"] = {"status": "Vérification partielle", "level": "warning",
                                             "details": "Statut SSHv1/SSHv2 à vérifier plus en détail."}
    except:
        security_audit["ssh_v2_only"] = {"status": "Erreur", "level": "warning",
                                         "details": "Impossible de vérifier 'show ip ssh'."}

    # Ajouter d'autres checks : NTP, logging, SNMP, HTTP/HTTPS, CDP/LLDP, banners, etc.
    # Exemple pour NTP (très basique)
    if "ntp server" in running_config:
        ntp_servers = len(re.findall(r"ntp server\s+", running_config))
        level = "good" if ntp_servers >= 2 else "warning" if ntp_servers == 1 else "bad"
        security_audit["ntp_config"] = {"status": f"{ntp_servers} serveurs configurés", "level": level,
                                        "details": f"{ntp_servers} serveurs NTP. Recommandé: >=2."}
    else:
        security_audit["ntp_config"] = {"status": "Non configuré", "level": "bad", "details": "NTP non configuré."}

    return security_audit


# --- Fonctions principales Aruba ---
# Copier load_inventory et load_passwords de cisco_audit.py (elles sont génériques)
# Adapter generate_excel_report si la structure des données Aruba est très différente,
# sinon elle peut être réutilisée en changeant le nom du fichier de sortie.

def perform_aruba_audit(aruba_devices_inventory, global_passwords_map, output_directory):
    all_devices_data = []
    for device_entry in aruba_devices_inventory:
        host, group = device_entry["host"], device_entry["group"]
        creds = global_passwords_map.get(group)
        print(f"\n[INFO Aruba] Traitement de {host} (groupe: {group})...")

        if not creds:
            print(f"  [ERREUR Aruba] Identifiants non trouvés pour groupe '{group}'. {host} ignoré.")
            all_devices_data.append({"attempted_host": host, "status": "error_connection",
                                     "error_message": f"Identifiants non trouvés pour groupe {group}"})
            continue

        device_params = {
            'device_type': 'aruba_aoscx',  # Spécifique à Aruba OS-CX
            'host': host, 'username': creds['username'], 'password': creds['password'],
            'secret': creds.get('enable_password'),  # OS-CX utilise 'enable'
            'global_delay_factor': 2, 'timeout': 45, 'session_timeout': 120
        }
        current_device_data = {"host": host}
        try:
            with ConnectHandler(**device_params) as net_connect:
                if creds.get(
                    'enable_password'): net_connect.enable()  # OS-CX peut ne pas nécessiter 'enable' si privilèges suffisants

                print(f"  [OK Aruba] Connecté à {net_connect.host} ({net_connect.base_prompt[:-1]}).")

                current_device_data["general_info"] = get_aruba_device_info(net_connect)
                current_device_data["general_info"]["ip_address_queried"] = host

                running_config = net_connect.send_command("show running-config", read_timeout=240, expect_string=r"#")
                if not running_config: running_config = ""

                current_device_data["interfaces"] = get_aruba_interfaces(net_connect)
                current_device_data["vlans"] = get_aruba_vlans(net_connect)
                current_device_data["arp_table"] = get_aruba_arp_table(net_connect)
                current_device_data["security_audit"] = check_aruba_security_features(net_connect, running_config)
                all_devices_data.append(current_device_data)

        except Exception as e:
            print(f"  [ERREUR Aruba] Inattendue avec {host}: {e}");
            traceback.print_exc()
            all_devices_data.append(
                {"attempted_host": host, "status": "error_connection", "error_message": f"Erreur inattendue: {e}"})

    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    json_filename = os.path.join(output_directory, f"audit_aruba_data_{timestamp}.json")
    try:
        with open(json_filename, 'w', encoding='utf-8') as f:
            json.dump(all_devices_data, f, indent=4, ensure_ascii=False)
        print(f"\n[+] Données JSON Aruba sauvegardées : {json_filename}")
    except Exception as e:
        print(f"\n[-] Erreur sauvegarde JSON Aruba: {e}")

    # Réutiliser generate_excel_report, en s'assurant que les clés des dictionnaires sont similaires
    # ou en adaptant la fonction. Pour l'instant, on suppose une structure de données similaire.
    # Vous nommerez le fichier différemment dans main_audit.py ou ici.
    # Pour la simplicité de ce script autonome:
    excel_report_path = os.path.join(output_directory, f"audit_aruba_report_{timestamp}.xlsx")
    if 'generate_excel_report' not in globals():  # Si la fonction n'est pas définie ici
        print(
            f"  [AVERTISSEMENT] La fonction generate_excel_report n'est pas définie dans aruba_audit.py. Rapport Excel non généré par ce script.")
    else:
        generate_excel_report(all_devices_data, excel_report_path)  # Passez le chemin complet

    print(f"\n[+] Audit Aruba terminé pour {len(aruba_devices_inventory)} équipement(s).")


# Fonctions load_inventory et load_passwords (identiques à cisco_audit.py, doivent gérer device_type)
# COPIEZ-LES ICI DEPUIS LA VERSION CORRIGÉE DE cisco_audit.py

def load_inventory(filepath="inventory.csv"):
    # ... (code complet de load_inventory qui lit la 3ème colonne device_type)
    inventory = []
    try:
        with open(filepath, mode='r', encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            try:
                header = next(reader)
                if len(header) < 3:
                    print(
                        f"Erreur: En-tête inventaire '{filepath}' doit avoir au moins 3 colonnes (hostname, group, device_type).")
                    return None
            except StopIteration:
                return []
            for row in reader:
                if len(row) >= 3 and row[0].strip():
                    inventory.append({
                        "host": row[0].strip(), "group": row[1].strip(),
                        "device_type": row[2].strip().lower()
                    })
                elif row and any(field.strip() for field in row):
                    print(f"Avertissement: Ligne inventaire mal formatée: {row}")
        return inventory
    except FileNotFoundError:
        print(f"Erreur: Fichier inventaire '{filepath}' non trouvé."); return None
    except Exception as e:
        print(f"Erreur lecture '{filepath}': {e}"); traceback.print_exc(); return None


def load_passwords(filepath="passwords.csv"):  # Identique
    # ... (code complet de load_passwords)
    passwords = {}
    try:
        with open(filepath, mode='r', encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            try:
                next(reader)
            except StopIteration:
                return {}
            for row in reader:
                if len(row) >= 3 and row[0].strip():
                    enable_pass = row[3].strip() if len(row) > 3 and row[3].strip() else None
                    passwords[row[0].strip()] = {"username": row[1].strip(), "password": row[2].strip(),
                                                 "enable_password": enable_pass}
                elif row and any(field.strip() for field in row):
                    print(f"Avertissement: Ligne passwords mal formatée: {row}")
        return passwords
    except FileNotFoundError:
        print(f"Erreur: Fichier passwords '{filepath}' non trouvé."); return None
    except Exception as e:
        print(f"Erreur lecture '{filepath}': {e}"); traceback.print_exc(); return None


# Copier generate_excel_report de cisco_audit.py ici si vous voulez qu'il soit autonome
# Sinon, main_audit.py devra gérer la génération de rapport combiné ou appeler des fonctions spécifiques.
# Pour que ce script soit autonome, copions-la.
def generate_excel_report(all_data, excel_filepath):  # Modifié pour prendre un chemin complet
    if not all_data: print("Aucune donnée pour rapport Excel Aruba."); return
    # ... (code complet de generate_excel_report, identique à cisco_audit.py)
    # ... (mais il sauvegardera avec le excel_filepath fourni)
    wb = openpyxl.Workbook();
    wb.remove(wb.active)
    ws_info = wb.create_sheet("Infos Générales")
    headers_info = ["Hostname", "IP Address", "Modèle", "Version OS", "Uptime", "Numéro de Série"]
    ws_info.append(headers_info);
    apply_header_style(ws_info)
    for dev_data in all_data:
        if dev_data.get('status') == 'error_connection':
            ws_info.append(
                [dev_data.get('attempted_host', 'N/A'), dev_data.get('attempted_host', 'N/A'), "ERREUR CONNEXION",
                 dev_data.get('error_message', 'N/A'), "N/A", "N/A"])
            for i in range(3, 5): ws_info.cell(row=ws_info.max_row, column=i).fill = RED_FILL
        else:
            info = dev_data.get("general_info", {})
            ws_info.append([info.get("hostname", dev_data.get("host", "N/A")), dev_data.get("host", "N/A"),
                            info.get("model", "N/A"), info.get("ios_version", "N/A"), info.get("uptime", "N/A"),
                            info.get("serial_number", "N/A")])
    auto_fit_columns(ws_info)
    # ... (Continuer pour les autres feuilles : Interfaces, VLANs, ARP, Audit Sécurité) ...
    # ... (Assurez-vous que les clés correspondent à ce que vos fonctions get_aruba... retournent)
    # Exemple pour la feuille Interfaces (adaptez les clés si nécessaire)
    ws_interfaces = wb.create_sheet("Interfaces")
    headers_interfaces = ["Hostname", "Nom Interface", "Type", "Description", "IP Address", "Statut Lien",
                          "Statut Protocole", "VLAN (Access)", "Duplex", "Vitesse"]
    ws_interfaces.append(headers_interfaces);
    apply_header_style(ws_interfaces)
    for dev_data in all_data:
        if dev_data.get('status') == 'error_connection': continue
        hostname = dev_data.get("general_info", {}).get("hostname", dev_data.get("host", "N/A"))
        for iface in dev_data.get("interfaces", []):  # Les clés doivent correspondre à get_aruba_interfaces
            ws_interfaces.append(
                [hostname, iface.get("name"), iface.get("type"), iface.get("description"), iface.get("ip_address"),
                 iface.get("status_link"), iface.get("status_protocol"), iface.get("vlan"), iface.get("duplex"),
                 iface.get("speed")])
            # Coloration (identique si les statuts sont similaires)
            lr, pr = ws_interfaces.cell(row=ws_interfaces.max_row, column=6), ws_interfaces.cell(
                row=ws_interfaces.max_row, column=7)
            sl, sp = str(iface.get("status_link", "")).lower(), str(iface.get("status_protocol", "")).lower()
            if sl == "up" or "connected" in sl:
                lr.fill = GREEN_FILL
            elif ("admin" in sl and "down" in sl) or sl == "disabled":
                lr.fill = ORANGE_FILL
            elif sl == "down" or "notconnect" in sl or "err-disabled" in sl or "error-disabled" in sl:
                lr.fill = RED_FILL
            if sp == "up":
                pr.fill = GREEN_FILL
            elif sp == "down":
                pr.fill = RED_FILL
    auto_fit_columns(ws_interfaces)

    # ... (Répétez pour VLANs, ARP, Audit Sécurité, en adaptant les clés si besoin)

    try:
        wb.save(excel_filepath); print(f"\n[+] Rapport Excel Aruba généré : {excel_filepath}")
    except Exception as e:
        print(f"\n[-] Erreur sauvegarde Excel Aruba: {e}")


def main_aruba():  # Renommé pour éviter conflit si importé
    inventory_file, password_file, output_directory = "inventory.csv", "passwords.csv", "audit_reports"
    if not os.path.exists(output_directory):
        try:
            os.makedirs(output_directory)
        except OSError as e:
            print(f"Erreur création répertoire '{output_directory}': {e}"); return

    full_inventory = load_inventory(inventory_file)
    passwords_map = load_passwords(password_file)

    if full_inventory is None or passwords_map is None:
        print("Arrêt (Aruba): Erreurs critiques chargement fichiers entrée.");
        return

    aruba_devices = [device for device in full_inventory if device.get("device_type") == "aruba_os-cx"]

    if not aruba_devices:
        print("Aucun équipement Aruba OS-CX trouvé dans l'inventaire pour audit autonome.")
        return

    perform_aruba_audit(aruba_devices, passwords_map, output_directory)


if __name__ == "__main__":
    main_aruba()