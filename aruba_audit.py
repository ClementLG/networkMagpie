# aruba_audit.py
# !/usr/bin/env python3

# NetworkMagpie
# Copyright (C) 2025 CLEMENT LE GRUIEC
#
# This program is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 3 of the License, or
# (at your option) any later version.
#
# This program is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with this program.  If not, see <https://www.gnu.org/licenses/>.

import json
import csv
import os
import datetime
import re
import logging
import traceback
from netmiko import ConnectHandler
from netmiko.exceptions import NetmikoTimeoutException, NetmikoAuthenticationException, SSHException
from textfsm.parser import TextFSMError
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# Configure logging
logger = logging.getLogger(__name__)

# --- Excel Constants and Helper Functions ---
GREEN_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
ORANGE_FILL = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
BLUE_FILL = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
BOLD_FONT = Font(bold=True)


def apply_header_style(ws, row_num=1):
    """
    Applies a standard bold, blue styling to the header row of an Excel worksheet.

    Args:
        ws (openpyxl.worksheet.worksheet.Worksheet): The worksheet to apply styling to.
        row_num (int, optional): The row number to style as header. Defaults to 1.
    """
    for cell in ws[row_num]:
        cell.font = BOLD_FONT;
        cell.fill = BLUE_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")


def auto_fit_columns(ws):
    """
    Automatically adjusts the width of all columns in an Excel worksheet based on their content.

    Args:
        ws (openpyxl.worksheet.worksheet.Worksheet): The worksheet to adjust.
    """
    for col in ws.columns:
        max_length = 0;
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length: max_length = len(str(cell.value))
            except:
                pass
        ws.column_dimensions[column].width = (max_length + 2)


def set_cell_status_color(cell, level):
    """
    Sets the background color of an Excel cell based on the security level/status.

    Args:
        cell (openpyxl.cell.cell.Cell): The cell to colorize.
        level (str): The security level ('good', 'warning', 'bad', 'error').
    """
    if level == "good":
        cell.fill = GREEN_FILL
    elif level == "warning":
        cell.fill = ORANGE_FILL
    elif level == "error" or level == "bad":
        cell.fill = RED_FILL


# --- Aruba OS-CX Collection Functions ---
def get_aruba_device_info(net_connect):
    """
    Retrieves general device information (hostname, model, version, serial, uptime).

    Args:
        net_connect (netmiko.ConnectHandler): The established Netmiko connection object.

    Returns:
        dict: A dictionary containing device information.
    """
    info = {'hostname': 'N/A', 'ios_version': 'N/A', 'model': 'N/A', 'serial_number': 'N/A', 'uptime': 'N/A'}
    raw_sys_data = ""  # To store raw 'show system' output if needed
    try:
        prompt_hostname = net_connect.base_prompt
        if prompt_hostname:
            cleaned_prompt = re.sub(r"[#>()\s]+$", "", prompt_hostname)
            cleaned_prompt = cleaned_prompt.splitlines()[-1]
            if cleaned_prompt: info['hostname'] = cleaned_prompt


        # 1. Attempt 'show system' with TextFSM
        system_out_textfsm = None
        try:
            system_out_textfsm = net_connect.send_command("show system", use_textfsm=True, expect_string=r"#")
            if system_out_textfsm and isinstance(system_out_textfsm, list) and system_out_textfsm[0]:
                sys_data_textfsm = system_out_textfsm[0]
                info['hostname'] = sys_data_textfsm.get('hostname', info['hostname'])
                info['model'] = sys_data_textfsm.get('product_name', sys_data_textfsm.get('model', 'N/A'))
                info['serial_number'] = sys_data_textfsm.get('chassis_serial_nbr',
                                                             sys_data_textfsm.get('serial_number', 'N/A'))
                if sys_data_textfsm.get('up_time') and sys_data_textfsm.get('up_time', 'N/A') != 'N/A':
                    info['uptime'] = sys_data_textfsm['up_time'].strip()
        except (TextFSMError, ValueError, IndexError) as e:
            pass

        # 2. If info is missing after TextFSM for 'show system', use raw parsing
        if info['model'] == 'N/A' or info['serial_number'] == 'N/A' or info['uptime'] == 'N/A':
            if not raw_sys_data:  # Retrieve raw output if not already done
                raw_sys_data = net_connect.send_command("show system", use_textfsm=False, expect_string=r"#")
            # logger.debug(f"DEBUG: Host {net_connect.host} - raw 'show system' (for fallback info): \n{raw_sys_data}\n--------------------")

            if info['model'] == 'N/A':
                match_model_sys = re.search(r"Product Name\s*:\s*([^\r\n]+)", raw_sys_data, re.IGNORECASE)
                if match_model_sys: info['model'] = match_model_sys.group(1).strip()

            if info['serial_number'] == 'N/A':
                match_serial_sys = re.search(r"Chassis Serial Nbr\s*:\s*(\S+)", raw_sys_data, re.IGNORECASE)
                if match_serial_sys: info['serial_number'] = match_serial_sys.group(1).strip()

            if info['uptime'] == 'N/A':  # Specific uptime from 'show system'
                match_uptime_sys = re.search(r"Up Time\s*:\s*(.+)", raw_sys_data, re.IGNORECASE)
                if match_uptime_sys: info['uptime'] = match_uptime_sys.group(1).strip()

        # 3. Get OS version from 'show version' (more specific for version)
        version_out_textfsm = None
        try:
            version_out_textfsm = net_connect.send_command("show version", use_textfsm=True, expect_string=r"#")
            if version_out_textfsm and isinstance(version_out_textfsm, list) and version_out_textfsm[0]:
                ver_data = version_out_textfsm[0]
                if info['hostname'] == 'N/A' and ver_data.get('hostname'): info['hostname'] = ver_data.get('hostname')
                info['ios_version'] = ver_data.get('version',
                                                   ver_data.get('software_version',
                                                                ver_data.get('os_version',
                                                                             ver_data.get('arubaos_cx_version',
                                                                                          'N/A'))))
            else:
                raise ValueError("TextFSM for 'show version' did not return valid data.")
        except (TextFSMError, ValueError, IndexError) as e:
            logger.debug(f"TextFSM/Parsing 'show version' on {net_connect.host} failed: {e}. Raw mode.")
            raw_ver_data = net_connect.send_command("show version", use_textfsm=False, expect_string=r"#")
            if info['hostname'] == 'N/A':
                match_hostname_ver = re.search(r"Hostname\s*:\s*(\S+)", raw_ver_data, re.IGNORECASE)
                if match_hostname_ver: info['hostname'] = match_hostname_ver.group(1)
            match_version_ver = re.search(r"(?:ArubaOS-CX Version|Software version|Version)\s*:\s*(\S+)", raw_ver_data,
                                          re.IGNORECASE)
            if match_version_ver: info['ios_version'] = match_version_ver.group(1)

        # 4. If uptime still not found, try 'show uptime' as last resort
        if info.get('uptime', 'N/A') == 'N/A':
            uptime_out_raw = net_connect.send_command("show uptime", expect_string=r"#", use_textfsm=False)
            match_uptime_cmd = re.search(r"(?:System Uptime|Uptime is)\s*:\s*(.+)", uptime_out_raw, re.IGNORECASE)
            if match_uptime_cmd: info['uptime'] = match_uptime_cmd.group(1).strip()

        return info
    except Exception as e:
        logger.critical(f"Critical error get_aruba_device_info: {e}")
        traceback.print_exc()
        return info  # Return what has been collected so far


def get_aruba_interfaces(net_connect):
    """
    Collecting detailed information about device interfaces (status, IP, description, etc.).

    Args:
        net_connect (netmiko.ConnectHandler): The established Netmiko connection object.

    Returns:
        list: A list of dictionaries, where each dictionary represents an interface.
    """
    interfaces = []
    logger.info(f"Collecting interface information for {net_connect.host}...")
    try:
        # 1. Get IP information
        ip_brief_textfsm_out = net_connect.send_command("show ip interface brief", use_textfsm=True, expect_string=r"#")
        ip_map = {}
        if isinstance(ip_brief_textfsm_out, list):
            for item in ip_brief_textfsm_out:
                ifname = item.get('interface', item.get('intf'))
                if ifname:
                    ip_addr = item.get('ip_address', item.get('ipaddr'))
                    protocol_l3 = item.get('protocol', 'N/A').lower()
                    status_l3 = item.get('status', 'N/A').lower()
                    entry = {"protocol_l3": protocol_l3, "status_l3": status_l3}
                    if ip_addr and ip_addr != 'N/A' and ip_addr != 'unassigned':
                        entry["ip"] = ip_addr
                    else:
                        entry["ip"] = "unassigned"
                    ip_map[ifname] = entry

        # 2. Get L2/Physical Status
        int_status_list = []
        try:
            int_status_out = net_connect.send_command("show interface status", use_textfsm=True, expect_string=r"#")
            if isinstance(int_status_out, list): int_status_list = int_status_out
        except TextFSMError:
            logger.debug("TextFSM parsing failed for 'show interface status', continuing without it.")
        except Exception as e:
            logger.error(f"Command 'show interface status' failed on {net_connect.host}: {e}.")

        status_l2_map = {item.get('port', item.get('interface')): item for item in int_status_list if
                         item.get('port', item.get('interface'))}

        # 3. Get Descriptions from Config
        running_config_interfaces = net_connect.send_command("show running-config interfaces", read_timeout=120,
                                                             expect_string=r"#")
        desc_map = {}
        current_if_desc = None
        for line in running_config_interfaces.splitlines():
            line_strip = line.strip()
            if line_strip.startswith("interface "):
                current_if_desc = line_strip.split()[-1]
            elif "description " in line_strip and current_if_desc:
                desc_map[current_if_desc] = line_strip.split("description ", 1)[1]
            elif not line_strip or line_strip.startswith("!"):
                current_if_desc = None

        # 4. Main Interface Loop via 'show interface brief' (Raw Parse)
        show_int_brief_raw = net_connect.send_command("show interface brief", use_textfsm=False, expect_string=r"#")

        # Regex to parse 'show interface brief'
        # Handles various column layouts loosely
        int_brief_re = re.compile(
            r"^(?P<interface>\S+)\s+"
            r"(?P<native_vlan>\S+)\s+"
            r"(?P<mode>\S+)\s+"
            r"(?P<type_col>\S+)\s+"
            r"(?P<enabled>yes|no)\s+"
            r"(?P<link_status_l2>\S+)\s+"
            r"(?:(?P<reason>.*?)\s{2,})?" 
            r"(?P<speed>\S+)\s+"
            r"(?P<description>.*)$"
        )

        interfaces_from_regex_count = 0
        header_skipped_brief = False
        processed_interfaces_in_brief = set()

        for line in show_int_brief_raw.splitlines():
            line_s = line.strip()
            if not line_s: continue
            # Skip headers
            if line_s.lower().startswith("port ") or \
                    line_s.lower().startswith("native") or \
                    line_s.lower().startswith("-----"):
                header_skipped_brief = True;
                continue
            if not header_skipped_brief: continue

            match = int_brief_re.match(line_s)
            
            if match:
                interfaces_from_regex_count += 1
                data = match.groupdict()
                name = data['interface']
                processed_interfaces_in_brief.add(name)

                ip_info_dict = ip_map.get(name, {})
                ip = ip_info_dict.get("ip", "N/A")
                final_proto_status = ip_info_dict.get("protocol_l3", "N/A")

                status_l2_detail = status_l2_map.get(name, {})

                admin_enabled = data.get('enabled', 'no').lower()
                link_s_l2 = data.get('link_status_l2', 'N/A').lower()
                reason = data.get('reason', '').strip().lower() if data.get('reason') else ""

                final_link_status = link_s_l2
                if admin_enabled == 'no':
                    final_link_status = "administratively down"
                elif reason == "administratively down":
                    final_link_status = "administratively down"
                elif reason == "no xcvr installed":
                    final_link_status = "down (no transceiver)"

                link_s_from_status_cmd = status_l2_detail.get('status', '').lower()
                if link_s_from_status_cmd: final_link_status = link_s_from_status_cmd

                status_l3_from_ip = ip_info_dict.get("status_l3", "N/A")
                if final_proto_status == 'n/a' or final_proto_status == '':
                    if status_l3_from_ip != 'n/a' and status_l3_from_ip != '':
                        final_proto_status = status_l3_from_ip
                    elif final_link_status == "up":
                        final_proto_status = "up"
                    elif final_link_status == "administratively down":
                        final_proto_status = "down"
                    elif final_link_status.startswith("down"):
                        final_proto_status = "down"

                intf_type_parsed = "Management" if name.lower() == "mgmt" else \
                    "Virtual" if name.lower().startswith(("vlan", "loopback", "lag", "tunnel")) else \
                        "Physical"

                vlan_info = status_l2_detail.get('vlan', 'N/A')
                if data.get('native_vlan') and data.get('native_vlan') != '--':
                    vlan_info = data.get('native_vlan')
                if name.lower().startswith("vlan") and name[4:].isdigit():
                    vlan_info = name[4:]

                interfaces.append({
                    "name": name, "ip_address": ip,
                    "status_link": final_link_status, "status_protocol": final_proto_status,
                    "description": desc_map.get(name, data.get('description', "N/A").strip()),
                    "type": intf_type_parsed,
                    "vlan": vlan_info,
                    "duplex": status_l2_detail.get('duplex', 'N/A'),
                    "speed": status_l2_detail.get('speed', data.get('speed', 'N/A')),
                })

        # Specific processing for mgmt interface if not found
        if "mgmt" not in processed_interfaces_in_brief:
            try:
                mgmt_raw = net_connect.send_command("show interface mgmt", use_textfsm=False, expect_string=r"#")
                
                mgmt_ip, mgmt_link, mgmt_admin, mgmt_proto = "N/A", "N/A", "N/A", "N/A"

                match_ip = re.search(r"IPv4 address/subnet-mask\s*:\s*(\S+)", mgmt_raw, re.IGNORECASE)
                if match_ip: mgmt_ip = match_ip.group(1)

                match_admin = re.search(r"Admin State\s*:\s*(\S+)", mgmt_raw, re.IGNORECASE)
                if match_admin: mgmt_admin = match_admin.group(1).lower()

                match_link = re.search(r"Link State\s*:\s*(\S+)", mgmt_raw, re.IGNORECASE)
                if match_link: mgmt_link = match_link.group(1).lower()

                if mgmt_admin == "down":
                    final_mgmt_link_status = "administratively down"
                else:
                    final_mgmt_link_status = mgmt_link

                if final_mgmt_link_status == "up":
                    mgmt_proto = "up"
                elif final_mgmt_link_status == "administratively down":
                    mgmt_proto = "down"
                elif final_mgmt_link_status == "down":
                    mgmt_proto = "down"

                # If mgmt interface has an IP via 'show ip interface brief', it takes precedence
                mgmt_ip_from_map = ip_map.get("mgmt", {}).get("ip", mgmt_ip)
                mgmt_proto_from_map = ip_map.get("mgmt", {}).get("protocol_l3", mgmt_proto)
                if mgmt_proto_from_map != 'N/A': mgmt_proto = mgmt_proto_from_map

                if mgmt_link != "N/A":  
                    interfaces.append({
                        "name": "mgmt", "ip_address": mgmt_ip_from_map,
                        "status_link": final_mgmt_link_status, "status_protocol": mgmt_proto,
                        "description": desc_map.get("mgmt", "Management Interface"), "type": "Management",
                        "vlan": "N/A", "duplex": "N/A", "speed": "N/A",
                    })
                    interfaces_from_regex_count += 1
            except Exception as e_mgmt:
                logger.warning(f"Error recovering 'show interface mgmt' on {net_connect.host}: {e_mgmt}")

        logger.info(f"Interface parsing processed/found {interfaces_from_regex_count} entries for {net_connect.host}.")
        return interfaces
    except Exception as e:
        logger.critical(f"Critical error in get_aruba_interfaces for {net_connect.host}: {e}")
        traceback.print_exc()
        return []


def get_vlans(net_connect):
    """
    Retrieves the list of VLANs and their status.

    Args:
        net_connect (netmiko.ConnectHandler): The established Netmiko connection object.

    Returns:
        list: A list of dictionaries representing VLANs.
    """
    vlans_list = []
    try:
        vlan_out_textfsm = None
        try:
            vlan_out_textfsm = net_connect.send_command("show vlan", use_textfsm=True, expect_string=r"#")
        except TextFSMError:
            logger.debug(f"TextFSM failed for 'show vlan' on {net_connect.host}, continuing with regex.")
        except Exception as e:
            logger.error(f"Command 'show vlan' (TextFSM) failed on {net_connect.host}: {e}.")

        vlan_data_textfsm = vlan_out_textfsm if isinstance(vlan_out_textfsm, list) else []

        if vlan_data_textfsm:
            for v_entry in vlan_data_textfsm:
                vlans_list.append({
                    "id": v_entry.get('vlan_id', v_entry.get('id', 'N/A')),
                    "name": v_entry.get('name', v_entry.get('vlan_name', 'N/A')),
                    "status": v_entry.get('status', 'N/A'),
                    "ports": ", ".join(v_entry.get('ports', [])) if v_entry.get('ports') and isinstance(
                        v_entry.get('ports'), list) else v_entry.get('ports', 'N/A')
                })
        else:
            raw_vlan_out = net_connect.send_command("show vlan", use_textfsm=False, expect_string=r"#")
            header_skipped = False
            for line in raw_vlan_out.splitlines():
                line_s = line.strip()
                if not line_s: continue
                if line_s.startswith("----") or line_s.lower().startswith("vlan name") or line_s.lower().startswith(
                        "vlan id name"):
                    header_skipped = True;
                    continue
                if not header_skipped: continue
                match_vlan = re.match(r"^\s*(\d+)\s+([\w\-\/\.]+)\s+(\S+)(?:\s+\S*){2}\s*(.*)", line_s)
                if match_vlan:
                    vid, vname, vstatus, vports_str = match_vlan.groups()
                    vname_clean = vname if not vname.startswith("<NO") else "N/A"
                    vports_clean = vports_str.strip() if vports_str.strip() and not vports_str.startswith(
                        "<NO") else "N/A"
                    vlans_list.append({"id": vid, "name": vname_clean, "status": vstatus, "ports": vports_clean})
        return vlans_list
    except Exception as e:
        logger.critical(f"Critical error in get_vlans for {net_connect.host}: {e}")
        traceback.print_exc()
        return []


def get_arp_table(net_connect):
    """
    Retrieves the ARP table from the device.

    Args:
        net_connect (netmiko.ConnectHandler): The established Netmiko connection object.

    Returns:
        list: A list of dictionaries representing ARP entries.
    """
    arp_table = []
    try:
        arp_out_textfsm = None
        try:
            arp_out_textfsm = net_connect.send_command("show arp", use_textfsm=True, expect_string=r"#")
        except TextFSMError:
            logger.debug(f"TextFSM failed for 'show arp' on {net_connect.host}, continuing with regex.")
        except Exception as e:
            logger.error(f"Command 'show arp' (TextFSM) failed on {net_connect.host}: {e}.")

        arp_data_textfsm = arp_out_textfsm if isinstance(arp_out_textfsm, list) else []

        if arp_data_textfsm:
            for entry in arp_data_textfsm:
                arp_table.append({
                    "protocol": entry.get('protocol', 'Internet'),
                    "address": entry.get('address', entry.get('ip_address', 'N/A')),
                    "age": entry.get('age', 'N/A'),
                    "mac_address": entry.get('mac', entry.get('mac_address', 'N/A')).replace(':', '').replace('-',
                                                                                                              '').replace(
                        '.', ''),
                    "type": entry.get('type', 'ARPA'),
                    "interface": entry.get('interface', 'N/A')})
        else:
            raw_arp_out = net_connect.send_command("show arp", use_textfsm=False, expect_string=r"#")
            if "No ARP entries found" in raw_arp_out or not raw_arp_out.strip() or \
                    ("Total ARP Entries" in raw_arp_out and "0" in raw_arp_out.split("Total ARP Entries")[-1]):
                return []

            header_skipped_arp = False
            for line in raw_arp_out.splitlines():
                line_s = line.strip()
                if not line_s or line_s.lower().startswith("ip address") or \
                        line_s.lower().startswith("total arp") or line_s.lower().startswith("arp entries"):
                    header_skipped_arp = True;
                    continue
                if not header_skipped_arp and not re.match(r"^\s*[\d\.]+", line_s): continue

                match_arp = re.match(r"^\s*([\d\.]+)\s+([0-9a-f\.\:\-]+)\s+(vlan\d+|\S+)\s+.*", line_s, re.IGNORECASE)
                if match_arp:
                    ip, mac, intf_arp = match_arp.groups()
                    arp_table.append({
                        "protocol": "Internet", "address": ip, "age": "N/A",
                        "mac_address": mac.replace(':', '').replace('-', '').replace('.', ''), "type": "ARPA",
                        "interface": intf_arp
                    })
        return arp_table
    except Exception as e:
        logger.critical(f"Critical error in get_arp_table for {net_connect.host}: {e}")
        traceback.print_exc()
        return []


def check_security_features(net_connect, running_config):
    """
    Performs a comprehensive security audit of the Aruba OS-CX device configuration.

    Args:
        net_connect (netmiko.ConnectHandler): The established Netmiko connection object.
        running_config (str): The full running configuration of the device.

    Returns:
        dict: A dictionary containing the results of various security checks.
    """
    security_audit = {}
    logger.info(f"Starting security audit for {net_connect.host}...")

    # --- I. AAA & Authentication & Management Access ---
    # 1. AAA Port Access (Edge Security)
    if "aaa authentication port-access" in running_config:
        security_audit["aaa_port_access_configured"] = {"status": True, "level": "good",
                                                        "details": "AAA for port access (dot1x/mac-auth) seems configured."}
    else:
        security_audit["aaa_port_access_configured"] = {"status": False, "level": "warning",
                                                        "details": "AAA for port access (dot1x/mac-auth) not detected. Recommended to secure network access."}

    # 2. AAA Login (Management Security) - CRITICAL MISSING CHECK ADDED
    if "aaa authentication login" in running_config:
        if "group tacacs" in running_config or "group radius" in running_config:
            security_audit["aaa_login_configured"] = {"status": "Centralized (TACACS+/RADIUS)", "level": "good",
                                                      "details": "AAA login authentication uses centralized server."}
        else:
            security_audit["aaa_login_configured"] = {"status": "Local/Other", "level": "warning",
                                                      "details": "AAA login configured but might be local only. Verify 'aaa authentication login' config."}
    else:
        security_audit["aaa_login_configured"] = {"status": "Not Configured", "level": "bad",
                                                  "details": "No 'aaa authentication login' found. Default local auth used?"}

    # 3. Local User Password Type
    local_user_passwords_encrypted = True
    # Check for plaintext passwords explicitly
    if re.search(r"user\s+\S+\s+password\s+plaintext", running_config):
        local_user_passwords_encrypted = False

    if local_user_passwords_encrypted and "password" in running_config:
        security_audit["local_user_password_encryption"] = {"status": "Encrypted (hashed)", "level": "good",
                                                            "details": "Local user passwords seem to be stored encrypted (hashed)."}
    elif not local_user_passwords_encrypted:
        security_audit["local_user_password_encryption"] = {"status": "Plaintext detected", "level": "bad",
                                                            "details": "At least one local user password is stored in plaintext. Use 'password ciphertext <hash>'."}
    else:
        # Case where no "password" keyword found or ambiguous
        security_audit["local_user_password_encryption"] = {
            "status": "No local user with password found", "level": "warning",
            "details": "Manually check local user configuration."}

    # Password complexity policy
    min_len_str, complexity_str = "Not configured", "Not configured"
    min_len_level, complexity_level = "bad", "bad"

    match_pass_len = re.search(r"password minimum-length\s+(\d+)", running_config)
    if match_pass_len:
        min_len = int(match_pass_len.group(1))
        min_len_str = f"Min length: {min_len}"
        min_len_level = "good" if min_len >= 12 else "warning" if min_len >= 8 else "bad"
    
    security_audit["password_min_length"] = {"status": min_len_str, "level": min_len_level,
                                             "details": f"{min_len_str}. Recommended: >=12."}

    if "password complexity" in running_config or "character-class-check" in running_config:
        complexity_str = "Enabled (check details)"
        complexity_level = "good"
    security_audit["password_complexity"] = {"status": complexity_str, "level": complexity_level,
                                             "details": f"Password complexity policy: {complexity_str}."}

    # --- II. Access Line Security (Console) ---
    console_config_text = ""
    console_match = re.search(r"line console\s*\n(.*?)(?=line|interface|vlan|router|exit|$)", running_config,
                              re.DOTALL | re.MULTILINE)
    if console_match:
        console_config_text = console_match.group(1)

    if console_config_text and ("password " in console_config_text or "login " in console_config_text or "aaa authentication login" in running_config):
         # Logic assumption: if global AAA login is set, console might use it imply default. 
         # But safer to look for specific line config or inheritance.
         # For Audit purposes, if "login" is there it's usually good.
        security_audit["console_auth"] = {"status": True, "level": "good",
                                          "details": "Console line protected by login method."}
    else:
        security_audit["console_auth"] = {"status": False, "level": "bad",
                                          "details": "Console line authentication not explicitly seen in 'line console'."}

    exec_timeout_con_match = re.search(r"session-timeout\s+(\d+)", console_config_text)
    if exec_timeout_con_match:
        minutes_con = int(exec_timeout_con_match.group(1))
        if 0 < minutes_con <= 15:
            security_audit["console_session_timeout"] = {"status": f"{minutes_con} minutes", "level": "good",
                                                         "details": "Session timeout configured on console."}
        elif minutes_con == 0:
            security_audit["console_session_timeout"] = {"status": "Disabled (0)", "level": "bad",
                                                         "details": "Console session timeout disabled. Risk."}
        else:
            security_audit["console_session_timeout"] = {"status": f"{minutes_con} minutes", "level": "warning",
                                                         "details": "Console session timeout high. Recommended: 5-15 min."}
    else:
        security_audit["console_session_timeout"] = {"status": "Not configured (default)", "level": "warning",
                                                     "details": "No session timeout on console. Recommended: 5-15 min."}

    # --- III. Management Services Security ---
    # SSH
    try:
        ssh_status_raw = net_connect.send_command("show ssh server all-vrfs", expect_string=r"#")
        
        ssh_enabled = False
        ssh_v1 = False
        
        if "SSH server configuration on VRF" in ssh_status_raw:
            ssh_enabled = True # Server is running
        
        # Check specific version disablement in config
        # "no ssh server v1 enable" -> Good
        # "ssh server v1 enable" -> Bad
        
        if "ssh server v1 enable" in running_config:
            ssh_v1 = True
        elif "no ssh server v1 enable" in running_config:
            ssh_v1 = False
        else:
            # Default behavior of OS-CX? Old versions enabled, newer disabled. 
            # We mark as warning if not explicitly disabled.
            ssh_v1 = "Unknown (Implicit)"

        if ssh_enabled:
            if ssh_v1 is True:
                 security_audit["ssh_status"] = {"status": "SSHv1 Enabled", "level": "bad", "details": "SSHv1 explicitly enabled. Disable it."}
            elif ssh_v1 is False:
                 security_audit["ssh_status"] = {"status": "SSHv2 Only", "level": "good", "details": "SSHv2 enabled, v1 disabled."}
            else:
                 security_audit["ssh_status"] = {"status": "SSH Enabled (v1 implicit)", "level": "warning", "details": "Verify if SSHv1 is disabled by default or add 'no ssh server v1 enable'."}
        else:
             # Weird if we are connected via SSH...
             security_audit["ssh_status"] = {"status": "Disabled?", "level": "warning", "details": "SSH server appears disabled in 'show ssh', yet we are connected?"}

    except Exception as e:
        logger.error(f"SSH check failed: {e}")
        security_audit["ssh_status"] = {"status": "Error", "level": "error", "details": "Failed to check SSH status."}

    # Telnet
    if "no telnet-server enable" in running_config:
        security_audit["telnet_server"] = {"status": "Disabled", "level": "good", "details": "Telnet server disabled."}
    elif "telnet-server enable" in running_config:
        security_audit["telnet_server"] = {"status": "Enabled", "level": "bad", "details": "Telnet server enabled. Disable it."}
    else:
        # Default is usually disabled on modern AOS-CX, but good verify.
        security_audit["telnet_server"] = {"status": "Default", "level": "warning", "details": "Explicit 'no telnet-server enable' recommended."}
        
    # TFTP - New Check
    if "tftp-server enable" in running_config:
        security_audit["tftp_server"] = {"status": "Enabled", "level": "bad", "details": "TFTP server enabled. Insecure."}
    else:
        security_audit["tftp_server"] = {"status": "Disabled", "level": "good", "details": "TFTP server not enabled."}

    # HTTP/HTTPS
    https_enabled = "https-server enable" in running_config or "https-server vrf" in running_config
    http_enabled = "http-server enable" in running_config or "http-server vrf" in running_config
    
    if https_enabled:
        security_audit["https_server"] = {"status": "Enabled", "level": "good", "details": "HTTPS server (web-mgmt) enabled."}
    else:
        security_audit["https_server"] = {"status": "Disabled", "level": "good", "details": "HTTPS server disabled."}
        
    if http_enabled:
        security_audit["http_server"] = {"status": "Enabled", "level": "bad", "details": "HTTP server enabled. Insecure."}
    else:
         security_audit["http_server"] = {"status": "Disabled", "level": "good", "details": "HTTP server disabled."}

    # Banners
    banners_set = []
    if "banner motd" in running_config: banners_set.append("MOTD")
    if "banner exec" in running_config: banners_set.append("Exec")
    
    if banners_set:
        security_audit["banners_configured"] = {"status": f"Configured: {', '.join(banners_set)}", "level": "good", "details": "Warning banners configured."}
    else:
        security_audit["banners_configured"] = {"status": "None", "level": "warning", "details": "No warning banners configured."}

    # --- IV. General Hardening ---
    # IP Source Route - Fixed Logic
    # On many Aruba CX, "no ip source-route" might not be a command or might be default.
    # We check if "ip source-route" IS present.
    if "ip source-route" in running_config and "no ip source-route" not in running_config:
         security_audit["ip_source_route"] = {"status": "Enabled", "level": "bad", "details": "IP Source Routing enabled."}
    else:
         security_audit["ip_source_route"] = {"status": "Disabled", "level": "good", "details": "IP Source Routing not found (disabled)."}

    # LLDP
    if "no lldp enable" in running_config:
        security_audit["lldp_global_status"] = {"status": "Globally Disabled", "level": "good", "details": "LLDP disabled."}
    else:
        security_audit["lldp_global_status"] = {"status": "Globally Enabled", "level": "warning", "details": "LLDP globally enabled. Secure untrusted ports."}

    # --- V. Logging & Monitoring ---
    if "logging syslog host" in running_config or "logging host" in running_config:
        security_audit["remote_logging"] = {"status": True, "level": "good", "details": "Remote logging (syslog) configured."}
    else:
        security_audit["remote_logging"] = {"status": False, "level": "bad", "details": "Remote logging (syslog) NOT configured."}

    # NTP
    ntp_lines = [line for line in running_config.splitlines() if line.strip().startswith("ntp server")]
    if len(ntp_lines) >= 2:
        security_audit["ntp_redundancy"] = {"status": f"{len(ntp_lines)} servers", "level": "good", "details": "NTP redundancy OK."}
    elif len(ntp_lines) == 1:
        security_audit["ntp_redundancy"] = {"status": "1 server", "level": "warning", "details": "Only one NTP server configured."}
    else:
        security_audit["ntp_redundancy"] = {"status": "None", "level": "bad", "details": "No NTP servers configured."}
        
    # NTP Auth (Extra)
    if any(" key " in line for line in ntp_lines):
         security_audit["ntp_auth"] = {"status": "Configured", "level": "good", "details": "NTP authentication seems used."}
    else:
         security_audit["ntp_auth"] = {"status": "Not Configured", "level": "warning", "details": "NTP authentication not detected."}

    # SNMP
    if "snmp-server community public" in running_config or "snmp-server community private" in running_config:
        security_audit["snmp_default_communities"] = {"status": "Found", "level": "bad", "details": "Default SNMP communities (public/private) present."}
    else:
        security_audit["snmp_default_communities"] = {"status": "Not Found", "level": "good", "details": "No default public/private communities found."}

    snmpv3_configured = "snmp-server user" in running_config
    if snmpv3_configured:
        security_audit["snmp_version"] = {"status": "SNMPv3", "level": "good", "details": "SNMPv3 users configured."}
    elif "snmp-server community" in running_config:
        security_audit["snmp_version"] = {"status": "SNMPv1/v2c", "level": "warning", "details": "Only SNMPv1/v2c communities found. Prefer v3."}
    else:
        security_audit["snmp_version"] = {"status": "None", "level": "good", "details": "SNMP not configured."}

    # --- VI. Layer 2 Security ---
    if "bpdu-protection" in running_config:
        security_audit["bpdu_protection"] = {"status": "Active", "level": "good", "details": "BPDU protection active globally or on ports."}
    else:
        security_audit["bpdu_protection"] = {"status": "Inactive", "level": "warning", "details": "BPDU protection not found."}

    if "dhcp-snooping" in running_config:
         security_audit["dhcp_snooping"] = {"status": "Active", "level": "good", "details": "DHCP Snooping configured."}
    else:
         security_audit["dhcp_snooping"] = {"status": "Inactive", "level": "bad", "details": "DHCP Snooping not configured."}

    return security_audit


def load_inventory(filepath="inventory.csv"):
    """
    Loads device inventory from a CSV file.

    Args:
        filepath (str, optional): Path to the inventory CSV file. Defaults to "inventory.csv".

    Returns:
        list: A list of dictionaries, each representing a device in the inventory.
              Returns None on error or empty list if file is empty.
    """
    inventory = []
    try:
        with open(filepath, mode='r', encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            try:
                header = next(reader)
                if len(header) < 3:
                    logger.error(
                        f"Inventory header '{filepath}' must have at least 3 columns (hostname, group, device_type).")
                    return None
            except StopIteration:
                logger.warning(f"Inventory file '{filepath}' is empty.")
                return []
            for row in reader:
                if len(row) >= 3 and row[0].strip():
                    inventory.append(
                        {"host": row[0].strip(), "group": row[1].strip(), "device_type": row[2].strip().lower()})
                elif row and any(field.strip() for field in row):
                    logger.warning(f"Malformed inventory row: {row}")
        return inventory
    except FileNotFoundError:
        logger.error(f"Inventory file '{filepath}' not found.")
        return None
    except Exception as e:
        logger.error(f"Error reading '{filepath}': {e}")
        traceback.print_exc()
        return None


def load_passwords(filepath="passwords.csv"):
    """
    Loads credentials from a CSV file mapping groups to usernames and passwords.

    Args:
        filepath (str, optional): Path to the passwords CSV file. Defaults to "passwords.csv".

    Returns:
        dict: A dictionary mapping group names to credential dictionaries (username, password, enable_password).
              Returns None on error or empty dict if file is empty.
    """
    passwords = {}
    try:
        with open(filepath, mode='r', encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            try:
                next(reader)
            except StopIteration:
                logger.warning(f"Passwords file '{filepath}' is empty.")
                return {}
            for row in reader:
                if len(row) >= 3 and row[0].strip():
                    enable_pass = row[3].strip() if len(row) > 3 and row[3].strip() else None
                    passwords[row[0].strip()] = {"username": row[1].strip(), "password": row[2].strip(),
                                                 "enable_password": enable_pass}
                elif row and any(field.strip() for field in row):
                    logger.warning(f"Malformed passwords row: {row}")
        return passwords
    except FileNotFoundError:
        logger.error(f"Passwords file '{filepath}' not found.")
        return None
    except Exception as e:
        logger.error(f"Error reading '{filepath}': {e}")
        traceback.print_exc()
        return None


def generate_excel_report(all_data, excel_filepath):
    """
    Generates a comprehensive Excel report from collected audit data.

    Args:
        all_data (list): A list of dictionaries containing audit data for all devices.
        excel_filepath (str): The file path where the Excel report will be saved.
    """
    if not all_data:
        logger.warning(f"No data for Excel report: {excel_filepath}.")
        return
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws_info = wb.create_sheet("General Info")
    headers_info = ["Hostname", "IP Address", "Model", "OS Version", "Uptime", "Serial Number"]
    ws_info.append(headers_info)
    apply_header_style(ws_info)
    for dev_data in all_data:
        if dev_data.get('status') == 'error_connection':
            ws_info.append(
                [dev_data.get('attempted_host', 'N/A'), dev_data.get('attempted_host', 'N/A'), "CONNECTION ERROR",
                 dev_data.get('error_message', 'N/A'), "N/A", "N/A"])
            for i in range(3, 5): ws_info.cell(row=ws_info.max_row, column=i).fill = RED_FILL
        else:
            info = dev_data.get("general_info", {})
            ws_info.append([info.get("hostname", dev_data.get("host", "N/A")), dev_data.get("host", "N/A"),
                            info.get("model", "N/A"), info.get("ios_version", "N/A"), info.get("uptime", "N/A"),
                            info.get("serial_number", "N/A")])
    auto_fit_columns(ws_info)
    ws_interfaces = wb.create_sheet("Interfaces")
    headers_interfaces = ["Hostname", "Interface Name", "Type", "Description", "IP Address", "Link Status",
                          "Protocol Status", "VLAN (Access)", "Duplex", "Speed"]
    ws_interfaces.append(headers_interfaces)
    apply_header_style(ws_interfaces)
    for dev_data in all_data:
        if dev_data.get('status') == 'error_connection': continue
        hostname = dev_data.get("general_info", {}).get("hostname", dev_data.get("host", "N/A"))
        for iface in dev_data.get("interfaces", []):
            ws_interfaces.append(
                [hostname, iface.get("name"), iface.get("type"), iface.get("description"), iface.get("ip_address"),
                 iface.get("status_link"), iface.get("status_protocol"), iface.get("vlan"), iface.get("duplex"),
                 iface.get("speed")])
            lr, pr = ws_interfaces.cell(row=ws_interfaces.max_row, column=6), ws_interfaces.cell(
                row=ws_interfaces.max_row, column=7)
            sl, sp = str(iface.get("status_link", "")).lower(), str(iface.get("status_protocol", "")).lower()
            if sl == "up" or "connected" in sl:
                lr.fill = GREEN_FILL
            elif ("admin" in sl and "down" in sl) or sl == "disabled" or "admin-down" in sl:
                lr.fill = ORANGE_FILL
            elif sl == "down" or "notconnect" in sl or "err-disabled" in sl or "error-disabled" in sl or "down (no transceiver)" in sl:
                lr.fill = RED_FILL
            if sp == "up":
                pr.fill = GREEN_FILL
            elif sp == "down":
                pr.fill = RED_FILL
    auto_fit_columns(ws_interfaces)
    ws_vlans = wb.create_sheet("VLANs")
    headers_vlans = ["Hostname", "VLAN ID", "VLAN Name", "Status", "Assigned Ports"]
    ws_vlans.append(headers_vlans)
    apply_header_style(ws_vlans)
    for dev_data in all_data:
        if dev_data.get('status') == 'error_connection': continue
        hostname = dev_data.get("general_info", {}).get("hostname", dev_data.get("host", "N/A"))
        for vlan in dev_data.get("vlans", []):
            ws_vlans.append([hostname, vlan.get("id"), vlan.get("name"), vlan.get("status"), vlan.get("ports")])
            sc = ws_vlans.cell(row=ws_vlans.max_row, column=4)
            if str(vlan.get("status", "")).lower() == "up" or str(vlan.get("status", "")).lower() == "active":
                sc.fill = GREEN_FILL
            else:
                sc.fill = ORANGE_FILL
    auto_fit_columns(ws_vlans)
    ws_arp = wb.create_sheet("ARP Table")
    headers_arp = ["Hostname", "Protocol", "IP Address", "Age (min)", "MAC Address", "Type", "Interface"]
    ws_arp.append(headers_arp)
    apply_header_style(ws_arp)
    for dev_data in all_data:
        if dev_data.get('status') == 'error_connection': continue
        hostname = dev_data.get("general_info", {}).get("hostname", dev_data.get("host", "N/A"))
        for entry in dev_data.get("arp_table", []):
            ws_arp.append(
                [hostname, entry.get("protocol"), entry.get("address"), entry.get("age"), entry.get("mac_address"),
                 entry.get("type"), entry.get("interface")])
    auto_fit_columns(ws_arp)
    ws_security = wb.create_sheet("Security Audit")
    headers_security = ["Hostname", "Check Point", "Status/Value", "Level", "Details/Recommendation"]
    ws_security.append(headers_security)
    apply_header_style(ws_security)
    for dev_data in all_data:
        if dev_data.get('status') == 'error_connection': continue
        hostname = dev_data.get("general_info", {}).get("hostname", dev_data.get("host", "N/A"))
        for check_name, check_data in dev_data.get("security_audit", {}).items():
            ws_security.append(
                [hostname, check_name.replace("_", " ").title(), str(check_data.get("status")), check_data.get("level"),
                 check_data.get("details")])
            lc = ws_security.cell(row=ws_security.max_row, column=4)
            set_cell_status_color(lc, check_data.get("level"))
    auto_fit_columns(ws_security)
    try:
        wb.save(excel_filepath)
        logger.info(f"[+] Aruba Excel report generated: {excel_filepath}")
    except Exception as e:
        logger.error(f"[-] Error saving Aruba Excel report: {e}")


def perform_aruba_audit(aruba_devices_inventory, global_passwords_map, output_directory):
    """
    Orchestrates the audit process for a list of Aruba devices.
    Connects to each device, collects data, performs security checks, and saves results.

    Args:
        aruba_devices_inventory (list): List of dictionaries representing Aruba devices to audit.
        global_passwords_map (dict): Dictionary mapping groups to credentials.
        output_directory (str): Directory to save JSON data and Excel reports.
    """
    all_devices_data = []
    for device_entry in aruba_devices_inventory:
        host, group = device_entry["host"], device_entry["group"]
        creds = global_passwords_map.get(group)
        logger.info(f"Processing {host} (group: {group})...")

        if not creds:
            logger.error(f"Credentials not found for group '{group}'. {host} ignored.")
            all_devices_data.append({"attempted_host": host, "status": "error_connection",
                                     "error_message": f"Credentials not found for group {group}"})
            continue

        device_params = {
            'device_type': 'aruba_aoscx_ssh',
            'host': host, 'username': creds['username'], 'password': creds['password'],
            'secret': creds.get('enable_password'),
            'global_delay_factor': 2, 'timeout': 45, 'session_timeout': 120
        }
        current_device_data = {"host": host}
        try:
            with ConnectHandler(**device_params) as net_connect:
                actual_host = net_connect.host
                actual_prompt = (
                    net_connect.base_prompt[:-1] if net_connect.base_prompt and net_connect.base_prompt.endswith(
                        ('#', '>')) else actual_host)
                logger.info(f"Connected to {actual_host} ({actual_prompt}).")

                current_device_data["general_info"] = get_aruba_device_info(net_connect)
                current_device_data["general_info"]["ip_address_queried"] = host

                running_config = net_connect.send_command("show running-config", read_timeout=240, expect_string=r"#")
                if not running_config: running_config = ""

                current_device_data["interfaces"] = get_aruba_interfaces(net_connect)
                current_device_data["vlans"] = get_vlans(net_connect)
                current_device_data["arp_table"] = get_arp_table(net_connect)
                current_device_data["security_audit"] = check_security_features(net_connect, running_config)
                all_devices_data.append(current_device_data)

        except (NetmikoTimeoutException, SSHException) as e:
            logger.error(f"Connection to {host} (Timeout/SSH): {e}")
            all_devices_data.append({"attempted_host": host, "status": "error_connection", "error_message": str(e)})
        except NetmikoAuthenticationException as e:
            logger.error(f"Authentication on {host}: {e}")
            all_devices_data.append(
                {"attempted_host": host, "status": "error_connection", "error_message": f"Authentication failed: {e}"})
        except Exception as e:
            if "Unsupported 'device_type'" in str(e):
                logger.error(
                    f"Unsupported 'device_type' for {host}: {e}. Check 'device_type' ('aruba_aoscx_ssh') or Netmiko version.")
                all_devices_data.append({"attempted_host": host, "status": "error_connection",
                                         "error_message": f"Unsupported 'device_type': {e}"})
            else:
                logger.error(f"Unexpected error with {host}: {e}")
                traceback.print_exc()
                all_devices_data.append(
                    {"attempted_host": host, "status": "error_connection", "error_message": f"Unexpected error: {e}"})

    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    json_filename = os.path.join(output_directory, f"audit_aruba_data_{timestamp}.json")
    try:
        with open(json_filename, 'w', encoding='utf-8') as f:
            json.dump(all_devices_data, f, indent=4, ensure_ascii=False)
        logger.info(f"[+] Aruba JSON data saved: {json_filename}")
    except Exception as e:
        logger.error(f"[-] Error saving Aruba JSON data: {e}")

    excel_report_path = os.path.join(output_directory, f"audit_aruba_report_{timestamp}.xlsx")
    generate_excel_report(all_devices_data, excel_report_path)

    logger.info(f"Aruba audit completed for {len(aruba_devices_inventory)} device(s).")


def main_aruba():
    """
    Main entry point for standalone execution of the Aruba audit script.
    Loads inventory, credentials, filter for Aruba devices, and starts the audit.
    """
    # Basic logging setup for standalone run
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')

    inventory_file, password_file, output_directory = "inventory.csv", "passwords.csv", "audit_reports"
    if not os.path.exists(output_directory):
        try:
            os.makedirs(output_directory)
        except OSError as e:
            logger.error(f"Error creating directory '{output_directory}': {e}")
            return

    full_inventory = load_inventory(inventory_file)
    passwords_map = load_passwords(password_file)

    if full_inventory is None or passwords_map is None:
        logger.error("Stop (Aruba): Critical errors loading input files.")
        return

    aruba_devices = [device for device in full_inventory if device.get("device_type") == "aruba_os-cx"]

    if not aruba_devices:
        logger.warning("No Aruba OS-CX devices found in inventory for standalone audit.")
        return

    perform_aruba_audit(aruba_devices, passwords_map, output_directory)


if __name__ == "__main__":
    main_aruba()
