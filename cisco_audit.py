#!/usr/bin/env python3

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
from netmiko import ConnectHandler
from netmiko.exceptions import NetmikoTimeoutException, NetmikoAuthenticationException, SSHException
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import traceback

# --- Logging Configuration ---
logger = logging.getLogger(__name__)
# --- Excel Configuration ---
GREEN_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
ORANGE_FILL = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
RED_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
BLUE_FILL = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
BOLD_FONT = Font(bold=True)


# --- Excel Helper Functions ---
def apply_header_style(ws, row_num=1):
    """
    Applies a standard bold, blue styling to the header row of an Excel worksheet.

    Args:
        ws (openpyxl.worksheet.worksheet.Worksheet): The worksheet to apply styling to.
        row_num (int, optional): The row number to style as header. Defaults to 1.
    """
    for cell in ws[row_num]:
        cell.font = BOLD_FONT
        cell.fill = BLUE_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")


def auto_fit_columns(ws):
    """
    Automatically adjusts the width of all columns in an Excel worksheet based on their content.

    Args:
        ws (openpyxl.worksheet.worksheet.Worksheet): The worksheet to adjust.
    """
    for col in ws.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length: max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)
        ws.column_dimensions[column].width = adjusted_width


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


# --- Information Collection Functions ---
def get_device_info(net_connect):
    """
    Retrieves general device information (hostname, model, version, serial, uptime).

    Args:
        net_connect (netmiko.ConnectHandler): The established Netmiko connection object.

    Returns:
        dict: A dictionary containing device information.
    """
    info = {}
    try:
        output_raw = net_connect.send_command("show version", use_textfsm=True)
        output = output_raw if isinstance(output_raw, list) and output_raw else [{}]
        if output:
            dev_info = output[0]
            info['hostname'] = dev_info.get('hostname', 'N/A')
            info['ios_version'] = dev_info.get('version', 'N/A')
            hardware_info = dev_info.get('hardware', ['N/A'])
            info['model'] = hardware_info[0] if isinstance(hardware_info, list) and hardware_info else 'N/A'
            info['uptime'] = dev_info.get('uptime', 'N/A')
            serial_info = dev_info.get('serial', ['N/A'])
            info['serial_number'] = serial_info[0] if isinstance(serial_info, list) and serial_info else 'N/A'
        if info.get('hostname', 'N/A') == 'N/A' or not info.get('hostname'):
            prompt = net_connect.base_prompt
            if prompt: info['hostname'] = prompt.strip("#> ")
        return info
    except Exception as e:
        logger.error(f"Error get_device_info for {net_connect.host}: {e}")
        return {'hostname': 'Error', 'ios_version': 'Error', 'model': 'Error', 'uptime': 'Error',
                'serial_number': 'Error'}


def get_interfaces(net_connect):
    """
    Collecting detailed information about device interfaces (status, IP, description, etc.).

    Args:
        net_connect (netmiko.ConnectHandler): The established Netmiko connection object.

    Returns:
        list: A list of dictionaries, where each dictionary represents an interface.
    """
    interfaces = []
    try:
        ip_interfaces_raw = net_connect.send_command("show ip interface brief", use_textfsm=True)
        ip_interfaces_list = ip_interfaces_raw if isinstance(ip_interfaces_raw, list) else []


        if not ip_interfaces_list:
            raw_output_ip_brief = net_connect.send_command("show ip interface brief", use_textfsm=False)
            ip_interfaces_list = []
            interface_regex = re.compile(
                r"^(?P<interface>\S+)\s+"
                r"(?P<ip_address>[\d\.]+|unassigned)\s+"
                r"\S+\s+\S+\s+"
                r"(?P<status>administratively down|up|down|[^ ]+)"
                r"\s+(?P<proto>up|down)$"
            )
            for line in raw_output_ip_brief.splitlines():
                line = line.strip()
                if not line or line.lower().startswith("interface"): continue
                match = interface_regex.match(line)
                if match:
                    data = match.groupdict()
                    if data['status'].lower() == "administratively" and data['proto'].lower() == "down":
                        data['status'] = "administratively down"
                    ip_interfaces_list.append(data)

        interface_status_raw = net_connect.send_command("show interfaces status", use_textfsm=True)
        interface_descriptions_raw = net_connect.send_command("show interfaces description", use_textfsm=True)
        interface_status_list = interface_status_raw if isinstance(interface_status_raw, list) else []
        interface_descriptions_list = interface_descriptions_raw if isinstance(interface_descriptions_raw, list) else []

        status_map = {item.get('port'): item for item in interface_status_list if item.get('port')}
        desc_map = {item.get('port'): str(item.get('descrip', item.get('description')))
                    for item in interface_descriptions_list
                    if item.get('port') and item.get('descrip', item.get('description')) and
                    str(item.get('descrip', item.get('description'))).strip() and
                    str(item.get('descrip', item.get('description'))) != '--'}

        for iface_data in ip_interfaces_list:
            name = iface_data.get('interface')
            if not name: continue

            protocol_status = iface_data.get('protocol', iface_data.get('proto', 'N/A')).lower()
            link_status_brief = iface_data.get('status', 'N/A').lower()
            ip_address_val = iface_data.get('ip_address', iface_data.get('ipaddr', 'unassigned'))

            phys_status_detail = status_map.get(name, {})
            current_link_status = link_status_brief

            if phys_status_detail:
                phys_status_val = phys_status_detail.get('status', '').lower()
                if phys_status_val:
                    if phys_status_val in ["connected", "disabled", "notconnect", "inactive",
                                           "monitoring"] or "err-disabled" in phys_status_val:
                        current_link_status = phys_status_val
                    if phys_status_val == "disabled": current_link_status = "administratively down"

            if "admin" in link_status_brief and "down" in link_status_brief:
                current_link_status = "administratively down"

            interface_type = "Virtual"
            name_lower = name.lower()
            physical_prefixes = ("eth", "gi", "fa", "te", "twe", "hu", "fo", "se")
            virtual_prefixes = ("vl", "lo", "tu", "po", "bv", "nu", "oo", "gr")
            if any(name_lower.startswith(p) for p in physical_prefixes):
                interface_type = "Physical"
            elif any(name_lower.startswith(p) for p in virtual_prefixes):
                interface_type = "Virtual"

            interfaces.append({
                "name": name, "ip_address": ip_address_val if ip_address_val != 'unassigned' else "N/A",
                "status_link": current_link_status, "status_protocol": protocol_status,
                "description": desc_map.get(name, "N/A"), "type": interface_type,
                "vlan": phys_status_detail.get('vlan', 'N/A'), "duplex": phys_status_detail.get('duplex', 'N/A'),
                "speed": phys_status_detail.get('speed', 'N/A'),
            })
        return interfaces
    except Exception as e:
        logger.critical(f"Critical error in get_interfaces for {net_connect.host}: {e}")
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
    vlans = []
    try:
        output_raw = net_connect.send_command("show vlan brief", use_textfsm=True)
        output = output_raw if isinstance(output_raw, list) else []
        if output:
            for vlan_entry in output:
                ports_list = vlan_entry.get('interfaces', [])
                ports_str = ", ".join(ports_list) if isinstance(ports_list, list) and ports_list else "N/A"
                vlan_name = vlan_entry.get('name', vlan_entry.get('vlan_name', 'N/A'))
                vlans.append({"id": vlan_entry.get('vlan_id', 'N/A'), "name": vlan_name,
                              "status": vlan_entry.get('status', 'N/A'), "ports": ports_str})
        return vlans
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
        output_raw = net_connect.send_command("show ip arp", use_textfsm=True)
        output = output_raw if isinstance(output_raw, list) else []
        if output:
            for entry in output:
                arp_table.append({
                    "protocol": entry.get('protocol', 'N/A'),
                    "address": entry.get('address', entry.get('ip_address', 'N/A')),
                    "age": entry.get('age', 'N/A'),
                    "mac_address": entry.get('mac', entry.get('mac_address', 'N/A')),
                    "type": entry.get('type', 'N/A'), "interface": entry.get('interface', 'N/A')})
        return arp_table
    except Exception as e:
        logger.critical(f"Critical error in get_arp_table for {net_connect.host}: {e}")
        traceback.print_exc()
        return []


def check_security_features(net_connect, running_config):
    """
    Performs a comprehensive security audit of the Cisco device configuration.

    Args:
        net_connect (netmiko.ConnectHandler): The established Netmiko connection object.
        running_config (str): The full running configuration of the device.

    Returns:
        dict: A dictionary containing the results of various security checks.
    """
    security_audit = {}
    virtual_prefixes = ("vl", "lo", "tu", "po", "bv", "nu", "oo", "gr")

    # --- I. AAA & Authentication ---
    if "aaa new-model" in running_config:
        security_audit["aaa_new_model"] = {"status": True, "level": "good",
                                           "details": "AAA new-model is enabled (prerequisite for TACACS+/RADIUS)."}
    else:
        security_audit["aaa_new_model"] = {"status": False, "level": "bad",
                                           "details": "AAA new-model is not enabled. Crucial for centralized and secure access management."}

    if "enable secret" in running_config:
        security_audit["enable_secret_configured"] = {"status": True, "level": "good",
                                                      "details": "An 'enable secret' is configured (strong hashing)."}
    elif "enable password" in running_config:
        security_audit["enable_secret_configured"] = {"status": "enable password only", "level": "bad",
                                                      "details": "'enable password' is used without 'enable secret'. Vulnerable even if 'service password-encryption' is active."}
    else:
        security_audit["enable_secret_configured"] = {"status": False, "level": "bad",
                                                      "details": "No 'enable secret' or 'enable password' configured. Privileged access unprotected."}

    if "service password-encryption" in running_config:
        security_audit["password_encryption_service"] = {"status": True, "level": "good",
                                                         "details": "Service 'password-encryption' enabled (obfuscates type 7 passwords, but does not secure them strongly)."}
    else:
        security_audit["password_encryption_service"] = {"status": False, "level": "bad",
                                                         "details": "Service 'password-encryption' NOT enabled. Passwords (except 'enable secret') stored in clear text."}

    # --- II. Access Line Security ---
    line_con_config_match = re.search(r"line con 0(.*?)!", running_config, re.DOTALL)
    con_config_text = line_con_config_match.group(1) if line_con_config_match else ""

    if "password" in con_config_text or "login local" in con_config_text or "login authentication" in con_config_text:
        security_audit["console_password"] = {"status": True, "level": "good",
                                              "details": "Console line protected by login method."}
    else:
        security_audit["console_password"] = {"status": False, "level": "bad",
                                              "details": "Console line not protected by password."}

    if "exec-timeout" in con_config_text:
        timeout_match = re.search(r"exec-timeout\s+(\d+)\s*(?:(\d+))?", con_config_text)
        if timeout_match:
            minutes, secondes = int(timeout_match.group(1)), int(timeout_match.group(2) or 0)
            if minutes > 0 or (minutes == 0 and secondes > 0):
                security_audit["console_exec_timeout"] = {"status": f"{minutes}m {secondes}s", "level": "good",
                                                          "details": "Console execution timeout configured."}
            else:
                security_audit["console_exec_timeout"] = {"status": "Disabled (0 0)", "level": "bad",
                                                          "details": "Console execution timeout disabled (0 0)."}
        else:
            security_audit["console_exec_timeout"] = {"status": "Configured (check value)", "level": "warning",
                                                      "details": "Console exec-timeout configured, check value."}
    else:
        security_audit["console_exec_timeout"] = {"status": False, "level": "warning",
                                                  "details": "No console execution timeout. Recommended: 5-15 minutes."}

    if "logging synchronous" in con_config_text:
        security_audit["console_logging_synchronous"] = {"status": True, "level": "good",
                                                         "details": "'logging synchronous' enabled on console."}
    else:
        security_audit["console_logging_synchronous"] = {"status": False, "level": "warning",
                                                         "details": "'logging synchronous' not enabled on console."}

    vty_config_text = ""
    vty_sections = re.findall(r"line vty\s+\d+\s*\d*(.*?)!", running_config, re.DOTALL)
    if vty_sections:
        vty_config_text = "\n".join(vty_sections)
    else:
        vty_config_text = running_config

    try:
        ssh_output_raw = net_connect.send_command("show ip ssh", use_textfsm=True)
        ssh_data = (ssh_output_raw[0] if isinstance(ssh_output_raw, list) and ssh_output_raw else {})
        ssh_version = ssh_data.get('protocol_version', 'N/A')
        if ssh_version == '2.0':
            security_audit["ssh_v2_only"] = {"status": True, "level": "good",
                                             "details": f"SSH version {ssh_version} enabled and seems to be the only version."}
        elif ssh_version != 'N/A':
            security_audit["ssh_v2_only"] = {"status": False, "level": "bad",
                                             "details": f"SSH version {ssh_version} detected. SSHv1 is enabled and vulnerable."}
        else:
            if "ip ssh version 2" in running_config and "no ip ssh version 1" in running_config:
                security_audit["ssh_v2_only"] = {"status": True, "level": "good",
                                                 "details": "SSH v2 explicitly configured, v1 disabled (config)."}
            elif "ip ssh version 2" in running_config:
                security_audit["ssh_v2_only"] = {"status": "v2 (v1 status unknown)", "level": "warning",
                                                 "details": "SSH v2 configured, but SSHv1 might still be active. Add 'no ip ssh version 1'."}
            elif "crypto key generate rsa" in running_config:
                security_audit["ssh_v2_only"] = {"status": "Unknown version", "level": "warning",
                                                 "details": "SSH seems enabled (RSA keys), version not confirmed v2 only."}
            else:
                security_audit["ssh_v2_only"] = {"status": False, "level": "bad",
                                                 "details": "SSH does not seem to be enabled or configured correctly."}
    except Exception as e:
        logger.warning(f"SSH check error for {net_connect.host}: {e}")
        security_audit["ssh_v2_only"] = {"status": "Error",
                                         "level": "warning",
                                         "details": "SSH check impossible."}

    if "transport input telnet" in vty_config_text.lower():
        security_audit["vty_transport_telnet"] = {"status": "Telnet Enabled", "level": "bad",
                                                  "details": "Telnet allowed on VTY lines. Unencrypted protocol, disable it."}
    elif "transport input ssh" in vty_config_text:
        security_audit["vty_transport_telnet"] = {"status": "SSH Only (Telnet not found)", "level": "good",
                                                  "details": "VTY lines seem configured for SSH only."}
    else:
        security_audit["vty_transport_telnet"] = {"status": "Transport Unclear", "level": "warning",
                                                  "details": "Non-standard VTY transport config. Ensure SSH only."}

    if "access-class" in vty_config_text and re.search(r"access-class\s+\S+\s+in", vty_config_text):
        security_audit["vty_acl"] = {"status": True, "level": "good",
                                     "details": "ACL (access-class) applied inbound on VTY lines."}
    else:
        security_audit["vty_acl"] = {"status": False, "level": "bad",
                                     "details": "No inbound ACL (access-class) on VTY lines. Risk of unfiltered access."}

    if "exec-timeout" in vty_config_text:
        vty_timeout_match = re.search(r"exec-timeout\s+(\d+)\s*(?:(\d+))?", vty_config_text)
        if vty_timeout_match:
            minutes, secondes = int(vty_timeout_match.group(1)), int(vty_timeout_match.group(2) or 0)
            if minutes > 0 or (minutes == 0 and secondes > 0):
                security_audit["vty_exec_timeout"] = {"status": f"{minutes}m {secondes}s", "level": "good",
                                                      "details": "VTY execution timeout configured."}
            else:
                security_audit["vty_exec_timeout"] = {"status": "Disabled (0 0)", "level": "bad",
                                                      "details": "VTY execution timeout disabled (0 0)."}
        else:
            security_audit["vty_exec_timeout"] = {"status": "Configured (check value)", "level": "warning",
                                                  "details": "VTY exec-timeout configured, check value."}
    else:
        security_audit["vty_exec_timeout"] = {"status": False, "level": "warning",
                                              "details": "No VTY execution timeout. Recommended: 5-15 minutes."}

    if "logging synchronous" in vty_config_text:
        security_audit["vty_logging_synchronous"] = {"status": True, "level": "good",
                                                     "details": "'logging synchronous' enabled on VTY."}
    else:
        security_audit["vty_logging_synchronous"] = {"status": False, "level": "warning",
                                                     "details": "'logging synchronous' not enabled on VTY."}

    # --- III. General Hardening ---
    if "no ip source-route" in running_config:
        security_audit["ip_source_route"] = {"status": "Disabled", "level": "good",
                                             "details": "Source routing disabled."}
    else:
        security_audit["ip_source_route"] = {"status": "Enabled (default)", "level": "bad",
                                             "details": "Source routing active. Configure 'no ip source-route'."}

    if "no service finger" in running_config or "service finger" not in running_config:
        security_audit["finger_service"] = {"status": "Disabled", "level": "good",
                                            "details": "Finger service disabled."}
    else:
        security_audit["finger_service"] = {"status": "Enabled", "level": "bad",
                                            "details": "Finger service enabled. Configure 'no service finger'."}

    small_servers_explicitly_disabled = "no service tcp-small-servers" in running_config and "no service udp-small-servers" in running_config
    small_servers_explicitly_enabled = "service tcp-small-servers" in running_config or "service udp-small-servers" in running_config
    if small_servers_explicitly_disabled or not small_servers_explicitly_enabled:
        security_audit["small_servers"] = {"status": "Disabled", "level": "good",
                                           "details": "tcp/udp-small-servers services disabled."}
    else:
        security_audit["small_servers"] = {"status": "Enabled", "level": "bad",
                                           "details": "tcp/udp-small-servers enabled. Disable them."}

    if "service timestamps log datetime msec" in running_config and "service timestamps debug datetime msec" in running_config:
        security_audit["service_timestamps"] = {"status": True, "level": "good",
                                                "details": "Precise timestamping (msec) logs/debug enabled."}
    elif "service timestamps log" in running_config:
        security_audit["service_timestamps"] = {"status": "Partial", "level": "warning",
                                                "details": "Partial logs timestamping. Recom: '... debug datetime msec' and '... log datetime msec'."}
    else:
        security_audit["service_timestamps"] = {"status": False, "level": "bad",
                                                "details": "Logs/debug timestamping not enabled. Essential for analysis."}

    if "banner motd" in running_config:
        security_audit["banner_motd"] = {"status": True, "level": "good", "details": "MOTD banner configured."}
    else:
        security_audit["banner_motd"] = {"status": False, "level": "warning", "details": "No MOTD banner."}

    if "no ip http server" in running_config:
        security_audit["http_server"] = {"status": "Disabled", "level": "good",
                                         "details": "HTTP server (insecure) disabled."}
    else:
        security_audit["http_server"] = {"status": "Enabled", "level": "bad",
                                         "details": "HTTP server (insecure) enabled. Use HTTPS or disable."}

    if "ip http secure-server" in running_config:
        security_audit["https_server"] = {"status": "Enabled", "level": "good",
                                          "details": "HTTPS server (secure) enabled."}
    elif "no ip http server" in running_config:
        security_audit["https_server"] = {"status": "Disabled (HTTP also disabled)", "level": "good",
                                          "details": "HTTPS server disabled (HTTP also disabled)."}
    else:
        security_audit["https_server"] = {"status": "Disabled (HTTP is Enabled)", "level": "bad",
                                          "details": "HTTPS server disabled while HTTP is active. Switch to HTTPS."}

    cdp_disabled, lldp_disabled = "no cdp run" in running_config, "no lldp run" in running_config
    security_audit["cdp_status"] = (
        {"status": "Globally Disabled", "level": "good", "details": "CDP globally disabled."} if cdp_disabled else
        {"status": "Globally Enabled", "level": "warning",
         "details": "CDP globally enabled. Filter on untrusted interfaces."})
    if lldp_disabled:
        security_audit["lldp_status"] = {"status": "Globally Disabled", "level": "good",
                                         "details": "LLDP globally disabled."}
    elif "lldp run" in running_config:
        security_audit["lldp_status"] = {"status": "Globally Enabled", "level": "warning",
                                         "details": "LLDP globally enabled. Filter on untrusted interfaces."}
    else:
        security_audit["lldp_status"] = {"status": "Potentially Enabled (default)", "level": "warning",
                                         "details": "LLDP potentially active by default."}

    try:
        int_status_raw = net_connect.send_command("show interfaces status", use_textfsm=True)
        int_status = int_status_raw if isinstance(int_status_raw, list) else []
        unused_ports_details, admin_down_ports_details = [], []
        for i_data in int_status:
            port_name, port_status = i_data.get('port'), i_data.get('status', '').lower()
            if not port_name or any(port_name.lower().startswith(p) for p in virtual_prefixes): continue
            if port_status in ['notconnect', 'disabled']: unused_ports_details.append(port_name)
            if port_status == 'disabled': admin_down_ports_details.append(port_name)
        active_unused = [p for p in unused_ports_details if p not in admin_down_ports_details]
        if not active_unused and not unused_ports_details:
            security_audit["unused_physical_ports"] = {"status": "All connected or N/A", "level": "good",
                                                       "details": "All physical ports are connected or status not detailed."}
        elif not active_unused:
            security_audit["unused_physical_ports"] = {"status": f"{len(admin_down_ports_details)} disabled",
                                                       "level": "good",
                                                       "details": f"Unused and disabled physical ports: {', '.join(admin_down_ports_details)}."}
        else:
            security_audit["unused_physical_ports"] = {"status": f"{len(active_unused)} active not connected",
                                                       "level": "bad",
                                                       "details": f"Physical ports not connected but active: {', '.join(active_unused)}. Risk. Shutdown them."}
    except Exception as e:
        logger.warning(f"Error unused_ports check for {net_connect.host}: {e}")
        security_audit["unused_physical_ports"] = {
            "status": "Error", "level": "warning", "details": f"Unused ports check impossible: {e}"}

    # --- IV. Logging & Monitoring ---
    if "logging buffered" in running_config:
        buflvl_match = re.search(r"logging buffered\s+(?:\d+|\S+)", running_config)
        if buflvl_match:
            security_audit["logging_buffered"] = {"status": f"Enabled ({buflvl_match.group(0).split()[-1]})",
                                                  "level": "good", "details": "Buffered logging enabled."}
        else:
            security_audit["logging_buffered"] = {"status": "Enabled (defaults)", "level": "good",
                                                  "details": "Buffered logging enabled (defaults)."}
    else:
        security_audit["logging_buffered"] = {"status": False, "level": "warning",
                                              "details": "Buffered logging not enabled."}

    if "logging host" in running_config or "logging server" in running_config:
        security_audit["remote_logging_configured"] = {"status": True, "level": "good",
                                                       "details": "Remote logging (syslog) configured."}
        if "logging source-interface" in running_config:
            src_int_match = re.search(r"logging source-interface\s+(\S+)", running_config)
            security_audit["logging_source_interface"] = {
                "status": src_int_match.group(1) if src_int_match else "Configured", "level": "good",
                "details": "Syslog source interface specified."}
        else:
            security_audit["logging_source_interface"] = {"status": False, "level": "warning",
                                                          "details": "No syslog source interface."}
    else:
        security_audit["remote_logging_configured"] = {"status": False, "level": "bad",
                                                       "details": "Remote logging (syslog) NOT configured."}
        security_audit["logging_source_interface"] = {"status": "N/A", "level": "bad",
                                                      "details": "Syslog not configured."}

    ntp_servers = [line for line in running_config.splitlines() if
                   "ntp server " in line.strip() and not line.strip().startswith("ntp server vrf")]
    num_ntp_servers = len(ntp_servers)
    ntp_sync_status, ntp_sync_details, ntp_sync_level = "Error", "NTP sync status unknown.", "warning"
    try:
        ntp_status_raw = net_connect.send_command("show ntp status", use_textfsm=True)
        ntp_data = (ntp_status_raw[0] if isinstance(ntp_status_raw, list) and ntp_status_raw else {})
        clock_state = ntp_data.get('clock_state', '').lower()
        if "synchronised" in clock_state or "synchronized" in clock_state:
            ntp_sync_status, ntp_sync_details, ntp_sync_level = True, f"NTP synchronized. Stratum: {ntp_data.get('stratum', 'N/A')}, Ref server: {ntp_data.get('reference_server', 'N/A')}.", "good"
        else:
            ntp_sync_status, ntp_sync_details = False, f"NTP not synchronized (state: {clock_state}). Logs incorrectly timestamped."
            ntp_sync_level = "bad"
    except Exception:
        ntp_sync_details = (
            "NTP configured, sync status unknown ('show ntp status' failed)." if num_ntp_servers > 0 else
            "NTP not configured and status unknown.")
        ntp_sync_level = "warning" if num_ntp_servers > 0 else "bad"
    security_audit["ntp_synchronization"] = {"status": ntp_sync_status, "level": ntp_sync_level,
                                             "details": ntp_sync_details}

    if num_ntp_servers >= 2:
        security_audit["ntp_redundancy"] = {"status": f"{num_ntp_servers} servers", "level": "good",
                                            "details": "NTP redundancy OK."}
    elif num_ntp_servers == 1:
        security_audit["ntp_redundancy"] = {"status": "1 server", "level": "warning",
                                            "details": "Only one NTP server. Recom: >=2."}
    else:
        security_audit["ntp_redundancy"] = {"status": "0 servers", "level": "bad",
                                            "details": "No NTP server configured. Unreliable time."}

    # --- V. Layer 2 Security (Indications) ---
    if "switchport port-security" in running_config:
        security_audit["port_security_feature_used"] = {"status": True, "level": "good",
                                                        "details": "Port Security feature used."}
    else:
        security_audit["port_security_feature_used"] = {"status": False, "level": "bad",
                                                        "details": "Port Security not used. Essential on access ports."}

    if "ip dhcp snooping" in running_config:
        security_audit["dhcp_snooping_global"] = {"status": True, "level": "good",
                                                  "details": "DHCP Snooping enabled globally."}
    else:
        security_audit["dhcp_snooping_global"] = {"status": False, "level": "bad",
                                                  "details": "DHCP Snooping not enabled globally. Required to prevent rogue DHCP servers."}

    if "spanning-tree portfast bpduguard default" in running_config:
        security_audit["bpduguard_default"] = {"status": True, "level": "good",
                                               "details": "BPDU Guard enabled by default on PortFast ports."}
    elif "spanning-tree bpduguard enable" in running_config:
        security_audit["bpduguard_default"] = {"status": "Per-interface (check)", "level": "good",
                                               "details": "BPDU Guard enabled on some interfaces."}
    else:
        security_audit["bpduguard_default"] = {"status": False, "level": "warning",
                                               "details": "BPDU Guard not enabled by default. Risk of STP loops."}

    if "storm-control" in running_config:
        security_audit["storm_control_used"] = {"status": True, "level": "good",
                                                "details": "Storm Control seems configured."}
    else:
        security_audit["storm_control_used"] = {"status": False, "level": "warning",
                                                "details": "Storm Control not used."}

    if "snmp-server community public RO" in running_config or "snmp-server community private RW" in running_config:
        security_audit["snmp_default_communities"] = {"status": True, "level": "bad",
                                                      "details": "SNMP uses DEFAULT communities. Major risk."}
    elif "snmp-server community" in running_config:
        security_audit["snmp_default_communities"] = {"status": False, "level": "good",
                                                      "details": "SNMP configured (no default communities)."}
    else:
        security_audit["snmp_default_communities"] = {"status": "Not Configured", "level": "good",
                                                      "details": "SNMP (v1/v2c community) not configured."}

    if "snmp-server group" in running_config and "v3 auth" in running_config:
        security_audit["snmp_version"] = {"status": "v3 (probable)", "level": "good",
                                          "details": "SNMPv3 seems used."}
    elif "snmp-server community" in running_config:
        security_audit["snmp_version"] = {"status": "v1/v2c", "level": "bad",
                                          "details": "SNMPv1/v2c used (plaintext communities). Prefer SNMPv3."}
    else:
        security_audit["snmp_version"] = {"status": "N/A", "level": "good", "details": "SNMP not configured."}
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


def generate_excel_report(all_data, excel_filepath):  # Modified to take full path
    """
    Generates a comprehensive Excel report from the collected audit data.

    Args:
        all_data (list): A list of dictionaries containing audit data for all devices.
        excel_filepath (str): The file path where the Excel report will be saved.
    """
    if not all_data:
        logger.warning("No data for Excel report.")
        return
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    ws_info = wb.create_sheet("General Info")
    headers_info = ["Hostname", "IP Address", "Model", "IOS Version", "Uptime", "Serial Number"]
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
    headers_interfaces = ["Hostname", "Interface", "Type", "Description", "IP Address", "Link Status",
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
            elif ("admin" in sl and "down" in sl) or sl == "disabled":
                lr.fill = ORANGE_FILL
            elif sl == "down" or "notconnect" in sl or "err-disabled" in sl or "error-disabled" in sl:
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
            if str(vlan.get("status", "")).lower() == "active":
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
    headers_security = ["Hostname", "Check Name", "Status/Value", "Level", "Details/Recommendation"]
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
    # excel_filepath is now passed as an argument
    try:
        wb.save(excel_filepath)
        logger.info(f"[+] Excel report generated: {excel_filepath}")
    except Exception as e:
        logger.error(f"[-] Error saving Excel: {e}")


def perform_cisco_audit(cisco_devices_inventory, global_passwords_map, output_directory):
    """
    Orchestrates the audit process for a list of Cisco devices.
    Connects to each device, collects data, performs security checks, and saves results.

    Args:
        cisco_devices_inventory (list): List of dictionaries representing Cisco devices to audit.
        global_passwords_map (dict): Dictionary mapping groups to credentials.
        output_directory (str): Directory to save JSON data and Excel reports.
    """
    all_devices_data = []
    for device_entry in cisco_devices_inventory:
        host, group = device_entry["host"], device_entry["group"]
        creds = global_passwords_map.get(group)
        logger.info(f"Processing {host} (group: {group})...")
        if not creds:
            logger.error(f"Credentials not found for group '{group}'. {host} skipped.")
            all_devices_data.append({"attempted_host": host, "status": "error_connection",
                                     "error_message": f"Credentials not found for group {group}"})
            continue
        dev_params = {'device_type': 'cisco_ios', 'host': host, 'username': creds['username'],
                      'password': creds['password'],
                      'secret': creds.get('enable_password'), 'global_delay_factor': 2, 'timeout': 45,
                      'session_timeout': 120}
        current_device_data = {"host": host}
        try:
            with ConnectHandler(**dev_params) as net_connect:
                actual_host, actual_prompt = net_connect.host, (
                    net_connect.base_prompt[:-1] if net_connect.base_prompt else net_connect.host)
                logger.info(f"Connected to {actual_host} ({actual_prompt}).")
                if creds.get('enable_password'): net_connect.enable()

                current_device_data["general_info"] = get_device_info(net_connect)
                current_device_data["general_info"]["ip_address_queried"] = host

                running_config = net_connect.send_command("show running-config", read_timeout=240)
                if not running_config: running_config = ""

                current_device_data["interfaces"] = get_interfaces(net_connect)
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
            logger.error(f"Unexpected with {host}: {e}")
            traceback.print_exc()
            all_devices_data.append(
                {"attempted_host": host, "status": "error_connection", "error_message": f"Unexpected error: {e}"})

    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    json_filename = os.path.join(output_directory, f"audit_cisco_data_{timestamp}.json")
    try:
        with open(json_filename, 'w', encoding='utf-8') as f:
            json.dump(all_devices_data, f, indent=4, ensure_ascii=False)
        logger.info(f"[+] Cisco JSON data saved: {json_filename}")
    except Exception as e:
        logger.error(f"[-] Error saving Cisco JSON: {e}")

    excel_report_path = os.path.join(output_directory, f"audit_cisco_report_{timestamp}.xlsx")
    generate_excel_report(all_devices_data, excel_report_path)
    logger.info(f"Cisco audit completed for {len(cisco_devices_inventory)} device(s).")


def main():
    """
    Main entry point for standalone execution of the Cisco audit script.
    Loads inventory, credentials, filter for Cisco devices, and starts the audit.
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
        logger.error("Stop (Cisco): Critical errors loading input files.")
        return

    cisco_devices = [device for device in full_inventory if device.get("device_type") in ["cisco_ios", "cisco_iosxe"]]

    if not cisco_devices:
        logger.warning("No Cisco devices (cisco_ios/cisco_iosxe) found in inventory for autonomous audit.")
        return

    perform_cisco_audit(cisco_devices, passwords_map, output_directory)


if __name__ == "__main__":
    main()
