import logging
import re
import traceback
from textfsm.parser import TextFSMError

logger = logging.getLogger(__name__)

def get_interfaces(net_connect):
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
