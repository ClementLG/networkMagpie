import logging
import re
import traceback
from textfsm.parser import TextFSMError

logger = logging.getLogger(__name__)

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
