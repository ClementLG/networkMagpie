import logging
import traceback

logger = logging.getLogger(__name__)

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
