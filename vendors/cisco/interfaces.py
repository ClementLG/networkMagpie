import logging
import re
import traceback

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
