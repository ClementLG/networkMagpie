import logging

logger = logging.getLogger(__name__)

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
