import logging
import re
import traceback
from textfsm.parser import TextFSMError

logger = logging.getLogger(__name__)

def get_device_info(net_connect):
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
