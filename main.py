import logging
import os
import json
import datetime
import traceback
from netmiko import ConnectHandler, NetmikoTimeoutException, NetmikoAuthenticationException
from common.file_utils import load_inventory, load_passwords
from common.report_generator import generate_excel_report
from vendors.cisco.cisco_audit import CiscoAudit
from vendors.aruba.aruba_audit import ArubaAudit

# Logging Configuration
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("network_magpie.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

import argparse
from common.html_report_generator import generate_html_report

def main():
    # Argument Parsing
    parser = argparse.ArgumentParser(description="Network Magpie Audit Tool")
    parser.add_argument(
        '-o', '--output',
        nargs='+',
        default=['excel'],
        help="Output formats: json, excel, html. Default: excel. Can be space or comma separated."
    )
    parser.add_argument(
        '-e', '--export-config',
        action='store_true',
        help="Export running configuration to text files."
    )
    args = parser.parse_args()

    # Process output formats (handle commas)
    selected_outputs = set()
    for fmt in args.output:
        for item in fmt.split(','):
            item = item.strip().lower()
            if item in ['json', 'excel', 'html']:
                selected_outputs.add(item)
            else:
                print(f"Warning: Unknown output format '{item}' ignored.")
    
    # If for some reason nothing valid was selected, default to excel
    if not selected_outputs:
        selected_outputs.add('excel')

    logger.info(f"Starting NetworkMagpie Audit... Selected outputs: {', '.join(selected_outputs)}")

    # Configuration
    inventory_file = "inventory.csv"
    password_file = "passwords.csv"
    output_directory = "audit_reports"

    if not os.path.exists(output_directory):
        os.makedirs(output_directory)

    # Configs directory
    configs_directory = os.path.join(output_directory, "configs")
    if args.export_config and not os.path.exists(configs_directory):
        os.makedirs(configs_directory)

    # Load Data
    inventory = load_inventory(inventory_file)
    passwords = load_passwords(password_file)

    if not inventory:
        logger.error("Inventory is empty or could not be loaded.")
        return
    if not passwords:
        logger.error("Passwords file is empty or could not be loaded.")
        return

    all_devices_data = []

    for device in inventory:
        host = device['host']
        group = device['group']
        dev_type_csv = device['device_type']
        
        logger.info(f"Processing {host} (Group: {group}, Type: {dev_type_csv})...")

        creds = passwords.get(group)
        if not creds:
            logger.error(f"Credentials not found for group '{group}'. Skipping {host}.")
            all_devices_data.append({
                "attempted_host": host, 
                "status": "error_connection", 
                "error_message": f"Credentials not found for group {group}"
            })
            continue

        # Determine Audit Class and Netmiko Device Type
        AuditClass = None
        netmiko_type = None

        if dev_type_csv in ['cisco_ios', 'cisco_iosxe']:
            AuditClass = CiscoAudit
            netmiko_type = 'cisco_ios'
        elif dev_type_csv in ['aruba_os-cx', 'aruba_aoscx_ssh']:
            AuditClass = ArubaAudit
            netmiko_type = 'aruba_aoscx_ssh'
        else:
            logger.warning(f"Unknown or unsupported device type '{dev_type_csv}' for {host}.")
            all_devices_data.append({
                "attempted_host": host, 
                "status": "error_connection", 
                "error_message": f"Unsupported device type: {dev_type_csv}"
            })
            continue

        # Prepare Connection Params
        device_params = {
            'device_type': netmiko_type,
            'host': host,
            'username': creds['username'],
            'password': creds['password'],
            'secret': creds.get('enable_password'),
            'global_delay_factor': 2,
            'timeout': 45,
            'session_timeout': 120
        }

        try:
            with ConnectHandler(**device_params) as net_connect:
                # Handle Enable Mode if needed
                if creds.get('enable_password'):
                    net_connect.enable()
                
                logger.info(f"Connected to {host}. Running audit...")
                auditor = AuditClass(net_connect)
                device_data = auditor.run_audit(host)
                # Mark status as success if not already set (run_audit returns dict without status key usually)
                device_data['status'] = 'success'

                # Handle Config Export if requested
                if args.export_config and 'running_config' in device_data:
                    config_content = device_data['running_config']
                    # Sanitize hostname for filename
                    safe_hostname = "".join([c for c in host if c.isalpha() or c.isdigit() or c in (' ', '-', '_')]).strip()
                    config_filename = os.path.join(configs_directory, f"{safe_hostname}_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.cfg")
                    try:
                        with open(config_filename, 'w', encoding='utf-8') as f:
                            f.write(config_content)
                        logger.info(f"Configuration exported to {config_filename}")
                    except Exception as e:
                        logger.error(f"Failed to export configuration for {host}: {e}")
                    
                    # Remove running_config from data to avoid bloating reports
                    del device_data['running_config']
                elif 'running_config' in device_data:
                     # Remove running_config even if not exporting, to keep reports clean
                    del device_data['running_config']

                all_devices_data.append(device_data)
                
        except (NetmikoTimeoutException, NetmikoAuthenticationException) as e:
            logger.error(f"Connection/Auth error for {host}: {e}")
            all_devices_data.append({
                "attempted_host": host, 
                "status": "error_connection", 
                "error_message": str(e)
            })
        except Exception as e:
            logger.error(f"Unexpected error for {host}: {e}")
            traceback.print_exc()
            all_devices_data.append({
                "attempted_host": host, 
                "status": "error_connection", 
                "error_message": f"Unexpected error: {e}"
            })

    # Save Results
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    
    # JSON
    if 'json' in selected_outputs:
        json_filename = os.path.join(output_directory, f"audit_data_{timestamp}.json")
        try:
            with open(json_filename, 'w', encoding='utf-8') as f:
                json.dump(all_devices_data, f, indent=4, ensure_ascii=False)
            logger.info(f"Full JSON data saved to {json_filename}")
        except Exception as e:
            logger.error(f"Error saving JSON data: {e}")

    # Excel
    if 'excel' in selected_outputs:
        excel_filename = os.path.join(output_directory, f"audit_report_{timestamp}.xlsx")
        try:
            generate_excel_report(all_devices_data, excel_filename)
            logger.info(f"Excel report saved to {excel_filename}")
        except Exception as e:
            logger.error(f"Error generating Excel report: {e}")

    # HTML
    if 'html' in selected_outputs:
        html_filename = os.path.join(output_directory, f"audit_report_{timestamp}.html")
        try:
            generate_html_report(all_devices_data, html_filename)
            logger.info(f"HTML report saved to {html_filename}")
        except Exception as e:
            logger.error(f"Error generating HTML report: {e}")

    logger.info("Audit completed.")

if __name__ == "__main__":
    main()
