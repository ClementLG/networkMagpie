# main_audit.py
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

import os
import datetime


# Import main functions from audit scripts
from cisco_audit import perform_cisco_audit, load_inventory as load_inventory_cisco, \
    load_passwords as load_passwords_cisco, generate_excel_report as generate_excel_cisco
from aruba_audit import perform_aruba_audit, load_inventory as load_inventory_aruba, \
    load_passwords as load_passwords_aruba, generate_excel_report as generate_excel_aruba


import logging

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def main():
    """
    Main entry point for the network audit script.
    Loads inventory and passwords, creates output directories, and triggers
    specific audit functions for Cisco and Aruba devices.
    """
    inventory_file = "inventory.csv"
    password_file = "passwords.csv"
    base_output_directory = "audit_reports"

    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    session_output_directory = base_output_directory

    if not os.path.exists(session_output_directory):
        try:
            os.makedirs(session_output_directory)
        except OSError as e:
            logger.error(f"Error: Unable to create output directory '{session_output_directory}': {e}")
            return


    # Load inventory and passwords
    full_inventory = load_inventory_cisco(inventory_file)
    passwords_map = load_passwords_cisco(password_file)

    if full_inventory is None or passwords_map is None:
        logger.critical("Critical error loading inventory or password files. Aborting.")
        return

    if not full_inventory:
        logger.warning("Inventory is empty. Nothing to do.")
        return

    cisco_devices_to_audit = []
    aruba_devices_to_audit = []

    for device in full_inventory:
        dev_type = device.get("device_type", "").lower()  # Normalize to lowercase
        if dev_type == "cisco_ios" or dev_type == "cisco_iosxe":  # Accept both for Cisco
            cisco_devices_to_audit.append(device)
        elif dev_type == "aruba_os-cx":
            aruba_devices_to_audit.append(device)
        else:
            logger.warning(
                f"Warning: Unknown or missing device type '{device.get('device_type')}' for {device.get('host')}. Ignored.")

    if cisco_devices_to_audit:
        logger.info(f"--- Starting audit for {len(cisco_devices_to_audit)} Cisco device(s) ---")
        perform_cisco_audit(cisco_devices_to_audit, passwords_map, session_output_directory)
    else:
        logger.info("--- No Cisco devices to audit ---")

    if aruba_devices_to_audit:
        logger.info(f"--- Starting audit for {len(aruba_devices_to_audit)} Aruba device(s) ---")
        perform_aruba_audit(aruba_devices_to_audit, passwords_map, session_output_directory)
    else:
        logger.info("--- No Aruba devices to audit ---")

    logger.info("All scheduled audits completed.")


if __name__ == "__main__":
    main()
