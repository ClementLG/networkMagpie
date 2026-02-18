import csv
import logging
import traceback

logger = logging.getLogger(__name__)

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
