import logging

def get_running_config(net_connect):
    """
    Retrieves the running configuration from an Aruba device.
    """
    logger = logging.getLogger(__name__)
    try:
        output = net_connect.send_command("show running-config", read_timeout=60)
        return output if output else ""
    except Exception as e:
        logger.error(f"Error getting running config (Aruba): {e}")
        return ""
