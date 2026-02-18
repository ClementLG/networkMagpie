import logging

def get_running_config(net_connect):
    """
    Retrieves the running configuration from a Cisco device.
    """
    logger = logging.getLogger(__name__)
    try:
        # Check privileges (optional but good practice)
        if not net_connect.check_enable_mode():
            net_connect.enable()
            
        output = net_connect.send_command("show running-config", read_timeout=60)
        return output if output else ""
    except Exception as e:
        logger.error(f"Error getting running config (Cisco): {e}")
        return ""
