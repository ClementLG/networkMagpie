from abc import ABC, abstractmethod
import logging

class BaseVendor(ABC):
    def __init__(self, net_connect):
        self.net_connect = net_connect
        self.logger = logging.getLogger(__name__)

    @abstractmethod
    def get_device_info(self):
        """Retrieves general device information."""
        pass

    @abstractmethod
    def get_interfaces(self):
        """Retrieves interface information."""
        pass

    @abstractmethod
    def get_vlans(self):
        """Retrieves VLAN information."""
        pass

    @abstractmethod
    def get_arp_table(self):
        """Retrieves ARP table."""
        pass

    @abstractmethod
    def check_security_features(self, running_config):
        """Performs security audit checks."""
        pass
    
    def run_audit(self, host):
        """
        Orchestrates the audit process for the device.
        """
        device_data = {"host": host}
        
        # 1. Device Info
        device_data["general_info"] = self.get_device_info()
        device_data["general_info"]["ip_address_queried"] = host

        # 2. Running Config (needed for security checks)
        # Using a default timeout, subclasses might override this specific command if needed
        # but usually 'show running-config' is standard-ish or netmiko handles it.
        # For granular control, we can make get_running_config a method.
        running_config = self.get_running_config()

        # 3. Other Data
        device_data["interfaces"] = self.get_interfaces()
        device_data["vlans"] = self.get_vlans()
        device_data["arp_table"] = self.get_arp_table()
        device_data["security_audit"] = self.check_security_features(running_config)
        
        return device_data

    def get_running_config(self):
        """
        Retrieves the running configuration.
        Can be overridden if the command is different.
        """
        try:
            # Default to cisco-like
            output = self.net_connect.send_command("show running-config", read_timeout=60)
            return output if output else ""
        except Exception as e:
            self.logger.error(f"Error getting running config: {e}")
            return ""
