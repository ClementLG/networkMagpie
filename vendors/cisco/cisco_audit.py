from vendors.base_vendor import BaseVendor
from .system import get_device_info
from .interfaces import get_interfaces
from .l2 import get_vlans, get_arp_table
from .security import check_security_features
from .config import get_running_config

class CiscoAudit(BaseVendor):
    def get_device_info(self):
        return get_device_info(self.net_connect)

    def get_interfaces(self):
        return get_interfaces(self.net_connect)

    def get_vlans(self):
        return get_vlans(self.net_connect)

    def get_arp_table(self):
        return get_arp_table(self.net_connect)

    def check_security_features(self, running_config):
        return check_security_features(self.net_connect, running_config)

    def get_running_config(self):
        return get_running_config(self.net_connect)
