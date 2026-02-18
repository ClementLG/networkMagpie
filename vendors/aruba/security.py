import logging
import re

logger = logging.getLogger(__name__)

def check_security_features(net_connect, running_config):
    """
    Performs a comprehensive security audit of the Aruba OS-CX device configuration.

    Args:
        net_connect (netmiko.ConnectHandler): The established Netmiko connection object.
        running_config (str): The full running configuration of the device.

    Returns:
        dict: A dictionary containing the results of various security checks.
    """
    security_audit = {}
    logger.info(f"Starting security audit for {net_connect.host}...")

    # --- I. AAA & Authentication & Management Access ---
    # 1. AAA Port Access (Edge Security)
    if "aaa authentication port-access" in running_config:
        security_audit["aaa_port_access_configured"] = {"status": True, "level": "good",
                                                        "details": "AAA for port access (dot1x/mac-auth) seems configured."}
    else:
        security_audit["aaa_port_access_configured"] = {"status": False, "level": "warning",
                                                        "details": "AAA for port access (dot1x/mac-auth) not detected. Recommended to secure network access."}

    # 2. AAA Login (Management Security) - CRITICAL MISSING CHECK ADDED
    if "aaa authentication login" in running_config:
        if "group tacacs" in running_config or "group radius" in running_config:
            security_audit["aaa_login_configured"] = {"status": "Centralized (TACACS+/RADIUS)", "level": "good",
                                                      "details": "AAA login authentication uses centralized server."}
        else:
            security_audit["aaa_login_configured"] = {"status": "Local/Other", "level": "warning",
                                                      "details": "AAA login configured but might be local only. Verify 'aaa authentication login' config."}
    else:
        security_audit["aaa_login_configured"] = {"status": "Not Configured", "level": "bad",
                                                  "details": "No 'aaa authentication login' found. Default local auth used?"}

    # 3. Local User Password Type
    local_user_passwords_encrypted = True
    # Check for plaintext passwords explicitly
    if re.search(r"user\s+\S+\s+password\s+plaintext", running_config):
        local_user_passwords_encrypted = False

    if local_user_passwords_encrypted and "password" in running_config:
        security_audit["local_user_password_encryption"] = {"status": "Encrypted (hashed)", "level": "good",
                                                            "details": "Local user passwords seem to be stored encrypted (hashed)."}
    elif not local_user_passwords_encrypted:
        security_audit["local_user_password_encryption"] = {"status": "Plaintext detected", "level": "bad",
                                                            "details": "At least one local user password is stored in plaintext. Use 'password ciphertext <hash>'."}
    else:
        # Case where no "password" keyword found or ambiguous
        security_audit["local_user_password_encryption"] = {
            "status": "No local user with password found", "level": "warning",
            "details": "Manually check local user configuration."}

    # Password complexity policy
    min_len_str, complexity_str = "Not configured", "Not configured"
    min_len_level, complexity_level = "bad", "bad"

    match_pass_len = re.search(r"password minimum-length\s+(\d+)", running_config)
    if match_pass_len:
        min_len = int(match_pass_len.group(1))
        min_len_str = f"Min length: {min_len}"
        min_len_level = "good" if min_len >= 12 else "warning" if min_len >= 8 else "bad"
    
    security_audit["password_min_length"] = {"status": min_len_str, "level": min_len_level,
                                             "details": f"{min_len_str}. Recommended: >=12."}

    if "password complexity" in running_config or "character-class-check" in running_config:
        complexity_str = "Enabled (check details)"
        complexity_level = "good"
    security_audit["password_complexity"] = {"status": complexity_str, "level": complexity_level,
                                             "details": f"Password complexity policy: {complexity_str}."}

    # --- II. Access Line Security (Console) ---
    console_config_text = ""
    console_match = re.search(r"line console\s*\n(.*?)(?=line|interface|vlan|router|exit|$)", running_config,
                              re.DOTALL | re.MULTILINE)
    if console_match:
        console_config_text = console_match.group(1)

    if console_config_text and ("password " in console_config_text or "login " in console_config_text or "aaa authentication login" in running_config):
         # Logic assumption: if global AAA login is set, console might use it imply default. 
         # But safer to look for specific line config or inheritance.
         # For Audit purposes, if "login" is there it's usually good.
        security_audit["console_auth"] = {"status": True, "level": "good",
                                          "details": "Console line protected by login method."}
    else:
        security_audit["console_auth"] = {"status": False, "level": "bad",
                                          "details": "Console line authentication not explicitly seen in 'line console'."}

    exec_timeout_con_match = re.search(r"session-timeout\s+(\d+)", console_config_text)
    if exec_timeout_con_match:
        minutes_con = int(exec_timeout_con_match.group(1))
        if 0 < minutes_con <= 15:
            security_audit["console_session_timeout"] = {"status": f"{minutes_con} minutes", "level": "good",
                                                         "details": "Session timeout configured on console."}
        elif minutes_con == 0:
            security_audit["console_session_timeout"] = {"status": "Disabled (0)", "level": "bad",
                                                         "details": "Console session timeout disabled. Risk."}
        else:
            security_audit["console_session_timeout"] = {"status": f"{minutes_con} minutes", "level": "warning",
                                                         "details": "Console session timeout high. Recommended: 5-15 min."}
    else:
        security_audit["console_session_timeout"] = {"status": "Not configured (default)", "level": "warning",
                                                     "details": "No session timeout on console. Recommended: 5-15 min."}

    # --- III. Management Services Security ---
    # SSH
    try:
        ssh_status_raw = net_connect.send_command("show ssh server all-vrfs", expect_string=r"#")
        
        ssh_enabled = False
        ssh_v1 = False
        
        if "SSH server configuration on VRF" in ssh_status_raw:
            ssh_enabled = True # Server is running
        
        # Check specific version disablement in config
        # "no ssh server v1 enable" -> Good
        # "ssh server v1 enable" -> Bad
        
        if "ssh server v1 enable" in running_config:
            ssh_v1 = True
        elif "no ssh server v1 enable" in running_config:
            ssh_v1 = False
        else:
            # Default behavior of OS-CX? Old versions enabled, newer disabled. 
            # We mark as warning if not explicitly disabled.
            ssh_v1 = "Unknown (Implicit)"

        if ssh_enabled:
            if ssh_v1 is True:
                 security_audit["ssh_status"] = {"status": "SSHv1 Enabled", "level": "bad", "details": "SSHv1 explicitly enabled. Disable it."}
            elif ssh_v1 is False:
                 security_audit["ssh_status"] = {"status": "SSHv2 Only", "level": "good", "details": "SSHv2 enabled, v1 disabled."}
            else:
                 security_audit["ssh_status"] = {"status": "SSH Enabled (v1 implicit)", "level": "warning", "details": "Verify if SSHv1 is disabled by default or add 'no ssh server v1 enable'."}
        else:
             # Weird if we are connected via SSH...
             security_audit["ssh_status"] = {"status": "Disabled?", "level": "warning", "details": "SSH server appears disabled in 'show ssh', yet we are connected?"}

    except Exception as e:
        logger.error(f"SSH check failed: {e}")
        security_audit["ssh_status"] = {"status": "Error", "level": "error", "details": "Failed to check SSH status."}

    # Telnet
    if "no telnet-server enable" in running_config:
        security_audit["telnet_server"] = {"status": "Disabled", "level": "good", "details": "Telnet server disabled."}
    elif "telnet-server enable" in running_config:
        security_audit["telnet_server"] = {"status": "Enabled", "level": "bad", "details": "Telnet server enabled. Disable it."}
    else:
        # Default is usually disabled on modern AOS-CX, but good verify.
        security_audit["telnet_server"] = {"status": "Default", "level": "warning", "details": "Explicit 'no telnet-server enable' recommended."}
        
    # TFTP - New Check
    if "tftp-server enable" in running_config:
        security_audit["tftp_server"] = {"status": "Enabled", "level": "bad", "details": "TFTP server enabled. Insecure."}
    else:
        security_audit["tftp_server"] = {"status": "Disabled", "level": "good", "details": "TFTP server not enabled."}

    # HTTP/HTTPS
    https_enabled = "https-server enable" in running_config or "https-server vrf" in running_config
    http_enabled = "http-server enable" in running_config or "http-server vrf" in running_config
    
    if https_enabled:
        security_audit["https_server"] = {"status": "Enabled", "level": "good", "details": "HTTPS server (web-mgmt) enabled."}
    else:
        security_audit["https_server"] = {"status": "Disabled", "level": "good", "details": "HTTPS server disabled."}
        
    if http_enabled:
        security_audit["http_server"] = {"status": "Enabled", "level": "bad", "details": "HTTP server enabled. Insecure."}
    else:
         security_audit["http_server"] = {"status": "Disabled", "level": "good", "details": "HTTP server disabled."}

    # Banners
    banners_set = []
    if "banner motd" in running_config: banners_set.append("MOTD")
    if "banner exec" in running_config: banners_set.append("Exec")
    
    if banners_set:
        security_audit["banners_configured"] = {"status": f"Configured: {', '.join(banners_set)}", "level": "good", "details": "Warning banners configured."}
    else:
        security_audit["banners_configured"] = {"status": "None", "level": "warning", "details": "No warning banners configured."}

    # --- IV. General Hardening ---
    # IP Source Route - Fixed Logic
    # On many Aruba CX, "no ip source-route" might not be a command or might be default.
    # We check if "ip source-route" IS present.
    if "ip source-route" in running_config and "no ip source-route" not in running_config:
         security_audit["ip_source_route"] = {"status": "Enabled", "level": "bad", "details": "IP Source Routing enabled."}
    else:
         security_audit["ip_source_route"] = {"status": "Disabled", "level": "good", "details": "IP Source Routing not found (disabled)."}

    # LLDP
    if "no lldp enable" in running_config:
        security_audit["lldp_global_status"] = {"status": "Globally Disabled", "level": "good", "details": "LLDP disabled."}
    else:
        security_audit["lldp_global_status"] = {"status": "Globally Enabled", "level": "warning", "details": "LLDP globally enabled. Secure untrusted ports."}

    # --- V. Logging & Monitoring ---
    if "logging syslog host" in running_config or "logging host" in running_config:
        security_audit["remote_logging"] = {"status": True, "level": "good", "details": "Remote logging (syslog) configured."}
    else:
        security_audit["remote_logging"] = {"status": False, "level": "bad", "details": "Remote logging (syslog) NOT configured."}

    # NTP
    ntp_lines = [line for line in running_config.splitlines() if line.strip().startswith("ntp server")]
    if len(ntp_lines) >= 2:
        security_audit["ntp_redundancy"] = {"status": f"{len(ntp_lines)} servers", "level": "good", "details": "NTP redundancy OK."}
    elif len(ntp_lines) == 1:
        security_audit["ntp_redundancy"] = {"status": "1 server", "level": "warning", "details": "Only one NTP server configured."}
    else:
        security_audit["ntp_redundancy"] = {"status": "None", "level": "bad", "details": "No NTP servers configured."}
        
    # NTP Auth (Extra)
    if any(" key " in line for line in ntp_lines):
         security_audit["ntp_auth"] = {"status": "Configured", "level": "good", "details": "NTP authentication seems used."}
    else:
         security_audit["ntp_auth"] = {"status": "Not Configured", "level": "warning", "details": "NTP authentication not detected."}

    # SNMP
    if "snmp-server community public" in running_config or "snmp-server community private" in running_config:
        security_audit["snmp_default_communities"] = {"status": "Found", "level": "bad", "details": "Default SNMP communities (public/private) present."}
    else:
        security_audit["snmp_default_communities"] = {"status": "Not Found", "level": "good", "details": "No default public/private communities found."}

    snmpv3_configured = "snmp-server user" in running_config
    if snmpv3_configured:
        security_audit["snmp_version"] = {"status": "SNMPv3", "level": "good", "details": "SNMPv3 users configured."}
    elif "snmp-server community" in running_config:
        security_audit["snmp_version"] = {"status": "SNMPv1/v2c", "level": "warning", "details": "Only SNMPv1/v2c communities found. Prefer v3."}
    else:
        security_audit["snmp_version"] = {"status": "None", "level": "good", "details": "SNMP not configured."}

    # --- VI. Layer 2 Security ---
    if "bpdu-protection" in running_config:
        security_audit["bpdu_protection"] = {"status": "Active", "level": "good", "details": "BPDU protection active globally or on ports."}
    else:
        security_audit["bpdu_protection"] = {"status": "Inactive", "level": "warning", "details": "BPDU protection not found."}

    if "dhcp-snooping" in running_config:
         security_audit["dhcp_snooping"] = {"status": "Active", "level": "good", "details": "DHCP Snooping configured."}
    else:
         security_audit["dhcp_snooping"] = {"status": "Inactive", "level": "bad", "details": "DHCP Snooping not configured."}

    return security_audit
