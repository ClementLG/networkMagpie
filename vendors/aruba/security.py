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

    # Helper for positive feature check (must be present and NOT negated)
    def check_feature_enabled(pattern, config):
        # Look for the pattern at the start of a line (ignoring whitespace)
        # We search for exact match at start of line.
        if re.search(fr"^\s*{pattern}", config, re.MULTILINE):
            return True
        return False

    # --- I. AAA & Authentication & Management Access ---
    # 1. AAA Port Access (Edge Security)
    # Check for "aaa authentication port-access dot1x authenticator" or "aaa authentication port-access mac-auth"
    if check_feature_enabled(r"aaa authentication port-access", running_config):
        security_audit["aaa_port_access_configured"] = {"status": True, "level": "good",
                                                        "details": "AAA for port access (dot1x/mac-auth) seems configured."}
    else:
        security_audit["aaa_port_access_configured"] = {"status": False, "level": "warning",
                                                        "details": "AAA for port access (dot1x/mac-auth) not detected. Recommended to secure network access."}

    # 2. AAA Login (Management Security)
    # Check for "aaa authentication login default group ..." or similar
    if check_feature_enabled(r"aaa authentication login", running_config):
        # checking for centralized auth
        if re.search(r"aaa authentication login \S+ group (tacacs|radius)", running_config):
            security_audit["aaa_login_configured"] = {"status": "Centralized (TACACS+/RADIUS)", "level": "good",
                                                      "details": "AAA login authentication uses centralized server."}
        else:
            security_audit["aaa_login_configured"] = {"status": "Local/Other", "level": "warning",
                                                      "details": "AAA login configured but might be local only."}
    else:
        security_audit["aaa_login_configured"] = {"status": "Not Configured", "level": "bad",
                                                  "details": "No 'aaa authentication login' found. Default local auth used?"}

    # 3. Local User Password Type
    local_user_passwords_encrypted = True
    # Check for plaintext passwords: "user <name> password plaintext <pass>"
    if re.search(r"^\s*user\s+\S+\s+password\s+plaintext", running_config, re.MULTILINE):
        local_user_passwords_encrypted = False

    # Check if we have any user with password
    has_user_password = re.search(r"^\s*user\s+\S+\s+password", running_config, re.MULTILINE)

    if has_user_password:
        if local_user_passwords_encrypted:
            security_audit["local_user_password_encryption"] = {"status": "Encrypted (hashed)", "level": "good",
                                                                "details": "Local user passwords seem to be stored encrypted (hashed)."}
        else:
            security_audit["local_user_password_encryption"] = {"status": "Plaintext detected", "level": "bad",
                                                                "details": "At least one local user password is stored in plaintext. Use 'password ciphertext <hash>'."}
    else:
        security_audit["local_user_password_encryption"] = {
            "status": "No local user with password found", "level": "warning",
            "details": "Manually check local user configuration."}

    # Password complexity policy
    min_len_match = re.search(r"^\s*password minimum-length\s+(\d+)", running_config, re.MULTILINE)
    if min_len_match:
        min_len = int(min_len_match.group(1))
        min_len_str = f"Min length: {min_len}"
        min_len_level = "good" if min_len >= 12 else "warning" if min_len >= 8 else "bad"
        security_audit["password_min_length"] = {"status": min_len_str, "level": min_len_level,
                                                 "details": f"{min_len_str}. Recommended: >=12."}
    else:
         security_audit["password_min_length"] = {"status": "Not configured", "level": "bad",
                                              "details": "Password minimum length not configured."}

    if check_feature_enabled(r"password complexity", running_config) or check_feature_enabled(r"password-complexity", running_config):
        security_audit["password_complexity"] = {"status": "Enabled", "level": "good",
                                                 "details": "Password complexity checks enabled."}
    else:
        security_audit["password_complexity"] = {"status": "Disabled", "level": "warning",
                                                 "details": "Password complexity not explicitly enabled."}

    # --- II. Access Line Security (Console) ---
    console_match = re.search(r"^line console\s*\n(.*?)(?=^line|^interface|^vlan|^router|^exit|$)", running_config,
                              re.DOTALL | re.MULTILINE)
    console_config_text = console_match.group(1) if console_match else ""

    if console_config_text and (re.search(r"^\s*password", console_config_text, re.MULTILINE) or 
                                re.search(r"^\s*login", console_config_text, re.MULTILINE) or 
                                check_feature_enabled(r"aaa authentication login", running_config)):
        security_audit["console_auth"] = {"status": True, "level": "good",
                                          "details": "Console line protected by login or password."}
    else:
        security_audit["console_auth"] = {"status": False, "level": "bad",
                                          "details": "Console line authentication not explicitly seen in 'line console'."}

    exec_timeout_con_match = re.search(r"^\s*session-timeout\s+(\d+)", console_config_text, re.MULTILINE)
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
        ssh_status_raw = net_connect.send_command("show ssh server", expect_string=r"#")
        
        # Check if SSH is globally enabled/disabled via config
        # "no ssh server" -> Disabled
        # "ssh server" -> Enabled (usually default)
        
        is_ssh_disabled_conf = re.search(r"^\s*no ssh server\s*$", running_config, re.MULTILINE)
        
        if is_ssh_disabled_conf:
             security_audit["ssh_status"] = {"status": "Disabled", "level": "bad", "details": "'no ssh server' found. SSH is disabled."}
        else:
             # Check for SSHv1
             # "ssh server v1 enable" -> v1 enabled (bad)
             # "no ssh server v1 enable" -> v1 disabled (good)
             if check_feature_enabled(r"ssh server v1 enable", running_config):
                  security_audit["ssh_status"] = {"status": "SSHv1 Enabled", "level": "bad", "details": "SSHv1 explicitly enabled. Disable it."}
             else:
                  security_audit["ssh_status"] = {"status": "SSHv2 Only", "level": "good", "details": "SSH enabled (v1 checks ok)."}

    except Exception as e:
        logger.error(f"SSH check failed: {e}")
        security_audit["ssh_status"] = {"status": "Error", "level": "error", "details": "Failed to check SSH status."}

    # Telnet
    # "no telnet-server enable" (Good)
    if re.search(r"^\s*no telnet-server", running_config, re.MULTILINE):
        security_audit["telnet_server"] = {"status": "Disabled", "level": "good", "details": "Telnet server disabled."}
    elif check_feature_enabled(r"telnet-server", running_config):
        security_audit["telnet_server"] = {"status": "Enabled", "level": "bad", "details": "Telnet server enabled. Disable it."}
    else:
        security_audit["telnet_server"] = {"status": "Default", "level": "warning", "details": "Telnet status unclear (legacy default?). Check 'show telnet'."}
        
    # TFTP
    if check_feature_enabled(r"tftp-server", running_config):
        security_audit["tftp_server"] = {"status": "Enabled", "level": "bad", "details": "TFTP server enabled. Insecure."}
    else:
        security_audit["tftp_server"] = {"status": "Disabled", "level": "good", "details": "TFTP server not enabled."}

    # HTTP/HTTPS
    https_enabled = check_feature_enabled(r"https-server", running_config)
    http_enabled = check_feature_enabled(r"http-server", running_config)
    
    if https_enabled:
        security_audit["https_server"] = {"status": "Enabled", "level": "good", "details": "HTTPS server (web-mgmt) enabled."}
    else:
        security_audit["https_server"] = {"status": "Disabled", "level": "good", "details": "HTTPS server disabled."}
        
    if http_enabled:
        # Check if HTTP is redirected to HTTPS (secure-redirect) - purely speculative command for Aruba, 
        # but generally HTTP should be disabled.
        if check_feature_enabled(r"no http-server", running_config):
             security_audit["http_server"] = {"status": "Disabled", "level": "good", "details": "HTTP server disabled."}
        else:
             security_audit["http_server"] = {"status": "Enabled", "level": "bad", "details": "HTTP server enabled. Insecure."}
    else:
          security_audit["http_server"] = {"status": "Disabled", "level": "good", "details": "HTTP server disabled."}

    # Banners
    banners_set = []
    if check_feature_enabled(r"banner motd", running_config): banners_set.append("MOTD")
    if check_feature_enabled(r"banner exec", running_config): banners_set.append("Exec")
    
    if banners_set:
        security_audit["banners_configured"] = {"status": f"Configured: {', '.join(banners_set)}", "level": "good", "details": "Warning banners configured."}
    else:
        security_audit["banners_configured"] = {"status": "None", "level": "warning", "details": "No warning banners configured."}

    # --- IV. General Hardening ---
    # IP Source Route
    # "no ip source-route" is GOOD.
    # If not present, it might be enabled by default.
    if re.search(r"^\s*no ip source-route", running_config, re.MULTILINE):
         security_audit["ip_source_route"] = {"status": "Disabled", "level": "good", "details": "IP Source Routing disabled."}
    else:
         security_audit["ip_source_route"] = {"status": "Enabled/Default", "level": "bad", "details": "IP Source Routing not explicitly disabled."}

    # LLDP
    if re.search(r"^\s*no lldp run", running_config, re.MULTILINE):
        security_audit["lldp_global_status"] = {"status": "Globally Disabled", "level": "good", "details": "LLDP disabled."}
    else:
        security_audit["lldp_global_status"] = {"status": "Globally Enabled", "level": "warning", "details": "LLDP globally enabled. Secure untrusted ports."}

    # --- V. Logging & Monitoring ---
    if check_feature_enabled(r"logging \S+", running_config): # Generic check for logging host/syslog
        security_audit["remote_logging"] = {"status": True, "level": "good", "details": "Remote logging (syslog) configured."}
    else:
        security_audit["remote_logging"] = {"status": False, "level": "bad", "details": "Remote logging (syslog) NOT configured."}

    # NTP
    ntp_lines = re.findall(r"^\s*ntp server", running_config, re.MULTILINE)
    if len(ntp_lines) >= 2:
        security_audit["ntp_redundancy"] = {"status": f"{len(ntp_lines)} servers", "level": "good", "details": "NTP redundancy OK."}
    elif len(ntp_lines) == 1:
        security_audit["ntp_redundancy"] = {"status": "1 server", "level": "warning", "details": "Only one NTP server configured."}
    else:
        security_audit["ntp_redundancy"] = {"status": "None", "level": "bad", "details": "No NTP servers configured."}
        
    # NTP Auth (Extra)
    if any("key" in line for line in ntp_lines):
         security_audit["ntp_auth"] = {"status": "Configured", "level": "good", "details": "NTP authentication seems used."}
    else:
         security_audit["ntp_auth"] = {"status": "Not Configured", "level": "warning", "details": "NTP authentication not detected."}

    # SNMP
    if check_feature_enabled(r"snmp-server community public", running_config) or check_feature_enabled(r"snmp-server community private", running_config):
        security_audit["snmp_default_communities"] = {"status": "Found", "level": "bad", "details": "Default SNMP communities (public/private) present."}
    else:
        security_audit["snmp_default_communities"] = {"status": "Not Found", "level": "good", "details": "No default public/private communities found."}

    snmpv3_configured = check_feature_enabled(r"snmp-server user", running_config)
    if snmpv3_configured:
        security_audit["snmp_version"] = {"status": "SNMPv3", "level": "good", "details": "SNMPv3 users configured."}
    elif check_feature_enabled(r"snmp-server community", running_config):
        security_audit["snmp_version"] = {"status": "SNMPv1/v2c", "level": "warning", "details": "Only SNMPv1/v2c communities found. Prefer v3."}
    else:
        security_audit["snmp_version"] = {"status": "None", "level": "good", "details": "SNMP not configured."}

    # --- VI. Layer 2 Security ---
    if check_feature_enabled(r"bpdu-protection", running_config) or check_feature_enabled(r"spanning-tree \S+ bpdu-protection", running_config):
        security_audit["bpdu_protection"] = {"status": "Active", "level": "good", "details": "BPDU protection active globally or on ports."}
    else:
        security_audit["bpdu_protection"] = {"status": "Inactive", "level": "warning", "details": "BPDU protection not found."}

    if check_feature_enabled(r"dhcp-snooping", running_config):
         security_audit["dhcp_snooping"] = {"status": "Active", "level": "good", "details": "DHCP Snooping configured."}
    else:
         security_audit["dhcp_snooping"] = {"status": "Inactive", "level": "bad", "details": "DHCP Snooping not configured."}

    return security_audit
