import re
import logging

logger = logging.getLogger(__name__)

def check_security_features(net_connect, running_config):
    """
    Performs a comprehensive security audit of the Cisco device configuration.

    Args:
        net_connect (netmiko.ConnectHandler): The established Netmiko connection object.
        running_config (str): The full running configuration of the device.

    Returns:
        dict: A dictionary containing the results of various security checks.
    """
    security_audit = {}
    virtual_prefixes = ("vl", "lo", "tu", "po", "bv", "nu", "oo", "gr")

    logger.info(f"Starting security audit for {net_connect.host}...")

    # Helper for positive feature check (must be present and NOT negated)
    def check_feature_enabled(pattern, config):
        # Look for the pattern at the start of a line (ignoring whitespace)
        # We assume if "no <feature>" is present, "feature" line is usually absent or handled.
        # But safest is:
        # 1. Check if "no <feature>" is present -> DISABLED
        # 2. Check if "<feature>" is present -> ENABLED
        
        # NOTE: In Cisco config, "no aaa new-model" explicitly disables it. 
        # "aaa new-model" enables it.
        # So we search for exact match at start of line.
        if re.search(fr"^\s*{pattern}", config, re.MULTILINE):
            return True
        return False

    # --- I. AAA & Authentication ---
    # 1. AAA New-Model
    if check_feature_enabled(r"aaa new-model", running_config):
        security_audit["aaa_new_model"] = {"status": True, "level": "good",
                                           "details": "AAA new-model is enabled."}
    else:
        # Explicit check for "no aaa new-model" just to be precise in detail, but logic is same
        if re.search(r"^\s*no aaa new-model", running_config, re.MULTILINE):
             detail = "AAA new-model is explicitly disabled ('no aaa new-model')."
        else:
             detail = "AAA new-model is not enabled."
        security_audit["aaa_new_model"] = {"status": False, "level": "bad", "details": detail}

    # 2. AAA Authentication Login
    aaa_login_match = re.search(r"^\s*aaa authentication login default\s+(group\s+(tacacs\+|radius)|local)", running_config, re.MULTILINE)
    if aaa_login_match:
        method = aaa_login_match.group(1)
        if "group" in method:
             security_audit["aaa_login_method"] = {"status": "Centralized", "level": "good",
                                                   "details": f"AAA login default uses: {method}."}
        else:
             security_audit["aaa_login_method"] = {"status": "Local", "level": "warning",
                                                   "details": "AAA login default uses local authentication."}
    else:
        if security_audit["aaa_new_model"]["status"]:
             security_audit["aaa_login_method"] = {"status": "Default/None", "level": "warning",
                                                   "details": "No explicit 'aaa authentication login default' found."}
        else:
             security_audit["aaa_login_method"] = {"status": "N/A", "level": "bad",
                                                   "details": "AAA not enabled."}

    # 3. Enable Secret
    if check_feature_enabled(r"enable secret", running_config):
        security_audit["enable_secret_configured"] = {"status": True, "level": "good",
                                                      "details": "An 'enable secret' is configured."}
    elif check_feature_enabled(r"enable password", running_config):
        security_audit["enable_secret_configured"] = {"status": "enable password only", "level": "bad",
                                                      "details": "'enable password' used without 'enable secret'. Vulnerable."}
    else:
        security_audit["enable_secret_configured"] = {"status": False, "level": "bad",
                                                      "details": "No 'enable secret' or 'enable password' configured."}

    # 4. Service Password-Encryption
    if check_feature_enabled(r"service password-encryption", running_config):
        security_audit["password_encryption_service"] = {"status": True, "level": "good",
                                                         "details": "Service 'password-encryption' enabled."}
    else:
        security_audit["password_encryption_service"] = {"status": False, "level": "bad",
                                                         "details": "Service 'password-encryption' NOT enabled."}
    
    # 5. Password Policy
    min_len_match = re.search(r"^\s*security passwords min-length\s+(\d+)", running_config, re.MULTILINE)
    if min_len_match:
        min_len = int(min_len_match.group(1))
        level = "good" if min_len >= 10 else "warning"
        security_audit["password_min_length"] = {"status": f"{min_len} chars", "level": level,
                                                 "details": f"Minimum password length set to {min_len}."}
    else:
         security_audit["password_min_length"] = {"status": "Default", "level": "warning",
                                                  "details": "No 'security passwords min-length' configured."}


    # --- II. Access Line Security ---
    # Console
    line_con_config_match = re.search(r"^line con 0(.*?)!", running_config, re.DOTALL | re.MULTILINE)
    con_config_text = line_con_config_match.group(1) if line_con_config_match else ""

    if "password" in con_config_text or "login local" in con_config_text or "login authentication" in con_config_text:
        security_audit["console_password"] = {"status": True, "level": "good",
                                              "details": "Console line protected."}
    else:
        security_audit["console_password"] = {"status": False, "level": "bad",
                                              "details": "Console line not explicitly protected."}

    # Console Timeout
    timeout_match = re.search(r"exec-timeout\s+(\d+)\s*(?:(\d+))?", con_config_text)
    if timeout_match:
        minutes, secondes = int(timeout_match.group(1)), int(timeout_match.group(2) or 0)
        if minutes > 0 or (minutes == 0 and secondes > 0):
            security_audit["console_exec_timeout"] = {"status": f"{minutes}m {secondes}s", "level": "good",
                                                      "details": "Console execution timeout configured."}
        else:
            security_audit["console_exec_timeout"] = {"status": "Disabled (0 0)", "level": "bad",
                                                      "details": "Console execution timeout disabled."}
    else:
        security_audit["console_exec_timeout"] = {"status": False, "level": "warning",
                                                  "details": "No console execution timeout."}

    # Console Logging Sync
    if "logging synchronous" in con_config_text:
        security_audit["console_logging_synchronous"] = {"status": True, "level": "good", "details": "Enabled."}
    else:
        security_audit["console_logging_synchronous"] = {"status": False, "level": "warning", "details": "Disabled."}

    # VTY
    vty_sections = re.findall(r"^line vty\s+\d+\s*\d*(.*?)!", running_config, re.DOTALL | re.MULTILINE)
    vty_config_text = "\n".join(vty_sections) if vty_sections else running_config

    # SSH Status (Hybrid check)
    ssh_status_known = False
    try:
        ssh_output_raw = net_connect.send_command("show ip ssh", use_textfsm=True)
        ssh_data = (ssh_output_raw[0] if isinstance(ssh_output_raw, list) and ssh_output_raw else {})
        if not ssh_data and isinstance(ssh_output_raw, str):
             if "SSH Enabled - version 2.0" in ssh_output_raw: ssh_data = {'protocol_version': '2.0'}
             elif "SSH Disabled" in ssh_output_raw: ssh_data = {'protocol_version': 'Disabled'}

        ssh_version = ssh_data.get('protocol_version', 'N/A')
        if ssh_version == '2.0':
            security_audit["ssh_v2_only"] = {"status": True, "level": "good", "details": "SSHv2 enabled."}
            ssh_status_known = True
        elif ssh_version == 'Disabled':
             security_audit["ssh_v2_only"] = {"status": "Disabled", "level": "bad", "details": "SSH disabled."}
             ssh_status_known = True
    except Exception:
        pass

    if not ssh_status_known:
         # Fallback to Config Logic
         if check_feature_enabled(r"ip ssh version 2", running_config):
              security_audit["ssh_v2_only"] = {"status": "Configured v2", "level": "good", "details": "Configured in running-config."}
         else:
              security_audit["ssh_v2_only"] = {"status": "Unknown", "level": "warning", "details": "SSH version unclear."}

    # SSH Timeout/Retries
    ssh_timeout = re.search(r"^\s*ip ssh time-out\s+(\d+)", running_config, re.MULTILINE)
    security_audit["ssh_timeout"] = {"status": f"{ssh_timeout.group(1)}s", "level": "good", "details": "Configured."} if ssh_timeout else {"status": "Default", "level": "warning", "details": "Default (120s)."}

    ssh_retries = re.search(r"^\s*ip ssh authentication-retries\s+(\d+)", running_config, re.MULTILINE)
    if ssh_retries:
         r_val = int(ssh_retries.group(1))
         security_audit["ssh_retries"] = {"status": f"{r_val}", "level": "good" if r_val <= 3 else "warning", "details": f"Set to {r_val}."}
    else:
         security_audit["ssh_retries"] = {"status": "Default", "level": "warning", "details": "Default (3)."}

    # VTY Transport
    if "transport input telnet" in vty_config_text.lower() or "transport input all" in vty_config_text.lower():
        security_audit["vty_transport_telnet"] = {"status": "Telnet Enabled", "level": "bad", "details": "Telnet enabled on VTY."}
    elif "transport input ssh" in vty_config_text.lower():
        security_audit["vty_transport_telnet"] = {"status": "SSH Only", "level": "good", "details": "SSH only."}
    else:
        security_audit["vty_transport_telnet"] = {"status": "Unclear", "level": "warning", "details": "Verify VTY transport."}

    # VTY ACL
    if "access-class" in vty_config_text and re.search(r"access-class\s+\S+\s+in", vty_config_text):
        security_audit["vty_acl"] = {"status": True, "level": "good", "details": "Inbound ACL present."}
    else:
        security_audit["vty_acl"] = {"status": False, "level": "bad", "details": "No inbound access-class."}
    
    
    # --- III. General Hardening (Negation Logic Fixes) ---
    
    # IP Source Route
    # We want "no ip source-route" (Good)
    # Default is often Enabled (Bad).
    if re.search(r"^\s*no ip source-route", running_config, re.MULTILINE):
        security_audit["ip_source_route"] = {"status": "Disabled", "level": "good", "details": "Source routing disabled."}
    else:
        security_audit["ip_source_route"] = {"status": "Enabled", "level": "bad", "details": "Source routing active (default)."}

    # Finger
    # We want "no service finger" (Good) or implicit default? usually enabled on old IOS.
    if re.search(r"^\s*no service finger", running_config, re.MULTILINE):
        security_audit["finger_service"] = {"status": "Disabled", "level": "good", "details": "Finger disabled."}
    elif check_feature_enabled(r"service finger", running_config):
        security_audit["finger_service"] = {"status": "Enabled", "level": "bad", "details": "Finger service enabled."}
    else:
        # If neither, assume disabled on modern, but warn on old.
        security_audit["finger_service"] = {"status": "Implicit/Unknown", "level": "good", "details": "Not explicitly enabled."}

    # Small Servers
    msg_small = "TCP/UDP small servers"
    if re.search(r"^\s*no service tcp-small-servers", running_config, re.MULTILINE) and re.search(r"^\s*no service udp-small-servers", running_config, re.MULTILINE):
         security_audit["small_servers"] = {"status": "Disabled", "level": "good", "details": f"{msg_small} disabled."}
    elif check_feature_enabled(r"service (tcp|udp)-small-servers", running_config):
         security_audit["small_servers"] = {"status": "Enabled", "level": "bad", "details": f"{msg_small} enabled. Disable them."}
    else:
         security_audit["small_servers"] = {"status": "Implicit", "level": "warning", "details": "Verify small servers."}

    # PAD
    if re.search(r"^\s*no service pad", running_config, re.MULTILINE):
        security_audit["pad_service"] = {"status": "Disabled", "level": "good", "details": "PAD disabled."}
    elif check_feature_enabled(r"service pad", running_config):
        security_audit["pad_service"] = {"status": "Enabled", "level": "bad", "details": "PAD enabled."}
    else:
        security_audit["pad_service"] = {"status": "Implicit", "level": "warning", "details": "Verify PAD service."}

    # BOOTP
    if re.search(r"^\s*no ip bootp server", running_config, re.MULTILINE):
        security_audit["bootp_server"] = {"status": "Disabled", "level": "good", "details": "BOOTP disabled."}
    else:
        security_audit["bootp_server"] = {"status": "Enabled (default)", "level": "warning", "details": "BOOTP might be active."}

    # HTTP Server
    # We want "no ip http server" (Good)
    if re.search(r"^\s*no ip http server", running_config, re.MULTILINE):
        security_audit["http_server"] = {"status": "Disabled", "level": "good", "details": "HTTP server disabled."}
    elif check_feature_enabled(r"ip http server", running_config):
        security_audit["http_server"] = {"status": "Enabled", "level": "bad", "details": "HTTP server enabled."}
    else:
        security_audit["http_server"] = {"status": "Enabled (default)", "level": "bad", "details": "HTTP server active by default."}

    # HTTPS Server
    if check_feature_enabled(r"ip http secure-server", running_config):
        security_audit["https_server"] = {"status": "Enabled", "level": "good", "details": "HTTPS enabled."}
    else:
        security_audit["https_server"] = {"status": "Disabled", "level": "warning", "details": "HTTPS not enabled."}

    # CDP / LLDP
    # "no cdp run" -> Disabled (Good)
    if re.search(r"^\s*no cdp run", running_config, re.MULTILINE):
        security_audit["cdp_status"] = {"status": "Disabled", "level": "good", "details": "CDP globally disabled."}
    else:
        security_audit["cdp_status"] = {"status": "Enabled", "level": "warning", "details": "CDP globally enabled."}
        
    if re.search(r"^\s*no lldp run", running_config, re.MULTILINE):
        security_audit["lldp_status"] = {"status": "Disabled", "level": "good", "details": "LLDP globally disabled."}
    elif check_feature_enabled(r"lldp run", running_config):
        security_audit["lldp_status"] = {"status": "Enabled", "level": "warning", "details": "LLDP globally enabled."}
    else:
        security_audit["lldp_status"] = {"status": "Disabled (default)", "level": "good", "details": "LLDP disabled by default."}


    # --- IV. Logging ---
    # Timestamps
    if check_feature_enabled(r"service timestamps log datetime msec", running_config) and check_feature_enabled(r"service timestamps debug datetime msec", running_config):
         security_audit["service_timestamps"] = {"status": True, "level": "good", "details": "Precise timestamps enabled."}
    else:
         security_audit["service_timestamps"] = {"status": "Partial", "level": "warning", "details": "Check timestamp settings."}

    # Banners
    if check_feature_enabled(r"banner motd", running_config):
        security_audit["banner_motd"] = {"status": True, "level": "good", "details": "MOTD configured."}
    else:
        security_audit["banner_motd"] = {"status": False, "level": "warning", "details": "No MOTD."}

    # Buffered Logging
    # "logging buffered <level>" or "logging buffered <size>"
    if re.search(r"^\s*logging buffered", running_config, re.MULTILINE):
        security_audit["logging_buffered"] = {"status": "Enabled", "level": "good", "details": "Buffered logging enabled."}
    else:
        security_audit["logging_buffered"] = {"status": "Default", "level": "warning", "details": "Buffered logging not explicitly set."}

    # Syslog
    if re.search(r"^\s*logging host|^\s*logging server", running_config, re.MULTILINE):
        security_audit["remote_logging_configured"] = {"status": True, "level": "good", "details": "Syslog configured."}
    else:
        security_audit["remote_logging_configured"] = {"status": False, "level": "bad", "details": "No syslog configured."}
    
    # NTP
    ntp_lines = re.findall(r"^\s*ntp server", running_config, re.MULTILINE)
    security_audit["ntp_redundancy"] = {"status": f"{len(ntp_lines)} servers", "level": "good" if len(ntp_lines) >= 2 else "warning", "details": "NTP servers."}


    # --- V. SNMP ---
    if check_feature_enabled(r"snmp-server community (public|private)", running_config):
        security_audit["snmp_default_communities"] = {"status": "Found", "level": "bad", "details": "Default communities found."}
    else:
        security_audit["snmp_default_communities"] = {"status": "Clean", "level": "good", "details": "No default communities."}

    # --- VI. Layer 2 ---
    # Unused Ports - Keeping existing logic wrapper
    try:
        int_status_raw = net_connect.send_command("show interfaces status", use_textfsm=True)
        if isinstance(int_status_raw, list) and int_status_raw:
            unused_ports_details = []
            admin_down_ports_details = []
            
            for i_data in int_status:
                port_name = i_data.get('port')
                port_status = i_data.get('status', '').lower()
                
                if not port_name or any(port_name.lower().startswith(p) for p in virtual_prefixes):
                    continue
                
                if port_status in ['notconnect', 'disabled', 'err-disabled']:
                    unused_ports_details.append(port_name)
                
                if port_status == 'disabled':
                    admin_down_ports_details.append(port_name)
            
            # Active unused = unused - admin_down
            active_unused = [p for p in unused_ports_details if p not in admin_down_ports_details]
            
            if not active_unused and not unused_ports_details:
                security_audit["unused_physical_ports"] = {"status": "All connected", "level": "good",
                                                           "details": "All physical ports are connected."}
            elif not active_unused:
                security_audit["unused_physical_ports"] = {"status": f"{len(admin_down_ports_details)} disabled", "level": "good",
                                                           "details": f"All unused ports are administratively down."}
            else:
                security_audit["unused_physical_ports"] = {"status": f"{len(active_unused)} active/unused", "level": "bad",
                                                           "details": f"Checking {len(active_unused)} ports found active but not connected (risk). Shutdown them."}
        else:
             security_audit["unused_physical_ports"] = {"status": "Check", "level": "info", "details": "Run manual check for unused ports."}
    except Exception:
        security_audit["unused_physical_ports"] = {"status": "Error", "level": "warning", "details": "Could not check ports."}

    logger.info(f"Security audit finished for {net_connect.host}.")
    return security_audit
