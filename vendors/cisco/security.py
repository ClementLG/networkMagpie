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

    # --- I. AAA & Authentication ---
    if "aaa new-model" in running_config:
        security_audit["aaa_new_model"] = {"status": True, "level": "good",
                                           "details": "AAA new-model is enabled (prerequisite for TACACS+/RADIUS)."}
    else:
        security_audit["aaa_new_model"] = {"status": False, "level": "bad",
                                           "details": "AAA new-model is not enabled. Crucial for centralized and secure access management."}

    if "enable secret" in running_config:
        security_audit["enable_secret_configured"] = {"status": True, "level": "good",
                                                      "details": "An 'enable secret' is configured (strong hashing)."}
    elif "enable password" in running_config:
        security_audit["enable_secret_configured"] = {"status": "enable password only", "level": "bad",
                                                      "details": "'enable password' is used without 'enable secret'. Vulnerable even if 'service password-encryption' is active."}
    else:
        security_audit["enable_secret_configured"] = {"status": False, "level": "bad",
                                                      "details": "No 'enable secret' or 'enable password' configured. Privileged access unprotected."}

    if "service password-encryption" in running_config:
        security_audit["password_encryption_service"] = {"status": True, "level": "good",
                                                         "details": "Service 'password-encryption' enabled (obfuscates type 7 passwords, but does not secure them strongly)."}
    else:
        security_audit["password_encryption_service"] = {"status": False, "level": "bad",
                                                         "details": "Service 'password-encryption' NOT enabled. Passwords (except 'enable secret') stored in clear text."}

    # --- II. Access Line Security ---
    line_con_config_match = re.search(r"line con 0(.*?)!", running_config, re.DOTALL)
    con_config_text = line_con_config_match.group(1) if line_con_config_match else ""

    if "password" in con_config_text or "login local" in con_config_text or "login authentication" in con_config_text:
        security_audit["console_password"] = {"status": True, "level": "good",
                                              "details": "Console line protected by login method."}
    else:
        security_audit["console_password"] = {"status": False, "level": "bad",
                                              "details": "Console line not protected by password."}

    if "exec-timeout" in con_config_text:
        timeout_match = re.search(r"exec-timeout\s+(\d+)\s*(?:(\d+))?", con_config_text)
        if timeout_match:
            minutes, secondes = int(timeout_match.group(1)), int(timeout_match.group(2) or 0)
            if minutes > 0 or (minutes == 0 and secondes > 0):
                security_audit["console_exec_timeout"] = {"status": f"{minutes}m {secondes}s", "level": "good",
                                                          "details": "Console execution timeout configured."}
            else:
                security_audit["console_exec_timeout"] = {"status": "Disabled (0 0)", "level": "bad",
                                                          "details": "Console execution timeout disabled (0 0)."}
        else:
            security_audit["console_exec_timeout"] = {"status": "Configured (check value)", "level": "warning",
                                                      "details": "Console exec-timeout configured, check value."}
    else:
        security_audit["console_exec_timeout"] = {"status": False, "level": "warning",
                                                  "details": "No console execution timeout. Recommended: 5-15 minutes."}

    if "logging synchronous" in con_config_text:
        security_audit["console_logging_synchronous"] = {"status": True, "level": "good",
                                                         "details": "'logging synchronous' enabled on console."}
    else:
        security_audit["console_logging_synchronous"] = {"status": False, "level": "warning",
                                                         "details": "'logging synchronous' not enabled on console."}

    vty_config_text = ""
    vty_sections = re.findall(r"line vty\s+\d+\s*\d*(.*?)!", running_config, re.DOTALL)
    if vty_sections:
        vty_config_text = "\n".join(vty_sections)
    else:
        vty_config_text = running_config

    try:
        ssh_output_raw = net_connect.send_command("show ip ssh", use_textfsm=True)
        ssh_data = (ssh_output_raw[0] if isinstance(ssh_output_raw, list) and ssh_output_raw else {})
        ssh_version = ssh_data.get('protocol_version', 'N/A')
        if ssh_version == '2.0':
            security_audit["ssh_v2_only"] = {"status": True, "level": "good",
                                             "details": f"SSH version {ssh_version} enabled and seems to be the only version."}
        elif ssh_version != 'N/A':
            security_audit["ssh_v2_only"] = {"status": False, "level": "bad",
                                             "details": f"SSH version {ssh_version} detected. SSHv1 is enabled and vulnerable."}
        else:
            if "ip ssh version 2" in running_config and "no ip ssh version 1" in running_config:
                security_audit["ssh_v2_only"] = {"status": True, "level": "good",
                                                 "details": "SSH v2 explicitly configured, v1 disabled (config)."}
            elif "ip ssh version 2" in running_config:
                security_audit["ssh_v2_only"] = {"status": "v2 (v1 status unknown)", "level": "warning",
                                                 "details": "SSH v2 configured, but SSHv1 might still be active. Add 'no ip ssh version 1'."}
            elif "crypto key generate rsa" in running_config:
                security_audit["ssh_v2_only"] = {"status": "Unknown version", "level": "warning",
                                                 "details": "SSH seems enabled (RSA keys), version not confirmed v2 only."}
            else:
                security_audit["ssh_v2_only"] = {"status": False, "level": "bad",
                                                 "details": "SSH does not seem to be enabled or configured correctly."}
    except Exception as e:
        logger.warning(f"SSH check error for {net_connect.host}: {e}")
        security_audit["ssh_v2_only"] = {"status": "Error",
                                         "level": "warning",
                                         "details": "SSH check impossible."}

    if "transport input telnet" in vty_config_text.lower():
        security_audit["vty_transport_telnet"] = {"status": "Telnet Enabled", "level": "bad",
                                                  "details": "Telnet allowed on VTY lines. Unencrypted protocol, disable it."}
    elif "transport input ssh" in vty_config_text:
        security_audit["vty_transport_telnet"] = {"status": "SSH Only (Telnet not found)", "level": "good",
                                                  "details": "VTY lines seem configured for SSH only."}
    else:
        security_audit["vty_transport_telnet"] = {"status": "Transport Unclear", "level": "warning",
                                                  "details": "Non-standard VTY transport config. Ensure SSH only."}

    if "access-class" in vty_config_text and re.search(r"access-class\s+\S+\s+in", vty_config_text):
        security_audit["vty_acl"] = {"status": True, "level": "good",
                                     "details": "ACL (access-class) applied inbound on VTY lines."}
    else:
        security_audit["vty_acl"] = {"status": False, "level": "bad",
                                     "details": "No inbound ACL (access-class) on VTY lines. Risk of unfiltered access."}

    if "exec-timeout" in vty_config_text:
        vty_timeout_match = re.search(r"exec-timeout\s+(\d+)\s*(?:(\d+))?", vty_config_text)
        if vty_timeout_match:
            minutes, secondes = int(vty_timeout_match.group(1)), int(vty_timeout_match.group(2) or 0)
            if minutes > 0 or (minutes == 0 and secondes > 0):
                security_audit["vty_exec_timeout"] = {"status": f"{minutes}m {secondes}s", "level": "good",
                                                      "details": "VTY execution timeout configured."}
            else:
                security_audit["vty_exec_timeout"] = {"status": "Disabled (0 0)", "level": "bad",
                                                      "details": "VTY execution timeout disabled (0 0)."}
        else:
            security_audit["vty_exec_timeout"] = {"status": "Configured (check value)", "level": "warning",
                                                  "details": "VTY exec-timeout configured, check value."}
    else:
        security_audit["vty_exec_timeout"] = {"status": False, "level": "warning",
                                              "details": "No VTY execution timeout. Recommended: 5-15 minutes."}

    if "logging synchronous" in vty_config_text:
        security_audit["vty_logging_synchronous"] = {"status": True, "level": "good",
                                                     "details": "'logging synchronous' enabled on VTY."}
    else:
        security_audit["vty_logging_synchronous"] = {"status": False, "level": "warning",
                                                     "details": "'logging synchronous' not enabled on VTY."}

    # --- III. General Hardening ---
    if "no ip source-route" in running_config:
        security_audit["ip_source_route"] = {"status": "Disabled", "level": "good",
                                             "details": "Source routing disabled."}
    else:
        security_audit["ip_source_route"] = {"status": "Enabled (default)", "level": "bad",
                                             "details": "Source routing active. Configure 'no ip source-route'."}

    if "no service finger" in running_config or "service finger" not in running_config:
        security_audit["finger_service"] = {"status": "Disabled", "level": "good",
                                            "details": "Finger service disabled."}
    else:
        security_audit["finger_service"] = {"status": "Enabled", "level": "bad",
                                            "details": "Finger service enabled. Configure 'no service finger'."}

    small_servers_explicitly_disabled = "no service tcp-small-servers" in running_config and "no service udp-small-servers" in running_config
    small_servers_explicitly_enabled = "service tcp-small-servers" in running_config or "service udp-small-servers" in running_config
    if small_servers_explicitly_disabled or not small_servers_explicitly_enabled:
        security_audit["small_servers"] = {"status": "Disabled", "level": "good",
                                           "details": "tcp/udp-small-servers services disabled."}
    else:
        security_audit["small_servers"] = {"status": "Enabled", "level": "bad",
                                           "details": "tcp/udp-small-servers enabled. Disable them."}

    if "service timestamps log datetime msec" in running_config and "service timestamps debug datetime msec" in running_config:
        security_audit["service_timestamps"] = {"status": True, "level": "good",
                                                "details": "Precise timestamping (msec) logs/debug enabled."}
    elif "service timestamps log" in running_config:
        security_audit["service_timestamps"] = {"status": "Partial", "level": "warning",
                                                "details": "Partial logs timestamping. Recom: '... debug datetime msec' and '... log datetime msec'."}
    else:
        security_audit["service_timestamps"] = {"status": False, "level": "bad",
                                                "details": "Logs/debug timestamping not enabled. Essential for analysis."}

    if "banner motd" in running_config:
        security_audit["banner_motd"] = {"status": True, "level": "good", "details": "MOTD banner configured."}
    else:
        security_audit["banner_motd"] = {"status": False, "level": "warning", "details": "No MOTD banner."}

    if "no ip http server" in running_config:
        security_audit["http_server"] = {"status": "Disabled", "level": "good",
                                         "details": "HTTP server (insecure) disabled."}
    else:
        security_audit["http_server"] = {"status": "Enabled", "level": "bad",
                                         "details": "HTTP server (insecure) enabled. Use HTTPS or disable."}

    if "ip http secure-server" in running_config:
        security_audit["https_server"] = {"status": "Enabled", "level": "good",
                                          "details": "HTTPS server (secure) enabled."}
    elif "no ip http server" in running_config:
        security_audit["https_server"] = {"status": "Disabled (HTTP also disabled)", "level": "good",
                                          "details": "HTTPS server disabled (HTTP also disabled)."}
    else:
        security_audit["https_server"] = {"status": "Disabled (HTTP is Enabled)", "level": "bad",
                                          "details": "HTTPS server disabled while HTTP is active. Switch to HTTPS."}

    cdp_disabled, lldp_disabled = "no cdp run" in running_config, "no lldp run" in running_config
    security_audit["cdp_status"] = (
        {"status": "Globally Disabled", "level": "good", "details": "CDP globally disabled."} if cdp_disabled else
        {"status": "Globally Enabled", "level": "warning",
         "details": "CDP globally enabled. Filter on untrusted interfaces."})
    if lldp_disabled:
        security_audit["lldp_status"] = {"status": "Globally Disabled", "level": "good",
                                         "details": "LLDP globally disabled."}
    elif "lldp run" in running_config:
        security_audit["lldp_status"] = {"status": "Globally Enabled", "level": "warning",
                                         "details": "LLDP globally enabled. Filter on untrusted interfaces."}
    else:
        security_audit["lldp_status"] = {"status": "Potentially Enabled (default)", "level": "warning",
                                         "details": "LLDP potentially active by default."}

    try:
        int_status_raw = net_connect.send_command("show interfaces status", use_textfsm=True)
        int_status = int_status_raw if isinstance(int_status_raw, list) else []
        unused_ports_details, admin_down_ports_details = [], []
        for i_data in int_status:
            port_name, port_status = i_data.get('port'), i_data.get('status', '').lower()
            if not port_name or any(port_name.lower().startswith(p) for p in virtual_prefixes): continue
            if port_status in ['notconnect', 'disabled']: unused_ports_details.append(port_name)
            if port_status == 'disabled': admin_down_ports_details.append(port_name)
        active_unused = [p for p in unused_ports_details if p not in admin_down_ports_details]
        if not active_unused and not unused_ports_details:
            security_audit["unused_physical_ports"] = {"status": "All connected or N/A", "level": "good",
                                                       "details": "All physical ports are connected or status not detailed."}
        elif not active_unused:
            security_audit["unused_physical_ports"] = {"status": f"{len(admin_down_ports_details)} disabled",
                                                       "level": "good",
                                                       "details": f"Unused and disabled physical ports: {', '.join(admin_down_ports_details)}."}
        else:
            security_audit["unused_physical_ports"] = {"status": f"{len(active_unused)} active not connected",
                                                       "level": "bad",
                                                       "details": f"Physical ports not connected but active: {', '.join(active_unused)}. Risk. Shutdown them."}
    except Exception as e:
        logger.warning(f"Error unused_ports check for {net_connect.host}: {e}")
        security_audit["unused_physical_ports"] = {
            "status": "Error", "level": "warning", "details": f"Unused ports check impossible: {e}"}

    # --- IV. Logging & Monitoring ---
    if "logging buffered" in running_config:
        buflvl_match = re.search(r"logging buffered\s+(?:\d+|\S+)", running_config)
        if buflvl_match:
            security_audit["logging_buffered"] = {"status": f"Enabled ({buflvl_match.group(0).split()[-1]})",
                                                  "level": "good", "details": "Buffered logging enabled."}
        else:
            security_audit["logging_buffered"] = {"status": "Enabled (defaults)", "level": "good",
                                                  "details": "Buffered logging enabled (defaults)."}
    else:
        security_audit["logging_buffered"] = {"status": False, "level": "warning",
                                              "details": "Buffered logging not enabled."}

    if "logging host" in running_config or "logging server" in running_config:
        security_audit["remote_logging_configured"] = {"status": True, "level": "good",
                                                       "details": "Remote logging (syslog) configured."}
        if "logging source-interface" in running_config:
            src_int_match = re.search(r"logging source-interface\s+(\S+)", running_config)
            security_audit["logging_source_interface"] = {
                "status": src_int_match.group(1) if src_int_match else "Configured", "level": "good",
                "details": "Syslog source interface specified."}
        else:
            security_audit["logging_source_interface"] = {"status": False, "level": "warning",
                                                          "details": "No syslog source interface."}
    else:
        security_audit["remote_logging_configured"] = {"status": False, "level": "bad",
                                                       "details": "Remote logging (syslog) NOT configured."}
        security_audit["logging_source_interface"] = {"status": "N/A", "level": "bad",
                                                      "details": "Syslog not configured."}

    ntp_servers = [line for line in running_config.splitlines() if
                   "ntp server " in line.strip() and not line.strip().startswith("ntp server vrf")]
    num_ntp_servers = len(ntp_servers)
    ntp_sync_status, ntp_sync_details, ntp_sync_level = "Error", "NTP sync status unknown.", "warning"
    try:
        ntp_status_raw = net_connect.send_command("show ntp status", use_textfsm=True)
        ntp_data = (ntp_status_raw[0] if isinstance(ntp_status_raw, list) and ntp_status_raw else {})
        clock_state = ntp_data.get('clock_state', '').lower()
        if "synchronised" in clock_state or "synchronized" in clock_state:
            ntp_sync_status, ntp_sync_details, ntp_sync_level = True, f"NTP synchronized. Stratum: {ntp_data.get('stratum', 'N/A')}, Ref server: {ntp_data.get('reference_server', 'N/A')}.", "good"
        else:
            ntp_sync_status, ntp_sync_details = False, f"NTP not synchronized (state: {clock_state}). Logs incorrectly timestamped."
            ntp_sync_level = "bad"
    except Exception:
        ntp_sync_details = (
            "NTP configured, sync status unknown ('show ntp status' failed)." if num_ntp_servers > 0 else
            "NTP not configured and status unknown.")
        ntp_sync_level = "warning" if num_ntp_servers > 0 else "bad"
    security_audit["ntp_synchronization"] = {"status": ntp_sync_status, "level": ntp_sync_level,
                                             "details": ntp_sync_details}

    if num_ntp_servers >= 2:
        security_audit["ntp_redundancy"] = {"status": f"{num_ntp_servers} servers", "level": "good",
                                            "details": "NTP redundancy OK."}
    elif num_ntp_servers == 1:
        security_audit["ntp_redundancy"] = {"status": "1 server", "level": "warning",
                                            "details": "Only one NTP server. Recom: >=2."}
    else:
        security_audit["ntp_redundancy"] = {"status": "0 servers", "level": "bad",
                                            "details": "No NTP server configured. Unreliable time."}

    # --- V. Layer 2 Security (Indications) ---
    if "switchport port-security" in running_config:
        security_audit["port_security_feature_used"] = {"status": True, "level": "good",
                                                        "details": "Port Security feature used."}
    else:
        security_audit["port_security_feature_used"] = {"status": False, "level": "bad",
                                                        "details": "Port Security not used. Essential on access ports."}

    if "ip dhcp snooping" in running_config:
        security_audit["dhcp_snooping_global"] = {"status": True, "level": "good",
                                                  "details": "DHCP Snooping enabled globally."}
    else:
        security_audit["dhcp_snooping_global"] = {"status": False, "level": "bad",
                                                  "details": "DHCP Snooping not enabled globally. Required to prevent rogue DHCP servers."}

    if "spanning-tree portfast bpduguard default" in running_config:
        security_audit["bpduguard_default"] = {"status": True, "level": "good",
                                               "details": "BPDU Guard enabled by default on PortFast ports."}
    elif "spanning-tree bpduguard enable" in running_config:
        security_audit["bpduguard_default"] = {"status": "Per-interface (check)", "level": "good",
                                               "details": "BPDU Guard enabled on some interfaces."}
    else:
        security_audit["bpduguard_default"] = {"status": False, "level": "warning",
                                               "details": "BPDU Guard not enabled by default. Risk of STP loops."}

    if "storm-control" in running_config:
        security_audit["storm_control_used"] = {"status": True, "level": "good",
                                                "details": "Storm Control seems configured."}
    else:
        security_audit["storm_control_used"] = {"status": False, "level": "warning",
                                                "details": "Storm Control not used."}

    if "snmp-server community public RO" in running_config or "snmp-server community private RW" in running_config:
        security_audit["snmp_default_communities"] = {"status": True, "level": "bad",
                                                      "details": "SNMP uses DEFAULT communities. Major risk."}
    elif "snmp-server community" in running_config:
        security_audit["snmp_default_communities"] = {"status": False, "level": "good",
                                                      "details": "SNMP configured (no default communities)."}
    else:
        security_audit["snmp_default_communities"] = {"status": "Not Configured", "level": "good",
                                                      "details": "SNMP (v1/v2c community) not configured."}

    if "snmp-server group" in running_config and "v3 auth" in running_config:
        security_audit["snmp_version"] = {"status": "v3 (probable)", "level": "good",
                                          "details": "SNMPv3 seems used."}
    elif "snmp-server community" in running_config:
        security_audit["snmp_version"] = {"status": "v1/v2c", "level": "bad",
                                          "details": "SNMPv1/v2c used (plaintext communities). Prefer SNMPv3."}
    else:
        security_audit["snmp_version"] = {"status": "N/A", "level": "good", "details": "SNMP not configured."}
    return security_audit
