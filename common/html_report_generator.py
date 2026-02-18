import os
import datetime
import html

def generate_html_report(all_data, html_filepath):
    """
    Generates a comprehensive HTML report from the collected audit data,
    formatted for A4 printing.
    
    Args:
        all_data (list): A list of dictionaries containing audit data for all devices.
        html_filepath (str): The file path where the HTML report will be saved.
    """
    if not all_data:
        return

    # Basic CSS for A4 printing
    css = """
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;700&display=swap');
        
        :root {
            --primary-color: #2c3e50;
            --secondary-color: #34495e;
            --accent-color: #3498db;
            --success-color: #27ae60;
            --warning-color: #f39c12;
            --error-color: #c0392b;
            --light-gray: #ecf0f1;
            --border-color: #bdc3c7;
        }

        body {
            font-family: 'Roboto', sans-serif;
            line-height: 1.6;
            color: #333;
            max-width: 210mm; /* A4 width */
            margin: 0 auto;
            padding: 20px;
            background-color: #f9f9f9;
        }

        h1, h2, h3 {
            color: var(--primary-color);
        }

        h1 {
            text-align: center;
            border-bottom: 2px solid var(--accent-color);
            padding-bottom: 10px;
            margin-bottom: 30px;
        }

        h2 {
            border-bottom: 1px solid var(--border-color);
            padding-bottom: 5px;
            margin-top: 40px;
            page-break-after: avoid;
        }

        .device-section {
            background: white;
            border: 1px solid var(--border-color);
            border-radius: 5px;
            padding: 20px;
            margin-bottom: 30px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.05);
            page-break-inside: avoid;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 20px;
            font-size: 12px;
        }

        th, td {
            padding: 8px 12px;
            text-align: left;
            border-bottom: 1px solid var(--light-gray);
        }

        th {
            background-color: var(--secondary-color);
            color: white;
            font-weight: bold;
        }

        tr:nth-child(even) {
            background-color: #f8f9fa;
        }

        .status-up, .status-active, .status-good, .status-secure {
            color: var(--success-color);
            font-weight: bold;
        }

        .status-down, .status-disabled, .status-bad, .status-critical {
            color: var(--error-color);
            font-weight: bold;
        }

        .status-warning, .status-admin-down {
            color: var(--warning-color);
            font-weight: bold;
        }

        .header-info {
            display: flex;
            justify-content: space-between;
            margin-bottom: 20px;
            font-size: 14px;
            color: var(--secondary-color);
        }

        @media print {
            body {
                background: white;
                max-width: 100%;
                padding: 0;
            }
            .device-section {
                box-shadow: none;
                border: none;
                page-break-inside: avoid;
            }
            h2 {
                page-break-before: auto;
            }
            a {
                text-decoration: none;
                color: black;
            }
        }
    </style>
    """

    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    html_content = [f"""
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Network Audit Report - {timestamp}</title>
        {css}
    </head>
    <body>
        <div class="header-info">
            <span><strong>Generated:</strong> {timestamp}</span>
            <span><strong>Network Magpie Audit</strong></span>
        </div>
        <h1>Network Audit Report</h1>
    """]

    for device in all_data:
        host = device.get('host', 'Unknown Host')
        status = device.get('status', 'unknown')
        
        # Determine status class for header
        status_class = "status-good" if status == 'success' else "status-bad"
        
        html_content.append(f"""
        <div class="device-section">
            <h2>Device: {html.escape(host)} <span style="font-size: 0.6em; float: right;" class="{status_class}">[{status.upper()}]</span></h2>
        """)

        if status == 'error_connection':
            error_msg = device.get('error_message', 'Unknown error')
            html_content.append(f"""
            <p style="color: var(--error-color);"><strong>Connection Failed:</strong> {html.escape(error_msg)}</p>
            </div>
            """)
            continue

        # General Info
        gen_info = device.get('general_info', {})
        html_content.append("""
            <h3>General Information</h3>
            <table>
                <tr><th>Property</th><th>Value</th></tr>
        """)
        for key, value in gen_info.items():
             html_content.append(f"<tr><td>{html.escape(key.replace('_', ' ').title())}</td><td>{html.escape(str(value))}</td></tr>")
        html_content.append("</table>")

        # Interfaces
        interfaces = device.get('interfaces', [])
        if interfaces:
            html_content.append("""
                <h3>Interfaces</h3>
                <table>
                    <tr>
                        <th>Name</th>
                        <th>IP Address</th>
                        <th>Status</th>
                        <th>Protocol</th>
                        <th>VLAN</th>
                        <th>Description</th>
                    </tr>
            """)
            for iface in interfaces:
                link_stat = str(iface.get('status_link', '')).lower()
                proto_stat = str(iface.get('status_protocol', '')).lower()
                
                link_class = ""
                if "up" in link_stat or "connected" in link_stat: link_class = "status-up"
                elif "admin" in link_stat: link_class = "status-admin-down"
                else: link_class = "status-down"

                proto_class = "status-up" if "up" in proto_stat else "status-down"

                html_content.append(f"""
                    <tr>
                        <td>{html.escape(iface.get('name', ''))}</td>
                        <td>{html.escape(iface.get('ip_address', ''))}</td>
                        <td class="{link_class}">{html.escape(iface.get('status_link', ''))}</td>
                        <td class="{proto_class}">{html.escape(iface.get('status_protocol', ''))}</td>
                        <td>{html.escape(str(iface.get('vlan', '')))}</td>
                        <td>{html.escape(iface.get('description', ''))}</td>
                    </tr>
                """)
            html_content.append("</table>")

        # Security Audit
        security = device.get('security_audit', {})
        if security:
            html_content.append("""
                <h3>Security Audit</h3>
                <table>
                    <tr>
                        <th>Check</th>
                        <th>Status</th>
                        <th>Level</th>
                        <th>Details</th>
                    </tr>
            """)
            for check_name, check_data in security.items():
                level = check_data.get('level', '').upper()
                level_class = ""
                if level == "CRITICAL" or level == "HIGH": level_class = "status-critical"
                elif level == "WARNING" or level == "MEDIUM": level_class = "status-warning"
                else: level_class = "status-good"

                html_content.append(f"""
                    <tr>
                        <td>{html.escape(check_name.replace('_', ' ').title())}</td>
                        <td class="{level_class}">{html.escape(str(check_data.get('status', '')))}</td>
                        <td>{html.escape(level)}</td>
                        <td>{html.escape(check_data.get('details', ''))}</td>
                    </tr>
                """)
            html_content.append("</table>")
            
        html_content.append("</div>") # End device-section

    html_content.append("""
    </body>
    </html>
    """)

    try:
        with open(html_filepath, 'w', encoding='utf-8') as f:
            f.write("".join(html_content))
    except Exception as e:
        print(f"Error saving HTML report: {e}")
