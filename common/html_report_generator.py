import os
import datetime
import html
import shutil

def generate_html_report(all_data, html_filepath):
    """
    Generates a comprehensive HTML report from the collected audit data,
    formatted for A4 printing with a cover page and logo.
    
    Args:
        all_data (list): A list of dictionaries containing audit data for all devices.
        html_filepath (str): The file path where the HTML report will be saved.
    """
    if not all_data:
        return

    # 1. Setup Resources
    report_dir = os.path.dirname(html_filepath)
    resources_dir = os.path.join(report_dir, "resources")
    if not os.path.exists(resources_dir):
        os.makedirs(resources_dir)
    
    # Copy logo if it exists
    # Assuming the script runs from the project root, the logo is in 'imgs/logomagpie.png'
    project_root = os.getcwd() # Or strictly relative to this file if needed, but cwd is standard here
    logo_src = os.path.join(project_root, "imgs", "logomagpie.png")
    logo_dest = os.path.join(resources_dir, "logo.png")
    
    logo_html = ""
    if os.path.exists(logo_src):
        try:
            shutil.copy2(logo_src, logo_dest)
            # Use relative path for HTML
            logo_html = f'<img src="resources/logo.png" alt="Network Magpie Logo" class="logo">'
        except Exception as e:
            print(f"Warning: Could not copy logo: {e}")

    # 2. Enhanced CSS for A4 and Cover Page
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

        * {
            box-sizing: border-box; 
        }

        body {
            font-family: 'Roboto', sans-serif;
            line-height: 1.4;
            color: #333;
            max-width: 210mm;
            margin: 0 auto;
            background-color: white;
        }
        
        /* A4 Page Setup */
        @page {
            size: A4;
            margin: 15mm;
            @bottom-center {
                content: "Page " counter(page);
                font-size: 9pt;
                color: #7f8c8d;
            }
        }
        
        /* Cover Page Styling */
        .cover-page {
            /* Use a fixed height for print, but flexible for screen */
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            text-align: center;
            page-break-after: always;
            border: 2px solid var(--primary-color);
            border-radius: 10px;
            padding: 40px;
            margin: 20px auto;
            min-height: 250mm; /* Ensure it takes up most of the page */
        }

        .logo {
            max-width: 250px;
            margin-bottom: 30px;
        }

        .report-title {
            font-size: 2.5em;
            color: var(--primary-color);
            margin-bottom: 10px;
        }

        .report-subtitle {
            font-size: 1.4em;
            color: var(--secondary-color);
            margin-bottom: 40px;
        }

        .report-meta {
            font-size: 1.1em;
            color: #7f8c8d;
            margin-top: auto;
            padding-bottom: 20px;
        }

        /* Content Styling */
        h1, h2, h3 {
            color: var(--primary-color);
        }

        h2 {
            border-bottom: 1px solid var(--border-color);
            padding-bottom: 5px;
            margin-top: 20px;
            margin-bottom: 15px;
            page-break-after: avoid;
        }

        h3 {
            margin-top: 15px;
            margin-bottom: 10px;
            page-break-after: avoid;
        }

        .device-section {
            padding: 5px 0;
            margin-bottom: 10px;
        }
        
        img {
            max-width: 100%;
            height: auto;
        }

        table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 15px;
            font-size: 11px;
            page-break-inside: auto;
        }
        
        tr {
            page-break-inside: avoid;
            page-break-after: auto;
        }

        th, td {
            padding: 6px 8px;
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

        @media print {
            body {
                width: 100%;
                margin: 0;
                padding: 0;
            }
            .cover-page {
                height: 250mm; /* Constrain height to avoid overflow */
                max-height: 260mm;
                margin: 0 auto;
                border: none;
                padding-top: 50mm;
            }
            /* Reset margins for headers */
            h2 { margin-top: 10px; }
        }
    </style>
    """

    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    device_count = len(all_data)
    
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
        <div class="cover-page">
            {logo_html}
            <h1 class="report-title">Network Audit Report</h1>
            <div class="report-subtitle">Comprehensive Security & Configuration Analysis</div>
            
            <div class="report-meta">
                <p><strong>Date:</strong> {timestamp}</p>
                <p><strong>Devices Audited:</strong> {device_count}</p>
                <p><strong>Generated By:</strong> Network Magpie</p>
            </div>
        </div>
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
