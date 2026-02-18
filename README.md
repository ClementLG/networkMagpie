# NetworkMagpie

[![Python Version](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![Netmiko Version](https://img.shields.io/badge/netmiko-4.5+-blue.svg)](https://github.com/ktbyers/netmiko)
[![License: GPL v3](https://img.shields.io/badge/License-GPLv3-blue.svg)](https://www.gnu.org/licenses/gpl-3.0)

<p align="center">
  <a href="https://www.buymeacoffee.com/clementlg">
    <img src="https://www.buymeacoffee.com/assets/img/custom_images/orange_img.png" alt="Buy Me A Coffee">
  </a>
</p>

![logo](imgs/logomagpie.png)

NetworkMagpie is a robust Python-based automation tool designed to perform comprehensive configuration and security audits on network devices across multiple vendors. Currently supporting **Cisco IOS-XE** and **Aruba OS-CX**, it aims to provide network administrators with clear, actionable reports to ensure compliance, security, and optimal performance of their infrastructure.

## Key Features

-   **Multi-Vendor Support**: Seamlessly audits Cisco IOS/IOS-XE and Aruba OS-CX devices.
-   **Comprehensive Data Collection**:
    -   Detailed interface inventory (physical, virtual, management) with status and IP addressing.
    -   VLAN configuration analysis.
    -   ARP table extraction.
-   **Security Auditing**: Performs extensive checks against best practices, including:
    -   AAA configuration.
    -   SSH and access management.
    -   Password policies and encryption.
    -   Service hardening (HTTP, CDP/LLDP, etc.).
    -   Logging, NTP, and SNMP configuration.
    -   Layer 2 security features.
-   **Flexible Reporting**: Generates reports in multiple formats:
    -   **Excel**: Formatted multi-tab reports with color-coded compliance status (Green/Orange/Red).
    -   **HTML**: printer-friendly reports for easy sharing and review.
    -   **JSON**: Raw data export for integration with other tools.
-   **Configuration Management**: Optional export of running configurations to text files.

## Installation

### Prerequisites

-   Python 3.8 or higher.
-   Git.

### Setup

1.  **Clone the repository**:
    ```bash
    git clone https://github.com/ClementLG/networkMagpie.git
    cd networkMagpie
    ```

2.  **Create a Virtual Environment**:
    It is highly recommended to run NetworkMagpie within a virtual environment to manage dependencies securely.

    *   **Windows**:
        ```powershell
        python -m venv venv
        .\venv\Scripts\activate
        ```

    *   **macOS / Linux**:
        ```bash
        python3 -m venv venv
        source venv/bin/activate
        ```

3.  **Install Dependencies**:
    ```bash
    pip install -r requirements.txt
    ```

## Configuration

Before running the audit, you need to configure your inventory and credentials.

### 1. Inventory File (`inventory.csv`)

Create a file named `inventory.csv` in the root directory. This file defines the devices to be audited.

**Format**: `hostname_or_ip,group_name,device_type`

-   `hostname_or_ip`: IP address or DNS name of the device.
-   `group_name`: Logical group name (used to map credentials in `passwords.csv`).
-   `device_type`: Supported OS type (`cisco_ios` or `aruba_os-cx`).

**Example**:
```csv
hostname_or_ip,group_name,device_type
192.168.1.10,core-switches,cisco_ios
10.0.0.5,access-aruba,aruba_os-cx
switch-01.local,dist-switches,cisco_ios
```

### 2. Passwords File (`passwords.csv`)

Create a file named `passwords.csv` in the root directory. This file maps credentials to the groups defined in the inventory.

**Format**: `group_name,username,password,enable_password`

-   `group_name`: Must match a group name from `inventory.csv`.
-   `username`: Login username.
-   `password`: Login password.
-   `enable_password`: (Optional) Enable secret. Leave empty if not applicable.

**Example**:
```csv
group_name,username,password,enable_password
core-switches,admin,SecretPassword123,EnableSecret456
access-aruba,manager,ArubaPass!,
dist-switches,audit_user,AuditPass2024,SecretKey
```

> **Security Note**: Ensure `passwords.csv` is secured with appropriate file permissions (e.g., restricted read access) and is excluded from version control (add to `.gitignore`).

## Usage

Run the main script to start the audit process. By default, it generates an Excel report.

### Basic Usage

```bash
python main.py
```

### Command Line Arguments

You can customize the output formats and enable configuration backup using command-line arguments.

| Argument | Description | Default |
| :--- | :--- | :--- |
| `-o`, `--output` | Specify output formats. Options: `json`, `excel`, `html`. You can specify multiple formats. | `excel` |
| `-e`, `--export-config` | Export the running configuration of each device to a text file. | Disabled |
| `-h`, `--help` | Show the help message and exit. | |

### Examples

**Generate Excel and HTML reports:**
```bash
python main.py -o excel html
```

**Generate JSON report and export device configurations:**
```bash
python main.py -o json -e
```

**Run everything (Excel, HTML, JSON, and Config Export):**
```bash
python main.py -o excel html json -e
```

## Output

All results are saved in the `audit_reports/` directory:

-   **Excel Reports**: `audit_report_YYYYMMDD_HHMMSS.xlsx` - Comprehensive audit results with color-coded security assessments.
-   **HTML Reports**: `audit_report_YYYYMMDD_HHMMSS.html` - Clean, printable summary of the audit.
-   **JSON Data**: `audit_data_YYYYMMDD_HHMMSS.json` - Full raw data export.
-   **Configurations**: `audit_reports/configs/` - Individual text files for device configurations (if `-e` is used).

## Contributing

Contributions are welcome! If you would like to add support for new vendors (Extreme, Fortinet, etc.) or improve existing audit checks:

1.  Fork the repository.
2.  Create a feature branch (`git checkout -b feature/AmazingFeature`).
3.  Commit your changes (`git commit -m 'Add some AmazingFeature'`).
4.  Push to the branch (`git push origin feature/AmazingFeature`).
5.  Open a Pull Request.

## License

Distributed under the GNU General Public License v3 (GPLv3). See `LICENSE` for more information.
