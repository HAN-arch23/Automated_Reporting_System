from netmiko import ConnectHandler
import pandas as pd
from datetime import datetime
import os

# ==============================
# Automated Network Reporting Script
# ==============================

# List of network devices to connect to
devices = [
    {
        'device_type': 'cisco_ios',
        'host': '192.168.1.1',
        'username': 'admin',
        'password': 'admin123',
    },
    {
        'device_type': 'cisco_ios',
        'host': '192.168.1.2',
        'username': 'admin',
        'password': 'admin123',
    }
]

# Store collected data
report_data = []

# Loop through each device and collect information
for device in devices:
    print(f"Connecting to {device['host']}...")
    try:
        connection = ConnectHandler(**device)
        output = connection.send_command("show ip interface brief")
        connection.disconnect()
        report_data.append({
            'Device IP': device['host'],
            'Interface Summary': output
        })
        print(f"✅ Data collected from {device['host']}")
    except Exception as e:
        print(f"⚠️ Failed to connect to {device['host']}: {e}")
        report_data.append({
            'Device IP': device['host'],
            'Interface Summary': 'Connection Failed'
        })

# ==============================
# Save Report
# ==============================

# Automatically create a Desktop folder for reports
save_path = os.path.expanduser("~/Desktop/NetworkReports")
os.makedirs(save_path, exist_ok=True)

# Create a timestamped Excel file
filename = os.path.join(
    save_path,
    f"Network_Report_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
)

# Convert to DataFrame and export to Excel
df = pd.DataFrame(report_data)
df.to_excel(filename, index=False)

print("\n====================================")
print(f"✅ Report generated successfully at:\n{filename}")
print("====================================")