from flask import Flask, render_template, request, send_from_directory
from netmiko import ConnectHandler
import pandas as pd
from datetime import datetime
from fpdf import FPDF
import os

app = Flask(__name__)

SAVE_DIR = os.path.expanduser("~/Desktop/NetworkReports")
os.makedirs(SAVE_DIR, exist_ok=True)

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate_report():
    devices = []
    ips = request.form.getlist('ip[]')
    usernames = request.form.getlist('username[]')
    passwords = request.form.getlist('password[]')

    for i in range(len(ips)):
        devices.append({
            'device_type': 'cisco_ios',
            'host': ips[i],
            'username': usernames[i],
            'password': passwords[i]
        })

    report_data = []

    for device in devices:
        try:
            connection = ConnectHandler(**device)
            output = connection.send_command("show ip interface brief")
            connection.disconnect()
            report_data.append({'Device IP': device['host'], 'Interface Summary': output})
        except Exception as e:
            report_data.append({'Device IP': device['host'], 'Interface Summary': f'Failed: {e}'})

    # Save Excel
    timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    excel_file = f"Network_Report_{timestamp}.xlsx"
    pdf_file = f"Network_Report_{timestamp}.pdf"

    df = pd.DataFrame(report_data)
    excel_path = os.path.join(SAVE_DIR, excel_file)
    df.to_excel(excel_path, index=False)

    # Create PDF
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(200, 10, "Network Status Report", ln=True, align='C')
    pdf.set_font("Arial", size=12)
    for entry in report_data:
        pdf.cell(0, 10, f"Device: {entry['Device IP']}", ln=True)
        pdf.multi_cell(0, 8, entry['Interface Summary'], border=1)
        pdf.ln(5)
    pdf_path = os.path.join(SAVE_DIR, pdf_file)
    pdf.output(pdf_path)

    return render_template('index.html', success=True, excel_file=excel_file, pdf_file=pdf_file)

@app.route('/download/<filename>')
def download(filename):
    return send_from_directory(SAVE_DIR, filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)