from flask import Flask, send_file, request
from openpyxl import Workbook
import tempfile
import os

app = Flask(__name__)

@app.route('/generate-excel', methods=['POST'])
def generate_excel():
    # Get data and optional filename from JSON
    data = request.json.get('rows', [])
    filename = request.json.get('filename', 'report.xlsx')

    # Ensure the filename ends with .xlsx
    if not filename.endswith('.xlsx'):
        filename += '.xlsx'

    # Create workbook
    wb = Workbook()
    ws = wb.active
    for row in data:
        ws.append(row)

    # Create a temporary file
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
    wb.save(tmp.name)
    tmp.seek(0)

    # Return the file with the requested filename
    response = send_file(tmp.name, download_name=filename, as_attachment=True)

    # Optional: clean up temp file on response close
    @response.call_on_close
    def cleanup():
        os.remove(tmp.name)

    return response

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8000)
