import os
from flask import Flask, send_file, request
from openpyxl import Workbook
import tempfile

app = Flask(__name__)

@app.route('/generate-excel', methods=['POST'])
def generate_excel():
    data = request.json.get('rows', [])
    filename = request.json.get('filename', 'report.xlsx')

    if not filename.endswith('.xlsx'):
        filename += '.xlsx'

    wb = Workbook()
    ws = wb.active
    for row in data:
        ws.append(row)

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
    wb.save(tmp.name)
    tmp.seek(0)

    response = send_file(tmp.name, download_name=filename, as_attachment=True)

    @response.call_on_close
    def cleanup():
        os.remove(tmp.name)

    return response

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 8000))
    app.run(host='0.0.0.0', port=port)
