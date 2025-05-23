from flask import Flask, request, send_file
from openpyxl import Workbook
import tempfile
import os

app = Flask(__name__)

@app.route('/generate-excel', methods=['POST'])
def generate_excel():
    try:
        data = request.get_json()
        rows = data.get('rows', [])

        # Create a workbook and worksheet
        wb = Workbook()
        ws = wb.active

        # Add the data to the sheet
        for row in rows:
            ws.append(row)

        # Save the file to a temporary location
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            wb.save(tmp.name)
            tmp.seek(0)

            response = send_file(
                tmp.name,
                download_name="generated.xlsx",  # Default file name
                as_attachment=True
            )

        return response

    except Exception as e:
        return {"error": str(e)}, 400

if __name__ == '__main__':
    app.run()
