from flask import Flask, request, render_template, redirect, url_for
import openpyxl
from datetime import datetime

app = Flask(__name__)

# Excel setup

wb = openpyxl.load_workbook('./attendance.xlsx')
sheet = wb.active

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit():
    data = request.form['data']
    name, reg_no = data.split(", ")
    name = name.split(": ")[1]
    reg_no = reg_no.split(": ")[1]

    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    sheet.append([name, reg_no, timestamp])
    wb.save('attendance.xlsx')

    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)