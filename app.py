from flask import Flask, render_template
import openpyxl
import requests
from io import BytesIO

app = Flask(__name__)

@app.route('/')
def index():
    wb = openpyxl.load_workbook('example.xlsx')
    sheet = wb.active

    items = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        name, price, image_url = row[0], row[3], row[1]
        if isinstance(price, (int, float)):
            response = requests.get(image_url)
            image = BytesIO(response.content)
            items.append({'name': name, 'price': price, 'image': image})

    return render_template('index.html', items=items)


if __name__ == '__main__':
    app.run(debug=True)
