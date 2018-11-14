from flask import Flask, render_template, request, redirect, Response, send_from_directory
from openpyxl import Workbook, load_workbook
from io import BytesIO
import pandas as pd
import numpy as np
from flask_cors import CORS, cross_origin
import urllib.request as urllib
from bs4 import BeautifulSoup
import json

wb = Workbook()
ws = wb.active

app = Flask(__name__)
CORS(app)


@app.route('/')
def index():
    return "hello"

@app.route('/api/take-ticker-list', methods='GET')
def take_ticker_list():
    print('ping')
    floorCode = ['10', '02', '03']

    list_ticker = []

    for code in floorCode:
        response = urllib.urlopen("https://price-fpt-08.vndirect.com.vn/priceservice/secinfo/snapshot/q=floorCode:{}".format(code))
        data = response.read()
        data = json.loads(data)
        data = data[code]

        for data_item in data:
            item_seperator = [pos for pos, char in enumerate(data_item) if char == "|"]
            if (item_seperator[0] + 1) != item_seperator[1]:
                lower_sep = item_seperator[2] + 1
                upper_sep = item_seperator[3]
                ticker = data_item[int(lower_sep): int(upper_sep)]
                list_ticker.append(ticker)

    df = pd.DataFrame(np.array(list_ticker))

    # Specify a writer
    writer = pd.ExcelWriter('Result.xlsx', engine='xlsxwriter')
    # Write your DataFrame to a file
    df.to_excel(writer, 'Sheet1')
    # Save the result
    writer.save()

    return send_from_directory('./','Result.xlsx', as_attachment=True)





if __name__ == '__main__':
  app.run(debug=True)
