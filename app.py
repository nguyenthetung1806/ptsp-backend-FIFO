from flask import Flask, render_template, request, redirect, Response, send_from_directory
from openpyxl import Workbook, load_workbook
from io import BytesIO
import pandas as pd
from flask_cors import CORS, cross_origin

wb = Workbook()
ws = wb.active

app = Flask(__name__)
CORS(app)




@app.route('/')
def index():
    return "hello"

@app.route('/api/fifo', methods=['POST'])
def api():

    print('ping')
    file_req = request.files['data']


    writer = pd.ExcelWriter('example_result.xlsx', engine='xlsxwriter')


    file = filename=BytesIO(file_req.read())

    xl = pd.ExcelFile(file)

    # Load a sheet into a DataFrame by name: df1
    df0 = xl.parse(xl.sheet_names[0])

    df1 = df0[['SYMBOL','SIDE','QTTY','PRICE']].copy()
    df_sell = df1.loc[df1['SIDE'] == 'S']
    df_buy = df1.loc[df1['SIDE'] == 'B']

    df1["TotalPrice"] = 0

    inventory = []
    inventory_side =  df1.iloc[0]['SIDE']
    for index, row in df1.iterrows():
        for i in range(row['QTTY']):
            if inventory_side == 'No_side':
                inventory_side = row['SIDE']
            if inventory_side == (row['SIDE']):
                inventory.append( row['PRICE'] )
            else:
                total_price_last = inventory[0] + df1.loc[index, 'TotalPrice']
                df1.loc[index, 'TotalPrice'] = total_price_last
                del inventory[0]
                if len(inventory) == 0:
                    inventory_side = 'No_side'

    df1['Average Matched Price'] = df1['TotalPrice'] / df1['QTTY']


    # Write your DataFrame to a file
    df1.to_excel(writer, 'Sheet1')
    # Save the result
    writer.save()

    return send_from_directory('./','example_result.xlsx', as_attachment=True)






if __name__ == '__main__':
  app.run(debug=True)
