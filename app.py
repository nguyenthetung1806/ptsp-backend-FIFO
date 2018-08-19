from flask import Flask, render_template, request, redirect, Response, send_from_directory
from openpyxl import Workbook, load_workbook
from io import BytesIO
import pandas as pd
import numpy as np
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

    df1 = df0[['TIME_EXEC','SYMBOL','SIDE','QTTY','PRICE']].copy()

    #repeat rows
    df1 = df1.loc[df1.index.repeat(df1['QTTY'])]
    #group by index with transform for date ranges
    df1 = df1.sort_values(by='TIME_EXEC')
    #unique default index
    df1 = df1.reset_index(drop=True)
    df1['QTTY'] = 1

    df1["Matched_Price"] = 0
    inventory = []
    inventory_side = df1.iloc[0]['SIDE']
    for index, row in df1.iterrows():
        if inventory_side == 'No_side':
            inventory_side = row['SIDE']
        if inventory_side == (row['SIDE']):
            inventory.append( row['PRICE'] )
        else:
            df1.loc[index, 'Matched_Price'] = inventory[0]
            del inventory[0]
            if len(inventory) == 0:
                inventory_side = 'No_side'

    conditions = [
        (df1['SIDE'] == 'S') & (df1['Matched_Price'] != 0),
        (df1['SIDE'] == 'B') & (df1['Matched_Price'] != 0)]
    choices = [(df1['PRICE'] - df1['Matched_Price']), (df1['Matched_Price'] - df1['PRICE'])]
    df1['Realized_Profit'] = np.select(conditions, choices, default=0)




    # Write your DataFrame to a file
    df1.to_excel(writer, 'Sheet1')
    # Save the result
    writer.save()

    return send_from_directory('./','example_result.xlsx', as_attachment=True)






if __name__ == '__main__':
  app.run(debug=True)
