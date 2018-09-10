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
    writer = pd.ExcelWriter('example_result.xlsx', engine='xlsxwriter')

    req_CF0079 = request.files['CF0079']
    req_OD0024 = request.files['OD0024']

    CF0079 = filename=BytesIO(req_CF0079.read())
    OD0024 = filename=BytesIO(req_OD0024.read())

    writer = pd.ExcelWriter('Result.xlsx', engine='xlsxwriter')

    # Read file Data Môi Giới
    DS_MG = "./Data Môi Giới.xlsx"
    DS_Hang = "./DS Hạng.xlsx"
    xl_DS_MG = pd.ExcelFile(DS_MG)
    xl_DS_Hang = pd.ExcelFile(DS_Hang)
    df_DS_MG = xl_DS_MG.parse(xl_DS_MG.sheet_names[0])
    df_DS_Hang = xl_DS_Hang.parse(xl_DS_Hang.sheet_names[0])

    # Read file CF0079 & OD0024 with dataframe only
    df_CF0079_Sheet1 = pd.read_excel("CF0079.xls", skiprows=[0,1,2,3], usecols=6, sheet_name='Sheet1')
    df_CF0079_Sheet2 = pd.read_excel("CF0079.xls", skipfooter=3, usecols=6, sheet_name='Sheet2', header=None)
    df_CF0079 = pd.DataFrame(np.concatenate(
        [df_CF0079_Sheet1.values, df_CF0079_Sheet2.values]),
                             columns=list(df_CF0079_Sheet1)
    )

    df_OD0024 = pd.read_excel("OD0024.xls",
                              skiprows=[0],
                              skipfooter=2)
    df_OD0024 = df_OD0024[df_OD0024['Thuế '].notnull()]

    # merge OD with CF (Số Tk, MG chính, Hạng)
    df_OD0024 = df_OD0024.merge(df_CF0079[['Số TK', 'MG chính', 'Hạng']],
                                left_on='Số \ntài khoản',
                                right_on='Số TK',
                                how='left')

    # merge OD with DS MG
    df_OD0024 = pd.merge(df_OD0024, df_DS_MG, how='left', on=['MG chính'])

    # classify KH FI to df_drop
    df_OD0024_drop = df_OD0024[
            (df_OD0024['MG chính'].notnull()) & (df_OD0024['Nhóm QL'].isnull())
        ].copy()
    df_OD0024_drop['Hạng(Abv)'] = 'Khác'
    # retain others customer to df_0024
    df_OD0024 = df_OD0024.drop(
        df_OD0024[
            (df_OD0024['MG chính'].notnull()) & (df_OD0024['Nhóm QL'].isnull())
        ].index)

    # merge OD24 with Hạng
    df_OD0024 = pd.merge(df_OD0024, df_DS_Hang, how='left', on=['Hạng'])
    # classify special customer class to df_drop_2
    df_OD0024_drop_1 = df_OD0024[(df_OD0024['Hạng(Abv)'].isnull())].copy()
    df_OD0024_drop_1['Hạng(Abv)'] = 'Khác'

    # retain refined data of Brokerage Division
    df_OD0024 = df_OD0024[df_OD0024['Hạng(Abv)'].notnull()].copy()

    # concate and last refine data
    df_OD0024 = df_OD0024.append([df_OD0024_drop, df_OD0024_drop_1], ignore_index=True)
    df_OD0024 = df_OD0024.drop(['Số TK', 'Hạng', 'Nhóm QL'], axis=1)


    # Write your DataFrame to a file
    df_OD0024.to_excel(writer, 'Sheet1')
    # Save the result
    writer.save()






    return send_from_directory('./','example_result.xlsx', as_attachment=True)






if __name__ == '__main__':
  app.run(debug=True)
