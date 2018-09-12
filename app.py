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

    CF0079 = BytesIO(req_CF0079.read())
    OD0024 = BytesIO(req_OD0024.read())

    writer = pd.ExcelWriter('Result.xlsx', engine='xlsxwriter')

    # Input Data MG, Hang KH
    df_DS_MG = pd.read_excel("Data Môi Giới.xlsx")
    df_DS_Hang = pd.read_excel("DS Hạng.xlsx")


    # Read file CF0079 & OD0024 with dataframe only
    # CF0079
    df_CF0079_Sheet1 = pd.read_excel(CF0079, skiprows=[0,1,2,3], usecols=6, sheet_name='Sheet1')
    df_CF0079_Sheet2 = pd.read_excel(CF0079, skipfooter=3, usecols=6, sheet_name='Sheet2', header=None)
    df_CF0079 = pd.DataFrame(np.concatenate(
        [df_CF0079_Sheet1.values, df_CF0079_Sheet2.values]),
                             columns=list(df_CF0079_Sheet1),
    )
    df_CF0079 = df_CF0079.merge(df_DS_Hang,  how='left', on=['Hạng'])

    # OD0024
    df_OD0024 = pd.read_excel(OD0024,
                              skiprows=[0],
                              skipfooter=2)
    df_OD0024 = df_OD0024[df_OD0024['Thuế '].notnull()]

    # rename to set OD0024 and CF0079's header "Số TK' same variable
    df_OD0024 = df_OD0024.rename(index=str, columns={'Số \ntài khoản': 'Số TK'})
    # merge OD with CF (Số Tk, MG chính, Hạng)
    df_OD0024 = df_OD0024.merge(df_CF0079[['Số TK', 'MG chính', 'Hạng(Abv)', 'Hạng(Short)']], how='left', on=['Số TK'])
    # merge OD with DS MG
    df_OD0024 = pd.merge(df_OD0024, df_DS_MG, how='left', on=['MG chính'])

    # Set that all TK that have no Hạng(Abv) or Hạng(Short) are Khác
    df_OD0024.fillna(value={'Hạng(Abv)': 'Khác', 'Hạng(Short)': 'Khác'}, inplace=True)



    # Clear FI
    df_OD0024['Hạng(Abv)'] = np.where((df_OD0024['MG chính'].notnull()) & (df_OD0024['Nhóm QL'].isnull()), 'Khác', df_OD0024['Hạng(Abv)'])
    df_OD0024['Hạng(Short)'] = np.where((df_OD0024['Hạng(Abv)'] == 'Khác'), 'Khác', df_OD0024['Hạng(Short)'])



    # # Speedup the speed by eliminate those whose inrelevant
    # df_CF0079_TVDT = df_CF0079.dropna(subset=['MG chính']).copy()
    #
    # for i in ['TVĐTMM', 'TVĐTVIP', 'TVĐT']:
    #     checkCondition = lambda row: len(df_CF0079_TVDT[
    #                                               (df_CF0079_TVDT['Hạng(Abv)'] == i) & (df_CF0079_TVDT['MG chính'] == row['MG chính'])
    #                                           ])
    #     head_name = 'Số khách ' + i
    #     df_OD0024[head_name] = df_OD0024.apply(checkCondition, axis=1)


    # Specify a writer
    writer = pd.ExcelWriter('Result.xlsx', engine='xlsxwriter')
    # Write your DataFrame to a file
    df_OD0024.to_excel(writer, 'Sheet1')
    # Save the result
    writer.save()





    return send_from_directory('./','Result.xlsx', as_attachment=True)






if __name__ == '__main__':
  app.run(debug=True)
