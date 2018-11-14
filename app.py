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

@app.route('/api/CF44', methods=['POST'])
def CF44():
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


    list =[]
    # specify the url
    quote_pages = [   'http://s.cafef.vn/Lich-su-giao-dich-Symbol-VNINDEX/Trang-1-0-tab-1.chn',
                      'http://s.cafef.vn/Lich-su-giao-dich-UPCOM-INDEX-1.chn',
                      'http://s.cafef.vn/Lich-su-giao-dich-HNX-INDEX-1.chn'
    ]

    for page in quote_pages:
       # query the website and return the html to the variable ‘page’
       page = urllib.urlopen(page)

       # parse the html using beautiful soup and store in variable `soup`
       soup = BeautifulSoup(page, 'html.parser')

       # Take out the <div> of name and get its value
       table = soup.select_one('table:nth-of-type(2)')
       name = table.select_one('tr:nth-of-type(4)')
       list.append(int(name.select_one("td:nth-of-type(6)").text.strip().replace(',','')))
       list.append(int(name.select_one("td:nth-of-type(8)").text.strip().replace(',','')))

    GTGD_Market = 2*sum(list)
    df_OD0024_GTGD = df_OD0024[df_OD0024['Hạng(Abv)'] != "Khác"].copy()
    GTGD_VCBS = df_OD0024_GTGD['Giá trị giao dịch'].sum()
    GTGD_M_exc_VCBS = GTGD_Market - GTGD_VCBS

    df_Market = pd.DataFrame([['Thị trường_exc VCBS', GTGD_M_exc_VCBS, 'Thị trường']], columns=['Tên khách hàng','Giá trị giao dịch','Hạng(Abv)'])
    print(df_Market)
    df_OD0024 = df_OD0024.append(df_Market, sort=False, ignore_index=True)
    print(df_OD0024)

    # Specify a writer
    writer = pd.ExcelWriter('Result.xlsx', engine='xlsxwriter')
    # Write your DataFrame to a file
    df_OD0024.to_excel(writer, 'Sheet1')
    # Save the result
    writer.save()

    return send_from_directory('./','Result.xlsx', as_attachment=True)


@app.route('/api/take-ticker-list', methods=['GET'])
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
