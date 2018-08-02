from flask import Flask, render_template, request, redirect, Response, send_from_directory
from openpyxl import Workbook, load_workbook
from io import BytesIO
from flask_cors import CORS, cross_origin

wb = Workbook()
ws = wb.active

app = Flask(__name__)
CORS(app)




@app.route('/')
def index():
    return "Hello World"

@app.route('/api/fifo', methods=['POST'])
def api():

    print('ping')
    file = request.files['data']

    wb_read = load_workbook(filename=BytesIO(file.read()))
    ws_read = wb_read['Sheet1']

    total_inventory = []
    total_orders = []
    for row in ws_read.rows:
        a = []
        a.append(row[3].value)
        a.append(row[2].value)
        a.append(row[5].value)
        a.append(row[6].value)
        total_orders.append(a)

    tickers = []
    for i in total_inventory:
        if i[1] not in tickers:
            tickers.append(i[1])
    for i in total_orders:
        if i[1] not in tickers:
            tickers.append(i[1])
    all_data = []
    for ticker in tickers:
        append_inv = []
        append_order = []
        for i in total_inventory:
            if i[1] == ticker:
                append_inv.append(i)
        for i in total_orders:
            if i[1] == ticker:
                append_order.append(i)
        a = {   'ticker' : ticker,
                'inventory' : append_inv,
                'orders': append_order
        }
        all_data.append(a)

        # start the FIFO method
    for data in all_data:
        ws.append(['Position', 'Ticker','Vol','Price','Average_Matched_Price', 'Profit FIFO'])

        # setupdata
        inventory = data['inventory']
        orders = data['orders']

        for order in orders:
            if not len(inventory) != 0 or order[0] == inventory[0][0]:
                inventory.append(order)
                order.extend(['inv stack',0])
                ws.append(order)
            else:
                prices = []
                total_inv = 0
                prices_if = []
                for inv in inventory:
                    total_inv += inv[2]
                    prices_if.append(inv[3])
                if total_inv < order[2]:
                    ws.append([
                                order[0],
                                order[1],
                                total_inv,
                                order[3],
                                "inv stack",
                                0
                             ])
                    if order[0] == "Mua":
                        profit = -(order[2] - total_inv)*(order[3] - sum(prices_if)/len(prices_if))
                    else:
                        profit = (order[2] - total_inv)*(order[3] - sum(prices_if)/len(prices_if))
                    ws.append([
                                order[0],
                                order[1],
                                order[2] - total_inv,
                                order[3],
                                sum(prices_if)/len(prices_if),
                                profit
                            ])
                else:
                    for i in range (0, order[2]):
                        prices.append(inventory[0][3])
                        inventory[0][2] -= 1
                        if inventory[0][2] == 0:
                            inventory.pop(0)
                    if order[0] == "Mua":
                        profit = -(order[2])*(order[3] - sum(prices)/len(prices))
                    else:
                        profit = (order[2])*(order[3] - sum(prices)/len(prices))
                    order.extend([sum(prices)/len(prices), profit  ])
                    ws.append(order)


    # wb.save("Result.xlsx")
    wb.save('Result.xlsx')

    return send_from_directory('./','Result.xlsx', as_attachment=True)






if __name__ == '__main__':
  app.run(debug=True)
