import datetime
import pandas
import openpyxl
import yfinance
import json
import supervised
import email.message
import smtplib
##讀入檔案 & 初始化
msg = ''
now = datetime.date.today()
one_day = datetime.timedelta(days=1)
read_wb = openpyxl.load_workbook('C:/股票/Investment.xlsx',data_only=True,read_only=True)
read_ws = read_wb['股票投資']
write_wb = openpyxl.load_workbook('C:/股票/Investment.xlsx')
write_ws = write_wb['股票投資']
with open('C:/股票/value.json',mode='r',encoding='utf-8') as value:
    value = json.load(value)
with open('C:/股票/stock.json',mode='r',encoding='utf-8') as stock:
    stock = json.load(stock)
line = value['stock_data_line']
if value['first'] == True:
    total_bill = read_ws[f'J{line}'].value
    value['first'] = False
else:
    total_bill = value['total_bill']
# print(total_bill)

##下載股票資料
stock_data = {}
for stock_name in stock.keys():
    if stock[stock_name] == '00888':
        stock_data[stock_name] = yfinance.download(f'{stock[stock_name]}.TWO',period='5d',interval='1d')
    else:
        stock_data[stock_name] = yfinance.download(f'{stock[stock_name]}.TW',period='5d',interval='1d')

##模型預測函式
def prediction():
    model = supervised.automl.AutoML(results_path='C:/股票/stock_model')
    pridict_dic = {}
    data_list = []
    for stock_name in stock_data.keys():
        if stock_name == '永豐台灣ESG':
            pass
        else:
            dic = {}
            dic['stock_id'] = stock[stock_name]
            for x in range(5):
                dic[f'date_{x+1}'] = stock_data[stock_name].index[x]
                dic[f'close_{x+1}'] = stock_data[stock_name]['Close'][x]
                dic[f'volume_{x+1}'] = stock_data[stock_name]['Volume'][x]
                pass
            new = pandas.DataFrame(dic,index=[1])
            data_list.append(new)
            # print(new)
    prediction = pandas.concat(data_list,ignore_index=True)
    prediction.to_csv('C:/股票/1.csv')
    prediction = pandas.read_csv('C:/股票/1.csv',index_col=0)
    predictions = model.predict(prediction)
    # print(predictions)
    x = 0
    for stock_name in stock_data.keys():
        if stock_name == '永豐台灣ESG':
            pass
        else:
            pridict_dic[stock_name] = (predictions[x] - stock_data[stock_name]['Close'][4]) / stock_data[stock_name]['Close'][4]
            x += 1
    return pridict_dic
    
##確定已購買股數及價格{股票名稱:[購買價值,購買股數]}
bought_dic = {}
for a in range(value["stock_data_line"]-2):
    if read_ws[f'D{a+3}'].value in bought_dic.keys():
        if read_ws[f'B{a+3}'].value == 'B':
            bought_dic[read_ws[f'D{a+3}'].value][0] += read_ws[f'E{a+3}'].value*read_ws[f'F{a+3}'].value * 1.001425
            bought_dic[read_ws[f'D{a+3}'].value][1] += read_ws[f'E{a+3}'].value
        else:
            bought_dic[read_ws[f'D{a+3}'].value][1] -= read_ws[f'E{a+3}'].value
            if bought_dic[read_ws[f'D{a+3}'].value][1] == 0:
                bought_dic[read_ws[f'D{a+3}'].value][0] = 0
    else:
        bought_dic[read_ws[f'D{a+3}'].value] = [read_ws[f'E{a+3}'].value * read_ws[f'F{a+3}'].value, read_ws[f'E{a+3}'].value]

##定義賣出函式
def sell(stock_name):
    global msg
    global total_bill
    value['stock_data_line'] += 1
    value['week'] = True
    b = value['stock_data_line']
    price = stock_data[stock_name].loc[now.isoformat(),'Close']
    earn = price * bought_dic[stock_name][1] * (1-0.001425-0.003) - bought_dic[stock_name][0] 
    write_ws[f'A{b}'] = now.isoformat()
    write_ws[f'B{b}'] = 'S'
    write_ws[f'C{b}'] = stock[stock_name]
    write_ws[f'D{b}'] = stock_name
    write_ws[f'E{b}'] = bought_dic[stock_name][1]
    write_ws[f'F{b}'] = price
    if earn >= 0:
        l = '賺'
    else:
        l = '賠'
        earn = -earn
    msg += f'賣出 {stock_name} {stock[stock_name]} 賣出股數 {bought_dic[stock_name][1]} 今日股價 {price} {l}:{earn} \n'
    total_bill += bought_dic[stock_name][1] * price * (1-0.001425-0.003)
    bought_dic[stock_name][1] -= bought_dic[stock_name][1]

##定義買入函式
def buy(stock_name,buy_num,ml = False):
    global msg
    global total_bill
    price = stock_data[stock_name].loc[now.isoformat(),'Close']
    if total_bill - price*buy_num*1.001425 > 50000:
        value['stock_data_line'] += 1
        value['week'] = True
        b = value['stock_data_line']
        write_ws[f'A{b}'] = now.isoformat()
        write_ws[f'B{b}'] = 'B'
        write_ws[f'C{b}'] = stock[stock_name]
        write_ws[f'D{b}'] = stock_name
        write_ws[f'E{b}'] = buy_num
        write_ws[f'F{b}'] = price
        msg += f'買入 {stock_name} {stock[stock_name]} 買入股數 {buy_num} 今日股價 {price} 共花 {int((price*buy_num)*1.001425)} \n'
        total_bill -= price*buy_num*1.001425
        if a in bought_dic.keys():
                bought_dic[stock_name][0] += buy_num*price
                bought_dic[stock_name][1] += buy_num
        else:
            bought_dic[stock_name] = [buy_num*price, buy_num] 
    else:
        if ml == True:
            msg += '本周無買賣股票請手動調整'
        else:
            pass

##股票賣出(漲2% or 降1%)
for a in bought_dic.keys():
    if bought_dic[a][1] != 0:
        if (stock_data[a].loc[now.isoformat(),'Close'] * bought_dic[a][1] - bought_dic[a][0]) / bought_dic[a][0] > 0.02:
            sell(a)
        elif (stock_data[a].loc[now.isoformat(),'Close'] * bought_dic[a][1] - bought_dic[a][0]) / bought_dic[a][0] < -0.01:
            sell(a)
            
##購買股票(連兩天漲幅大於0.1%-->買進1張{1000股} or 連四天降福大於0.3%小於1%-->買進0.5張{500股}),若股價大於二十元則買進數量折半
for a in stock.keys():
    cal_day = now
    up = 0
    down = 0
    for x in range(4):
        x1 = (stock_data[a].iloc[4-x, 3] - stock_data[a].iloc[4-(x+1), 3]) / stock_data[a].iloc[x+1, 3]
        if x1 > 0.001:
            if x >= 2:
                pass
            else:
                cal_day -= one_day
                up += 1
        elif x1 < -0.003 and x1 > -0.01:
            cal_day -= one_day
            down += 1
        else:
            break
    if up == 2:
        price = stock_data[a].loc[now.isoformat(),'Close']
        if price >= 20:
            c = 500
        else:
            c = 1000
        buy(a,c)
    elif down == 4:
        price = stock_data[a].loc[now.isoformat(),'Close']
        if price >= 20:
            c = 250
        else:
            c = 500
        buy(a,c)

##一星期沒購買則用機器學習買/賣一支股票(星期三檢查),一股大於20元買500股,小於買1000股
if now.isoweekday() == 3 and value['week'] == False:
    predic_dic = prediction()
    max_key = ''
    max_value = 0
    for key, value in predic_dic.items():
        if value > max_value:
            max_value = value
            max_key = key
    if max_value == 0:
        msg += '這周無買賣請前往解決 \n'
    else:
        if stock_data[max_key].loc[now.isoformat(),'Close'] >= 20:
            c = 500
        else:
            c = 1000
        buy(a,c,ml=True)
elif now.isoweekday() == 3 and value['week'] == True:
    value['week'] = False

if now.isoweekday() == 5:
    import requests
    import bs4
    value['fund_data_line'] += 1
    b = value['fund_data_line']
    found_wb = write_wb['基金投資']
    url = 'https://tw.stock.yahoo.com/fund/summary/F0HKG05X2G:FO'
    r = requests.session()
    text = bs4.BeautifulSoup(r.get(url).text,"html.parser")
    num = text.find('span',attrs="Fz(40px) Fw(b) Lh(1) C($c-primary-text)")
    found_wb[f'B{b}'] = now.isoformat()
    found_wb[f'G{b}'] = float(num.string)
    msg += f'基金交易 價格:{float(num.string)} \n'


##儲存檔案
msg += f'餘額:{total_bill}'
value['total_bill'] = total_bill
with open('C:/股票/value.json',mode='w',encoding='utf-8') as save:
    json.dump(value,save,ensure_ascii=False)
write_wb.save('C:/股票/Investment.xlsx')
mail = email.message.EmailMessage()
mail['From'] = '輸入gmail帳號' # TODO 1
mail['To'] = '輸入gmail帳號' # TODO 2
mail['Subject'] = f'{now} 股票買賣報表'
mail.set_content(msg)
sever = smtplib.SMTP_SSL('smtp.gmail.com',465)
sever.login('輸入gmail帳號','輸入密碼') # TODO 3
sever.send_message(mail)
sever.close()