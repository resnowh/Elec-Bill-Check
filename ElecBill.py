# 使用 pyinstaller --onefile ElecBill.py 生成可执行文件
# 利用 windows 任务计划程序创建一个查询电费任务，设置重复周期
import json
import requests
from dingtalkchatbot.chatbot import DingtalkChatbot
from datetime import datetime, timedelta
import openpyxl
from openpyxl import load_workbook

limit = 20  # 欠费预警阈值
room = '3S527'  # 房间号
room_id = '300352711'  # 电费查询号码
difference = 0  # 和昨日用电差额

# 机器人参数,webhook 和 secret,使用时在PC端创建机器人后可以查看并替换成自己的
webhook = 'https://oapi.dingtalk.com/robot/send?access_token=2e1a8e3bf5c77c5d3e341b63494239d537c8a18dc76653f1231bd52dec4bcfb9'
secret = 'SEC8834a56af271fb0246db77347726811e3aa13080b888db6c44204db7ab8c0f93'
# 查询请求数据
data_dict = {
    "query_elec_roominfo": {
        "aid": "0030000000007301",
        "account": "158086",
        "room": {
            "roomid": room_id,
            "room": "roomid"
        },
        "floor": {
            "floorid": "",
            "floor": ""
        },
        "area": {
            "area": "",
            "areaname": ""
        },
        "building": {
            "buildingid": "",
            "building": ""
        },
        "extdata": "info1="
    }
}
# 将字典转换为 JSON 字符串
jsondata = json.dumps(data_dict)

def get_electricity_bill():
    url = "http://172.31.248.26:8988/web/Common/Tsm.html"
    headers = {}
    data = {
        "jsondata": jsondata,
        "funname": "synjones.onecard.query.elec.roominfo",
    }

    response = requests.post(url, headers=headers, data=data)

    if response.status_code == 200:
        # 在这里处理响应数据,解析电费信息等
        print("电费信息获取成功")
        electric_bill = response.text  # 将电费信息保存到变量中
        return electric_bill
    else:
        print("电费信息获取失败")
        return None

# 调用函数获取电费信息
bill = get_electricity_bill()

# 解析电费信息
data = json.loads(bill)
remaining_amount = float(data['query_elec_roominfo']['errmsg'].split('剩余金额:')[1])

# 获取当前日期和时间
current_date = datetime.now().strftime('%Y-%m-%d')
current_time = datetime.now().strftime('%H:%M:%S')

# 输出剩余金额信息
print(f"剩余金额: {remaining_amount:.2f} 元")

# 写入Excel文件
excel_file = 'electricity_log.xlsx'

# 检查Excel文件是否存在,如果不存在则创建一个新的
try:
    wb = load_workbook(excel_file)
except FileNotFoundError:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "ElectricityLog"
    ws.append(['Date', 'Time', 'Remaining Amount'])
    wb.save(excel_file)

# 获取工作表
ws = wb["ElectricityLog"]

# 获取前一天的最后一次查询余额(如果存在)
previous_day = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
previous_amount = None
for row in reversed(list(ws.iter_rows(values_only=True))):
    if row[0] == previous_day:
        previous_amount = row[2]
        break

# 如果当前查询的余额与上一次查询的余额不同,则写入Excel
if ws.max_row == 1 or remaining_amount != ws.cell(row=ws.max_row, column=3).value:
    ws.append([current_date, current_time, remaining_amount])
    wb.save(excel_file)
    print("已将查询结果写入Excel文件")
else:
    print("余额与上一次查询相同,不写入Excel")

# 计算当前查询的余额与前一天最后一次查询的余额的差值(如果存在)
if previous_amount is not None:
    difference = remaining_amount - previous_amount
    print(f"与前一天最后一次查询的余额差值: {difference:.6f} 元")

msgText = "【🔋电费信息】\n"
xiaoding = DingtalkChatbot(webhook, secret=secret)

if difference > 0:
    msgText += f"💰️充电费日\n"
msgText += f"目前剩余电费 {remaining_amount} 元,\n"
msgText += f"较昨日差费额 {difference:.6f} 元。"
xiaoding.send_text(msg=msgText)

if remaining_amount < limit:
    xiaoding.send_text(msg=f'⚠️ 宿舍用电即将欠费,请尽快充值', is_at_all=True)

print("电费检查程序结束")