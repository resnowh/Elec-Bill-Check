import json
import requests
from dingtalkchatbot.chatbot import DingtalkChatbot

limit = 20  # 欠费预警阈值
room = '3S527'  # 房间号
room_id = '300352711'  # 电费查询号码
# 机器人参数，webhook 和 secret，使用时在PC端创建机器人后可以查看并替换成自己的
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
        # 在这里处理响应数据，解析电费信息等
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
remaining_amount = data['query_elec_roominfo']['errmsg'].split('剩余金额:')[1]
# remaining_amount = '2'
# 输出剩余金额信息
print("剩余金额:", remaining_amount)

xiaoding = DingtalkChatbot(webhook, secret=secret)
xiaoding.send_text(msg='【电费】' + room + ' 目前剩余金额' + remaining_amount + '元')
if float(remaining_amount) < limit:
    xiaoding.send_text(msg='⚠️ ' + room + ' 宿舍用电即将欠费，请尽快充值', is_at_all=True)

print("电费检查程序结束")
