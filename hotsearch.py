import requests
from html.parser import HTMLParser
import json
import xlwt
import time

import pyautogui
import pyperclip
#获取的cookie值存放在这
myHeader = {"Cookie":"SINAGLOBAL=2061253498329.181.1634522467076; UOR=,,www.baidu.com; _ga=GA1.2.639394034.1659491694; __gads=ID=631df5ba423c1782:T=1659491731:S=ALNI_Mb9RCMFImZJgcIJHXS8eYk3yVKmOg; __gpi=UID=000007962ec05ea8:T=1659491731:RT=1659491731:S=ALNI_MYXjjP6D6oUSA_WBKwiZs_YcElTPw; XSRF-TOKEN=0avbTHutlWbxfwYl8v1B_xH6; ALF=1697677726; SSOLoginState=1666141729; SCF=AkobTsIObPYG38U6_MqlbipiPbjsQGODpxa0VLHqWsDP5Fa8iA0NvL1ZVksClYXb0rCYAK_Z3bi5K5HOMbpWkko.; SUB=_2A25OSz5yDeRhGedJ4lIW-CnKwz2IHXVtISi6rDV8PUNbmtAKLRT-kW9NUbS_rCjbwQm-waYeBH2G6t8phybfXAAu; SUBP=0033WrSXqPxfM725Ws9jqgMF55529P9D9WhhS-zx5GeJZomy8Cu4pKRG5JpX5KzhUgL.Fo2N1K5N1hMc1h22dJLoIEBLxKML1KBL1KnLxKnL1hzL1heLxKBLBo.L12zLxKMLB-2L1KMt; _s_tentry=weibo.com; Apache=3944333526977.104.1666141748853; ULV=1666141748879:56:6:2:3944333526977.104.1666141748853:1665987132553; UPSTREAM-V-WEIBO-COM=35846f552801987f8c1e8f7cec0e2230; WBPSESS=EJ5-8oyr00qcxJZSvf4K_6CrMAfGUAOYGocZKX4b2I3fIVpHmjRhN_OD3OMsmaLK7b2qlVO9kMV35FgnI-QyehYhrl5hJgRhx9CSxxPoXmz9rkuQfYAchUZw5ffsstBwwDYZhvfvVGutEqQfyIFR9Q=="}
#要爬去的账号的粉丝列表页面的地址
r = requests.get('https://weibo.com/ajax/side/hotSearch',headers=myHeader)
f = open("test.html", "w", encoding="UTF-8")
parser = HTMLParser()
parser.feed(r.text)
htmlStr = r.text
dataFrom = json.loads(htmlStr)['data']['realtime']
# print(dataFrom)
book = xlwt.Workbook(encoding='utf-8', style_compression=0)
sheet = book.add_sheet('热点', cell_overwrite_ok=True)
col = ('序号', '热搜话题', '热度值', '类别')
for k in range(0,4):
    sheet.write(0,k,col[k])
for i in range(len(dataFrom)):
  data = dataFrom[i]
  print(data)
  if dataFrom[i]['category'] == '社会新闻':
    def send(msg):
      pyperclip.copy(msg)  # 复制需要发送的内容到粘贴板
      pyautogui.hotkey('ctrl', 'v')  # 模拟键盘 ctrl + v 粘贴内容
      pyautogui.press('enter')  # 发送消息
    def send_msg(friend):
      pyautogui.hotkey('ctrl', 'alt', 'w')
      pyautogui.hotkey('ctrl', 'f')  # 搜索好友
      pyperclip.copy(friend)  # 复制好友昵称到粘贴板
      pyautogui.hotkey('ctrl', 'v')  # 模拟键盘 ctrl + v 粘贴
      time.sleep(1)
      pyautogui.press('enter')  # 回车进入好友消息界面
      msg = dataFrom[i]['word_scheme']
      send(msg)
    if __name__ == '__main__':
      friend_name = "zero9one0"  # 对方用户名称：与微信备注保持一致，尽量使用英文
      send_msg(friend_name)
    print(dataFrom[i]['word_scheme'])
    time.sleep(2)
    pyautogui.hotkey('ctrl', 'alt', 'w')
  for j in range(0, 4):
    sheet.write(i + 1, 0, i+1)
    sheet.write(i + 1, 1, data['word_scheme'])
    sheet.write(i + 1, 2, data['num'])
    sheet.write(i + 1, 3, data['category'])
savepath = './hot.xls'
book.save(savepath)