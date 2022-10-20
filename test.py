import requests
from html.parser import HTMLParser
import json
import xlwt
#获取的cookie值存放在这
myHeader = {"Cookie":"SINAGLOBAL=2061253498329.181.1634522467076; UOR=,,www.baidu.com; _ga=GA1.2.639394034.1659491694; __gads=ID=631df5ba423c1782:T=1659491731:S=ALNI_Mb9RCMFImZJgcIJHXS8eYk3yVKmOg; __gpi=UID=000007962ec05ea8:T=1659491731:RT=1659491731:S=ALNI_MYXjjP6D6oUSA_WBKwiZs_YcElTPw; SCF=AkobTsIObPYG38U6_MqlbipiPbjsQGODpxa0VLHqWsDPFQz0JRqboDBu_MhySHLIas5DdjO_4N9SjyBRL1YFVRM.; ULV=1665730328097:54:4:3:4042241409845.393.1665730328094:1665452542654; XSRF-TOKEN=X28mDpEyL--TpYNRNXWQSzI1; SUB=_2A25OSLNlDeRhGedJ4lIW-CnKwz2IHXVtP6OtrDV8PUNbmtANLRLkkW9NUbS_rFB8azzXtsdez9jRruH61J-iZu7k; SUBP=0033WrSXqPxfM725Ws9jqgMF55529P9D9WhhS-zx5GeJZomy8Cu4pKRG5JpX5KzhUgL.Fo2N1K5N1hMc1h22dJLoIEBLxKML1KBL1KnLxKnL1hzL1heLxKBLBo.L12zLxKMLB-2L1KMt; ALF=1697511093; SSOLoginState=1665975093; WBPSESS=EJ5-8oyr00qcxJZSvf4K_6CrMAfGUAOYGocZKX4b2I3fIVpHmjRhN_OD3OMsmaLK7b2qlVO9kMV35FgnI-QyeudM-Fp3EkjEkg4q8AA3tDD-_ScuKp4leCXVedp4Ia_rAx5TPXJv1o3B6Xy_EUaLyA=="}
#要爬去的账号的粉丝列表页面的地址
r = requests.get('https://weibo.com/ajax/friendships/friends?uid=1790787681&relate=fans&count=20&page=10&type=fans&fansSortType=fansCount',headers=myHeader)
f = open("test.html", "w", encoding="UTF-8")
parser = HTMLParser()
parser.feed(r.text)
htmlStr = r.text
dataFrom = json.loads(htmlStr)['users']
# print(dataFrom)
book = xlwt.Workbook(encoding='utf-8', style_compression=0)
sheet = book.add_sheet('关注大V', cell_overwrite_ok=True)
col = ('序号', '微博名', '粉丝数', '介绍')
for k in range(0,4):
    sheet.write(0,k,col[k])
for i in range(len(dataFrom)):
  data = dataFrom[i]
  for j in range(0, 4):
    sheet.write(i + 1, 0, i+1)
    sheet.write(i + 1, 1, data['screen_name'])
    sheet.write(i + 1, 2, data['followers_count_str'])
    sheet.write(i + 1, 3, data['verified_reason'])
savepath = './fansesd.xls'
book.save(savepath)

