# This is a sample Python script.
# coding:utf-8
import xlrd
import requests
import xlwt
import datetime
current_time = datetime.datetime.now()

data = xlrd.open_workbook('ytb.xlsx')
table = data.sheets()[0]
nrows = table.nrows
current_time = datetime.datetime.now()
for i in range(nrows):
    if i == 0:
        continue
    print(table.row_values(i)[:6][0])
    print(table.row_values(i)[:6][2])
    print(table.row_values(i)[:6][3])
    print(table.row_values(i)[:6][4])
    print(table.row_values(i)[:6][5])
    datalist = [[str(current_time), table.row_values(i)[:6][0], '成功'], [str(current_time), table.row_values(i)[:6][0], '成功']]
    url = 'https://ytb.xian-industrycloud.com/sxdxytb/formParser?status=update&formid=4a1d22b4-8647-4d5c-82c9-bca3f4fa&workflowAction=none&workitemid=&process='
    h = {
        'Content-Type': 'application/json;charset=UTF-8',
        'Accept': 'application/json, text/javascript, */*; q=0.01',
        'Origin': 'https://ytb.xian-industrycloud.com',
        'Referer': 'https://ytb.xian-industrycloud.com/sxdxytb/formParser?showTab=true&submitter=oJbJN53dULS-ZuDr06kNtDr6ymag&validateType=wechatValidate&assignment_id=74a5968d-9384-4a1e-8fef-cda197dc034b&formid=4a1d22b4-8647-4d5c-82c9-bca3f4fa&submitter=oJbJN53dULS-ZuDr06kNtDr6ymag&validateType=wechatValidate&submitterName=roc',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/53.0.2785.116 Safari/537.36 QBCore/4.0.1326.400 QQBrowser/9.0.2524.400 Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/53.0.2875.116 Safari/537.36 NetType/WIFI MicroMessenger/7.0.20.1781(0x6700143B) WindowsWechat(0x63010200)',
        'X-Requested-With': 'XMLHttpRequest'
    }
    d = {
        "header": {
            "code": 0,
            "message": {
                "title": "",
                "detail": ""
            }
        },
        "body": {
            "dataStores": {
                "variable": {
                    "rowSet": {
                        "primary": [{
                            "name": "SYS_USER",
                            "source": "interface",
                            "type": "string",
                            "value": ""
                        }, {
                            "name": "SYS_UNIT",
                            "source": "interface",
                            "type": "string",
                            "value": ""
                        }, {
                            "name": "SYS_UNIT_PATH",
                            "source": "interface",
                            "type": "string",
                            "value": ""
                        }, {
                            "name": "SYS_DATE",
                            "source": "interface",
                            "type": "date",
                            "value": "2022-10-12 09:05:17"
                        }, {
                            "name": "SYS_ROLE",
                            "source": "interface",
                            "type": "string",
                            "value": ""
                        }],
                        "filter": [],
                        "delete": []
                    },
                    "name": "variable",
                    "pageNumber": 1,
                    "pageSize": 2147483647,
                    "recordCount": 0,
                    "parameters": {}
                },
                "036f7c06-6e4a-49de-96c4-5d63bc18": {
                    "rowSet": {
                        "primary": [{
                            "aa7eHRj": "22063481401618432,云中台",
                            "_t": 3,
                            "aaRbHcy": table.row_values(i)[:6][0],
                            "aacT7Ew": "男",
                            "aasBz7d": "第三方支撑人员（指包含划小承包以外的业务合作伙伴）",
                            "aaR7MTE": table.row_values(i)[:6][2],
                            "aaX9rx2": "绿码",
                            "aasAEmF": "正常",
                            "aaFAHAH": "正常",
                            "aa8mm3J": ",坐标为:34.34127, 108.93984",
                            "aacSAf6": "",
                            "aaR1EJS": ",坐标为:34.34127, 108.93984",
                            "aayaGEw": "",
                            "aaKTBie": ",坐标为:34.34127, 108.93984",
                            "aa8YCbi": "",
                            "aa1Cthd": "否",
                            "aayakCt": "",
                            "aankr7d": "",
                            "aa4iYCe": "否",
                            "aajrGPE": "",
                            "aaH1yE4": "",
                            "aayHsh2": "本人承诺：遵守“非必要，不出市，非必要，不聚集”的规定，出现核酸检测阳性、被通知为密接者、次密接者等情况，要主动、及时地向所在单位和社区报告；主动做好个人防护，按照属地要求做好流调、隔离等自我防控工作。能够主动按照当地政府、防疫部门要求的频次进行核酸检测。保证个人及同住家属做到次次参加，100%按规定核检。如实上报本人、同住人员从重点地区来（返）信息。疫情期间服从公司安排和抗疫要求，填报内容属实，如有违反将承担相应责任。",
                            "aaSE66f": "",
                            "_o": {
                                "aa7eHRj": "null",
                                "aaRbHcy": "null",
                                "aacT7Ew": "null",
                                "aasBz7d": "null",
                                "aaR7MTE": "null",
                                "aaX9rx2": "null",
                                "aasAEmF": "null",
                                "aaFAHAH": "null",
                                "aa8mm3J": "null",
                                "aacSAf6": "null",
                                "aaR1EJS": "null",
                                "aayaGEw": "null",
                                "aaKTBie": "null",
                                "aa8YCbi": "null",
                                "aa1Cthd": "null",
                                "aayakCt": "null",
                                "aankr7d": "null",
                                "aa4iYCe": "null",
                                "aajrGPE": "null",
                                "aaH1yE4": "null",
                                "aayHsh2": "null",
                                "aaSE66f": "null"
                            }
                        }],
                        "filter": [],
                        "delete": []
                    },
                    "name": "036f7c06-6e4a-49de-96c4-5d63bc18",
                    "pageNumber": 1,
                    "pageSize": 2147483647,
                    "recordCount": 1,
                    "rowSetName": "55ff3e04-e6c2-4d55-a49c-d7a3af70",
                    "parameters": {
                        "relatedcontrols": "body_0",
                        "primarykey": "pk_id",
                        "foreignkey": "fk_id",
                        "queryds": "036f7c06-6e4a-49de-96c4-5d63bc18"
                    }
                }
            },
            "parameters": {
                "formid": "4a1d22b4-8647-4d5c-82c9-bca3f4fa",
                "print_settings": "",
                "assignment_id": "74a5968d-9384-4a1e-8fef-cda197dc034b",
                "submitter": table.row_values(i)[:6][3],
                "validateType": "wechatValidate",
                "submitterName": table.row_values(i)[:6][4],
                "vs": table.row_values(i)[:6][5]
            }
        }
    }

    r = requests.post(url, headers=h, json=d)
print(r.text)
book = xlwt.Workbook(encoding='utf-8', style_compression=0)
sheet = book.add_sheet('填报成功', cell_overwrite_ok=True)
col = ('日期', '姓名', '是否成功')

for j in range(0, 3):
    sheet.write(0, j, col[j])
# for k in range(0, nrows-1):
#       data = datalist[k]
#       for x in range(0, 3):
#           sheet.write(k+1, x, data[x])

savepath = 'D:/personal/pushData/success.xls'
book.save(savepath)
