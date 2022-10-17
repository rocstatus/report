import itchat

# 打开微信二维码登陆图片
itchat.auto_login(True)
# 获取除自己以外的好友信息，update=True字段标识储存微信登陆信息到itchat.pkl文件，后续无需重复验证
friends = itchat.get_friends(update=True)[1:]
sex = {'0': '未设置', '1': '男', '2': '女'}
# 循环分析每个微信好友的信息
for i in friends:
    # 判断微信好友名称是否能解析，个别微信名使用图片等其他文字，导致无法解析
    try:
        print('微信名：' + str(i['NickName']))
    except:
        print('微信名：无法解析')
    print('微信名首拼：' + str(i['PYInitial']))
    print('微信名全拼：' + str(i['PYQuanPin']))
    print('备注名：' + str(i['RemarkName']))
    print('备注名首拼：' + str(i['RemarkPYInitial']))
    print('备注名全拼：' + str(i['RemarkPYQuanPin']))
    print('个性签名：' + str(i['Signature']))
    print('城市：' + str(i['City']))
    print('性别：' + sex[str(i['Sex'])])
    print('省份：' + str(i['Province']))
    print('城市：' + str(i['City']))
    print('-----------------------------------------------')