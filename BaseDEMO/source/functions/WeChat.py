from wxauto import WeChat

import functions.Init

# wx = WeChat()
# print("初始化成功，获取到已登录窗口：", wx.nickname)

# # # for friend in friends:
# # #     print(friend.keys())  # 打印所有字段名称，查看是否有缺失的字段

# # # WXauto 模拟点击获取好友群列表

# friends = wx.GetAllFriends()
# for friend in friends:
#     print(f"好友昵称: {friend.get('nickname')}, 备注: {friend.get('remark', 'tag')}")

##########################################################################################

# from wxpy import *

# # 初始化机器人并登录
# bot = Bot(cache_path=True)

# # 获取所有好友
# friends = bot.friends()
# print("好友列表:")
# for friend in friends:
#     print(friend.nick_name, friend.user_name)

# # 获取所有群聊
# groups = bot.groups()
# print("群组列表:")
# for group in groups:
#     print(group.name, group.user_name)

############################################################################################

# import itchat

# import xml.dom.minidom

# def process_login_info(r):
#     try:
#         return xml.dom.minidom.parseString(r.text)
#     except xml.parsers.expat.ExpatError as e:
#         print(f"XML 解析错误: {e}")
#         print("返回的原始数据:")
#         print(r.text)  # 打印出原始的响应数据
#         return None

# itchat.auto_login(hotReload=True)

# itchat.send('Hello, filehelper', toUserName='filehelper')

# import itchat

# @itchat.msg_register(itchat.content.TEXT)
# def print_content(msg):
#     print(msg['Text'])

# itchat.auto_login()
# itchat.run()

def Auto_send():
    # SendType = functions.Init.SendType
    LogLocation = functions.Init.LogFiles_Path + "\\" + functions.Init.LogFiles
    
    try:
        for i in functions.Init.SendList:
            functions.Init.wx.SendFiles( LogLocation + '\\Tasks.txt', who=functions.Init.friends[i]['remark'])
            print(f"{i} Send Success.")
        
    except Exception as e:
        print(e)

#==============================================================================================================

if __name__ == "__main__":
    functions.Init.init()
    
    # for i in functions.Init.SendList:
    #     functions.Init.wx.SendMsg( msg="RPA 自动化信息收发测试", who=functions.Init.friends[i]['remark'])
        
    # for i in functions.Init.SendList:
    #     functions.Init.wx.SendFiles( r'C:\Users\AdmOA\Downloads\test.txt', who=functions.Init.friends[i]['remark'])
    
    Auto_send()