from wxauto import WeChat

#----------------------------

from functions import Init

#===============================

def GetAllFriends():
    # 获取当前微信信息
    wx = WeChat()

    # 获取所有好友
    Friends = wx.GetAllFriends()
    for friend in Friends:
        print(f"好友昵称: {friend.get('nickname')}, 备注: {friend.get('remark', 'tag')}")

    # 计算数量
    Friends_num = len(Friends)

    return Friends_num, Friends

def AutoSend( path: str, Friends: list):
    # # 定义变量存储发送文本内容
    # Msg = ""
    #
    # # 读取文件内容
    # try:
    #     with open( path, 'r', encoding='utf-8') as file:  # 打开文件，指定编码为 utf-8
    #         for line in file:  # 逐行读取文件
    #             Msg += line  # 将每一行追加到字符串中
    #
    # except FileNotFoundError:
    #     print(f"文件未找到: {path}")
    # except Exception as e:
    #     print(f"发生错误: {e}")

    for friend in Friends:
        # if friend['type'] > 2:

        # 发送信息（文件/文本）
        try:
            # pass
            # Init.wx.SendMsg()
            Init.wx.SendFiles( filepath=path, who=friend['friend'])

        except Exception as e:
            print(f"发生错误: {e}")
    # pass

if __name__ == '__main__':
    friends_num, friends = GetAllFriends()
    print(f"好友数量: {friends_num}")
