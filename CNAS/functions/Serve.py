import asyncio
import websockets

import json

#---------------------------

from functions import config
from functions import Init

#=================================

operate ={
    "OperateType": None,
    "command": None
}

response = {
    "type": "status",
    "data": "received successfully"
}

error_response = {
    "type": "status",
    "data": "received FAILED"
}

async def start_operate( test):
    # 格式逻辑测试
    print( test)

# WebSocket Server
class WSS:
    def __init__( self, host="localhost"):
        with open( r'Z:\CPS-OA_RPA-20250312\config\default.json', 'r', encoding='utf-8') as file:
            source = json.load(file)

        self.host = host
        self.port = source['port']
        self.clients = set()                # 存储所有连接的客户端
        self.lock = asyncio.Lock()          # 异步锁防止并发冲突
        self.server = None                  # WebSocket服务器实例
        self.running = False                # 标志服务器是否在运行

        # # 直接用当前事件循环来启动 WebSocket 服务器
        # loop = asyncio.get_event_loop()
        # if loop.is_running():
        #     # 如果事件循环已经运行，直接启动任务
        #     loop.create_task(self.start_server())
        # else:
        #     # 如果没有事件循环，使用 asyncio.run 启动
        #     asyncio.run(self.start_server())

    async def handle_client(self, websocket):
        # 新的客户端连接
        self.clients.add(websocket)

        try:
            async for message in websocket:
                print(f"收到消息: {message}")
                # 回显消息给客户端
                # await websocket.send(f"服务器已收到: {message}")

                # 1. 接收消息并解析为 JSON
                data = json.loads(message)

                # 2. 检查是否是退出信号
                if data.get("type") == "signal" and data.get("data") == "exit":
                    print("接收到退出信号，停止服务器...")
                    await self.stop_server()  # 停止服务器
                    return  # 退出当前连接

                # 使用异步锁保护全局变量写入
                async with self.lock:
                    # 提取 OperateType 和 command 字段
                    operate["OperateType"] = data.get("type")
                    operate["command"] = data.get("data")

                start_operate(data)

                # 3. 处理业务逻辑（示例：回显消息）
                await websocket.send( json.dumps(response))

        except json.JSONDecodeError:
            # 处理非 JSON 格式的消息
            await websocket.send(json.dumps(error_response))
        except websockets.ConnectionClosed as e:
            print(f"客户端断开连接: {e}")
        finally:
            # 移除断开的客户端
            self.clients.remove(websocket)

    async def send_client( self, message):
        """向所有连接的客户端发送消息"""
        if not self.clients:
            print("没有客户端连接，无法发送消息")
            return

        disconnected_clients = []

        for client in self.clients:
            try:
                await client.send(message)
            except websockets.ConnectionClosed:
                disconnected_clients.append(client)

        # 清理断开连接的客户端
        for client in disconnected_clients:
            self.clients.remove(client)

    async def stop_server(self):
        """停止 WebSocket 服务器"""
        print("正在停止 WebSocket 服务器...")

        if self.server:
            print("关闭 WebSocket 服务器...")
            self.server.close()  # 关闭 WebSocket 服务器

            # 强制关闭所有客户端连接
            # async with self.lock:
            #     print(f"强制关闭 {len(self.clients)} 个客户端连接...")
            #     for client in self.clients:
            #         try:
            #             await client.close()  # 主动关闭客户端连接
            #         except Exception as e:
            #             print(f"关闭客户端连接失败: {e}")

            print("等待所有连接关闭...")
            # 设置超时等待服务器关闭
            try:
                await asyncio.wait_for( self.server.wait_closed(), timeout=1)  # 超时1秒
                print("WebSocket 服务器已关闭")
            except asyncio.TimeoutError:
                print("关闭 WebSocket 服务器超时，强制停止服务器")

            self.running = False

        # 停止事件循环
        # print("停止事件循环...")
        # asyncio.get_event_loop().stop()
        # print("事件循环已停止")

    async def start_server(self):
        """启动服务器并保持运行"""
        print(f"启动WebSocket服务器: ws://{self.host}:{self.port}")
        async with websockets.serve(self.handle_client, self.host, self.port) as server:
            self.server = server
            self.running = True

            # 持续运行直到服务器关闭
            # await asyncio.Future()  # 永久等待，直到被外部取消

            # i = 0
            while self.running:  # 保持服务器运行

                await asyncio.sleep(1)  # 避免CPU占用过高

                # print(i)
                # i = i + 1

async def main():
    server = WSS()
    try:
        await server.start_server()
    except KeyboardInterrupt:
        print("手动中断: 程序终止")
    finally:
        await server.stop_server()

if __name__ == "__main__":
    pass
    # Init.Init()
    # server = WSS()

    asyncio.run(main())