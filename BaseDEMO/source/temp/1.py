import asyncio
import json
import websockets

# 假设以下全局变量和函数已定义
operate = {}
def start_operate(data): pass

class WSS:
    def __init__(self, host="localhost"):
        with open(r'Z:\CPS-OA_RPA-20250312\config\default.json', 'r', encoding='utf-8') as file:
            source = json.load(file)

        self.host = host
        self.port = source['port']
        self.clients = set()
        self.lock = asyncio.Lock()
        self.server = None  # 存储WebSocket服务器实例
        self.running = False  # 标志服务器是否在运行

    async def handle_client(self, websocket):
        self.clients.add(websocket)
        try:
            async for message in websocket:
                print(f"收到消息: {message}")
                data = json.loads(message)

                # 处理退出信号
                if data.get("type") == "signal" and data.get("data") == "exit":
                    print("接收到退出信号，停止服务器...")
                    await self.stop_server()
                    return  # 退出当前连接处理

                # 更新操作信息
                async with self.lock:
                    operate["OperateType"] = data.get("type")
                    operate["command"] = data.get("data")

                start_operate(data)
                # 示例响应（需根据实际逻辑修改）
                response = {"status": "success"}
                await websocket.send(json.dumps(response))

        except json.JSONDecodeError:
            error_response = {"error": "无效的JSON格式"}
            await websocket.send(json.dumps(error_response))
        except websockets.ConnectionClosed:
            print("客户端断开连接")
        finally:
            self.clients.remove(websocket)

    async def send_client(self, message):
        """广播消息给所有客户端"""
        if not self.clients:
            print("无客户端连接")
            return

        disconnected = []
        for client in self.clients:
            try:
                await client.send(message)
            except websockets.ConnectionClosed:
                disconnected.append(client)
        # 清理断开连接的客户端
        for client in disconnected:
            self.clients.remove(client)

    async def stop_server(self):
        """安全关闭服务器"""
        if self.running:
            print("停止服务器...")
            self.running = False
            if self.server:
                self.server.close()
                await self.server.wait_closed()
                print("服务器已停止")

    async def start_server(self):
        """启动服务器并保持运行"""
        print(f"启动WebSocket服务器: ws://{self.host}:{self.port}")
        async with websockets.serve(self.handle_client, self.host, self.port) as server:
            self.server = server
            self.running = True
            print("服务器已启动，等待连接...")
            while self.running:  # 保持服务器运行
                await asyncio.sleep(1)  # 避免CPU占用过高

async def main():
    server = WSS()
    try:
        await server.start_server()
    except KeyboardInterrupt:
        print("手动中断: 程序终止")
    finally:
        await server.stop_server()

if __name__ == "__main__":
    asyncio.run(main())
    
#========================================

import asyncio
import websockets
from enum import Enum

class TaskState(Enum):
    RUNNING = 1
    PAUSED = 2
    STOPPED = 3

class AsyncTask:
    def __init__(self):
        self.state = TaskState.RUNNING
        self.lock = asyncio.Lock()
        self.pause_event = asyncio.Event()
        self.pause_event.set()  # 初始设置为已触发状态
        self.task = None

    async def task_function(self):
        try:
            while self.state != TaskState.STOPPED:
                # 将锁的作用域扩展到整个循环
                async with self.lock:
                    if self.state == TaskState.PAUSED:
                        self.pause_event.clear()  # 进入暂停时重置事件
                
                # 等待恢复信号的正确位置
                if self.state == TaskState.PAUSED:
                    await self.pause_event.wait()

                # 实际任务逻辑
                print("Task is running...")
                await asyncio.sleep(1)
        except asyncio.CancelledError:
            print("Task was aborted.")

    async def control_handler(self, websocket):
        async for message in websocket:
            async with self.lock:  # 保证原子操作
                if message == "pause":
                    self.state = TaskState.PAUSED
                    self.pause_event.clear()
                elif message == "resume":
                    self.state = TaskState.RUNNING
                    self.pause_event.set()  # 触发事件唤醒任务
                elif message == "abort":
                    self.task.cancel()
                    self.state = TaskState.STOPPED

    async def start(self):
        async with websockets.serve(self.control_handler, "localhost", 6565):
            self.task = asyncio.create_task(self.task_function())
            await self.task