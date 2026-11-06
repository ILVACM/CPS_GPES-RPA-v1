import asyncio
import websockets

import json

from enum import Enum

#----------------------------------------

#========================================

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

# WebSocket Server
class WSS:
    def __init__( self, port=8765, host="localhost"):
        # 通过json配置文件 读取端口号
        with open( 'default.json', 'r', encoding='utf-8') as file:              # 内存盘 Z: -> 测试 
        # with open( 'config\default.json', 'r', encoding='utf-8') as file:       # 正式盘 D: -> 实际运行
            source = json.load(file)
        
        self.host = host
        self.port = source['port']
        
        self.clients = set()
        
        self.lock = asyncio.Lock()
        self.server = None
        
        self.test_couter = 0

    async def start_server(self):
        # pass
        
        # 启动服务器并保持长连接
        self.server = await websockets.serve(
            self.handle_client,
            self.host,
            self.port
        )
        
        print(f'WebSocket Server started at {self.host}:{self.port}')
        
        # 关键：保持服务器运行
        await self.server.wait_closed()
    
    async def stop_server(self):
        # pass
        # 关闭所有客户端连接
        # for client in self.clients:
        #     try:
        #         await client.close()
        #     except Exception as e:
        #         print(f"关闭客户端连接时出错: {e}")
        
        # 关闭服务器
        async with self.lock:
            if self.server:
                self.server.close()
                self.server = None
                
                # asyncio.get_event_loop.stop()
                
                print(f'Closed at {self.host}:{self.port}')
    
    # 异步上下文管理器支持 --------------------------------
    async def __aenter__(self):
        """进入异步上下文时自动启动服务器"""
        await self.start_server()
        return self

    async def __aexit__(self):
        """退出异步上下文时自动关闭服务器"""
        await self.stop_server()
    # ----------------------------------------------------
    async def handle_client(self, websocket):
        # pass
        
        # 客户端连接处理器
        self.clients.add(websocket)
        
        try:
            async for message in websocket:
                # 显示消息
                print(f'Received: \n{message}\n')
                
                # 向客户端发送回应
                await websocket.send(json.dumps(response))
                
                # 消息格式预处理
                message = json.loads(message)
                
                # 2. 检查是否是退出信号
                if message.get("type") == "signal" and message.get("data") == "exit":
                    print("接收到退出信号，停止服务器...")
                    
                    await self.stop_server()  # 停止服务器
                    
                    return  # 退出当前连接

                # 3. 处理业务逻辑
                self.test_couter += 1
                print(self.test_couter) 
                 
        except json.JSONDecodeError:
            # 处理非 JSON 格式的消息
            await websocket.send(json.dumps(error_response))
            
        except websockets.ConnectionClosed as e:
            print(f"客户端断开连接: {e}")
            
        finally:
            # 移除断开的客户端
            self.clients.remove(websocket)
            
async def main():
    # 运行服务器
    # async with WSS() as server:
    #     await server.start_server()
    
    server = WSS()
    await server.start_server()

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
            i = 0 # 输出标识

            while self.state != TaskState.STOPPED:
                # 将锁的作用域扩展到整个循环
                async with self.lock:
                    if self.state == TaskState.PAUSED:
                        self.pause_event.clear()  # 进入暂停时重置事件

                # 等待恢复信号的正确位置
                if self.state == TaskState.PAUSED:
                    await self.pause_event.wait()

                # 实际任务逻辑（示例：周期性输出）
                print("Task is running...")
                print(i)
                i = i + 1

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

if __name__ == "__main__":
    pass

    # asyncio.run(main())

    task = AsyncTask()
    asyncio.run(task.start())