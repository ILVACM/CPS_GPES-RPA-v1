import asyncio
import websockets
import json
import time

from concurrent.futures import ThreadPoolExecutor
from enum import Enum

# from functions import config, Init  # 保持原有导入

#---------------------------
# 增强型状态管理
#---------------------------
class TaskStatus(Enum):
    IDLE = 0      # 空闲状态
    RUNNING = 1   # 运行中
    PAUSED = 2    # 已暂停
    STOPPED = 3   # 已终止
    
class TaskController:
    def __init__(self):
        self.status = "idle"  # idle/running/paused/stopped
        self.progress = 0
        self.task = None
        self.pause_event = asyncio.Event()
        self.stop_event = asyncio.Event()
        self.lock = asyncio.Lock()

    async def start_task(self, command):
        async with self.lock:
            if self.status in ["running", "paused"]:
                return
            self.status = "running"
            self.stop_event.clear()
            self.task = asyncio.create_task(self._execute(command))

    async def _execute(self, command):
        try:
            while self.progress < 100 and not self.stop_event.is_set():
                if self.pause_event.is_set():
                    await self.pause_event.wait()
                    continue
                
                # 模拟耗时操作
                await asyncio.sleep(0.5)
                self.progress += 1
                print(f"当前进度: {self.progress}%")
                
        except asyncio.CancelledError:
            print("任务被终止")
        finally:
            self.status = "idle"
            self.progress = 0

    async def pause_task(self):
        async with self.lock:
            if self.status == "running":
                self.status = "paused"
                self.pause_event.set()

    async def resume_task(self):
        async with self.lock:
            if self.status == "paused":
                self.status = "running"
                self.pause_event.clear()

    async def stop_task(self):
        async with self.lock:
            if self.task and not self.task.done():
                self.stop_event.set()
                self.task.cancel()
                try:
                    await self.task
                except asyncio.CancelledError:
                    pass

class OperationManager:
    def __init__(self):
        self.status = TaskStatus.IDLE     # 初始状态为空闲
        self.progress = 0                 # 执行进度百分比（0-100）
        self.checkpoints = []             # 断点记录（用于恢复执行）
        self.lock = asyncio.Lock()        # 异步锁保证状态修改原子性
        self.current_task = None          # 当前运行的任务对象
        self.pause_event = asyncio.Event()# 暂停事件标志（set表示需要暂停）
        self.stop_event = asyncio.Event() # 终止事件标志
        
        # 在OperationManager中添加
        self.executor = ThreadPoolExecutor(max_workers=4)  # CPU密集型操作
        self.io_executor = ThreadPoolExecutor(max_workers=8)  # IO密集型操作

# 执行时选择执行器
# await loop.run_in_executor(executor, func)

async def start_operate( test):
    # 格式逻辑测试
    print( test)

#---------------------------
# WebSocket服务增强
#---------------------------
class EnhancedWSS:
    def __init__(self, host="localhost"):
        with open( 'default.json', 'r', encoding='utf-8') as file:
            source = json.load(file)

        self.host = host
        self.port = source['port']
        self.clients = set()  # 存储所有连接的客户端
        self.lock = asyncio.Lock()  # 异步锁防止并发冲突
        self.server = None  # WebSocket服务器实例
        
        self.op_manager = OperationManager()
        self.task_queue = asyncio.Queue()  # 任务队列
        
        self.task_ctl = TaskController()  # 新增任务控制器
        
    async def handle_client(self, websocket):
        async for message in websocket:
            # 协议解析层
            command = await self._parse_command(message)
            
            # 命令路由层
            await self._command_router(websocket, command)

    async def send_client( self, message):
        """向所有连接的客户端发送消息"""
        if not self.clients:
            print("没有客户端连接，无法发送消息")
            return
        disconnected_clients = set()

        for client in self.clients:
            try:
                await client.send(message)
            except websockets.ConnectionClosed:
                disconnected_clients.add(client)
        # 清理已断开的客户端
        self.clients -= disconnected_clients

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

        # 停止事件循环
        # print("停止事件循环...")
        # asyncio.get_event_loop().stop()
        # print("事件循环已停止")

    async def start_server(self):
        """启动 WebSocket 服务器并设置信号处理"""
        print(f"启动 WebSocket 服务器: ws://{self.host}:{self.port}")
        self.server = await websockets.serve( self.handle_client, self.host, self.port)

        # 进入事件循环，等待 WebSocket 连接
        await self.server.wait_closed()

    async def _parse_command(self, raw_message):
        try:
            # 1. 反序列化JSON
            command = json.loads(raw_message)
            
            # 2. 校验协议格式
            if "type" not in command or "data" not in command:
                raise ValueError("命令必须包含 'type' 和 'data' 字段")
                
            # 3. 校验操作类型合法性
            valid_commands = {"start", "pause", "resume", "stop", "query"}
            if command["type"] not in valid_commands:
                raise ValueError(f"无效命令类型: {command['type']}")
                
            return command
            
        except json.JSONDecodeError as e:
            # JSON解析失败处理
            raise ValueError("消息格式错误: 非法的JSON") from e

    async def _command_router(self, websocket, command):
        # 使用状态机处理命令
        async with self.op_manager.lock:
            if command["type"] == "start":
                await self._handle_start(command)
            elif command["type"] == "pause":
                await self._handle_pause()
            elif command["type"] == "resume":
                await self._handle_resume()
            elif command["type"] == "stop":
                await self._handle_stop()
                
        await self._send_status_update(websocket)

    async def _handle_start(self, command):
        if self.op_manager.status == TaskStatus.IDLE:
            self.op_manager.current_task = asyncio.create_task(
                self._exec_operation(command)
            )
            self.op_manager.status = TaskStatus.RUNNING

# 添加性能监控装饰器
    def async_monitor(func):
        async def wrapper(*args, **kwargs):
            start = time.monotonic()
            try:
                return await func(*args, **kwargs)
            finally:
                duration = time.monotonic() - start
                print(f"{func.__name__} executed in {duration:.2f}s")
        return wrapper

    async def _exec_operation(self, command):
        try:
            # 支持断点续传的执行逻辑
            while self.op_manager.progress < 100:
                if self.op_manager.stop_event.is_set():
                    break
                
                if self.op_manager.pause_event.is_set():
                    await self.op_manager.pause_event.wait()
                
                # 执行原子操作
                await start_operate({
                    **command,
                    "progress": self.op_manager.progress
                })
                
                self.op_manager.progress += 1
                await asyncio.sleep(0.1)  # 模拟耗时操作
                
        except asyncio.CancelledError:
            # 任务取消时的清理逻辑
            await self._cleanup_operation()
        finally:
            self.op_manager.status = TaskStatus.IDLE

    async def _handle_pause(self):
        self.op_manager.pause_event.set()
        self.op_manager.status = TaskStatus.PAUSED
        self.op_manager.checkpoints.append(self.op_manager.progress)

    async def _handle_resume(self):
        self.op_manager.pause_event.clear()
        self.op_manager.status = TaskStatus.RUNNING

    async def _handle_stop(self):
        self.op_manager.stop_event.set()
        if self.op_manager.current_task:
            self.op_manager.current_task.cancel()
        await self._cleanup_operation()

    async def _cleanup_operation(self):
        self.op_manager.progress = 0
        self.op_manager.checkpoints.clear()
        self.op_manager.stop_event.clear()

if __name__ == "__main__":
    # 创建实例并运行
    enhanced_wss = EnhancedWSS()
    asyncio.run(enhanced_wss.start_server())