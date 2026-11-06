import asyncio
import websockets
import json

# 存储计时状态的变量
timer_running = False
elapsed_time = 0

# 后端处理WebSocket连接的函数
async def handle_message(websocket, path):
    global timer_running, elapsed_time

    async for message in websocket:
        data = json.loads(message)
        
        if data['action'] == 'start_timer':
            # 启动计时器，开始计时
            timer_running = True
            print("计时器开始")
        elif data['action'] == 'pause_timer':
            # 暂停计时器
            timer_running = False
            print("计时器暂停")
        
        # 每次状态改变时发送计时器状态给前端
        await websocket.send(json.dumps({
            'action': 'update_timer',
            'running': timer_running,
            'elapsed_time': elapsed_time
        }))
    
    # print(elapsed_time)

# 每秒更新计时器
async def timer_update():
    global elapsed_time
    while True:
        if timer_running:
            elapsed_time += 1  # 更新秒数
        await asyncio.sleep(1)

# 启动WebSocket服务器和计时器
async def main():
    server = await websockets.serve(handle_message, 'localhost', 8765)
    asyncio.create_task(timer_update())  # 启动定时器更新
    await server.wait_closed()

if __name__ == "__main__":
    asyncio.run(main())
