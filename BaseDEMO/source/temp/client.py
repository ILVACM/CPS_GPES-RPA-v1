import asyncio
import json
import websockets

from websockets.asyncio.client import connect

operate ={
    "type": "signal",
    "data": "exit"
}

# async def hello():
#     uri = "ws://localhost:6565"
#     async with connect(uri) as websocket:
#         # name = input("What's your name? ")

#         await websocket.send(json.dumps(operate))
#         print(f"send >>> {operate}")

#         greeting = await websocket.recv()
#         print(f"recv <<< {greeting}")

# if __name__ == "__main__":
#     asyncio.run(hello())
    
async def send_control(command):
    async with websockets.connect("ws://localhost:6565") as ws:
        await ws.send(command)

# asyncio.run(send_control("pause"))
asyncio.run(send_control("resume"))
# asyncio.run(send_control("abort"))