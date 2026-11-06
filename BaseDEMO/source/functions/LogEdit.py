import logging
from datetime import datetime

import functions.Init

#==================================================================

def setup_logging():
    try:
        # 配置日志格式和日志级别
        logging.basicConfig(
            level=logging.DEBUG,  # 日志级别
            format='%(asctime)s [%(levelname)s] [%(module)s] [%(funcName)s] - %(message)s',  # 自定义日志格式
            datefmt='%Y-%m-%d %H:%M:%S',  # 日期格式
            filename= functions.Init.LogFiles_Path + "\\" + functions.Init.LogFiles + '.log',
            filemode='a',
            encoding='utf-8'
        )
        
        # logging.info("日志设置成功")
        
    except Exception as e:
        logging.error(f"设置日志配置时发生错误: {e}")
        
def Record( content: str, level: int):
    if level == 1:
        logging.info(content)
    
    elif level == 2:
        logging.warning(content)
    
    elif level == 3:
        logging.error(content)
    
    elif level == 4:
        logging.critical(content)
    
    elif level == 5:
        logging.debug(content)
    
    else:
        print("Logging exception!!!")
        
def TaskLog( id: int, Task_mode: str, Total: int, Compeleted: int):
    # file=functions.Init.Tasklogger
    
    # 获取当前系统时间
    current_time = datetime.now()

    # 将当前时间格式化为字符串
    time_str = current_time.strftime('%Y-%m-%d %H:%M:%S')
    
    functions.Init.Tasklogger.write( "\n" + time_str + "\n")
    
    functions.Init.Tasklogger.write( "############################################################" + "\n")
    functions.Init.Tasklogger.write( "Task ID:           " + str(id) + "\n")
    functions.Init.Tasklogger.write( "Task Mode:         " + Task_mode + "\n")
    functions.Init.Tasklogger.write( "Total OP num:      " + str(Total) + "\n")
    functions.Init.Tasklogger.write( "Compeleted OP num: " + str(Compeleted) + "\n")
    functions.Init.Tasklogger.write( "Task Schedule:     " + str( Compeleted/Total) + "\n")

    # print( "Task Status:       " + str(status), file=functions.Init.Tasklogger)    
    if Compeleted == Total:
        functions.Init.Tasklogger.write( "Task Status:       " + "Success")
    
    else:
        if Compeleted == 0:
            functions.Init.Tasklogger.write( "Task Status:       " + "Failed")
        else:
            functions.Init.Tasklogger.write( "Task Status:       " + "Aborted")
    
    functions.Init.Tasklogger.write( "\n------------------------------------------------------------\n")
    
    functions.Init.Tasklogger.flush()
    # pass

if __name__ == "__main__":
    functions.Init.init()
    
    TaskLog( 1, "Add", 10, 5)