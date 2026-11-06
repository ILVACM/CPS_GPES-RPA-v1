import logging
from datetime import datetime

#----------------------------

from functions import Init

#===========================================


def Record(content: str, level: int):
    try:
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
            raise ValueError(f"Invalid logging level: {level}")

    except Exception as e:
        # 捕获日志记录过程中发生的异常并打印错误信息
        print(f"Logging exception: {e}")


def TaskLog( ID: int, Task_mode: str, Total: int, Compeleted: int, status: int):
    # file=functions.Init.Tasklogger

    # 获取当前系统时间
    current_time = datetime.now()
    # 将当前时间格式化为字符串
    time_str = current_time.strftime('%Y-%m-%d %H:%M:%S')

    Init.Tasklogger.write("\n" + time_str + "\n")

    Init.Tasklogger.write("############################################################" + "\n")
    Init.Tasklogger.write("Task ID:           " + str(ID) + "\n")
    Init.Tasklogger.write("Task Mode:         " + Task_mode + "\n")
    Init.Tasklogger.write("Total OP num:      " + str(Total) + "\n")
    Init.Tasklogger.write("Compeleted OP num: " + str(Compeleted) + "\n")
    Init.Tasklogger.write("Task Schedule:     " + str(Compeleted*100 / Total) + "%\n")

    # print( "Task Status:       " + str(status), file=Init.Tasklogger)

    if status == 0:
        Init.Tasklogger.write("Task Status:       " + "Success")
    elif status == 1:
        Init.Tasklogger.write("Task Status:       " + "Aborted")
    elif status == 2:
        Init.Tasklogger.write("Task Status:       " + "Failed")

    Init.Tasklogger.write("\n------------------------------------------------------------\n")

    Init.Tasklogger.flush()
    # pass