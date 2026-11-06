import logging

import Init

#==================================================================

def setup_logging():
    try:
        # 配置日志格式和日志级别
        logging.basicConfig(
            level=logging.DEBUG,  # 日志级别
            format='%(asctime)s [%(levelname)s] [%(module)s] [%(funcName)s] - %(message)s',  # 自定义日志格式
            datefmt='%Y-%m-%d %H:%M:%S',  # 日期格式
            filename= Init.LogFiles_Path + "\\" + Init.LogFiles + '.log',
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

if __name__ == "__main__":
    pass