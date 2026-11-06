import os
import time
import re
import json
import requests

import functions.Init

def default():
    print("Please enter your wechat nickname:")
    functions.Init.Name = input()
    
def custom():
    print("Please enter your wechat nickname:")
    
    print("Invalid input! Please try again.")
    os.system("cls") # clear screen

def display():
    
    # Init.init()
    # print(Init.Name)
    
    print("Settings:")
    
    print( "\tName(Your wechat nickname):\n\t   ", functions.Init.Name, "\t")
    
    print( "\tURL(The target website for the file download operation):\n\t   ", functions.Init.URL, "\t")
    
    print( "\tDownload file storage location:\n\t   ", functions.Init.Download, "\t")
    
    print( "\tSendList(The wechat FRIENDS you want to send a message to):\n")
    for index in functions.Init.SendList:
        print("\t   ", index, "\t", functions.Init.friends[index],"\n")

    # print(len(functions.Init.friends))

    print("\tOperateList(A series of operations to be performed on each file unit):\n")
    for i in functions.Init.OperateList:
        print("\t   ", i, "\n")

def prompt():
    print("\t mode       \t  Unit          \t  Index                \t  postfix \n")
    # print("\n")
    print("\t 1 add      \t 1 column       \t 1 coord               \t 1 same ")
    print("\t            \t 2 row          \t 2 Regular expression  \t 2 Increase by degrees ")
    print("\t            \t 3 single(down) \t 3 Whole word matching \t 3 Decrease progressively \n")
    # print("\n")
    print("\t 2 search   \t 1 column       \t 1 coord               \t 1 single ")
    print("\t            \t 2 row          \t 2 Regular expression  \t 2 all ")
    print("\t            \t 3 single       \t 3 Whole word matching \t \n")
    # print("\n")
    print("\t 3 change   \t 1 column       \t 1 coord               \t ")
    print("\t            \t 2 row          \t 2 Regular expression  \t ")
    print("\t            \t 3 single       \t 3 Whole word matching \t \n")
    # print("\n")
    print("\t 4 delete   \t 1 column       \t 1 coord               \t 1 clean ")
    print("\t            \t 2 row          \t 2 Regular expression  \t 2 disappear")
    print("\t            \t 3 single       \t 3 Whole word matching \t ")
    print("\n")

def set(): 
    while True:
        print("Do you need to change the Settings? (y/n)")
        
        choice = input()
        if choice == "y":
            os.system("cls") # clear screen
            
            #
            print("1. WeChat Name 2. URL 3. Download file storage location 4. OperateList 5. SendList")
            print("Please enter the parameter number that you want to modify:")
            choice = input()
            
            if choice == "1":
                # WeChat Name
                print( "Please enter your new wechat nickname(If you want to skip this step, press enter directly):  " ,functions.Init.Name)
                
                temp = input()
                if temp != "":
                    functions.Init.Name = temp
                    
                os.system("cls") # clear screen
                
            elif choice == "2":
                # URL
                print( "Please enter your new URL(If you want to skip this step, press enter directly):  " ,functions.Init.URL)
                
                temp = input()
                if temp != "":
                    response = requests.get( temp, timeout=5)  # 设置超时为5秒
                    
                    if response.status_code == 200:
                        # print(f"网址 {URL} 可以正常访问。")
                        functions.Init.URL = temp
                    
                    else:
                        print(f"Website {temp} is inaccessible, status code: {response.status_code}\n Change is invalid.")
                        
                os.system("cls") # clear screen
                    
            elif choice == "3":
                # Download file storage location
                print( "Please enter your new Download file storage location:  " +functions.Init.Download)
                print("If you want to skip this step, press enter directly.")
                print("If you want to use the system's default download folder, enter only one space and press Enter.")
                
                temp = input()
                if temp != "":
                    if temp == " ":
                        functions.Init.Download = "C:\\Users\\AdmOA\\Downloads"
                    else:
                        functions.Init.Download = temp

                os.system("cls") # clear screen
                
            elif choice == "4":
                # OperateList
                print("Enter the preset operation code and related values as prompted:  ")
                prompt()
                
                while True:
                    os.system("cls") # clear screen
                    
                    print("Whether to set a new preset action: (y/n)")
                    choice = input()
                    
                    if choice == "y":
                        
                        print("reset?(y/n)")
                        choice = input()
                        if choice == "y":
                            functions.Init.OperateList = []
                        
                        os.system("cls") # clear screen
                            
                        print("Enter the preset operation code and related values as prompted:  ")
                        print("If multiple entries are selected, separate the two adjacent entries with Spaces")
                        prompt()
                        
                        choice = input()
                        temp = { "mode":choice.split()[0], "OperateCode":choice.split()[1], "row":choice.split()[2], "column":choice.split()[3], "match":choice.split()[4], "value":choice.split()[5]}
                        
                        # print(temp)
                        # choice = input()
                        # ########???
                        
                        functions.Init.OperateList.append(temp)
                    
                    elif choice == "n":
                        print("\tCheck the Settings again.\n")
                        for i in functions.Init.OperateList:
                            print("\t   ", i, "\n")
                        
                        time.sleep(3)
                        break
                    
                    else:
                        print("Invalid input! Please try again.")
                        time.sleep(1)
                
                os.system("cls") # clear screen
                
            elif choice == "5":
                os.system("cls") # clear screen
                
                i = 0
                for friend in functions.Init.friends:
                    print( i, friend)
                    i += 1
                    
                print( len(functions.Init.friends), "\n" )
                
                print("\nPlease input the wechat friend number you want to send the message to:")
                print("If multiple entries are selected, separate the two adjacent entries with Spaces")
                
                choice = input()
                # functions.Init.SendList = [int(num) for num in re.findall(r'\d+', choice) if int(num) <= functions.Init.friends_num]
                choice = choice.split()
                for i in choice:
                    functions.Init.SendList.append(int(i))
                functions.Init.SendList.sort()
                
                print("\tCheck the Settings again.\n")
                print(functions.Init.SendList)
                time.sleep(3)
                
                os.system("cls") # clear screen
                
            else:
                print("Invalid input")
                time.sleep(1)

        elif choice == "n":
            print("Check the Settings again.\n")
            
            display()
            
            choice = input("Press Enter to continue...")
            
            break
        
        else:
            print("Invalid input! Please try again.")
            os.system("cls") # clear screen
            
        # 保存修改到参数文件中
        with open( 'config\default.json', 'r', encoding='utf-8') as file:
            data = json.load(file)  # 读取文件并转换为字典
        
        data["name"]        = functions.Init.Name
        data["url"]         = functions.Init.URL
        data["download"]    = functions.Init.Download
        data["OperateList"] = functions.Init.OperateList
        data["SendList"]    = functions.Init.SendList
        
        with open( 'config\default.json', 'w', encoding='utf-8') as file:
            json.dump(data, file, ensure_ascii=False, indent=4)  # 写回 JSON 文件
        
if __name__ == "__main__":
    
    functions.Init.init()
    # # print(functions.Init.Name)
    
    # display()
    # prompt()
    
    set()
    