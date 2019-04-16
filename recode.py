#utf-8

def search_name(init_name):
    """查找原有表中与输入是否有重复的项
    Args:
    name: 输入的一个字符或字符串
    Returns: 
    """

    search = False
    with open('E:/ExcelOp/recoded_name.txt', 'r') as f:
        for line_name in f:
            if line_name[:-1] == init_name:
                # print(line_name)
                # print(init_name)
                search = True
                break

        if search == True:
            return '存在'
        else:
            return '不存在'


"""
记录人员名单 + 查找是否有重复项
"""

while True:
    print("-"*16)
    print("输入 0： 输入已录人员名单， 1：查找是否有重复项， 2：列出现有的名单， 3： 不输入任何东西，结束程序")
    print("-"*16)
    
    tag = int(input("输入数字："))

    if tag == 0:
        print("输入姓名，结束时按 # \n")
        count = 0
        while True:
            name = input("姓名: ")
            if name == '#':
                break
            else:
                count = count + 1
                with open('E:/ExcelOp/recoded_name.txt', 'a') as f:
                    if count == 1:
                        f.write('\n')
                        f.write(name + '\n')
                    else:
                        f.write(name + '\n')

        print("输入的姓名数目：%d"%(count))
        print("\n")
        print("是否继续：1：继续， 0：取消")
        n = int(input("输入："))
        if n == 1:
            pass
        else:
            break

    elif tag == 1:
        print("输入姓名\n")
        name = input("姓名: ")
        print(search_name(name))
        print("是否继续：1：继续， 0：取消")
        n = int(input("输入："))
        if n == 1:
            pass
        else:
            break

    elif tag == 2:
        with open('E:/ExcelOp/recoded_name.txt', 'r') as f:
            for line_name in f:
                print(line_name)
        print("-"*16)
        print("结束")

    elif tag == 3:
        break

    else:
        print("重输")

