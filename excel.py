#utf-8
# 对excel文件进行读取，写入的操作
import xlrd
import xlwt
import re
import requests


# def StrTransInt(string):
#     try:
#         Tran = str(string)
#         STR = re.findall(r"\d+\.?\d*",Tran)
        
#     except:
#         return 0
#     else:
#         if STR == []:
#             return 0
#         else:
#             return float(STR[0])

def str_trans_int(string):
    """字符串转数字函数
    Args:
    string:一个字符串
    Returns:提取到的字符串中的数字
    """

    # 字符库
    numlib = {'零': '0', '一': '1', '二': '2', '三': '3',
            '四': '4', '五': '5', '六': '6', '七': '7',
            '八': '8', '九': '9', '十': '0',
            '岁': ''
            }
    try:
        Tran = str(string)  
        STR = re.findall(r"\d+\.?\d*", Tran)   # 提取其中的数字
    except Exception as error:
        print(error)
    else:  
        if STR == []:
            if Tran == 0:
                return 0
            else:
                a = []
                for i in range(len(Tran)):
                    a.append(numlib[Tran[i]])
                b = ''.join(a)
                return float(b)
        else:
            return float(STR[0])


def search_name(init_name):
    """查找原有表中与输入是否有重复的项
    Args:
    name: 输入的一个字符或字符串
    Returns: 
    1: 该人已经存在
    0：该人不存在
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
            return 1
        else:
            return 0


"""汉字处理的工具:
判断unicode是否是汉字，数字，英文，或者其他字符。
"""
def is_chinese(uchar):
    """判断一个unicode是否是汉字"""
    if uchar >= u'\u4e00' and uchar<=u'\u9fa5':
        return True
    else:
        return False

def is_number(uchar):
    """判断一个unicode是否是数字"""
    if uchar >= u'\u0030' and uchar<=u'\u0039':
        return True
    else:
        return False

def is_alphabet(uchar):
    """判断一个unicode是否是英文字母"""
    if (uchar >= u'\u0041' and uchar<=u'\u005a') or (uchar >= u'\u0061' and uchar<=u'\u007a'):
        return True
    else:
        return False

def is_other(uchar):
    """判断是否非汉字，数字和英文字符"""
    if not (is_chinese(uchar) or is_number(uchar) or is_alphabet(uchar)):
        return True
    else:
        return False

def string2List(ustring):
    """将ustring按照中文，字母，数字分开"""
    retList=[]
    utmp=[]
    for uchar in ustring:
        if is_other(uchar):
            if len(utmp)==0:
                continue
            else:
                retList.append("".join(utmp))
                utmp=[]
        else:
            utmp.append(uchar)
        if len(utmp)!=0:
            retList.append("".join(utmp))
        return retList


# 记录录过音的人
'''
recordedName = [
                '晶晶', '李书平', '张金芝', '张俊波', '刘燕', '庞宝英', '王长福', '季学东',
                '张冉', '陈元青', '黄曼', '高浩','姜九龙', '沈秀萍', '王喜莲', '方勉', '于小英',
                '张振兴', '李晓伟', '王桂英', '田亚凤', '刘秀芹', '刘春华', '任福洋','郝淑香',
                '苗海燕', '王大雷', '王娇', '张方田', '庄明珍', '庄明环','黄瑞云','刘春芳',
                '赵莲芝', '赵小强', '黄瑞云', '孙逊', '张璐', '刘秀荣', '曾国庆', '刘芸', '王明',
                '赵东锁', '王玉', '程晓芳'#, '陈瑞', '于世军', '赵丽云', '何昕', '何荫林' ,
                '张轩', '葛喜成', '徐玲', '宋强', '杨洁林', '付景芬', '聂国兵', '苏立铮', '孙梅',
                '张新', '陈雄辉', '马一琛', '刘金秋'
赵雪燕
孟卫
王元龙
樊珍珍         张典
吴凡
石榴
刘玉明    周静      郝志德     陆洪文        田亚其       
 李爱兵        闫素丽 苏卫革 李秋菊 李岩
 叶玉兵
王丹
卞西朋
李敬春
王阔
                ]
'''

data = xlrd.open_workbook('原始数据.xls')
table = data.sheets()[0]   # 提取第一个sheet
nrows = table.nrows   # 行数
ncols = table.ncols   # 列数

# print("输入姓名，结束时按 # \n")
# count = 0
# while True:
#     name = input("姓名: ")
#     if name == '#':
#         break
#     else:
#         count = count + 1
#         with open('E:/ExcelOp/recoded_name.txt', 'a') as f:
#             if count == 1:
#                 f.write('\n')
#                 f.write(name + '\n')
#             else:
#                 f.write(name + '\n')

# print("输入的姓名数目：%d"%(count))


sum_match_person = 0  # 符合条件的人数
sum_match_boy = 0  # 符合条件的男生人数
sum_match_girl = 0  # 符合条件的女生人数
sum_no_match_person = 0 # 不符合条件的人数
sum_no_match_boy = 0  # 不符合条件的男生人数
sum_no_match_girl = 0  # 不符合条件的女生人数
# 写入数据
workbook_1 = xlwt.Workbook(encoding='utf-8', style_compression=0)   # 创建第一个表格
sheet_match = workbook_1.add_sheet('DATAmatch', cell_overwrite_ok=True)   # 合格人数表格
workbook_2 = xlwt.Workbook(encoding='utf-8', style_compression=1)   # 创建第二个表格
sheet_nomatch = workbook_2.add_sheet('DATAnomatch', cell_overwrite_ok=True) # 不合格人数表格

# 在表格第一行写入标记
sheet_match.write(0, 0, '姓名')
sheet_match.write(0, 1, '年龄')
sheet_match.write(0, 2, '性别')
sheet_match.write(0, 3, '分数')

sheet_nomatch.write(0, 0, '姓名')
sheet_nomatch.write(0, 1, '年龄')
sheet_nomatch.write(0, 2, '性别')
sheet_nomatch.write(0, 3, '分数')

sum_person = 0  # 总人数

for i in range(406, nrows):
    num = table.row(i)[0].value
    if num > 1:
        Score = table.row(i)[ncols - 1].value
        Age = table.row(i)[7].value
        Age= str_trans_int(Age)
        Name = table.row(i)[6].value
        Sex = table.row(i)[8].value

        sum_person = sum_person + 1
#
# (Name not in recordedName)
        if Score < 5 and Score >= 0 and (Age >= 30 or Age == 0) and Age <= 60 and (not search_name(Name)):
            sum_match_person = sum_match_person + 1
            if Sex == 1:
                sex = "男"
                sum_match_boy = sum_match_boy + 1
            if Sex == 2:
                sex = "女"
                sum_match_girl = sum_match_girl + 1

            sheet_match.write(sum_match_person, 0, Name)
            
            if Age == 0:
                Age = table.row(i)[7].value
                sheet_match.write(sum_match_person, 1, Age)
            else:
                sheet_match.write(sum_match_person, 1, Age)
   
            sheet_match.write(sum_match_person, 2, sex)
            sheet_match.write(sum_match_person, 3, Score)

        # 不合格人数统计  
        else:
            Age = table.row(i)[7].value
            # # if Name in recordedName:
            # if search_name(Name):
            #     continue
            # else:
            sum_no_match_person = sum_no_match_person + 1
            if Sex == 1:
                sex = "男"
                sum_no_match_boy = sum_no_match_boy + 1
            if Sex == 2:
                sex = "女"
                sum_no_match_girl = sum_no_match_girl + 1

            sheet_nomatch.write(sum_no_match_person , 0, Name)
            sheet_nomatch.write(sum_no_match_person , 1, Age)
            sheet_nomatch.write(sum_no_match_person, 2, sex)
            sheet_nomatch.write(sum_no_match_person, 3, Score)

sheet_match.write(sum_match_person + 1, 4, "问卷总人数")
sheet_match.write(sum_match_person + 1, 5, sum_person)   
sheet_match.write(sum_match_person + 2, 4, "符合条件的人数")
sheet_match.write(sum_match_person + 2, 5, sum_match_person)
sheet_match.write(sum_match_person + 3, 4, "男生人数")
sheet_match.write(sum_match_person + 3, 5, sum_match_boy)
sheet_match.write(sum_match_person + 4, 4, "女生人数")
sheet_match.write(sum_match_person + 4, 5, sum_match_girl)

sheet_nomatch.write(sum_no_match_person + 1, 4, "不符合条件的人数")
sheet_nomatch.write(sum_no_match_person + 1, 5, sum_no_match_person)
sheet_nomatch.write(sum_no_match_person + 2, 4, "男生人数")
sheet_nomatch.write(sum_no_match_person + 2, 5, sum_no_match_boy)
sheet_nomatch.write(sum_no_match_person + 3, 4, "女生人数")
sheet_nomatch.write(sum_no_match_person + 3, 5, sum_no_match_girl)


workbook_1.save(r'合格人.xls')
workbook_2.save(r'不合格人.xls')
print("运行完毕")
print("生成两个文件： 合格人.xls   不合格人.xls")