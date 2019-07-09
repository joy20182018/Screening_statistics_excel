#utf-8
# 对excel文件进行读取，写入的操作
import xlrd
import xlwt
import re
import requests


def str_trans_int(string):
    """字符串转数字函数
    Args:
    string:一个字符串
    Returns:提取到的字符串中的数字
    """

    # 字符库
    # 做这个的原因：有些人在填表或制表过程中写作不规范导致
    numlib = {'零': '0', '一': '1', '二': '2', '三': '3',
            '四': '4', '五': '5', '六': '6', '七': '7',
            '八': '8', '九': '9', '十': '0',
            '岁': ''
            }
    # 有时会出现更加意想不到的情况
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


# 读取原始数据集，从网上或者其他渠道获取的excel文件
data = xlrd.open_workbook('原始数据.xls')
table = data.sheets()[0]   # 提取第一个sheet
nrows = table.nrows   # 行数
ncols = table.ncols   # 列数

sum_match_person = 0  # 符合条件的人数
sum_match_boy = 0  # 符合条件的男生人数
sum_match_girl = 0  # 符合条件的女生人数
sum_no_match_person = 0 # 不符合条件的人数
sum_no_match_boy = 0  # 不符合条件的男生人数
sum_no_match_girl = 0  # 不符合条件的女生人数
# 写入数据
# 创建第一个表格
workbook_1 = xlwt.Workbook(encoding='utf-8', style_compression=0)   
# 合格人数表格
sheet_match = workbook_1.add_sheet('DATAmatch', cell_overwrite_ok=True) 
# 创建第二个表格
workbook_2 = xlwt.Workbook(encoding='utf-8', style_compression=1)   
# 不合格人数表格
sheet_nomatch = workbook_2.add_sheet('DATAnomatch', cell_overwrite_ok=True) 

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

check_count = 466 #  从第几行开始检查

for i in range(check_count, nrows):
    # 从excel中提取有用信息进行比较
    num = table.row(i)[0].value
    if num > 1:
        Score = table.row(i)[ncols - 1].value
        Age = table.row(i)[7].value # 表示第i行第7列
        Age= str_trans_int(Age)
        Name = table.row(i)[6].value
        Sex = table.row(i)[8].value

        sum_person = sum_person + 1


        # 合格人数统计
        if Score < 5 and Score >= 0 and (not search_name(Name)):
            sum_match_person = sum_match_person + 1

            # 根据表格制作者的信息进行修改
            if Sex == 1:
                sex = "男"
                sum_match_boy = sum_match_boy + 1
            if Sex == 2:
                sex = "女"
                sum_match_girl = sum_match_girl + 1
            # 写入数据
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

# 保存两个文件
workbook_1.save(r'合格人.xls')
workbook_2.save(r'不合格人.xls')
print("运行完毕")
print("生成两个文件： 合格人.xls   不合格人.xls")
