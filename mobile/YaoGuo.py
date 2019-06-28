import requests
import xlwt


# 创建单页第一行的标题
def create_tab(title, wb):
    sheet = wb.add_sheet(title, cell_overwrite_ok=True)
    sheet.write(0, 0, '题干')
    sheet.write(0, 1, '答案')
    sheet.write(0, 2, '选项')
    sheet.write(0, 3, '解析')
    sheet.write(0, 4, '来源')
    sheet.write(0, 5, '资料(资料题特有)')
    return sheet


# 获取题号映射
def get_options(item):
    kw = {
        0: 'A',
        1: 'B',
        2: 'C',
        3: 'D',
        4: 'E',
        5: 'F',
        6: 'G',
        7: 'H',
        8: 'I',
        9: 'J'
    }
    return kw.get(item)


# 设置表格样式
def set_style(name, height, bold=False):
    style = xlwt.XFStyle
    font = xlwt.Font
    font.height = height
    font.bold = bold
    font.colour_index = 4
    style.font = font
    return style


# 写入数据
def write_data(title, url, questions_count, wb):
    # 创建一页
    sheet = create_tab(title, wb)
    # 当前已经得到的题的id数组
    id_list = []
    # 如果小于需要的题的数量就继续请求
    while id_list.__len__() < questions_count:
        r = requests.get(url)
        data = r.json()
        questions = data.get('questions')
        print('paper_id = %s' % data.get('paper_id'))
        for i in range(0, len(questions)):
            if bool(1 - (questions[i]['id'] in id_list)):
                print('正在写入 id = %s 的问题' % questions[i]['id'])
                cur_question_count = id_list.__len__()
                # 题干
                question = questions[i]['question']
                sheet.write(cur_question_count + 1, 0, question)
                # 答案
                answer = questions[i]['answer']
                sheet.write(cur_question_count + 1, 1, answer)
                # 选项
                options = questions[i]['options']
                f = ''
                for j in range(0, len(options)):
                    f += '%s:%s  ' % (get_options(j), options[j])
                sheet.write(cur_question_count + 1, 2, f)
                # 解析
                analysis = questions[i]['analysis']
                sheet.write(cur_question_count + 1, 3, analysis)
                # 来源
                source = questions[i]['source']
                sheet.write(cur_question_count + 1, 4, source)
                # 资料
                material = questions[i]['material']
                sheet.write(cur_question_count + 1, 5, material)
                # 该题写入完成，把题号记录下来
                id_list.append(questions[i]['id'])
                print('id = %s 写入完成' % questions[i]['id'])
                print('题库总数 = %s' % cur_question_count)


# 写入excel
def create_excel(title, bank_id, questions_count):
    url = 'http://spark.appublisher.com/quizbank/get_note_questions?terminal_type=android_phone' \
          '&app_type=quizbank&app_version=3.13.0&uuid=075c74b74359a52c&user_id=7017385' \
          '&user_token=95a798974cfed2a111677de9f33e8fda&cid=null&timestamp=1557128269098' \
          '&note_id=%s&type=note' % bank_id
    print(url)
    wb = xlwt.Workbook()
    write_data(title, url, questions_count, wb)
    wb.save('data/%s.xls' % title)


#  腰果题库数据抓包
if __name__ == '__main__':
    """
    常识判断-
            | 科技 id=1457,total=963
            | 政治 id=647,total=489
            | 历史 id=1064,total=167
            | 法律 id=145,total=407
            | 经济 id=1455,total=179
            | 人文 id=1456,total=490
            | 管理公文 id=1458,total=72
    言语理解与表达-
                | 语句表达 id=135,total=453
                | 逻辑填空 id=103,total=1754
                | 片段阅读 id=138,total=1888
    数量关系-
            | 数学关系 id=60,total=1534
            | 数字推理 id=1453,total=318
    判断推理-
            | 定义判断 id=1466,total=1235
            | 逻辑判断 id=1460,total=1591
            | 类比推理 id=147,total=1349
            | 图形推荐 id=706,total=1101
    资料分析-
            | 计算类 id=1449,total=1213
            | 大小类比类 id=1450,total=385
            | 读数类 id=1448,total=319
            | 综合分析类 id=1451,total=478
    """

    # create_excel('腰果-行测-常识判断-科技', 1457, 900)
    # create_excel('腰果-行测-常识判断-政治', 647, 400)
    # create_excel('腰果-行测-常识判断-历史', 1064, 150)
    # create_excel('腰果-行测-常识判断-法律', 145, 350)
    # create_excel('腰果-行测-常识判断-经济', 1455, 150)
    # create_excel('腰果-行测-常识判断-人文', 1456, 450)
    # create_excel('腰果-行测-常识判断-管理公文', 1458, 50)
    #
    # create_excel('腰果-行测-言语理解与表达-语句表达', 135, 400)
    # create_excel('腰果-行测-言语理解与表达-逻辑填空', 103, 1500)
    # create_excel('腰果-行测-言语理解与表达-片段阅读', 138, 1500)
    #
    # create_excel('腰果-行测-数量关系-数学关系', 60, 1200)
    # create_excel('腰果-行测-数量关系-数字推理', 1453, 270)
    #
    # create_excel('腰果-行测-判断推理-定义判断', 1466, 1000)
    # create_excel('腰果-行测-判断推理-逻辑判断', 1460, 1200)
    # create_excel('腰果-行测-判断推理-类比推理', 147, 1100)
    # create_excel('腰果-行测-判断推理-图形推荐', 706, 900)

    create_excel('腰果-行测-资料分析-计算类', 1449, 1000)
    create_excel('腰果-行测-资料分析-大小类比类', 1450, 300)
    create_excel('腰果-行测-资料分析-读数类', 1448, 270)
    create_excel('腰果-行测-资料分析-综合分析类', 1451, 400)
