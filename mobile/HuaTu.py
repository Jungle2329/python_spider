import requests
import xlwt

header = {
    'token': '729e542c5724482bb73b581f1c8a51fc',
    'uid': '236204252',
    'cv': '7.1.140',
    'pixel': '1920_1080',
    'terminal': '1',
    'device': 'OnePlusONEPLUS A3010',
    'system': '9',
    'channelId': 'oppo',
    'appType': '2',
    'subject': '1',
    'catgory': '1',
}

# 获取题的接口
question_url = 'https://ns.huatu.com/q/v1/questions/?ids='
# 获取国考模考题列表
country_list_url = 'https://ns.huatu.com/p/v3/matches/past?page=1&tag=1'
# 获取省考模考题列表
province_list_url = 'https://ns.huatu.com/p/v3/matches/past?page=1&tag=2'
# 申论列表
mock_list_url = 'https://ns.huatu.com/e/api/v1/mock/papers?page=1'


# 创建单页第一行的标题
def create_tab(title, wb):
    sheet = wb.add_sheet(title, cell_overwrite_ok=True)
    sheet.write(0, 0, '题干')
    sheet.write(0, 1, '答案')
    sheet.write(0, 2, '选项')
    sheet.write(0, 3, '解析')
    sheet.write(0, 4, '来源')
    sheet.write(0, 5, '单/多选')
    sheet.write(0, 6, '分类')
    sheet.write(0, 7, '资料(资料题特有)')
    return sheet


# 替换html标签
def replace_html_tag(s):
    return s.replace('<p>', '').replace('</p>', '').replace('<br/>', '\r\n')


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


# 华图题号获取的是数字，需要转成英文大写
def switch_answer(item):
    kw = {
        1: 'A',
        2: 'B',
        3: 'C',
        4: 'D',
        5: 'E',
        6: 'F',
        7: 'G',
        8: 'H',
        9: 'I'
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
def write_data(title, url, wb):
    # 创建一页
    sheet = create_tab(title, wb)

    data = requests.get(url, headers=header).json()
    questions = data.get('data')
    for i in range(0, len(questions)):
        print('正在写入 id = %s 的问题' % questions[i]['id'])

        # 题干
        question = questions[i]['stem']
        sheet.write(i + 1, 0, replace_html_tag(question))
        # 答案
        answer = questions[i]['answer']
        sheet.write(i + 1, 1, switch_answer(answer))
        # 选项
        options = questions[i]['choices']
        f = ''
        for j in range(0, len(options)):
            f += '%s:%s  ' % (get_options(j), options[j])
        sheet.write(i + 1, 2, replace_html_tag(f))
        # 解析
        analysis = questions[i]['analysis']
        sheet.write(i + 1, 3, replace_html_tag(analysis))
        # 来源
        source = questions[i]['from']
        sheet.write(i + 1, 4, source)
        # 单/多选
        teach_type = questions[i]['teachType']
        sheet.write(i + 1, 5, teach_type)
        # 分类
        question_type = questions[i]['pointList'][0]['pointsName']
        type_str = ''
        for j in range(0, len(question_type)):
            if j != len(question_type) - 1:
                type_str += '%s - ' % question_type[j]
            else:
                type_str += question_type[j]
        sheet.write(i + 1, 6, type_str)
        # 资料
        materials = questions[i]['materials']
        material_str = ''
        for j in range(0, len(materials)):
            material_str += '%s \r\n' % materials[j]
        sheet.write(i + 1, 7, replace_html_tag(material_str))

        print('id = %s 写入完成' % questions[i]['id'])
        print('题库总数 = %s' % i)


# 写入excel
def create_excel(title, url):
    wb = xlwt.Workbook()
    write_data(title, url, wb)
    wb.save('data/%s.xls' % title)


def get_bank_list(list_url):
    list_data = requests.get(list_url, headers=header).json()
    result = list_data.get('data').get('result')
    for i in range(0, len(result)):
        question_aar = result[i].get('questions')
        question_str = ','.join(str(j) for j in question_aar)
        bank_name = result[i].get('name')
        request_url = question_url + question_str
        create_excel(bank_name, request_url)


def start_spider(list_url):
    get_bank_list(list_url)


# 华图题库数据抓包
if __name__ == '__main__':
    # start_spider(country_list_url)
    start_spider(province_list_url)
