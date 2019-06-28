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

# 申论题库列表
bank_list_url = 'https://ns.huatu.com/e/api/v1/mock/papers?page=1'
# 申论材料 需要后跟id
material_list_baseurl = 'https://ns.huatu.com/e/api/v1/paper/materialList/'
# 申论问题 需要后跟id
question_list_baseurl = 'https://ns.huatu.com/e/api/v2/paper/questionList/'


# 创建单页第一行的标题
def create_tab(title, wb):
    sheet = wb.add_sheet(title, cell_overwrite_ok=True)
    sheet.write(0, 0, '材料')
    sheet.write(0, 1, '问题')
    sheet.write(0, 2, '解析1')
    sheet.write(0, 3, '解析2')
    sheet.write(0, 4, '解析3')
    return sheet


# 替换html标签
def replace_html_tag(s):
    return s.replace('<p>', '').replace('</p>', '').replace('<br/>', '\r\n').replace('&nbsp;', ' ')


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
def write_data(title, material_url, question_url, wb):
    # 创建一页
    sheet = create_tab(title, wb)

    material_data = requests.get(material_url, headers=header).json()
    question_data = requests.get(question_url, headers=header).json()

    material_list = material_data['data']
    for i in range(len(material_list)):
        # 材料
        sheet.write(i + 1, 0, replace_html_tag(material_list[i]['content']))

    question_list = question_data['data']['essayQuestions']
    for i in range(len(question_list)):
        cur = question_list[i]
        # 问题
        question = cur['answerRequire']
        sheet.write(i + 1, 1, replace_html_tag(question))
        # 解析
        answer_list = cur['answerList']
        for j in range(len(answer_list)):
            sheet.write(i + 1, j + 2, replace_html_tag(answer_list[j]['answerComment']))

        print('题库总数 = %s' % i)


# 写入excel
def create_excel(title, material_url, question_url):
    wb = xlwt.Workbook()
    write_data(title, material_url, question_url, wb)
    wb.save('data/%s.xls' % title)


def get_bank_list(list_url):
    list_data = requests.get(list_url, headers=header).json()
    result = list_data.get('data').get('result')
    for i in range(0, len(result)):
        paper_name = result[i].get('paperName')
        paper_id = result[i].get('paperId')

        material_url = material_list_baseurl + str(paper_id)
        question_url = question_list_baseurl + str(paper_id)

        create_excel(paper_name, material_url, question_url)


def start_spider(list_url):
    get_bank_list(list_url)


# 华图申论模拟题数据抓包
if __name__ == '__main__':
    start_spider(bank_list_url)
