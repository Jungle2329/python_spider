import os
import re
from functools import reduce

import requests
import xlwt
from bs4 import BeautifulSoup as bs
from requests.cookies import RequestsCookieJar

from web.DownloadImage import download_img

user_agent = 'Mozilla/5.0 (Windows NT 6.1; Win64; x64; Trident/7.0; rv:11.0) like Gecko'

header = {
    "User-Agent": user_agent
}

cookies = '1pqA_2132_saltkey=efXchfXC; 1pqA_2132_lastvisit=1561599634;' \
          ' 1pqA_2132_lastact=1561603263%09plugin.php%09; 1pqA_2132_sendmail=1;' \
          ' 1pqA_2132_ulastactivity=1561603262%7C0; 1pqA_2132_auth=dd38ojIxbrKd5q0' \
          'y9b%2FO4a0HkIHYp8Otonhtm5wB9znZXbM4QgPyoS6osZPiY2TTa9hLI%2BbVMWCbAF8rtXFdp' \
          '9nQsbeK; 1pqA_2132_lastcheckfeed=1290934%7C1561603262; 1pqA_2132_check' \
          'follow=1; 1pqA_2132_lip=110.178.211.35%2C1561603262'
d = dict(map(lambda x: x.split('='), cookies.split(';')))
# 添加登录需要的cookie
cookie_jar = RequestsCookieJar()
for k, v in d.items():
    cookie_jar.set(k, v)

base_url = 'https://bbs.shiyebian.org'

# 获取题的解析
question_result = '/plugin.php?id=kaoshi:showbox&from=exam&mod=showcontent' \
                  '&infloat=yes&handlekey=showbox&inajax=1&ajaxtarget=fwin_content_showbox&cid='
# 历年真题
true_exam = "/plugin.php?id=kaoshi:list&tid=92&sid=79&mode=2&page="

# 真题
exam = "/plugin.php?id=kaoshi:paper&sid=79&pid="

# 模块特训
model_chapter = '/plugin.php?id=kaoshi:chapter&sid=79'

# 模块下的分类
model_list = '/plugin.php?id=kaoshi:list&sid=79&cid=62'

# 模块下的分类下的题
model_paper = '/plugin.php?id=kaoshi:paper&sid=79&pid=11249&cid=62'


def start_spider(list_url):
    for i in range(1, 18):
        if i == 10:
            return
        req = requests.get(base_url + true_exam + str(i), headers=header, cookies=cookie_jar)
        soup = bs(str(req.content, 'GBK'), 'html.parser')
        for data in re.findall(r'pid=([0-9]*)', str(soup.find('ul', class_='ul-li'))):
            get_bank_list(list_url + str(data))


def get_bank_list(list_url):
    req = requests.get(list_url, headers=header, cookies=cookie_jar)
    if req.status_code == 200:
        soup = bs(str(req.content, 'GBK'), 'html.parser')
        # <div class="paper-bt">2019年浙江省公务员录用考试《行测》真题及解析（B类）</div>
        bank_name = soup.find_all('div', class_='paper-bt')[0].get_text()
        create_excel(bank_name, soup)


# 写入excel
def create_excel(title, soup):
    os.makedirs('./data/', exist_ok=True)
    wb = xlwt.Workbook()
    write_data(title, soup, wb)
    wb.save('./data/%s.xls' % (title))


# 写入数据
def write_data(title, soup, wb):
    # 创建一页
    sheet = create_tab(wb)

    for i, value in enumerate(soup.find_all('div', class_='exam-box')):
        if i == 2:
            return

        print('正在写入 eid = %s 的问题' % value.map['eid'])
        # 保存该题的所有图片
        for image_data in value.find_all('img'):
            download_img(image_data['src'])

        # 题干
        question = replace_html_tag(
            reduce(lambda x, y: str(x) + str(y), value.find('div', class_='exam-subject').children))
        sheet.write(i + 1, 0, question)

        # 答案
        answer = value.find('span', class_='result').get_text()
        sheet.write(i + 1, 1, answer)

        # 选项
        option_labels = value.find('ul', class_='exam-options').find_all('label')
        options = ''
        for j in option_labels:
            options += j.get_text()
        sheet.write(i + 1, 2, replace_html_tag(options))

        # 解析
        # 题id
        eid = value.map['eid']
        # 单独创建请求获取解析
        req = requests.get(base_url + question_result + eid, headers=header, cookies=cookie_jar)
        if req.status_code == 200:
            soup = bs(replace_cdata(str(req.content, 'GBK')), 'html.parser')
            # 先下载下来图片
            for image_data in soup.find_all('img'):
                download_img(image_data['src'])
            for data in soup.find_all('b'):
                if data.get_text().find('解析') == 1:
                    s = ''
                    for j in data.parent.contents:
                        if str(j).find('<b>') == -1 & str(j).find('<br/>') == -1:
                            s += str(j)
                    sheet.write(i + 1, 3, s)
        print('写入完成 eid = %s 的问题' % value.map['eid'])


# 创建单页第一行的标题
def create_tab(wb):
    sheet = wb.add_sheet("卷1", cell_overwrite_ok=True)
    sheet.write(0, 0, '题干')
    sheet.write(0, 1, '答案')
    sheet.write(0, 2, '选项')
    sheet.write(0, 3, '解析')
    return sheet


# 替换html标签
def replace_html_tag(s):
    result = s.replace('<p>', '') \
        .replace('</p>', '') \
        .replace('<br/>', '\r\n')
    # <u>替换下划线</u>
    return re.sub(re.compile(r"<u.*?</u>", re.S), "_______", result)


def replace_cdata(s):
    return s.replace("<![CDATA[", "").replace(']]>', '')


if __name__ == '__main__':
    start_spider(base_url + exam)
