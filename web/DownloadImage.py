import os
import re
from functools import reduce

import requests


# img_path 完整url路径
def download_img(img_path):
    print(img_path)
    req_img = requests.get(img_path)
    img_path = img_path.replace('https://zhenti.oss-cn-qingdao.aliyuncs.com', './data/image')
    img_path = img_path.replace('https://360kao.oss-cn-hangzhou.aliyuncs.com', './data/image')
    path = img_path.replace('/' + img_path.split('/')[-1], '')
    os.makedirs(path, exist_ok=True)
    with open(img_path, 'wb') as f:
        f.write(req_img.content)
    return img_path


def get_img_url(img_tag):
    return img_tag['src']
