import requests

from bs4 import BeautifulSoup

import json

import xlsxwriter


def get_data(pname, pid):
    base_url = 'http://www.mafengwo.cn'

    headers = {
        'Connection': 'keep-alive',
        'Upgrade-Insecure-Requests': '1',
        'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_3) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/73.0.3683.75 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3',
        'Referer': 'http://www.mafengwo.cn/yj/10267/2-0-1.html',
        'Accept-Encoding': 'gzip, deflate',
        'Accept-Language': 'zh-CN,zh;q=0.9,en-US;q=0.8,en;q=0.7',
    }
    id_list = []
    fail_list = []
    # 获取总页数
    resp = requests.get(base_url + '/yj/%s/2-0-1.html' % str(pid), headers=headers)

    soup = BeautifulSoup(resp.text)
    total_count = soup.find('span', 'count').get_text('-').split('-')[1]

    #
    def get_detail(p):
        resp = requests.get('http://www.mafengwo.cn/yj/%s/2-0-%d.html' % (str(pid), p), headers=headers)
        soup = BeautifulSoup(resp.text)
        post_list = soup.find_all('li', 'post-item clearfix')

        l = []

        for post in post_list:
            author = post.find('span', 'author').find_all('a')[-1].get_text()
            title = post.find('a', 'title-link').get_text()
            publish_date = post.find('span', 'comment-date').get_text()
            pvc = post.find('span', 'status').get_text('/')
            pv, comment_count = pvc.split('/')
            data = {
                'author': author,
                'title': title,
                'publish_date': publish_date,
                'pv': pv,
                'comment_count': comment_count,
            }
            l.append(data)

        return l

    for i in range(1, int(total_count) + 1):
        try:
            id_list += get_detail(i)
            print(i, 'success...')
        except Exception as e:
            print(i, 'fail...', e)
            fail_list.append(i)

    with open('%s.json' % pname, 'w') as f1:
        json.dump(id_list, f1, ensure_ascii=False)

    with open('fail%s.json' % pname, 'w') as f2:
        json.dump(fail_list, f2, ensure_ascii=False)


def save_excel(pname):
    f = open('%s.json' % pname, 'r')

    data = json.load(f)

    result = {}

    for d in data:
        date = d.get('publish_date')[:7]
        if not result.get(date):
            result[date] = {
                'total_count': 0,
                'total_pv': 0,
                'total_comment': 0,
                'year': date[:4],
                'month': date[5:]
            }
        result[date]['total_count'] += 1
        result[date]['total_pv'] += int(d.get('pv'))
        result[date]['total_comment'] += int(d.get('comment_count'))

    for k, v in result.items():
        v['avg_pv'] = '%.2f' % (v['total_pv'] / v['total_count'])
        v['avg_comment'] = '%.2f' % (v['total_comment'] / v['total_count'])

    workbook = xlsxwriter.Workbook('%s.xlsx' % pname)
    worksheet = workbook.add_worksheet()
    row = 1
    worksheet.write_row(0, 0, ['年份', '月份', '游记总数', '浏览总数', '评论总数', '浏览平均数', '评论平均数'])
    for k, v in result.items():
        worksheet.write_row(row, 0, [
            v['year'],
            v['month'],
            v['total_count'],
            v['total_pv'],
            v['total_comment'],
            v['avg_pv'],
            v['avg_comment']
        ])
        row += 1
    workbook.close()


plist = [
    {
        'pname': '张家界',
        'pid': '10267',
    },
    {
        'pname': '峨眉山',
        'pid': '10143',
    },
    {
        'pname': '桂林旅游',
        'pid': '10095',
    },
    {
        'pname': '黄山旅游',
        'pid': '10440',
    },

    {
        'pname': '天目湖',
        'pid': '50311',
    },
    {
        'pname': '九华山',
        'pid': '11264',
    },{
        'pname': '宋城演艺',
        'pid': '20650',
    },
    {
        'pname': '大连圣亚',
        'pid': '10301',
    },
    {
        'pname': '西安曲江文旅',
        'pid': '10301',
    },
    {
        'pname': '丽江旅游',
        'pid': '10186',
    },
    
]

for p in plist[0:]:
    get_data(p['pname'], p['pid'])
    save_excel(p['pname'])
