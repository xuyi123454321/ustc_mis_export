'''
Draft of script tool for exporting scores from ustc mis website.
Function limited
No any exception handling

Author: Xu Yi
Institute: Univeristy of Science and Technology of China
Email: turambar@mail.ustc.edu.cn
'''

import urllib3
import pandas as pd
from pyquery import PyQuery

def request_html(sessionid):
    '''
    send post request
    JSESSINID needed
    return html page
    '''

    header = {
        'Host': 'mis.teach.ustc.edu.cn',
        'Connection': 'keep-alive',
        'Content-Length': '42',
        'Cache-Control': 'max-age=0',
        'Origin': 'http://mis.teach.ustc.edu.cn',
        'Upgrade-Insecure-Requests': '1',
        'Content-Type': 'application/x-www-form-urlencoded',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/66.0.3359.181 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
        'DNT': '1',
        'Referer': 'http://mis.teach.ustc.edu.cn/initquerycjxx.do',
        'Accept-Encoding': 'gzip, deflate',
        'Accept-Language': 'en-US,en;q=0.9,zh-CN;q=0.8,zh;q=0.7',
        'Cookie': 'JSESSIONID=%s'%sessionid
    }

    http = urllib3.PoolManager(headers=header)

    r = http.request(
        'POST',
        'http://mis.teach.ustc.edu.cn/querycjxx.do',
        fields={'xuenian': ''} # semester not supported now
    )

    return r.data.decode('gbk')

def query_page(page):
    '''
    query page for information using PyQuery
    '''
    column_num = 8
    d = PyQuery(page)
    col_t = d('table')[1]
    t = d('table')[2]
    results = []
    col = []

    for td in col_t.find('tr').iterfind('td'):
        col.append(td.find('b').text)

    for tr in t.iterfind('tr'):
        row = []
        for td in tr.iterfind('td'):
            row.append(td.text)

        if len(row) == column_num:
            results.append(row)

    return pd.DataFrame(results, columns=col)

def write_excel(df):
    '''
    write DataFrame to excel
    '''
    xlsx = pd.ExcelWriter('mis.xlsx')
    df.to_excel(xlsx, '成绩')
    xlsx.save()

if __name__ == '__main__':
    sessionid = input('JSESSIONID = ')
    df = query_page(request_html(sessionid))
    write_excel(df)