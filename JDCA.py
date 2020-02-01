# coding:utf-8
import datetime
import json
import re
import requests
import time

import xlwings as xw

s = requests.session()
product_url = input('请粘贴需要抓取的商品详情页地址：')
product_id_find = re.compile('https://item.jd.com/(.*?).html').findall(product_url)
product_id = product_id_find[0]

start_content_page = 0  # 起始抓取页码,0是第一页

time_now = datetime.datetime.now().strftime('%Y%m%d%H%M')
save_file_name = './' + product_id + '京东评价' + time_now
# 抓取当前评论，可能是合并多个SKU的评论
url = 'https://sclub.jd.com/comment/productPageComments.action'

# 抓取当前SKU评论，相当于京东中勾选只看当前商品评论
url_sku = 'https://club.jd.com/comment/skuProductPageComments.action'

headers = {
    'user-agent': 'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/65.0.3325.181 '
                  'Safari/537.36',
    'referer': ''
}

data = {
    'productId': '',
    'score': 0,  # 抓取中评为2 差评为1 好评为3 所有为0
    'sortType': 5,  # 5表示按照推荐排序，6表示按照时间排序
    'pageSize': 10,  # 单页最多显示条数，最多支持10条
    'isShadowSku': 0,
    'page': 0,  # 对应抓取的页面0-99，京东只允许抓取100页，所以最多抓取1000条
    'fold': 1
}

data['productId'] = product_id
data['page'] = start_content_page
headers['referer'] = 'https://item.jd.com/' + product_id + '.html'

# 读取模板，Excel模板
# app = xw.App(visible=False, add_book=False)
# wb = app.books.open(r'./template.xlsx')
wb = xw.Book(r'./template.xlsx')
sheet_summary = wb.sheets['评价概况']

# 抓取评价概要信息
t = s.get(url, params=data, headers=headers).text
j = json.loads(t)
shop_link = 'https://item.jd.com/%s.html' % product_id
print('开始抓取：'+ shop_link)
sheet_summary.range('B1').add_hyperlink(shop_link, shop_link, '提示：点击访问商品详情页')
sheet_summary.range('B2').value = j['productCommentSummary']['goodRate']
sheet_summary.range('B3').value = j['productCommentSummary']['commentCount']
sheet_summary.range('B4').value = j['productCommentSummary']['goodCount']
sheet_summary.range('B5').value = j['productCommentSummary']['generalCount']
sheet_summary.range('B6').value = j['productCommentSummary']['poorCount']
sheet_summary.range('B7').value = j['productCommentSummary']['showCount']
sheet_summary.range('B8').value = j['productCommentSummary']['videoCount']
sheet_summary.range('B9').value = j['productCommentSummary']['afterCount']
sheet_summary.range('D2').value = j['productCommentSummary']['averageScore']
sheet_summary.range('D3').value = j['productCommentSummary']['score1Count']
sheet_summary.range('D4').value = j['productCommentSummary']['score2Count']
sheet_summary.range('D5').value = j['productCommentSummary']['score3Count']
sheet_summary.range('D6').value = j['productCommentSummary']['score4Count']
sheet_summary.range('D7').value = j['productCommentSummary']['score5Count']
sheet_summary.range('B10').value = j['imageListCount']
sheet_summary.range('D11').value = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')

wb.save(save_file_name)

sheets_array = ['好评', '中评', '差评']
data_score = 3
for sheet_num in sheets_array:
    sheet_comment = wb.sheets[sheet_num]
    sheet_comment.activate()
    data['score'] = data_score
    data_score -= 1
    line_num = 2
    try_time = 3  # 抓取为空时的重试次数
    data['page'] = 0
    while True:
        try:
            t = s.get(url, params=data, headers=headers).text
        except Exception as e:
            print(e)
            time.sleep(3)
            continue
        flag = re.search('content', t)
        print(flag)
        if flag == None:
            last_shop_page = data['page']
            print("抓取停止，最后页码为：" + str(last_shop_page))
            try_time -= 1
            if try_time == 0:
                break
        else:
            j = json.loads(t)
            maxPage = j['maxPage']
            commentSummary = j['comments']
            for comment in commentSummary:
                nickname = comment['nickname']  # 用户名
                creationTime = comment['creationTime']
                referenceTime = comment['referenceTime']  # 购买日期
                days = comment['days']
                score = comment['score']
                usefulVoteCount = comment['usefulVoteCount']
                replyCount = comment['replyCount']
                content = comment['content']  # 评论内容
                referenceName = comment['referenceName']
                productColor = comment['productColor']

                print('{}  {}\n{}\n{}\n'.format(nickname, creationTime, content, referenceTime))
                comment_range = 'A' + str(line_num)
                sheet_comment.range(comment_range).value = [nickname, creationTime, referenceTime, days, score,
                                                            usefulVoteCount, replyCount, content, referenceName,
                                                            productColor]
                line_num += 1
            print('正在抓取页面:' + str(data['page']) + '总页面：' + str(maxPage))
            data['page'] += 1
    #         # time.sleep(1)

sheet_summary.range('B1').add_hyperlink(shop_link, referenceName, '提示：点击访问商品详情页')
wb.sheets['评价概况'].activate()
wb.save(save_file_name)
# wb.close()
# app.quit()
