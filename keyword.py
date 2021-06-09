#-*- coding = utf-8 -*-
import urllib.request
import json
import xlwings as xw

#扫描关键词下所有商品标题
#商品信息参数
def main():
    print ("开始爬取中...")
    Keyword = 'iphone 12 手機殼'
    #取第X页
    page = 3 #页码
    html = getData(Keyword,page)
    savepath = 'E:\CodeProjects\shopee标题.xls'  #保存路径
    saveData(html,savepath)

#获取指定链接解析
def getData(Keyword,page):
    print ('正在解析链接...')
    newest = (page-1)*50
    word = urllib.parse.quote(Keyword)
    url = 'https://xiapi.xiapibuy.com/api/v2/search_items/?by=relevancy&keyword=' + str(word)+'&limit=50&newest='+str(newest)+'&order=desc&page_type=search&version=2'
    text = askURL(url)
    title = explain(text)
    return title

#获取指定链接内容
def askURL(url):
    print("正在获取内容...")
    headers = {
    'User-Agent': "User-Agent,Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36"
    }
    request = urllib.request.Request(url,headers=headers)
    datalist = ""
    try:
        response = urllib.request.urlopen(request)
        datalist = response.read().decode('utf-8')
    except urllib.error.URLError as e:
        print(e)
    return datalist

#解析数据
def explain(html):
    print("正在解析数据...")
    data = json.loads(html)
    result = [item.get('name','NA') for item in data['items']]
    return result

#保存数据
def saveData(html,savepath):
    print("正在保存...")
    app = xw.App(visible=False,add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    wb = app.books.add()
    sheet = wb.sheets.add('虾皮标题')
    sheet.range('A1').value = '标题'
    for i in range(0,len(html)):
        print ("第%d条" %(i+1))
        sheet.range(i+1,1).value = html[i]
    wb.save(savepath)
    wb.close()

if __name__ == "__main__":
    main()