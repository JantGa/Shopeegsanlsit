#-*- coding = utf-8 -*-
import urllib.request
import json 
import xlrd
import xlwings as xw

#商品信息参数
def main():
    print('开始读取文件数据...')
    #关键词文件
    fileMZ = 'shopee标题.xls'
    keyword = listdata(fileMZ)
    html = []
    for length in range (1,len(keyword)):
        htmlLs = getData(keyword[length])
        html.extend(htmlLs)
        print(html)

    #保存文件名
    savepath = 'E:\CodeProjects\shopee关键词.xls'
    saveData(html,savepath)

#获取关键词列表
def listdata(filename):
    listData = xlrd.open_workbook(filename)
    table = listData.sheet_by_index(0)
    list = table.col_values(0, start_rowx=0, end_rowx=None)
    return list
    
#获取指定链接解析
def getData(Keyword):
    word = urllib.parse.quote(Keyword)
    url = 'https://www.dny001.com/f/tools/keywordsSearch?country=1&keyword=' + str(word)+'&pageNum=NaN&orderByColumn=&isAsc=asc&_=1608317535961'
    text = askURL(url)
    title = explain(text)
    return title

#获取指定链接内容
def askURL(url):
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
    data = json.loads(html)
    result = [(item.get('keyword','NA'),item.get('searchVolume','NA'),item.get('recommendPrice','NA')) for item in data['rows']]
    print (result)
    return result


#保存数据
def saveData(html,savepath):
    app = xw.App(visible=False,add_book=False)
    app.display_alerts = False
    app.screen_updating = False
    wb = app.books.add()
    sheet = wb.sheets.add('数据分析')
    col = ('标题','热度','推荐出价','商品数')
    for x in range(0,4):
        sheet.range(1,x+1).value = col[x]
    for i in range(1,len(html)):
        for j in range(0,3):
            data = html[i]
            sheet.range(i+1,j+1).value = data[j]
    wb.save(savepath)
    wb.close()

if __name__ == "__main__":
    main()