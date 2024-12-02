# -*- coding = utf-8 -*-
from bs4 import BeautifulSoup
import re
import requests
from openpyxl import Workbook


#通过正则表达式对https源码进行匹配并创建对象
findLink = re.compile(r'<a href="(.*?)">')
findImgSrc = re.compile(r'<img.*src="(.*?)"', re.S)
findTitle = re.compile(r'<span class="title">(.*)</span>')
findRating = re.compile(r'<span class="rating_num" property="v:average">(.*)</span>')
findJudge = re.compile(r'<span>"(.\d*)"人评价</span>')
findInq = re.compile(r'<span class="inq">(.*)</span>')
findBd = re.compile(r'<p class="">(.*?)</p>', re.S)


#主函数 分别为 1爬取网页 2创建xlsx 3存储数据
def main():
    baseurl = "https://movie.douban.com/top250"
    datalist = getData(baseurl)
    savepath = "doubanTop250.xlsx"
    saveData(datalist, savepath)
    print("done")


#爬取网页函数
def getData(baseurl):
    datalist = []  #创建列表储存数据
    for i in range(0, 10):  #利用迭代翻页收集数据
        url = f"{baseurl}?start={i * 25}"  #豆瓣电影 Top 250 的分页 URL
        print(f"Fetching URL: {url}")  #提示成功获取到URL
        html = askURL(url)  #将新读取url的html内容存为新的html方便创建正则表达式对象并进行匹配
        if not html:  #防止askURL(url)返回空字符串及请求失败的情况，对其返回值进行检查
            print(f"Failed to fetch HTML for URL: {url}")
            continue
        soup = BeautifulSoup(html, "html.parser")

        for item in soup.find_all('div', class_='item'):
            try:
                data = []
                item = str(item)

                link = re.findall(findLink, item)
                data.append(link[0] if link else "N/A")

                imgSrc = re.findall(findImgSrc, item)
                data.append(imgSrc[0] if imgSrc else "N/A")

                titles = re.findall(findTitle, item)
                if len(titles) == 2:
                    data.append(titles[0])
                    data.append(titles[1].replace("/", ""))
                elif len(titles) == 1:
                    data.append(titles[0])
                    data.append("N/A")
                else:
                    data.append("N/A")
                    data.append("N/A")

                rating = re.findall(findRating, item)
                data.append(rating[0] if rating else "N/A")

                judgeNum = re.findall(findJudge, item)
                data.append(judgeNum[0] if judgeNum else "N/A")

                inq = re.findall(findInq, item)
                data.append(inq[0].replace('。', '').strip() if inq else 'N/A')

                bd = re.findall(findBd, item)[0]
                bd_cleaned = re.sub(r'<br(\s+)?/>(\s+)?', '', bd).strip() if bd else 'N/A'
                data.append(bd_cleaned)

                datalist.append(data)
            except Exception as e:
                print(f"Error parsing item: {e}")
                continue
    print(f"Fetched {len(datalist)} items")
    return datalist


def askURL(url):
    header = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36'
    }
    #request = urllib.request.Request(url, headers=head)  # 创建一个http请求对象
    html = ""  # 初始化一个空字符串html用于储存目标网页的HTML数据
    try:
        response = requests.get(url, headers=header, timeout=10) # 发送http请求并获得响应（使用requests库）
        response.raise_for_status() # 检查http响应码
        html = response.text
        return html
        # response = urllib.request.urlopen(request)  # 发送http请求并获得响应 （使用urllib库）
        # html = response.read.decode('utf-8')  # 将HTML数据解码为utf-8字节格式并读取内容
    except requests.exceptions.RequestException as e:
        print(f"Error fetching URL {url}: {e}")

    """(使用urllib库进行错误提示)
    except urllib.error.HTTPError as e:
        print(f"Error fetching URL: {url}")  # 加入错误提示
        if hasattr(e, 'code'):  # 检查是否有code属性并返回状态码
            print(e.code)
        if hasattr(e, 'reason'):  # 检查是否有出错的原因并返回
            print(e.reason)
    return html  # 返回读取到的html内容
    """

def saveData(datalist, savepath):
    print("Saving data...")

    try:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Movie of Douban Top250"

        #创建列名
        col = ("电影详情链接", "图片链接", "影片中文名", "影片外国名", "评分", "评价数", "概况", "相关信息")
        sheet.append(col)

        #写入数据
        for data in datalist:
            sheet.append(data)

        #保存文件
        workbook.save(savepath)
        print(f"Data saved successfully to {savepath}")
    except Exception as e:
        print(f"An error occurred while saving dataL {savepath}: {e}")

""""
def saveData(datalist, savepath):
    print("save...")
    book = xlwt.Workbook(encoding='utf-8', style_compression=0)
    sheet = book.add_sheet('Movie of DouBan Top250', cell_overwrite_ok=True)
    col = ("电影详情链接", "图片链接", "影片中文名", "影片外国名", "评分", "评价数", "概况", "相关信息")
    for i in range(0, 8):
        sheet.write(0, i, col[i])
    for i in range(0, 250):
        data = datalist[i]
        for j in range(0, 8):
            sheet.write(i + 1, j, data[j])
    book.save(savepath)
"""
if __name__ == "__main__":
    main()
    print("done")