import requests
import re
from urllib import parse


session = requests.Session(); ##session 对象

def getCourtInfo(DocID):
    """
    根据文书DocID获取相关信息：标题、时间、浏览次数、内容等详细信息
    """
    url = 'http://wenshu.court.gov.cn/CreateContentJS/CreateContentJS.aspx?DocID={0}'.format(DocID)
    headers = {
        'Host':'wenshu.court.gov.cn',
        'Origin':'http://wenshu.court.gov.cn',
        'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.84 Safari/537.36',
    }## 构造报文头
    req = session.get(url,headers=headers) ## 传入头部参数
    req.encoding = 'uttf-8' ## 设定编码格式
    return_data = req.text.replace('\\','') ## 把返回的 \\ 号替换成空格
    read_count = re.findall(r'"浏览\：(\d*)次"',return_data)[0]
    court_title = re.findall(r'\"Title\"\:\"(.*?)\"',return_data)[0]
    court_date = re.findall(r'\"PubDate\"\:\"(.*?)\"',return_data)[0]
    court_content = re.findall(r'\"Html\"\:\"(.*?)\"',return_data)[0]
    return [court_title,court_date,read_count,court_content]

def download(DocID):
    """
    根据文书DocID下载doc文档
    """
    courtInfo = getCourtInfo(DocID)
    print(courtInfo)
    url = 'http://wenshu.court.gov.cn/Content/GetHtml2Word'
    headers = {
        'Host':'wenshu.court.gov.cn',
        'Origin':'http://wenshu.court.gov.cn',
        'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.84 Safari/537.36',
    }
    fp = open('content.html','r',encoding='utf-8')
    htmlStr = fp.read()
    #print(htmlStr)
    fp.close()
    htmlStr = htmlStr.replace('court_title',courtInfo[0]).replace('court_date',courtInfo[1]).\
        replace('read_count',courtInfo[2]).replace('court_content',courtInfo[3])
    #print(htmlStr)
    htmlName = courtInfo[0]
    data = {
        'htmlStr':parse.quote(htmlStr),
        'htmlName':parse.quote(htmlName),
        'DocID':DocID
    }
    req = session.post(url,headers=headers,data=data)
    filename = './download/{}.doc'.format(htmlName)
    fp = open('{}.doc'.format(htmlName),'wb')
    fp.write(req.content)
    fp.close()
    print('"{}"文件下载完成...'.format(filename))

download('532bd8ed-4ba8-48b7-ad70-0063f64ede05')