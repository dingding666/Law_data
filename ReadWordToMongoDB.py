import fnmatch
import pymongo
import win32com
from bson import ObjectId
from pymongo import MongoClient
from win32com import client as wc
import os
from win32com.client import Dispatch,constants
import pydoc
import json
import uuid


def connect_MongoDB():##连接数据库
    client = MongoClient('192.168.0.21', port=27017)
    db = client.get_database('Lawdata')
    mycol = db.get_collection('LawwordData')
    return db , mycol;##返回连接对象，以及集合
    #for doc in mycol.find():
     # print(doc)

# def read_word():#读取docx文件名称
#     w = win32com.client.Dispatch('Word.Application')
#
#     Path = "E:\Spider\Data"
#     doc_files = os.listdir(Path)
#     for doc in doc_files:
#         if os.path.splitext(doc)[1] == '.docx':
#             print(doc)
#             result = doc.to__json()
# str = "This is test word"
# meta = {
#     "title":"吕宗勇等非法吸收公众存款申诉、申请刑事通知书",
#     "creater":"刘宇宁",
#     "type":"通知书",
#     "publisher":"华南师范大学",
#     "data":"2018-08-12",
#     "language":"zh",
#     "data":str
# }


def insert_MongoDB_Wordname(collection_name,dataSet):##参数传递为被插入数据的集合以及数据
    try:
        collection = collection_name;
        collection.insert(dataSet)
    except:
        print("数据插入失败")

mypath = "E:\Spider\Data0.3"##文件路径


def DocxToTxt():##将docx文件转换成txt文件
    wordapp = win32com.client.gencache.EnsureDispatch("Word.Application")
    try:
        for root, dirs,files in os.walk(mypath):
            for _dir in dirs:
                pass
            for _file in files:
                if not fnmatch.fnmatch(_file,'*.docx'):
                    continue
                word_file = os.path.join(root, _file)
                wordapp.Documents.Open(word_file)
                docastxt = word_file[:-4] +'txt'
                wordapp.ActiveDocument.SaveAs(docastxt,FileFormat = win32com.client.constants.wdFormatText)
                wordapp.ActiveDocument.Close()
    finally:
        wordapp.Quit()##
mete = {
        #
        '_id': ObjectId(),
        '案件名称':'str0',
        '法院名称': 'str1',
        '通知书名称': 'str2',
        '审判号': 'str3',
        '被通知人名称':'str4',
        '通知书内容':' '
    }##插入数据格式

def Test_writeInMongoDB():#对单个文件的分段读取内容
    Name = '驳回杨飞申诉通知书.txt'
    path =  'E:\Spider\Data0.3\\'+ Name
    fr = open(path)
    lines = 1

    mete2 = {
        #
        '_id': ObjectId(),
        '案件名称': 'str0',
        '法院名称': 'str1',
        '通知书名称': 'str2',
        '审判号': 'str3',
        '被通知人名称': 'str4',
        '通知书内容': ' '
    }  ##插入数据格式

    for line in fr.readlines():  # 二层循环是按行读取txt文档中的每一行
        # print(line)##type = str
        if (lines == 1):
            print(line)
            mete['法院名称'] = line;
            mete['案件名称'] = Name
        if (lines == 2):
            print(line)
            mete['通知书名称'] = line;
        if (lines == 3):
            print(line)
            mete['审判号'] = line;
        if (lines == 4):
            print(line)
            mete['被通知人名称'] = line;
        if(lines > 4):
            print(line)
            mete['通知书内容'] += line;
        lines = lines + 1;

    print(mete2)

def Test_writeInMongoDB2(thispath,filesname):
    fr = open(thispath)
    lines = 1

    mete2 = {
        #
        '_id': ObjectId(),
        '案件名称': 'str0',
        '法院名称': 'str1',
        '通知书名称': 'str2',
        '审判号': 'str3',
        '被通知人名称': 'str4',
        '通知书内容': ' '
    }  ##插入数据格式

    for line in fr.readlines():  # 二层循环是按行读取txt文档中的每一行
        # print(line)##type = str
        if (lines == 1):
            if (line.find('法院') != -1):
                print(line)
                mete2['法院名称'] = line;
                mete2['案件名称'] = filesname
            else:
                break
        if (lines == 2):
            #print(line)
            mete2['通知书名称'] = line;
        if (lines == 3):
            #print(line)
            mete2['审判号'] = line;
        if (lines == 4):
            #print(line)
            mete2['被通知人名称'] = line;
        if (lines > 4):
            #print(line)
            mete2['通知书内容'] += line;
        lines = lines + 1;

    # print(mete2)
    # print(mete2['_id'])
    if(mete2['案件名称'] != 'str0'):
        MYCOL.insert(mete2)
    fr.close()



def ReadAllTxt():
    files = os.listdir(mypath)##遍历路径中的文件
    filesNum = 0
    for txt_files in files:
        if os.path.splitext(txt_files)[1] == '.txt':
            filepath = mypath+'\\'+txt_files
            # print(txt_files)
            Test_writeInMongoDB2(filepath,txt_files)##把路径和文件名当成参数传入
            filesNum = filesNum + 1
    print(filesNum)


def ReadTxt():#读取txt文件
    #DB, MYCOL = connect_MongoDB();
    doc_files = os.listdir(mypath)
    TxtFileNums = 0;# 记录文件数量
    for doc in doc_files:
        #遍历整个文件夹
        if os.path.splitext(doc)[1] == '.txt':
            #如果文件夹中文件后缀名为txt
            # print(doc)
            filename = mypath+'\\'+doc # 将该文件路径记录下来
            fr = open(filename)#打开该文件
            lines = 1 #记录在文件中的第几行
            #print(fr.read())
    #print(TxtFileNums)
            for line in fr.readlines():#二层循环是按行读s取txt文档中的每一行
                # print(line)##type = str
                print(lines)
                if (lines == 1):
                    flag = line.find('通知书')##p判断不规范通知书
                    if flag == -1:
                        mete['法院名称'] = line;
                        mete['_id'] = TxtFileNums
                        mete['案件名称'] = doc
                        TxtFileNums += 1
                    else:
                        break;
                elif (lines == 2):
                    mete['通知书名称'] = line;
                elif (lines == 3):
                    mete['审判号'] = line;
                elif (lines == 4):
                    mete['被通知人名称'] = line;
                else:
                    mete['通知书内容'] += line;
                lines = lines + 1;
            #print(mete)
        #print(mete)
            #print(TxtFileNums)
        #continue
      #MYCOL.insert(mete)
    #print(TxtFileNums)
    # print(type(mete))
    # json_data = json.dumps(mete)
    # print(json_data)
    # print(type(json_data))
##这个先放着，有BUG

DB,MYCOL = connect_MongoDB();
if __name__ =='__main__':
    #Test_writeInMongoDB()
    #ReadTxt()
    ReadAllTxt()
    #MYCOL.insert(mete)
    #DocxToTxt()
