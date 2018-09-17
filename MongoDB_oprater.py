import pymongo

class MongoDB(object):
    def __init__(self,host,port):
        self.host = host
        self.port = port

    def mongoDB_connect(self,dbName):
        client = pymongo.MongoClient(self.host,self.port)
        self.db = client[dbName]

    def mongoDB_query(self,collection_name,query={}):
        try:
            collection = self.db[collection_name]
        except:
            print("集合获取失败")

        try:
            result_obj = collection.find(query)

            result = []
            for item in result_obj:
                result.append(item)
            return result
        except:
            print("查询失败")

    def mongoDB_insert(self,collection_name,dataSet):
        try:
            collection = self.db[collection_name]
            collection.insert(dataSet)
        except:
            print("数据插入失败")
