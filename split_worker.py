# -*- encoding: utf-8 -*-
import sys
reload(sys)
sys.setdefaultencoding("utf-8")

from split_util import split_worker

from pymongo import MongoClient

mongo_con = MongoClient("103.29.133.171", 30001)

tasks = {
        "tianyu": {
            "db": "bbscrawl1",
            "collection": "comments",
            "find_": {"forum_id": "tianyu"}},

        "nubiya": {
            "db": "bbscrawl1",
            "collection": "comments",
            "find_": {"forum_id": "nubiya"}},

        "jiwu": {
            "db": "bbscrawl1",
            "collection": "comments",
            "find_": {"forum_id": "jiwu"}},
        
        "thl": {
            "db": "bbscrawl1",
            "collection": "comments",
            "find_": {"forum_id": "thl"}},

        "zhidao_xiaomi": {
            "db": "zhidaocrawl",
            "collection": "comments",
            "find_": {"forum_id": "xiaomi"}}
        }


def iterComments(task):
    mongo_db = mongo_con[task["db"]]
    mongo_collection_comment = mongo_db[task["collection"]]
    count = 10000
    for comment in mongo_collection_comment.find(task["find_"]):
        if "context" in comment:
            comment = comment.split("F")[0]
        if count:
            count = count - 1
            yield comment["comment"]
     

if __name__ == "__main__":
    split_worker(iterComments(tasks["zhidao_xiaomi"]))
