#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import glob
import json
import logging
import datetime
import dateutil.parser
import elasticsearch
import hashlib
import pandas as pd

from elasticsearch.helpers import bulk


logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')


class Manager():
    def __init__(self):
        self.elasticsearch_host = "localhost"
        self.elasticsearch_port = "9200"

    def load(self, data):
        self.data = data

    def insert_into_es(self):
        es_hosts = [{
            'host': self.elasticsearch_host,
            'port': self.elasticsearch_port
        }]

        es_client = elasticsearch.Elasticsearch("http://localhost:9200")
        index_name = 'seven'

        if not es_client.indices.exists(index=index_name):
            es_client.indices.create(
                index=index_name,
                body={"settings": {"index": {"number_of_replicas": 0, "number_of_shards": 3}}}
            )

        def doc_generator(data: pd.DataFrame):
            for index, document in data.iterrows():
                document = document.to_dict()
                for k, v in document.items():
                    doc = {
                            'type': index,
                            'date': pd.to_datetime(k).isoformat(),
                            'value': v
                        }

                    mydata = {
                            '_index': index_name,
                            '_id': f"{pd.to_datetime(k).isoformat()}{index}",
                            '_source': doc
                            }
                    #print(mydata)
                    yield mydata

            return StopIteration

        for _ in range(5):
            # retry 5 times in case network issue
            try:
                bulk(es_client, doc_generator(self.data))
                break
            except Exception as e:
                logging.error(e)


def insert_into_es(data):
    manager = Manager()
    manager.load(data)
    # manager.dump()
    manager.insert_into_es()

def main():
    print(elasticsearch.__version__)
    data = pd.read_excel('杜子期血常规数据统计.xlsx')
    data.set_index('Unnamed: 0', inplace=True)
    data = data.filter(regex='^((?!_参考范围$).)*$', axis=0).astype(float)
    #print(data)
    insert_into_es(data)

main()
