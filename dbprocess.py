import logging
import os
import re
from datetime import datetime

import yaml

import dbconfig
import getAppPath
from db_lib import MySQLDataBase
from loadconfig import getconfig


class DataBase:
    def __init__(self):
        self.blogger = logging.getLogger("mAnalyser")
        app_path = getAppPath.getapppath()
        self.configFolder, self.breakDownConfig = getconfig()
        db_yaml = os.path.join(self.configFolder, "config\\dbConfig.yaml")
        with open(db_yaml, "r") as stream:
            self.db_config = yaml.safe_load(stream)
        self.isDbConfig = self.db_config['db_system']


    # def add_db(self, customer, ms_id, jrn_id, art_id, unique_id):
    #     add_sql = f"INSERT INTO breakdown_data(Customer_name, Manuscript_id, Journal_id, Article_id, Unique_id) VALUES ('{customer}', '{ms_id}', '{jrn_id}', '{art_id}', '{unique_id}')"
    #     dbs = MySQLDataBase()
    #     conn = dbs.connect(dbconfig.read_db_config())
    #     cur = conn.cursor()
    #     try:
    #         cur.execute(add_sql)
    #         conn.commit()
    #         self.blogger.info(f"Customer: {customer}, MS_NO: {ms_id}, Journal ID: {jrn_id}, Article ID: {art_id}, "
    #                           f"Unique id: {unique_id} added in db")
    #     except Exception as e:
    #         print(e)
    #         conn.rollback()
    #     cur.close()
    #     conn.close()

    def add_db(self, package_id, unique_id, date, time):
        if self.isDbConfig is True:
            add_sql = f"INSERT INTO breakdown_data(Package_name, unique_id, Date, Time) VALUES ('{package_id}', '{unique_id}', '{date}', '{time}')"
            dbs = MySQLDataBase()
            conn = dbs.connect(dbconfig.read_db_config())
            cur = conn.cursor()
            try:
                cur.execute(add_sql)
                conn.commit()
                self.blogger.info(f"Package Name: {package_id}, unique_id: {unique_id}, Date: {date}, Time: {time}")
            except Exception as e:
                print(e)
                conn.rollback()
            cur.close()
            conn.close()

    def update_db(self, unique_id, process_name, process_status, info_log, error_log):
        if self.isDbConfig is True:
            date = datetime.today().strftime('%Y-%m-%d')
            time = datetime.today().strftime('%H:%M:%S')
            print(f"{unique_id}, {process_name}, {process_status}, {date}: {time}")
            pattern = r"[a-z0-9]+"
            match = re.search(pattern, unique_id, re.IGNORECASE)
            if match:
                table_name = dbconfig.read_table_config(process_name)
                update_sql = f"UPDATE breakdown_data SET {process_name} = '{process_status}' WHERE Unique_id = '{unique_id}'"
                process_sql = f"INSERT INTO {table_name}(unique_id, log_info, log_error) VALUES ('{unique_id}', '{info_log}', '{error_log}')"
                # conn = MySQLDataBase().connection(dbconfig.read_db_config())
                conn = MySQLDataBase().connect(dbconfig.read_db_config())
                cursor = conn.cursor()
                try:
                    cursor.execute(update_sql)
                    cursor.execute(process_sql)
                    conn.commit()
                except Exception as e:
                    print(e)
                    conn.rollback()
                cursor.close()
                conn.close()

    def update_data(self, customer, ms_no, journal_id, article_id, unique_id):
        if self.isDbConfig is True:
            update_sql = f"UPDATE breakdown_data SET Customer_name = '{customer}', Manuscript_id = '{ms_no}', Journal_id = '{journal_id}', Article_id = '{article_id}' WHERE unique_id = '{unique_id}'"
            conn = MySQLDataBase().connect(dbconfig.read_db_config())
            cursor = conn.cursor()
            try:
                cursor.execute(update_sql)
                conn.commit()
            except Exception as e:
                print(e)
                conn.rollback()
            cursor.close()
            conn.close()

    def update_remark(self, unique_id, error_log):
        if self.isDbConfig is True:
            update_sql = f"UPDATE breakdown_data SET Remark = '{error_log}' WHERE Unique_id = '{unique_id}'"
            conn = MySQLDataBase().connect(dbconfig.read_db_config())
            cursor = conn.cursor()
            try:
                cursor.execute(update_sql)
                conn.commit()
            except Exception as e:
                print(e)
                conn.rollback()
            cursor.close()
            conn.close()

    def delete_db(self, unique_id, process_name, process_status):
        if self.isDbConfig is True:
            delete_sql = f"DELETE FROM breakdown_data WHERE WHERE Unique_id = '{unique_id}'"
            conn = MySQLDataBase().connection(dbconfig.read_db_config())
            cursor = conn.cursor()
            try:
                cursor.execute(delete_sql)
                conn.commit()
            except Exception as e:
                print(e)
                conn.rollback()
            cursor.close()
            conn.close()

    def get_uniqueid(self, jrn_id, art_id):
        if self.isDbConfig is True:
            uniq_id = ""
            sql = f"SELECT unique_id FROM breakdown_data WHERE Article_id = '{art_id}' and Journal_id = '{jrn_id}' ORDER BY id DESC LIMIT 1"
            conn = MySQLDataBase().connect(dbconfig.read_db_config())
            cursor = conn.cursor()
            try:
                cursor.execute(sql)
                value = cursor.fetchone()
                uniq_id = value[0]
            except Exception as e:
                print(e)
            conn.close()
            return uniq_id

# update_databas = DataBase()
# update_databas.update_data("SAGE","SADFS","MMS","1234","18b23ac6f8f54693a40990e8db3869a9")