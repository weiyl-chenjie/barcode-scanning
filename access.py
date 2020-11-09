#encoding=utf-8
import pyodbc
class ODBC_MS:
    def __init__(self, DRIVER, DBQ):
        self.DRIVER = DRIVER
        self.DBQ = DBQ

    def __GetConnect(self):
        if not self.DRIVER:
            raise (NameError, "no setting DRIVER info")
        self.conn = pyodbc.connect(DRIVER=self.DRIVER, DBQ=self.DBQ)
        crsr = self.conn.cursor()
        if not crsr:
            raise (NameError, "connected failed!")
        else:
            return crsr

    def select_query(self, sql):
        crsr = self.__GetConnect()
        crsr.execute(sql)
        rows = crsr.fetchall()
        crsr.close()
        self.conn.close()
        return rows

    def update_query(self, sql):
        crsr = self.__GetConnect()
        crsr.execute(sql)
        self.conn.commit()
        crsr.close()
        self.conn.close()

    def insert_query(self, sql):
        crsr = self.__GetConnect()
        crsr.execute(sql)
        self.conn.commit()
        crsr.close()
        self.conn.close()

    def select_status_query(self, sql):  # 返回DB_ezCad表中第一行的status的值
        crsr = self.__GetConnect()
        crsr.execute(sql)
        status = crsr.fetchone().status
        crsr.close()
        self.conn.close()
        return status

    def select_one_query(self, sql):  # 返回DB_ezCad表中某一行的keycode的值
        crsr = self.__GetConnect()
        crsr.execute(sql)
        row = list(crsr.fetchone())
        crsr.close()
        self.conn.close()
        return row

