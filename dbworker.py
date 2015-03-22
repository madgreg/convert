#coding: utf-8

import sqlite3 as sqlite3
import sys


class DBWorker:


    def __init__(self, db_path):
        "Func connect bd "

        self.text_select = u'''select inn,nameorg,sklad,vigr,sohr
                      from vigr
                       order by inn'''
        self.firststr = 10
        self.skipstr = 0
        self.kolvostr = [0,0]
        self.tekzap = -1

        #= kinterbasdb.connect(dsn=db_path, user='SYSDBA', password='masterkey',  charset='WIN1251', role='NONE', dialect=3)
        self.connect=sqlite3.connect(db_path, isolation_level=None, detect_types=sqlite3.PARSE_DECLTYPES|sqlite3.PARSE_COLNAMES)
        #con = sqlite3.connect(self.DB_PATH, isolation_level='DEFERRED', detect_types=sqlite3.PARSE_DECLTYPES|sqlite3.PARSE_COLNAMES)

        self.connect.row_factory = sqlite3.Row

        self.cursor=self.connect.cursor()
        self.cursor.execute('PRAGMA journal_mode = MEMORY;')
        self.cursor.execute('PRAGMA synchronous = OFF;')


    def field_dokn(self):
        if self.kolvostr[0] == 0:
            self.count_dokn()
        sel = sel= self.text_select.upper()
        pos = sel.find('SELECT')
        #~ if pos >= 0 :
            #~ selr = sel[pos+6:]
            #~ print sel
            #~ print selr

        #sel = sel+''' limit  %s offset %s''' % (self.firststr, self.skipstr)

        self.cursor.execute(sel)
        #return (f for f in self.cursor.description)
        return self.cursor.description

    def sel_dokn(self):
        #self.cursor.execute(u'select * from SPR_TOVARY')
        return self.cursor.fetchone()

    def count_dokn(self):
        "Функция подсчета количества строк в запросе"
        self.kolvostr = (-1,)
        sel= self.text_select.upper()
        pos = sel.find('SELECT')
        if pos >= 0 :
            pos = sel.find('FROM')
            if pos >= 0 :
                sel = 'select count(*) ' + sel[pos:]
                pos = sel.find('ORDER')
                if pos >= 0 :
                    sel = sel[:pos]
                pos = sel.find('GROUP')
                if pos >= 0 :
                    sel = sel[:pos]
                self.cursor.execute(sel)
                self.kolvostr = self.cursor.fetchone()
        return self.kolvostr

    def updateVigr(self,strUpd,QueryParams):
        try:
            self.cursor.execute(strUpd,QueryParams)
        except sqlite3.DatabaseError,x:
            print "Oshibka:  ",x
            print strUpd
    
    def ttest(self, inn):
        self.cursor.execute(u'select * from dokn')
        arr = self.cursor.fetchall()
  
        self.cursor.execute(u'select inn,nameorg,sklad,vigr,sohr from vigr order by inn')
        arr1 = self.cursor.fetchall()
        i = 0
        for rec in arr:
            tmpbool = False
            for rec1 in arr1:
                if rec[5] == rec1[1] and rec[7] == rec1[2]:
                    tmpbool = True
            if tmpbool == False:
                i += 1
                try:
                    self.cursor.execute(u"""INSERT INTO vigr VALUES(%s, '%s', '%s', 1, 1)""" % (inn, rec[5], rec[7]) ) 
                    self.connect.commit()
                except:
                    continue
        return i
                
if __name__ == '__main__':
    d = DBWorker('e:/work/convert xls/ConvBD.db3')
    #d.field_dokn()
    print d.ttest('7107000373')



