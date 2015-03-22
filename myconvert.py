#!/usr/bin/python
# -*- coding: utf-8 -*-

import sys

reload(sys)
sys.setdefaultencoding('utf8')


import os
#import simplejson as json
import datetime


def pyUpper(x):
    return x.upper()

class Convertor:
    IMPORT_PATH = "d:/work/convert/in"
    EXPORT_PATH = "d:/work/convert/out"
    PERIOD = 1

    def unZip(self):
        import zipfile 
        path = self.IMPORT_PATH
        print 'RUN UNZIP %s' % path
        try:
            files = os.listdir(path);
            zipfiles = filter(lambda x: x.endswith(('.zip', '.ZIP')), files)
	    print len(zipfiles)
            for t in zipfiles:
                #farr = t.split('+')
                zfile = zipfile.ZipFile( '%s/%s' % (path, t), "r" )
                for info in zfile.infolist():
                    fname = info.filename
                    data = zfile.read(fname)
                    filename = path + '/' + fname
                    fout = open(filename, "w")
                    fout.write(data)
                    fout.close()
                zfile.close()
                os.rename('%s/%s' % (path, t), '%s/%s' % (path, t)+'.old')
        except Exception, e:
                    print str(e)
        print 'END UNZIP %s' % path

    def RunWorkConv(self):
        
        path = self.IMPORT_PATH
        files = os.listdir(path);
        listFiles = filter(lambda x: x.endswith(('.txt', '.TXT')), files)
        for fName in listFiles:
            inn = fName.split('-')[1][0:-4]
            print "%s: %s" % (inn, self.getTime())
            address = ''
            seller = ''
           
            dt = ''            
            
            with open(path+"/"+fName, 'rt') as f:                
                for row in f.xreadlines():
                    tmparr = row.decode('1251').split('\t')
                    if len(tmparr) < 3:
                        continue
                    strana = ''
                    proizv = ''
                    if tmparr[1] != "ID_DOK" and tmparr[1] != "ID_PARTIYA":
                        if tmparr[0] == "DOKN":
                            try:
                                outf.close()
                            except:
                                pass
                            address = tmparr[8]
                            seller = tmparr[6]
                            dt = tmparr[3]
                            outname = self.EXPORT_PATH+"/" + u"Отчет по " + inn + "_" + address.replace('"', '').replace('/', '').replace('\\', "").replace('?', "") + u" за "+ str(self.PERIOD) + u" месяц.csv"
                            if os.path.isfile(outname):
                                outf = open(outname, 'a')
                            else:
                                outf = open(outname, 'a')
                                outf.write("Наименование;Кол-во;Цена пост;Цена розн;Дата;Поставщик;Производитель;Страна\n".encode('1251'))
                        if tmparr[0] == "DOKS":
                            try:
                                if '{' in tmparr[8]:
                                    tmparr1 = tmparr[8].split('{')
                                    if tmparr1 > 1:
                                       tmparr[8] = tmparr1[0]
                                       if ',' in tmparr1[1]:
                                            tmparr2 = tmparr1[1].split(',')
                                            strana = tmparr2[0]
                                            if tmparr2 > 1:
                                                proizv = tmparr2[1].replace('}', '')                                                               
                                tmpstr = tmparr[8] + ";"+ tmparr[18]+";"+ tmparr[11]+";"+ tmparr[17]+";"+ dt + ";"+ seller+";"+ proizv +";"+ strana +"\n"      
                                outf.write(tmpstr.encode('1251'))
                            except Exception, e:
                                #continue
                                print e
                                
        
                        
    def getTime(self):
        now_time = datetime.datetime.now()
        cur_hour = now_time.hour # Час текущий
        cur_minute = now_time.minute # Минута текущая
        cur_second = now_time.second # Секунда текущие
        return "%s:%s:%s" % (cur_hour, cur_minute, cur_second)
        

if __name__ == "__main__":
    
    convertor = Convertor()
    print "Begin:", convertor.getTime()
    convertor.unZip()
    convertor.RunWorkConv()
    print "End:", convertor.getTime()
