#!/usr/bin/env python
# -*- coding: utf-8 -*-
from __future__ import with_statement

from datetime import datetime
import sys,time,os
import sqlite3 as sqlite3


import glob
import zipfile

import xlwt as pycel

# Excel has issues when creating too many styles/fonts, hence use
# a factory to reuse instances (see FAQ#13 http://poi.apache.org/faq.html )
STYLE_FACTORY = {}
FONT_FACTORY = {}

    # Background color
    # See http://groups.google.com.au/group/python-excel/attach/93621400bdddf464/palette_trial.xls?part=2
    # for colour indices
BLACK = 0
GREY = 22
#============================excel

def write(ws, row, col, data, style=None):
    """
    Write data to row, col of worksheet (ws) using the style
    information.

    Again, I'm wrapping this because you'll have to do it if you
    create large amounts of formatted entries in your spreadsheet
    (else Excel, but probably not OOo will crash).
    """
    if style:
        s = get_style(style)
        ws.write(row, col, data, s)
    else:
        ws.write(row, col, data)

def get_style(style):
    """
    Style is a dict maping key to values.
    Valid keys are: background, format, alignment, border

    The values for keys are lists of tuples containing (attribute,
    value) pairs to set on model instances...
    """
    #print "KEY", style
    style_key = tuple(style.items())
    s = STYLE_FACTORY.get(style_key, None)
    if s is None:
        s = pycel.XFStyle()
        for key, values in style.items():
            if key == "background":
                p = pycel.Pattern()
                for attr, value in values:
                    p.__setattr__(attr, value)
                s.pattern = p
            elif key == "format":
                s.num_format_str = values
            elif key == "alignment":
                a = pycel.Alignment()
                for attr, value in values:
                    a.__setattr__(attr, value)
                s.alignment = a
            elif key == "border":
                b = pycel.Formatting.Borders()
                for attr, value in values:
                    b.__setattr__(attr, value)
                s.borders = b
            elif key == "font":
                f = get_font(values)
                s.font = f
        STYLE_FACTORY[style_key] = s
    return s

def get_font(values):
    """
    'height' 10pt = 200, 8pt = 160
    """
    font_key = values
    f = FONT_FACTORY.get(font_key, None)
    if f is None:
        f = pycel.Font()
        for attr, value in values:
            f.__setattr__(attr, value)
        FONT_FACTORY[font_key] = f
    return f

def create_wb_ws():
    #============== Create a workbook
    wb = pycel.Workbook(encoding='cp1251')

    # Add a sheet
    ws = wb.add_sheet("Sheet")

    # Tweak printer settings
    # following makes a landscape layout on Letter paper
    # the width of the columns
    ws.fit_num_pages = 1
    ws.fit_height_to_pages = 0
    ws.fit_width_to_pages = 1
    # Set to Letter paper
    # See BiffRecords.SetupPageRecord for paper types/orientation
    ws.paper_size_code = 1
    # Set to landscape
    ws.portrait = 0
    StileZagol={"border": (("bottom",pycel.Formatting.Borders.THIN),
                      ("bottom_colour", BLACK),
                      ("right",pycel.Formatting.Borders.THIN),
                      ("right_colour", BLACK),("left",pycel.Formatting.Borders.THIN),
                      ("left_colour", BLACK)),"background": (("pattern", pycel.Pattern.SOLID_PATTERN),
                           ("pattern_fore_colour", GREY) )}
    # Freeze/split headers when scrolling
    write(ws, 0, 0, u"Наименование",StileZagol)
    write(ws, 0, 1, u"Кол-во",StileZagol)
    write(ws, 0, 2, u"Цена пост",StileZagol)
    write(ws, 0, 3, u"Цена розн",StileZagol)
    write(ws, 0, 4, u"Дата",StileZagol)
    write(ws, 0, 5, u"Поставщик",StileZagol)
    write(ws, 0, 6, u"Производитель",StileZagol)
    write(ws, 0, 7, u"Страна",StileZagol)
    ws.panes_frozen = True
    ws.horz_split_pos = 1
    return [wb,ws]
#=======================

#=====================================end func for excel

# function for filling table sohr
def zapolnit_vigr(FileName):
    inn = os.path.splitext(os.path.basename(FileName))[0].split('-')[1]

    connectDb=sqlite3.connect('ConvBD.db3', isolation_level=None, detect_types=sqlite3.PARSE_DECLTYPES|sqlite3.PARSE_COLNAMES)
    #con = sqlite3.connect(self.DB_PATH, isolation_level='DEFERRED', detect_types=sqlite3.PARSE_DECLTYPES|sqlite3.PARSE_COLNAMES)
    connectDb.row_factory = sqlite3.Row
    curs=connectDb.cursor()
    curs.execute('PRAGMA journal_mode = MEMORY;')
    curs.execute('PRAGMA synchronous = OFF;')


    fgHead = 1
    head = {}
    errorsIns=0

    with open(FileName, 'rt') as f:
        for row in f.xreadlines():
            if fgHead:
                row = row.strip().decode('1251').split('\t')
                if len(row)<2:
                    fgHead = 0
                    continue
                else:
                    if 'DOKN'==row[0]:
                        r = {}
                        rr = row[1:]
                        for nm in [u'FG_DOKTIP', u'ID_DOK', u'VK_POSTAVSCHIK', u'P_POSTAVSCHIK', u'P_DOKDAT', u'P_POLUCHATEL']:
                            i = rr.index(nm)
                            r[nm]=i
                        head[row[0]] = r
            else:

                row = row.decode('1251').split('\t')
                if len(row)<2:
                    continue
                tb = row[0]
                row = row[1:]

                #if 'DOKN'==tb and row[head['DOKN']['FG_DOKTIP']]<>'=':
                print "insert vigr run"
                if row[head['DOKN']['FG_DOKTIP']]=='-':
                    strInsDokn='insert into vigr (inn,nameorg,sklad) values(?,?,?)'
                    post= row[head['DOKN']['P_POSTAVSCHIK']]
                    poluchatel=row[head['DOKN']['P_POLUCHATEL']]

                    try:
                        curs.execute(strInsDokn,[inn,post,poluchatel]);
                    except sqlite3.DatabaseError,x:
                        errorsIns=errorsIns+1

                if 'DOKN'==tb and row[head['DOKN']['FG_DOKTIP']]=='+':
                    strInsDokn='insert into vigr (inn,nameorg,sklad) values(?,?,?)'
                    poluchatel=row[head['DOKN']['P_POLUCHATEL']]
                    try:
                        curs.execute(strInsDokn,[inn,poluchatel,'+']);
                    except sqlite3.DatabaseError,x:
                        #print "Oshibka:  ",x
                        errorsIns=errorsIns+1

    
        
    
#function for unziping files
def unzipfile(nmZip, nmDst=None, aDecode=None, aEncode=None):
    zf = None
    try:
        zf = zipfile.ZipFile(nmZip, "r")
        nm = zf.namelist()[0]
        if not nmDst:
            nmDst = nm
        zi = zf.getinfo(nm)
        # Получаем время src-файла из архива
        dt = list(zi.date_time); dt.append(0); dt.append(0); dt.append(1);
        with open(nmDst, 'wb') as f:
            if aDecode and aEncode:
                f.write(zf.read(nm).decode(aDecode).encode(aEncode))
            else:
                f.write(zf.read(nm))
        # Устанавливаем время dst-файла из архива
        dt = time.mktime(dt);
        os.utime(nmDst, (dt, dt))
    finally:
        if zf:
            zf.close()
#function for read dokn into temp table
def readDokN(FileName,listFilesPrihod,listFilesRashod,perBegin,perEnd,connectDb):

    inn = os.path.splitext(os.path.basename(FileName))[0].split('-')[1]

    curs=connectDb.cursor()
    cursVigr=connectDb.cursor()

    fgHead = 1
    head = {}
    fieldList = []
    with open(FileName, 'rt') as f:
        for row in f.xreadlines():
            if fgHead:
                row = row.strip().decode('1251').split('\t')
                if len(row)<2:
                    fgHead = 0
                    continue
                else:
                    if 'DOKN'==row[0]:
                        r = {}
                        rr = row[1:]
                        for nm in [u'FG_DOKTIP', u'ID_DOK', u'VK_POSTAVSCHIK', u'P_POSTAVSCHIK', u'P_DOKDAT', u'P_POLUCHATEL']:
                            i = rr.index(nm)
                            if u'FG_DOKTIP'<>nm:
                                fieldList.append(i)
                            r[nm]=i

                        #for bd
                        try:
                            curs.execute('''
                            drop table if exists "dokn"
                            ''')
                        except sqlite3.DatabaseError,x:
                            print "Oshibka:  ",x

                        strCreateDokn='CREATE  TABLE "DOKN" ('
                        for nm in row[1:]:
                            if nm=='ID_DOK':
                                strCreateDokn=strCreateDokn+nm+' INTEGER,'
                            elif nm=='VK_POSTAVSCHIK':
                                strCreateDokn=strCreateDokn+nm+' INTEGER,'
                            else:
                                strCreateDokn=strCreateDokn+nm+' text,'
                        strCreateDokn=strCreateDokn[:-1]+')'
                        #print strCreateDokn
                        try:
                            curs.execute(strCreateDokn);
                            curs.execute('CREATE INDEX "DOKN_ID_DOK" on "DOKN" (ID_DOK ASC)')
                            curs.execute('CREATE INDEX "DOKN_VK_POSTAVSCHIK" on "DOKN" (VK_POSTAVSCHIK ASC)')
                        except sqlite3.DatabaseError,x:
                            print "Oshibka:  ",x
                            #print strCreateDokn
                            sys.exit()
                        #curs.close()

                        #for bd

                        head[row[0]] = r

            else:
                row = row.decode('1251').split('\t')
                if len(row)<2:
                    continue
                tb = row[0]
                row = row[1:]

                #for bd
                if 'DOKN'==tb: #and row[head['DOKN']['FG_DOKTIP']]<>'=':


                    dt_objbeg=perBegin
                    dt_objend=perEnd
                    d=row[head['DOKN']['P_DOKDAT']].split(' ');
                    dt_objdokn = datetime.strptime(d[0], "%Y-%m-%d")
                    #~ if dt_objbeg<=dt_objdokn<dt_objend:

                    strInsDokn='insert into dokn values('
                    for valueCell in row:
                        strInsDokn=strInsDokn+"?,"
                    strInsDokn=strInsDokn[:-1]+');'
                    #print strInsDokn
                    try:
                        curs.execute(strInsDokn,row);
                    except sqlite3.DatabaseError,x:
                        print "Oshibka:  ",x
                        #curs.close()



                #for bd

                #создать словарь файлов расход
                if 'DOKN'==tb and row[head['DOKN']['FG_DOKTIP']]=='-':
                    #Поменять на переменные период
                    dt_objbeg=perBegin
                    dt_objend=perEnd

                    d=row[head['DOKN']['P_DOKDAT']].split(' ');
                    dt_objdokn = datetime.strptime(d[0], "%Y-%m-%d")

                    try:
                        strVigr='''select vigr from vigr where inn=?
                            and nameOrg=? and sklad=?'''
                        cursVigr.execute(strVigr,[inn,row[head['DOKN']['P_POSTAVSCHIK']],row[head['DOKN']['P_POLUCHATEL']],])
                    except sqlite3.DatabaseError,x:
                        print "Oshibka:  ",x


                    for vigr in cursVigr.fetchall():
                        if vigr['vigr']==1:
                            if dt_objbeg<=dt_objdokn<dt_objend:

                                filename = u'%s_%s' % (inn, row[head['DOKN']['P_POLUCHATEL']].replace(' ','_'))
                                filename=filename.replace('/','_')
                                filename=filename.replace('\\','_')
                                filename=filename.replace('"',"'")
                                filename=filename.replace('?',"_")

                                listFilesRashod[filename] = None
                #создать словарь файлов приход
                if 'DOKN'==tb and row[head['DOKN']['FG_DOKTIP']]=='+':
                    #Поменять на переменные период
                    dt_objbeg=perBegin
                    dt_objend=perEnd

                    d=row[head['DOKN']['P_DOKDAT']].split(' ');
                    dt_objdokn = datetime.strptime(d[0], "%Y-%m-%d")

                    try:
                        strVigr='''select vigr from vigr where inn=?
                            and nameOrg=? and sklad=?'''
                        cursVigr.execute(strVigr,[inn,row[head['DOKN']['P_POLUCHATEL']],'+',])
                    except sqlite3.DatabaseError,x:
                        print "Oshibka:  ",x

                    for vigr in cursVigr.fetchall():
                        if vigr['vigr']==1:
                            if dt_objbeg<=dt_objdokn<dt_objend:

                                filename = u'%s_%s' % (inn, row[head['DOKN']['P_POLUCHATEL']].replace(' ','_'))
                                filename=filename.replace('/','_')
                                filename=filename.replace('\\','_')
                                filename=filename.replace('"',"'")
                                filename=filename.replace('?',"_")


                                listFilesPrihod[filename] = None


    return [listFilesPrihod,listFilesRashod]
#function for read doks into temp table
def readDokS(FileName,connectDb):
    curs=connectDb.cursor()
    fgHead = 1
    head = {}
    fieldList = []
    with open (FileName,'rt') as f:
        for row in f.xreadlines():
            if fgHead:
                row=row.strip().decode('1251').split('\t')
                if len(row)<2:
                    fgHead=0
                    continue
                else:
                    if 'DOKS'==row[0]:

                        #for bd
                        try:
                            curs.execute('''
                            drop table if exists "doks"
                            ''')
                        except sqlite3.DatabaseError,x:
                            print "Oshibka:  ",x
                        connectDb.commit

                        strCreateDoks='CREATE  TABLE "doks" ('
                        for nm in row[1:]:
                            if nm=='ID_PARTIYA':
                                strCreateDoks=strCreateDoks+nm+' INTEGER,'
                            elif nm=='ID_PRIHOD':
                                strCreateDoks=strCreateDoks+nm+' INTEGER,'
                            elif nm=='VK_DOK':
                                strCreateDoks=strCreateDoks+nm+' INTEGER,'
                            elif nm=='VK_POSTAVSCHIK':
                                strCreateDoks=strCreateDoks+nm+' INTEGER,'
                            else:
                                strCreateDoks=strCreateDoks+nm+' text,'
                        strCreateDoks=strCreateDoks[:-1]+')'
                        #print strCreateDokn
                        try:
                            curs.execute(strCreateDoks);
                            curs.execute('CREATE INDEX "DOKS_ID_PARTIYA" on "DOKS" (ID_PARTIYA ASC)')
                            curs.execute('CREATE INDEX "DOKS_ID_PRIHOD" on "DOKS" (ID_PRIHOD ASC)')
                            curs.execute('CREATE INDEX "DOKS_VK_DOK" on "DOKS" (VK_DOK ASC)')
                            curs.execute('CREATE INDEX "DOKS_VK_POSTAVSCHIK" on "DOKS" (VK_POSTAVSCHIK ASC)')
                        except sqlite3.DatabaseError,x:
                            print "Oshibka:  ",x
                        #curs.close()

                        #for bd
            else:
                row=row.decode('1251').split('\t')
                if len(row)<2:
                    continue
                tb=row[0]
                row=row[1:]
                if 'DOKS'==tb:

                #for bd

                    strInsDoks='insert into doks values('
                    for valueCell in row:
                        strInsDoks=strInsDoks+"?,"
                    strInsDoks=strInsDoks[:-1]+')'
                    #print strInsDokn
                    try:
                        curs.execute(strInsDoks,row);
                        connectDb.commit
                    except sqlite3.DatabaseError,x:
                        #print strInsDoks
                        print "Oshibka:  ",x
                    #curs.close()



                #for bd
# function for filling table sohr
def makeBdSohr(connectDb):
    tmpPar=[]


    cursExec=connectDb.cursor()

    cursDokn=connectDb.cursor()
    cursDokn.execute('select * from dokn')



    cursDoks=connectDb.cursor()

    cursSelectDokPost=connectDb.cursor()

    cursSohrVigr=connectDb.cursor()




    for DOKN in cursDokn.fetchall():
        cursSohrVigr.execute("select * from vigr where sohr=1 and sklad='+'")
        for sohr in cursSohrVigr.fetchall():
            if DOKN['FG_DOKTIP']=='+' or DOKN['FG_DOKTIP']=='-' or DOKN['P_POSTAVSCHIK']==sohr['nameOrg'] or DOKN['P_POLUCHATEL']==sohr['nameOrg'] :
                tmpPar[:]=[]
                tmpPar.append(DOKN['ID_DOK'])
                tmpPar.append(DOKN['FG_DOKTIP'])
                tmpPar.append(DOKN['P_DOKDAT'])
                tmpPar.append(DOKN['P_POLUCHATEL'])
                tmpPar.append(DOKN['P_POSTAVSCHIK'])
                try:
                    cursExec.execute("insert into DOKN_SOHR (ID_DOK,FG_DOKTIP,P_DOKDAT,P_POLUCHATEL,P_POSTAVSCHIK) values(?,?,?,?,?)",tmpPar)
                except sqlite3.DatabaseError,x:
                    #print "Oshibka:  ",x
                    None

                strDoks='select * from doks where VK_DOK=?'
                tmpPar[:]=[]
                tmpPar.append(DOKN['ID_DOK'])
                cursDoks.execute(strDoks,tmpPar)
                for DOKS in cursDoks.fetchall():
                    tmpPar[:]=[]
                    tmpPar.append(DOKS['ID_PARTIYA'])
                    tmpPar.append(DOKS['VK_DOK'])
                    try:tmpPar.append(DOKS['ID_PARTIYAFIRST'])
                    except: tmpPar.append('')

                    tmpPar.append(DOKS['ID_PRIHOD'])
                    try:
                        cursExec.execute("insert into DOKS_SOHR (ID_PARTIYA,VK_DOK,ID_PARTIYAFIRST,id_prihod) values(?,?,?,?)",tmpPar)
                    except sqlite3.DatabaseError,x:
                        None
                        #print "Oshibka:  ",x



def makeFiles1(FileName,listFilesPrihod,listFilesRashod,mes,connectDb,outdirPrihod,outdirRashod):
    tmpPar=[]
    #===for debugging (empty poluchatel============================
    #~ fpusto=open('pusto.txt', 'a')
    #===============================

    f=file
    inn=os.path.splitext(os.path.basename(FileName))[0].split('-')[1]

    cursDokn=connectDb.cursor()
    cursDokn.execute('select * from dokn')


    cursDoks=connectDb.cursor()

    cursSelectDokPost=connectDb.cursor()
    cursSelectDokPost_sohr=connectDb.cursor()
    for fName in listFilesRashod:
        line=1
        wb,ws=create_wb_ws()
        cursDokn.execute("select * from dokn where FG_DOKTIP='-'")
        for DOKN in cursDokn.fetchall():
            if fName==u'%s_%s'%(inn,DOKN['P_POLUCHATEL'].replace(' ','_').replace('\\','_').replace('/','_').replace('"',"'").replace('?',"_")):

                strDoks='select * from doks where VK_DOK=?'
                tmpPar[:]=[]
                tmpPar.append(DOKN['ID_DOK'])
                cursDoks.execute(strDoks,tmpPar)
                for DOKS in cursDoks.fetchall():

                    naim=DOKS['P_TOVAR']
                    if len(DOKS['P_TOVAR'].split('{'))>1 and len(DOKS['P_TOVAR'].split('}'))>1:
                        #строка наимеование  без производителя
                        naim=DOKS['P_TOVAR'].split('{')[0]
                    write(ws, line, 0, naim,{"border": (("bottom",pycel.Formatting.Borders.THIN),
                          ("bottom_colour", BLACK),
                          ("right",pycel.Formatting.Borders.THIN),
                          ("right_colour", BLACK),("left",pycel.Formatting.Borders.THIN),
                          ("left_colour", BLACK)),"alignment":(("wrap", pycel.Alignment.WRAP_AT_RIGHT),("horz", pycel.Alignment.HORZ_FILLED),)})

                    write(ws, line, 1, DOKS['P_KOLVO'],{"border": (("bottom",pycel.Formatting.Borders.THIN),
                          ("bottom_colour", BLACK),
                          ("right",pycel.Formatting.Borders.THIN),
                          ("right_colour", BLACK),("left",pycel.Formatting.Borders.THIN),
                          ("left_colour", BLACK))})


                    write(ws, line, 2, DOKS['P_CENAPOS'],{"border": (("bottom",pycel.Formatting.Borders.THIN),
                          ("bottom_colour", BLACK),
                          ("right",pycel.Formatting.Borders.THIN),
                          ("right_colour", BLACK),("left",pycel.Formatting.Borders.THIN),
                          ("left_colour", BLACK))})


                    write(ws, line, 3, DOKS['P_CENAROZ'],{"border": (("bottom",pycel.Formatting.Borders.THIN),
                          ("bottom_colour", BLACK),
                          ("right",pycel.Formatting.Borders.THIN),
                          ("right_colour", BLACK),("left",pycel.Formatting.Borders.THIN),
                          ("left_colour", BLACK))})

                    tmpdate=(DOKN['P_DOKDAT'].split(' ')[0]).split('-')
                    newdate=u'.'.join([tmpdate[2],tmpdate[1],tmpdate[0],])

                    write(ws, line, 4, newdate,{"border": (("bottom",pycel.Formatting.Borders.THIN),
                          ("bottom_colour", BLACK),
                          ("right",pycel.Formatting.Borders.THIN),
                          ("right_colour", BLACK),("left",pycel.Formatting.Borders.THIN),
                          ("left_colour", BLACK))})

                    strSelPost="""
select dokn.P_POSTAVSCHIK as P_POSTAVSCHIK from dokn join doks on dokn.ID_DOK=DOKS.VK_DOK
where DOKS.ID_partiya=? limit 1"""

                    tmpPar[:]=[]
                    tmpPar.append(DOKS['id_Prihod'])

                    try:
                        cursSelectDokPost.execute(strSelPost,tmpPar)
                    except sqlite3.DatabaseError,x:
                        print "Oshibka:  ",x

                    post=''
                    for DoksPost in cursSelectDokPost.fetchall():
                        post=DoksPost['P_POSTAVSCHIK']
#=========================select post from sohr=================================

                    if post=='':
                        strSelPost="""
                            select dokn_sohr.P_POSTAVSCHIK as P_POSTAVSCHIK from dokn_sohr join doks_sohr on dokn_sohr.ID_DOK=doks_sohr.VK_DOK
                            where doks_sohr.ID_partiyafirst=?"""
                            #and length(dokn.p_postavschik)>2  limit 1"""
                        tmpPar[:]=[]
                        tmpPar.append(DOKS['id_Prihod'])
                        #~ print DOKS['id_Prihod']
                        #tmpPar.append(DOKS['ID_PARTIYA'])
                        #~ print DOKS['id_Prihod']
                        try:
                            cursSelectDokPost_sohr.execute(strSelPost,tmpPar)
                        except sqlite3.DatabaseError,x:
                            print "Oshibka:  ",x
                        for DoksPost1 in cursSelectDokPost_sohr.fetchall():
                            post=DoksPost1['P_POSTAVSCHIK']
                            #print post
#==================for debug empty post into file=================
                        #~ if post='':
                            #~ dokn=[]
                            #~ doks=[]
                            #~ for d in DOKN:
                                #~ dokn.append(str(d))
                            #~ for d in DOKS:
                                #~ doks.append(str(d))
#~
                            #~ fpusto.write('\t'.join(dokn))
                            #~ fpusto.write('\t'.join(doks))
                            #~ fpusto.write('\n')

                        #post=str(DOKS['id_Prihod'])+'____'+str(DOKS['id_partiyafirst'])
#===========================================================

                    #~ if post='':
                        #~ post=DOKS['id_Prihod']


                    write(ws, line, 5,post,{"border": (("bottom",pycel.Formatting.Borders.THIN),
                          ("bottom_colour", BLACK),
                          ("right",pycel.Formatting.Borders.THIN),
                          ("right_colour", BLACK),("left",pycel.Formatting.Borders.THIN),
                          ("left_colour", BLACK))})

                    proizv=''
                    strana=''
                    a=DOKS['P_TOVAR'].split('{')
                    if len(a)>1:
                        aa=a[1].split('}')
                        aaa=aa[0].split(',')
                        if len(aaa)>1:
                            proizv=aaa[1]
                            strana=aaa[0]

                    write(ws, line, 6, proizv,{"border": (("bottom",pycel.Formatting.Borders.THIN),
                          ("bottom_colour", BLACK),
                          ("right",pycel.Formatting.Borders.THIN),
                          ("right_colour", BLACK),("left",pycel.Formatting.Borders.THIN),
                          ("left_colour", BLACK))})

                    write(ws, line, 7, strana,{"border": (("bottom",pycel.Formatting.Borders.THIN),
                          ("bottom_colour", BLACK),
                          ("right",pycel.Formatting.Borders.THIN),
                          ("right_colour", BLACK),("left",pycel.Formatting.Borders.THIN),
                          ("left_colour", BLACK))})


                    line=line+1
        ws.col(0).width=256*50
        ws.col(1).width=256*20
        ws.col(2).width=256*20
        ws.col(3).width=256*20
        ws.col(4).width=256*20
        ws.col(5).width=256*50
        ws.col(6).width=256*35
        ws.col(7).width=256*35


        wb.save(u'%s/Отчет по %s за %s мес.xls' % (outdirRashod,fName,mes))

    for fName in listFilesPrihod:

        line=1
        wb,ws=create_wb_ws()
        cursDokn.execute("select * from dokn where FG_DOKTIP='+'")
        for DOKN in cursDokn.fetchall():
            if fName==u'%s_%s'%(inn,DOKN['P_POLUCHATEL'].replace(' ','_').replace('\\','_').replace('/','_').replace('"',"'").replace('?',"_")):
                strDoks='select * from doks where VK_DOK=?'
                tmpPar[:]=[]
                tmpPar.append(DOKN['ID_DOK'])
                cursDoks.execute(strDoks,tmpPar)
                for DOKS in cursDoks.fetchall():



                    naim=DOKS['P_TOVAR']

                    if len(DOKS['P_TOVAR'].split('{'))>1 and len(DOKS['P_TOVAR'].split('}'))>1:
                        # строка наименование без производителя
                        naim=DOKS['P_TOVAR'].split('{')[0]

                    write(ws, line, 0, naim,{"border": (("bottom",pycel.Formatting.Borders.THIN),
                          ("bottom_colour", BLACK),
                          ("right",pycel.Formatting.Borders.THIN),
                          ("right_colour", BLACK),("left",pycel.Formatting.Borders.THIN),
                          ("left_colour", BLACK))})


                    write(ws, line, 1, DOKS['P_KOLVO'],{"border": (("bottom",pycel.Formatting.Borders.THIN),
                          ("bottom_colour", BLACK),
                          ("right",pycel.Formatting.Borders.THIN),
                          ("right_colour", BLACK),("left",pycel.Formatting.Borders.THIN),
                          ("left_colour", BLACK))})


                    write(ws, line, 2, DOKS['P_CENAPOS'],{"border": (("bottom",pycel.Formatting.Borders.THIN),
                          ("bottom_colour", BLACK),
                          ("right",pycel.Formatting.Borders.THIN),
                          ("right_colour", BLACK),("left",pycel.Formatting.Borders.THIN),
                          ("left_colour", BLACK))})

                    write(ws, line, 3, DOKS['P_CENAROZ'],{"border": (("bottom",pycel.Formatting.Borders.THIN),
                          ("bottom_colour", BLACK),
                          ("right",pycel.Formatting.Borders.THIN),
                          ("right_colour", BLACK),("left",pycel.Formatting.Borders.THIN),
                          ("left_colour", BLACK))})

                    tmpdate=(DOKN['P_DOKDAT'].split(' ')[0]).split('-')
                    newdate=u'.'.join([tmpdate[2],tmpdate[1],tmpdate[0],])

                    write(ws, line, 4, newdate,{"border": (("bottom",pycel.Formatting.Borders.THIN),
                          ("bottom_colour", BLACK),
                          ("right",pycel.Formatting.Borders.THIN),
                          ("right_colour", BLACK),("left",pycel.Formatting.Borders.THIN),
                          ("left_colour", BLACK))})


                    write(ws, line, 5, DOKN['P_POSTAVSCHIK'],{"border": (("bottom",pycel.Formatting.Borders.THIN),
                          ("bottom_colour", BLACK),
                          ("right",pycel.Formatting.Borders.THIN),
                          ("right_colour", BLACK),("left",pycel.Formatting.Borders.THIN),
                          ("left_colour", BLACK))})

                    proizv=''
                    strana=''
                    a=DOKS['P_TOVAR'].split('{')
                    if len(a)>1:
                        aa=a[1].split('}')
                        aaa=aa[0].split(',')
                        if len(aaa)>1:
                            proizv=aaa[1]
                            strana=aaa[0]

                    write(ws, line, 6, proizv,{"border": (("bottom",pycel.Formatting.Borders.THIN),
                          ("bottom_colour", BLACK),
                          ("right",pycel.Formatting.Borders.THIN),
                          ("right_colour", BLACK),("left",pycel.Formatting.Borders.THIN),
                          ("left_colour", BLACK))})

                    write(ws, line, 7, strana,{"border": (("bottom",pycel.Formatting.Borders.THIN),
                          ("bottom_colour", BLACK),
                          ("right",pycel.Formatting.Borders.THIN),
                          ("right_colour", BLACK),("left",pycel.Formatting.Borders.THIN),
                          ("left_colour", BLACK))})


                    line=line+1
        ws.col(0).width=256*50
        ws.col(1).width=256*20
        ws.col(2).width=256*20
        ws.col(3).width=256*20
        ws.col(4).width=256*20
        ws.col(5).width=256*50
        ws.col(6).width=256*35
        ws.col(7).width=256*35
        wb.save(u'%s/Отчет по %s за %s мес.xls' % (outdirPrihod,fName,mes))




#function for making all operations from WorkConv.py
def convertation_make(FileName,perBegin,perEnd,mes,outdir):
#================for debug empty postavschik==========================
    #~ with open('pusto.txt', 'w') as f:
        #~ f.write('')
        #~ f.close()
#==========================================

    outdirPrihod=u'/'.join([outdir,'prihod'])
    outdirRashod=u'/'.join([outdir,'rashod'])
    if not os.path.exists(outdir):
        os.mkdir(outdir)
    if not os.path.exists(outdirPrihod):
        os.mkdir(outdirPrihod)
    if not os.path.exists(outdirRashod):
        os.mkdir(outdirRashod)



    listFilesPrihod={}
    listFilesRashod={}


    connectDb=sqlite3.connect('ConvBD.db3', isolation_level=None, detect_types=sqlite3.PARSE_DECLTYPES|sqlite3.PARSE_COLNAMES)
    #con = sqlite3.connect(self.DB_PATH, isolation_level='DEFERRED', detect_types=sqlite3.PARSE_DECLTYPES|sqlite3.PARSE_COLNAMES)
    connectDb.row_factory = sqlite3.Row
    curs=connectDb.cursor()
    curs.execute('PRAGMA journal_mode = MEMORY;')
    curs.execute('PRAGMA synchronous = OFF;')


    listFilesPrihod,listFilesRashod=readDokN(FileName,listFilesPrihod,listFilesRashod,perBegin,perEnd,connectDb)


    readDokS(FileName,connectDb)

    #~ fshapka = open("shapka.xml", "r")
    #~ fkonec = open("konec.xml", "r")

    #~ for fName, v in listFilesPrihod.iteritems():
#~
        #~ filename = '%s/Отчет по %s за %s мес.xls' % (outdirPrihod,fName,mes)
        #~ with open(filename, 'w') as f:
            #~ f.write('')
#~
    #~ for fName, v in listFilesRashod.iteritems():
#~
        #~ filename = '%s/Отчет по %s за %s мес.xls' % (outdirRashod,fName,mes)
        #~ with open(filename, 'w') as f:
            #~ f.write('')


    #~ for line in fshapka.readlines():
        #~ for fName, v in listFilesPrihod.iteritems():
            #~ filename = '%s/Отчет по %s за %s мес.xls' % (outdirPrihod,fName,mes)
            #~ with open(filename, 'a') as f:
                #~ f.write(line)
        #~ for fName, v in listFilesRashod.iteritems():
            #~ filename = '%s/Отчет по %s за %s мес.xls' % (outdirRashod,fName,mes)
            #~ with open(filename, 'a') as f:
                #~ f.write(line)



    makeBdSohr(connectDb)
    print "=========sohranenie dannih zakoncheno============="
    makeFiles1(FileName,listFilesPrihod,listFilesRashod,mes,connectDb,outdirPrihod,outdirRashod)

    curs.execute('select count(*) from DOKN')
    for i in curs.fetchall():
        print "Dokn: ",i

    curs.execute('select count(*) from DOKS')
    for i in curs.fetchall():
        print "Doks: ",i

    #~ for line in fkonec.readlines():
        #~ for fName, v in listFilesPrihod.iteritems():
            #~ filename = '%s/Отчет по %s за %s мес.xls' % (outdirPrihod,fName,mes)
            #~ with open(filename, 'a') as f:
                #~ f.write(line)
        #~ for fName, v in listFilesRashod.iteritems():
            #~ filename = '%s/Отчет по %s за %s мес.xls' % (outdirRashod,fName,mes)
            #~ with open(filename, 'a') as f:
                #~ f.write(line)

