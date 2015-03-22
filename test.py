#!/usr/bin/env python
# -*- coding: cp1251 -*-

from __future__ import with_statement

from datetime import datetime
import sys,time,os,glob
if os.name == 'nt':
    if os.path.isdir('gtk'):
        # Used to create windows installer with GTK included
        paths = os.environ['PATH']
        print paths
        list_ = paths.split(';')
        new_list = []
        for p in list_:
            if p.find('gtk') < 0 and p.find('GTK') < 0:
                new_list.append(p)
        new_list.insert(0, 'gtk/lib')
        new_list.insert(0, 'gtk/bin')
        os.environ['PATH'] = ';'.join(new_list)
        os.environ['GTK_BASEPATH'] = 'gtk'

import gtk
import workConv
import dbworker



import decimal
import threading
import time
import copy


import gobject
from os.path import dirname as path_dirname, join as path_join, abspath as path_abspath
from os import chdir
_file_ = path_abspath(__file__)

# poluchit spisok dokn dokumentov s postavshikami i znkom(vidDoc) iz faila

sys.ms71_db = dbworker.DBWorker('ConvBD.db3')



def SelectFile():
    dialog = gtk.FileChooserDialog("Open..",
                              None,
                              gtk.FILE_CHOOSER_ACTION_OPEN,
                              (gtk.STOCK_CANCEL, gtk.RESPONSE_CANCEL,
                               gtk.STOCK_OPEN, gtk.RESPONSE_OK))
    dialog.set_default_response(gtk.RESPONSE_OK)
    filter = gtk.FileFilter()
    filter.set_name("txt or zip")
    filter.add_pattern("*.txt")
    filter.add_pattern("*.zip")
    dialog.add_filter(filter)
    dialog.set_select_multiple(1)
    listFiles=[]

    response = dialog.run()
    if response == gtk.RESPONSE_OK:
        filenames = dialog.get_filenames()
        #dialog.hide()
        dialog.destroy()

        while gtk.events_pending():
            gtk.main_iteration()

        for fName in filenames:
            print listFiles.append(fName)

    elif response == gtk.RESPONSE_CANCEL:
        print 'Closed, no files selected'
        dialog.destroy()
    return listFiles


def SelectDir():
    listFiles=[]
    dialog = gtk.FileChooserDialog("Open..",
                              None,
                              gtk.FILE_CHOOSER_ACTION_SELECT_FOLDER,
                              (gtk.STOCK_CANCEL, gtk.RESPONSE_CANCEL,
                               gtk.STOCK_OPEN, gtk.RESPONSE_OK))
    dialog.set_default_response(gtk.RESPONSE_OK)



    response = dialog.run()
    if response == gtk.RESPONSE_OK:
        mask=(os.path.join(dialog.get_current_folder(),'*.txt')).decode('utf-8')
        maskZip=(os.path.join(dialog.get_current_folder(),'*.zip')).decode('utf-8')

        #dialog.hide()
        dialog.destroy()

        while gtk.events_pending():
            gtk.main_iteration()

        for fName in glob.iglob(mask):
            listFiles.append(fName)
        for fName in glob.iglob(maskZip):
            listFiles.append(fName)

    elif response == gtk.RESPONSE_CANCEL:
        print ''
        dialog.destroy()

    return listFiles

class adderGui:

    def __init__( self ):

        self.builder = gtk.Builder()
        self.builder.add_from_file("main.ui")
        self.window = self.head
        self.window.connect("destroy", gtk.main_quit)
        self.ListFileNames=[]

        dic = {
            "on_button3_clicked" : self.openwindowFile,
            "on_button1_clicked" : self.openwindowDir,
            "on_button2_clicked" : self.RunWorkConv,
            "on_button4_clicked" : self.RunRefreshVigr,
        }

        self.builder.connect_signals( dic )
        #self.builder.get_object("entry2").set_text(time.strftime('%Y-%m-%d', time.localtime()))
        row = time.strftime('%m', time.localtime())
        if int(row)==1:
            self.builder.get_object("entry1").set_text('12')


        else:
            self.builder.get_object("entry1").set_text(str(int(row)-1))
        self.SelectVigr()



    def run(self,widget):
        print "run"

        self.builder.get_object("entry2").set_text(time.strftime('%Y-%m-%d', time.localtime()))
        row = time.strftime('%Y-%m-%d', time.localtime()).split('-')
        self.builder.get_object("entry1").set_text('-'.join([row[0],'0'+str(int(row[1])-1),row[2]]))


    def openwindowFile(self,widget):

        self.ListFileNames=SelectFile()
        if len(self.ListFileNames)==0:
            self.builder.get_object("label2").set_text('Выберите файл/папку')
        if len(self.ListFileNames)==1:
            self.builder.get_object("label2").set_text('Файл:'+self.ListFileNames[0])
        if len(self.ListFileNames)>1:
            self.builder.get_object("label2").set_text('Выбрано несколько файлов')



    def openwindowDir(self,widget):

        self.ListFileNames=SelectDir()
        if len(self.ListFileNames)==0:
            self.builder.get_object("label2").set_text('Выберите файл/папку')
        if len(self.ListFileNames)==1:
            self.builder.get_object("label2").set_text('Файл:'+self.ListFileNames[0])
        if len(self.ListFileNames)>1:
            self.builder.get_object("label2").set_text('Выбрано несколько файлов')

    def RunWorkConv(self,widget):
        tmpListFileNames=[]
        if len(self.ListFileNames)>0:

            rowBeg='-'.join([time.strftime('%Y', time.localtime()),adderGui.builder.get_object("entry1").get_text(),'01'])
            rowEnd='-'.join([time.strftime('%Y', time.localtime()),str(int(adderGui.builder.get_object("entry1").get_text())+1),'01'])
            dt_objbeg=datetime.strptime( rowBeg, "%Y-%m-%d")
            dt_objend=datetime.strptime( rowEnd, "%Y-%m-%d")
            mes=adderGui.builder.get_object("entry1").get_text()

            for fName in self.ListFileNames:
                print fName
                sys.stdout.flush()

                if '.ZIP'==os.path.splitext(fName)[1].upper():
                    dstName = os.path.splitext(fName)[0] + '.txt'
                    workConv.unzipfile(fName, dstName)
                    oldName = fName + '.old'
                    try: os.remove(oldName)
                    except: pass
                    try: os.rename(fName, oldName)
                    except: pass
                else:
                    dstName = fName

                tmpListFileNames.append(dstName)

                with open(dstName, 'rt') as f:
                    for row in f.xreadlines():
                        row = row.strip().split('\t')
                        if 'DOKN'==row[0]:
                            f.close()
                            outdir='./out1'
                            workConv.convertation_make(dstName,dt_objbeg,dt_objend,mes,outdir)
                            break
            print 'END.'
            self.ListFileNames=tmpListFileNames
            sys.stdout.flush()



    def closewindow2(self,widget):
        print "run"
        adderGui.window2.hide()

    def quit(self, widget):
        sys.exit(0)

    def RunRefreshVigr(self,widget):
        tmpListFileNames=[]
        for fName in self.ListFileNames:
            print fName
            if '.ZIP'==os.path.splitext(fName)[1].upper():
                dstName = os.path.splitext(fName)[0] + '.txt'
                workConv.unzipfile(fName, dstName)
                oldName = fName + '.old'
                try: os.remove(oldName)
                except: pass
                try: os.rename(fName, oldName)
                except: pass
            else:
                dstName = fName
            tmpListFileNames.append(dstName)
            with open(dstName, 'rt') as f:
                for row in f.xreadlines():
                    row = row.strip().split('\t')
                    if 'DOKN'==row[0]:
                        f.close()
                        workConv.zapolnit_vigr(dstName)
                        break
        self.ListFileNames=tmpListFileNames
        print 'Zakoncheno obnonovlenie tablici vigruzka'
        self.SelectVigr()




    def SelectVigr(self):

        # Чтение результата будет с первой записи
        sys.ms71_db.skipstr = 0
        sys.ms71_db.kolvostr = [0,0]
        self.type_field = []  # список типов полей
        self.name_field = []  # список наименование полей
        self.cn = []          # список описаний колонок
        #self.ls =[]           # список данных их типов строк
        #Удаление старых колонок
        try:
            pos =self.tvgrid.get_columns()
            for i in pos:
                self.tvgrid.remove_column(i)
        except:
            pass
        #Считываем список полей
        g= sys.ms71_db.field_dokn()
        #self.set_statusbar()

        #Преобразование типов десятичных и юникоды создается список типов полей
        for i in g:
            self.name_field.append(i[0])
            if i[0]=='vigr':
                self.type_field.append(bool)
            elif i[0]=='sohr':
                self.type_field.append(bool)
            else:
                self.type_field.append(str)


        #Создаем gtk.ListStore со списком полей - колонок пока данных здесь нет
        #self.ls = gtk.ListStore(*(self.type_field+[str,str]))#<type str>

        # Прочитываем первую порцию записей
        ls = self.selectstr()

        for i in xrange(len(self.type_field)):
            if(i==3):
                self.cell = gtk.CellRendererToggle()
            elif(i==4):
                self.cell = gtk.CellRendererToggle()
            else:
                self.cell = gtk.CellRendererText()
            #~ self.cell = gtk.CellRendererText()
            #Для первой колонки задаем цвет фона и шрифта, которые будут указыватся в доп. полях
            if i == 0:
                self.cn = gtk.TreeViewColumn(self.name_field[i], self.cell, markup = i,
                    background= len(self.type_field),foreground = len(self.type_field)+1)

                #self.cn.set_sizing(gtk.TREE_VIEW_COLUMN_AUTOSIZE) #TREE_VIEW_COLUMN_AUTOSIZE   TREE_VIEW_COLUMN_FIXED
                #self.cn.set_sizing(gtk.TREE_VIEW_COLUMN_GROW_ONLY)
                self.cn.set_resizable(True)
                #self.cn.set_fixed_width(64)

                #Задаем у колонки и заголовок
                self.cn.set_widget(gtk.Label())
                self.cn.get_widget().set_markup(u'<b>inn</b>') #self.tvgrid.set_cursor(0)
                self.cn.get_widget().show()
            elif i == 3:
                #print 'self.name_field[i]', self.name_field[i]
                self.cn = gtk.TreeViewColumn(self.name_field[i], self.cell, active = i)
                self.cell.connect('toggled', self.toggled_cb, (ls, i))
            elif i == 4:
                #print 'self.name_field[i]', self.name_field[i]
                self.cn = gtk.TreeViewColumn(self.name_field[i], self.cell, active = i)
                self.cell.connect('toggled', self.toggled_cb1, (ls, i))
            else:
                self.cn = gtk.TreeViewColumn(self.name_field[i], self.cell, markup = i)




            if i==1:
                self.cell.set_property('background','DarkOliveGreen')
                self.cell.set_property('foreground','White')
                self.cn.set_resizable(True)
            if i==2:
                self.cell.set_property('background','Moccasin')
                self.cell.set_property('foreground','Navy')
                self.cn.set_resizable(True)

            if self.type_field == str:
                self.cn.set_alignment(0.0)
            else:
                self.cn.set_alignment(0.5)

            self.tvgrid.append_column(self.cn)

        #Устанавливаем свойства gtk.TreeeView
        self.tvgrid.set_headers_visible(True)
        self.tvgrid.set_enable_search(False)
        self.tvgrid.set_enable_tree_lines(True)



        #Связываем gtk.TreeeView с gtk.ListStore
        self.tvgrid.set_model(ls)
        sys.ms71_db.tekzap = 1
        #print dir(self.vscrollbardata)
        #print self.vscrollbardata.get_min_slider_size()

    #Читаем порцию данных
    def selectstr(self, *args):
        #Считываем первую строку из порции
        row = sys.ms71_db.sel_dokn()
        #self.ls.clear()
        ls = gtk.ListStore(*(self.type_field+[str,str]))

        while row :
            row = list(row)
            #Преобразование типов десятичных и дат, создается в каждом поле строки
            for i, v in enumerate(row):
                if i==3:
                    row[i]=bool(v)
                elif i==4:
                    row[i]=bool(v)
                    #~ print 'row[i]=bool(v)', row[i], bool(v)
                elif isinstance(v, decimal.Decimal):
                    row[i] = float(v)
                elif isinstance(v, unicode):
                    row[i] = gobject.markup_escape_text(v)
                    pass
                #~ elif isinstance(v, datetime.date):
                    #~ row[i] = str(v)
                #~ elif v is None:
                    #~ if self.type_field[i] in [int, float]:
                        #~ row[i] = 0
                else:
                    pass
            if row[0] < 300:
                ls.append(row+['SlateBlue','Snow'])
            else:
                # gtk.ListStore - содержит данные и типы данных строк
                if row[3]==True or row[3]==False:
                    ls.append(row+[None,None])
                elif row[4]==True or row[4]==False:
                    ls.append(row+[None,None])
                else:
                    ls.append(row+[None,None])

            #Считываем очередную строку порции
            row = sys.ms71_db.sel_dokn()
        return ls
#======================================================================================
    def toggled_cb(self, render, path, user_data):
        model, column = user_data

        #path строка
        #coulum столбец
        model[path][column] = not model[path][column]

        inn=model[path][0]
        nameOrg=model[path][1].replace('&apos;',"'")
        nameOrg=nameOrg.replace('&quot;','"')

        sklad=model[path][2].replace('&apos;',"'")
        sklad=sklad.replace('&quot;','"')

        strUpd='update vigr set vigr=? where inn=? and nameorg=? and sklad=?'
        QueryParams=[int(model[path][3]),inn.decode('utf-8'),nameOrg.decode('utf-8'),sklad.decode('utf-8'),]

        sys.ms71_db.updateVigr(strUpd,QueryParams)

        return
    def toggled_cb1(self, render, path, user_data):
        model, column = user_data

        #path строка
        #coulum столбец
        model[path][column] = not model[path][column]

        inn=model[path][0]
        nameOrg=model[path][1].replace('&apos;',"'")
        nameOrg=nameOrg.replace('&quot;','"')

        sklad=model[path][2].replace('&apos;',"'")
        sklad=sklad.replace('&quot;','"')

        strUpd='update vigr set sohr=? where inn=? and nameorg=? and sklad=?'
        QueryParams=[int(model[path][4]),inn.decode('utf-8'),nameOrg.decode('utf-8'),sklad.decode('utf-8'),]

        sys.ms71_db.updateVigr(strUpd,QueryParams)

        return


#функция для работы смены знака на кнопку
#получает данные из selection = self.tvgrid.get_selection()
    def ChangeVigr(self,widget):
        selection = self.tvgrid.get_selection()
        model, iter, = selection.get_selected()
        if iter:
            path = model.get_path(iter)
            #print model
            vigr=0
            if int(model.get_value(iter,3))==1:
                vigr=0
            if int(model.get_value(iter,3))==0:
                vigr=1

            model.set_value(iter, 3, vigr)

            nameOrg=model.get_value(iter,1).replace('&apos;',"'")
            nameOrg=nameOrg.replace('&quot;','"')

            sklad=model.get_value(iter,2).replace('&apos;',"'")
            sklad=sklad.replace('&quot;','"')

            strUpd='update vigr set vigr=? where inn=? and nameorg=? and sklad=?'
            QueryParams=[vigr,model.get_value(iter,0).decode('utf-8'),nameOrg.decode('utf-8'),sklad.decode('utf-8'),]
            sys.ms71_db.updateVigr(strUpd,QueryParams)
            #model.remove(iter)
      # now that we removed the selection, play nice with
      # the user and select the next item
        selection.select_path(path)
#~ #~
      # well, if there was no selection that meant the user
      # removed the last entry, so we try to select the
      # last item
        if not selection.path_is_selected(path):
            row = path[0]-1
            # test case for empty lists
            if row >= 0:
                selection.select_path((row,))


    def __getattr__(self, attr):
        # удобней писать self.window1, чем self.get_object('window1')
        obj = self.builder.get_object(attr)
        if not obj:
            raise AttributeError('object %r has no attribute %r' % (self,attr))
        #if not obj:
        #obj = lambda *args: self._method_any(attr, args)
            #raise AttributeError('object %r has no attribute %r' % (self,attr))
        setattr(self, attr, obj)
        return obj

    def _method_any(self, attr, args):
        #print '_method_any:', self, attr, args
        #print '  %s' % attr, args
        try:
            pytxt = args[0].get_tooltip_text().strip()
        except:
            pytxt = None
        if pytxt:
            exec pytxt

if __name__=='__main__':
    adderGui = adderGui()
    adderGui.window.show()
    gtk.main()





