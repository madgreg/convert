#!/usr/bin/env python
# coding: utf-8

import sys
reload(sys)
sys.setdefaultencoding('utf-8')
import os
import rpc
import time
import getmail
import subprocess
import pycs.common
import mkmsg
import sndmsg
import datetime
import xmpp
import glob
import shutil
import traceback
import dropbox


IMP_PATH = pycs.common.IMP_PATH
EXP_PATH = pycs.common.EXP_PATH
UPCOMING = pycs.common.UPCOMING


def get_mail():
    res = rpc.send_request('get_imp_data', ['mail'])
    if res['result']:
        for cfg in res['result']:
            tmp = cfg[0].split(';')
            port = None
            if ':' in tmp[0]:
                server, port = tmp[0].split(':')
            else:
                server, port = tmp[0], 110
            #print server, port, tmp[1], tmp[2], tmp[3], os.path.join(IMP_PATH, cfg[1])
            path = os.path.join(IMP_PATH, cfg[1])
            if not os.path.exists(path):
                os.mkdir(path)
            print 'getting mail... ', path
            getmail.get_mail(server, int(port), tmp[1], tmp[2], tmp[3], path)


def get_ftp():
    # user, passwd, server, path, remote_path
    res = rpc.send_request('get_imp_data', ['ftp'])
    if res['result']:
        for cfg in res['result']:
            print " ---getFtp ",cfg
            tmp = cfg[0].split(';')
            path = os.path.join(IMP_PATH, cfg[1])
            if not os.path.exists(path):
                os.mkdir(path)
            cmd = 'ncftpget -DD -u %s -p %s %s %s \'%s\'' % (tmp[1], tmp[2], tmp[0], path, '\' \''.join(tmp[3:]))
            #print cmd
            p = subprocess.Popen(cmd, shell = True)
            p.wait()
            if 'pharmpartner71' in cfg[0]:
                pass
            elif 'lekrus77' in cfg[0]:
                pass
            else:
                for file in glob.glob(os.path.join(path, '*')):
                    os.rename(file, file.lower())
            #for file in os.listdir(path):
            #	print 1, file
            #	file = os.path.join(path, file.decode('cp866'))
            #	print 2, file
            #	os.rename(file, file.lower())
            '''
            if 'geofarm77' in cfg[0]:
                copy_path = os.path.join(IMP_PATH, '_geofarm77_copy')
                copy_path_t = os.path.join(copy_path, time.strftime('%Y-%m-%d-%H-%M'))
                if not os.path.exists(copy_path):
                    os.mkdir(copy_path)
                if not os.path.exists(copy_path_t):
                    os.mkdir(copy_path_t)
                for file in glob.glob(os.path.join(path, '*.dbf')):
                    shutil.copy(file, copy_path_t)
                if not os.listdir(copy_path_t):
                    os.rmdir(copy_path_t)
            '''

def get_path():
    res = rpc.send_request('get_imp_data', ['path'])
    if res['result']:
        for cfg in res['result']:
            print "---- get_path", cfg
            tmp = cfg[0].split(';')
            path = os.path.join(IMP_PATH, cfg[1])
            if not os.path.exists(path):
                os.mkdir(path)
            for inc_path in tmp:
                for file in glob.glob(inc_path.encode('utf-8')):
                    #print file, type(file), path, type(path)
                    shutil.move(file, path.encode('utf-8'))

def run_conv(mod, path, exp_path):
    if not os.path.exists(exp_path):
        os.mkdir(exp_path)
    mod = __import__(mod)
    if os.path.exists(IMP_PATH+'/'+path):
        path = os.path.join(IMP_PATH, path)
        for name in os.listdir(path):
            filename = os.path.join(path, name)
            #print os.path.splitext(filename)[1]
            if os.path.splitext(filename)[1][1:2] not in ('~', '!'):
                if os.path.isfile(filename):
                    #print filename
                    try:
                        mod.conv(filename, exp_path)
                        os.rename(filename, os.path.splitext(filename)[0] + '.~' + os.path.splitext(filename)[1][1:])
                    except:
                        os.rename(filename, os.path.splitext(filename)[0] + '.!' + os.path.splitext(filename)[1][1:])

def run_imp_new(mod, path, exp_path, cl_id = None, vnd_id = None):
    mod = __import__(mod)
    #print 'run_imp RUN - ',IMP_PATH+'/'+path
    import pycs.dbworker as dbw
    dbworker = dbw.DBWorker()

    if os.path.exists(IMP_PATH+'/'+path):
        path = os.path.join(IMP_PATH, path)
        for name in os.listdir(path):
            filename = os.path.join(path, name)
            if os.path.splitext(filename)[1][1:2] not in ('~', '!'):
                if os.path.isfile(filename):
                #    print '---run_imp_new'
                    tmpname = name.split("+")
                    sellid = dbworker.get_ids_by_mail(tmpname[0])
                    
                    if sellid == vnd_id:
                	cupid = dbworker.get_idc_by_mail(tmpname[1])
                	print cupid, tmpname[1]
                	if cupid != 0:
                	    print name                    
        	            try:
        	        	newname = os.path.join(path, str(cupid)+'_'+name)
            		        os.rename(filename, newname)
                	        filename = newname
                	        print filename
                	        mod.imp(filename, exp_path, cupid, vnd_id )
	                        print "--- run_imp imp end"
    		                newname = os.path.splitext(filename)[0] + '.~' + os.path.splitext(filename)[1][1:]
            		        if os.path.exists(newname):
                    	    	     newname = newname = os.path.splitext(filename)[0] + time.strftime('%Y%m%d%H%M%S') + '.~' + os.path.splitext(filename)[1][1:]
                    		os.rename(filename, newname)
                	    except Exception, e:
                    		print "---run_imp2", filename#, newname
                    		traceback.print_exc()
                    		print "---run_imp3", str(e)
                    		newname = os.path.splitext(filename)[0] + '.!' + os.path.splitext(filename)[1][1:]
                		if os.path.exists(newname):
            			    newname = os.path.splitext(filename)[0] + time.strftime('%Y%m%d%H%M%S') + '.!' + os.path.splitext(filename)[1][1:]
                    		os.rename(filename, newname)
                		continue
                    


def run_imp(mod, path, exp_path, cl_id = None, vnd_id = None):
    mod = __import__(mod)
    #print 'run_imp RUN - ',IMP_PATH+'/'+path

    if os.path.exists(IMP_PATH+'/'+path):
        path = os.path.join(IMP_PATH, path)
        for name in os.listdir(path):
            filename = os.path.join(path, name)
            if os.path.splitext(filename)[1][1:2] not in ('~', '!'):
                if os.path.isfile(filename):
                    bn = os.path.basename(filename)
                    if bn[0] == '.' and bn[-5:] == '.part':
                        continue
                    print '---run_imp1', filename
                    
                    try:
                        mod.imp(filename, exp_path, cl_id, vnd_id )
                        print "--- run_imp imp end"
                        newname = os.path.splitext(filename)[0] + '.~' + os.path.splitext(filename)[1][1:]
                        if os.path.exists(newname):
                            newname = newname = os.path.splitext(filename)[0] + time.strftime('%Y%m%d%H%M%S') + '.~' + os.path.splitext(filename)[1][1:]
                        os.rename(filename, newname)
                    except Exception, e:
                        print "---run_imp2", filename#, newname
                        traceback.print_exc()
                        print "---run_imp3", str(e)
                        newname = os.path.splitext(filename)[0] + '.!' + os.path.splitext(filename)[1][1:]
                        if os.path.exists(newname):
                            newname = os.path.splitext(filename)[0] + time.strftime('%Y%m%d%H%M%S') + '.!' + os.path.splitext(filename)[1][1:]
                        os.rename(filename, newname)
                        continue

def run_exp():
    res = rpc.send_request('get_unexp_invoices')
    if res['result']:
        for line in res['result']:
            try:
                ger_res = rpc.send_request('export_invoice', [line[0]])
                if ger_res['result']:                    
                    head, body = ger_res['result']
                    #print head
                    #print body
                    mod = __import__(line[5])
                    print "---run_exp--", line[4]
                    # создадим папку на случай если ее нет
                    if line[4].startswith("/"):
                        exp_path = line[4]
                    else:
                        exp_path = os.path.join(EXP_PATH, line[4])
                    if not os.path.exists(exp_path):
                        os.mkdir(exp_path)
                    print 'exp_path = ', exp_path
                    print 'run_exp ', line[0], line[2], line[3], line[6], exp_path, line[7]
                    if mod.exp(line[0], line[2], line[3], exp_path, line[6], head,
                        body, seller=line[7]):
                        print 'GO'
                        rpc.send_request('update_exp_status', [line[0]])
			print 'GO_END'
	    except Exception, e:
                traceback.print_exc()
                print str(e)
                print 'ошибка'
                continue


def jabber_report(msg):
    FROM_JID = "status@ms71.ru"
    JID_PASS = "eR#z*U5"
    J_SERVER = "localhost"
    TO_JID = "maslov@ms71.ru"
    jid = xmpp.protocol.JID(FROM_JID)
    cl = xmpp.Client(jid.getDomain(), debug = [])
    if not cl.connect((J_SERVER,5222)):
        raise IOError('Can not connect to server.')
    if not cl.auth(jid.getNode(),JID_PASS, 'plx'):
        raise IOError('Can not auth with server.')

    cl.send(xmpp.Message(TO_JID, msg, 'chat'))
    #cl.send( xmpp.Message( TO_JID2, datetime.datetime.now().strftime("%H:%M") + " OK", 'chat' ) )
    cl.disconnect()

def stconv():
    """
    конвертируем накладные в обход базы, письма создаем самостоятельно
    """
    print 'stconv run'
    run_conv('katrendbf2ahold', 'edoc20577', os.path.join(EXP_PATH, 'matveeva_protvino'))
    run_conv('profitmed2ahold', 'edoc22077', os.path.join(EXP_PATH, 'matveeva_protvino'))
    run_conv('apteka2k2ahold', 'edoc28177', os.path.join(EXP_PATH, 'matveeva_protvino'))
    #run_conv('apteka2k2ahold', os.path.join(IMP_PATH, 'katren_matveeva'), os.path.join(EXP_PATH, 'matveeva_protvino'))
    run_conv('apteka2k2ahold', 'edoc20377-10677', os.path.join(EXP_PATH, 'matveeva_protvino'))
    run_conv('nadezhdadbf2ahold', '34157-40077', os.path.join(EXP_PATH, 'matveeva_protvino'))
    run_conv('imperiadbf2ahold', '23477-40077', os.path.join(EXP_PATH, 'matveeva_protvino'))
    print 'stconv 3'
    mkmsg.strun(EXP_PATH, 40077)
    print 'stconv end'
    
def unZip(sp_serch):
		import zipfile 
		print 'RUN UNZIP %s' % sp_serch
		try:
		    files = os.listdir(sp_serch);
		    zipfiles = filter(lambda x: x.endswith(('.zip', '.ZIP')), files)
		    for t in zipfiles:
			    farr = t.split('+')
			    zfile = zipfile.ZipFile( '%s/%s' % (sp_serch, t), "r" )
			    for info in zfile.infolist():
				    fname = info.filename
				    data = zfile.read(fname)
				    filename = sp_serch + '/'+ farr[0]+ '+' + farr[1]+ '+' + fname
				    fout = open(filename, "w")
				    fout.write(data)
				    fout.close()
			    zfile.close()
			    #print '%s/%s' % (sp_serch, t), '%s/%s' % (sp_serch, t)+'.old'
			    os.rename('%s/%s' % (sp_serch, t), '%s/%s' % (sp_serch, t)+'.old')
		except Exception, e:
                    print str(e)
		print 'END UNZIP %s' % sp_serch

	
if __name__ == '__main__':
    try:
        print "START"
        get_path()
	print 1
        #dropbox.imp()
        print "IMPORT RUN"
        unZip(IMP_PATH+"/edocs")
        import pycs.dbworker as dbw
	dbworker = dbw.DBWorker()

        res = dbworker.get_seller_list()
	#print res
	for line in res:
	    if line[5] != '' and "_sklad" in line[5]:
		print line[5], line[0]
    		run_imp_new(line[5], 'edocs', None, vnd_id = line[0])
        #print 'NEXT Level'
	run_imp_new('profidmed_sklad', 'edocs', None, vnd_id = 22077)
#        run_imp_new('nadejda_f_sklad', 'edocs', None, vnd_id = 34157)
        
        run_imp('alliance-k', 'edoc31971', None,
            vnd_id = 31971)        
        run_imp('csmedika', 'edoc30371', None,
            vnd_id = 30371)
        run_imp('brizxls', 'edoc31471', None, vnd_id = 31471)
        run_imp('geofarm77', 'geofarm77', None, vnd_id = 21177)
        run_imp('genezis77', 'genezis77', None, vnd_id = 21777)
        run_imp('rosfarm', 'edoc29477', None, vnd_id = 29477)
        run_imp('vektorxls', 'edoc28871', None, vnd_id = 28871)
        run_imp('torg12-belita', 'edoc31871', None, vnd_id = 31871)
        run_imp('torg12-kosmetikopt', 'edoc33571', None, vnd_id = 33571)
        run_imp('pharmp_sst', 'pharmpartner71', None, vnd_id = 21271)
        run_imp('ee_holding_xls', 'edoc33571', None, vnd_id = 33677)
        run_imp('torg12-kosmetikopt', 'edoc33771', None, vnd_id = 33771)
        run_imp('epihin-xls', 'edoc33971', None, vnd_id = 33971)
        run_imp('protek_sst', 'protek71', None, vnd_id = 20271)
        run_imp('torg12-demidovskaya', '34271', None, vnd_id = 34271)
        run_imp('lekrus77', 'lekrus77', None, vnd_id = 33877)
        run_imp('torg12-bobkov', '34371', None, vnd_id = 34371)
        run_imp('kolinz71', '34471', None, vnd_id = 34471)
        run_imp('torg12-isfera', '34677', None, vnd_id = 34677)
        run_imp('lugovoi_metiz_xls', 'lugovoi_metiz', None, vnd_id=36171)
        run_imp('legattm-torg12', 'legattm', None, vnd_id=36271)
        run_imp('gribachev-torg12', 'gribachev', None, vnd_id=36471)
        run_imp('grandpostavka-xls', 'grandpostavka', None, vnd_id=36377)
        run_imp('avtomatika-torg12', 'avtomatika', None, vnd_id=36571)
        run_imp('nikolaev-torg12', 'nikolaev-ip71', None, vnd_id=36671)
        run_imp('himservis-xls', 'himservis', None, vnd_id=36871)
        run_imp('batteryteam-xls', 'batteryteam71', None, vnd_id=37071)
        run_imp('abretov-torg12', 'abretov71', None, vnd_id=36971)
        run_imp('master-xls', '37271', None, vnd_id = 37271)
        run_imp('yakshina_xls', '37371', None, vnd_id = 37371)
        
        run_imp('torg12-okey', 'edoc37571', None, vnd_id = 37571)
	run_imp('torg12-vesna', '40777', None, vnd_id = 40777)
        run_imp('torg12-lipovsk', 'edoc28971', None, vnd_id = 28971)
        run_imp('torg12-stelmas77', 'edoc39177', None, vnd_id = 39177)
        run_imp('torg12-bosomi', 'edoc39577', None, vnd_id = 39577)
        run_imp('torg12-evroklin', 'evroklin', None, vnd_id = 39771)
        run_imp('torg12-imperiya71', 'imperiya71', None, vnd_id = 39871)
        run_imp('torg12-doronin71', 'edoc40171', None, vnd_id = 40171)
        run_imp('torg12-nashamarka', 'nashamarka', None, vnd_id = 40377)
        run_imp('torg12-kavrigin', 'kavrigin', None, vnd_id = 40471)
        run_imp('torg12-azbuka', 'azbuka_insh', None, vnd_id = 41071)

        print 'running  export...'
        #run_exp1()
        run_exp()
        print 'making messages...'
        mkmsg.run(EXP_PATH)
        print 'stand alone converting...'
        stconv()
        print 'sending messages...'
       # print ' --- ', os.path.join(EXP_PATH, UPCOMING)
        sndmsg.send_messages(os.path.join(EXP_PATH, UPCOMING))
        print 'sending report to jabber...'
        jabber_report(datetime.datetime.now().strftime('%Y-%m-%d %H:%M') + \
            ' pyconv status OK')
    except Exception, e:
        import sys, traceback
        jabber_report(datetime.datetime.now().strftime('%Y-%m-%d %H:%M') + ' ' + str(e) + '\n' + traceback.format_exc())
