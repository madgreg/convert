	# -*- coding: utf-8 -*-
import glob, os, sys, time, shutil   # ����������� ����������� ������
a, c, x = 123, "", 1
t = 10                        # ��������� �������� � ��������; 5400 - 1,5 ����

fost = open('ostatok.txt', 'a')

files = os.listdir('.');
tmp = []
for f in files:
	fost.write(f+'\n')
    
print 'end'
fost.close()

    
    