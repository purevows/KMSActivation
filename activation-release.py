#coding=utf-8
import os
import sys
import ctypes
import platform
import subprocess
import searchwinreg
import shutil

def checkisoinput(setuplist, choose):
	try:
		if(choose == '0'):
			return 0;
		if(setuplist[int(choose) - 1]):
			return 1;
	except:
		return 0;

def getsetupiso():
	setuplist = [];
	localpath = os.getcwd();
	remotepath = '\\\\192.168.3.103\\public\\Tools\\Imagine';
	print("当前内网服务端提供的安装包有：")
	for root, dirs, files in os.walk(remotepath):
		if 'SW_DVD5_Office_Professional_Plus_2013_64Bit_ChnSimp_MLF_X18-55285.ISO' in files:
			setuplist.append("Office2013 Pro Plus VOL");
		if 'SW_DVD5_Visio_Pro_2013w_SP1_64Bit_ChnSimp_MLF_X19-36392.ISO' in files:
			setuplist.append("Office2013 Visio Pro VOL");
		if 'SW_DVD5_Office_Professional_Plus_2010_64Bit_ChnSimp_MLF_X16-52534.ISO' in files:
			setuplist.append("Office2010 Pro Plus VOL");
		if 'SW_DVD5_Office_Professional_Plus_2016_64Bit_ChnSimp_MLF_X20-42426.ISO' in files:
			setuplist.append("Office2016 Pro Plus VOL");
		if 'SW_DVD5_Visio_Pro_2016_64Bit_ChnSimp_MLF_X20-42759.ISO' in files:
			setuplist.append("Office2016 Visio Pro VOL");
	num = 1;flag = 0;
	for i in setuplist:
		print(num, i);
		num = num + 1;
	choose = input("请选择版本:");
	while(checkisoinput(setuplist, choose) == 0):
		choose = input("请选择版本:");
	if(setuplist[int(choose) - 1] == 'Office2013 Pro Plus VOL'):
		print("正在下载安装包...，下载完成前请不要关闭本工具");
		shutil.copy(remotepath + '\\SW_DVD5_Office_Professional_Plus_2013_64Bit_ChnSimp_MLF_X18-55285.ISO', localpath);
	elif(setuplist[int(choose) - 1] == 'Office2013 Visio Pro VOL'):
		print("正在下载安装包...，下载完成前请不要关闭本工具");
		shutil.copy(remotepath + '\\SW_DVD5_Visio_Pro_2013w_SP1_64Bit_ChnSimp_MLF_X19-36392.ISO', localpath);
	elif(setuplist[int(choose) - 1] == 'Office2010 Pro Plus VOL'):
		print("正在下载安装包...，下载完成前请不要关闭本工具");
		shutil.copy(remotepath + '\\SW_DVD5_Office_Professional_Plus_2010_64Bit_ChnSimp_MLF_X16-52534.ISO', localpath);
	elif(setuplist[int(choose) - 1] == 'Office2016 Pro Plus VOL'):
		print("正在下载安装包...，下载完成前请不要关闭本工具");
		shutil.copy(remotepath + '\\SW_DVD5_Office_Professional_Plus_2016_64Bit_ChnSimp_MLF_X20-42426.ISO', localpath);
	elif(setuplist[int(choose) - 1] == 'Office2016 Visio Pro VOL'):
		print("正在下载安装包...，下载完成前请不要关闭本工具");
		shutil.copy(remotepath + '\\SW_DVD5_Visio_Pro_2016_64Bit_ChnSimp_MLF_X20-42759.ISO', localpath);
	print("文件获取完成，文件存放于本工具所在目录（" + localpath + "),现在可以关闭本工具。请将安装包用ISO工具打开或压缩工具解压后运行setup.exe进行安装，然后重新运行本工具进行激活。")
	os.system("pause");

def checkpathinput(path):
	if(path == 'exit'):
		return 0;
	else:
		for i in path:
			if(i == ''):
				return -1;
	return 1;

def scanospp(officelist, path, ver, exe):
	global officepathflag
	target = 'OSPP.VBS';
	if(ver == exe == 0):
		for root, dirs, files in os.walk(path):
			if target in files:
				if((root.find('Office15')) != -1):
					officelist['15.0'] = {'Unknow':[root, '1']};
					officepathflag = 1;
				elif((root.find('Office16')) != -1):
					officelist['16.0'] = {'Unknow':[root, '1']};
					officepathflag = 1;
				elif((root.find('Office14')) != -1):
					officelist['14.0'] = {'Unknow':[root, '1']};
					officepathflag = 1;
				elif((root.find('Office17')) != -1):
					officelist['15.0'] = {'Unknow':[root, '1']};
					officepathflag = 1;
	else:
		for root, dirs, files in os.walk(path):
			if target in files:
				officelist[ver][exe][1] = '1';
				officepathflag = 1;
			else:
				if 'WINWORD.EXE' in files:
					officelist[ver][exe][1] = '2';	
	return officelist;

def actwindows(act):
	global kmshost, sysver;
	if(act == '1'):
		os.system("slmgr /dlv");
		windowsmenu();
	elif(act == '2'):
		if(sysver == '7'):
			os.system("slmgr /ipk FJ82H-XT6CR-J8D7P-XQJJ2-GPDD4");
			windowsmenu();
		elif(sysver == '8'):
			os.system("slmgr /ipk NG4HW-VH26C-733KW-K6F98-J8CK4");
			windowsmenu();
		elif(sysver == '8.1'):
			os.system("slmgr /ipk GCRJD-8NW9H-F2CDX-CCM8D-9D6T9");
			windowsmenu();
		elif(sysver == '10'):
			os.system("slmgr /ipk W269N-WFGWX-YVC9B-4J6C9-T83GX");
			windowsmenu();
	elif(act == '3'):
		os.system("slmgr /upk");
		windowsmenu();
	elif(act == '4'):
		if(kmshost == '0'):
			os.system("slmgr /skms 192.168.3.103");
			print("未设置kms服务器地址，将通过默认服务器激活");
		os.system("slmgr /ato");
		windowsmenu();
	elif(act == '5'):
		if(kmshost == '0'):
			print("当前为默认kms服务器");
		else:
			print("当前kms服务器设置为：" + kmshost);
		kmshost = input("请输入kms服务器地址，请自行确认所输入的地址是否有效，输入0重置为默认服务器 ");
		if(kmshost == '0'):
			os.system("slmgr /skms 192.168.3.103");
		else:
			os.system("slmgr /skms " + kmshost);
		windowsmenu();
	elif(act == '6'):
		os.system("slmgr /xpr");
		windowsmenu();
	elif(act == '7'):
		mainmenu();
	else:
		print("喵喵喵喵喵?");
		windowsmenu();
	os.system("pause");

def actoffice(act):
	global kmshost, selectver, rootlist;
	ospppath = "cscript \"" + rootlist[selectver] + "\\OSPP.VBS\"";
	if(act == '1'):
		print(subprocess.check_output(ospppath + " /dstatus", universal_newlines = True));
		officemenu();
	elif(act == '2'):
		print("安装" + selectver + "的GVLK");
		if(selectver == 'Office2013'):
			print(subprocess.check_output(ospppath + " /inpkey:YC7DK-G2NP3-2QQC3-J6H88-GVGXT", universal_newlines = True));
			officemenu();
		elif(selectver == 'Office2016'):
			print(subprocess.check_output(ospppath + " /inpkey:XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99", universal_newlines = True));
			officemenu();
		elif(selectver == 'Office2010'):
			print(subprocess.check_output(ospppath + " /inpkey:VYBBJ-TRJPB-QFQRF-QFT4D-H3GVB", universal_newlines = True));
			officemenu();
		elif(selectver == 'Office2019'):
			print(subprocess.check_output(ospppath + " /inpkey:NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP", universal_newlines = True));
			officemenu();
		elif(selectver == 'Office2010 Visio'):
			print(subprocess.check_output(ospppath + " /inpkey:7MCW8-VRQVK-G677T-PDJCM-Q8TCP", universal_newlines = True));
			officemenu();
		elif(selectver == 'Office2013 Visio'):
			print(subprocess.check_output(ospppath + " /inpkey:C2FG9-N6J68-H8BTJ-BW3QX-RM3B3", universal_newlines = True));
			officemenu();
		elif(selectver == 'Office2016 Visio'):
			print(subprocess.check_output(ospppath + " /inpkey:PD3PC-RHNGV-FXJ29-8JK7D-RJRJK", universal_newlines = True));
			officemenu();
	elif(act == '3'):
		if(kmshost == '0'):
			print("未设置kms服务器地址，将通过默认服务器激活");
			print(subprocess.check_output(ospppath + " /sethst:192.168.3.103", universal_newlines = True));
		print(subprocess.check_output(ospppath + " /act", universal_newlines = True));
		officemenu();
	elif(act == '4'):
		if(kmshost == '0'):
			print("当前为默认kms服务器");
		else:
			print("当前kms服务器设置为：" + kmshost);
		kmshost = input("请输入kms服务器地址，请自行确认所输入的地址是否有效，输入0重置为默认服务器 ");
		if(kmshost == '0'):
			print(subprocess.check_output(ospppath + " /sethst:192.168.3.103", universal_newlines = True));
		else:
			print(subprocess.check_output(ospppath + " /sethst:" + kmshost, universal_newlines = True));
		officemenu();
	elif(act == '5'):
		officevermenu();
	elif(act == '6'):
		mainmenu();
	else:
		print("喵喵喵喵喵?");
		officemenu();
	os.system("pause")

def checkofficeverinput(choose, num):
	try:
		choose = int(choose);
		for i in range(1, num + 1):
			if(i == choose):
				return 1;
		return 0;
	except:
		return 0;

def officevermenu():
	global rootlist, officepathflag, officelist, selectver;
	flag = 0;#是否只找到零售版
	if(officepathflag == 0):
		print("正在获取Office安装路径...");
		path = '';
		officelist = searchwinreg.searchwinreg(officelist);
		for ver in officelist:
			for exe in officelist[ver]:
				try:
					if(officelist[ver][exe][1] == '0'):
						path = officelist[ver][exe][0];
						officelist = scanospp(officelist, path, ver, exe);
				except:
					continue;
		if(path == ''):
			path1 = 'c:\\Program Files\\Microsoft Office';
			path2 = 'c:\\Program Files (x86)\\Microsoft Office';
			officelist = scanospp(officelist, path1, 0, 0);
			officelist = scanospp(officelist, path2, 0, 0);
		if(officepathflag == 0):
			for ver in officelist:
				for exe in officelist[ver]:
					try:
						if(officelist[ver][exe][1] == '2'):
							flag = 1;#只找到了零售版
					except:
						continue;
			if(flag == 0):
				print("貌似没找到Office呢QwQ，是不是注册表信息丢失了呢");
			elif(flag == 1):
				print("貌似只找到了零售版的Office呢，不支持零售版的激活呢，是不是注册表信息丢失了呢QAQ");
			path3 = input("手动输入office所在的目录，如C:\\\\Program Files\\\\Microsoft Office,要用双反斜杠分隔目录哦~,输入exit返回主菜单\n");
			while(officepathflag == 0):
				if(checkpathinput(path3) == 1):
					officelist = scanospp(officelist, path3, 0, 0);
					if(officepathflag == 0):
						print("这个目录下好像没找到可用于激活的office组件呢");
						path3 = input("请重新输入一个路径，输入exit返回主菜单\n");
				elif(checkpathinput(path3) == 0):
					mainmenu();
				elif(checkpathinput(path3) == -1):
					path3 = input("请输入合法的路径,输入exit返回主菜单\n");
	print("当前安装的Office版本有：");
	a = 1;
	for i in officelist:
		for j in officelist[i]:
			try:
				if(officelist[i][j][1] == '1'):
					if(j == 'Visio'):
						if(i == '14.0'):
							print(a, 'Office2010 ' + j);
							rootlist['Office2010 ' + j] = officelist[i][j][0];
							a = a + 1;
						elif(i == '15.0'):
							print(a, 'Office2013 ' + j);
							rootlist['Office2013 ' + j] = officelist[i][j][0];
							a = a + 1;
						elif(i == '16.0'):
							print(a, 'Office2016 ' + j);
							rootlist['Office2016 ' + j] = officelist[i][j][0];
							a = a + 1;
						elif(i == '17.0'):
							print(a, 'Office2019 ' + j);
							rootlist['Office2019 ' + j] = officelist[i][j][0];
							a = a + 1;
					elif(j == 'Unknow'):
						if(i == '14.0'):
							print(a, 'Office2010 (手动寻址，未能确定具体版本)');
							rootlist['Office2010 ' + j] = officelist[i][j][0];
							a = a + 1;
						elif(i == '15.0'):
							print(a, 'Office2013 (手动寻址，未能确定具体版本)');
							rootlist['Office2013 ' + j] = officelist[i][j][0];
							a = a + 1;
						elif(i == '16.0'):
							print(a, 'Office2016 (手动寻址，未能确定具体版本)');
							rootlist['Office2016 ' + j] = officelist[i][j][0];
							a = a + 1;
						elif(i == '17.0'):
							print(a, 'Office2019 (手动寻址，未能确定具体版本)');
							rootlist['Office2019 ' + j] = officelist[i][j][0];
							a = a + 1;
					else:
						if(i == '14.0'):
							print(a, 'Office2010');
							rootlist['Office2010'] = officelist[i][j][0];
							a = a + 1;
						elif(i == '15.0'):
							print(a, 'Office2013');
							rootlist['Office2013'] = officelist[i][j][0];
							a = a + 1;
						elif(i == '16.0'):
							print(a, 'Office2016');
							rootlist['Office2016'] = officelist[i][j][0];
							a = a + 1;
						elif(i == '17.0'):
							print(a, 'Office2019');
							rootlist['Office2019'] = officelist[i][j][0];
							a = a + 1;
				elif(officelist[i][j][1] == '2'):
					if(j == 'Visio'):
						if(i == '14.0'):
							print('* Office2010 ' + j + "零售版");
						elif(i == '15.0'):
							print('* Office2013 ' + j + "零售版");
						elif(i == '16.0'):
							print('* Office2016 ' + j + "零售版");
						elif(i == '17.0'):
							print('* Office2019 ' + j + "零售版");
					else:
						if(i == '14.0'):
							print('* Office2010' + "零售版");
						elif(i == '15.0'):
							print('* Office2013' + "零售版");
						elif(i == '16.0'):
							print('* Office2016' + "零售版");
						elif(i == '17.0'):
							print('* Office2019' + "零售版");
			except:
				continue;
	if(len(rootlist) > 1):
		choose = input("请选择要激活的版本：");
		while(checkofficeverinput(choose, len(rootlist)) == 0):
			choose = input("请选择要激活的版本：");
	else:
		choose = '1';
	m = 1;
	for n in rootlist:
		if(m == int(choose)):
			selectver = n;
		m = m + 1;
	officemenu();

def officemenu():
	print("当前选择版本为：" + selectver);
	print("1.查看激活信息\n2.安装GVLK\n3.激活Office\n4.设置kms服务器地址并激活\n5.重新选择版本\n6.返回主菜单");
	act = input("select a choice:");
	actoffice(act);

def windowsmenu():
	print("1.查看激活信息\n2.安装GVLK\n3.卸载当前key\n4.激活Windows\n5.设置kms服务器地址并激活\n6.显示当前激活的截止日期\n7.返回");
	act = input("select a choice:");
	actwindows(act);

def mainmenu():
	global debug;
	print("Tips:办公区域内的计算机，可直接选激活通过服务端与内置的GVLK激活VOL的Windows与Office");
	print("如安装的是零售版，由于License不同，请先将零售版更换为VOL版再使用本工具激活");
	print("若曾经使用其它的VOLkey进行过激活，请先安装GVLK再选择激活");
	print("请选择激活产品\n1.Windows(7 or later)\n2.Office(2010 or later)\n3.从内网服务端获取Office安装程序");
	choose = input("select a choice:");
	while((choose != '1') & (choose != '2') & (choose != '3') & (choose != 'debug')):
		print("喵喵喵喵喵?");
		choose = input("select a choice:");
	if(choose == '1'):
		windowsmenu();
	elif(choose == '2'):
		officevermenu();
	elif(choose == '3'):
		getsetupiso();
	elif(choose == 'debug'):
		print("debug on");
		debug = 1;
		mainmenu();

def main():
	global sysinfo, sysplatform, sysver, sysos, kmshost, rootlist, officepathflag, selectver, officelist;
	global debug;
	debug = 0;
	kmshost = '0';
	officepathflag = 0;
	sysinfo = platform.platform().split('-');
	sysplatform = sysinfo[0];
	sysver = sysinfo[1];
	sysos = sysinfo[2];
	selectver = 0;
	rootlist = {};
	officelist = {};
	print("当前操作系统版本:", sysplatform, sysver, "内部版本号:", sysos);
	print("本工具支持的产品有:\nwindows7以上的专业版\nOffice2010以上的专业版、Visio(VOL channel)，不支持零售版与其它Office单件的激活");
	print("请勿将本工具用于非法拷贝，请支持正版！");
	print("Created by ZERO-L,2018.11.20");
	print("changelog:\n2018.12.24:添加从办公区域内网服务端获取office安装包的功能\n2018.12.25:优化了部分提示，适当增加激活说明");
	mainmenu();

def check_admin():
	try:
		return ctypes.windll.shell32.IsUserAnAdmin();
	except:
		return 0;

if __name__ == '__main__':
	#main();#change it before install
	if(check_admin()):
		main();
	else:
		if(sys.version_info[0] == 3):
			ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, "", None, 1);
		else:
			ctypes.windll.shell32.ShellExecuteW(None, u"runas", unicode(sys.executable), "", None, 1);
