#coding:utf-8
import winreg

def getpath(keypath, officelist, ver, exe):
    try:
        i = 0;
        while(1):
            if(winreg.EnumValue(keypath, i)[0] == 'Path'):
                officelist[ver][exe] = [winreg.EnumValue(keypath, i)[1], '0'];
            i = i + 1;
    except:
        return officelist;

def scanpath(key, officelist, target, ver, exe):
    try:
        if(key == 0):
            return officelist;
        i = 0;
        while(1):
            if(winreg.EnumKey(key, i) == 'InstallRoot'):
                keypath = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, target + "\\" + ver + "\\" + exe + "\\InstallRoot");
                officelist = getpath(keypath, officelist, ver, exe);
            i = i + 1;
    except:
        return officelist;

def scanexe(key, officelist, ver):
    try:
        if(key == 0):
            return officelist;
        i = 0;
        while(1):
            if(ver == '14.0'):
                if(winreg.EnumKey(key, i) == 'Word'):
                    officelist["14.0"] = {"Word":[]};
                if(winreg.EnumKey(key, i) == 'Visio'):
                    officelist["14.0"] = {"Visio":[]};
            elif(ver == '15.0'):
                if(winreg.EnumKey(key, i) == 'Word'):
                    officelist["15.0"] = {"Word":[]};
                if(winreg.EnumKey(key, i) == 'Visio'):
                    officelist["15.0"] = {"Visio":[]};
            elif(ver == '16.0'):
                if(winreg.EnumKey(key, i) == 'Word'):
                    officelist["16.0"] = {"Word":[]};
                if(winreg.EnumKey(key, i) == 'Visio'):
                    officelist["16.0"] = {"Visio":[]};
            elif(ver == '17.0'):
                if(winreg.EnumKey(key, i) == 'Word'):
                    officelist["17.0"] = {"Word":[]};
                if(winreg.EnumKey(key, i) == 'Visio'):
                    officelist["17.0"] = {"Visio":[]};
            i = i + 1;
    except:
        return officelist;

def scanver(key, officelist):
    try:
        if(key == 0):
            return officelist;
        i = 0;
        while(1):
            if(winreg.EnumKey(key, i) == '14.0'):
                officelist["14.0"] = {};
            elif(winreg.EnumKey(key, i) == '15.0'):
                officelist["15.0"] = {};
            elif(winreg.EnumKey(key, i) == '16.0'):
                officelist["16.0"] = {};
            elif(winreg.EnumKey(key, i) == '17.0'):
                officelist["17.0"] = {};
            i = i + 1;
    except:
        return officelist;

def searchwinreg(officelist):
    global debug;
    target1 = "SOFTWARE\\Microsoft\\Office";
    target2 = "SOFTWARE\\WOW6432Node\\Microsoft\\Office"
    try:
        key1 = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, target1);
    except:
        key1 = 0;
    try:
        key2 = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, target2);
    except:
        key2 = 0;
    #if(debug == 0):
    #   key1 = 0;
    #   key2 = 0;
    if(key1 == key2 == 0):
        #print("未找到office");
        return officelist;
    officelist = scanver(key1, officelist);
    officelist = scanver(key2, officelist);
    for ver in officelist:
        try:
            keytemp1 = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, target1 + "\\" + ver);
        except:
            keytemp1 = 0;
        try:
            keytemp2 = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, target2 + "\\" + ver);
        except:
            keytemp2 = 0;
        if(keytemp1 == keytemp2 == 0):
            continue;
        officelist = scanexe(keytemp1, officelist, ver);
        officelist = scanexe(keytemp2, officelist, ver);
        for exe in officelist[ver]:
            try:
                path1 = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, target1 + "\\" + ver + "\\" + exe);
            except:
                path1 = 0;
            try:
                path2 = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, target2 + "\\" + ver + "\\" + exe);
            except:
                path2 = 0;
            if(path1 == path2 == 0):
                continue;
            officelist = scanpath(path1, officelist, target1, ver, exe);
            officelist = scanpath(path2, officelist, target2, ver, exe);
    return officelist;

#officelist = {};
#office = serchwinreg(officelist)
#print(officelist);
