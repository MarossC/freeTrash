import os, configparser, winreg, zipfile, zlib, shutil, pathlib, numpy, ctypes, sys
from datetime import datetime
from typing import Union
config = configparser.ConfigParser()

try:
    is_admin = ctypes.windll.shell32.IsUserAnAdmin()
except:
    is_admin = False

if not is_admin:
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
    sys.exit()


config.read("config.cfg")

if not os.path.isfile("config.cfg"):
    print("Config does not exist, generating.")
    config["Location"] = {"Desktop": "",
                          "Downloads": "",
                          "Output": ""}
    with open("config.cfg","w") as fileconfig:
        config.write(fileconfig)

if not os.path.isdir(config["Location"]["Desktop"]):
    print("Desktop not present in config, setting Desktop location from registry.")
    config["Location"]["Desktop"] = winreg.QueryValueEx(winreg.OpenKeyEx(winreg.HKEY_CURRENT_USER,
                                                                           r'SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Folders'),
                                                                           r'{Desktop}')[0]
    with open("config.cfg","w") as fileconfig:
        config.write(fileconfig)

if not os.path.exists(config["Location"]["Downloads"]):
    print("Downloads not present in config, setting location from registry.")
    config["Location"]["Downloads"] = winreg.QueryValueEx(winreg.OpenKeyEx(winreg.HKEY_CURRENT_USER,
                                                                           r'SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Folders'),
                                                                           r'{374DE290-123F-4565-9164-39C4925E467B}')[0]
    with open("config.cfg","w") as fileconfig:
        config.write(fileconfig)

if not os.path.exists(config["Location"]["Output"]):
    while True:
        oInput = input("Output not present in config, please enter:").replace('"', "")
        if os.path.exists(oInput):
            config["Location"]["Output"] = oInput
            print("\n")
            break
        else:
            print("Incorrect path, try again please.")
    with open("config.cfg","w") as fileconfig:
        config.write(fileconfig)

lDesktop = config["Location"]["Desktop"]
lDownloads = config["Location"]["Downloads"]
lOutput = config["Location"]["Output"]
lFinalOutput = os.path.join(lOutput, datetime.today().strftime('%Y-%m-%d'))

if os.path.exists(os.path.join(lOutput, datetime.today().strftime('%Y-%m-%d'))):
    lFinalOutput = lFinalOutput + "-" + str(numpy.random.default_rng().integers(low=0, high=9999, size=1))
    os.mkdir(lFinalOutput)

######################## Funcs


def zip_dir(dir: Union[pathlib.Path, str], filename: Union[pathlib.Path, str]): # credits: https://stackoverflow.com/questions/1855095/how-to-create-a-zip-archive-of-a-directory
    dir = pathlib.Path(dir)

    with zipfile.ZipFile(filename, "w", zipfile.ZIP_DEFLATED) as zip_file:
        for entry in dir.rglob("*"):
            zip_file.write(entry, entry.relative_to(dir))
    os.system("rmdir /s /q "+str(dir))

"""
def zipdir(path, ziph): # credits: https://stackoverflow.com/questions/1855095/how-to-create-a-zip-archive-of-a-directory
    for root, dirs, files in os.walk(path):
        for file in files:
            ziph.write(os.path.join(root, file), 
                       os.path.relpath(os.path.join(root, file), 
                                       os.path.join(path, '..')))
"""

def checkDir(dir):
    kbdskip = False
    os.system("attrib -r -h "+ dir +"\\*.* /s")

    for iter in os.listdir(dir):
        iterpath = os.path.join(dir, iter)

        if iter == "desktop.ini":
            continue
        
        if os.path.isdir(iterpath):
        
            while True:
                print(iterpath)
                kbdinput = input(iter + " => Is a directory; (A)rchive/(M)ove/(D)elete/(I)gnore: ")
                if kbdinput.lower() == "a":
                    zip_dir(iterpath, os.path.join(lFinalOutput,iter + ".zip"))
                    break
                
                if kbdinput.lower() == "m":
                    os.mkdir(os.path.join(lFinalOutput, iter))
                    shutil.move(iterpath, lFinalOutput)
                    break

                if kbdinput.lower() == "d":
                    try:
                        shutil.rmtree(iterpath)
                    except OSError as e:
                        print("Error: %s - %s." % (e.filename, e.strerror))
                    break
                
                if kbdinput.lower() == "i":
                    break
                
                print("Incorrect input. Please try again.")
            continue
                

        if iter[-4::] == ".lnk" or iter[-4::] == ".url":
            print(iter + " => Is a shortcut, ignoring.")
            continue
        
        while True:
            if kbdskip:
                kbdinput = "m"
            else:
                kbdinput = input(iter + " => Is a file; (M)ove/(M)ove (A)ll/(D)elete/(I)gnore: ")

            if kbdinput.lower() == "m":
                shutil.move(iterpath, os.path.join(lFinalOutput, iter))
                if kbdskip:
                    print(iter + " => Is a file; included in Move all")
                break

            if kbdinput.lower() == "ma":
                shutil.move(iterpath, os.path.join(lFinalOutput, iter))
                kbdskip = True
                break

            if kbdinput.lower() == "d":
                os.remove(iterpath)
                break

            if kbdinput.lower() == "i":
                break

            print("Incorrect input. Please try again.")
        continue
    return
        


checkDir(lDesktop)
checkDir(lDownloads)

while True:
    final = input("Do you want to compress the moved files? (y/n) : ")

    if final.lower() == "y":
        zip_dir(lFinalOutput,lFinalOutput + ".zip")
        break
    elif final.lower() == "n":
        print("Not compressing.")
        break
    else:
        print("Wrong input. Please try again.")

print("Finished.")
input()





