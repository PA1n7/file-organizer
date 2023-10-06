import os, json, time, winshell, sys
from win32com.client import Dispatch
import win32api

drives = win32api.GetLogicalDriveStrings()
drives = drives.split('\000')[:-1]

def check_dir(path):
    #created a function that checks if a directory exists, if not it is then created
    if not os.path.exists(path):
        os.mkdir(path)

#create external folder variable
cwd = "\\".join(os.getcwd().split("\\")[:-1])
external_folder = cwd.split("\\")[0] + "\\file-organizer"

#check if it's first time user
if not "file-organizer.lnk" in os.listdir(winshell.startup()):
    #save directory of file
    os.chdir(cwd)
    #move settings file to be more accessible
    check_dir(external_folder)
    os.rename("settings.json", external_folder+"\\settings.json")
    #add to startup
    newPath = winshell.startup() + r"\file-organizer.lnk"
    target = os.getcwd() + r"\dist\file-organizer.exe"
    shell = Dispatch("WScript.Shell")
    shortcut = shell.CreateShortcut(newPath)
    shortcut.Targetpath = target
    shortcut.save()
    #set to administrator only run
    with open(newPath, "rb") as f2:
        ba = bytearray(f2.read())
    ba[0x15] = ba[0x15] | 0x20
    with open(newPath, "wb") as f3:
        f3.write(ba)
    sys.exit(1)

#import all settings
#checking all drives for settings file
for drive in drives:
    if os.path.exists(f"{drive}file-organizer\\settings.json"):
        settings = json.loads(open(f"{drive}file-organizer\\settings.json").read())

#change to working directory
os.chdir(settings["cwd"])

#creating curr_files empty so it runs once on start
curr_files = []

while True:
    if not curr_files == os.listdir():
        curr_files = os.listdir()
        for elem in curr_files:
            #checking if it is a folder
            if not "." in elem:
                newPath = settings["exports"]["folder"]
                check_dir(newPath)
                os.rename(f'{settings["cwd"]}\\{elem}', f'{newPath}\\{elem}')
            #creating variable for easier use
            extension = elem.split(".")[-1].lower()
            if extension in settings["exports"].keys():
                #checking for directory
                newPath = settings["exports"][extension]
                check_dir(newPath)
                #moving file
                counter = 0
                if not os.path.exists(f"{newPath}\\{elem}"):
                    os.rename(f"{settings['cwd']}\\{elem}", f"{newPath}\\{elem}")
    #added a sleep so it doesn't take that much CPU
    time.sleep(1)