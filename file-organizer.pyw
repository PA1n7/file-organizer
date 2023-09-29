import os, json, time, winshell
from win32com.client import Dispatch

os.chdir("\\".join(os.getcwd().split("\\")[:-1]))

#check if it's first time user
if not "file-organizer.lnk" in os.listdir(winshell.startup()):
    #add to startup
    newPath = winshell.startup() + r"\file-organizer.lnk"
    target = os.getcwd() + r"\build\file-organizer\dist\file-organizer.exe"
    shell = Dispatch("WScript.Shell")
    shortcut = shell.CreateShortcut(newPath)
    shortcut.Targetpath = target
    shortcut.save()

#import all settings
settings = json.loads(open("settings.json").read())

#change to working directory
os.chdir(settings["cwd"])

#creating curr_files empty so it runs once on start
curr_files = []

def check_dir(path):
    #created a function that checks if a directory exists, if not it is then created
    if not os.path.exists(path):
        os.mkdir(path)

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
                os.rename(f"{settings['cwd']}\\{elem}", f"{newPath}\\{elem}")
    #added a sleep so it doesn't take that much CPU
    time.sleep(1)