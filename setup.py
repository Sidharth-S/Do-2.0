from cx_Freeze import setup, Executable
import os,sys
import subprocess
import shutil
import yaml
import winreg

print("Setting up pre-requisites.")
PACKAGES = ['argparse','yaml','os','datetime','win32com','sys','logging','shutil']

installed_packages = subprocess.check_output([sys.executable, '-m', 'pip', 'freeze']).decode('utf-8')
installed_packages = installed_packages.split('\r\n')

EXCLUDES = {pkg.split('==')[0] for pkg in installed_packages if pkg != ''}
EXCLUDES.add('tkinter')

for i in PACKAGES:
    try:
        EXCLUDES.os.remove(i)
    except:pass

build_exe_options = {'packages':PACKAGES, #packages used
                     'excludes':list(EXCLUDES),
                     'include_files':['icon.ico','doo.bat','dow.bat',],
                     "build_exe" : "C:\\Program Files\\Prosid\\Do"}

path = "C:\\ProgramData\\Prosid\\Do"

if os.path.isdir("C:\\Program Files\\Prosid\\Do"):shutil.rmtree("C:\\Program Files\\Prosid\\Do")

if os.path.isfile(os.path.join(path,"dosettings.yaml")):
    sample = open('sampledo.yaml')
    sampledata = yaml.full_load(sample)

    existing = open(os.path.join(path,"dosettings.yaml"))
    existingdata = yaml.full_load(existing)

    for i,j in sampledata.items():
        existingdata[i] = j | existingdata[i] 

    with open(os.path.join(path,"dosettings.yaml"),'w') as file:
        documents = yaml.dump(existingdata, file)

else:
    os.makedirs(path,exist_ok=True)
    newfile = open(os.path.join(path,"dosettings.yaml"),'x')
    sampledata = yaml.full_load(open('sampledo.yaml'))
    with open(os.path.join(path,"dosettings.yaml"),'w') as file:
        documents = yaml.dump(sampledata, file)

base = None
executables = [Executable(
            "cli.py",
            target_name = "do",
            copyright="Copyright (C) 2022 Prosid",
            base=base,
            icon="icon.ico",
            shortcut_name="do",
            shortcut_dir="do",)]

print("Starting installation")

setup(
    name="do",
    version="2.0",
    author = "Sidharth S",
    description="Customizable macro tasker",
    executables=executables,
    options={
        "build_exe": build_exe_options,
    },
)

print("Adding software to registry...")
app_path_registry = winreg.ConnectRegistry(None,winreg.HKEY_CURRENT_USER)
app_path_key = winreg.OpenKey(app_path_registry,r"Software\Microsoft\Windows\CurrentVersion\App Paths")

do_key = winreg.CreateKeyEx(app_path_key,"do.exe",)
defv = winreg.SetValueEx(do_key,"",0,winreg.REG_SZ,r"C:\Program Files\Prosid\Do\do.exe")
pathv = winreg.SetValueEx(do_key,"Path",0,winreg.REG_SZ,r"C:\Program Files\Prosid\Do")


print("Installation complete !")
