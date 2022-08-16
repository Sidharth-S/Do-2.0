import argparse
import shutil
import yaml
import os 
from win32com.client import Dispatch

class docli:
    def __init__(self) -> None:
        parser = argparse.ArgumentParser(description='All in one macro tasker')
        
        parser.add_argument('--list','-ls', 
                            action = "store_true",
                            dest   = "list",
                            help   = "List all saved settings")

        parser.add_argument('--version', 
                            action = "store_true",
                            dest   = "version",
                            help   = "View version of running application")

        parser.add_argument('--uninstall', 
                            action = "store_true",
                            dest   = "unins",
                            help   = "Uninstall the application.")

        parser.add_argument('--open','-o', 
                            action  = "store",
                            dest    = "open", 
                            metavar = "brave",
                            help    = "Run command with prescribed shortcut",)
                            
        parser.add_argument('--settings','-s', 
                            action = "store_true",
                            dest   = "settings",
                            help   = "Open settings of program / modify shortcuts")          

        parser.add_argument('--web','-w',
                            action  = "store",
                            dest    = "web",   
                            metavar = "youtube",
                            help    = "Open websites from predefined shortcuts",)

        args = parser.parse_args()

        self.file = 'C:\\ProgramData\\Prosid\\Do\\dosettings.yaml'
        self.args = args
        self.values = yaml.full_load(open('C:\\ProgramData\\Prosid\\Do\\dosettings.yaml'))


        if args.open :
            term = self.args.open
            if term not in self.values["open"] :  raise ValueError("No shortcut saved for term")
            array = self.values["open"][term] 

            for item in array:
                os.startfile(item)

        if args.settings :
            os.startfile(self.file)

        if args.web : 
            term = self.args.web
            if term  in self.values["web"] :  
                array = self.values["web"][term] 
            else: array = [term]
            bpath = self.values['websettings']['browserpath']
            webjoin = " ".join(array)
            os.system(f'"{bpath}" {webjoin}')
            

        if args.list :
            for i in self.values :
                print(f"{i} =>")
                print(self.values[i])
                print("")
            
        if args.version : 
            parser = Dispatch("Scripting.FileSystemObject")
            version = parser.GetFileVersion('C:\\Program Files\\Prosid\\Do\\do.exe')
            print(version)

        if args.unins : 
            x = str(input("Do you want to uninstall the application and its associated files irreversibly (Y/N): "))
            if x not in ['Y','y'] : return
            shutil.rmtree("C:\Program Files\Prosid\Do")
            x = input('Do you want to remove the settings file as well("Y/N") ?')
            if x not in ['Y','y'] : return
            shutil.rmtree('C:\ProgramData\Prosid\Do')
            print("This does not remove the registry key for your computer's safety.You may delete it by yourself by \
                   going to 'HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\App Paths\do.exe' in your registry editor")
            print("Do program removed.")
        
        return


if __name__ == "__main__":
    x = docli()
    os.abort()