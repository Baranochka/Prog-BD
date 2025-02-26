import os
import sys
import requests
import subprocess
from pathlib import Path
from tkinter import messagebox

__latest_version__ = None

class UpdateApp:
    def __init__(self, version: str) -> None:
        if self.check_update(version):
            if messagebox.askyesno("Обновление", "Доступно новое обновление. Хотите его установить?"):
                self.update()
    
    def check_update(self, version: str) -> bool:
        
        url = "https://api.github.com/repos/Baranochka/Prog-BD/releases/latest"
        try:
            response = requests.get(url)
            if response.status_code == 200:
                latest_release = response.json()
                global __latest_version__
                __latest_version__ = latest_release['tag_name']
                if __latest_version__ != version:
                    return True
                else:
                    return False
        except:
            return False


    def update(self) -> None:
        url = f"https://github.com/Baranochka/Prog-BD/releases/download/{__latest_version__}/Program.BD.exe"
        path_cur_exe = os.path.dirname(sys.executable)
        path_new_exe = os.path.join(path_cur_exe, 'update\\Program BD.exe')
        response = requests.get(url, stream=True)
        if response.status_code == 200:
            if os.path.isdir(".\\update"):
                pass
            else:
                os.mkdir(f".\\update")
            with open(path_new_exe, "wb") as file:
                for chunk in response.iter_content(1024):
                    file.write(chunk)
            if os.path.isfile(path_new_exe):
                self.update_self()
        else:
            return False


    def update_self(self) -> None:
        updater_script = "updater.bat"

        with open(updater_script, "w") as f:
            f.write(f"""
            @echo off
            setlocal
            set EXE_NAME=Program BD.exe
            set UPDATE_FOLDER=.\\update\\
            set TARGET_FOLDER=.\\

            :: Убиваем процесс, если он запущен
            taskkill /F /IM %EXE_NAME% > nul 2>&1
            
            timeout /t 2 /nobreak > nul

            copy "%UPDATE_FOLDER%%EXE_NAME%" 

            rmdir /Q /S "%UPDATE_FOLDER%"

            :: start "" "%TARGET_FOLDER%%EXE_NAME%"

            cmd /c del "%~f0"
            """)
        try:
            subprocess.Popen(updater_script, shell=True)
            sys.exit()
        except PermissionError as e:
            print(f"PermissionError: {e}")
            messagebox.showerror("Ошибка", f"Отказано в доступе: {e}")
