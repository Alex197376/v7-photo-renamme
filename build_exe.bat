@echo off
cd /d "C:\Users\USER\Documents\Python\renomage photo v 6.2  auto apprentissage"

pyinstaller ^
--noconfirm ^
--clean ^
--windowed ^
--onefile ^
--hidden-import=win32com.client ^
--hidden-import=pythoncom ^
--add-data "corrections.json;." ^
--add-data "button_presets.json;." ^
Renommage_Photo_V6.1.py

pause
