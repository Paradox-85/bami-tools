@echo off
pip install pyinstaller openpyxl
pyinstaller --onefile --name equipment_unpivot equipment_unpivot.py
echo Done! Check dist\equipment_unpivot.exe