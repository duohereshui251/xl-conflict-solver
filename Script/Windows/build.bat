@ECHO OFF
pyinstaller --onefile .\src\Diff\diff.py --name=ExlDiff.exe  --icon .\Script\Windows\logo.ico
pyinstaller --onefile .\src\merge\merge.py --name=ExlMerge.exe  --icon .\Script\Windows\logo.ico