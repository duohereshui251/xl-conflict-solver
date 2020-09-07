@ECHO OFF
pyinstaller --onefile .\src\Diff\diff.py --name=ExlDiff-git.exe  --icon .\Script\Windows\logo.ico
pyinstaller --onefile .\src\merge\merge.py --name=ExlMerge-git.exe  --icon .\Script\Windows\logo.ico