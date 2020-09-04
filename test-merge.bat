git checkout master
git branch -D demo-branch-1
git branch -D demo-branch-2
python.exe test/example.py 6
git add test/a.xlsx
git commit -m"master: change a.xlsx"

:: Create 'my-file.mrg' on branch 1
git checkout -b demo-branch-1
echo "change a.xlsx on demo-branch-1" 
python.exe test/example.py 7 -f
git add test/a.xlsx
git commit -m"demo-branch-1: change a.xlsx"

:: Create 'my-file.mrg' on branch 2
git checkout master
git checkout -b demo-branch-2 
echo "change a.xlsx on demo-branch-2" 
python.exe test/example.py 8 
git add test/a.xlsx
git commit -m"demo-branch-2: change a.xlsx"

:: Merge the two branches, causing a conflict
git merge -m"Merged in demo-branch-1" demo-branch-1
REM git reset --merge  