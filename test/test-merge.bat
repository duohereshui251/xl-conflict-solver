git checkout master
git branch -D demo-branch-1
git branch -D demo-branch-2

:: Create 'my-file.mrg' on branch 1
git checkout -b demo-branch-1
echo "change a.xlsx on demo-branch-1" 
python.exe test/example.py 11111111
git add test/a.xlsx
git commit -m"demo-branch-1: change a.xlsx"

:: Create 'my-file.mrg' on branch 2
git checkout master
git checkout -b demo-branch-2
echo "change a.xlsx on demo-branch-1" 
python.exe test/example.py 22222222
git add test/a.xlsx
git commit -m"demo-branch-2: change a.xlsx"

:: Merge the two branches, causing a conflict
git merge -m"Merged in demo-branch-1" demo-branch-1