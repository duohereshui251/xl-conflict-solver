# ExlMerge

## 使用环境

windows10

目前只支持微软Excel文件后缀为xlsx的文件，其他的没测试过

## 如何使用

1. 下载 `ExlDiff.exe` 和 `ExlMerge.exe`两个程序，放在任何你想放的地方。

2. 配置`.git/config`文件，末尾添加上

```
[diff "excel"]
    command = <文件路径>/ExlDiff.exe 
[merge "ExcelMerge"]
    name = A custom merge driver used to resolve conflicts in excel files
	driver  = <文件路径>/ExlMerge.exe %O %A %B %P

```



3. 在根目录下新增文件 `.gitattributes`，内容如下：

   ```
   *.xls* diff=excel
   *.xls* merge=ExcelMerge
   ```



## 功能说明

**excel对比**

例如刚刚修改一个excel文件，但是还未提交，使用

```
git diff 文件名
```

就能来查看差异，效果如下

```
(venv) PS D:\code\xl-conflict-solver\xl-conflict-solver> git diff .\test\a.xlsx
in sheet Sheet1
+++ a/a.xlsx/D1
--- b/nRGmNe_a.xlsx/D1
+5555.0
-4.0
```

如果想在图形界面进行对比，请使用其他工具

**excel合并处理**

假如在合并其他分支，或者对远程分支进行拉取的时候，如果同时修改了同一文件同一单元格造成了冲突，则合并完成后会提示合并失败，需要手动修改。

