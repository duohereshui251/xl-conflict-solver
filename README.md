## 关于开发与运行
首先电脑上要安装python3，使用pip安装好virtualenv，然后在当前目录下新建venv环境并且激活，windows使用：
```
pip3 install virtualenv
virtualenv  venv
.\venv\Scripts\activate.ps1
```
激活命令不同系统不一样。
进入虚拟环境后，安装好依赖
```
pip -r requirements.txt
```
然后运行test文件夹中的脚本

## 参考项目
```
https://github.com/xlwings/git-xl
```

## 思路
1. 编写好python代码实现基本的对比功能，可以的话想要实现以下功能：
    - 每个单元格文字内容改动检测
    - excel样式改动检测（颜色，字体粗细，大小等）
    - 宏改动检测（VBA）
    - 用户自定义函数改动检测
    - 图表改动检测
    - 图片改动检测

2. 将编写好的代码使用工具生成exe格式的可执行文件
3. 将生成的程序添加到git diff, git merge等组件中中，可以做到git diff对比xlsx格式文件会调用到我们的程序输出差异，可以参考相关教程：

4. 有能力的话，实现web界面查看每一次commit的改动，可以参考 `https://www.xltrail.com/` 网站的效果（但我估计八成做不出来）

