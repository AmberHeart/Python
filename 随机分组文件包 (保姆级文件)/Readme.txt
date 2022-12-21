需要用到的模块
xlwings
PyQt5

附- pyinstall 终端输入代码（需要对应模块生成exe icon文件已经打包)
pyinstaller -F -w -i icon.ico mainapp.py -p GUI.py -p classify.py