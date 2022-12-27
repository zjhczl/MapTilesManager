import os
import shutil
import win32com.client
# from win32com.client import Dispatch


def checkDir(dirName):
    try:
        int(dirName)
        return True
    except:
        return False


def copy_dir(dir1, dir2):  # 复制该目录到另外一目录下
    basedir = os.path.basename(dir1)
    dlist = os.listdir(dir1)
    if not os.path.exists(dir2):
        os.mkdir(dir2)

    if not os.path.exists(dir2+"/"+basedir):
        os.mkdir(dir2+"/"+basedir)

    for f in dlist:
        file1 = os.path.join(dir1, f)  # 源文件
        file2 = os.path.join(dir2+"/"+basedir, f)  # 目标文件
        if os.path.isfile(file1):
            shutil.copyfile(file1, file2)
        if os.path.isdir(file1):
            copy_dir2(file1, os.path.join(dir2+"/"+basedir, f))


def copy_dir2(dir1, dir2):  # 复制目录下的所有文件到另外一目录
    dlist = os.listdir(dir1)

    if not os.path.exists(dir2):
        os.mkdir(dir2)

    for f in dlist:
        file1 = os.path.join(dir1, f)  # 源文件
        file2 = os.path.join(dir2, f)  # 目标文件
        if os.path.isfile(file1):
            shutil.copyfile(file1, file2)
        if os.path.isdir(file1):
            copy_dir2(file1, os.path.join(dir2, f))


def copyTiles(dirPath, savePath, nameSet):
    dirs = os.listdir(dirPath)
    for dir in dirs:
        current_path = os.path.join(dirPath, dir)
        if (checkDir(dir) == False):
            if (dir in nameSet):
                copyTiles(current_path, savePath, nameSet)
        else:
            print(current_path+"==>"+savePath+"\\")
            copy_dir(current_path, savePath)


current_directory = os.path.dirname(os.path.abspath(__file__))
mapPath = os.path.join(current_directory, 'world')
path = current_directory+"/"+"list.xlsx"
print(mapPath)

# 国家
# copyTiles(mapPath, current_directory+"\\maptiles")
xlApp = win32com.client.Dispatch('Excel.Application')
# xlApp.DisplayAlerts = True
xlApp.Visible = False


xlBook = xlApp.Workbooks.Open(path)
sht = xlBook.Worksheets(1)  # 打开sheet
row = sht.UsedRange.Rows.Count
col = sht.UsedRange.Columns.Count
name = set()

for r in range(row):
    for c in range(col):
        name.add(sht.Cells(r+1, c+1).Value)

print(name)

copyTiles(mapPath, current_directory+"\\maptiles", name)
