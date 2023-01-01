import win32com.client as win32
import os
import PyQt5
from pathlib import Path
import re
import math

import win32com.client as win32
import os
import PyQt5
from pathlib import Path
import re
import math


hwp=win32.gencache.EnsureDispatch("hwpframe.hwpobject")
hwp.XHwpWindows.Item(0).Visible=True
hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule")
path=os.getcwd()
hwp.Open(path+"\\사진.hwp")
list_picture=[]
path=Path(os.getcwd())


for file in path.glob('*.jpg'):
    list_picture.append(file)

for file in path.glob('*.png'):
    list_picture.append(file)


print(list_picture)

number=math.ceil(len(list_picture)/3)

print(number,"@@")

hwp.HAction.GetDefault("PageSetup", hwp.HParameterSet.HSecDef.HSet)  # 액션생성
hwp.HParameterSet.HSecDef.HSet.SetItem("ApplyClass", 24)  # 적용범위 구분. 없어도 됨
hwp.HParameterSet.HSecDef.HSet.SetItem("ApplyTo", 3)  # 적용범위. 필수
hwp.HAction.Execute("PageSetup", hwp.HParameterSet.HSecDef.HSet)  # 해당액션 실행(파라미터셋 적용)
hwp.HParameterSet.HSecDef.PageDef.LeftMargin = hwp.MiliToHwpUnit(30.0)  # 파라미터셋 설정
hwp.HParameterSet.HSecDef.PageDef.RightMargin = hwp.MiliToHwpUnit(30.0)

table = hwp.HAction.GetDefault("TableCreate", hwp.HParameterSet.HTableCreation.HSet)
hwp.HParameterSet.HTableCreation.Rows = 2
hwp.HParameterSet.HTableCreation.Cols = 3
hwp.HParameterSet.HTableCreation.WidthType = 2
hwp.HParameterSet.HTableCreation.HeightType = 1
hwp.HParameterSet.HTableCreation.WidthValue = hwp.MiliToHwpUnit(148.0)
hwp.HParameterSet.HTableCreation.HeightValue = hwp.MiliToHwpUnit(150)
hwp.HParameterSet.HTableCreation.CreateItemArray("ColWidth", 3)
hwp.HParameterSet.HTableCreation.ColWidth.SetItem(0, hwp.MiliToHwpUnit(47.0))

hwp.HParameterSet.HTableCreation.ColWidth.SetItem(1, hwp.MiliToHwpUnit(44.0))

hwp.HParameterSet.HTableCreation.ColWidth.SetItem(2, hwp.MiliToHwpUnit(47.0))
hwp.HParameterSet.HTableCreation.CreateItemArray("RowHeight", 5)
hwp.HParameterSet.HTableCreation.RowHeight.SetItem(0, hwp.MiliToHwpUnit(40.0))
hwp.HParameterSet.HTableCreation.RowHeight.SetItem(1, hwp.MiliToHwpUnit(5.0))

hwp.HParameterSet.HTableCreation.TableProperties.TreatAsChar = 1  # 글자처럼 취급
hwp.HParameterSet.HTableCreation.TableProperties.Width = hwp.MiliToHwpUnit(148)
hwp.HAction.Execute("TableCreate", hwp.HParameterSet.HTableCreation.HSet)
hwp.Run("CopyPage")

for i in range(3):
    hwp.InsertPicture(list_picture[i].absolute(),sizeoption=3)
    hwp.HAction.Run("MoveRight")
    p=i

print(p)


print(len(list_picture))
if number>=2:
    for i in range(number):
        r = 3 + i
        print(r)

        if(r<len(list_picture) or r==len(list_picture)):
            p=0
            hwp.Run("PastePage")
            hwp.Run("MoveUp")
            hwp.Run("MoveUp")
            p=r
            if (p >= len(list_picture)):
                break
            hwp.InsertPicture(list_picture[r].absolute(),sizeoption=3)
            hwp.HAction.Run("MoveRight")
            p=r+1
            if(p>=len(list_picture)):
                break
            hwp.InsertPicture(list_picture[r+1].absolute(),sizeoption=3)
            hwp.HAction.Run("MoveRight")
            if (p >= len(list_picture)):
                break
            p=r+2
            if (p >= len(list_picture)):
                break
            hwp.InsertPicture(list_picture[r+2].absolute(),sizeoption=3)

        else:
            break
