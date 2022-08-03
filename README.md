# Everyday Task Program
這是一個每次喚醒電腦會"提醒你今天的行程"的程式
每日事項記在 excel 資料夾內的 Calender.xlsx 檔案中

## 安裝第三方模組 **openpyxl**
```pip install openpyxl```\
需先安裝 openpyxl 才可執行 everyday.py 和 create.py

## openpyxl 使用說明
**[openpyxl - Python library directions](https://openpyxl.readthedocs.io/en/stable/)**

## 示意圖
![image](https://user-images.githubusercontent.com/99878799/182423356-4aba58e8-b1ec-4a25-b449-65eafb5f2c11.png)

# 若要設定喚醒自動執行

## 打包Python程式檔
安裝第三方模組 **pyinstaller**\
```pip intall pyinstaller```\
\
在專案資料夾中開啟cmd\
```pyinstaller –F 專案名稱.py```\
\
即可在dist資料夾中得到*.exe檔

## 設定電腦自動執行
點選左方"**工作排程器程式庫**"\
![image](https://user-images.githubusercontent.com/99878799/182544736-1cad9536-b39c-4528-9c81-e887a1d4c83f.png)

點選右方"**新增資料夾**"\
![image](https://user-images.githubusercontent.com/99878799/182544913-cdd22636-34e6-4d32-a448-f3e77ab7ddf1.png)

在建立好的資料夾中點選"**建立新工作**"\
![image](https://user-images.githubusercontent.com/99878799/182544521-357c874c-f752-4c9d-b1af-8da41ab339b4.png)\
![image](https://user-images.githubusercontent.com/99878799/182544535-5fb64edd-5d3f-4dd3-b794-d45e34f6e17b.png)\
以我為例，我clone在桌面\
![image](https://user-images.githubusercontent.com/99878799/182544538-b596c735-d160-4e86-9e1d-7761009f28ae.png)\
![image](https://user-images.githubusercontent.com/99878799/182544540-fc11c6de-74f6-4f4c-8628-3a5cef613d6f.png)\
\
**設置完成**\
電腦會在開機、喚醒休眠時自動播報提醒今天的排程

## 練習內容
- 用 Python 操作 excel (openpyxl)
- 使用 class物件
- REGEX 正規表達式
