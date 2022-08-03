# Everyday Task Program
這是一個每次喚醒電腦會"提醒你今天的行程"的程式
每日事項記在 excel 資料夾內的 Calender.xlsx 檔案中

## 安裝第三方模組 openpyxl
```pip install openpyxl```\
需先安裝 openpyxl 才可執行 everyday.py 和 create.py

## openpyxl 使用說明
**[openpyxl - Python library directions](https://openpyxl.readthedocs.io/en/stable/)**

## 示意圖
![image](https://user-images.githubusercontent.com/99878799/182423356-4aba58e8-b1ec-4a25-b449-65eafb5f2c11.png)

## 若要設定喚醒自動執行

### 打包Python程式檔
安裝第三方模組 pyinstaller
```pip intall pyinstaller```
在專案資料夾中開啟cmd
```pyinstaller –F 專案名稱.py```
即可在dist資料夾中得到*.exe檔

### 設定電腦自動執行


## 練習內容
- 用 Python 操作 excel (openpyxl)
- 使用 class物件
- REGEX 正規表達式
