# Kai‘s py project
 Kai's Python project, reading and writing Excel file by openpyxl plugin
 

# 運行環境：
如果要運行此程序，必須安裝Python的openpyxl插件。
1. 最好安裝最新版Python3。
2. 然後在終端窗口輸入下列命令：  
   pip3 install openpyxl  

# 運行前提： 
1. 使程序順利運行， 需要首先在“Kai_Excel.xlsx"文件中選擇正確的月份，在表格中“F1”點擊箭頭標籤來改變當前月份：  
![alt text](https://github.com/HuGuo-UNC-Chapel-Hill/Kai-s-py-project/blob/main/Screenshots/Screen%20Shot%202022-07-14%20at%209.56.39%20AM.png?raw=true).   

   改變後的效果，每月的週日天數會隨月份改變而自動改變  
   ![alt text](https://github.com/HuGuo-UNC-Chapel-Hill/Kai-s-py-project/blob/main/Screenshots/Screen%20Shot%202022-07-14%20at%209.56.48%20AM.png?raw=true)

2. 選擇好正確月份，可用下來菜單來選擇所有人員的優先安排事項：  
![alt text](https://github.com/HuGuo-UNC-Chapel-Hill/Kai-s-py-project/blob/main/Screenshots/Screen%20Shot%202022-07-14%20at%209.59.56%20AM.png?raw=true)  
   要清理上月的優先安排，只需要框選整個表格區域，按鍵盤上的“delete”按鍵。  
   整個區域清空後，開始用下來菜單安排新月份的優先事項即可。  
![alt text](https://github.com/HuGuo-UNC-Chapel-Hill/Kai-s-py-project/blob/main/Screenshots/Screen%20Shot%202022-07-14%20at%2010.01.18%20AM.png?raw=true).    


3. 選者玩當月所有人員的優先安排後，表格會自動顯示本月每一個週日可用的人員列表。  
   成勳會自動讀取有效日期下的人員列表。  
![alt text](https://github.com/HuGuo-UNC-Chapel-Hill/Kai-s-py-project/blob/main/Screenshots/Screen%20Shot%202022-07-14%20at%2010.00.36%20AM.png?raw=true)

4. 確認當前月份與所有人員優先安排無誤後請保存當前文檔。  
![alt text](/Screenshots/Screen%20Shot%202022-07-14%20at%2010.24.21%20AM.png?raw=true).   
請一定要保存當前文檔！  
請一定要保存當前文檔！  
請一定要保存當前文檔！  

# 運行程序 
保存好文檔之後就可以點擊”kai_read.py"運行。
運行“Kai_read.py"程序，將會在當前目錄下自動生成單（當）月排班表的Excel文件.  
 ![alt text](/Screenshots/Screen%20Shot%202022-07-14%20at%2010.03.39%20AM.png?raw=true)  
 
打開後可以查閱當月安排詳情.  
 ![alt text](/Screenshots/Screen%20Shot%202022-07-14%20at%2010.03.19%20AM.png?raw=true)  
