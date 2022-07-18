# Kai‘s Python Project
 Kai's Python project, reading, computing, and writing Excel file by openpyxl plugin
 
# 安排規則
  原則上每月每人擔班不超過兩次。如果某人本月實際擔班兩次，那麼下月將只擔班一次。  
  結果： 如果某人連續每月擔班，那麼平均每月擔班次數為1.5次。  

# 綜合報告
  在程序生成的Excel文件中添加了每月綜合擔班報告， 顯示擔班的成員本月所有擔班次數與種類。


# 運行環境：
如果要運行此程序，必須安裝Python的openpyxl插件。
1. 最好安裝最新版Python3。
2. 然後在終端窗口輸入下列命令：  
   pip3 install openpyxl  

# 運行前提： 
1. 下載本項目的zip，解壓後進入目錄“Kai-s-py-project”，  
   ![alt text](https://github.com/HuGuo-UNC-Chapel-Hill/Kai-s-py-project/blob/main/Screenshots/Screen%20Shot%202022-07-18%20at%2011.17.42%20AM.png?raw=true)   
   
   再進入目錄“kai_project_0.2”：  
   ![alt text](https://github.com/HuGuo-UNC-Chapel-Hill/Kai-s-py-project/blob/main/Screenshots/Screen%20Shot%202022-07-18%20at%2011.18.03%20AM.png?raw=true)    
   
   首先使用Microsoft offic軟件打開“kai_Excel_2_Marco.xlsm"進行編輯，  
   ![alt text](https://github.com/HuGuo-UNC-Chapel-Hill/Kai-s-py-project/blob/main/Screenshots/Screen%20Shot%202022-07-18%20at%202.56.59%20PM.png?raw=true)     
   
   “kai_Excel_2_Marco.xlsm"中添加了一鍵清除當前優先安排的功能。 如果要使用這項功能，必須在打開文件的時候允許“Enable Macros." 否則只能手動使用“delete”鍵清除。
   ![alt text](https://github.com/HuGuo-UNC-Chapel-Hill/Kai-s-py-project/blob/main/Screenshots/Screen%20Shot%202022-07-18%20at%2011.17.06%20AM.png?raw=true)
   ![alt text](https://github.com/HuGuo-UNC-Chapel-Hill/Kai-s-py-project/blob/main/Screenshots/Screen%20Shot%202022-07-17%20at%205.30.32%20PM.png?raw=true)   
   
   使程序順利運行， 需要首先在“kai_Excel_2_Marco.xlsm"文件中選擇正確的月份，在表格中“F1”點擊箭頭標籤來改變當前月份：  
   ![alt text](https://github.com/HuGuo-UNC-Chapel-Hill/Kai-s-py-project/blob/main/Screenshots/Screen%20Shot%202022-07-18%20at%202.56.59%20PM.png?raw=true)    
![alt text](https://github.com/HuGuo-UNC-Chapel-Hill/Kai-s-py-project/blob/main/Screenshots/Screen%20Shot%202022-07-14%20at%209.56.39%20AM.png?raw=true)   

   改變後的效果，每月的週日天數會隨月份改變而自動改變  
   ![alt text](https://github.com/HuGuo-UNC-Chapel-Hill/Kai-s-py-project/blob/main/Screenshots/Screen%20Shot%202022-07-14%20at%209.56.48%20AM.png?raw=true)

2. 選擇好正確月份，可用下拉菜單來選擇所有人員的優先安排事項：  
![alt text](https://github.com/HuGuo-UNC-Chapel-Hill/Kai-s-py-project/blob/main/Screenshots/Screen%20Shot%202022-07-14%20at%209.59.56%20AM.png?raw=true)  

   要清理上月的優先安排，只需要框選整個表格區域，按鍵盤上的“delete”按鍵。  
   整個區域清空後，開始用下拉菜單安排新月份的優先事項即可。  
![alt text](https://github.com/HuGuo-UNC-Chapel-Hill/Kai-s-py-project/blob/main/Screenshots/Screen%20Shot%202022-07-14%20at%2010.01.18%20AM.png?raw=true).    


3. 確定當月所有人員的優先安排後，表格會自動顯示本月每一個週日可用的人員列表。  
   “kai_ready.py”程序運行時會自動讀取有效日期下的人員列表。  
![alt text](https://github.com/HuGuo-UNC-Chapel-Hill/Kai-s-py-project/blob/main/Screenshots/Screen%20Shot%202022-07-14%20at%2010.00.36%20AM.png?raw=true)

4. 確認當前月份的優先安排後，請保存當前文檔。  
![alt text](https://github.com/HuGuo-UNC-Chapel-Hill/Kai-s-py-project/blob/main/Screenshots/Screen%20Shot%202022-07-14%20at%2010.24.21%20AM.png?raw=true)    
請一定要保存當前文檔！  
請一定要保存當前文檔！  
請一定要保存當前文檔！  

#  運行程序 
   運行方法有兩種：  
1. 使用Pycharm一類的IDE軟件直接運行：  
   ![alt text](https://github.com/HuGuo-UNC-Chapel-Hill/Kai-s-py-project/blob/main/Screenshots/Screen%20Shot%202022-07-15%20at%209.49.00%20AM.png?raw=true)
2. 在終端窗口下進入“kai_project_0.1”目錄，然後輸入：  
   python3 kai_ready.py
  
   程序會開開始運行，並且在終端窗口中顯示一些提示信息：  
   ![alt twxt](https://github.com/HuGuo-UNC-Chapel-Hill/Kai-s-py-project/blob/main/Screenshots/Screen%20Shot%202022-07-14%20at%2012.09.42%20PM.png?raw=true)   
  
#  運行結果  
   “Kai_read.py"運行完畢後，將會在當前目錄下自動生成單（當）月排班表的Excel文件.  
   ![alt text](https://github.com/HuGuo-UNC-Chapel-Hill/Kai-s-py-project/blob/main/Screenshots/Screen%20Shot%202022-07-14%20at%2011.08.26%20AM.png?raw=true)   
  
   打開後可以查閱當月安排詳情.  
   ![alt text](https://user-images.githubusercontent.com/86079744/179245280-948da2af-7ef3-45f3-9b11-5503923baa7f.png)   
   請保存上月的排班表，以便進行本月份安排時可以優化本月值班次數。  
