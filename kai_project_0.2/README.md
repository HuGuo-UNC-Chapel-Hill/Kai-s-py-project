# Kai‘s Python Project（v0.2）
  Kai's Python project, reading, computing, and writing Excel file by openpyxl plugin.  
 
# V0.2002 
  Fixed January report reading probelm.  
# V0.2003  
  Add current year to the name of the output Excel file.   
 
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
   ![alt text](https://github.com/HuGuo-UNC-Chapel-Hill/Kai-s-py-project/blob/main/Screenshots/Screen%20Shot%202022-07-17%20at%205.30.32%20PM.png?raw=true)   
   
   使程序順利運行， 需要首先在“kai_Excel_2_Marco.xlsm"文件中選擇正確的月份，在表格中“F1”點擊箭頭標籤來改變當前月份：     
![alt text](https://github.com/HuGuo-UNC-Chapel-Hill/Kai-s-py-project/blob/main/Screenshots/Screen%20Shot%202022-07-18%20at%203.07.50%20PM.png?raw=true)   

   改變後的效果，每月的週日天數會隨月份改變而自動改變  
   ![alt text](https://github.com/HuGuo-UNC-Chapel-Hill/Kai-s-py-project/blob/main/Screenshots/Screen%20Shot%202022-07-18%20at%203.25.25%20PM.png?raw=true)

2. 要清理上月的優先安排，點擊“清理當前所有優先安排。“  
   ![alt text](https://github.com/HuGuo-UNC-Chapel-Hill/Kai-s-py-project/blob/main/Screenshots/Screen%20Shot%202022-07-18%20at%203.47.18%20PM.png?raw=true).   
   整個區域清空後，開始用下拉菜單安排新月份的優先事項即可。  
   ![alt text](https://github.com/HuGuo-UNC-Chapel-Hill/Kai-s-py-project/blob/main/Screenshots/Screen%20Shot%202022-07-18%20at%203.27.39%20PM.png?raw=true).    


3. 確定當月所有人員的優先安排後，表格會自動顯示本月每一個週日可用的人員列表。  
   “kai_ready.py”程序運行時會自動讀取有效日期下的人員列表。  
   ![alt text](https://github.com/HuGuo-UNC-Chapel-Hill/Kai-s-py-project/blob/main/Screenshots/Screen%20Shot%202022-07-18%20at%203.10.17%20PM.png?raw=true)

4. 確認當前月份的優先安排後，請保存當前文檔。  
   ![alt text](https://github.com/HuGuo-UNC-Chapel-Hill/Kai-s-py-project/blob/main/Screenshots/Screen%20Shot%202022-07-14%20at%2010.24.21%20AM.png?raw=true)    
請一定要保存當前文檔！  
請一定要保存當前文檔！  
請一定要保存當前文檔！  

#  運行程序 
   運行方法有兩種：  
1. 使用Pycharm一類的IDE軟件直接運行：  
   ![alt text](https://github.com/HuGuo-UNC-Chapel-Hill/Kai-s-py-project/blob/main/Screenshots/Screen%20Shot%202022-07-18%20at%203.33.27%20PM.png?raw=true)
   
2. 在終端窗口下進入“kai_project_0.2”目錄，然後輸入：  
   python3 kai_ready_02.py
  
   程序會開開始運行，並且在終端窗口中顯示一些提示信息：  
   ![alt twxt](https://github.com/HuGuo-UNC-Chapel-Hill/Kai-s-py-project/blob/main/Screenshots/Screen%20Shot%202022-07-18%20at%203.37.54%20PM.png?raw=true)   
  
#  運行結果  
   “kai_read_02.py"運行完畢後，將會在當前目錄下自動生成單（當）月排班表的Excel文件.  
   ![alt text](https://github.com/HuGuo-UNC-Chapel-Hill/Kai-s-py-project/blob/main/Screenshots/Screen%20Shot%202022-07-18%20at%203.39.12%20PM.png?raw=true)   
  
   打開後可以查閱當月安排詳情以及本月綜合擔班的詳情與次數。  
   ![alt text](https://github.com/HuGuo-UNC-Chapel-Hill/Kai-s-py-project/blob/main/Screenshots/Screen%20Shot%202022-07-18%20at%204.22.14%20PM.png?raw=true)   
   請保存上月的排班表，以便進行本月份安排時可以優化本月值班次數。  
