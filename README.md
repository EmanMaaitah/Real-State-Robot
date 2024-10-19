# OBJECTIVE 
Get a lot of information and fill it automatically with Excel spreadsheets.Getting to know the input methodes, best practise for naming variables.
## MANUALLY
1) Access to the open sooq Enter in the search for an apartment for rent in a particular city.
2) The information is displayed in the form of cards and filtered, and here the goal is to take the information and fill it with Excel tables in addition to getting the best results.
# TECHNICAL REQUIREMENTS
## UI Automation Activities
1) [Input Dialog](https://docs.uipath.com/activities/other/latest/workflow/input-dialog)
2) [Use Application/Browser](https://docs.uipath.com/activities/other/latest/ui-automation%22/n-application-card)
3) [Type Into](https://docs.uipath.com/activities/other/latest/ui-automation%22/type-into)
4) [Select Item](https://docs.uipath.com/activities/other/latest/ui-automation%22/n-select-item)

## Using Data Table Activities
1) [Extract Table Data](https://docs.uipath.com/activities/other/latest/ui-automation%22/n-extract-data)
2) [Filter Data Table](https://docs.uipath.com/ACTIVITIES/other/latest/workflow/filter-data-table)

## Productivity Activities
1) [Write Range Workbook](https://docs.uipath.com/activities/other/latest/productivity/write-range)

## Using Excel Activities
1) [Excel Process Scope](https://docs.uipath.com/activities/other/latest/productivity/excel-process-scope-x)
2) [Use Excel File](https://docs.uipath.com/activities/other/latest/productivity/excel-application-card)
3) [Read Range](https://docs.uipath.com/activities/other/latest/productivity/excel-read-range)
4) [Delete Sheet](https://docs.uipath.com/activities/other/latest/productivity/delete-sheet-x)
5) [Write DataTable to Excel](https://docs.uipath.com/activities/other/latest/productivity/write-range-x)
6) [Insert Column](https://docs.uipath.com/ACTIVITIES/other/latest/productivity/insert-column-x)
7) [For Each Excel Row](https://docs.uipath.com/activities/other/latest/productivity/excel-for-each-row)
8) [Write Cell](https://docs.uipath.com/activities/other/latest/productivity/write-cell-x)
   
# PROCEDURES
1) Insert the Input Dialog activity and select the input type (Text Box), save it in a string variable.
   
   ![image](https://github.com/user-attachments/assets/daf61559-9238-4f73-863e-daa55e8bb919)

2) Use Application/Browser>Type into (In order to filter some information)> Select item (For multiple choice fields/ drop down menue)
   
![image](https://github.com/user-attachments/assets/e6075d28-f3e2-4b32-9d9d-3a9a06c2ad35)

![image](https://github.com/user-attachments/assets/0e212b6b-23ec-423e-aebf-6e33d27ba284)

You can enter the enter command inside the type into box

![image](https://github.com/user-attachments/assets/54b87c9c-0502-471b-a9c5-5aff2a1c435a)


3) You can use [Comment Out activity](https://docs.uipath.com/activities/other/latest/workflow/comment-out) using (ctrl+D) To deactivate some unwanted activities.
4) Extract Table Data: In order to get all the information and arrange it in rows and columns
   ![image](https://github.com/user-attachments/assets/43f6f080-271e-4589-aafe-366f62834d7c)

Name the headers and make a select on the desired information
   ![Screenshot (852)](https://github.com/user-attachments/assets/1f8744ab-1cf0-46f8-ba4b-22314502901f)

Through the settings, you can determine the number of rows and columns and determine the quality of variables.

5) Use Fiter Date
   To filter the data and call the created variable in Extract Table Data

   ![image](https://github.com/user-attachments/assets/e87c697f-3f07-457d-b1fe-d4e6163615ea)

 6) To make sure the number of rows
    dt_Datas.RowCount.ToString in Write line box.

 7) To store filtered data in a new sheet use Write Range Workbook

 8) Excel Process Scope> Use Excel File> Delete Sheet (to get the updated sheet)>Write DataTable to Excel>Insert Column
    
    ![image](https://github.com/user-attachments/assets/8ff56765-21f4-4914-9ead-d3caef3cecb2)

 9) For Each Excel Row> Write Cell
     
     ![image](https://github.com/user-attachments/assets/71eb54e6-666a-40de-80d1-c91b208b86f4)






# Results
# CONCLUSION
1) You can speed up the process by changing the [input mode](https://docs.uipath.com/studio/standalone/2023.4/user-guide/input-methods) Through the  Properties Panel:(Simulate, Chromium API, SendWindowMessages, Hardware Events, Background)
2) The difference between ancher and target is to match the exact label.
3) You can start by naming variables with one of the phrases depending on the type of variable (str/dt/Int/bool)
   
