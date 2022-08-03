# excel_schedule_clean
Takes a pdf-to-xls conversion from Convertio.co and cleans for use in MRP.

Goal: output delivery schedule in clean format to excel for production plan scheduling purposes.  Easier to test production plan against delivery dates, manual adjustment to ERP system is slow and cumbersome in comparison.

Process:
1. Print PDF copy of delivery schedule from ERP/MRP software.

2. Convert it to xls using Convertio.co

3. Manually convert from rows to columns using space delimiters.

4.Run schedule_clean() macro.

Output = Clean table of data 

Customer Order :: Part Number :: QTY :: Dock Date
