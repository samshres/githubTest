'====================================================
'HP UFT/QTP - Create Text File
Set fso=CreateObject ("Scripting.FileSystemObject")
Set tfile1=fso.CreateTextFile("C:\Users\bhail\OneDrive\Desktop\Files\Trainings\VB Script\VBS-Training/demofile1", False)
MsgBox "FileCreated"

'====================================================
'HP UFT/QTP - Write Data in Text File - This may only work on the UFT EID.

Set fso=CreateObject("Scripting.FileSystemObject")
Set stream=fso.CreateObject ("C:\Users\bhail\OneDrive\Desktop\Files\Trainings\VB Script\VBS-Training/demofile1")
	stream.Write "This is the first test that I am writing in the Text Files to see if I can automate the writing in the TextFile"
	
Stream.Close

Set stream = Nothing
Set fso = Nothing

'====================================================
'HP UFT/QTP - Write data to Excel.Application
Set tExcel = CreateObject(Excel.Application)

tExcel.visible = True
tExcel.workbook.add
tExcel.sheet.add

tExcel.Cells(1,1).Value = "Monday"
tExcel.Cells(1,1).Value = "Tuesday"
tExcel.Cells(1,1).Value = "Wednesday"
tExcel.Cells(1,1).Value = "Thursday"
tExcel.Cells(1,1).Value = "Friday"
tExcel.Cells(1,1).Value = "Saturday"
tExcel.Cells(1,1).Value = "Sunday"

tExcel.ActiveWorkbook.SaveAs "C:\Users\bhail\OneDrive\Desktop\Files\Trainings\VB Script\VBS-Training\vbdemo1"
tExcel.Workbook.colse
tExcel.Application.Quit

'====================================================
'HP UFT/QTP - String Functions
Dim varString

varString = InputBox("Enter the text")
MsgBox len (varString)

MsgBox UCase(varString)
MsgBox LCase(varString)

MsgBox left ("My Name is Sanam Bhaila" ,15)
MsgBox right ("My Name is Sanam Bhaila" ,17)
MsgBox mid ("My Name is Sanam Bhaila" ,9, 14)

'====================================================
'HP UFT/QTP - Date Function 

MsgBox Date
MsgBox now()

MsgBox dateadd("m", 1, "1/14/2024")
MsgBox dateadd("d", 1, "1/14/2024")
MsgBox dateadd("ww", 1, "1/14/2024")
MsgBox dateadd("yyyy", 1, "1/14/2024")

MsgBox dateadd("h", 1, "1/14/2024 9:00:00")
MsgBox dateadd("h", 1, "1/14/2024 9:00:00")
MsgBox dateadd("h", 1, "1/14/2024 9:00:00")

MsgBox datediff("yyyy","1/14/2024, 1/15/2025")
MsgBox datediff("m","1/14/2024, 1/15/2025")
MsgBox datediff("d","1/14/2024, 1/16/2025")



