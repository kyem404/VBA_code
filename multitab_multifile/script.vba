
Sub import1()


'Dir
Dim wbO As Workbook
Dim wbD As Workbook
Dim wsO As Worksheet
Dim wsD As Worksheet
Dim i As Integer
Dim sheetNames() As Variant
Dim myrow As Long
Dim mycol As Long
Dim outputrow As Long
Dim MyAmount As Variant
Dim fileName1234 As Variant
 

fileName1234 = Dir("filepath/*.xlsx")
save_path = "filepath"

While fileName1234 <> ""


'Setting Workbook
Set wbO = ThisWorkbook
Set wbD = Workbooks.Open(save_path & fileName1234, False)


'Setting Worksheet
Set wsO = wbO.Sheets("Sheet1")


'Resize the array to the number of worksheets
ReDim sheetNames(1 To wbD.Worksheets.Count)

    ' Loop through each worksheet in the active workbook
    For Each wsD In wbD.Worksheets
        i = i + 1
'        sheetNames(i) = wsD.Name
 
'Start of MyAmounts
myrow = 11
mycol = 3


'Output sheet starting row
outputrow = wsO.Range("A1048576").End(xlUp).Row
outputrow = outputrow + 1
Set wsD = wbD.Sheets(wsD.Name)
 
'Where to stop
Do Until myrow = 70
    Do Until mycol = 18


    'Name splitter
'    MyName = wsD.Cells(myrow, 1)
'    SplitName = Split(MyName, " ")
'    MySurname = SplitName(0)
'    MyForename = SplitName(1)


    'MyAmount, not to include zeros
    MyAmount = wsD.Cells(myrow, mycol)

        If MyAmount <> 0 And IsNumeric(MyAmount) Then

'        Rest of the Data
        MyForename = wsD.Cells(4, 2)
        MySurname = wsD.Cells(6, 2)
'        MyName = wsD.Cells(myrow, 1)
        MyDescription = wsD.Cells(myrow, 2)
        MyMonth = wsD.Cells(8, mycol)
'        MyTrackerID = wsD.Cells(myrow, 2)
'        MyCurrency = wsD.Cells(myrow, 4)
 
        'Outputting data into Output tab
'        wsO.Cells(outputrow, 1) = Trim(MyName)
         wsO.Cells(outputrow, 1) = Trim(MySurname)
         wsO.Cells(outputrow, 2) = Trim(MyForename)
         wsO.Cells(outputrow, 3) = Trim(MyMonth)
         wsO.Cells(outputrow, 4) = Trim(MyDescription)
         wsO.Cells(outputrow, 5) = Trim(MyAmount)
'        wsO.Cells(outputrow, 6) = Trim(MyTrackerID)
'        wsO.Cells(outputrow, 7) = Trim(MyCurrency)



         outputrow = outputrow + 1

         End If

    mycol = mycol + 1

    Loop

    'Always the same as the top
    mycol = 3

    myrow = myrow + 1

Loop

    'Sheet
    Next wsD
    For i = 1 To UBound(sheetNames)
    Next i
    'Workbook
    fileName1234 = Dir

    wbD.Close False
 

Wend
 
End Sub
