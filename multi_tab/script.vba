 

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



Application.EnableEvents = False
Application.DisplayAlerts = False
Application.ScreenUpdating = False
Application.AskToUpdateLinks = False

 

 

'Setting Workbook
Set wbO = ThisWorkbook
Set wbD = Workbooks.Open("filepath", False)



'Setting Worksheet
Set wsO = wbO.Sheets("Witts End Output")


'Resize the array to the number of worksheets
ReDim sheetNames(1 To wbD.Worksheets.Count)

    ' Loop through each worksheet in the active workbook
    For Each wsD In wbD.Worksheets
        i = i + 1
        sheetNames(i) = wsD.Name

 



'Start of MyAmounts
myrow = 3
mycol = 2

 

 

'Output sheet starting row
outputrow = wsO.Range("A1048576").End(xlUp).Row
outputrow = outputrow + 1
 
Set wsD = wbD.Sheets(wsD.Name)

Dim lrow As Long
Dim lcol As Long

'lrow = wsD.Range("A1").End(xlDown).Row
lrow = wsD.Range("A1048576").End(xlUp).Row
lcol = wsD.Range("XFD3").End(xlToLeft).Column

 

 


'Where to stop
Do Until myrow = lrow
    Do Until mycol = lcol

 


'Where to stop
'Do Until myrow = 113
'    Do Until mycol = 14

 

    'Name splitter
    MyName = wsD.Name
    SplitName = Split(MyName, " ")
    MySurname = SplitName(0)
    MyForename = SplitName(1)



    'MyAmount, not to include zeros
    MyAmount = wsD.Cells(myrow, mycol)

 

 

        If MyAmount <> 0 And IsNumeric(MyAmount) Then

        Debug.Print My

 

'        Rest of the Data
        MyForename = SplitName(0)
        MySurname = SplitName(1)
'        MyName = wsD.Cells(myrow, 1)
        MyDescription = wsD.Cells(3, mycol)
        MyMonth = wsD.Cells(myrow, 1)
        MyTrackerID = wsD.Cells(myrow, 1)
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
    mycol = 5

 

 

    myrow = myrow + 1

 

 


Loop

 
 


    Next wsD
    For i = 1 To UBound(sheetNames)
    Next i
 
 
Application.EnableEvents = True
Application.DisplayAlerts = True
Application.ScreenUpdating = True
Application.AskToUpdateLinks = True


wbD.Close False


 

End Sub
