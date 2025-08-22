Sub import()

 

'Dir
Dim wbO As Workbook
Dim wsO As Worksheet
Dim wsD As Worksheet

 

'Setting Workbook
Set wbO = ThisWorkbook

 

'Setting Worksheet
Set wsO = wbO.Sheets("Output")
Set wsD = wbO.Sheets("Data")

 

'Start of MyAmounts
myrow = 2
mycol = 7

 

'Output sheet starting row
outputrow = 2

 

'Where to stop
Do Until myrow = 17
    Do Until mycol = 113
    
    'Name splitter
    MyName = wsD.Cells(myrow, 1)
    SplitName = Split(MyName, ", ")
    MySurname = SplitName(0)
    MyForename = SplitName(1)


    'MyAmount, not to include zeros
    MyAmount = wsD.Cells(myrow, mycol)

        If MyAmount <> 0 And IsNumeric(MyAmount) Then

        'Rest of the Data
        'MyForename = wsD.Cells(myrow, 4)
        'MySurname = wsD.Cells(myrow, 5)
        'MyName = wsD.Cells(myrow, 3)
        MyDescription = wsD.Cells(myrow, 6)
        MyMonth = wsD.Cells(1, mycol)
        MyTrackerID = wsD.Cells(myrow, 2)
        MyCurrency = wsD.Cells(myrow, 4)

        'Outputting data into Output tab
        wsO.Cells(outputrow, 1) = Trim(MySurname)
        wsO.Cells(outputrow, 2) = Trim(MyForename)
        wsO.Cells(outputrow, 3) = Trim(MyMonth)
        wsO.Cells(outputrow, 4) = Trim(MyDescription)
        wsO.Cells(outputrow, 5) = Trim(MyAmount)
        wsO.Cells(outputrow, 6) = Trim(MyTrackerID)
        wsO.Cells(outputrow, 7) = Trim(MyCurrency)

        outputrow = outputrow + 1

        End If

    mycol = mycol + 1

 

    Loop

    'Always the same as the top
    mycol = 7

    myrow = myrow + 1


Loop

 

End Sub
