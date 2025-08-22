Sub importTkIDs()
Dim objFSO As Object
Dim objFolder As Object
Dim objSubFolder As Object
Dim wbD As Workbook
Dim wsD As Worksheet
Dim i As Integer

Set wbD = ThisWorkbook
Set wsD = wbD.Sheets("Sheet1")

'Create an instance of the FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")
'Get the folder object
Set objFolder = objFSO.GetFolder("filepath")
i = 1
'loops through each file in the directory and prints their names and path
For Each objSubFolder In objFolder.subfolders
Dim Splitname As Variant
Dim SplitLenght As Integer

    'Split objFolderName
    Splitname = Split(objSubFolder.Name, ",")
    SplitLenght = (UBound(Splitname) - LBound(Splitname) + 1)
    
    
    'print folder name
   If SplitLenght > 2 Then
    wsD.Cells(i + 1, 1) = Splitname(0)
   
    wsD.Cells(i + 1, 2) = Splitname(1)
   
    wsD.Cells(i + 1, 3) = Splitname(2)
End If

 If SplitLenght = 2 Then
    wsD.Cells(i + 1, 1) = Splitname(0)
   
    wsD.Cells(i + 1, 2) = Splitname(1)
   
   End If
 If SplitLenght < 2 Then
    wsD.Cells(i + 1, 1) = Splitname(0)
   
 End If

     
      i = i + 1
Next objSubFolder
End Sub
