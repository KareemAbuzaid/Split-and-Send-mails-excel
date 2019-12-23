Sub split()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Split a sheet by categories into different workbooks then ' 
' move those workbooks to new files. The name of the file   '
' is the name of the category.                              '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Step 1 - Name your ranges and Copy sheet
'Step 2 - Filter by Department and delete rows not applicable
'Step 3 - Loop until the end of the list

'variable to represent the range that will be splited
Dim Splitcode As Range
Sheets("Master").Select
Set Splitcode = Range("Splitcode")

'loop over all cells in the range 
For Each cell In Splitcode
 Sheets("Master").Copy After:=Worksheets(Sheets.Count)
 
 'create a new worksheet and give it a name
 ActiveSheet.Name = cell.Value

 'move the data to the worksheet
 With ActiveWorkbook.Sheets(cell.Value).Range("MasterData")
  .AutoFilter Field:=3, Criteria1:="<>" & cell.Value, Operator:=xlFilterValues
  .Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete
 End With

 'remove the autofilter to show the data that remains
 ActiveSheet.AutoFilter.ShowAllData
Next cell

Dim WS_Count As Integer
Dim I As Integer
Dim path As String
path = Application.ActiveWorkbook.path

' Set WS_Count equal to the number of worksheets in the active
' workbook.
WS_Count = ActiveWorkbook.Worksheets.Count

'Begin the loop and start at the 3rd worksheet.
For I = 3 To WS_Count
    'Create a new workbook and save the data in the worksheet into it.
    Dim wb As Workbook
    Set wb = Workbooks.Add
    ThisWorkbook.Sheets(I).Copy Before:=wb.Sheets(1)
    'The new workbook gets the name of the worksheet and is closed after saving
    wb.SaveAs Filename:=path & "\" & ThisWorkbook.Sheets(I).Name
    wb.Close Savechanges:=True
Next I
End Sub
