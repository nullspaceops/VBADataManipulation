Attribute VB_Name = "Modul1"
Sub DownloadFixings()
'
Dim Name As String
Dim Path As String
Dim UnderlyingName As String
Dim Quote As String
Dim StartDate As Date
Dim EndDate As Date
Dim TempDate As Date
Dim Row As Integer

Dim Count As Integer
Dim TempCount As Integer
Dim CompoundName As String
Dim File As String
Dim TempColumn As Integer
Dim TempRow2 As Integer
Dim TempRow As Integer

Application.ScreenUpdating = False

Sheets("Settings").Select
Name = Cells(1, 2).Value
Path = Cells(2, 2).Value
UnderlyingName = Cells(3, 2).Value
Quote = Cells(4, 2).Value
StartDate = Cells(5, 2).Value
EndDate = Cells(6, 2).Value
CompoundName = Name + "-" + Quote + "-" + UnderlyingName
File = Path + CompoundName

' Select the sheet to be prepared for conversion
Sheets("CSV").Select
Range("A1").Select
TempDate = StartDate

' Write the dates for our fixings and get the counter for our row number
    While TempDate <= EndDate
        ActiveCell.Value = TempDate
        TempDate = TempDate + 1
        ActiveCell.Offset(1, 0).Select
    Wend

Row = ActiveCell.Row
Range("B1").Select
    
' Write the formula in the rows (atm an arbitrary value) and recalculate the sheet, wait a second or two
    For Count = 1 To Row - 1
        ActiveCell.Value = 100
        ActiveCell.Offset(1, 0).Select

    Next

Worksheets("CSV").Calculate
Application.Wait (500)

'Do CleanUp. A spot price is usually an integer or a double, errors not. Cleans unnecessary rows (e.g. non business days)

    For Count = 1 To Row - 1
    TempRow = ActiveCell.Row
    TempColumn = ActiveCell.Column
    
    ActiveCell.Offset(1, 0).Select
        If Not TypeName(ActiveCell.Value) = "Double" Then
        TempRow2 = ActiveCell.Row
        Rows(TempRow2).Select
        Selection.Delete
        Cells(TempRow, TempColumn).Select
        End If
    
    Next
    
'Add some arbitrary name to it
    
Range("C1").Select
    
    For Count = 1 To Row - 1
    ActiveCell.Value = CompoundName
    ActiveCell.Offset(1, 0).Select
    Next

' Save the data locally in this template and delete the old file

ActiveWorkbook.Save
Application.ScreenUpdating = True
ChDir Path
    If Dir(File + ".csv") <> "" Then
    Kill File + ".csv"
    End If

ActiveWorkbook.SaveAs Filename:=Path + CompoundName + ".csv", FileFormat:=xlCSVMSDOS _
        , CreateBackup:=False
        
ActiveWorkbook.Saved = True
Application.Quit

End Sub





