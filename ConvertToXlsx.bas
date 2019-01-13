Attribute VB_Name = "Modul1"
Sub ConvertToXlsx()

Dim strFileName As String
Dim strFolder As String: strFolder =
Dim strFileSpec As String: strFileSpec = strFolder & "*.csv"
Dim Appendix As String
Dim CompundName As String

strFileName = Dir(strFileSpec)
Appendix = "-"

Do While Len(strFileName) > 0

    Application.Workbooks.Open strFolder + strFileName
    Columns("A:A").Select
    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=False, FieldInfo _
        :=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1), _
        Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12, 1), Array(13, 1 _
        ), Array(14, 1), Array(15, 1), Array(16, 1), Array(17, 1)), TrailingMinusNumbers:= _
        True
    
    ActiveWorkbook.SaveAs Filename:=strFolder + strFileName + Appendix, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWorkbook.Close True
    CompoundName = strFolder + strFileName + Appendix
    Name CompoundName As CompoundName + ".xlsx"

    strFileName = Dir
    

Loop

End Sub
