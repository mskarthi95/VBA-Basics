VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public Sub test()
Dim R As Range

Worksheets("Sample-spreadsheet-file").Activate

Set R = Worksheets("Sample-spreadsheet-file").Range("A1:J" & Cells(Rows.Count, 1).End(xlUp).Row)

R.Select
R.Copy

Worksheets.Add
ActiveSheet.Paste

End Sub

Private Sub Workbook_NewSheet(ByVal Sh As Object)
MsgBox "New WorkSheet is Added"

End Sub


