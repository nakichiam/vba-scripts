Attribute VB_Name = "Module1"
'Replace demoXXXX with License
Sub ReplaceDemoCodes()
    Dim cell As Range
    Dim regex As Object
    Dim ws As Worksheet

    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.Pattern = "Key\d{4}"

    Set ws = ActiveSheet
    For Each cell In ws.UsedRange
        If Not IsEmpty(cell.Value) And VarType(cell.Value) = vbString Then
            If regex.test(cell.Value) Then
                cell.Value = regex.Replace(cell.Value, "license")
            End If
        End If
    Next cell
End Sub

'Remove everything with redacted
Sub ClearRowsWithRedactedInWord()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim r As Long, c As Long
    Dim cellVal As String

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For r = lastRow To 2 Step -1 ' Assuming row 1 is headers
        For c = 1 To lastCol
            cellVal = ws.Cells(r, c).Value
            If InStr(cellVal, "[email redacted]") > 1 Then ' in the middle of a word
                ws.Rows(r).ClearContents
                Exit For
            End If
        Next c
    Next r
End Sub

