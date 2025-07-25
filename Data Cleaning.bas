Attribute VB_Name = "Module1"
'Replace demoXXXX with License
Sub ReplaceDemoCodes()
    Dim cell As Range
    Dim regex As Object
    Dim ws As Worksheet

    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.Pattern = "demo\d{4}"

    Set ws = ActiveSheet
    For Each cell In ws.UsedRange
        If Not IsEmpty(cell.Value) And VarType(cell.Value) = vbString Then
            If regex.test(cell.Value) Then
                cell.Value = regex.Replace(cell.Value, "license")
            End If
        End If
    Next cell
End Sub
