Attribute VB_Name = "Module1"
'Replace demoXXXX with License
Sub ReplaceDemoCodes()
    Dim cell As Range
    Dim regex As Object
    Dim ws As Worksheet

    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    ' switches between demo and key
    regex.IgnoreCase = True ' ?? Makes it case-insensitive
    ' regex.Pattern = "demo\d{4}"
    regex.Pattern = "key\d{4}"
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

Sub RemoveRedactedWords()
    Dim cell As Range
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    
    re.Pattern = "\S*email\s*redacted\S*"
    re.Global = True
    re.IgnoreCase = True
    
    For Each cell In Selection
        If Not IsEmpty(cell.Value) Then
            cell.Value = re.Replace(cell.Value, "")
        End If
    Next cell
End Sub

Option Compare Text

Option Compare Text

Sub RemoveGreetingLines()
    Dim cell As Range
    Dim lines As Variant
    Dim result As String
    Dim i As Integer
    Dim skip As Boolean

    For Each cell In Selection
        lines = Split(cell.Value, vbCrLf)
        result = ""
        For i = LBound(lines) To UBound(lines)
            skip = False
            If lines(i) Like "Hello*" _
            Or lines(i) Like "Hi*" _
            Or lines(i) Like "Good morning*" _
            Or lines(i) Like "Good afternoon*" _
            Or lines(i) Like "Best regards*" _
            Or lines(i) Like "Kind regards*" _
            Or lines(i) Like "Dear*" Then
                skip = True
            End If
            If Not skip Then result = result & lines(i) & vbCrLf
        Next i
        cell.Value = Trim(result)
    Next cell
End Sub


