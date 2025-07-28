Attribute VB_Name = "Module1"
Option Compare Text
' is a directive you can place at the top of a VBA module to make string comparisons case-insensitive within that module.
' "apple" = "Apple"    ' ? False (by default)
' With Option Compare Test:  "apple" = "Apple"    ' ? True

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

Sub AnonymizeNamesAndEmails()
    Dim cell As Range
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.IgnoreCase = True

    For Each cell In Selection
        If Not IsEmpty(cell.Value) Then
            Dim txt As String
            txt = cell.Value

            ' Replace names in angle brackets (emails): <...>
            re.Pattern = "<[^>]+>"
            txt = re.Replace(txt, "")

            ' Replace full names: e.g., Kamensky Pavol, Benjamin Haaske ? User
            re.Pattern = "\b[A-Z][a-z]+(?:\s+[A-Z][a-z]+)+\b"
            txt = re.Replace(txt, "User")

            ' Remove "Mr./Ms./Dr." titles followed by names
            re.Pattern = "\b(Mr|Ms|Mrs|Dr)\.?\s+[A-Z][a-z]+\b"
            txt = re.Replace(txt, "User")

            cell.Value = Trim(txt)
        End If
    Next cell

    MsgBox "Names and emails replaced with 'User'."
End Sub

Sub MasterAnonymizeSupportLog()
    Dim cell As Range
    Dim re As Object
    Dim matches As Object
    Dim match As Variant
    Dim line As Variant
    Dim cleaned As String
    Dim i As Integer
    Dim skip As Boolean

    Set re = CreateObject("VBScript.RegExp")
    
    For Each cell In Selection
        If Not IsEmpty(cell.Value) Then
            cleaned = ""
            For Each line In Split(cell.Value, vbCrLf)
                skip = False
                
                ' ----- Greeting/Sign-off Removal -----
                If line Like "Hello*" Or line Like "Hi*" Or line Like "Dear*" _
                Or line Like "Best regards*" Or line Like "Kind regards*" _
                Or line Like "Mit freundlichen Grüßen*" Then
                    skip = True
                End If
                
                ' If line is not a greeting/sign-off, process it
                If Not skip Then
                    Dim lineText As String
                    lineText = line
                    
                    ' ----- 1. Remove emails in <> or [] -----
                    Set re = CreateObject("VBScript.RegExp")
                    re.Pattern = "<[^>]+>|[\[][^]]+[\]]"
                    re.Global = True
                    lineText = re.Replace(lineText, "[email redacted]")
                    
                    ' ----- 2. Replace demoXXXX or keyXXXX -----
                    re.Pattern = "\b(?:demo|key)\d{4}\b"
                    lineText = re.Replace(lineText, "license")
                    
                    ' ----- 3. Replace names with 'User' -----
                    re.Pattern = "\b([A-Z][a-z]+(?:\s[A-Z][a-z]+)+)\b"
                    Set matches = re.Execute(lineText)
                    For Each match In matches
                        Select Case LCase(match.Value)
                            Case "Lucerne", "Hospital", "Cantonal", "Exit", "File", "PACS", "Support", "Application"
                                ' Ignore common non-names
                            Case Else
                                lineText = Replace(lineText, match.Value, "User")
                        End Select
                    Next
                    
                    ' ----- 4. Replace phone numbers -----
                    re.Pattern = "\b\d{3,4}[\s\-]?\d{3}[\s\-]?\d{3,4}\b"
                    lineText = re.Replace(lineText, "[phone redacted]")
                    
                    ' Add to result
                    cleaned = cleaned & lineText & vbCrLf
                End If
            Next
            
            ' Final assignment (trim trailing line)
            cell.Value = Trim(cleaned)
        End If
    Next cell
    
    MsgBox "Anonymization complete!"
End Sub


