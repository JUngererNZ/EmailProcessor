Sub ExportSelectedEmailsToBoth()
    Dim objMail As Object, fso As Object, oFile As Object
    Dim folderPath As String, baseFileName As String
    Dim mdContent As String, jsonContent As String
    Dim i As Integer, j As Integer
    Dim fullJsonPath As String ' Added to track the specific file path
    
    ' --- CONFIGURATION ---
    folderPath = "C:\Users\Jason\Obsidian\Logic\Inbox\"
    ' We define the base name once so both files match
    baseFileName = "Logistics_Export_" & Format(Now, "yyyy-mm-dd_hh-nn-ss")
    fullJsonPath = folderPath & baseFileName & ".json"
    
    If Application.ActiveExplorer.Selection.Count = 0 Then Exit Sub

    ' Initialize Contents [cite: 1, 2]
    mdContent = "---" & vbCrLf & "type: logistics_export" & vbCrLf & "---" & vbCrLf & vbCrLf
    jsonContent = "[" & vbCrLf
    
    ' Loop through selected emails [cite: 1]
    For i = 1 To Application.ActiveExplorer.Selection.Count
        Set objMail = Application.ActiveExplorer.Selection.Item(i)
        Dim bodyText As String: bodyText = ConvertHTMLToMarkdown(objMail.HTMLBody)
        
        ' --- DYNAMIC EXTRACTION [cite: 3, 4] ---
        Dim fields As Object: Set fields = CreateObject("Scripting.Dictionary")
        Dim fieldKeys As Variant
        fieldKeys = Array("File Ref", "Client Ref", "Consignee", "Description", "PIN No", _
                          "Serial No", "Vessel", "Voy", "Bill No", "ETA", _
                          "Tarrif Code, Weight & Cube", "Client PO", "QUOTATION NR", "TRANSPORTER")
        
        Dim k As Variant, foundVal As String
        For Each k In fieldKeys
            foundVal = ExtractField(bodyText, CStr(k))
            fields.Add k, foundVal
        Next k

        ' 1. Append to Markdown [cite: 5]
        mdContent = mdContent & "## " & objMail.Subject & vbCrLf
        mdContent = mdContent & "> [!info] Logistics Metadata" & vbCrLf
        For Each k In fieldKeys
            mdContent = mdContent & "> **" & k & ":** " & fields(k) & vbCrLf
        Next k
        mdContent = mdContent & vbCrLf & "---" & vbCrLf & bodyText & vbCrLf & vbCrLf
        
        ' 2. Append to JSON [cite: 6, 7]
        jsonContent = jsonContent & "  {" & vbCrLf & _
                      "    ""subject"": """ & EscapeJSON(objMail.Subject) & """," & vbCrLf
        
        For Each k In fieldKeys
            jsonContent = jsonContent & "    """ & LCase(Replace(k, " ", "_")) & """: """ & EscapeJSON(fields(k)) & """," & vbCrLf
        Next k
        
        jsonContent = jsonContent & "    ""body"": """ & EscapeJSON(bodyText) & """" & vbCrLf & "  }"
        If i < Application.ActiveExplorer.Selection.Count Then jsonContent = jsonContent & ","
        jsonContent = jsonContent & vbCrLf
    Next i
    
    jsonContent = jsonContent & "]" [cite: 8]

    ' --- SAVE FILES ---
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Save Markdown
    Set oFile = fso.CreateTextFile(folderPath & baseFileName & ".md", True, True)
    oFile.Write mdContent
    oFile.Close
    
    ' Save JSON [cite: 12]
    Set oFile = fso.CreateTextFile(fullJsonPath, True, True)
    oFile.Write jsonContent
    oFile.Close
    
    ' --- CALL PYTHON SCRIPT  ---
    Dim shell As Object
    Set shell = CreateObject("WScript.Shell")
    
    Dim pythonExe As String: pythonExe = "python.exe"
    Dim scriptPath As String: scriptPath = """C:\Users\Jason\Projects\EmailProcessor\zScripts\email_processor-msg-perplex-original.py"""
    Dim command As String
    
    ' Build command using the specific JSON file path we just saved 
    command = pythonExe & " " & scriptPath & " """ & fullJsonPath & """"
    
    ' Run script (1 = show window, True = wait for Python to finish) 
    shell.Run command, 1, True
    
    MsgBox "Export complete! Python processed " & Application.ActiveExplorer.Selection.Count & " items.", vbInformation
End Sub

' --- HELPERS (Keep these exactly as you had them) ---

Function ExtractField(ByVal txt As String, ByVal fieldName As String) As String [cite: 9]
    Dim regEx As Object, matches As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = fieldName & "[:\s]+([^\r\n]+)"
    regEx.IgnoreCase = True: regEx.Multiline = True
    If regEx.Test(txt) Then
        Set matches = regEx.Execute(txt): ExtractField = Trim(matches(0).SubMatches(0))
    Else
        ExtractField = "N/A"
    End If
End Function

Function ConvertHTMLToMarkdown(ByVal html As String) As String [cite: 10]
    Dim regEx As Object: Set regEx = CreateObject("VBScript.RegExp")
    With regEx
        .Global = True: .IgnoreCase = True: .Multiline = True
        .Pattern = "<br\s*/?>|<\/p>|<\/div>": html = .Replace(html, vbCrLf)
        .Pattern = "<[^>]*>": html = .Replace(html, "")
    End With
    ConvertHTMLToMarkdown = Trim(html)
End Function

Function EscapeJSON(ByVal txt As String) As String [cite: 11]
    txt = Replace(txt, "\", "\\")
    txt = Replace(txt, """", "\""")
    txt = Replace(txt, vbCrLf, "\n"): txt = Replace(txt, vbLf, "\n"): txt = Replace(txt, vbCr, "\n")
    EscapeJSON = txt
End Function