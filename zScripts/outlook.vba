Option Explicit

Sub ExportSelectedEmailsToBoth()
    Dim objMail As Object
    Dim fso As Object
    Dim oFile As Object
    Dim shell As Object

    Dim folderPath As String
    Dim baseFileName As String
    Dim mdFilePath As String
    Dim fullJsonPath As String

    Dim mdContent As String
    Dim jsonContent As String

    Dim i As Long
    Dim k As Variant
    Dim bodyText As String
    Dim foundVal As String
    Dim fields As Object
    Dim fieldKeys As Variant

    Dim pythonExe As String
    Dim scriptPath As String
    Dim command As String

    ' --- CONFIGURATION ---
    folderPath = "C:\Users\Jason\Obsidian\Logic\Inbox\"
    baseFileName = "Logistics_Export_" & Format(Now, "yyyy-mm-dd_hh-nn-ss")
    mdFilePath = folderPath & baseFileName & ".md"
    fullJsonPath = folderPath & baseFileName & ".json"

    pythonExe = "C:\Users\Jason\AppData\Local\Programs\Python\Python313\python.exe"
    scriptPath = "C:\Users\Jason\Projects\EmailProcessor\zScripts\email_processor-msg-perplex-original.py"

    If Application.ActiveExplorer.Selection.Count = 0 Then
        MsgBox "Please select at least one email.", vbExclamation
        Exit Sub
    End If

    ' Initialize contents
    mdContent = "---" & vbCrLf & _
                "type: logistics_export" & vbCrLf & _
                "---" & vbCrLf & vbCrLf

    jsonContent = "[" & vbCrLf

    fieldKeys = Array( _
        "File Ref", _
        "Client Ref", _
        "Consignee", _
        "Description", _
        "PIN No", _
        "Serial No", _
        "Vessel", _
        "Voy", _
        "Bill No", _
        "ETA", _
        "Tarrif Code, Weight & Cube", _
        "Client PO", _
        "QUOTATION NR", _
        "TRANSPORTER" _
    )

    ' Loop through selected emails
    For i = 1 To Application.ActiveExplorer.Selection.Count
        Set objMail = Application.ActiveExplorer.Selection.Item(i)

        bodyText = ConvertHTMLToMarkdown(objMail.HTMLBody)

        Set fields = CreateObject("Scripting.Dictionary")

        For Each k In fieldKeys
            foundVal = ExtractField(bodyText, CStr(k))
            fields.Add CStr(k), foundVal
        Next k

        ' Append to Markdown
        mdContent = mdContent & "## " & objMail.Subject & vbCrLf
        mdContent = mdContent & "> [!info] Logistics Metadata" & vbCrLf

        For Each k In fieldKeys
            mdContent = mdContent & "> **" & CStr(k) & ":** " & fields(CStr(k)) & vbCrLf
        Next k

        mdContent = mdContent & vbCrLf & "---" & vbCrLf & bodyText & vbCrLf & vbCrLf

        ' Append to JSON
        jsonContent = jsonContent & "  {" & vbCrLf
        jsonContent = jsonContent & "    ""subject"": """ & EscapeJSON(objMail.Subject) & """," & vbCrLf

        For Each k In fieldKeys
            jsonContent = jsonContent & _
                "    """ & LCase(Replace(CStr(k), " ", "_")) & """: """ & EscapeJSON(fields(CStr(k))) & """," & vbCrLf
        Next k

        jsonContent = jsonContent & _
            "    ""body"": """ & EscapeJSON(bodyText) & """" & vbCrLf & _
            "  }"

        If i < Application.ActiveExplorer.Selection.Count Then
            jsonContent = jsonContent & ","
        End If

        jsonContent = jsonContent & vbCrLf
    Next i

    jsonContent = jsonContent & "]"

    ' Save Markdown
    Set fso = CreateObject("Scripting.FileSystemObject")

    Set oFile = fso.CreateTextFile(mdFilePath, True, True)
    oFile.Write mdContent
    oFile.Close

    ' Save JSON
    Set oFile = fso.CreateTextFile(fullJsonPath, True, True)
    oFile.Write jsonContent
    oFile.Close

    ' Call Python script
    Set shell = CreateObject("WScript.Shell")
    command = """" & pythonExe & """" & " " & _
              """" & scriptPath & """" & " " & _
              """" & fullJsonPath & """" & _
              " > ""C:\Users\Jason\Obsidian\Logic\Inbox\python_run.log"" 2>&1"

    shell.Run command, 1, True

    MsgBox "Export complete!" & vbCrLf & _
           "Markdown: " & mdFilePath & vbCrLf & _
           "JSON: " & fullJsonPath, vbInformation
End Sub

Function ExtractField(ByVal txt As String, ByVal fieldName As String) As String
    Dim regEx As Object
    Dim matches As Object

    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Pattern = fieldName & "[:\s]+([^\r\n]+)"
    regEx.IgnoreCase = True
    regEx.Multiline = True

    If regEx.Test(txt) Then
        Set matches = regEx.Execute(txt)
        ExtractField = Trim(matches(0).SubMatches(0))
    Else
        ExtractField = "N/A"
    End If
End Function

Function ConvertHTMLToMarkdown(ByVal html As String) As String
    Dim regEx As Object

    Set regEx = CreateObject("VBScript.RegExp")

    With regEx
        .Global = True
        .IgnoreCase = True
        .Multiline = True

        .Pattern = "<br\s*/?>|</p>|</div>"
        html = .Replace(html, vbCrLf)

        .Pattern = "<[^>]*>"
        html = .Replace(html, "")
    End With

    ConvertHTMLToMarkdown = Trim(html)
End Function

Function EscapeJSON(ByVal txt As String) As String
    txt = Replace(txt, "\", "\\")
    txt = Replace(txt, """", "\""")
    txt = Replace(txt, vbCrLf, "\n")
    txt = Replace(txt, vbLf, "\n")
    txt = Replace(txt, vbCr, "\n")
    EscapeJSON = txt
End Function