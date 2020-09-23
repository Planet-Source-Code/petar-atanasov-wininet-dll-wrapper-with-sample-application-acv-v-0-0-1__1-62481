Attribute VB_Name = "mdGlobals"
Option Explicit
Public Const DUMP_FILES             As String = "\FILES_DUMP.LOG"
Public Const DUMP_LOG_FILE          As String = "\DUMP_LOG.LOG"
Public Const REF_LOG_FILE           As String = "\REF_LOG.LOG"
Private Const MODULE_NAME           As String = "mdGlobals"

'--- Error handling section
Public Sub PrintError(ByVal sModName As String, ByVal sFuncName As String)
    Dim sErrFile As String
    Dim vDateTime As Variant
    Dim sDateTime As String
    On Error Resume Next
    vDateTime = CStr(Now())
    sDateTime = Replace(CStr(Now()), ".", "-")
    sDateTime = Replace(sDateTime, ":", "_")
    sDateTime = Replace(sDateTime, " ", "_")
    '---Print err to Immediate
    Debug.Print "========= ERROR STATUS =========" & vbCrLf
    Debug.Print "======= Error in module : " & sModName & vbCrLf
    Debug.Print "======= Error in function: " & sFuncName & vbCrLf
    Debug.Print Err.Number & vbCrLf
    Debug.Print Err.Description & vbCrLf
    Debug.Print Err.Source & vbCrLf
    Debug.Print Err.LastDllError & vbCrLf
    Debug.Print "================================" & vbCrLf
    '--- log error to file
    sErrFile = App.Path & "\ERR_" & sDateTime & ".log"
    Open sErrFile For Binary Access Write As #1
    Put #1, , "========= ERROR STATUS =========" & vbCrLf
    Put #1, , "> Error generated on:" & CStr(vDateTime) & vbCrLf
    Put #1, , "> Error in module : " & sModName & vbCrLf
    Put #1, , "> Error in function: " & sFuncName & vbCrLf
    Put #1, , "> Err. No: " & Err.Number & vbCrLf
    Put #1, , "> Err. Desc: " & Err.Description & vbCrLf
    Put #1, , "> Err. Source: " & Err.Source & vbCrLf
    Put #1, , "> Err. LastDllErr: " & Err.LastDllError & vbCrLf
    Put #1, , "================================" & vbCrLf
    Close #1
End Sub

Public Sub DumpToFile(ByVal s1 As String, ByVal s2 As String, ByVal s3 As String, ByVal s4 As String)
    Dim sLogFile As String
    Dim vDateTime As Variant
    Dim sDateTime As String
    
    vDateTime = CStr(Now())
    sDateTime = Replace(CStr(Now()), ".", "-")
    sDateTime = Replace(sDateTime, ":", "_")
    sDateTime = Replace(sDateTime, " ", "_")
    '--- log to file
    sLogFile = App.Path & DUMP_LOG_FILE
    Open sLogFile For Binary Access Write As #1
    Put #1, , "========= LOG DUMP =========" & vbCrLf
    Put #1, , "> Log generated on:" & CStr(vDateTime) & vbCrLf
    Put #1, , s1 & vbCrLf
    Put #1, , s2 & vbCrLf
    Put #1, , s3 & vbCrLf
    Put #1, , s4 & vbCrLf
    Put #1, , "================================" & vbCrLf
    Close #1
End Sub

Public Function HandleErr(sModuleName As String, sFunctionName As String) As VbMsgBoxResult
    Dim sPrompt As String
    Dim lRes As VbMsgBoxResult
    sPrompt = "Program error in module " & sModuleName & " in function " & sFunctionName
    lRes = MsgBox(sPrompt, vbRetryCancel, "Error message")
    HandleErr = lRes
End Function

Public Function ClearString(ByVal sString As String) As String
    Const FUNC_NAME As String = "ClearString"
    
    On Error GoTo EH

    '--- http://www.everything2.com/index.pl?node_id=1350053
    '--- Url escape sequences list
    
    sString = Replace(sString, "%20", " ")
    sString = Replace(sString, "%26", "&")
    sString = Replace(sString, "%3C", "<")
    sString = Replace(sString, "%3E", ">")
    sString = Replace(sString, "%22", """")

    '--- when used as 1 symbol of folder name,
    '--- must not be streaped, otherwise
    '--- no bad link reported - goes to parent folder
    '--- so we do not replace diez
    'sString = Replace(sString, "%23", "#")

    sString = Replace(sString, "%24", "$")
    sString = Replace(sString, "%25", "%")
    sString = Replace(sString, "%27", "'")
    sString = Replace(sString, "%2B", "+")
    sString = Replace(sString, "%2b", "+")
    sString = Replace(sString, "%2C", ",")
    sString = Replace(sString, "%2c", ",")
    sString = Replace(sString, "%2F", "/")
    sString = Replace(sString, "%2f", "/")
    sString = Replace(sString, "%3A", ":")
    sString = Replace(sString, "%3a", ":")
    sString = Replace(sString, "%3B", ";")
    sString = Replace(sString, "%3b", ";")
    sString = Replace(sString, "%3D", "=")
    sString = Replace(sString, "%3d", "=")
    sString = Replace(sString, "%3F", "?")
    sString = Replace(sString, "%3f", "?")
    sString = Replace(sString, "%40", "@")
    sString = Replace(sString, "%5B", "[")
    sString = Replace(sString, "%5b", "[")
    sString = Replace(sString, "%5C", "\")
    sString = Replace(sString, "%5c", "\")
    sString = Replace(sString, "%5D", "]")
    sString = Replace(sString, "%5d", "]")
    sString = Replace(sString, "%5E", "^")
    sString = Replace(sString, "%5e", "^")
    sString = Replace(sString, "%60", "`")
    sString = Replace(sString, "%7B", "{")
    sString = Replace(sString, "%7b", "{")
    sString = Replace(sString, "%7C", "|")
    sString = Replace(sString, "%7c", "|")
    sString = Replace(sString, "%7D", "}")
    sString = Replace(sString, "%7d", "}")
    sString = Replace(sString, "%7E", "~")
    sString = Replace(sString, "%7e", "~")

    '--- check for single quotation mark or whatever is it
    'sString = Replace(sString, "%60", "'")
    ClearString = sString
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function ClearString2(ByVal sString As String) As String
    Const FUNC_NAME As String = "ClearString"
    
    On Error GoTo EH
    '--- http://wdvl.internet.com/Authoring/HTML/Entities/common.html
    '--- HTML escape sequences list
    '--- ampersand
    sString = Replace(sString, "&amp;", "&")
    sString = Replace(sString, "&#38;", "&")
    '--- acute accent
    sString = Replace(sString, "&acute;", "`")
    sString = Replace(sString, "&#180;", "`")
    '--- broken vertical bar
    sString = Replace(sString, "&brvbar;", "|")
    sString = Replace(sString, "&#166;", "|")
    '--- cedilla
    sString = Replace(sString, "&cedil;", ",")
    sString = Replace(sString, "&#184;", ",")
    '--- less than
    sString = Replace(sString, "&lt;", "<")
    sString = Replace(sString, "&#60;", "<")
    '--- greater than
    sString = Replace(sString, "&gt;", ">")
    sString = Replace(sString, "&#62;", ">")
    '--- left-pointing double angle quotation mark
    sString = Replace(sString, "&laquo;", "«")
    sString = Replace(sString, "&#171;", "«")
    '--- right-pointing double angle quotation mark
    sString = Replace(sString, "&raquo;", "»")
    sString = Replace(sString, "&#187;", "»")
    '--- copyright
    sString = Replace(sString, "&copy;", "©")
    sString = Replace(sString, "&#169;", "©")
    '--- registered
    sString = Replace(sString, "&reg;", "®")
    sString = Replace(sString, "&#174;", "®")
    '--- trade mark
    sString = Replace(sString, "&trade;", "™")
    sString = Replace(sString, "&#8482;", "™")
    '--- to be continued if needed...
    ClearString2 = sString
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Function IsFilePresent(ByVal sFilePath As String)
    Const FUNC_NAME As String = "IsFilePresent"
    
    On Error GoTo EH
    If Len(Dir(sFilePath)) <> 0 Then
        IsFilePresent = True
    Else
        IsFilePresent = False
    End If
    '--- exit function
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Public Sub glFixWindowSize(ByRef lWidth As Long, ByRef lHeight As Long, oForm As Form)
    On Error Resume Next
    If oForm.Width <= lWidth Then
        oForm.Width = lWidth
    End If
    If oForm.Height <= lHeight Then
        oForm.Height = lHeight
    End If
End Sub



