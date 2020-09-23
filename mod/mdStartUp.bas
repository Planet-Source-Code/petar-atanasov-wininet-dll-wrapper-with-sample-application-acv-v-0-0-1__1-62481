Attribute VB_Name = "mdStartUp"
Option Explicit
'--- Consts
Private Const MODULE_NAME           As String = "mdStartUp"

Public Sub Main()
    Const FUNC_NAME As String = "Main()"
    Dim oFrm As New frmMain
    
    On Error GoTo EH
    If IsFilePresent(App.Path & REF_LOG_FILE) Then
        Kill App.Path & REF_LOG_FILE
    End If
    oFrm.frInit
    Exit Sub
EH:
    PrintError MODULE_NAME, FUNC_NAME
    Select Case HandleErr(MODULE_NAME, FUNC_NAME)
        Case vbRetry
            Resume
        Case vbCancel
            Exit Sub
    End Select
End Sub
