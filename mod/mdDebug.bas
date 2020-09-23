Attribute VB_Name = "mdDebug"
Option Explicit
Private Const MODULE_NAME           As String = "mdDebug"

'---------------------------------------------------------
'--- DEBUG SECTION - API DECLARE
'---------------------------------------------------------

Private Declare Function GetLastError Lib "kernel32" () As Long

Private Declare Sub OutputDebugString Lib "kernel32" _
                Alias "OutputDebugStringA" (ByVal lpOutputString As String)

'---------------------------------------------------------
'--- DEBUG SECTION - FUNCTION DECLARE
'---------------------------------------------------------

Private Function pvGetDebugID() As Long
    Static lDebugID As Long
    lDebugID = lDebugID + 1
    pvGetDebugID = lDebugID
End Function

Public Sub DebugInit(sObject As String, sDebugID As String)
    sDebugID = pvGetDebugID()
    DebugPrint "Init " & sObject & " " & sDebugID
End Sub

Public Sub DebugTerm(sObject As String, sDebugID As String)
    DebugPrint "Term " & sObject & " " & sDebugID
End Sub

Public Sub DebugPrint(sText As String)
    OutputDebugString App.ThreadID & ": " & sText & " (" & Timer & ")" & vbCrLf
End Sub
                
Public Function DebugGetLastError() As Long
    DebugGetLastError = GetLastError()
End Function
