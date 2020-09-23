VERSION 5.00
Begin VB.Form frmNetDisplay 
   ClientHeight    =   3504
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   9192
   Icon            =   "frmNetDisplay.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3504
   ScaleWidth      =   9192
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frmNetDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--- Consts
Const MODULE_NAME                   As String = "frmNetDisplay"
Private Const MSG_CTL_NOT_CREATED   As String = "Could not create WebBrowser control"
'--- variables
Private WithEvents m_oWebControl As VBControlExtender
Attribute m_oWebControl.VB_VarHelpID = -1
Private m_sURL As String
Private m_sDebugID As String

Friend Sub frInit(sUrl As String)
    m_sURL = sUrl
    Me.Caption = "Content of: " & m_sURL
    Me.Show vbModal
End Sub

Private Sub Form_Load()
    Const FUNC_NAME As String = "Form_Load"
    On Error GoTo EH
    
    Set m_oWebControl = Controls.Add("Shell.Explorer", "webctl", Me)
    If Not m_oWebControl Is Nothing Then
        m_oWebControl.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
        m_oWebControl.Visible = True
        m_oWebControl.object.navigate m_sURL
    End If
    Exit Sub
EH:
    PrintError MODULE_NAME, FUNC_NAME
    MsgBox MSG_CTL_NOT_CREATED, vbInformation
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Not m_oWebControl Is Nothing Then
        m_oWebControl.Width = Me.ScaleWidth
        m_oWebControl.Height = Me.ScaleHeight
    End If
    glFixWindowSize 8500, 6500, Me
End Sub

Private Sub Form_Initialize()
    Put #5, , "> " & MODULE_NAME
    Put #5, , "> Form_Initialize() " & vbCrLf
    DebugInit MODULE_NAME, m_sDebugID
End Sub

Private Sub Form_Terminate()
    Put #5, , "---" & MODULE_NAME
    Put #5, , "> Form_Terminate() " & vbCrLf
    DebugTerm MODULE_NAME, m_sDebugID
End Sub
