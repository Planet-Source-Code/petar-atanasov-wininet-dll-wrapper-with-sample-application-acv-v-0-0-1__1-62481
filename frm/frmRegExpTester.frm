VERSION 5.00
Begin VB.Form frmRegExpTester 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parse server name and tokenize folders"
   ClientHeight    =   2736
   ClientLeft      =   36
   ClientTop       =   420
   ClientWidth     =   6600
   Icon            =   "frmRegExpTester.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2736
   ScaleWidth      =   6600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMain 
      Caption         =   "Server information"
      Height          =   1860
      Left            =   84
      TabIndex        =   1
      Top             =   756
      Width           =   6396
      Begin VB.TextBox txtRegExpPattern 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   10.2
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   396
         Left            =   1932
         TabIndex        =   7
         Text            =   "[a-zA-Z][\s\w\.-]*[a-zA-Z0-9]\/"
         Top             =   1260
         Width           =   4296
      End
      Begin VB.CommandButton CmdTest 
         Caption         =   "Tokenize"
         Height          =   348
         Left            =   168
         TabIndex        =   6
         Top             =   1260
         Width           =   1020
      End
      Begin VB.ComboBox cobServerParent 
         Height          =   288
         Left            =   1932
         TabIndex        =   4
         Top             =   840
         Width           =   4296
      End
      Begin VB.TextBox txtServerAddress 
         Height          =   288
         Left            =   1932
         TabIndex        =   2
         Top             =   336
         Width           =   4296
      End
      Begin VB.Label Label2 
         Caption         =   "Tokenized Address:"
         Height          =   264
         Left            =   168
         TabIndex        =   5
         Top             =   840
         Width           =   1608
      End
      Begin VB.Label lblLink 
         Caption         =   "Address:"
         Height          =   264
         Left            =   168
         TabIndex        =   3
         Top             =   336
         Width           =   1608
      End
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   432
      Left            =   84
      TabIndex        =   0
      Top             =   168
      Width           =   6396
   End
End
Attribute VB_Name = "frmRegExpTester"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MODULE_NAME           As String = "frmRegExpTester"
Private Const G_SRV_FOLDER_TOKENIZER As String = "[a-zA-Z][\s\w\.-]*[a-zA-Z0-9]\/"

Private m_sDebugID As String

Friend Sub frInit(ByVal sServerAddress As String)
    txtServerAddress = sServerAddress
    Me.Show
End Sub

Private Sub CmdTest_Click()
    Const FUNC_NAME As String = "Form_DblClick"
    
    On Error GoTo EH

    Dim oRegExp      As RegExp
    Dim oMatch       As Match
    Dim oMatches     As MatchCollection
    Dim sRetValue    As String
    Dim lIdx        As Long
    '---
    Set oRegExp = New RegExp
    oRegExp.Pattern = txtRegExpPattern
    '---
    With oRegExp
        .IgnoreCase = True
        .Global = True
    End With
    cobServerParent.Clear
    '---Todo: First Clear the shit from the url  whitespace %20 and etc
    '--- add / az last char if needed - inache ne e folder i ne go razpoznava
    txtServerAddress.Text = ClearString(txtServerAddress.Text)
    '--- append /
    If Right(txtServerAddress.Text, 1) <> "/" Then
        txtServerAddress.Text = txtServerAddress.Text & "/"
    End If
    '--- make RegExp
    If (oRegExp.Test(txtServerAddress) = True) Then
        Set oMatches = oRegExp.Execute(txtServerAddress)
        For lIdx = 0 To oMatches.Count - 1
            cobServerParent.AddItem oMatches.Item(lIdx)
        Next
        cobServerParent.Text = oMatches.Count & " items added."
        lblStatus.Caption = " RegExp Test call finished succesfully: "
        lblStatus.Caption = lblStatus.Caption & oMatches.Count & " items found."
    Else
        lblStatus.Caption = " Regular expression failed !"
    End If
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

