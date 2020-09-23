VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDisplay 
   Caption         =   "Content of: "
   ClientHeight    =   6012
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   9780
   Icon            =   "frmDisplay.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6012
   ScaleWidth      =   9780
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cobRes 
      Height          =   288
      Left            =   2436
      TabIndex        =   3
      Top             =   504
      Width           =   7236
   End
   Begin VB.TextBox txtSearch 
      Height          =   288
      Left            =   1344
      TabIndex        =   2
      Text            =   "*.txt"
      Top             =   504
      Width           =   936
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   288
      Left            =   84
      TabIndex        =   1
      Top             =   504
      Width           =   1188
   End
   Begin RichTextLib.RichTextBox oRTF 
      Height          =   4968
      Left            =   84
      TabIndex        =   0
      Top             =   924
      Width           =   9588
      _ExtentX        =   16912
      _ExtentY        =   8763
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmDisplay.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Status"
      ForeColor       =   &H80000008&
      Height          =   348
      Left            =   84
      TabIndex        =   4
      Top             =   84
      Width           =   9588
   End
End
Attribute VB_Name = "frmDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const MODULE_NAME           As String = "frmDisplay"
Private Const MSG_NO_FILE           As String = "Dump file not found."
Private Const MSG_SEARCH_TEXT       As String = "Enter keyword for search, *.ext or name.* for group search."
Private Const MSG_BIG_SEARCH        As String = "The searched text is too big and this may slow down the search. Do you want to proceed anyway ?"
Private Const CAP_BIG_SEARCH        As String = "WARNING !!!"
Private Const SEARCH_TXT_LEN        As Long = 15000

Private m_sDebugID As String

Friend Sub frInit(ByVal sLink As String)
    Const FUNC_NAME As String = "Form_Load()"
    
    On Error GoTo EH
    '--- set form caption
    Me.Caption = Me.Caption & ClearString(ClearString2(sLink))
    If IsFilePresent(App.Path & DUMP_FILES) Then
        oRTF.FileName = App.Path & DUMP_FILES
    Else
        lblStatus.Caption = MSG_NO_FILE
    End If
    '--- show form
    Me.Show
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

Private Sub Form_Load()
    Const FUNC_NAME As String = "Form_Load()"

    On Error GoTo EH
    lblStatus.Caption = MSG_SEARCH_TEXT
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

Private Sub Form_Resize()
    On Error Resume Next
    oRTF.Width = Width - 250
    oRTF.Height = Height - 1450
    lblStatus.Width = oRTF.Width
    cobRes.Width = lblStatus.Width - (cmdSearch.Width + txtSearch.Width) - 200
End Sub

Private Sub cmdSearch_Click()
    Const FUNC_NAME As String = "cmdSearch_Click()"
    Dim sSearchPattern As String
    Dim vUserResponse As VbMsgBoxResult
    
    On Error GoTo EH
    Me.MousePointer = vbHourglass
    cobRes.Clear
    '--- capture input and create pattern
    If Left(txtSearch.Text, 2) = "*." Then
        sSearchPattern = ".*\." & Mid(txtSearch.Text, 3)
    ElseIf Right(txtSearch.Text, 2) = "\.*" Then
        sSearchPattern = txtSearch.Text
    Else
        sSearchPattern = ".*" & txtSearch.Text & ".*"
    End If
    '--- extra long text
    '--- Todo: consider splitting results later
    If Len(oRTF.Text) > CLng(SEARCH_TXT_LEN) Then
        vUserResponse = MsgBox(MSG_BIG_SEARCH, vbOKCancel, CAP_BIG_SEARCH)
        If vUserResponse = vbOK Then
        '--- do search
        pvRxpSearch sSearchPattern, oRTF.Text
        End If
    Else
        '--- do search
        pvRxpSearch sSearchPattern, oRTF.Text
    End If
    '--- restore mouse pointer
    Me.MousePointer = vbDefault
    '--- exit sub
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

Private Sub pvRxpSearch(ByVal sSearchCriteria As String, ByVal sSearchText As String)
    Const FUNC_NAME As String = "pvRxpSearch"
    Dim oRegExp      As RegExp
    Dim oMatches     As MatchCollection
    Dim lIdx        As Long
    
    On Error GoTo EH
    '--- Create RegExp obj
    Set oRegExp = New RegExp
    '--- set pattern
    oRegExp.Pattern = sSearchCriteria
    '--- prepare RegExp Object
    With oRegExp
        .IgnoreCase = True
        .Global = True
    End With
    '--- make RegExp
    If (oRegExp.Test(sSearchText) = True) Then
        Set oMatches = oRegExp.Execute(sSearchText)
        For lIdx = 0 To oMatches.Count - 1
            cobRes.AddItem oMatches.Item(lIdx)
        Next
        cobRes.Text = oMatches.Count & " items added."
        lblStatus.Caption = " RegExp Test call finished succesfully: "
        lblStatus.Caption = lblStatus.Caption & oMatches.Count & " items found."
    Else
        lblStatus.Caption = " Regular expression failed !"
    End If
    '--- exit sub
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
