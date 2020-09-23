VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Apache Content Viewer"
   ClientHeight    =   4884
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   7236
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4884
   ScaleWidth      =   7236
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraHeader 
      Caption         =   "Navigation"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3036
      Left            =   84
      TabIndex        =   4
      Top             =   84
      Width           =   7068
      Begin VB.ComboBox cobServerAddressTokenized 
         Height          =   288
         Left            =   1344
         TabIndex        =   11
         Text            =   "Link tokenized"
         Top             =   2604
         Width           =   2200
      End
      Begin VB.ComboBox cobMails 
         Height          =   288
         Left            =   4872
         TabIndex        =   10
         Text            =   "Mails"
         Top             =   2184
         Width           =   2100
      End
      Begin VB.ComboBox cobNA 
         Height          =   288
         Left            =   3696
         TabIndex        =   9
         Text            =   "Not categorized"
         Top             =   2604
         Width           =   2200
      End
      Begin VB.ComboBox cobFiles 
         Height          =   288
         Left            =   2520
         TabIndex        =   8
         Text            =   "Files"
         Top             =   2184
         Width           =   2200
      End
      Begin VB.ComboBox cobFolders 
         Height          =   288
         Left            =   168
         TabIndex        =   7
         Text            =   "Folders"
         Top             =   2184
         Width           =   2200
      End
      Begin VB.TextBox txtURL 
         Height          =   288
         Left            =   672
         TabIndex        =   5
         Tag             =   "http://rack8.free.evro.net/ROT/"
         Text            =   "http://archive.apache.org/dist/"
         Top             =   1260
         Width           =   6312
      End
      Begin VB.CommandButton cmdLink 
         Caption         =   "Reap Links"
         Height          =   800
         Left            =   168
         Picture         =   "frmMain.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   252
         Width           =   1400
      End
      Begin VB.CommandButton cmdDisplay 
         Caption         =   "Display content"
         Height          =   800
         Left            =   3780
         Picture         =   "frmMain.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   252
         Width           =   1400
      End
      Begin VB.CommandButton cmdTraverse 
         Caption         =   "Traverse"
         Enabled         =   0   'False
         Height          =   800
         Left            =   2016
         Picture         =   "frmMain.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   252
         Width           =   1400
      End
      Begin VB.CommandButton cmdRegExpTester 
         Caption         =   "Parse URL"
         Height          =   800
         Left            =   5544
         Picture         =   "frmMain.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   252
         Width           =   1400
      End
      Begin VB.Image Image1 
         Height          =   384
         Left            =   3024
         Picture         =   "frmMain.frx":14F2
         Top             =   1680
         Width           =   3108
      End
      Begin VB.Label Label1 
         Caption         =   "Categorized content:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Left            =   168
         TabIndex        =   13
         Top             =   1764
         Width           =   5976
      End
      Begin VB.Label lblLink 
         Caption         =   "URL:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   264
         Left            =   168
         TabIndex        =   12
         Top             =   1260
         Width           =   600
      End
   End
   Begin VB.Label lblStatus 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   1608
      Left            =   84
      TabIndex        =   6
      Top             =   3192
      Width           =   7068
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--- Consts
Private Const MODULE_NAME           As String = "frmMain"
Private Const STR_ITM_LOAD          As String = " items loaded."
Private m_sHTML As String
Private m_bIsLinkAbsolute As Boolean
Private m_sDebugID As String

Friend Sub frInit()
    Const FUNC_NAME As String = "frInit()"
    
    On Error GoTo EH
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
    Const FUNC_NAME As String = "Form_Load"
    
    On Error GoTo EH
    Me.lblStatus.Caption = " Form Loaded: " & "{" & GetGUID & "}"
    '--- ToDo: use the GUID for ROOT key, later
    Me.cmdDisplay.Enabled = IsFilePresent(App.Path & DUMP_FILES)
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

Private Sub cmdLink_Click()
    Const FUNC_NAME As String = "cmdLink_Click"
    '--- test link class here
    Dim oLink As New cLink
    Dim sHTML As String, sHTMBuf As String
    
    On Error GoTo EH
    pvClear True
    Me.MousePointer = vbHourglass
    '--- basic check for folder
    If Right(txtURL.Text, 1) <> "/" Then
        txtURL.Text = txtURL.Text & "/"
    End If
    '--- get link
    acvGetLink sHTMBuf, Trim(ClearString2(txtURL.Text))
    sHTML = Trim(sHTMBuf)
    m_sHTML = sHTML
    '--- init link
    oLink.frInit Trim(ClearString2(txtURL.Text)), sHTML
    m_bIsLinkAbsolute = oLink.IsLinkAbsolute
    '--- fix mousepointer
    Me.MousePointer = vbDefault
    '--- get children
    '--- oLink.LinkBaseAddress & oLink.LinkFolders(0) is special folder - parent of current link
    '--- 1st folder to n -> UBound(oLink.LinkFolders)
    '--- oLink.LinkBaseAddress & oLink.LinkFolders(1)
    pvClear True
    FillControls oLink
    cmdTraverse.Enabled = True
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

Private Sub cmdDisplay_Click()
    Const FUNC_NAME As String = "cmdDisplay_Click()"
    Dim oFrm As New frmDisplay
    
    On Error GoTo EH
    oFrm.frInit ClearString(ClearString2(Trim(txtURL.Text)))
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

Private Sub cmdTraverse_Click()
    Const FUNC_NAME As String = "cmdTraverse_Click()"
    Dim oFrm As New frmNavigator
    
    On Error GoTo EH
    Me.MousePointer = vbHourglass
    oFrm.frInit ClearString(ClearString2(Trim(txtURL.Text))), m_sHTML, m_bIsLinkAbsolute, Me
    'lblStatus.Caption = vbNullString
    lblStatus.Caption = lblStatus.Caption & " Link Trace Completed"
    cmdDisplay.Enabled = IsFilePresent(App.Path & DUMP_FILES)
    Me.MousePointer = vbDefault
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

Private Sub cmdRegExpTester_Click()
    Const FUNC_NAME As String = "cmdRegExpTester_Click"
    Dim oFrm As New frmRegExpTester
    
    On Error GoTo EH
    
    oFrm.frInit ClearString(ClearString2(Trim(txtURL.Text)))
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

Private Sub FillControls(ByRef oLink As cLink)
    Const FUNC_NAME As String = "FillControls"
    Dim lIdx As Long
    
    On Error GoTo EH
        
    pvClear False
    If oLink.LinkProcessed Then
        '--- fill folders
        If UBound(oLink.LinkArrFolders) <> 0 Then
            For lIdx = 0 To UBound(oLink.LinkArrFolders)
                cobFolders.AddItem (oLink.LinkArrFolders(lIdx))
            Next lIdx
        End If
        '--- fill files
        If UBound(oLink.LinkArrFiles) <> 0 Then
            For lIdx = 0 To UBound(oLink.LinkArrFiles)
                cobFiles.AddItem (oLink.LinkArrFiles(lIdx))
            Next lIdx
        End If
        '--- fill N/A
        If UBound(oLink.LinkArrNA) <> 0 Then
            For lIdx = 0 To UBound(oLink.LinkArrNA)
                cobNA.AddItem (oLink.LinkArrNA(lIdx))
            Next lIdx
        End If
        '--- fill mails
        If UBound(oLink.LinkArrMails) <> 0 Then
            For lIdx = 0 To UBound(oLink.LinkArrMails)
                cobMails.AddItem (oLink.LinkArrMails(lIdx))
            Next lIdx
        End If
        If UBound(oLink.ServerAddressTokenized) <> 0 Then
            For lIdx = 0 To UBound(oLink.ServerAddressTokenized)
                cobServerAddressTokenized.AddItem (oLink.ServerAddressTokenized(lIdx))
            Next lIdx
        End If
        '--- status
        If UBound(oLink.LinkArrFiles) <> 0 Then
            cobFiles.Text = cobFiles.Text & UBound(oLink.LinkArrFiles) + 1 & STR_ITM_LOAD
        Else
            cobFiles.Text = cobFiles.Text & "0 " & STR_ITM_LOAD
        End If
        If UBound(oLink.LinkArrFolders) <> 0 Then
            cobFolders.Text = cobFolders.Text & UBound(oLink.LinkArrFolders) + 1 & STR_ITM_LOAD
        Else
            cobFolders.Text = cobFolders.Text & "0 " & STR_ITM_LOAD
        End If
        If UBound(oLink.LinkArrNA) <> 0 Then
            cobNA.Text = cobNA.Text & UBound(oLink.LinkArrNA) + 1 & STR_ITM_LOAD
        Else
            cobNA.Text = cobNA.Text & "0 " & STR_ITM_LOAD
        End If
        If UBound(oLink.LinkArrMails) <> 0 Then
            cobMails.Text = cobMails.Text & UBound(oLink.LinkArrMails) + 1 & STR_ITM_LOAD
        Else
            cobMails.Text = cobMails.Text & "0 " & STR_ITM_LOAD
        End If
        If UBound(oLink.ServerAddressTokenized) <> 0 Then
            cobServerAddressTokenized.Text = cobServerAddressTokenized.Text & _
            UBound(oLink.ServerAddressTokenized) + 1 & STR_ITM_LOAD
        Else
            cobServerAddressTokenized.Text = cobServerAddressTokenized.Text & _
            " 0 " & STR_ITM_LOAD
        End If
        '--- status
        lblStatus.Caption = lblStatus.Caption & " " _
                           & UBound(oLink.LinkArrFolders) & " folders, " _
                           & UBound(oLink.LinkArrFiles) & " files, " _
                           & UBound(oLink.LinkArrNA) & " N/A. " _
                           & "Status: " & oLink.LinkStatus
        Else
        lblStatus.Caption = "Link not processed. " & oLink.LinkStatus
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

Private Sub pvClear(Optional bClearStatusLabelas As Boolean)
    Const FUNC_NAME As String = "pvClear"
    
    On Error GoTo EH
    
    If bClearStatusLabelas Then
        lblStatus.Caption = vbNullString
    End If
    cobFolders.Clear
    cobFiles.Clear
    cobNA.Clear
    cobMails.Clear
    cobServerAddressTokenized.Clear
    cobFolders.Text = "Folders: "
    cobFiles.Text = "Files: "
    cobNA.Text = "N/A: "
    cobMails.Text = "Mails: "
    cobServerAddressTokenized.Text = "Link tokens: "
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
    Height = 5364
    Width = 7332
End Sub

Private Sub Form_Initialize()
    Dim sREF As String
    sREF = App.Path & REF_LOG_FILE
    Open sREF For Binary Access Write As #5
    Put #5, , "========= REF STATUS =========" & vbCrLf
    Put #5, , "> " & MODULE_NAME
    Put #5, , "> Form_Initialize() " & vbCrLf
    DebugInit MODULE_NAME, m_sDebugID
End Sub

Private Sub Form_Terminate()
    Put #5, , "---" & MODULE_NAME
    Put #5, , "> Form_Terminate() " & vbCrLf
    Put #5, , "================================" & vbCrLf
    DebugTerm MODULE_NAME, m_sDebugID
End Sub

