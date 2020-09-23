VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "Comctl32.ocx"
Begin VB.Form frmNavigator 
   Caption         =   "Content"
   ClientHeight    =   6900
   ClientLeft      =   48
   ClientTop       =   432
   ClientWidth     =   9612
   Icon            =   "frmNavigator.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6900
   ScaleWidth      =   9612
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraStruct 
      Caption         =   "Structure:"
      Height          =   5388
      Left            =   84
      TabIndex        =   2
      Top             =   1428
      Width           =   9420
      Begin VB.CheckBox chbNavigate 
         Caption         =   "On Doble click navigate to address."
         Height          =   192
         Left            =   168
         TabIndex        =   6
         Top             =   252
         Value           =   1  'Checked
         Width           =   5136
      End
      Begin ComctlLib.TreeView tvNav 
         Height          =   4548
         Left            =   168
         TabIndex        =   3
         Top             =   588
         Width           =   9084
         _ExtentX        =   16023
         _ExtentY        =   8022
         _Version        =   327682
         Indentation     =   0
         Style           =   7
         Appearance      =   1
      End
   End
   Begin VB.Frame fraHeader 
      Caption         =   "Status"
      Height          =   1272
      Left            =   84
      TabIndex        =   0
      Top             =   84
      Width           =   9420
      Begin VB.ComboBox cobBadLinks 
         Height          =   288
         Left            =   1176
         TabIndex        =   1
         Top             =   840
         Width           =   8076
      End
      Begin VB.Label lblLink 
         Caption         =   "Bad links list:"
         Height          =   264
         Left            =   168
         TabIndex        =   5
         Top             =   840
         Width           =   1104
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
         Left            =   168
         TabIndex        =   4
         Top             =   252
         Width           =   9084
      End
   End
End
Attribute VB_Name = "frmNavigator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--- consts
Private Const MODULE_NAME           As String = "frmNavigator"
Private Const MSG_BAD_LINKS         As String = "Bad links loaded."
Private Const MSG_NO_BAD_LINKS      As String = "All links processed and correct."
Private Const MSG_LINK_NOT_PROC     As String = "Link not processed."
'--- member variables
Private m_sLink As String
Private m_sHTML As String
Private m_bIsLinkAbsolute As Boolean
Private m_asBadLinks() As String
Private m_lBadLinkIndex As Long
Private m_sDebugID As String
Private lCounter As Long
Private m_oParentFrm As frmMain

Private Sub pvTraverseChild(ByVal oLink As cLink, _
                            ByVal sParentKey As String, _
                            ByVal sChildKey As String, _
                            ByVal oNode As Node, _
                            ByVal IsLinkAbsolute As Boolean, _
                            Optional ByVal sLogFile As String, _
                            Optional sFileDumpFile As String, _
                            Optional oParentForm As frmMain)
    Const FUNC_NAME As String = "pvTraverseChild()"
    Dim a_oLinkLevelChildren() As New cLink
    Dim a_oLinkKeys() As String
    Dim sHTML As String
    Dim lChilCount As Long, lIdx As Long
    Dim sPrefChildKey As String
    Dim lJdx As Long
    Dim sTemp As String, sFile As String, sTempHTMBuff As String
    Dim sLinkAddress As String

    On Error GoTo EH
    '---
    lChilCount = UBound(oLink.LinkArrFolders)
    sPrefChildKey = GetGUID
    '--- size arrays
    ReDim Preserve a_oLinkLevelChildren(lChilCount)
    ReDim Preserve a_oLinkKeys(lChilCount)
    '--- cycle level
    For lIdx = 1 To lChilCount
        If IsLinkAbsolute Then
            sLinkAddress = ClearString2(oLink.LinkArrFolders(lIdx))
        Else
            sLinkAddress = ClearString2(oLink.LinkBaseAddress & oLink.LinkArrFolders(lIdx))
        End If
        '--- Get Content here
        acvGetLink sTempHTMBuff, sLinkAddress
        sHTML = Trim(sTempHTMBuff)
        '--- init link obj
        a_oLinkLevelChildren(lIdx).frInit sLinkAddress, sHTML
        If a_oLinkLevelChildren(lIdx).LinkProcessed Then
            pvAddChild oNode, tvNav, sParentKey, sChildKey, a_oLinkLevelChildren(lIdx).LinkBaseAddress
            pvDisplayInfo oParentForm, a_oLinkLevelChildren(lIdx).LinkBaseAddress
            '--- add files here
            For lJdx = 0 To UBound(a_oLinkLevelChildren(lIdx).LinkArrFiles)
                If IsLinkAbsolute Then
                    sFile = a_oLinkLevelChildren(lIdx).LinkArrFiles(lJdx)
                    sTemp = a_oLinkLevelChildren(lIdx).LinkArrFiles(lJdx) & vbCrLf
                Else
                    sFile = a_oLinkLevelChildren(lIdx).LinkBaseAddress & a_oLinkLevelChildren(lIdx).LinkArrFiles(lJdx)
                    sTemp = a_oLinkLevelChildren(lIdx).LinkBaseAddress & a_oLinkLevelChildren(lIdx).LinkArrFiles(lJdx) & vbCrLf
                End If
                Put #101, , ClearString(ClearString2(sTemp))
                pvAddChild oNode, tvNav, sChildKey, GetGUID, ClearString(ClearString2(sFile))
                'pvDisplayInfo oParentForm, vbNullString
                pvDisplayInfo oParentForm, ClearString(ClearString2(sFile))
            Next lJdx
            '--- log to file
            Put #1, , a_oLinkLevelChildren(lIdx).LinkBaseAddress & vbCrLf
            Put #1, , "Parent Key: " & sParentKey & vbCrLf
            Put #1, , "Child Key: " & sChildKey & vbCrLf
            Put #1, , "Children count: " & UBound(a_oLinkLevelChildren(lIdx).LinkArrFolders) & vbCrLf
            Put #1, , " " & vbCrLf
            '--- key arr
            a_oLinkKeys(lIdx) = sChildKey
            '--- create next key
            sChildKey = GetGUID
            '--- if has children - then go, go, go
            If UBound(a_oLinkLevelChildren(lIdx).LinkArrFolders) > 0 Then
                '--- new key
                sPrefChildKey = GetGUID
                '--- log to file
                Put #1, , a_oLinkLevelChildren(lIdx).LinkBaseAddress & vbCrLf
                Put #1, , lIdx & " has " & UBound(a_oLinkLevelChildren(lIdx).LinkArrFolders) & " children" & vbCrLf
                Put #1, , " " & vbCrLf
                Put #1, , "------------------" & vbCrLf
                '--- recursive call
                pvTraverseChild a_oLinkLevelChildren(lIdx), a_oLinkKeys(lIdx), GetGUID, oNode, IsLinkAbsolute, sLogFile, sFileDumpFile, oParentForm
            End If
        Else
            m_lBadLinkIndex = m_lBadLinkIndex + 1
            ReDim Preserve m_asBadLinks(m_lBadLinkIndex)
            m_asBadLinks(m_lBadLinkIndex) = a_oLinkLevelChildren(lIdx).LinkBaseAddress
            '--- log to file
            Put #1, , "+++++++++++++++++++++++++" & vbCrLf
            Put #1, , "Bad link or something" & vbCrLf
            Put #1, , a_oLinkLevelChildren(lIdx).LinkBaseAddress & vbCrLf
            Put #1, , "+++++++++++++++++++++++++" & vbCrLf
        End If
    Next lIdx
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

Private Sub pvAddChild(ByVal oNode As Node, ByVal oTreeView As TreeView, _
                    ByVal sParentKey As String, ByVal sChildKey As String, _
                    ByVal sChildItem As String)
    Const FUNC_NAME As String = "pvAddChild"
    
    On Error GoTo EH
    
    Set oNode = oTreeView.Nodes.Add(sParentKey, tvwChild, sChildKey, ClearString(ClearString2(sChildItem)))
    '---exit
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

Friend Sub frInit(sLink As String, sHTML As String, bIsLinkAbsolute As Boolean, oParentForm As frmMain)
    Const FUNC_NAME As String = "frInit()"
    Dim oLink As New cLink
    Dim oNode As Node
    Dim sParentKey As String, sChildKey As String, sLogFile As String, sLogFileDump As String
    Dim lIdx As Long
    Dim sTemp As String
    
    On Error GoTo EH
    '--- init member variables
    m_sLink = sLink
    m_sHTML = sHTML
    m_bIsLinkAbsolute = bIsLinkAbsolute
    '--- set form caption
    Me.Caption = ClearString(ClearString2(sLink))
    
    '--- define keys
    sParentKey = "ROOT"
    sChildKey = "Child"
    cobBadLinks.Clear
    '--- init link obj
    oLink.frInit m_sLink, m_sHTML
    If oLink.LinkProcessed Then
        '--- add root
        Set oNode = tvNav.Nodes.Add(, , sParentKey, ClearString2(oLink.LinkBaseAddress))
        '--- cleanup if needed
        If IsFilePresent(App.Path & DUMP_LOG_FILE) Then
            Kill App.Path & DUMP_LOG_FILE
        End If
        '--- log file
        sLogFile = App.Path & DUMP_LOG_FILE
        '--- cleanup if needed
        If IsFilePresent(App.Path & DUMP_FILES) Then
            Kill App.Path & DUMP_FILES
        End If
        '--- log file with files
        sLogFileDump = App.Path & DUMP_FILES
        '---open file
        Open sLogFile For Binary Access Write As #1
        Open sLogFileDump For Binary Access Write As #101
        '--- begin log
        Put #1, , "========= LOG DUMP =========" & vbCrLf
        Put #1, , "> Log generated on:" & CStr(Now()) & vbCrLf
        '--- dump files
        For lIdx = 0 To UBound(oLink.LinkArrFiles)
            If m_bIsLinkAbsolute Then
                sTemp = oLink.LinkArrFiles(lIdx) & vbCrLf
            Else
                sTemp = oLink.LinkBaseAddress & oLink.LinkArrFiles(lIdx) & vbCrLf
            End If
            Put #101, , sTemp
        Next lIdx
        '--- traverse child
        pvTraverseChild oLink, sParentKey, sChildKey, oNode, m_bIsLinkAbsolute, sLogFile, sLogFileDump, oParentForm
        '--- close log
        Put #1, , "================================" & vbCrLf
        Put #1, , CStr(Now()) & vbCrLf
        Put #1, , "================================" & vbCrLf
        '--- close file
        Close #1
        Close #101
        '--- fill bad links combo
        If m_lBadLinkIndex <> 0 Then
            For lIdx = 0 To m_lBadLinkIndex
                cobBadLinks.AddItem m_asBadLinks(lIdx)
            Next lIdx
            lblStatus.Caption = MSG_BAD_LINKS & " " & cobBadLinks.ListCount & " " & "items."
        Else
            lblStatus.Caption = MSG_NO_BAD_LINKS
        End If
    Else
        Me.Caption = MSG_LINK_NOT_PROC
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
    
    Me.MousePointer = vbDefault
    cobBadLinks.Text = "Not processed links"
    '--- exit
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
    If Me.Width < 9708 Then
        Me.Width = 9708
    ElseIf Me.Height < 7400 Then
        Me.Height = 7400
    Else
        fraHeader.Width = Width - 250
        fraStruct.Width = fraHeader.Width
        tvNav.Width = Width - 550
        fraStruct.Height = Height - 2000
        tvNav.Height = fraStruct.Height - 650
        lblStatus.Width = fraHeader.Width - 250
        cobBadLinks.Width = lblStatus.Width - 1000
    End If
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

Private Sub tvNav_DblClick()
    Const FUNC_NAME As String = "tvNav_DblClick()"
    Dim oFrmNetDisplay As frmNetDisplay
    
    On Error GoTo EH
    Set oFrmNetDisplay = New frmNetDisplay
    If CBool(chbNavigate.Value) Then
        If Right(tvNav.SelectedItem.Text, 1) = "/" Then
            '--- folder, so display content
            oFrmNetDisplay.frInit tvNav.SelectedItem.Text
        Else
            '--- file, so display it's parent
            oFrmNetDisplay.frInit tvNav.SelectedItem.Parent.Text
        End If
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

Private Sub pvDisplayInfo(ByVal oFrm As frmMain, ByVal sDisplayText As String)
    Const FUNC_NAME As String = "pvDisplayInfo()"
    
    On Error GoTo EH
    '--- todo add code
    lCounter = lCounter + 1
    If lCounter = 8 Then
        oFrm.lblStatus = vbNullString
        lCounter = 0
    End If
    If Not oFrm Is Nothing Then
        oFrm.lblStatus = oFrm.lblStatus & sDisplayText & vbCrLf
        oFrm.Refresh
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
