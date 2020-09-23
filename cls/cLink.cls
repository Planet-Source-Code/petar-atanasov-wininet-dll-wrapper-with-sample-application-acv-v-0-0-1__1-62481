VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'--- Consts
Private Const MODULE_NAME           As String = "cLink"
Private Const G_AHREF_LINK          As String = "< ?A[^\v]*?HREF=""([^\v]*?)""[^\v]*?>"
Private Const G_SRV_FOLDER_TOKENIZER As String = "[a-zA-Z][\s\w\.-]*[a-zA-Z0-9]\/"
Private Const MSG_NOAHREF_FOUND     As String = "No matching link tags found (a href)"
Private Const MSG_OK                As String = "Link processed"
'--- Member variables
Private m_sLinkBaseAddress As String
Private m_asLinkFolders() As String
Private m_asLinkFiles() As String
Private m_asLinkNA() As String
Private m_asLinkMails() As String
Private m_asServerAddressTokenized() As String
Private m_bLinkProcessed As Boolean
Private m_sLinkStatus As String
Private m_sDebugID As String
Private m_bIsLinkAbsolute As Boolean
'--- Properties

Public Property Get IsLinkAbsolute() As String
    IsLinkAbsolute = m_bIsLinkAbsolute
End Property

Public Property Get LinkBaseAddress() As String
    LinkBaseAddress = m_sLinkBaseAddress
End Property

Public Property Get LinkArrFolders() As Variant
    LinkArrFolders = m_asLinkFolders()
End Property

Public Property Get LinkArrFiles() As Variant
    LinkArrFiles = m_asLinkFiles()
End Property

Public Property Get LinkArrNA() As Variant
    LinkArrNA = m_asLinkNA()
End Property

Public Property Get ServerAddressTokenized() As Variant
    ServerAddressTokenized = m_asServerAddressTokenized()
End Property

Public Property Get LinkArrMails() As Variant
    LinkArrMails = m_asLinkMails()
End Property

Public Property Get LinkProcessed() As String
    LinkProcessed = m_bLinkProcessed
End Property

Public Property Get LinkStatus() As String
    LinkStatus = m_sLinkStatus
End Property

'--- Events
Private Sub Class_Initialize()
    DebugInit MODULE_NAME, m_sDebugID
End Sub

Private Sub Class_Terminate()
    DebugTerm MODULE_NAME, m_sDebugID
End Sub

'--- Methods
Friend Sub frInit(ByVal sAddress As String, sHTML As String)
    Const FUNC_NAME As String = "frInit"
    Dim lIdx As Long
    Dim lJdx As Long
    Dim lKdx As Long
    Dim lLdx As Long
    Dim lMdx As Long
    Dim a_sLinks() As String
    Dim a_sChildrenLinkFolders() As String
    Dim a_sChildrenLinkFiles() As String
    Dim a_sChildrenLinkNA() As String
    Dim a_sChildrenLinkMails() As String
    
    On Error GoTo EH
    
    m_sLinkBaseAddress = sAddress
    If rxpPrepareLinkParse(sHTML, a_sLinks()) Then
        '--- classify links
        ReDim Preserve a_sChildrenLinkFiles(UBound(a_sLinks))
        ReDim Preserve a_sChildrenLinkFolders(UBound(a_sLinks))
        ReDim Preserve a_sChildrenLinkMails(UBound(a_sLinks))
        ReDim Preserve a_sChildrenLinkNA(UBound(a_sLinks))
        For lIdx = 0 To UBound(a_sLinks) - 1
            If InStr(a_sLinks(lIdx), "?") Then
               '--- N/A link
                a_sChildrenLinkNA(lJdx) = a_sLinks(lIdx)
                lJdx = lJdx + 1
                '--- If Last symbol is / then folder else file
                ElseIf Len(a_sLinks(lIdx)) = InStrRev(a_sLinks(lIdx), "/") Then
                    '--- folder link
                    a_sChildrenLinkFolders(lKdx) = a_sLinks(lIdx)
                    lKdx = lKdx + 1
                ElseIf InStr(a_sLinks(lIdx), "@") Then
                    '--- mail -> consider storing
                    a_sChildrenLinkMails(lMdx) = a_sLinks(lIdx)
                    lMdx = lMdx + 1
                Else
                    '--- all other are files, todo mail and extra tags parse here or above
                    a_sChildrenLinkFiles(lLdx) = a_sLinks(lIdx)
                    lLdx = lLdx + 1
            End If
        Next
        If lJdx > 0 Then
            ReDim Preserve a_sChildrenLinkNA(lJdx)
        Else
            lJdx = 1
            ReDim Preserve a_sChildrenLinkNA(lJdx)
        End If
        If lKdx > 0 Then
            ReDim Preserve a_sChildrenLinkFolders(lKdx)
        Else
            lKdx = 1
            ReDim Preserve a_sChildrenLinkFolders(lKdx)
        End If
        If lLdx > 0 Then
            ReDim Preserve a_sChildrenLinkFiles(lLdx)
        Else
            lLdx = 1
            ReDim Preserve a_sChildrenLinkFiles(lLdx)
        End If
        If lMdx > 0 Then
            ReDim Preserve a_sChildrenLinkMails(lMdx)
        Else
            lMdx = 1
            ReDim Preserve a_sChildrenLinkMails(lMdx)
        End If
        '--- init class children folders
        ReDim Preserve m_asLinkFolders(UBound(a_sChildrenLinkFolders) - 1)
        For lIdx = 0 To UBound(a_sChildrenLinkFolders) - 1
            m_asLinkFolders(lIdx) = a_sChildrenLinkFolders(lIdx)
        Next
        '--- init class children files
        ReDim Preserve m_asLinkFiles(UBound(a_sChildrenLinkFiles) - 1)
        For lIdx = 0 To UBound(a_sChildrenLinkFiles) - 1
            m_asLinkFiles(lIdx) = a_sChildrenLinkFiles(lIdx)
        Next
        '--- init class children N/A links
        ReDim Preserve m_asLinkNA(UBound(a_sChildrenLinkNA) - 1)
        For lIdx = 0 To UBound(a_sChildrenLinkNA) - 1
            m_asLinkNA(lIdx) = a_sChildrenLinkNA(lIdx)
        Next
        '--- init class children mails
        ReDim Preserve m_asLinkMails(UBound(a_sChildrenLinkMails) - 1)
        For lIdx = 0 To UBound(a_sChildrenLinkMails) - 1
            m_asLinkMails(lIdx) = a_sChildrenLinkMails(lIdx)
        Next
        '--- init class children N/A links
        ReDim Preserve m_asLinkNA(UBound(a_sChildrenLinkNA) - 1)
        For lIdx = 0 To UBound(a_sChildrenLinkNA) - 1
            m_asLinkNA(lIdx) = a_sChildrenLinkNA(lIdx)
        Next
        '--- determine if collected links are absolute or relative
        If UBound(a_sChildrenLinkFolders) > 1 Then
            '--- if the first folder is relative - than all others are too
            m_bIsLinkAbsolute = CBool(InStr(a_sChildrenLinkFolders(1), "http://"))
        ElseIf UBound(a_sChildrenLinkFiles) > 1 Then
            '--- no folders avalilable - same with files
            m_bIsLinkAbsolute = CBool(InStr(a_sChildrenLinkFiles(1), "http://"))
        Else
            '--- empty location - not determined, leave the default value
            m_bIsLinkAbsolute = True
        End If
        m_bLinkProcessed = True
        m_sLinkStatus = MSG_OK
        '--- tokenize address
        rxpTokenizeServerAddress
    Else
        m_bLinkProcessed = False
        m_sLinkStatus = MSG_NOAHREF_FOUND
    End If
    Exit Sub
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Sub

Private Function rxpPrepareLinkParse(sLinkHTML As String, _
                                    asLinks() As String) As Boolean
    Const FUNC_NAME As String = "rxpPrepareLinkParse"
    
    On Error GoTo EH
    
    Dim oRegExp As RegExp
    Dim oMatch As Match
    Dim oMatches As MatchCollection
    Dim sRetValue As String
    Dim lIdx As Long
    Dim lJdx As Long
    Dim lKdx As Long
    Dim lLdx As Long
    Dim lMdx As Long
    Dim lNdx As Long

    '---
    Set oRegExp = New RegExp
    '---
    With oRegExp
        .Pattern = G_AHREF_LINK
        .IgnoreCase = True
        .Global = True
    End With
    '---
    If (oRegExp.Test(sLinkHTML) = True) Then
        Set oMatches = oRegExp.Execute(sLinkHTML)
        lIdx = 0
        ReDim Preserve asLinks(oMatches.Count)
        For Each oMatch In oMatches
            If Not IsEmpty(oMatch) Then
                asLinks(lIdx) = oMatch.SubMatches(0)
                lIdx = lIdx + 1
            End If
        Next
        rxpPrepareLinkParse = True
    Else
        rxpPrepareLinkParse = False
        '--- Regular expression failed !
    End If
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function

Private Sub rxpTokenizeServerAddress()
    Const FUNC_NAME As String = "rxpTokenizeServerAddress"
    
    On Error GoTo EH

    Dim oRegExp      As RegExp
    Dim oMatch       As Match
    Dim oMatches     As MatchCollection
    Dim sRetValue    As String
    Dim lIdx        As Long
    '--- create RegExp
    Set oRegExp = New RegExp
    '--- tune RegExp
    With oRegExp
        .Pattern = G_SRV_FOLDER_TOKENIZER
        .IgnoreCase = True
        .Global = True
    End With
    '--- add / az last char if needed - inache ne e folder i ne go razpoznava
    m_sLinkBaseAddress = m_sLinkBaseAddress
    '--- append /
    If Right(m_sLinkBaseAddress, 1) <> "/" Then
        m_sLinkBaseAddress = m_sLinkBaseAddress & "/"
    End If
    '--- make RegExp
    If (oRegExp.Test(m_sLinkBaseAddress) = True) Then
        Set oMatches = oRegExp.Execute(m_sLinkBaseAddress)
        ReDim Preserve m_asServerAddressTokenized(oMatches.Count - 1)
        For lIdx = 0 To oMatches.Count - 1
            m_asServerAddressTokenized(lIdx) = oMatches.Item(lIdx)
        Next
        '--- Regular expression ok !
'        rxpTokenizeServerAddress = True
    Else
        '--- Regular expression failed !
'        rxpTokenizeServerAddress = True
    End If
    Exit Sub
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Sub
