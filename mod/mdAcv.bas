Attribute VB_Name = "mdApachecontentViewer"
Option Explicit
Private Const MODULE_NAME           As String = "mdApacheContentViewer"
Private Const USER_AGENT            As String = "Apache Content Viewer v.0.0.1."

Public Function acvGetLink(ByRef sHtmBuffer, ByVal sLinkAddress As String)
    Const FUNC_NAME As String = "acvGetLink"
    Dim sUrl As String, sTempBuf As String, sHTMLCollector As String
    Dim hOpen As Long, hFile As Long, lRet As Long
    Dim lRemaining As Long, lSize As Long
   
    On Error GoTo EH
    '--- get URL
    sUrl = Trim(ClearString2(sLinkAddress))
    '--- Create an internet connection
    hOpen = netInternetOpen(USER_AGENT, INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
    '--- Open the url
    hFile = netInternetOpenUrl(hOpen, sUrl, vbNullString, ByVal 0&, INTERNET_FLAG_RELOAD, ByVal 0&)
    '--- get content
    lRemaining = 1
    While lRemaining <> 0
        netInternetQueryDataAvailable hFile, lRemaining, 0, 0
        lSize = lSize + lRemaining
        sTempBuf = Space(lSize)
        netInternetReadFile hFile, sTempBuf, lSize, lRet
        sHTMLCollector = sHTMLCollector + sTempBuf
    Wend
    '--- clean up
    netInternetCloseHandle hFile
    netInternetCloseHandle hOpen
    sHtmBuffer = Trim(sHTMLCollector)
    Exit Function
EH:
    PrintError MODULE_NAME, FUNC_NAME
End Function
