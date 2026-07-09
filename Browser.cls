VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Browser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mDriverUrl As String
Private mSessionId As String
Private mBrowserType As BrowserType
Private mPid As Long

Public Sub Init(driverUrl As String, browserType As BrowserType, pid As Long, headless As Boolean, args As String)
    mDriverUrl = driverUrl
    mBrowserType = browserType
    mPid = pid
    
    Dim browserName As String
    Select Case browserType
        Case Chromium
            browserName = "chrome"
        Case Edge
            browserName = "MicrosoftEdge"
        Case Firefox
            browserName = "firefox"
    End Select
    
    Dim capabilities As String
    If browserType = Chromium Or browserType = Edge Then
        Dim chromeArgs As String
        chromeArgs = "["
        If headless Then chromeArgs = chromeArgs & """--headless=new"","
        If args <> "" Then
            Dim parts() As String
            parts = Split(args, " ")
            Dim i As Integer
            For i = LBound(parts) To UBound(parts)
                If parts(i) <> "" Then
                    chromeArgs = chromeArgs & """" & parts(i) & ""","
                End If
            Next i
        End If
        If Right(chromeArgs, 1) = "," Then chromeArgs = Left(chromeArgs, Len(chromeArgs) - 1)
        chromeArgs = chromeArgs & "]"
        
        capabilities = "{""capabilities"":{""alwaysMatch"":{""browserName"":""" & browserName & """,""goog:chromeOptions"":{""args"":" & chromeArgs & "}}}}"
    Else
        capabilities = "{""capabilities"":{""alwaysMatch"":{""browserName"":""" & browserName & """}}}"
    End If
    
    Dim resp As String
    resp = WebDriverCore.HttpRequest("POST", mDriverUrl & "/session", capabilities)
    mSessionId = WebDriverCore.JsonExtract(resp, "sessionId")
    If mSessionId = "" Then
        Err.Raise vbObjectError + 2, "Browser", "Failed to create session: " & resp
    End If
End Sub

Public Function NewPage() As Page
    Dim p As New Page
    p.Init mDriverUrl, mSessionId
    Set NewPage = p
End Function

Public Property Get SessionId() As String
    SessionId = mSessionId
End Property

Public Sub CloseBrowser()
    On Error Resume Next
    WebDriverCore.HttpRequest "DELETE", mDriverUrl & "/session/" & mSessionId
    WebDriverCore.KillProcessById mPid
    On Error GoTo 0
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    CloseBrowser
    On Error GoTo 0
End Sub
