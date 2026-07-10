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

' ============================================
' Browser v2.0
' 浏览器会话管理：多标签页、配置、Cookie
' ============================================

Private mDriverUrl As String
Private mSessionId As String
Private mBrowserType As BrowserType
Private mPid As Long
Private mHeadless As Boolean
Private mArgs As String

' -------- 初始化 --------
Public Sub Init(driverUrl As String, browserType As BrowserType, pid As Long, headless As Boolean, args As String)
    mDriverUrl = driverUrl
    mBrowserType = browserType
    mPid = pid
    mHeadless = headless
    mArgs = args

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
        WebDriverCore.ParseWebDriverError resp
        Err.Raise vbObjectError + 2, "Browser", "Failed to create session: " & WebDriverCore.LastError
    End If
End Sub

' -------- 页面管理 --------
Public Function NewPage() As Page
    Dim p As New Page
    p.Init mDriverUrl, mSessionId
    Set NewPage = p
End Function

' 列出所有窗口句柄
Public Function WindowHandles() As Collection
    Dim col As New Collection
    Dim resp As String
    resp = WebDriverCore.HttpRequest("GET", mDriverUrl & "/session/" & mSessionId & "/window/handles")
    ' 简单解析 ["a","b",...]
    Dim pos As Long
    pos = 1
    Do
        Dim p1 As Long, p2 As Long
        p1 = InStr(pos, resp, """")
        If p1 = 0 Then Exit Do
        p2 = InStr(p1 + 1, resp, """")
        If p2 = 0 Then Exit Do
        col.Add Mid(resp, p1 + 1, p2 - p1 - 1)
        pos = p2 + 1
    Loop
    Set WindowHandles = col
End Function

' 切换窗口
Public Sub SwitchToWindow(handle As String)
    Dim body As String
    body = "{""handle"":""" & handle & """}"
    WebDriverCore.HttpRequest "POST", mDriverUrl & "/session/" & mSessionId & "/window", body
End Sub

' 关闭当前窗口
Public Sub CloseWindow()
    WebDriverCore.HttpRequest "DELETE", mDriverUrl & "/session/" & mSessionId & "/window"
End Sub

' 新开标签页
Public Function NewTab(Optional url As String = "about:blank") As Page
    Dim body As String
    body = "{""type"":""tab""}"
    WebDriverCore.HttpRequest "POST", mDriverUrl & "/session/" & mSessionId & "/window/new", body
    Set NewTab = NewPage()
    If url <> "" And url <> "about:blank" Then NewTab.GotoUrl url
End Function

' -------- Cookie 管理 --------
Public Sub SetCookie(name As String, value As String, Optional domain As String = "", Optional path As String = "/")
    Dim body As String
    body = "{""cookie"":{""name"":""" & WebDriverCore.JsonEscape(name) & """,""value"":""" & WebDriverCore.JsonEscape(value) & """"
    If domain <> "" Then body = body & ",""domain"":""" & domain & """"
    If path <> "" Then body = body & ",""path"":""" & path & """"
    body = body & "}}"
    WebDriverCore.HttpRequest "POST", mDriverUrl & "/session/" & mSessionId & "/cookie", body
End Sub

Public Function GetCookie(name As String) As String
    Dim resp As String
    resp = WebDriverCore.HttpRequest("GET", mDriverUrl & "/session/" & mSessionId & "/cookie/" & name)
    GetCookie = WebDriverCore.JsonExtract(resp, "value")
End Function

Public Sub DeleteCookie(name As String)
    WebDriverCore.HttpRequest "DELETE", mDriverUrl & "/session/" & mSessionId & "/cookie/" & name
End Sub

Public Sub DeleteAllCookies()
    WebDriverCore.HttpRequest "DELETE", mDriverUrl & "/session/" & mSessionId & "/cookie"
End Sub

' -------- 配置 --------
Public Property Get SessionId() As String
    SessionId = mSessionId
End Property

Public Property Get BrowserTypeValue() As BrowserType
    BrowserTypeValue = mBrowserType
End Property

Public Property Get IsHeadless() As Boolean
    IsHeadless = mHeadless
End Property

' 关闭浏览器
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
