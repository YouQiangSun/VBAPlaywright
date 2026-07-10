VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PWElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ============================================
' PWElement v2.0
' 单元素交互：扩展基本操作
' ============================================

Private mDriverUrl As String
Private mSessionId As String
Private mElementId As String

Friend Sub Init(driverUrl As String, sessionId As String, elementId As String)
    mDriverUrl = driverUrl
    mSessionId = sessionId
    mElementId = elementId
End Sub

Private Function Endpoint(path As String) As String
    Endpoint = mDriverUrl & "/session/" & mSessionId & "/element/" & mElementId & path
End Function

' ============= 基础 =============

Public Sub Click()
    WebDriverCore.HttpRequest "POST", Endpoint("/click"), "{}"
End Sub

Public Sub Fill(text As String)
    WebDriverCore.HttpRequest "POST", Endpoint("/clear"), "{}"
    Dim body As String
    body = "{""text"":""" & WebDriverCore.JsonEscape(text) & """}"
    WebDriverCore.HttpRequest "POST", Endpoint("/value"), body
End Sub

Public Function TextContent() As String
    Dim resp As String
    resp = WebDriverCore.HttpRequest("GET", Endpoint("/text"))
    TextContent = WebDriverCore.JsonExtract(resp, "value")
End Function

Public Function GetAttribute(attr As String) As String
    Dim resp As String
    resp = WebDriverCore.HttpRequest("GET", Endpoint("/attribute/" & attr))
    GetAttribute = WebDriverCore.JsonExtract(resp, "value")
End Function

Public Function IsVisible() As Boolean
    Dim resp As String
    resp = WebDriverCore.HttpRequest("GET", Endpoint("/displayed"))
    IsVisible = (WebDriverCore.JsonExtract(resp, "value") = "true")
End Function

Public Function IsEnabled() As Boolean
    Dim resp As String
    resp = WebDriverCore.HttpRequest("GET", Endpoint("/enabled"))
    IsEnabled = (WebDriverCore.JsonExtract(resp, "value") = "true")
End Function

Public Function IsSelected() As Boolean
    Dim resp As String
    resp = WebDriverCore.HttpRequest("GET", Endpoint("/selected"))
    IsSelected = (WebDriverCore.JsonExtract(resp, "value") = "true")
End Function

' ============= 操作 =============

' 鼠标悬停
Public Sub Hover()
    Dim body As String
    body = "{""actions"":[{""id"":""mouse1"",""type"":""pointer"",""parameters"":{""pointerType"":""mouse""},""actions"":[" & _
           "{""type"":""pointerMove"",""duration"":0,""origin"":{""element-6066-11e4-a52e-4f735466cecf"":""" & mElementId & """}}]}]}"
    WebDriverCore.HttpRequest "POST", mDriverUrl & "/session/" & mSessionId & "/actions", body
End Sub

' 设置复选框/单选框
Public Sub SetChecked(checked As Boolean)
    Dim cur As Boolean
    cur = IsSelected()
    If cur <> checked Then Click
End Sub

' 获取 CSS 属性
Public Function GetCssProperty(prop As String) As String
    Dim resp As String
    resp = WebDriverCore.HttpRequest("GET", Endpoint("/css/" & prop))
    GetCssProperty = WebDriverCore.JsonExtract(resp, "value")
End Function

' 获取内部 HTML
Public Function InnerHTML() As String
    Dim resp As String
    resp = WebDriverCore.HttpRequest("GET", Endpoint("/property/innerHTML"))
    InnerHTML = WebDriverCore.JsonExtract(resp, "value")
End Function

' ============= 等待 =============

' 等待元素可见
Public Sub WaitForVisible(Optional timeout As Long = 30000)
    Dim startTime As Double
    startTime = Timer
    Do While (Timer - startTime) * 1000 < timeout
        If IsVisible() Then Exit Sub
        WebDriverCore.Sleep 200
    Loop
    Err.Raise vbObjectError + 8, "PWElement", "Timeout waiting for element visible"
End Sub

' ============= 截图 =============

Public Sub Screenshot(filePath As String)
    Dim resp As String
    resp = WebDriverCore.HttpRequest("GET", Endpoint("/screenshot"))
    Dim base64 As String
    base64 = WebDriverCore.JsonExtract(resp, "value")
    WebDriverCore.SaveBase64ToFile base64, filePath
End Sub
