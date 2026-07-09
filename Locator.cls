VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Locator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mDriverUrl As String
Private mSessionId As String
Private mSelector As String

Friend Sub Init(driverUrl As String, sessionId As String, selector As String)
    mDriverUrl = driverUrl
    mSessionId = sessionId
    mSelector = selector
End Sub

Private Function Endpoint(path As String) As String
    Endpoint = mDriverUrl & "/session/" & mSessionId & path
End Function

Public Sub Click()
    Dim elemId As String
    elemId = FindElement()
    WebDriverCore.HttpRequest "POST", Endpoint("/element/" & elemId & "/click"), "{}"
End Sub

Public Sub Fill(text As String)
    Dim elemId As String
    elemId = FindElement()
    WebDriverCore.HttpRequest "POST", Endpoint("/element/" & elemId & "/clear"), "{}"
    Dim body As String
    body = "{""text"":""" & text & """}"
    WebDriverCore.HttpRequest "POST", Endpoint("/element/" & elemId & "/value"), body
End Sub

Public Function TextContent() As String
    Dim elemId As String
    elemId = FindElement()
    Dim resp As String
    resp = WebDriverCore.HttpRequest("GET", Endpoint("/element/" & elemId & "/text"))
    TextContent = WebDriverCore.JsonExtract(resp, "value")
End Function

Public Function GetAttribute(attr As String) As String
    Dim elemId As String
    elemId = FindElement()
    Dim resp As String
    resp = WebDriverCore.HttpRequest("GET", Endpoint("/element/" & elemId & "/attribute/" & attr))
    GetAttribute = WebDriverCore.JsonExtract(resp, "value")
End Function

Public Function IsVisible() As Boolean
    On Error Resume Next
    Dim elemId As String
    elemId = FindElement()
    On Error GoTo 0
    IsVisible = (elemId <> "")
End Function

Public Function Count() As Long
    Dim body As String
    body = "{""using"":""css selector"",""value"":""" & mSelector & """}"
    Dim resp As String
    resp = WebDriverCore.HttpRequest("POST", Endpoint("/elements"), body)
    Dim cnt As Long
    Dim pos As Long
    pos = 1
    Do
        Dim p1 As Long, p2 As Long
        p1 = InStr(pos, resp, """element-6066-11e4-a52e-4f735466cecf""")
        p2 = InStr(pos, resp, """ELEMENT""")
        If p1 = 0 And p2 = 0 Then Exit Do
        If p1 > 0 And (p2 = 0 Or p1 < p2) Then
            pos = p1 + 1
        Else
            pos = p2 + 1
        End If
        cnt = cnt + 1
    Loop While pos > 1
    Count = cnt
End Function

Public Function Element() As PWElement
    Dim elemId As String
    elemId = FindElement()
    Dim el As New PWElement
    el.Init mDriverUrl, mSessionId, elemId
    Set Element = el
End Function

Private Function FindElement() As String
    Dim body As String
    body = "{""using"":""css selector"",""value"":""" & mSelector & """}"
    Dim resp As String
    resp = WebDriverCore.HttpRequest("POST", Endpoint("/element"), body)
    Dim elemId As String
    elemId = WebDriverCore.JsonExtractElementId(resp)
    If elemId = "" Then
        Err.Raise vbObjectError + 4, "Locator", "Element not found: " & mSelector
    End If
    FindElement = elemId
End Function
