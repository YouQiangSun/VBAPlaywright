VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Page"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mDriverUrl As String
Private mSessionId As String

Friend Sub Init(driverUrl As String, sessionId As String)
    mDriverUrl = driverUrl
    mSessionId = sessionId
End Sub

Private Function Endpoint(path As String) As String
    Endpoint = mDriverUrl & "/session/" & mSessionId & path
End Function

Public Sub GotoUrl(url As String)
    Dim body As String
    body = "{""url"":""" & url & """}"
    WebDriverCore.HttpRequest "POST", Endpoint("/url"), body
End Sub

Public Function PageTitle() As String
    Dim resp As String
    resp = WebDriverCore.HttpRequest("GET", Endpoint("/title"))
    PageTitle = WebDriverCore.JsonExtract(resp, "value")
End Function

Public Function PageUrl() As String
    Dim resp As String
    resp = WebDriverCore.HttpRequest("GET", Endpoint("/url"))
    PageUrl = WebDriverCore.JsonExtract(resp, "value")
End Function

Public Sub Click(selector As String)
    Dim elemId As String
    elemId = FindElement(selector)
    WebDriverCore.HttpRequest "POST", Endpoint("/element/" & elemId & "/click"), "{}"
End Sub

Public Sub Fill(selector As String, text As String)
    Dim elemId As String
    elemId = FindElement(selector)
    WebDriverCore.HttpRequest "POST", Endpoint("/element/" & elemId & "/clear"), "{}"
    Dim body As String
    body = "{""text"":""" & text & """}"
    WebDriverCore.HttpRequest "POST", Endpoint("/element/" & elemId & "/value"), body
End Sub

Public Sub TypeText(selector As String, text As String)
    Dim elemId As String
    elemId = FindElement(selector)
    Dim body As String
    body = "{""text"":""" & text & """}"
    WebDriverCore.HttpRequest "POST", Endpoint("/element/" & elemId & "/value"), body
End Sub

Public Function TextContent(selector As String) As String
    Dim elemId As String
    elemId = FindElement(selector)
    Dim resp As String
    resp = WebDriverCore.HttpRequest("GET", Endpoint("/element/" & elemId & "/text"))
    TextContent = WebDriverCore.JsonExtract(resp, "value")
End Function

Public Function GetAttribute(selector As String, attr As String) As String
    Dim elemId As String
    elemId = FindElement(selector)
    Dim resp As String
    resp = WebDriverCore.HttpRequest("GET", Endpoint("/element/" & elemId & "/attribute/" & attr))
    GetAttribute = WebDriverCore.JsonExtract(resp, "value")
End Function

Public Function Evaluate(expression As String) As String
    Dim body As String
    body = "{""script"":""return " & expression & """,""args"":[]}"
    Dim resp As String
    resp = WebDriverCore.HttpRequest("POST", Endpoint("/execute/sync"), body)
    Evaluate = WebDriverCore.JsonExtract(resp, "value")
End Function

Public Sub Screenshot(filePath As String)
    Dim resp As String
    resp = WebDriverCore.HttpRequest("GET", Endpoint("/screenshot"))
    Dim base64 As String
    base64 = WebDriverCore.JsonExtract(resp, "value")
    SaveBase64ToFile base64, filePath
End Sub

Public Function Locator(selector As String) As Locator
    Dim loc As New Locator
    loc.Init mDriverUrl, mSessionId, selector
    Set Locator = loc
End Function

Public Function QuerySelector(selector As String) As PWElement
    Dim elemId As String
    elemId = FindElement(selector)
    Dim el As New PWElement
    el.Init mDriverUrl, mSessionId, elemId
    Set QuerySelector = el
End Function

Public Sub WaitForSelector(selector As String, Optional timeout As Long = 30000)
    Dim startTime As Double
    startTime = Timer
    Do While (Timer - startTime) * 1000 < timeout
        On Error Resume Next
        Dim elemId As String
        elemId = FindElement(selector)
        On Error GoTo 0
        If elemId <> "" Then Exit Sub
        WebDriverCore.Sleep 200
    Loop
    Err.Raise vbObjectError + 3, "Page", "Timeout waiting for selector: " & selector
End Sub

Public Sub WaitForTimeout(ms As Long)
    WebDriverCore.Sleep ms
End Sub

Public Sub ClosePage()
    WebDriverCore.HttpRequest "DELETE", Endpoint("/window")
End Sub

Private Function FindElement(selector As String) As String
    Dim body As String
    body = "{""using"":""css selector"",""value"":""" & selector & """}"
    Dim resp As String
    resp = WebDriverCore.HttpRequest("POST", Endpoint("/element"), body)
    Dim elemId As String
    elemId = WebDriverCore.JsonExtractElementId(resp)
    If elemId = "" Then
        Err.Raise vbObjectError + 4, "Page", "Element not found: " & selector
    End If
    FindElement = elemId
End Function

Private Sub SaveBase64ToFile(base64Str As String, filePath As String)
    Dim xml As Object
    Dim el As Object
    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    Set el = xml.createElement("tmp")
    el.DataType = "bin.base64"
    el.Text = base64Str
    Dim bytes() As Byte
    bytes = el.NodeTypedValue
    Dim ff As Integer
    ff = FreeFile
    Open filePath For Binary As #ff
    Put #ff, , bytes
    Close #ff
End Sub
