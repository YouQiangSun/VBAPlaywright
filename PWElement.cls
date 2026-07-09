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

Public Sub Click()
    WebDriverCore.HttpRequest "POST", Endpoint("/click"), "{}"
End Sub

Public Sub Fill(text As String)
    WebDriverCore.HttpRequest "POST", Endpoint("/clear"), "{}"
    Dim body As String
    body = "{""text"":""" & text & """}"
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

Public Sub Screenshot(filePath As String)
    Dim resp As String
    resp = WebDriverCore.HttpRequest("GET", Endpoint("/screenshot"))
    Dim base64 As String
    base64 = WebDriverCore.JsonExtract(resp, "value")
    SaveBase64ToFile base64, filePath
End Sub

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
