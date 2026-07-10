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

' ============================================
' Locator v2.0
' 元素定位器：链式调用、批量操作、过滤
' ============================================

Private mDriverUrl As String
Private mSessionId As String
Private mSelector As String
Private mStrategy As String

Friend Sub Init(driverUrl As String, sessionId As String, selector As String, strategy As String)
    mDriverUrl = driverUrl
    mSessionId = sessionId
    mSelector = selector
    mStrategy = strategy
End Sub

Private Function Endpoint(path As String) As String
    Endpoint = mDriverUrl & "/session/" & mSessionId & path
End Function

' 链式：第 N 个元素
Public Function Nth(index As Long) As Locator
    Dim loc As New Locator
    loc.Init mDriverUrl, mSessionId, mSelector & ":nth-of-type(" & index & ")", mStrategy
    Set Nth = loc
End Function

' 链式：包含某文本的元素
Public Function WithText(text As String) As Locator
    Dim loc As New Locator
    loc.Init mDriverUrl, mSessionId, mSelector, mStrategy
    ' 用 JS 过滤
    loc.SelectorOverride = mSelector
    loc.TextFilter = text
    Set WithText = loc
End Sub

Private mSelectorOverride As String
Private mTextFilter As String

Public Property Let SelectorOverride(val As String)
    mSelectorOverride = val
End Property

Public Property Get SelectorOverride() As String
    If mSelectorOverride = "" Then SelectorOverride = mSelector Else SelectorOverride = mSelectorOverride
End Property

Public Property Let TextFilter(val As String)
    mTextFilter = val
End Property

Public Property Get TextFilter() As String
    TextFilter = mTextFilter
End Property

' ============= 元素查找 =============

Private Function FindElement() As String
    Dim body As String
    body = "{""using"":""" & mStrategy & """,""value"":""" & WebDriverCore.JsonEscape(SelectorOverride) & """}"
    Dim resp As String
    resp = WebDriverCore.HttpRequest("POST", Endpoint("/element"), body)
    Dim elemId As String
    elemId = WebDriverCore.JsonExtractElementId(resp)
    If elemId = "" Then
        ' 尝试在多个结果中按文本过滤
        If mTextFilter <> "" Then
            Dim id As String
            id = FindElementByText(mSelector, mTextFilter)
            If id <> "" Then
                FindElement = id
                Exit Function
            End If
        End If
        WebDriverCore.ParseWebDriverError resp
        Err.Raise vbObjectError + 4, "Locator", "Element not found: " & SelectorOverride & " (" & WebDriverCore.LastError & ")"
    End If
    FindElement = elemId
End Function

Private Function FindElementByText(selector As String, text As String) As String
    Dim script As String
    script = "(function(s,t){var els=document.querySelectorAll(s);for(var i=0;i<els.length;i++){if(els[i].innerText&&els[i].innerText.indexOf(t)>=0)return els[i];}return null;})('" & Replace(selector, "'", "\'") & "','" & Replace(text, "'", "\'") & "')"
    Dim body As String
    body = "{""script"":""return " & WebDriverCore.JsonEscape(script) & """,""args"":[]}"
    Dim resp As String
    resp = WebDriverCore.HttpRequest("POST", Endpoint("/execute/sync"), body)
    ' 从返回的 element 提取 elementId
    FindElementByText = WebDriverCore.JsonExtractElementId(resp)
End Function

' ============= 单元素操作 =============

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
    body = "{""text"":""" & WebDriverCore.JsonEscape(text) & """}"
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

Public Function Element() As PWElement
    Dim elemId As String
    elemId = FindElement()
    Dim el As New PWElement
    el.Init mDriverUrl, mSessionId, elemId
    Set Element = el
End Function

' ============= 批量操作 =============

Public Function Count() As Long
    Dim body As String
    body = "{""using"":""" & mStrategy & """,""value"":""" & WebDriverCore.JsonEscape(SelectorOverride) & """}"
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

' 获取所有元素 ID
Public Function AllElementIds() As Collection
    Dim col As New Collection
    Dim body As String
    body = "{""using"":""" & mStrategy & """,""value"":""" & WebDriverCore.JsonEscape(SelectorOverride) & """}"
    Dim resp As String
    resp = WebDriverCore.HttpRequest("POST", Endpoint("/elements"), body)
    Dim pos As Long
    pos = 1
    Do
        Dim keyPos As Long
        keyPos = InStr(pos, resp, """element-6066-11e4-a52e-4f735466cecf""")
        If keyPos = 0 Then
            keyPos = InStr(pos, resp, """ELEMENT""")
            If keyPos = 0 Then Exit Do
        End If
        Dim valStart As Long
        valStart = InStr(keyPos, resp, """")
        If valStart <= keyPos Then Exit Do
        valStart = InStr(valStart + 1, resp, """")
        If valStart = 0 Then Exit Do
        Dim valEnd As Long
        valEnd = InStr(valStart + 1, resp, """")
        If valEnd = 0 Then Exit Do
        col.Add Mid(resp, valStart + 1, valEnd - valStart - 1)
        pos = valEnd + 1
    Loop
    Set AllElementIds = col
End Function

' 获取第 N 个元素
Public Function GetAt(index As Long) As PWElement
    Dim ids As Collection
    Set ids = AllElementIds()
    If index < 1 Or index > ids.Count Then
        Err.Raise vbObjectError + 7, "Locator", "Index out of range: " & index
    End If
    Dim el As New PWElement
    el.Init mDriverUrl, mSessionId, ids(index)
    Set GetAt = el
End Function

' 批量点击所有元素
Public Sub ClickAll()
    Dim ids As Collection
    Set ids = AllElementIds()
    Dim i As Long
    For i = 1 To ids.Count
        WebDriverCore.HttpRequest "POST", Endpoint("/element/" & ids(i) & "/click"), "{}"
    Next i
End Sub

' 批量获取所有元素的文本
Public Function AllText() As Collection
    Dim col As New Collection
    Dim ids As Collection
    Set ids = AllElementIds()
    Dim i As Long
    For i = 1 To ids.Count
        Dim resp As String
        resp = WebDriverCore.HttpRequest("GET", Endpoint("/element/" & ids(i) & "/text"))
        col.Add WebDriverCore.JsonExtract(resp, "value")
    Next i
    Set AllText = col
End Function
