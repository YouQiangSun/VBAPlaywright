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

' ============================================
' Page v2.0
' 页面操作：导航、元素操作、Frame/Alert/拖拽/表格
' ============================================

Private mDriverUrl As String
Private mSessionId As String

Friend Sub Init(driverUrl As String, sessionId As String)
    mDriverUrl = driverUrl
    mSessionId = sessionId
End Sub

Private Function Endpoint(path As String) As String
    Endpoint = mDriverUrl & "/session/" & mSessionId & path
End Function

' ============= 导航 =============

Public Sub GotoUrl(url As String)
    Dim body As String
    body = "{""url"":""" & WebDriverCore.JsonEscape(url) & """}"
    Dim resp As String
    resp = WebDriverCore.HttpRequest("POST", Endpoint("/url"), body)
    If InStr(resp, """error""") > 0 And InStr(resp, "no such") = 0 Then
        WebDriverCore.ParseWebDriverError resp
    End If
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

Public Sub Back()
    WebDriverCore.HttpRequest "POST", Endpoint("/back"), "{}"
End Sub

Public Sub Forward()
    WebDriverCore.HttpRequest "POST", Endpoint("/forward"), "{}"
End Sub

Public Sub Refresh()
    WebDriverCore.HttpRequest "POST", Endpoint("/refresh"), "{}"
End Sub

' ============= 元素查找 =============

Private Function FindElement(selector As String, strategy As String) As String
    Dim body As String
    body = "{""using"":""" & strategy & """,""value"":""" & WebDriverCore.JsonEscape(selector) & """}"
    Dim resp As String
    resp = WebDriverCore.HttpRequest("POST", Endpoint("/element"), body)
    Dim elemId As String
    elemId = WebDriverCore.JsonExtractElementId(resp)
    If elemId = "" Then
        WebDriverCore.ParseWebDriverError resp
        Err.Raise vbObjectError + 4, "Page", "Element not found: " & selector & " (" & WebDriverCore.LastError & ")"
    End If
    FindElement = elemId
End Function

Public Function QuerySelector(selector As String) As PWElement
    Dim elemId As String
    elemId = FindElement(selector, "css selector")
    Dim el As New PWElement
    el.Init mDriverUrl, mSessionId, elemId
    Set QuerySelector = el
End Function

Public Function QueryByXPath(xpath As String) As PWElement
    Dim elemId As String
    elemId = FindElement(xpath, "xpath")
    Dim el As New PWElement
    el.Init mDriverUrl, mSessionId, elemId
    Set QueryByXPath = el
End Function

Public Function Locator(selector As String, Optional strategy As String = "css selector") As Locator
    Dim loc As New Locator
    loc.Init mDriverUrl, mSessionId, selector, strategy
    Set Locator = loc
End Function

' ============= 基础元素操作 =============

Public Sub Click(selector As String)
    Dim elemId As String
    elemId = FindElement(selector, "css selector")
    WebDriverCore.HttpRequest "POST", Endpoint("/element/" & elemId & "/click"), "{}"
End Sub

Public Sub Fill(selector As String, text As String)
    Dim elemId As String
    elemId = FindElement(selector, "css selector")
    WebDriverCore.HttpRequest "POST", Endpoint("/element/" & elemId & "/clear"), "{}"
    Dim body As String
    body = "{""text"":""" & WebDriverCore.JsonEscape(text) & """}"
    WebDriverCore.HttpRequest "POST", Endpoint("/element/" & elemId & "/value"), body
End Sub

Public Sub TypeText(selector As String, text As String)
    Dim elemId As String
    elemId = FindElement(selector, "css selector")
    Dim body As String
    body = "{""text"":""" & WebDriverCore.JsonEscape(text) & """}"
    WebDriverCore.HttpRequest "POST", Endpoint("/element/" & elemId & "/value"), body
End Sub

Public Function TextContent(selector As String) As String
    Dim elemId As String
    elemId = FindElement(selector, "css selector")
    Dim resp As String
    resp = WebDriverCore.HttpRequest("GET", Endpoint("/element/" & elemId & "/text"))
    TextContent = WebDriverCore.JsonExtract(resp, "value")
End Function

Public Function GetAttribute(selector As String, attr As String) As String
    Dim elemId As String
    elemId = FindElement(selector, "css selector")
    Dim resp As String
    resp = WebDriverCore.HttpRequest("GET", Endpoint("/element/" & elemId & "/attribute/" & attr))
    GetAttribute = WebDriverCore.JsonExtract(resp, "value")
End Function

Public Function IsVisible(selector As String) As Boolean
    On Error Resume Next
    Dim elemId As String
    elemId = FindElement(selector, "css selector")
    On Error GoTo 0
    If elemId = "" Then
        IsVisible = False
        Exit Function
    End If
    Dim resp As String
    resp = WebDriverCore.HttpRequest("GET", Endpoint("/element/" & elemId & "/displayed"))
    IsVisible = (WebDriverCore.JsonExtract(resp, "value") = "true")
End Function

' ============= 高级操作 =============

' 执行 JS
Public Function Evaluate(expression As String) As String
    Dim body As String
    body = "{""script"":""return " & WebDriverCore.JsonEscape(expression) & """,""args"":[]}"
    Dim resp As String
    resp = WebDriverCore.HttpRequest("POST", Endpoint("/execute/sync"), body)
    Evaluate = WebDriverCore.JsonExtract(resp, "value")
End Function

' 异步执行 JS（带 await）
Public Function EvaluateAsync(expression As String) As String
    Dim body As String
    body = "{""script"":""return (async function() { " & expression & " })()"",""args"":[]}"
    Dim resp As String
    resp = WebDriverCore.HttpRequest("POST", Endpoint("/execute/async"), body)
    EvaluateAsync = WebDriverCore.JsonExtract(resp, "value")
End Function

' 滚动到元素
Public Sub ScrollToElement(selector As String)
    Evaluate "(function(el){el.scrollIntoView({block:'center'});})(document.querySelector('" & Replace(selector, "'", "\'") & "'))"
End Sub

' 滚动到底部
Public Sub ScrollToBottom()
    Evaluate "(function(){window.scrollTo(0, document.body.scrollHeight);})()"
End Sub

' 滚动到顶部
Public Sub ScrollToTop()
    Evaluate "(function(){window.scrollTo(0, 0);})()"
End Sub

' 拖拽（从一个元素到另一个元素）
Public Sub DragAndDrop(sourceSelector As String, targetSelector As String)
    Dim srcId As String, tgtId As String
    srcId = FindElement(sourceSelector, "css selector")
    tgtId = FindElement(targetSelector, "css selector")
    Dim body As String
    body = "{""actions"":[{""id"":""mouse1"",""type"":""pointer"",""parameters"":{""pointerType"":""mouse""},""actions"":[" & _
           "{""type"":""pointerMove"",""duration"":0,""x"":0,""y"":0,""origin"":{""element-6066-11e4-a52e-4f735466cecf"":""" & srcId & """}}," & _
           "{""type"":""pointerDown"",""button"":0}," & _
           "{""type"":""pointerMove"",""duration"":500,""origin"":{""element-6066-11e4-a52e-4f735466cecf"":""" & tgtId & """}}," & _
           "{""type"":""pointerUp"",""button"":0}]}]}"
    WebDriverCore.HttpRequest "POST", Endpoint("/actions"), body
End Sub

' 鼠标移动到元素
Public Sub Hover(selector As String)
    Dim elemId As String
    elemId = FindElement(selector, "css selector")
    Dim body As String
    body = "{""actions"":[{""id"":""mouse1"",""type"":""pointer"",""parameters"":{""pointerType"":""mouse""},""actions"":[" & _
           "{""type"":""pointerMove"",""duration"":0,""origin"":{""element-6066-11e4-a52e-4f735466cecf"":""" & elemId & """}}]}]}"
    WebDriverCore.HttpRequest "POST", Endpoint("/actions"), body
End Sub

' 键盘按键
Public Sub PressKey(key As String)
    Dim body As String
    body = "{""actions"":[{""id"":""keyboard1"",""type"":""key"",""actions"":[{""type"":""keyDown"",""value"":""" & key & """},{""type"":""keyUp"",""value"":""" & key & """}]}]}"
    WebDriverCore.HttpRequest "POST", Endpoint("/actions"), body
End Sub

' 设置窗口大小
Public Sub SetWindowSize(width As Long, height As Long)
    Dim body As String
    body = "{""width"":" & width & ",""height"":" & height & "}"
    WebDriverCore.HttpRequest "POST", Endpoint("/window/rect"), body
End Sub

' 上传文件
Public Sub UploadFile(selector As String, filePath As String)
    Dim elemId As String
    elemId = FindElement(selector, "css selector")
    Dim body As String
    body = "{""files"":[""" & Replace(filePath, "\", "\\") & """]}"
    WebDriverCore.HttpRequest "POST", Endpoint("/element/" & elemId & "/value"), body
End Sub

' ============= Frame 处理 =============

Public Sub SwitchToFrame(selectorOrId As Variant)
    Dim body As String
    If IsNumeric(selectorOrId) Then
        body = "{""id"":" & CLng(selectorOrId) & "}"
    Else
        body = "{""element"":{""element-6066-11e4-a52e-4f735466cecf"":""" & FindElement(CStr(selectorOrId), "css selector") & """}}"
    End If
    WebDriverCore.HttpRequest "POST", Endpoint("/frame"), body
End Sub

Public Sub SwitchToParentFrame()
    WebDriverCore.HttpRequest "POST", Endpoint("/frame/parent"), "{}"
End Sub

' ============= Alert 处理 =============

Public Sub AcceptAlert()
    WebDriverCore.HttpRequest "POST", Endpoint("/alert/accept"), "{}"
End Sub

Public Sub DismissAlert()
    WebDriverCore.HttpRequest "POST", Endpoint("/alert/dismiss"), "{}"
End Sub

Public Function GetAlertText() As String
    Dim resp As String
    resp = WebDriverCore.HttpRequest("GET", Endpoint("/alert/text"))
    GetAlertText = WebDriverCore.JsonExtract(resp, "value")
End Sub

Public Sub SetAlertText(text As String)
    Dim body As String
    body = "{""text"":""" & WebDriverCore.JsonEscape(text) & """}"
    WebDriverCore.HttpRequest "POST", Endpoint("/alert/text"), body
End Sub

' ============= 表格数据提取 =============

' 提取表格所有内容为二维数组
Public Function ExtractTable(selector As String) As Variant
    Dim script As String
    script = "(function(t){if(!t)return[];var rows=t.rows,r=[];for(var i=0;i<rows.length;i++){var cells=rows[i].cells,c=[];for(var j=0;j<cells.length;j++)c.push(cells[j].innerText.trim());r.push(c);}return r;})(document.querySelector('" & Replace(selector, "'", "\'") & "'))"
    Dim body As String
    body = "{""script"":""return " & WebDriverCore.JsonEscape(script) & """,""args"":[]}"
    Dim resp As String
    resp = WebDriverCore.HttpRequest("POST", Endpoint("/execute/sync"), body)
    ' 解析二维数组（简化处理：返回原始 JSON 字符串给用户解析）
    ExtractTable = Parse2DArray(resp)
End Function

Private Function Parse2DArray(json As String) As Variant
    ' 简化：从 JSON 中提取 value 数组内容
    Dim pos As Long
    pos = InStr(json, """value"":[")
    If pos = 0 Then
        Parse2DArray = Array()
        Exit Function
    End If
    pos = pos + Len("""value"":")
    ' 找到 value 对应的 [ ]
    Dim depth As Long
    depth = 0
    Dim startPos As Long
    startPos = pos
    Dim endPos As Long
    endPos = pos
    Do While endPos <= Len(json)
        Dim c As String
        c = Mid(json, endPos, 1)
        If c = "[" Then depth = depth + 1
        If c = "]" Then
            depth = depth - 1
            If depth = 0 Then Exit Do
        End If
        endPos = endPos + 1
    Loop
    Dim arrJson As String
    arrJson = Mid(json, startPos, endPos - startPos + 1)
    Parse2DArray = arrJson
End Function

' ============= 等待机制 =============

Public Sub WaitForSelector(selector As String, Optional timeout As Long = 30000)
    Dim startTime As Double
    startTime = Timer
    Do While (Timer - startTime) * 1000 < timeout
        On Error Resume Next
        Dim elemId As String
        elemId = FindElement(selector, "css selector")
        On Error GoTo 0
        If elemId <> "" Then Exit Sub
        WebDriverCore.Sleep 200
    Loop
    Err.Raise vbObjectError + 3, "Page", "Timeout waiting for selector: " & selector
End Sub

Public Sub WaitForUrl(url As String, Optional timeout As Long = 30000)
    Dim startTime As Double
    startTime = Timer
    Do While (Timer - startTime) * 1000 < timeout
        If InStr(PageUrl(), url) > 0 Then Exit Sub
        WebDriverCore.Sleep 200
    Loop
    Err.Raise vbObjectError + 5, "Page", "Timeout waiting for URL: " & url
End Sub

Public Sub WaitForText(text As String, Optional timeout As Long = 30000)
    Dim startTime As Double
    startTime = Timer
    Do While (Timer - startTime) * 1000 < timeout
        If InStr(Me.PageSource(), text) > 0 Then Exit Sub
        WebDriverCore.Sleep 200
    Loop
    Err.Raise vbObjectError + 6, "Page", "Timeout waiting for text: " & text
End Sub

Public Sub WaitForTimeout(ms As Long)
    WebDriverCore.Sleep ms
End Sub

Public Function PageSource() As String
    Dim resp As String
    resp = WebDriverCore.HttpRequest("GET", Endpoint("/source"))
    PageSource = WebDriverCore.JsonExtract(resp, "value")
End Function

' ============= 截图 =============

Public Sub Screenshot(filePath As String)
    Dim resp As String
    resp = WebDriverCore.HttpRequest("GET", Endpoint("/screenshot"))
    Dim base64 As String
    base64 = WebDriverCore.JsonExtract(resp, "value")
    WebDriverCore.SaveBase64ToFile base64, filePath
End Sub

' 元素截图
Public Sub ScreenshotElement(selector As String, filePath As String)
    Dim elemId As String
    elemId = FindElement(selector, "css selector")
    Dim resp As String
    resp = WebDriverCore.HttpRequest("GET", Endpoint("/element/" & elemId & "/screenshot"))
    Dim base64 As String
    base64 = WebDriverCore.JsonExtract(resp, "value")
    WebDriverCore.SaveBase64ToFile base64, filePath
End Sub

' ============= 关闭 =============

Public Sub ClosePage()
    WebDriverCore.HttpRequest "DELETE", Endpoint("/window")
End Sub
