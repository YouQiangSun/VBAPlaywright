Attribute VB_Name = "WebDriverCore"
Option Explicit

' ============================================
' WebDriverCore v2.0
' 通用 HTTP 通信、JSON 解析、进程与超时管理
' ============================================

' -------- Win32 API 声明 --------
Private Declare PtrSafe Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare PtrSafe Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Const PROCESS_TERMINATE As Long = &H1

' -------- 全局配置 --------
Public DefaultRequestTimeoutMs As Long   ' 默认 HTTP 请求超时
Public DefaultRetryCount As Long         ' 默认重试次数
Public LastError As String               ' 最近一次错误描述
Public LastErrorCode As String           ' WebDriver 错误码

Public Sub InitCore()
    DefaultRequestTimeoutMs = 30000
    DefaultRetryCount = 2
    LastError = ""
    LastErrorCode = ""
End Sub

' -------- HTTP 请求（带超时与重试） --------
Public Function HttpRequest(method As String, url As String, Optional body As String = "", Optional timeoutMs As Long = -1, Optional retryCount As Long = -1) As String
    If timeoutMs < 0 Then timeoutMs = DefaultRequestTimeoutMs
    If retryCount < 0 Then retryCount = DefaultRetryCount

    Dim attempt As Long
    For attempt = 0 To retryCount
        Dim xhr As Object
        Set xhr = CreateObject("MSXML2.XMLHTTP.6.0")
        On Error Resume Next
        xhr.Open method, url, False
        xhr.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
        xhr.SetRequestHeader "Accept", "application/json"
        xhr.setTimeouts 5000, 5000, timeoutMs, timeoutMs
        If body <> "" Then
            xhr.Send body
        Else
            xhr.Send
        End If
        If Err.Number = 0 Then
            HttpRequest = xhr.responseText
            LastError = ""
            LastErrorCode = ""
            Exit Function
        Else
            LastError = Err.Description
            LastErrorCode = CStr(Err.Number)
            Err.Clear
        End If
        On Error GoTo 0
        Sleep 300
    Next attempt
    HttpRequest = ""
End Function

' -------- JSON 解析 --------
' 解析字符串字段值
Public Function JsonExtract(json As String, key As String) As String
    If Len(json) = 0 Then Exit Function
    Dim pattern As String
    pattern = """" & key & """:"""
    Dim pos As Long
    pos = InStr(json, pattern)
    If pos > 0 Then
        pos = pos + Len(pattern)
        Dim sb As String
        sb = ""
        Do While pos <= Len(json)
            Dim ch As String
            ch = Mid(json, pos, 1)
            If ch = "\" Then
                pos = pos + 1
                If pos <= Len(json) Then
                    sb = sb & Mid(json, pos, 1)
                End If
            ElseIf ch = """" Then
                Exit Do
            Else
                sb = sb & ch
            End If
            pos = pos + 1
        Loop
        JsonExtract = sb
        Exit Function
    End If

    ' 尝试数字 / bool / null
    pattern = """" & key & """:"
    pos = InStr(json, pattern)
    If pos > 0 Then
        pos = pos + Len(pattern)
        Dim endPos As Long
        endPos = pos
        Do While endPos <= Len(json)
            Dim c As String
            c = Mid(json, endPos, 1)
            If c = "," Or c = "}" Or c = "]" Or c = " " Or c = Chr(10) Or c = Chr(13) Then
                Exit Do
            End If
            endPos = endPos + 1
        Loop
        If endPos > pos Then
            JsonExtract = Trim(Mid(json, pos, endPos - pos))
        End If
    End If
End Function

' 解析整数字段
Public Function JsonExtractInt(json As String, key As String) As Long
    Dim s As String
    s = JsonExtract(json, key)
    If IsNumeric(s) Then JsonExtractInt = CLng(s)
End Function

' 解析布尔字段
Public Function JsonExtractBool(json As String, key As String) As Boolean
    JsonExtractBool = (LCase(JsonExtract(json, key)) = "true")
End Function

' 提取 W3C 元素 ID
Public Function JsonExtractElementId(json As String) As String
    Dim id As String
    id = JsonExtract(json, "element-6066-11e4-a52e-4f735466cecf")
    If id = "" Then id = JsonExtract(json, "ELEMENT")
    JsonExtractElementId = id
End Function

' 解析 WebDriver 错误信息
Public Sub ParseWebDriverError(resp As String)
    If Len(resp) = 0 Then
        LastError = "No response from WebDriver"
        LastErrorCode = "no_response"
        Exit Sub
    End If
    LastError = JsonExtract(resp, "message")
    LastErrorCode = JsonExtract(resp, "error")
    If LastError = "" Then
        LastError = resp
    End If
End Sub

' JSON 字符串转义
Public Function JsonEscape(s As String) As String
    Dim r As String
    r = Replace(s, "\", "\\")
    r = Replace(r, """", "\""")
    r = Replace(r, Chr(10), "\n")
    r = Replace(r, Chr(13), "\r")
    r = Replace(r, Chr(9), "\t")
    JsonEscape = r
End Function

' -------- Base64 解码到文件 --------
Public Sub SaveBase64ToFile(base64Str As String, filePath As String)
    On Error Resume Next
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
    On Error GoTo 0
End Sub

' -------- 进程管理 --------
Public Sub KillProcessById(pid As Long)
    On Error Resume Next
    Dim hProcess As Long
    hProcess = OpenProcess(PROCESS_TERMINATE, 0, pid)
    If hProcess <> 0 Then
        TerminateProcess hProcess, 0
        CloseHandle hProcess
    End If
    On Error GoTo 0
End Sub

' 等待端口就绪
Public Function WaitForPort(host As String, port As Long, timeoutMs As Long) As Boolean
    Dim startTime As Double
    startTime = Timer
    Do While (Timer - startTime) * 1000 < timeoutMs
        On Error Resume Next
        Dim xhr As Object
        Set xhr = CreateObject("MSXML2.XMLHTTP.6.0")
        xhr.Open "GET", "http://" & host & ":" & port & "/status", False
        xhr.setTimeouts 1000, 1000, 2000, 2000
        xhr.Send
        If Err.Number = 0 And xhr.Status = 200 Then
            WaitForPort = True
            Exit Function
        End If
        Err.Clear
        On Error GoTo 0
        Sleep 200
    Loop
    WaitForPort = False
End Function

' 非阻塞睡眠
Public Sub Sleep(ms As Long)
    Dim startTime As Double
    startTime = Timer
    Do While (Timer - startTime) * 1000 < ms
        DoEvents
    Loop
End Sub
