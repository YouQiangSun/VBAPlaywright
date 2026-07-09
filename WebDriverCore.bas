Attribute VB_Name = "WebDriverCore"
Option Explicit

Private Declare PtrSafe Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare PtrSafe Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Const PROCESS_TERMINATE = &H1

Public Function HttpRequest(method As String, url As String, Optional body As String = "") As String
    Dim xhr As Object
    Set xhr = CreateObject("MSXML2.XMLHTTP.6.0")
    On Error Resume Next
    xhr.Open method, url, False
    xhr.SetRequestHeader "Content-Type", "application/json; charset=utf-8"
    xhr.SetRequestHeader "Accept", "application/json"
    If body <> "" Then
        xhr.Send body
    Else
        xhr.Send
    End If
    HttpRequest = xhr.responseText
    On Error GoTo 0
End Function

Public Function JsonExtract(json As String, key As String) As String
    Dim pattern As String
    pattern = """" & key & """:"""
    Dim pos As Long
    pos = InStr(json, pattern)
    If pos = 0 Then
        pattern = """" & key & """:"
        pos = InStr(json, pattern)
        If pos > 0 Then
            pos = pos + Len(pattern)
            Dim endPos As Long
            endPos = InStr(pos, json, ",")
            If endPos = 0 Then endPos = InStr(pos, json, "}")
            If endPos = 0 Then endPos = Len(json) + 1
            If endPos > 0 Then
                Dim val As String
                val = Trim(Mid(json, pos, endPos - pos))
                val = Replace(val, """", "")
                JsonExtract = val
            End If
        End If
        Exit Function
    End If
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
End Function

Public Function JsonExtractElementId(json As String) As String
    Dim id As String
    id = JsonExtract(json, "element-6066-11e4-a52e-4f735466cecf")
    If id = "" Then id = JsonExtract(json, "ELEMENT")
    JsonExtractElementId = id
End Function

Public Sub KillProcessById(pid As Long)
    Dim hProcess As Long
    hProcess = OpenProcess(PROCESS_TERMINATE, 0, pid)
    If hProcess <> 0 Then
        TerminateProcess hProcess, 0
        CloseHandle hProcess
    End If
End Sub

Public Function WaitForPort(host As String, port As Long, timeoutMs As Long) As Boolean
    Dim startTime As Double
    startTime = Timer
    Do While (Timer - startTime) * 1000 < timeoutMs
        On Error Resume Next
        Dim xhr As Object
        Set xhr = CreateObject("MSXML2.XMLHTTP.6.0")
        xhr.Open "GET", "http://" & host & ":" & port & "/status", False
        xhr.Send
        If xhr.Status = 200 Then
            WaitForPort = True
            Exit Function
        End If
        On Error GoTo 0
        Sleep 200
    Loop
    WaitForPort = False
End Function

Public Sub Sleep(ms As Long)
    Dim startTime As Double
    startTime = Timer
    Do While (Timer - startTime) * 1000 < ms
        DoEvents
    Loop
End Sub
