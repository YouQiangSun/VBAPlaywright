Attribute VB_Name = "TestReporter"
Option Explicit

' ============================================
' TestReporter v2.0（新增）
' 简单的 HTML 测试报告生成器
' 用法：TestReporter.InitReport → TestReporter.LogStep ... → TestReporter.SaveReport
' ============================================

Private mReportPath As String
Private mHtml As String
Private mTotal As Long
Private mPassed As Long
Private mFailed As Long
Private mStartTime As Date

Public Sub InitReport(Optional reportPath As String = "")
    WebDriverCore.InitCore
    If reportPath = "" Then
        reportPath = Environ("TEMP") & "\VBAPlaywright_Report_" & Format(Now, "yyyymmdd_hhnnss") & ".html"
    End If
    mReportPath = reportPath
    mTotal = 0
    mPassed = 0
    mFailed = 0
    mStartTime = Now

    mHtml = "<!DOCTYPE html><html><head><meta charset='utf-8'>" & _
        "<title>VBAPlaywright Test Report</title>" & _
        "<style>" & _
        "body{font-family:Segoe UI,Arial,sans-serif;margin:20px;background:#f5f5f5;}" & _
        "h1{color:#333;}" & _
        ".summary{background:#fff;padding:15px;border-radius:5px;margin-bottom:20px;box-shadow:0 1px 3px rgba(0,0,0,0.1);}" & _
        ".stat{display:inline-block;margin-right:30px;font-size:16px;}" & _
        ".stat-num{font-size:24px;font-weight:bold;}" & _
        ".pass{color:#28a745;}.fail{color:#dc3545;}.total{color:#007bff;}" & _
        "table{width:100%;border-collapse:collapse;background:#fff;box-shadow:0 1px 3px rgba(0,0,0,0.1);}" & _
        "th{background:#007bff;color:#fff;padding:10px;text-align:left;}" & _
        "td{padding:10px;border-bottom:1px solid #eee;}" & _
        "tr.pass{background:#f0fff4;}tr.fail{background:#fff5f5;}" & _
        ".badge{display:inline-block;padding:3px 8px;border-radius:3px;font-size:12px;color:#fff;}" & _
        ".badge-pass{background:#28a745;}.badge-fail{background:#dc3545;}" & _
        "</style></head><body>" & _
        "<h1>VBAPlaywright Test Report</h1>" & _
        "<div class='summary'>" & _
        "<div class='stat'><div>Start Time</div><div class='stat-num'>" & Format(mStartTime, "yyyy-mm-dd hh:nn:ss") & "</div></div>" & _
        "<div class='stat total'><div>Total</div><div class='stat-num' id='total'>0</div></div>" & _
        "<div class='stat pass'><div>Passed</div><div class='stat-num' id='pass'>0</div></div>" & _
        "<div class='stat fail'><div>Failed</div><div class='stat-num' id='fail'>0</div></div>" & _
        "</div><table><thead><tr><th>#</th><th>Status</th><th>Step</th><th>Message</th><th>Duration</th></tr></thead><tbody>"
End Sub

' 记录一步操作
Public Sub LogStep(stepName As String, success As Boolean, Optional message As String = "")
    mTotal = mTotal + 1
    If success Then mPassed = mPassed + 1 Else mFailed = mFailed + 1

    Dim rowClass As String, badgeClass As String, statusText As String
    If success Then
        rowClass = "pass"
        badgeClass = "badge-pass"
        statusText = "PASS"
    Else
        rowClass = "fail"
        badgeClass = "badge-fail"
        statusText = "FAIL"
    End If

    Dim row As String
    row = "<tr class='" & rowClass & "'>" & _
          "<td>" & mTotal & "</td>" & _
          "<td><span class='badge " & badgeClass & "'>" & statusText & "</span></td>" & _
          "<td>" & HtmlEncode(stepName) & "</td>" & _
          "<td>" & HtmlEncode(message) & "</td>" & _
          "<td>-</td></tr>"

    mHtml = mHtml & row
End Sub

' 用例级别记录（带时长）
Public Sub LogCase(caseName As String, success As Boolean, durationMs As Long, Optional message As String = "")
    mTotal = mTotal + 1
    If success Then mPassed = mPassed + 1 Else mFailed = mFailed + 1

    Dim rowClass As String, badgeClass As String, statusText As String
    If success Then
        rowClass = "pass" : badgeClass = "badge-pass" : statusText = "PASS"
    Else
        rowClass = "fail" : badgeClass = "badge-fail" : statusText = "FAIL"
    End If

    Dim row As String
    row = "<tr class='" & rowClass & "'>" & _
          "<td>" & mTotal & "</td>" & _
          "<td><span class='badge " & badgeClass & "'>" & statusText & "</span></td>" & _
          "<td>" & HtmlEncode(caseName) & "</td>" & _
          "<td>" & HtmlEncode(message) & "</td>" & _
          "<td>" & durationMs & " ms</td></tr>"
    mHtml = mHtml & row
End Sub

Public Sub SaveReport()
    If mReportPath = "" Then
        Debug.Print "Report not initialized"
        Exit Sub
    End If
    Dim endTime As Date
    endTime = Now

    mHtml = mHtml & "</tbody></table>" & _
        "<script>" & _
        "document.getElementById('total').innerText=" & mTotal & ";" & _
        "document.getElementById('pass').innerText=" & mPassed & ";" & _
        "document.getElementById('fail').innerText=" & mFailed & ";" & _
        "</script>" & _
        "<p style='margin-top:20px;color:#666;'>End Time: " & Format(endTime, "yyyy-mm-dd hh:nn:ss") & " | Duration: " & Format(endTime - mStartTime, "hh:nn:ss") & "</p>" & _
        "</body></html>"

    Dim ff As Integer
    ff = FreeFile
    Open mReportPath For Output As #ff
    Print #ff, mHtml
    Close #ff
    Debug.Print "Report saved: " & mReportPath
End Sub

' 包装一个 Sub，自动捕获异常并写入报告
Public Sub RunStep(stepName As String, subRef As String)
    On Error GoTo handleErr
    Application.Run subRef
    LogStep stepName, True, ""
    Exit Sub
handleErr:
    LogStep stepName, False, Err.Description
End Sub

Public Property Get ReportPath() As String
    ReportPath = mReportPath
End Property

Public Property Get TotalCount() As Long
    TotalCount = mTotal
End Property

Public Property Get PassedCount() As Long
    PassedCount = mPassed
End Property

Public Property Get FailedCount() As Long
    FailedCount = mFailed
End Property

Private Function HtmlEncode(s As String) As String
    If s = "" Then Exit Function
    Dim r As String
    r = Replace(s, "&", "&amp;")
    r = Replace(r, "<", "&lt;")
    r = Replace(r, ">", "&gt;")
    r = Replace(r, """", "&quot;")
    HtmlEncode = r
End Function
