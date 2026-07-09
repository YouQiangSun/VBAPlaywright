Attribute VB_Name = "Playwright"
Option Explicit

Public Enum BrowserType
    Chromium
    Edge
    Firefox
End Enum

Public Function Launch(browserType As BrowserType, Optional driverPath As String = "", Optional headless As Boolean = False, Optional args As String = "") As Browser
    Dim driverExe As String
    Dim port As Long
    port = 9515
    
    Select Case browserType
        Case Chromium
            driverExe = "chromedriver.exe"
        Case Edge
            driverExe = "msedgedriver.exe"
        Case Firefox
            driverExe = "geckodriver.exe"
            port = 4444
    End Select
    
    ' 自动检查并下载/更新 WebDriver
    Dim actualPath As String
    actualPath = driverPath
    WebDriverUpdater.EnsureDriver browserType, actualPath
    driverExe = actualPath
    
    Dim cmd As String
    cmd = """" & driverExe & """ --port=" & port
    
    Dim pid As Long
    pid = Shell(cmd, vbMinimizedNoFocus)
    
    WebDriverCore.Sleep 1500
    
    If Not WebDriverCore.WaitForPort("localhost", port, 15000) Then
        Err.Raise vbObjectError + 1, "Playwright", "WebDriver failed to start on port " & port
    End If
    
    Dim b As New Browser
    b.Init "http://localhost:" & port, browserType, pid, headless, args
    Set Launch = b
End Function
