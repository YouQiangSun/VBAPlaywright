Attribute VB_Name = "WebDriverUpdater"
Option Explicit

' ============================================
' WebDriverUpdater v2.0
' 自动检测/下载/缓存 WebDriver
' 特性：下载重试、版本清单、Edge 内置驱动探测
' ============================================

Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Private Const S_OK = 0
Private Const MAX_DOWNLOAD_RETRY As Long = 3

Public Sub EnsureDriver(browserType As BrowserType, ByRef driverPath As String)
    Dim defaultDir As String
    defaultDir = Environ("LOCALAPPDATA") & "\VBAPlaywright\drivers"

    Dim driverExe As String
    Select Case browserType
        Case Chromium
            driverExe = "chromedriver.exe"
        Case Edge
            driverExe = "msedgedriver.exe"
        Case Firefox
            driverExe = "geckodriver.exe"
    End Select

    If driverPath = "" Then
        driverPath = defaultDir & "\" & driverExe
    End If

    Dim needDownload As Boolean
    needDownload = (Dir(driverPath) = "")

    If Not needDownload And (browserType = Chromium Or browserType = Edge) Then
        Dim browserVer As String
        browserVer = GetBrowserVersion(browserType)
        If browserVer <> "" Then
            Dim driverVer As String
            driverVer = GetDriverVersion(driverPath)
            If driverVer = "" Or Not VersionsCompatible(browserVer, driverVer) Then
                needDownload = True
            End If
        End If
    End If

    If needDownload Then
        MkDirRecursive defaultDir
        Select Case browserType
            Case Chromium
                DownloadChromeDriver defaultDir, driverPath
            Case Edge
                DownloadEdgeDriver defaultDir, driverPath
            Case Firefox
                DownloadFirefoxDriver defaultDir, driverPath
        End Select
    End If
End Sub

Private Sub MkDirRecursive(path As String)
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(path) Then
        fso.CreateFolder path
    End If
    On Error GoTo 0
End Sub

Private Function GetBrowserVersion(browserType As BrowserType) As String
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    Dim regPath As String
    On Error Resume Next
    Select Case browserType
        Case Chromium
            regPath = "HKCU\Software\Google\Chrome\BLBeacon\version"
        Case Edge
            regPath = "HKCU\Software\Microsoft\Edge\BLBeacon\version"
    End Select
    GetBrowserVersion = wsh.RegRead(regPath)
    On Error GoTo 0
End Function

' Edge 新版使用 WebView2 内置驱动，优先探测其位置
Private Function TryEdgeBundledDriver() As String
    On Error Resume Next
    Dim paths As Variant
    paths = Array( _
        Environ("ProgramFiles(x86)") & "\Microsoft\Edge\Application\msedgedriver.exe", _
        Environ("ProgramFiles") & "\Microsoft\Edge\Application\msedgedriver.exe")
    Dim i As Long
    For i = LBound(paths) To UBound(paths)
        If Dir(paths(i)) <> "" Then
            TryEdgeBundledDriver = paths(i)
            Exit Function
        End If
    Next i
    On Error GoTo 0
End Function

Private Function GetDriverVersion(driverPath As String) As String
    On Error Resume Next
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    Dim exec As Object
    Set exec = wsh.Exec("""" & driverPath & """ --version")
    Dim output As String
    If Not exec.StdOut Is Nothing Then
        output = exec.StdOut.ReadAll()
    End If
    GetDriverVersion = ExtractVersion(output)
    On Error GoTo 0
End Function

Private Function ExtractVersion(text As String) As String
    On Error Resume Next
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "\d+\.\d+\.\d+\.\d+"
    regex.Global = False
    If regex.Test(text) Then
        ExtractVersion = regex.Execute(text)(0).Value
    End If
    On Error GoTo 0
End Function

Private Function VersionsCompatible(browserVer As String, driverVer As String) As Boolean
    Dim bMajor As String, dMajor As String
    bMajor = Split(browserVer, ".")(0)
    dMajor = Split(driverVer, ".")(0)
    VersionsCompatible = (bMajor = dMajor)
End Function

' 带重试的下载
Private Function DownloadWithRetry(url As String, target As String) As Boolean
    Dim attempt As Long
    For attempt = 1 To MAX_DOWNLOAD_RETRY
        Dim ret As Long
        ret = URLDownloadToFile(0, url, target, 0, 0)
        If ret = S_OK Then
            DownloadWithRetry = True
            Exit Function
        End If
        WebDriverCore.Sleep 1000
    Next attempt
    DownloadWithRetry = False
End Function

Private Sub DownloadChromeDriver(targetDir As String, driverPath As String)
    Dim browserVer As String
    browserVer = GetBrowserVersion(Chromium)

    Dim majorVer As String
    If browserVer <> "" Then
        majorVer = Split(browserVer, ".")(0)
    Else
        majorVer = "stable"
    End If

    Dim versionUrl As String
    If majorVer = "stable" Or majorVer = "" Then
        versionUrl = "https://googlechromelabs.github.io/chrome-for-testing/LATEST_RELEASE_STABLE"
    Else
        versionUrl = "https://googlechromelabs.github.io/chrome-for-testing/LATEST_RELEASE_" & majorVer
    End If

    Dim xhr As Object
    Set xhr = CreateObject("MSXML2.XMLHTTP.6.0")
    xhr.Open "GET", versionUrl, False
    xhr.send
    Dim driverVer As String
    driverVer = Trim(xhr.responseText)

    Dim downloadUrl As String
    downloadUrl = "https://storage.googleapis.com/chrome-for-testing-public/" & driverVer & "/win64/chromedriver-win64.zip"

    Dim zipPath As String
    zipPath = targetDir & "\chromedriver.zip"

    If Not DownloadWithRetry(downloadUrl, zipPath) Then
        Err.Raise vbObjectError + 10, "WebDriverUpdater", "Failed to download ChromeDriver after retries"
    End If

    ExtractZip zipPath, targetDir

    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim extractedDriver As String
    extractedDriver = targetDir & "\chromedriver-win64\chromedriver.exe"
    If fso.FileExists(extractedDriver) Then
        If fso.FileExists(driverPath) Then fso.DeleteFile driverPath
        fso.MoveFile extractedDriver, driverPath
        If fso.FolderExists(targetDir & "\chromedriver-win64") Then
            fso.DeleteFolder targetDir & "\chromedriver-win64", True
        End If
    End If

    On Error Resume Next
    Kill zipPath
    On Error GoTo 0
End Sub

Private Sub DownloadEdgeDriver(targetDir As String, driverPath As String)
    ' 优先使用 Edge 自带驱动
    Dim bundled As String
    bundled = TryEdgeBundledDriver()
    If bundled <> "" Then
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FileExists(driverPath) Then fso.DeleteFile driverPath
        fso.CopyFile bundled, driverPath
        Exit Sub
    End If

    Dim browserVer As String
    browserVer = GetBrowserVersion(Edge)

    Dim versionUrl As String
    If browserVer <> "" Then
        Dim majorVer As String
        majorVer = Split(browserVer, ".")(0)
        versionUrl = "https://msedgedriver.azureedge.net/LATEST_RELEASE_" & majorVer & "_WINDOWS"
    Else
        versionUrl = "https://msedgedriver.azureedge.net/LATEST_STABLE"
    End If

    Dim xhr As Object
    Set xhr = CreateObject("MSXML2.XMLHTTP.6.0")
    xhr.Open "GET", versionUrl, False
    xhr.send
    Dim driverVer As String
    driverVer = Trim(xhr.responseText)

    Dim downloadUrl As String
    downloadUrl = "https://msedgedriver.azureedge.net/" & driverVer & "/edgedriver_win64.zip"

    Dim zipPath As String
    zipPath = targetDir & "\edgedriver.zip"

    If Not DownloadWithRetry(downloadUrl, zipPath) Then
        Err.Raise vbObjectError + 11, "WebDriverUpdater", "Failed to download EdgeDriver after retries"
    End If

    ExtractZip zipPath, targetDir

    On Error Resume Next
    Kill zipPath
    On Error GoTo 0
End Sub

Private Sub DownloadFirefoxDriver(targetDir As String, driverPath As String)
    Dim downloadUrl As String
    downloadUrl = "https://github.com/mozilla/geckodriver/releases/download/v0.33.0/geckodriver-v0.33.0-win64.zip"

    Dim zipPath As String
    zipPath = targetDir & "\geckodriver.zip"

    If Not DownloadWithRetry(downloadUrl, zipPath) Then
        Err.Raise vbObjectError + 12, "WebDriverUpdater", "Failed to download GeckoDriver after retries"
    End If

    ExtractZip zipPath, targetDir

    On Error Resume Next
    Kill zipPath
    On Error GoTo 0
End Sub

Private Sub ExtractZip(zipPath As String, targetDir As String)
    Dim psCmd As String
    psCmd = "powershell -Command ""Expand-Archive -Path '" & zipPath & "' -DestinationPath '" & targetDir & "' -Force"""
    Shell psCmd, vbHide
    WebDriverCore.Sleep 3000
End Sub
