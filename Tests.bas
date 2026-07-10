Attribute VB_Name = "Tests"
Option Explicit

' ============================================
' VBAPlaywright v2.0 完整测试用例
' ============================================
' 运行方式：在 VBA 编辑器中按 F5 执行对应 Sub
' 注意：部分测试需要网络连接和已安装的浏览器
' ============================================

' -------- 测试报告变量 --------
Private mTestPassed As Long
Private mTestFailed As Long
Private mTestResults As String

' ============================================
' 单元测试入口
' ============================================

Public Sub RunAllTests()
    mTestPassed = 0
    mTestFailed = 0
    mTestResults = ""
    
    Debug.Print String(50, "=")
    Debug.Print "VBAPlaywright v2.0 Test Suite"
    Debug.Print String(50, "=")
    
    ' 浏览器操作测试
    Test_BasicNavigation
    Test_MultiTab
    Test_CookieManagement
    
    ' 元素操作测试
    Test_ElementInteraction
    Test_LocatorChain
    Test_BatchOperations
    
    ' 高级功能测试
    Test_FrameSwitching
    Test_AlertHandling
    Test_JavaScriptExecution
    Test_ScrollAndHover
    
    ' 总结
    Debug.Print String(50, "=")
    Debug.Print "Results: " & mTestPassed & " passed, " & mTestFailed & " failed"
    Debug.Print String(50, "=")
    
    If mTestFailed > 0 Then
        Debug.Print "Failed tests:"
        Debug.Print mTestResults
    End If
End Sub

' ============================================
' 1. 基础导航测试
' ============================================

Public Sub Test_BasicNavigation()
    TestHeader "Test_BasicNavigation"
    
    Dim browser As Browser
    Set browser = Playwright.Launch(Edge, , True)
    
    Dim page As Page
    Set page = browser.NewPage()
    
    On Error GoTo Cleanup
    
    ' 测试打开页面
    page.GotoUrl "https://www.example.com"
    Assert page.PageTitle() = "Example Domain", "Page title should be 'Example Domain'"
    Assert InStr(page.PageUrl(), "example.com") > 0, "URL should contain example.com"
    
    ' 测试刷新
    page.Refresh
    Assert page.PageTitle() = "Example Domain", "Title should remain after refresh"
    
    ' 测试前进后退
    page.GotoUrl "https://www.iana.org/domains/reserved"
    page.Back
    Assert InStr(page.PageUrl(), "example.com") > 0, "Back should return to example.com"
    
    TestPass
    
Cleanup:
    If Err.Number <> 0 Then TestFail Err.Description
    browser.CloseBrowser
End Sub

' ============================================
' 2. 多标签页测试
' ============================================

Public Sub Test_MultiTab()
    TestHeader "Test_MultiTab"
    
    Dim browser As Browser
    Set browser = Playwright.Launch(Edge, , True)
    
    On Error GoTo Cleanup
    
    Dim page1 As Page
    Set page1 = browser.NewPage()
    page1.GotoUrl "https://www.example.com"
    
    ' 新开标签页
    Dim page2 As Page
    Set page2 = browser.NewTab("https://www.iana.org/domains/reserved")
    
    ' 检查窗口数
    Dim handles As Collection
    Set handles = browser.WindowHandles()
    Assert handles.Count >= 2, "Should have at least 2 windows"
    
    ' 切换窗口
    browser.SwitchToWindow handles(1)
    Assert InStr(page1.PageUrl(), "example.com") > 0, "Should be on example.com after switch"
    
    TestPass
    
Cleanup:
    If Err.Number <> 0 Then TestFail Err.Description
    browser.CloseBrowser
End Sub

' ============================================
' 3. Cookie 管理测试
' ============================================

Public Sub Test_CookieManagement()
    TestHeader "Test_CookieManagement"
    
    Dim browser As Browser
    Set browser = Playwright.Launch(Edge, , True)
    
    On Error GoTo Cleanup
    
    Dim page As Page
    Set page = browser.NewPage()
    page.GotoUrl "https://www.example.com"
    
    ' 设置 cookie
    browser.SetCookie "test_cookie", "test_value", "", "/"
    
    ' 验证 cookie 存在（通过 JS）
    Dim cookieVal As String
    cookieVal = page.Evaluate("document.cookie")
    Assert InStr(cookieVal, "test_cookie") > 0, "Cookie should be set"
    
    ' 删除 cookie
    browser.DeleteCookie "test_cookie"
    cookieVal = page.Evaluate("document.cookie")
    Assert InStr(cookieVal, "test_cookie") = 0, "Cookie should be deleted"
    
    ' 删除所有
    browser.SetCookie "c1", "v1", "", "/"
    browser.SetCookie "c2", "v2", "", "/"
    browser.DeleteAllCookies
    cookieVal = page.Evaluate("document.cookie")
    Assert cookieVal = "" Or InStr(cookieVal, "c1") = 0, "All cookies should be cleared"
    
    TestPass
    
Cleanup:
    If Err.Number <> 0 Then TestFail Err.Description
    browser.CloseBrowser
End Sub

' ============================================
' 4. 元素交互测试
' ============================================

Public Sub Test_ElementInteraction()
    TestHeader "Test_ElementInteraction"
    
    Dim browser As Browser
    Set browser = Playwright.Launch(Edge, , True)
    
    On Error GoTo Cleanup
    
    Dim page As Page
    Set page = browser.NewPage()
    page.GotoUrl "https://the-internet.herokuapp.com/login"
    
    ' 输入用户名密码
    page.Fill "#username", "tomsmith"
    page.Fill "#password", "SuperSecretPassword!"
    
    ' 验证输入值
    Assert page.GetAttribute("#username", "value") = "tomsmith", "Username should be filled"
    
    ' 点击登录
    page.Click "button[type='submit']"
    
    ' 等待结果
    page.WaitForSelector "#flash"
    Dim flashText As String
    flashText = page.TextContent("#flash")
    Assert InStr(flashText, "secure area") > 0, "Should login successfully"
    
    ' 测试元素可见性
    Assert page.IsVisible("#flash") = True, "Flash message should be visible"
    Assert page.IsVisible("#nonexistent") = False, "Nonexistent element should not be visible"
    
    TestPass
    
Cleanup:
    If Err.Number <> 0 Then TestFail Err.Description
    browser.CloseBrowser
End Sub

' ============================================
' 5. Locator 链式调用测试
' ============================================

Public Sub Test_LocatorChain()
    TestHeader "Test_LocatorChain"
    
    Dim browser As Browser
    Set browser = Playwright.Launch(Edge, , True)
    
    On Error GoTo Cleanup
    
    Dim page As Page
    Set page = browser.NewPage()
    page.GotoUrl "https://the-internet.herokuapp.com/"
    
    ' 测试 Count
    Dim loc As Locator
    Set loc = page.Locator("a")
    Dim cnt As Long
    cnt = loc.Count()
    Assert cnt > 0, "Should find multiple links"
    
    ' 测试 GetAt
    Dim el As PWElement
    Set el = loc.GetAt(1)
    Assert el.TextContent() <> "", "First link should have text"
    
    ' 测试 Element
    Dim el2 As PWElement
    Set el2 = loc.Element()
    Assert el2.TextContent() <> "", "Element() should return first element"
    
    ' 测试 QuerySelector 返回 PWElement
    Dim link As PWElement
    Set link = page.QuerySelector("a[href='/login']")
    Assert InStr(link.GetAttribute("href"), "login") > 0, "QuerySelector should find login link"
    link.Click
    
    page.WaitForUrl "/login"
    Assert InStr(page.PageUrl(), "login") > 0, "Should navigate to login page"
    
    TestPass
    
Cleanup:
    If Err.Number <> 0 Then TestFail Err.Description
    browser.CloseBrowser
End Sub

' ============================================
' 6. 批量操作测试
' ============================================

Public Sub Test_BatchOperations()
    TestHeader "Test_BatchOperations"
    
    Dim browser As Browser
    Set browser = Playwright.Launch(Edge, , True)
    
    On Error GoTo Cleanup
    
    Dim page As Page
    Set page = browser.NewPage()
    page.GotoUrl "https://the-internet.herokuapp.com/checkboxes"
    
    ' 测试 AllText
    Dim texts As Collection
    Set texts = page.Locator("label").AllText()
    Assert texts.Count >= 2, "Should find checkbox labels"
    
    ' 测试批量获取元素
    Dim ids As Collection
    Set ids = page.Locator("input[type='checkbox']").AllElementIds()
    Assert ids.Count >= 2, "Should find checkboxes"
    
    ' 测试 XPath
    Dim cb As PWElement
    Set cb = page.QueryByXPath("//input[@type='checkbox'][1]")
    cb.SetChecked True
    Assert cb.IsSelected() = True, "First checkbox should be checked"
    
    TestPass
    
Cleanup:
    If Err.Number <> 0 Then TestFail Err.Description
    browser.CloseBrowser
End Sub

' ============================================
' 7. Frame 切换测试
' ============================================

Public Sub Test_FrameSwitching()
    TestHeader "Test_FrameSwitching"
    
    Dim browser As Browser
    Set browser = Playwright.Launch(Edge, , True)
    
    On Error GoTo Cleanup
    
    Dim page As Page
    Set page = browser.NewPage()
    page.GotoUrl "https://the-internet.herokuapp.com/iframe"
    
    ' 切换到 Frame（通过索引）
    page.SwitchToFrame 0
    
    ' 在 Frame 内操作
    Dim frameText As String
    frameText = page.TextContent("p")
    Assert frameText <> "", "Should find text inside iframe"
    
    ' 切回父 Frame
    page.SwitchToParentFrame
    
    ' 父级元素应可访问
    Assert page.PageTitle() <> "", "Should be back in parent frame"
    
    TestPass
    
Cleanup:
    If Err.Number <> 0 Then TestFail Err.Description
    browser.CloseBrowser
End Sub

' ============================================
' 8. Alert 处理测试
' ============================================

Public Sub Test_AlertHandling()
    TestHeader "Test_AlertHandling"
    
    Dim browser As Browser
    Set browser = Playwright.Launch(Edge, , True)
    
    On Error GoTo Cleanup
    
    Dim page As Page
    Set page = browser.NewPage()
    page.GotoUrl "https://the-internet.herokuapp.com/javascript_alerts"
    
    ' 测试 JS Alert
    page.Click "button[onclick='jsAlert()']"
    page.WaitForTimeout 500
    
    Dim alertText As String
    alertText = page.GetAlertText()
    Assert alertText <> "", "Alert should have text"
    page.AcceptAlert
    
    ' 验证结果
    Dim result As String
    result = page.TextContent("#result")
    Assert InStr(result, "successfully clicked") > 0, "Should show success message"
    
    ' 测试 Confirm
    page.Click "button[onclick='jsConfirm()']"
    page.WaitForTimeout 500
    page.DismissAlert
    result = page.TextContent("#result")
    Assert InStr(result, "Cancel") > 0, "Should show cancel message"
    
    ' 测试 Prompt
    page.Click "button[onclick='jsPrompt()']"
    page.WaitForTimeout 500
    page.SetAlertText "Hello VBA"
    page.AcceptAlert
    result = page.TextContent("#result")
    Assert InStr(result, "Hello VBA") > 0, "Should show entered text"
    
    TestPass
    
Cleanup:
    If Err.Number <> 0 Then TestFail Err.Description
    browser.CloseBrowser
End Sub

' ============================================
' 9. JavaScript 执行测试
' ============================================

Public Sub Test_JavaScriptExecution()
    TestHeader "Test_JavaScriptExecution"
    
    Dim browser As Browser
    Set browser = Playwright.Launch(Edge, , True)
    
    On Error GoTo Cleanup
    
    Dim page As Page
    Set page = browser.NewPage()
    page.GotoUrl "https://www.example.com"
    
    ' 同步执行 JS
    Dim url As String
    url = page.Evaluate("document.URL")
    Assert InStr(url, "example.com") > 0, "Evaluate should return current URL"
    
    ' 获取标题
    Dim title As String
    title = page.Evaluate("document.title")
    Assert title = "Example Domain", "Evaluate should return title"
    
    ' 修改页面内容
    page.Evaluate "document.body.style.backgroundColor = 'yellow'"
    page.WaitForTimeout 100
    
    ' 验证修改
    Dim bg As String
    bg = page.Evaluate("document.body.style.backgroundColor")
    Assert bg <> "", "Background should be modified"
    
    TestPass
    
Cleanup:
    If Err.Number <> 0 Then TestFail Err.Description
    browser.CloseBrowser
End Sub

' ============================================
' 10. 滚动与悬停测试
' ============================================

Public Sub Test_ScrollAndHover()
    TestHeader "Test_ScrollAndHover"
    
    Dim browser As Browser
    Set browser = Playwright.Launch(Edge, , True)
    
    On Error GoTo Cleanup
    
    Dim page As Page
    Set page = browser.NewPage()
    page.GotoUrl "https://the-internet.herokuapp.com/hovers"
    
    ' 测试 Hover
    page.Hover "div.figure:nth-of-type(1)"
    page.WaitForTimeout 500
    
    Dim caption As String
    caption = page.TextContent("div.figure:nth-of-type(1) .figcaption h5")
    Assert caption <> "", "Hover should reveal caption"
    
    ' 测试滚动到底部（需要长页面）
    page.GotoUrl "https://the-internet.herokuapp.com/infinite_scroll"
    page.WaitForTimeout 1000
    page.ScrollToBottom
    page.WaitForTimeout 500
    
    TestPass
    
Cleanup:
    If Err.Number <> 0 Then TestFail Err.Description
    browser.CloseBrowser
End Sub

' ============================================
' 11. 截图测试
' ============================================

Public Sub Test_Screenshot()
    TestHeader "Test_Screenshot"
    
    Dim browser As Browser
    Set browser = Playwright.Launch(Edge, , True)
    
    On Error GoTo Cleanup
    
    Dim page As Page
    Set page = browser.NewPage()
    page.GotoUrl "https://www.example.com"
    
    ' 页面截图
    page.Screenshot "C:\Temp\test_page.png"
    Assert Dir("C:\Temp\test_page.png") <> "", "Page screenshot should be saved"
    
    ' 元素截图
    page.ScreenshotElement "h1", "C:\Temp\test_element.png"
    Assert Dir("C:\Temp\test_element.png") <> "", "Element screenshot should be saved"
    
    ' 清理
    On Error Resume Next
    Kill "C:\Temp\test_page.png"
    Kill "C:\Temp\test_element.png"
    On Error GoTo Cleanup
    
    TestPass
    
Cleanup:
    If Err.Number <> 0 Then TestFail Err.Description
    browser.CloseBrowser
End Sub

' ============================================
' 12. 带报告的完整流程测试
' ============================================

Public Sub Test_WithReporter()
    TestHeader "Test_WithReporter"
    
    TestReporter.InitReport "C:\Temp\vba_playwright_test_report.html"
    
    Dim browser As Browser
    Set browser = Playwright.Launch(Edge, , True)
    
    On Error GoTo Cleanup
    
    Dim page As Page
    Set page = browser.NewPage()
    
    ' 用例 1
    Dim t As Double
    t = Timer
    page.GotoUrl "https://www.example.com"
    TestReporter.LogCase "Open Example Domain", page.PageTitle() = "Example Domain", (Timer - t) * 1000
    
    ' 用例 2
    t = Timer
    page.GotoUrl "https://the-internet.herokuapp.com/login"
    page.Fill "#username", "tomsmith"
    page.Fill "#password", "SuperSecretPassword!"
    page.Click "button[type='submit']"
    page.WaitForSelector "#flash"
    Dim success As Boolean
    success = InStr(page.TextContent("#flash"), "secure area") > 0
    TestReporter.LogCase "Login Test", success, (Timer - t) * 1000
    
    TestReporter.SaveReport
    
    Debug.Print "Report saved to: " & TestReporter.ReportPath
    Assert TestReporter.PassedCount >= 1, "At least one test should pass"
    
    TestPass
    
Cleanup:
    If Err.Number <> 0 Then TestFail Err.Description
    browser.CloseBrowser
End Sub

' ============================================
' 断言辅助函数
' ============================================

Private Sub Assert(condition As Boolean, message As String)
    If Not condition Then
        Err.Raise vbObjectError + 100, "Tests", "ASSERT FAILED: " & message
    End If
End Sub

Private Sub TestHeader(name As String)
    Debug.Print "Running: " & name & " ..."
End Sub

Private Sub TestPass()
    mTestPassed = mTestPassed + 1
    Debug.Print "  [PASS]"
End Sub

Private Sub TestFail(msg As String)
    mTestFailed = mTestFailed + 1
    mTestResults = mTestResults & "- " & msg & vbCrLf
    Debug.Print "  [FAIL] " & msg
    Err.Clear
End Sub
