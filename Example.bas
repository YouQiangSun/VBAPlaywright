Attribute VB_Name = "Example"
Option Explicit

' ============================================
' VBAPlaywright v2.0 使用示例
' ============================================
' 1. 在 VBA 编辑器中导入所有 .bas 和 .cls 文件
' 2. 确保系统已安装 Chrome / Edge / Firefox
' 3. 框架会自动检测/下载匹配的 WebDriver
' ============================================

Sub Demo_BasicSearch()
    Dim browser As Browser
    Set browser = Playwright.Launch(Edge, , False)
    
    Dim page As Page
    Set page = browser.NewPage()
    
    page.GotoUrl "https://www.baidu.com"
    Debug.Print "页面标题: " & page.PageTitle()
    
    page.Fill "#kw", "VBA 浏览器自动化"
    page.Click "#su"
    
    page.WaitForTimeout 2000
    page.Screenshot "C:\Temp\baidu_result.png"
    
    ' Locator 批量操作
    Dim loc As Locator
    Set loc = page.Locator(".result")
    Debug.Print "结果数量: " & loc.Count()
    
    ' 收集所有结果标题
    Dim texts As Collection
    Set texts = loc.AllText()
    Dim i As Long
    For i = 1 To texts.Count
        Debug.Print i & ": " & texts(i)
    Next i
    
    browser.CloseBrowser
End Sub

Sub Demo_AdvancedActions()
    Dim browser As Browser
    Set browser = Playwright.Launch(Chromium, , True)
    
    Dim page As Page
    Set page = browser.NewPage()
    page.SetWindowSize 1280, 800
    
    page.GotoUrl "https://the-internet.herokuapp.com/login"
    
    ' 基础登录
    page.Fill "#username", "tomsmith"
    page.Fill "#password", "SuperSecretPassword!"
    page.Click "button[type='submit']"
    
    page.WaitForUrl "/secure"
    page.WaitForSelector "#flash"
    Debug.Print "登录结果: " & page.TextContent("#flash")
    
    ' 截图
    page.Screenshot "C:\Temp\login_success.png"
    
    ' 提取表格
    Dim tableData As Variant
    tableData = page.ExtractTable("table")
    Debug.Print "表格 JSON: " & tableData
    
    ' 滚动到底部
    page.ScrollToBottom
    page.WaitForTimeout 500
    
    ' 键盘事件
    page.PressKey "Escape"
    
    browser.CloseBrowser
End Sub

Sub Demo_MultiTab()
    Dim browser As Browser
    Set browser = Playwright.Launch(Edge, , True)
    
    Dim page As Page
    Set page = browser.NewPage()
    page.GotoUrl "https://www.baidu.com"
    
    ' 新开标签页
    Dim newTab As Page
    Set newTab = browser.NewTab("https://www.bing.com")
    Debug.Print "新标签页 URL: " & newTab.PageUrl()
    
    ' 列出所有窗口
    Dim handles As Collection
    Set handles = browser.WindowHandles()
    Debug.Print "窗口数: " & handles.Count
    
    ' 切换回第一个标签
    browser.SwitchToWindow handles(1)
    Debug.Print "当前 URL: " & page.PageUrl()
    
    browser.CloseBrowser
End Sub

Sub Demo_FrameAlert()
    Dim browser As Browser
    Set browser = Playwright.Launch(Edge, , True)
    
    Dim page As Page
    Set page = browser.NewPage()
    
    ' 演示 alert 处理
    page.GotoUrl "https://the-internet.herokuapp.com/javascript_alerts"
    page.Click "button[onclick='jsAlert()']"
    page.WaitForTimeout 500
    Debug.Print "Alert 文本: " & page.GetAlertText()
    page.AcceptAlert
    
    ' Frame 切换
    page.GotoUrl "https://the-internet.herokuapp.com/iframe"
    page.SwitchToFrame 0
    Debug.Print "Frame 内文本: " & page.TextContent("p")
    page.SwitchToParentFrame
    
    browser.CloseBrowser
End Sub

Sub Demo_Cookies()
    Dim browser As Browser
    Set browser = Playwright.Launch(Edge, , True)
    
    Dim page As Page
    Set page = browser.NewPage()
    page.GotoUrl "https://www.example.com"
    
    ' 设置 cookie
    browser.SetCookie "session_id", "abc123", ".example.com"
    Debug.Print "Cookie: " & browser.GetCookie("session_id")
    
    ' 清理
    browser.DeleteAllCookies
    
    browser.CloseBrowser
End Sub

Sub Demo_WithReport()
    ' 带测试报告的完整流程
    TestReporter.InitReport
    
    Dim browser As Browser
    Set browser = Playwright.Launch(Edge, , True)
    
    Dim page As Page
    Set page = browser.NewPage()
    
    Dim t As Double
    t = Timer
    page.GotoUrl "https://www.baidu.com"
    TestReporter.LogCase "打开百度", page.PageTitle() <> "", (Timer - t) * 1000
    
    t = Timer
    page.Fill "#kw", "VBA"
    page.Click "#su"
    page.WaitForTimeout 2000
    TestReporter.LogCase "搜索 VBA", page.PageUrl() Like "*wd=VBA*", (Timer - t) * 1000
    
    TestReporter.SaveReport
    Debug.Print "报告: " & TestReporter.ReportPath
    
    browser.CloseBrowser
End Sub

Sub Demo_XPathAndChain()
    Dim browser As Browser
    Set browser = Playwright.Launch(Edge, , True)
    
    Dim page As Page
    Set page = browser.NewPage()
    page.GotoUrl "https://the-internet.herokuapp.com/"
    
    ' 使用 XPath
    Dim link As PWElement
    Set link = page.QueryByXPath("//a[contains(text(),'Login')]")
    link.Click
    
    page.WaitForUrl "/login"
    
    ' 链式 Locator：第 N 个元素 + 文本过滤
    Dim loc As Locator
    Set loc = page.Locator("a")
    Dim el As PWElement
    Set el = loc.GetAt(1)  ' 第二个链接
    Debug.Print "第二个链接: " & el.TextContent()
    
    ' 复选框
    page.GotoUrl "https://the-internet.herokuapp.com/checkboxes"
    Dim cb As PWElement
    Set cb = page.QuerySelector("#checkboxes input")
    cb.SetChecked True
    
    browser.CloseBrowser
End Sub

Sub Demo_ManualDriverPath()
    Dim browser As Browser
    Set browser = Playwright.Launch(Chromium, "C:\Tools\chromedriver.exe", True)
    Dim page As Page
    Set page = browser.NewPage()
    page.GotoUrl "https://www.example.com"
    Debug.Print page.PageTitle()
    browser.CloseBrowser
End Sub
