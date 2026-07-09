Attribute VB_Name = "Example"
Option Explicit

' ============================================
' VBA Playwright 框架使用示例
' ============================================
' 前提条件：
' 1. 在 VBA 编辑器中导入所有模块和类
' 2. 确保系统已安装对应浏览器（Chrome/Edge/Firefox）
' 3. 框架会自动检查并下载匹配的 WebDriver，无需手动准备
'    - 驱动默认保存路径：%LOCALAPPDATA%\VBAPlaywright\drivers\
'    - 如果驱动已存在但版本不匹配，会自动更新
' ============================================

Sub Demo_BaiduSearch()
    Dim browser As Browser
    ' 启动 Edge 浏览器（非 headless 模式便于观察）
    ' 不指定 driverPath，框架会自动检测浏览器版本并下载对应驱动
    Set browser = Playwright.Launch(Edge, , False)
    
    Dim page As Page
    Set page = browser.NewPage()
    
    ' 访问百度
    page.GotoUrl "https://www.baidu.com"
    Debug.Print "页面标题: " & page.PageTitle()
    
    ' 在搜索框输入内容
    page.Fill "#kw", "VBA 浏览器自动化"
    
    ' 点击搜索按钮
    page.Click "#su"
    
    ' 等待结果加载
    page.WaitForTimeout 3000
    
    ' 截图保存
    page.Screenshot "C:\Temp\baidu_result.png"
    
    ' 使用 Locator 风格 API
    Dim loc As Locator
    Set loc = page.Locator(".result")
    Debug.Print "搜索结果数量: " & loc.Count()
    
    ' 获取第一个搜索结果的标题
    Dim firstResult As PWElement
    Set firstResult = loc.Element()
    Debug.Print "第一条结果文本: " & firstResult.TextContent()
    
    ' 执行 JS
    Debug.Print "页面URL: " & page.Evaluate("document.URL")
    
    ' 关闭浏览器
    browser.CloseBrowser
End Sub

Sub Demo_LoginForm()
    Dim browser As Browser
    Set browser = Playwright.Launch(Chromium, , True) ' headless 模式
    
    Dim page As Page
    Set page = browser.NewPage()
    
    page.GotoUrl "https://the-internet.herokuapp.com/login"
    
    page.Fill "#username", "tomsmith"
    page.Fill "#password", "SuperSecretPassword!"
    page.Click "button[type='submit']"
    
    page.WaitForSelector "#flash"
    Debug.Print "登录结果: " & page.TextContent("#flash")
    
    browser.CloseBrowser
End Sub

Sub Demo_ManualDriverPath()
    ' 如果你希望使用自定义路径的 WebDriver，仍可手动指定
    Dim browser As Browser
    Set browser = Playwright.Launch(Chromium, "C:\Tools\chromedriver.exe", True)
    
    Dim page As Page
    Set page = browser.NewPage()
    
    page.GotoUrl "https://www.example.com"
    Debug.Print page.PageTitle()
    
    browser.CloseBrowser
End Sub
