# VBAPlaywright v2.0

Playwright 风格的 VBA 浏览器自动化框架，底层基于 WebDriver 协议驱动 Chrome / Edge / Firefox。

## v2.0 新增功能

### 功能增强
- **多标签页管理**（`Browser.NewTab` / `SwitchToWindow` / `WindowHandles`）
- **XPath 支持**（`Page.QueryByXPath`）
- **Frame 切换**（`Page.SwitchToFrame` / `SwitchToParentFrame`）
- **Alert 处理**（`AcceptAlert` / `DismissAlert` / `GetAlertText` / `SetAlertText`）
- **拖拽与高级交互**（`DragAndDrop` / `Hover` / `PressKey`）
- **滚动控制**（`ScrollToElement` / `ScrollToBottom` / `ScrollToTop`）
- **表格数据提取**（`ExtractTable`）
- **文件上传**（`UploadFile`）
- **元素级 WaitForVisible**（`PWElement.WaitForVisible`）
- **CSS 属性读取**（`GetCssProperty`）
- **Locator 文本过滤**（`WithText`）

### 稳定性提升
- HTTP 请求**超时控制**（`setTimeouts`）与**自动重试**
- **WebDriver 错误信息**解析（`WebDriverCore.ParseWebDriverError`）
- 改进 JSON 解析：支持数字、布尔、转义
- **Edge 内置驱动**自动探测（WebView2 路径）
- **下载重试机制**（最多 3 次）

### 易用性提升
- 链式 Locator：`Locator.Nth()` / `Locator.WithText()` / `Locator.GetAt(i)`
- 批量操作：`ClickAll` / `AllText` / `AllElementIds`
- 测试报告生成器 `TestReporter.bas`：自动生成 HTML 报告
- 全部 JSON 字符串自动转义
- 用例级日志：`TestReporter.LogCase` 自动记录耗时

## 文件结构

| 文件 | 类型 | 说明 |
|---|---|---|
| `Playwright.bas` | 标准模块 | 入口 |
| `WebDriverCore.bas` | 标准模块 | HTTP / JSON / 进程管理 |
| `WebDriverUpdater.bas` | 标准模块 | 驱动自动下载 |
| `Browser.cls` | 类模块 | 浏览器会话 + 多标签 + Cookie |
| `Page.cls` | 类模块 | 页面操作（导航/Frame/Alert/拖拽等） |
| `Locator.cls` | 类模块 | 元素定位器（链式 + 批量） |
| `PWElement.cls` | 类模块 | 单元素交互 |
| `TestReporter.bas` | 标准模块 | **新增** — HTML 测试报告 |
| `Example.bas` | 标准模块 | 使用示例 |

## 快速开始

```vba
Sub Demo()
    Dim browser As Browser
    Set browser = Playwright.Launch(Edge, , False)  ' 自动下载驱动
    
    Dim page As Page
    Set page = browser.NewPage()
    
    page.GotoUrl "https://www.baidu.com"
    page.Fill "#kw", "VBA 自动化"
    page.Click "#su"
    
    Debug.Print page.PageTitle()
    
    browser.CloseBrowser
End Sub
```

## API 参考（v2.0）

### Browser
```vba
Playwright.Launch(browserType, [driverPath], [headless], [args]) As Browser
browser.NewPage() As Page
browser.NewTab([url]) As Page
browser.WindowHandles() As Collection
browser.SwitchToWindow handle
browser.CloseWindow
browser.SetCookie name, value, [domain], [path]
browser.GetCookie(name) As String
browser.DeleteCookie name
browser.DeleteAllCookies
browser.CloseBrowser
```

### Page
```vba
' 导航
page.GotoUrl url
page.Back
page.Forward
page.Refresh
page.PageTitle() As String
page.PageUrl() As String
page.PageSource() As String

' 元素查找
page.QuerySelector(selector) As PWElement
page.QueryByXPath(xpath) As PWElement
page.Locator(selector, [strategy]) As Locator

' 基础操作
page.Click selector
page.Fill selector, text
page.TypeText selector, text
page.TextContent(selector) As String
page.GetAttribute(selector, attr) As String
page.IsVisible(selector) As Boolean

' 高级
page.Evaluate(jsExpression) As String
page.EvaluateAsync(jsExpression) As String
page.Hover selector
page.DragAndDrop sourceSelector, targetSelector
page.PressKey key
page.ScrollToElement selector
page.ScrollToBottom
page.ScrollToTop
page.SetWindowSize w, h
page.UploadFile selector, filePath

' Frame / Alert
page.SwitchToFrame selectorOrIndex
page.SwitchToParentFrame
page.AcceptAlert
page.DismissAlert
page.GetAlertText() As String
page.SetAlertText text

' 表格
page.ExtractTable(selector) As Variant

' 等待
page.WaitForSelector selector, [timeout]
page.WaitForUrl url, [timeout]
page.WaitForText text, [timeout]
page.WaitForTimeout ms

' 截图
page.Screenshot filePath
page.ScreenshotElement selector, filePath
```

### Locator
```vba
loc.Click
loc.Fill text
loc.TextContent() As String
loc.GetAttribute(attr) As String
loc.IsVisible() As Boolean
loc.Count() As Long
loc.Element() As PWElement

' 链式 / 批量
loc.Nth(index) As Locator
loc.WithText(text) As Locator
loc.GetAt(index) As PWElement
loc.AllElementIds() As Collection
loc.AllText() As Collection
loc.ClickAll
```

### PWElement
```vba
elem.Click
elem.Fill text
elem.Hover
elem.SetChecked checked
elem.TextContent() As String
elem.GetAttribute(attr) As String
elem.GetCssProperty(prop) As String
elem.InnerHTML() As String
elem.IsVisible() As Boolean
elem.IsEnabled() As Boolean
elem.IsSelected() As Boolean
elem.WaitForVisible [timeout]
elem.Screenshot filePath
```

### TestReporter
```vba
TestReporter.InitReport [reportPath]
TestReporter.LogStep stepName, success, [message]
TestReporter.LogCase caseName, success, durationMs, [message]
TestReporter.SaveReport
```

## 系统要求

- Windows
- Microsoft Excel / Access（支持 VBA）
- 已安装 Chrome / Edge / Firefox
- PowerShell（用于解压 zip）
