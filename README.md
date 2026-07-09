# VBAPlaywright

Playwright 风格的 VBA 浏览器自动化框架，底层基于 WebDriver 协议驱动 Chrome / Edge / Firefox。

## 特性

- **Playwright 风格 API** — `Browser` / `Page` / `Locator` / `Element` 四层抽象
- **自动驱动管理** — 自动检测浏览器版本，下载并更新对应 WebDriver
- **零第三方依赖** — 仅使用 VBA 内置对象，无需引用额外库
- **支持多浏览器** — Chrome、Edge、Firefox

## 文件结构

| 文件 | 类型 | 说明 |
|---|---|---|
| `Playwright.bas` | 标准模块 | 入口，`Launch` 函数 |
| `WebDriverCore.bas` | 标准模块 | HTTP 请求、JSON 解析、进程管理 |
| `WebDriverUpdater.bas` | 标准模块 | WebDriver 自动检测与下载更新 |
| `Browser.cls` | 类模块 | 浏览器会话 |
| `Page.cls` | 类模块 | 页面操作 |
| `Locator.cls` | 类模块 | 元素定位器（链式风格） |
| `PWElement.cls` | 类模块 | 单元素交互 |
| `Example.bas` | 标准模块 | 使用示例 |

## 快速开始

1. 在 VBA 编辑器中，**文件 → 导入文件**，导入所有 `.bas` 和 `.cls` 文件
2. 确保系统已安装 Chrome / Edge / Firefox
3. 运行示例：

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

## API 参考

### Browser

```vba
Playwright.Launch(browserType, [driverPath], [headless], [args]) As Browser
browser.NewPage() As Page
browser.CloseBrowser
```

### Page

```vba
page.GotoUrl url
page.PageTitle() As String
page.PageUrl() As String
page.Click selector
page.Fill selector, text
page.TypeText selector, text
page.TextContent(selector) As String
page.GetAttribute(selector, attr) As String
page.Evaluate(expression) As String
page.Screenshot filePath
page.WaitForSelector selector, [timeout]
page.WaitForTimeout ms
page.Locator(selector) As Locator
page.QuerySelector(selector) As PWElement
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
```

### PWElement

```vba
elem.Click
elem.Fill text
elem.TextContent() As String
elem.GetAttribute(attr) As String
elem.IsVisible() As Boolean
elem.Screenshot filePath
```

## WebDriver 自动更新

框架会在启动时自动：

1. 从注册表读取已安装浏览器的版本
2. 检查本地驱动是否存在且版本匹配
3. 不匹配时自动从官方源下载对应版本
4. 默认保存路径：`%LOCALAPPDATA%\VBAPlaywright\drivers\`

如需使用自定义路径，在 `Launch` 时传入第二个参数即可：

```vba
Set browser = Playwright.Launch(Edge, "C:\Tools\msedgedriver.exe", False)
```

## 系统要求

- Windows
- Microsoft Excel / Access（支持 VBA）
- 已安装 Chrome / Edge / Firefox 浏览器
- PowerShell（用于解压 zip，Windows 10+ 自带）
