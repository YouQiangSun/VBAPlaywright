Attribute VB_Name = "Installer"
Option Explicit

' ============================================
' VBAPlaywright 一键安装器
' ============================================
' 功能：自动将所有 .bas / .cls 文件导入当前工作簿
' 前置条件：
'   1. 需要引用 "Microsoft Visual Basic for Applications Extensibility 5.3"
'   2. Excel 需开启 "信任对 VBA 项目对象模型的访问"
'      路径：文件 → 选项 → 信任中心 → 信任中心设置 → 宏设置
'            → 勾选 "信任对 VBA 项目对象模型的访问"
' ============================================

' -------- 常量 --------
Private Const VBIDE_GUID As String = "{0002E157-0000-0000-C000-000000000046}"
Private Const VBIDE_NAME As String = "VBIDE"
Private Const VBIDE_MAJOR As Long = 5
Private Const VBIDE_MINOR As Long = 3

' ============================================
' 主入口：从文件夹一键安装
' ============================================

Public Sub InstallFromFolder()
    On Error GoTo HandleError
    
    ' 1. 检查/添加 VBIDE 引用
    If Not EnsureVBIDEReference Then
        MsgBox "无法添加 VBIDE 引用，安装中止。" & vbCrLf & _
               "请手动添加引用：工具 → 引用 → 勾选 'Microsoft Visual Basic for Applications Extensibility 5.3'", vbCritical
        Exit Sub
    End If
    
    ' 2. 选择代码文件夹
    Dim folderPath As String
    folderPath = SelectFolder("请选择 VBAPlaywright 代码文件夹")
    If folderPath = "" Then
        MsgBox "未选择文件夹，安装已取消。", vbInformation
        Exit Sub
    End If
    
    ' 3. 确认当前工作簿
    Dim wb As Workbook
    Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "请先打开或新建一个 Excel 工作簿！", vbExclamation
        Exit Sub
    End If
    
    ' 4. 执行导入
    Dim result As String
    result = ImportAllComponents(wb, folderPath)
    
    ' 5. 提示保存为 xlsm
    MsgBox result & vbCrLf & vbCrLf & _
           "建议立即将工作簿另存为 .xlsm 格式！" & vbCrLf & _
           "按 F12 → 保存类型选择 'Excel 启用宏的工作簿 (*.xlsm)'", vbInformation, "安装完成"
    
    Exit Sub
    
HandleError:
    MsgBox "安装出错: " & Err.Description & vbCrLf & "错误号: " & Err.Number, vbCritical
End Sub

' ============================================
' 智能安装：自动从 GitHub 下载最新版并导入
' ============================================

Public Sub InstallFromGitHub()
    On Error GoTo HandleError
    
    Dim resp As String
    resp = DownloadText("https://raw.githubusercontent.com/YouQiangSun/VBAPlaywright/main/README.md")
    If resp = "" Then
        MsgBox "无法连接 GitHub，请检查网络或手动下载代码。", vbExclamation
        Exit Sub
    End If
    
    MsgBox "检测到网络连接正常。" & vbCrLf & _
           "由于 VBA 限制，无法直接解压 zip。" & vbCrLf & _
           "请手动下载 https://github.com/YouQiangSun/VBAPlaywright/archive/refs/heads/main.zip 并解压，" & vbCrLf & _
           "然后运行 InstallFromFolder 选择解压后的文件夹。", vbInformation
    
    Exit Sub
    
HandleError:
    MsgBox "出错: " & Err.Description, vbCritical
End Sub

' ============================================
' 核心：导入所有组件
' ============================================

Private Function ImportAllComponents(wb As Workbook, folderPath As String) As String
    Dim vbProj As Object
    Set vbProj = wb.VBProject
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    Dim folder As Object
    Set folder = fso.GetFolder(folderPath)
    
    Dim imported As Long, skipped As Long, failed As Long
    imported = 0: skipped = 0: failed = 0
    
    Dim file As Object
    For Each file In folder.Files
        Dim ext As String
        ext = LCase(fso.GetExtensionName(file.Name))
        
        If ext = "bas" Or ext = "cls" Then
            Dim compName As String
            compName = fso.GetBaseName(file.Name)
            
            ' 检查是否已存在同名组件
            If ComponentExists(vbProj, compName) Then
                ' 询问是否覆盖
                Dim ans As VbMsgBoxResult
                ans = MsgBox("组件 '" & compName & "' 已存在，是否覆盖？" & vbCrLf & _
                             "是=覆盖, 否=跳过, 取消=中止安装", vbYesNoCancel + vbQuestion)
                If ans = vbCancel Then
                    ImportAllComponents = "安装已取消。"
                    Exit Function
                ElseIf ans = vbNo Then
                    skipped = skipped + 1
                    GoTo NextFile
                Else
                    ' 删除旧组件
                    RemoveComponent vbProj, compName
                End If
            End If
            
            ' 导入新组件
            On Error Resume Next
            vbProj.VBComponents.Import file.Path
            If Err.Number = 0 Then
                imported = imported + 1
                Debug.Print "[导入] " & compName
            Else
                failed = failed + 1
                Debug.Print "[失败] " & compName & " - " & Err.Description
                Err.Clear
            End If
            On Error GoTo 0
        End If
NextFile:
    Next file
    
    ImportAllComponents = "导入完成: " & imported & " 成功, " & skipped & " 跳过, " & failed & " 失败"
End Function

' ============================================
' 检查/添加 VBIDE 引用
' ============================================

Private Function EnsureVBIDEReference() As Boolean
    On Error Resume Next
    
    Dim vbProj As Object
    Set vbProj = ActiveWorkbook.VBProject
    
    ' 检查是否已有 VBIDE 引用
    Dim ref As Object
    For Each ref In vbProj.References
        If ref.Name = VBIDE_NAME Then
            EnsureVBIDEReference = True
            Exit Function
        End If
    Next ref
    
    ' 尝试添加引用
    Err.Clear
    vbProj.References.AddFromGuid VBIDE_GUID, VBIDE_MAJOR, VBIDE_MINOR
    If Err.Number = 0 Then
        EnsureVBIDEReference = True
    Else
        ' 可能 GUID 不同，尝试从文件路径添加
        Err.Clear
        Dim vbidePath As String
        vbidePath = Environ("CommonProgramFiles") & "\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB"
        If Dir(vbidePath) = "" Then
            vbidePath = Environ("CommonProgramFiles(x86)") & "\Microsoft Shared\VBA\VBA6\VBE6EXT.OLB"
        End If
        If Dir(vbidePath) <> "" Then
            vbProj.References.AddFromFile vbidePath
            If Err.Number = 0 Then EnsureVBIDEReference = True
        End If
    End If
    
    On Error GoTo 0
End Function

' ============================================
' 组件管理辅助函数
' ============================================

Private Function ComponentExists(vbProj As Object, name As String) As Boolean
    On Error Resume Next
    Dim comp As Object
    Set comp = vbProj.VBComponents(name)
    ComponentExists = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Sub RemoveComponent(vbProj As Object, name As String)
    On Error Resume Next
    vbProj.VBComponents.Remove vbProj.VBComponents(name)
    On Error GoTo 0
End Sub

' ============================================
' 文件夹选择对话框
' ============================================

Private Function SelectFolder(prompt As String) As String
    Dim shellApp As Object
    Set shellApp = CreateObject("Shell.Application")
    
    Dim folder As Object
    Set folder = shellApp.BrowseForFolder(0, prompt, 0)
    
    If Not folder Is Nothing Then
        SelectFolder = folder.Self.Path
    Else
        SelectFolder = ""
    End If
End Function

' ============================================
' 下载文本（用于 GitHub 检测）
' ============================================

Private Function DownloadText(url As String) As String
    On Error Resume Next
    Dim xhr As Object
    Set xhr = CreateObject("MSXML2.XMLHTTP.6.0")
    xhr.Open "GET", url, False
    xhr.setTimeouts 5000, 5000, 10000, 10000
    xhr.Send
    If xhr.Status = 200 Then
        DownloadText = xhr.responseText
    End If
    On Error GoTo 0
End Function

' ============================================
' 快速安装（假设代码在固定路径）
' ============================================

Public Sub QuickInstall()
    ' 假设用户把代码解压到了桌面
    Dim desktop As String
    desktop = Environ("USERPROFILE") & "\Desktop\VBAPlaywright"
    
    If Dir(desktop, vbDirectory) = "" Then
        desktop = Environ("USERPROFILE") & "\Downloads\VBAPlaywright"
    End If
    
    If Dir(desktop, vbDirectory) = "" Then
        MsgBox "未找到默认代码文件夹。" & vbCrLf & _
               "请先将代码解压到桌面或下载文件夹，" & vbCrLf & _
               "或运行 InstallFromFolder 手动选择。", vbExclamation
        Exit Sub
    End If
    
    ' 检查/添加引用
    If Not EnsureVBIDEReference Then
        MsgBox "无法添加 VBIDE 引用。请手动添加后重试。", vbCritical
        Exit Sub
    End If
    
    Dim result As String
    result = ImportAllComponents(ActiveWorkbook, desktop)
    MsgBox result & vbCrLf & vbCrLf & "请另存为 .xlsm 格式！", vbInformation
End Sub
