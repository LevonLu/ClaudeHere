# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## 项目概述

Claude Here 是一个 Windows 右键菜单工具，通过 PowerShell 脚本在 Windows 注册表中创建右键菜单项，使用户可以在任意文件夹中快速启动 Claude Code。

## 常用命令

```powershell
# 直接运行脚本 - 进入交互式菜单（推荐）
.\Install-ClaudeHere.ps1

# 安装右键菜单（需要管理员权限）
.\Install-ClaudeHere.ps1 -Action install

# 指定语言安装
.\Install-ClaudeHere.ps1 -Action install -Language en-US  # 英文
.\Install-ClaudeHere.ps1 -Action install -Language zh-CN  # 中文
.\Install-ClaudeHere.ps1 -Action install -Language auto   # 自动检测

# 卸载右键菜单
.\Install-ClaudeHere.ps1 -Action uninstall

# 更新现有配置
.\Install-ClaudeHere.ps1 -Action update

# 显示帮助信息
.\Install-ClaudeHere.ps1 -Action help
```

## 代码架构

### 核心脚本结构

[Install-ClaudeHere.ps1](Install-ClaudeHere.ps1) 是唯一的 PowerShell 脚本文件，按功能区域组织：

1. **参数定义** (行 24-33): 定义 `Action`（默认为 `interactive`）和 `Language` 参数
2. **国际化支持** (行 44-313):
   - `Get-Messages`: 返回指定语言的消息哈希表
   - `Initialize-Localization`: 初始化本地化设置
   - `Get-LocalizedString`: 获取本地化字符串
3. **交互式菜单** (行 316-383):
   - `Show-InteractiveMenu`: 显示交互式菜单让用户选择操作
   - `Pause-AfterAction`: 在操作完成后暂停，等待用户按键返回菜单
4. **权限和依赖检查** (行 473-509):
   - `Test-AdminPrivileges`: 检查管理员权限
   - `Test-WindowsTerminal`: 检查 Windows Terminal 是否安装
   - `Test-ClaudeCommand`: 检查 claude 命令是否在 PATH 中
5. **Git Bash 路径检测** (行 514-530):
   - `Find-GitBashPath`: 自动检测 Git Bash 安装路径
6. **注册表操作** (行 535-634):
   - `Set-ClaudeHereRegistry`: 创建右键菜单注册表项
   - `Remove-ClaudeHereRegistry`: 删除右键菜单注册表项
7. **安装/卸载/更新功能** (行 639-883):
   - `Install-ClaudeHere`: 交互式安装向导
   - `Uninstall-ClaudeHere`: 卸载向导
   - `Update-ClaudeHere`: 更新配置向导
   - `Save-UserConfig`: 保存用户配置到注册表

### 注册表结构

脚本在以下位置创建注册表项：
- `HKEY_CLASSES_ROOT\Directory\Background\shell\ClaudeHere` - 文件夹背景右键
- `HKEY_CLASSES_ROOT\Directory\shell\ClaudeHere` - 文件夹本身右键
- `HKEY_CURRENT_USER\Software\ClaudeHere` - 用户配置存储

启动命令格式：
```
wt.exe -w 0 nt --title "Claude" "<bash.exe路径>" -c "cd '<路径>' && exec claude"
```

### 国际化 (i18n)

脚本支持多语言界面，通过以下方式实现：
- 在 `Get-Messages` 函数中定义每种语言的消息哈希表
- 使用 `$supportedCultures` 数组定义支持的语言（目前支持 `en-US` 和 `zh-CN`）
- 通过 `Get-LocalizedString` 函数获取本地化字符串，支持格式化参数

添加新语言的步骤：
1. 在 `Get-Messages` 函数中添加新的语言哈希表
2. 更新 `$supportedCultures` 数组

## 开发注意事项

- 脚本需要管理员权限才能修改注册表
- 所有用户界面文本必须使用 `Get-LocalizedString` 函数以支持多语言
- 修改注册表前应检查路径是否存在
- Git Bash 路径检测应覆盖常见的安装位置
- 错误处理使用 `try-catch` 块，并显示本地化的错误消息
