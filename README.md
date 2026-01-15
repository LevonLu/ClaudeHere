# Claude Here

在Windows右键菜单添加"Claude Here"选项，快速在当前目录启动Claude Code。

## 功能特性

- **双位置支持**: 支持文件夹背景和文件夹本身两处右键菜单
- **快速启动**: 右键即可在当前目录启动Claude Code
- **使用Windows Terminal**: 使用现代化的终端体验
- **Git Bash集成**: 在Git Bash环境中运行Claude
- **可自定义**: 支持自定义菜单名称和图标
- **易于管理**: 提供安装、卸载、更新功能
- **自动检测**: 自动检测Git Bash安装路径和默认图标
- **多语言支持**: 支持中文和英文界面，自动检测系统语言

## 国际化 (Internationalization)

### 支持的语言

- 🇨🇳 简体中文 (zh-CN)
- 🇺🇸 English (en-US)

### 使用方法

脚本会自动检测Windows系统语言并显示对应的界面。如果需要手动指定语言：

```powershell
# 强制使用英文界面
.\Install-ClaudeHere.ps1 -Action install -Language en-US

# 强制使用中文界面
.\Install-ClaudeHere.ps1 -Action install -Language zh-CN

# 自动检测（默认）
.\Install-ClaudeHere.ps1 -Action install -Language auto
```

### 贡献翻译

欢迎贡献新的语言翻译：

1. 在脚本目录创建语言文件夹（如 `fr-FR`、`ja-JP` 等）
2. 复制 `en-US/Install-ClaudeHere.psd1` 到新文件夹
3. 翻译所有字符串值，保持键名不变
4. 更新脚本中的 `$supportedCultures` 数组
5. 提交Pull Request

## 系统要求

- Windows 10/11
- Windows Terminal（从Microsoft Store安装）
- Git for Windows
- Claude Code CLI（已安装并添加到PATH）
- PowerShell 5.1 或更高版本

## 快速开始

### 安装

以管理员身份运行PowerShell，执行：

```powershell
.\Install-ClaudeHere.ps1 -Action install
```

安装程序会引导你完成配置：
1. 菜单显示名称（默认"Claude Here"）
2. Git Bash路径（自动检测）
3. 图标路径（可选）

### 使用

安装完成后，在任何文件夹中右键，选择"Claude Here"即可启动Claude Code。

## 命令

```powershell
# 安装右键菜单（自动检测语言）
.\Install-ClaudeHere.ps1 -Action install

# 指定英文界面安装
.\Install-ClaudeHere.ps1 -Action install -Language en-US

# 指定中文界面安装
.\Install-ClaudeHere.ps1 -Action install -Language zh-CN

# 卸载右键菜单
.\Install-ClaudeHere.ps1 -Action uninstall

# 更新现有配置
.\Install-ClaudeHere.ps1 -Action update

# 显示帮助信息
.\Install-ClaudeHere.ps1 -Action help
```

### 参数说明

| 参数 | 可选值 | 说明 |
|------|--------|------|
| `-Action` | `install`, `uninstall`, `update`, `help` | 要执行的操作 |
| `-Language` | `auto`, `en-US`, `zh-CN` | 界面语言，默认为 `auto`（自动检测系统语言） |

## 配置说明

### 菜单名称
自定义在右键菜单中显示的文字。默认为"Claude Here"。

### Git Bash路径
脚本会自动检测Git Bash的安装位置。常见路径：
- `C:\Program Files\Git\bin\bash.exe`
- `C:\Program Files (x86)\Git\bin\bash.exe`
- `%LOCALAPPDATA%\Programs\Git\bin\bash.exe`

### 图标
可指定.ico格式的图标文件，让右键菜单更美观。

## 工作原理

脚本通过在Windows注册表中创建以下键值实现右键菜单：

```
HKEY_CLASSES_ROOT\Directory\Background\shell\ClaudeHere    # 文件夹背景右键
HKEY_CLASSES_ROOT\Directory\shell\ClaudeHere              # 文件夹本身右键
```

启动命令使用Windows Terminal的命令行参数：

```cmd
wt.exe -w 0 nt --title "Claude" "C:\Program Files\Git\bin\bash.exe" -c "cd '%V' && exec claude"
```

参数说明：
- `-w 0`: 在新窗口打开
- `nt`: 使用默认配置文件
- `--title "Claude"`: 设置终端标题
- `bash.exe`: Git Bash可执行文件的完整路径
- `%V` / `%1`: 当前目录路径（背景用%V，文件夹用%1）
- `exec`: 用claude进程替换bash进程

## 故障排除

### 权限错误
**问题**: 提示需要管理员权限
**解决**: 右键脚本，选择"以管理员身份运行"

### 未找到Windows Terminal
**问题**: 提示未找到Windows Terminal
**解决**: 从Microsoft Store安装[Windows Terminal](https://aka.ms/terminal)

### 未找到Claude命令
**问题**: 提示未找到claude命令
**解决**: 确保Claude Code CLI已安装并添加到系统PATH

### 右键菜单未出现
**问题**: 安装后右键菜单没有出现
**解决**:
1. 检查是否以管理员权限运行
2. 重启Windows资源管理器（任务管理器 > 重启）
3. 检查注册表编辑器中是否存在以下路径：
   - `HKEY_CLASSES_ROOT\Directory\Background\shell\ClaudeHere`
   - `HKEY_CLASSES_ROOT\Directory\shell\ClaudeHere`

## 安全说明

此脚本仅修改用户自己的注册表设置，不会修改系统文件或安装任何软件。

## AI

本仓库代码和README均由AI生成，使用Claude Code + GLM完成。

## 许可证

MIT License
