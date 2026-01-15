<#
.SYNOPSIS
    Claude Here - Windows右键菜单工具

.DESCRIPTION
    在Windows右键菜单添加"Claude Here"选项，快速启动Claude Code

.PARAMETER Action
    要执行的操作：install（安装）、uninstall（卸载）、update（更新）、help（帮助）

.EXAMPLE
    .\Install-ClaudeHere.ps1 -Action install
    安装右键菜单

.EXAMPLE
    .\Install-ClaudeHere.ps1 -Action uninstall
    卸载右键菜单

.EXAMPLE
    .\Install-ClaudeHere.ps1 -Action update
    更新配置
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [ValidateSet('install', 'uninstall', 'update', 'help')]
    [string]$Action = 'help',

    [Parameter(Mandatory=$false)]
    [ValidateSet('en-US', 'zh-CN', 'auto')]
    [string]$Language = 'auto'
)

# 脚本配置
$ErrorActionPreference = 'Stop'
# 文件夹背景右键菜单
$RegistryKeyPath_Background = 'Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\ClaudeHere'
# 文件夹本身右键菜单
$RegistryKeyPath_Directory = 'Registry::HKEY_CLASSES_ROOT\Directory\shell\ClaudeHere'
$ConfigRegistryPath = 'Registry::HKEY_CURRENT_USER\Software\ClaudeHere'


#region 国际化支持

function Get-Messages {
    <#
    .SYNOPSIS
        返回指定语言的消息哈希表
    .PARAMETER Culture
        语言代码（如 'en-US', 'zh-CN'）
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$Culture
    )

    # 英文消息
    $enMessages = @{
        ScriptTitle = "Claude Here - Windows Right-Click Menu Tool"
        Error = "Error"
        ErrorAdminRequired = "This script requires administrator privileges."
        ErrorRunAsAdmin = "Please right-click the script and select 'Run as Administrator'."
        ErrorTerminalNotFound = "Windows Terminal not found."
        ErrorInstallTerminal = "Please install Windows Terminal from Microsoft Store."
        ErrorPathNotFound = "The specified path does not exist."
        ErrorInvalidPath = "Must provide a valid Git Bash path."
        ErrorIconNotFound = "Icon file does not exist."
        ErrorIconWrongFormat = "Icon file should be in .ico format."
        ErrorNotInstalled = "No installed right-click menu found, please run install first."
        ErrorRegistryCreateFailed = "Failed to create registry entries"
        ErrorRegistryRemoveFailed = "Failed to remove registry entries"
        Warning = "Warning"
        WarningClaudeNotFound = "claude command not found in system PATH."
        WarningClaudeInstallNote = "Make sure Claude Code is installed and added to PATH after installation."
        PromptContinueInstall = "Continue with installation? (y/N)"
        WarningGitBashNotFound = "Git Bash not automatically detected."
        WarningConfigSaveFailed = "Failed to save configuration"
        WarningIconNotSet = "Icon file does not exist, no icon will be set"
        WarningIconWrongFormatNotSet = "Icon file should be in .ico format, no icon will be set"
        WarningOldMenuCleaned = "Cleaned up old drive menu item"
        WarningNoMenuFound = "No installed right-click menu found"
        Info = "Info"
        InfoCreatingRegistry = "Creating registry entries..."
        InfoDeletingRegistry = "Deleting registry entries..."
        InfoDetectingGitBash = "Detecting Git Bash installation path..."
        InfoIconSetup = "Icon setup (optional)"
        InfoCurrentConfig = "Current Configuration"
        InfoKeepCurrentValue = "Press Enter to keep current value"
        InfoFolderBackground = "folder background"
        InfoFolder = "folder"
        Success = "Success"
        SuccessMenuCreated = "Successfully created right-click menu for"
        SuccessAllMenusCreated = "All right-click menus created successfully!"
        SuccessMenuName = "Menu Name"
        SuccessBashPath = "Git Bash"
        SuccessIconPath = "Icon Path"
        SuccessMenusDeleted = "Successfully deleted right-click menu items!"
        SuccessFoundGitBash = "Found Git Bash"
        PromptMenuName = "Enter menu display name (press Enter for default 'Claude Here')"
        PromptUseDetectedPath = "Use this path? (Y/n)"
        PromptEnterBashPath = "Enter full path to bash.exe"
        PromptEnterIconPath = "Enter icon file path (.ico file"
        PromptEnterIconPathDefault = "use default icon: claude.ico"
        PromptEnterIconPathSkip = "or press Enter to skip)"
        PromptUninstallConfirm = "Are you sure you want to uninstall Claude Here right-click menu? (y/N)"
        PromptNewMenuName = "Enter new menu name"
        PromptNewBashPath = "Enter new Git Bash path"
        PromptNewIconPath = "Enter new icon path (press Enter to keep current, type 'none' to remove icon)"
        WizardInstallTitle = "Claude Here - Installation Wizard"
        WizardUninstallTitle = "Claude Here - Uninstallation Wizard"
        WizardUpdateTitle = "Claude Here - Update Configuration Wizard"
        WizardInstallationComplete = "Installation Complete!"
        WizardUninstallationComplete = "Uninstallation Complete!"
        WizardUpdateComplete = "Configuration Updated!"
        WizardUninstallCancelled = "Uninstallation cancelled."
        WizardUsageInfo = "Now you can right-click in any folder and select '{0}' to launch Claude Code."
        LabelMenuName = "Menu Name"
        LabelGitBash = "Git Bash"
        LabelIconPath = "Icon Path"
        LabelIcon = "Icon"
        LabelNotSet = "(not set)"
        Tip = "Tip"
        TipDefaultIconFound = "Found default icon in script directory: claude.ico"
        HelpUsage = "Usage:"
        HelpCommand = ".\Install-ClaudeHere.ps1 -Action <Action> [-Language <Language>]"
        HelpAvailableActions = "Available Actions:"
        HelpActionInstall = "install    Install right-click menu (interactive configuration)"
        HelpActionUninstall = "uninstall  Uninstall right-click menu"
        HelpActionUpdate = "update     Update existing configuration"
        HelpActionHelp = "help       Display this help information"
        HelpAvailableLanguages = "Available Languages:"
        HelpLanguageEnUS = "en-US      English (United States)"
        HelpLanguageZhCN = "zh-CN      Chinese (Simplified)"
        HelpLanguageAuto = "auto       Auto-detect based on system locale (default)"
        HelpExamples = "Examples:"
        HelpNotes = "Notes:"
        HelpNoteAdmin = "- Requires administrator privileges to run"
        HelpNoteTerminal = "- Windows Terminal must be installed"
        HelpNoteGit = "- Git for Windows must be installed"
        HelpNoteClaude = "- Claude command must be available in system PATH"
        HelpSeparator1 = "==============================================================================="
        HelpSeparator2 = ""
    }

    # 中文消息
    $zhMessages = @{
        ScriptTitle = "Claude Here - Windows 右键菜单工具"
        Error = "错误"
        ErrorAdminRequired = "此脚本需要管理员权限。"
        ErrorRunAsAdmin = "请右键点击脚本并选择'以管理员身份运行'。"
        ErrorTerminalNotFound = "未找到 Windows Terminal。"
        ErrorInstallTerminal = "请从 Microsoft Store 安装 Windows Terminal。"
        ErrorPathNotFound = "指定的路径不存在。"
        ErrorInvalidPath = "必须提供有效的 Git Bash 路径。"
        ErrorIconNotFound = "图标文件不存在。"
        ErrorIconWrongFormat = "图标文件应为 .ico 格式。"
        ErrorNotInstalled = "未找到已安装的右键菜单，请先运行安装。"
        ErrorRegistryCreateFailed = "创建注册表项失败"
        ErrorRegistryRemoveFailed = "删除注册表项失败"
        Warning = "警告"
        WarningClaudeNotFound = "在系统 PATH 中未找到 claude 命令。"
        WarningClaudeInstallNote = "请确保已安装 Claude Code 并在安装后添加到 PATH。"
        PromptContinueInstall = "是否继续安装？(y/N)"
        WarningGitBashNotFound = "未自动检测到 Git Bash。"
        WarningConfigSaveFailed = "保存配置失败"
        WarningIconNotSet = "图标文件不存在，将不设置图标"
        WarningIconWrongFormatNotSet = "图标文件应为 .ico 格式，将不设置图标"
        WarningOldMenuCleaned = "已清理旧的驱动器菜单项"
        WarningNoMenuFound = "未找到已安装的右键菜单"
        Info = "信息"
        InfoCreatingRegistry = "正在创建注册表项..."
        InfoDeletingRegistry = "正在删除注册表项..."
        InfoDetectingGitBash = "正在检测 Git Bash 安装路径..."
        InfoIconSetup = "图标设置（可选）"
        InfoCurrentConfig = "当前配置"
        InfoKeepCurrentValue = "按 Enter 保持当前值"
        InfoFolderBackground = "文件夹背景"
        InfoFolder = "文件夹"
        Success = "成功"
        SuccessMenuCreated = "成功为以下位置创建右键菜单："
        SuccessAllMenusCreated = "所有右键菜单创建成功！"
        SuccessMenuName = "菜单名称"
        SuccessBashPath = "Git Bash"
        SuccessIconPath = "图标路径"
        SuccessMenusDeleted = "成功删除右键菜单项！"
        SuccessFoundGitBash = "找到 Git Bash"
        PromptMenuName = "输入菜单显示名称（按 Enter 使用默认名称 'Claude Here'）"
        PromptUseDetectedPath = "使用此路径？(Y/n)"
        PromptEnterBashPath = "输入 bash.exe 的完整路径"
        PromptEnterIconPath = "输入图标文件路径（.ico 文件"
        PromptEnterIconPathDefault = "使用默认图标：claude.ico"
        PromptEnterIconPathSkip = "或按 Enter 跳过）"
        PromptUninstallConfirm = "确定要卸载 Claude Here 右键菜单吗？(y/N)"
        PromptNewMenuName = "输入新菜单名称"
        PromptNewBashPath = "输入新的 Git Bash 路径"
        PromptNewIconPath = "输入新图标路径（按 Enter 保持当前值，输入 'none' 移除图标）"
        WizardInstallTitle = "Claude Here - 安装向导"
        WizardUninstallTitle = "Claude Here - 卸载向导"
        WizardUpdateTitle = "Claude Here - 更新配置向导"
        WizardInstallationComplete = "安装完成！"
        WizardUninstallationComplete = "卸载完成！"
        WizardUpdateComplete = "配置已更新！"
        WizardUninstallCancelled = "已取消卸载。"
        WizardUsageInfo = "现在您可以在任何文件夹中右键点击并选择 '{0}' 来启动 Claude Code。"
        LabelMenuName = "菜单名称"
        LabelGitBash = "Git Bash"
        LabelIconPath = "图标路径"
        LabelIcon = "图标"
        LabelNotSet = "（未设置）"
        Tip = "提示"
        TipDefaultIconFound = "在脚本目录中找到默认图标：claude.ico"
        HelpUsage = "用法："
        HelpCommand = ".\Install-ClaudeHere.ps1 -Action <操作> [-Language <语言>]"
        HelpAvailableActions = "可用操作："
        HelpActionInstall = "install    安装右键菜单（交互式配置）"
        HelpActionUninstall = "uninstall  卸载右键菜单"
        HelpActionUpdate = "update     更新现有配置"
        HelpActionHelp = "help       显示此帮助信息"
        HelpAvailableLanguages = "可用语言："
        HelpLanguageEnUS = "en-US      英语（美国）"
        HelpLanguageZhCN = "zh-CN      中文（简体）"
        HelpLanguageAuto = "auto       自动检测系统语言（默认）"
        HelpExamples = "示例："
        HelpNotes = "说明："
        HelpNoteAdmin = "- 需要管理员权限运行"
        HelpNoteTerminal = "- 必须安装 Windows Terminal"
        HelpNoteGit = "- 必须安装 Git for Windows"
        HelpNoteClaude = "- Claude 命令必须在系统 PATH 中可用"
        HelpSeparator1 = "==============================================================================="
        HelpSeparator2 = ""
    }

    switch ($Culture) {
        'zh-CN' { return $zhMessages }
        default { return $enMessages }
    }
}

function Initialize-Localization {
    <#
    .SYNOPSIS
        初始化本地化设置，加载对应语言的消息
    #>

    # 确定最终使用的语言
    $script:UICulture = if ($Language -eq 'auto') {
        # 自动检测：优先使用 $PSUICulture，如果不可用则回退到 en-US
        $detectedCulture = $PSUICulture

        # 检查是否支持该语言
        $supportedCultures = @('en-US', 'zh-CN')
        if ($detectedCulture -in $supportedCultures) {
            $detectedCulture
        } else {
            # 如果检测到的语言不支持（如 fr-FR），检查是否为中文变体
            if ($detectedCulture -like 'zh-*') {
                'zh-CN'
            } else {
                # 其他情况回退到英文
                'en-US'
            }
        }
    } else {
        # 用户指定了语言
        $Language
    }

    # 加载对应语言的消息
    $script:Messages = Get-Messages -Culture $script:UICulture

    Write-Debug "Loaded localization for culture: $script:UICulture"

    return $script:Messages
}

function Get-LocalizedString {
    <#
    .SYNOPSIS
        获取本地化字符串，支持格式化参数
    .PARAMETER Key
        消息键名
    .PARAMETER Arguments
        格式化参数（可选）
    .EXAMPLE
        Get-LocalizedString -Key 'WizardUsageInfo' -Arguments $menuName
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$Key,

        [Parameter(Mandatory=$false)]
        [array]$Arguments = $null
    )

    if ($null -eq $script:Messages -or -not $script:Messages.ContainsKey($Key)) {
        # 回退到键名本身
        $message = $Key
    } else {
        $message = $script:Messages[$Key]
    }

    if ($Arguments) {
        return $message -f $Arguments
    }

    return $message
}

# 初始化本地化
Initialize-Localization | Out-Null

#endregion


#region 主函数

function Main {
    switch ($Action) {
        'install'  {
            Test-AdminPrivileges
            Install-ClaudeHere
        }
        'uninstall' {
            Test-AdminPrivileges
            Uninstall-ClaudeHere
        }
        'update'   {
            Test-AdminPrivileges
            Update-ClaudeHere
        }
        'help'     {
            Show-Help
        }
    }
}

#endregion


#region 帮助函数

function Show-Help {
    $separator = Get-LocalizedString -Key 'HelpSeparator1'
    $separator2 = Get-LocalizedString -Key 'HelpSeparator2'

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine()
    [void]$sb.AppendLine($separator)
    [void]$sb.AppendLine("                    $(Get-LocalizedString -Key 'ScriptTitle')")
    [void]$sb.AppendLine($separator)
    [void]$sb.AppendLine()
    [void]$sb.AppendLine("$(Get-LocalizedString -Key 'HelpUsage')")
    [void]$sb.AppendLine("    $(Get-LocalizedString -Key 'HelpCommand')")
    [void]$sb.AppendLine()
    [void]$sb.AppendLine("$(Get-LocalizedString -Key 'HelpAvailableActions')")
    [void]$sb.AppendLine("    $(Get-LocalizedString -Key 'HelpActionInstall')")
    [void]$sb.AppendLine("    $(Get-LocalizedString -Key 'HelpActionUninstall')")
    [void]$sb.AppendLine("    $(Get-LocalizedString -Key 'HelpActionUpdate')")
    [void]$sb.AppendLine("    $(Get-LocalizedString -Key 'HelpActionHelp')")
    [void]$sb.AppendLine()
    [void]$sb.AppendLine("$(Get-LocalizedString -Key 'HelpAvailableLanguages')")
    [void]$sb.AppendLine("    $(Get-LocalizedString -Key 'HelpLanguageEnUS')")
    [void]$sb.AppendLine("    $(Get-LocalizedString -Key 'HelpLanguageZhCN')")
    [void]$sb.AppendLine("    $(Get-LocalizedString -Key 'HelpLanguageAuto')")
    [void]$sb.AppendLine()
    [void]$sb.AppendLine("$(Get-LocalizedString -Key 'HelpExamples')")
    [void]$sb.AppendLine("    .\Install-ClaudeHere.ps1 -Action install")
    [void]$sb.AppendLine("    .\Install-ClaudeHere.ps1 -Action install -Language en-US")
    [void]$sb.AppendLine("    .\Install-ClaudeHere.ps1 -Action uninstall")
    [void]$sb.AppendLine("    .\Install-ClaudeHere.ps1 -Action update")
    [void]$sb.AppendLine()
    [void]$sb.AppendLine("$(Get-LocalizedString -Key 'HelpNotes')")
    [void]$sb.AppendLine("    $(Get-LocalizedString -Key 'HelpNoteAdmin')")
    [void]$sb.AppendLine("    $(Get-LocalizedString -Key 'HelpNoteTerminal')")
    [void]$sb.AppendLine("    $(Get-LocalizedString -Key 'HelpNoteGit')")
    [void]$sb.AppendLine("    $(Get-LocalizedString -Key 'HelpNoteClaude')")
    [void]$sb.AppendLine()
    [void]$sb.AppendLine($separator)
    [void]$sb.AppendLine($separator2)
    [void]$sb.AppendLine()

    # Execute the string to expand variables
    $helpText = $ExecutionContext.InvokeCommand.ExpandString($sb.ToString())

    Write-Host $helpText -ForegroundColor Cyan
}

#endregion


#region 权限和依赖检查

function Test-AdminPrivileges {
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if (-not $isAdmin) {
        Write-Host "[$(Get-LocalizedString -Key 'Error')] $(Get-LocalizedString -Key 'ErrorAdminRequired')" -ForegroundColor Red
        Write-Host (Get-LocalizedString -Key 'ErrorRunAsAdmin') -ForegroundColor Yellow
        exit 1
    }
}

function Test-WindowsTerminal {
    $wtPath = Join-Path $env:LOCALAPPDATA "Microsoft\WindowsApps\wt.exe"

    if (-not (Test-Path $wtPath)) {
        Write-Host "[$(Get-LocalizedString -Key 'Error')] $(Get-LocalizedString -Key 'ErrorTerminalNotFound')" -ForegroundColor Red
        Write-Host (Get-LocalizedString -Key 'ErrorInstallTerminal') -ForegroundColor Yellow
        return $false
    }

    return $true
}

function Test-ClaudeCommand {
    try {
        $null = Get-Command claude -ErrorAction Stop
        return $true
    }
    catch {
        Write-Host "[$(Get-LocalizedString -Key 'Warning')] $(Get-LocalizedString -Key 'WarningClaudeNotFound')" -ForegroundColor Yellow
        Write-Host (Get-LocalizedString -Key 'WarningClaudeInstallNote') -ForegroundColor Yellow
        $response = Read-Host "$(Get-LocalizedString -Key 'PromptContinueInstall')"

        return $response -eq 'y' -or $response -eq 'Y'
    }
}

#endregion


#region Git Bash 路径检测

function Find-GitBashPath {
    $possiblePaths = @(
        "$env:ProgramFiles\Git\bin\bash.exe",
        "$env:ProgramFiles(x86)\Git\bin\bash.exe",
        "${env:LOCALAPPDATA}\Programs\Git\bin\bash.exe"
    )

    foreach ($path in $possiblePaths) {
        if (Test-Path $path) {
            return $path
        }
    }

    return $null
}

#endregion


#region 注册表操作

function Set-ClaudeHereRegistry {
    param(
        [Parameter(Mandatory=$true)]
        [string]$MenuName,

        [Parameter(Mandatory=$true)]
        [string]$BashPath,

        [Parameter(Mandatory=$false)]
        [string]$IconPath = ''
    )

    # 创建注册表项
    Write-Host "`n[$(Get-LocalizedString -Key 'Info')] $(Get-LocalizedString -Key 'InfoCreatingRegistry')" -ForegroundColor Cyan

    try {
        # 定义所有需要创建的注册表位置
        $registryPaths = @(
            @{ Path = $RegistryKeyPath_Background; Arg = '%V'; Desc = Get-LocalizedString -Key 'InfoFolderBackground' },
            @{ Path = $RegistryKeyPath_Directory;   Arg = '%1'; Desc = Get-LocalizedString -Key 'InfoFolder' }
        )

        foreach ($entry in $registryPaths) {
            $regPath = $entry.Path
            $arg = $entry.Arg
            $desc = $entry.Desc

            # 生成启动命令
            $startupCommand = "wt.exe -w 0 nt --title `"Claude`" `"$BashPath`" -c `"cd '$arg' && exec claude`""

            # 删除旧的注册表项（如果存在）
            if (Test-Path $regPath) {
                Remove-Item -Path $regPath -Recurse -Force
            }

            # 创建主键
            $null = New-Item -Path $regPath -Force
            Set-ItemProperty -Path $regPath -Name "(default)" -Value $MenuName

            # 设置图标
            if ($IconPath -and (Test-Path $IconPath)) {
                Set-ItemProperty -Path $regPath -Name "Icon" -Value $IconPath
            }

            # 创建command子键
            $commandPath = "$regPath\command"
            $null = New-Item -Path $commandPath -Force
            Set-ItemProperty -Path $commandPath -Name "(default)" -Value $startupCommand

            Write-Host "[$(Get-LocalizedString -Key 'Success')] $(Get-LocalizedString -Key 'SuccessMenuCreated') $desc" -ForegroundColor Green
        }

        Write-Host "`n[$(Get-LocalizedString -Key 'Success')] $(Get-LocalizedString -Key 'SuccessAllMenusCreated')" -ForegroundColor Green
        Write-Host "  $(Get-LocalizedString -Key 'SuccessMenuName'): $MenuName" -ForegroundColor Gray
        Write-Host "  $(Get-LocalizedString -Key 'SuccessBashPath'): $BashPath" -ForegroundColor Gray
        if ($IconPath) {
            Write-Host "  $(Get-LocalizedString -Key 'SuccessIconPath'): $IconPath" -ForegroundColor Gray
        }
    }
    catch {
        Write-Host "[$(Get-LocalizedString -Key 'Error')] $(Get-LocalizedString -Key 'ErrorRegistryCreateFailed'): $_" -ForegroundColor Red
        exit 1
    }
}

function Remove-ClaudeHereRegistry {
    Write-Host "`n[$(Get-LocalizedString -Key 'Info')] $(Get-LocalizedString -Key 'InfoDeletingRegistry')" -ForegroundColor Cyan

    try {
        $deletedCount = 0
        $allPaths = @($RegistryKeyPath_Background, $RegistryKeyPath_Directory)

        foreach ($path in $allPaths) {
            if (Test-Path $path) {
                Remove-Item -Path $path -Recurse -Force
                $deletedCount++
            }
        }

        # 同时清理旧的驱动器注册表项（如果存在）
        $oldDrivePath = 'Registry::HKEY_CLASSES_ROOT\Drive\shell\ClaudeHere'
        if (Test-Path $oldDrivePath) {
            Remove-Item -Path $oldDrivePath -Recurse -Force
            Write-Host "[$(Get-LocalizedString -Key 'Tip')] $(Get-LocalizedString -Key 'WarningOldMenuCleaned')" -ForegroundColor Yellow
        }

        if ($deletedCount -gt 0) {
            Write-Host "[$(Get-LocalizedString -Key 'Success')] $(Get-LocalizedString -Key 'SuccessMenusDeleted') $deletedCount" -ForegroundColor Green
        }
        else {
            Write-Host "[$(Get-LocalizedString -Key 'Tip')] $(Get-LocalizedString -Key 'WarningNoMenuFound')" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "[$(Get-LocalizedString -Key 'Error')] $(Get-LocalizedString -Key 'ErrorRegistryRemoveFailed'): $_" -ForegroundColor Red
        exit 1
    }
}

#endregion


#region 安装功能

function Install-ClaudeHere {
    $separator = Get-LocalizedString -Key 'HelpSeparator1'

    Write-Host "`n$separator" -ForegroundColor Cyan
    Write-Host "                    $(Get-LocalizedString -Key 'WizardInstallTitle')" -ForegroundColor Cyan
    Write-Host "$separator`n" -ForegroundColor Cyan

    # 检查 Windows Terminal
    if (-not (Test-WindowsTerminal)) {
        exit 1
    }

    # 检查 Claude 命令
    if (-not (Test-ClaudeCommand)) {
        exit 1
    }

    # 菜单名称
    $menuName = Read-Host "`n$(Get-LocalizedString -Key 'PromptMenuName')"
    if ([string]::IsNullOrWhiteSpace($menuName)) {
        $menuName = "Claude Here"
    }

    # Git Bash 路径
    Write-Host "`n[$(Get-LocalizedString -Key 'Info')] $(Get-LocalizedString -Key 'InfoDetectingGitBash')" -ForegroundColor Cyan
    $bashPath = Find-GitBashPath

    if ($bashPath) {
        Write-Host "[$(Get-LocalizedString -Key 'Success')] $(Get-LocalizedString -Key 'SuccessFoundGitBash'): $bashPath" -ForegroundColor Green
        $useDefault = Read-Host "$(Get-LocalizedString -Key 'PromptUseDetectedPath')"

        if ($useDefault -ne 'n' -and $useDefault -ne 'N') {
            # 使用检测到的路径
        }
        else {
            $bashPath = Read-Host "$(Get-LocalizedString -Key 'PromptEnterBashPath')"
            if (-not (Test-Path $bashPath)) {
                Write-Host "[$(Get-LocalizedString -Key 'Error')] $(Get-LocalizedString -Key 'ErrorPathNotFound')" -ForegroundColor Red
                exit 1
            }
        }
    }
    else {
        Write-Host "[$(Get-LocalizedString -Key 'Warning')] $(Get-LocalizedString -Key 'WarningGitBashNotFound')" -ForegroundColor Yellow
        $bashPath = Read-Host "$(Get-LocalizedString -Key 'PromptEnterBashPath')"

        if ([string]::IsNullOrWhiteSpace($bashPath) -or -not (Test-Path $bashPath)) {
            Write-Host "[$(Get-LocalizedString -Key 'Error')] $(Get-LocalizedString -Key 'ErrorInvalidPath')" -ForegroundColor Red
            exit 1
        }
    }

    # 图标路径
    # 自动检测脚本所在目录的 claude.ico
    $defaultIconPath = Join-Path $PSScriptRoot "claude.ico"

    Write-Host "`n[$(Get-LocalizedString -Key 'Info')] $(Get-LocalizedString -Key 'InfoIconSetup')" -ForegroundColor Cyan
    $iconPrompt = "$(Get-LocalizedString -Key 'PromptEnterIconPath')"

    if (Test-Path $defaultIconPath) {
        $iconPrompt += "$(Get-LocalizedString -Key 'PromptEnterIconPathDefault')"
        Write-Host "[$(Get-LocalizedString -Key 'Tip')] $(Get-LocalizedString -Key 'TipDefaultIconFound')" -ForegroundColor Green
    }
    else {
        $iconPrompt += "$(Get-LocalizedString -Key 'PromptEnterIconPathSkip')"
    }
    $iconPrompt += ")"

    $iconPath = Read-Host $iconPrompt

    if ([string]::IsNullOrWhiteSpace($iconPath)) {
        # 用户直接回车，如果有默认图标就使用
        if (Test-Path $defaultIconPath) {
            $iconPath = $defaultIconPath
        }
        else {
            $iconPath = ''
        }
    }
    else {
        if (-not (Test-Path $iconPath)) {
            Write-Host "[$(Get-LocalizedString -Key 'Warning')] $(Get-LocalizedString -Key 'WarningIconNotSet')" -ForegroundColor Yellow
            $iconPath = ''
        }
        elseif (-not $iconPath.EndsWith('.ico')) {
            Write-Host "[$(Get-LocalizedString -Key 'Warning')] $(Get-LocalizedString -Key 'WarningIconWrongFormatNotSet')" -ForegroundColor Yellow
            $iconPath = ''
        }
    }

    # 创建注册表项
    Set-ClaudeHereRegistry -MenuName $menuName -BashPath $bashPath -IconPath $iconPath

    # 保存配置到用户注册表
    Save-UserConfig -MenuName $menuName -BashPath $bashPath -IconPath $iconPath

    Write-Host "`n$separator" -ForegroundColor Green
    Write-Host "                          $(Get-LocalizedString -Key 'WizardInstallationComplete')" -ForegroundColor Green
    Write-Host "$separator" -ForegroundColor Green
    Write-Host "`n$(Get-LocalizedString -Key 'WizardUsageInfo' -Arguments $menuName)`n" -ForegroundColor White
}

function Save-UserConfig {
    param(
        [string]$MenuName,
        [string]$BashPath,
        [string]$IconPath
    )

    try {
        $null = New-Item -Path $ConfigRegistryPath -Force
        Set-ItemProperty -Path $ConfigRegistryPath -Name "MenuName" -Value $MenuName
        Set-ItemProperty -Path $ConfigRegistryPath -Name "BashPath" -Value $BashPath
        Set-ItemProperty -Path $ConfigRegistryPath -Name "IconPath" -Value $IconPath
    }
    catch {
        Write-Host "[$(Get-LocalizedString -Key 'Warning')] $(Get-LocalizedString -Key 'WarningConfigSaveFailed'): $_" -ForegroundColor Yellow
    }
}

#endregion


#region 卸载功能

function Uninstall-ClaudeHere {
    $separator = Get-LocalizedString -Key 'HelpSeparator1'

    Write-Host "`n$separator" -ForegroundColor Cyan
    Write-Host "                    $(Get-LocalizedString -Key 'WizardUninstallTitle')" -ForegroundColor Cyan
    Write-Host "$separator`n" -ForegroundColor Cyan

    $response = Read-Host "$(Get-LocalizedString -Key 'PromptUninstallConfirm')"

    if ($response -ne 'y' -and $response -ne 'Y') {
        Write-Host "`n$(Get-LocalizedString -Key 'WizardUninstallCancelled')" -ForegroundColor Yellow
        return
    }

    Remove-ClaudeHereRegistry

    # 清理配置
    try {
        if (Test-Path $ConfigRegistryPath) {
            Remove-Item -Path $ConfigRegistryPath -Recurse -Force
        }
    }
    catch {
        # 忽略配置清理错误
    }

    Write-Host "`n$separator" -ForegroundColor Green
    Write-Host "                          $(Get-LocalizedString -Key 'WizardUninstallationComplete')" -ForegroundColor Green
    Write-Host "$separator`n" -ForegroundColor Green
}

#endregion


#region 更新功能

function Update-ClaudeHere {
    $separator = Get-LocalizedString -Key 'HelpSeparator1'

    Write-Host "`n$separator" -ForegroundColor Cyan
    Write-Host "                    $(Get-LocalizedString -Key 'WizardUpdateTitle')" -ForegroundColor Cyan
    Write-Host "$separator`n" -ForegroundColor Cyan

    # 检查是否已安装（检查任意一个路径）
    $allPaths = @($RegistryKeyPath_Background, $RegistryKeyPath_Directory)
    $installedPath = $null

    foreach ($path in $allPaths) {
        if (Test-Path $path) {
            $installedPath = $path
            break
        }
    }

    if (-not $installedPath) {
        Write-Host "[$(Get-LocalizedString -Key 'Error')] $(Get-LocalizedString -Key 'ErrorNotInstalled')" -ForegroundColor Red
        exit 1
    }

    # 读取当前配置
    $currentMenuName = (Get-ItemProperty -Path $installedPath -ErrorAction SilentlyContinue)."(default)"
    $currentBashPath = $null
    $currentIconPath = (Get-ItemProperty -Path $installedPath -ErrorAction SilentlyContinue)."Icon"

    # 从command中提取Bash路径
    $commandValue = (Get-ItemProperty -Path "$installedPath\command" -ErrorAction SilentlyContinue)."(default)"
    if ($commandValue -match '"([^"]+\\bash\.exe)"') {
        $currentBashPath = $matches[1]
    }

    Write-Host "[$(Get-LocalizedString -Key 'InfoCurrentConfig')]" -ForegroundColor Yellow
    Write-Host "  $(Get-LocalizedString -Key 'LabelMenuName'): $currentMenuName" -ForegroundColor Gray
    Write-Host "  $(Get-LocalizedString -Key 'LabelGitBash'): $currentBashPath" -ForegroundColor Gray
    Write-Host "  $(Get-LocalizedString -Key 'LabelIconPath'): $(if ($currentIconPath) { $currentIconPath } else { (Get-LocalizedString -Key 'LabelNotSet') })" -ForegroundColor Gray

    # 新的菜单名称
    Write-Host "`n[$(Get-LocalizedString -Key 'Tip')] $(Get-LocalizedString -Key 'InfoKeepCurrentValue')" -ForegroundColor Cyan
    $newMenuName = Read-Host "`n$(Get-LocalizedString -Key 'PromptNewMenuName')"

    if ([string]::IsNullOrWhiteSpace($newMenuName)) {
        $newMenuName = $currentMenuName
    }

    # 新的 Bash 路径
    $newBashPath = Read-Host "`n$(Get-LocalizedString -Key 'PromptNewBashPath')"

    if ([string]::IsNullOrWhiteSpace($newBashPath)) {
        $newBashPath = $currentBashPath
    }
    elseif (-not (Test-Path $newBashPath)) {
        Write-Host "[$(Get-LocalizedString -Key 'Error')] $(Get-LocalizedString -Key 'ErrorPathNotFound')" -ForegroundColor Red
        exit 1
    }

    # 新的图标路径
    $newIconPath = Read-Host "`n$(Get-LocalizedString -Key 'PromptNewIconPath')"

    if ([string]::IsNullOrWhiteSpace($newIconPath)) {
        $newIconPath = $currentIconPath
    }
    elseif ($newIconPath -eq 'none') {
        $newIconPath = ''
    }
    elseif (-not (Test-Path $newIconPath)) {
        Write-Host "[$(Get-LocalizedString -Key 'Error')] $(Get-LocalizedString -Key 'ErrorIconNotFound')" -ForegroundColor Red
        exit 1
    }

    # 更新注册表
    Set-ClaudeHereRegistry -MenuName $newMenuName -BashPath $newBashPath -IconPath $newIconPath

    # 保存配置
    Save-UserConfig -MenuName $newMenuName -BashPath $newBashPath -IconPath $newIconPath

    Write-Host "`n$separator" -ForegroundColor Green
    Write-Host "                          $(Get-LocalizedString -Key 'WizardUpdateComplete')" -ForegroundColor Green
    Write-Host "$separator`n" -ForegroundColor Green
}

#endregion


# 执行主函数
Main



