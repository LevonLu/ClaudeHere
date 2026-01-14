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
    [string]$Action = 'help'
)

# 脚本配置
$ErrorActionPreference = 'Stop'
# 文件夹背景右键菜单
$RegistryKeyPath_Background = 'Registry::HKEY_CLASSES_ROOT\Directory\Background\shell\ClaudeHere'
# 文件夹本身右键菜单
$RegistryKeyPath_Directory = 'Registry::HKEY_CLASSES_ROOT\Directory\shell\ClaudeHere'
$ConfigRegistryPath = 'Registry::HKEY_CURRENT_USER\Software\ClaudeHere'


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
    Write-Host @"

=============================================================================
                    Claude Here - Windows右键菜单工具
=============================================================================

使用方法:
    .\Install-ClaudeHere.ps1 -Action <操作>

可用操作:
    install    安装右键菜单（交互式配置）
    uninstall  卸载右键菜单
    update     更新现有配置
    help       显示此帮助信息

示例:
    .\Install-ClaudeHere.ps1 -Action install
    .\Install-ClaudeHere.ps1 -Action uninstall
    .\Install-ClaudeHere.ps1 -Action update

注意事项:
    - 需要管理员权限运行
    - 需要安装 Windows Terminal
    - 需要安装 Git for Windows
    - Claude 命令需要在系统 PATH 中可用

=============================================================================
"@ -ForegroundColor Cyan
}

#endregion


#region 权限和依赖检查

function Test-AdminPrivileges {
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

    if (-not $isAdmin) {
        Write-Host "[错误] 此脚本需要管理员权限运行。" -ForegroundColor Red
        Write-Host "请右键点击脚本，选择'以管理员身份运行'。" -ForegroundColor Yellow
        exit 1
    }
}

function Test-WindowsTerminal {
    $wtPath = Join-Path $env:LOCALAPPDATA "Microsoft\WindowsApps\wt.exe"

    if (-not (Test-Path $wtPath)) {
        Write-Host "[错误] 未找到 Windows Terminal。" -ForegroundColor Red
        Write-Host "请先从 Microsoft Store 安装 Windows Terminal。" -ForegroundColor Yellow
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
        Write-Host "[警告] 未在系统 PATH 中找到 claude 命令。" -ForegroundColor Yellow
        Write-Host "安装后需要确保 Claude Code 已安装并添加到 PATH。" -ForegroundColor Yellow
        $response = Read-Host "是否继续安装？(y/N)"

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
    Write-Host "`n[信息] 正在创建注册表项..." -ForegroundColor Cyan

    try {
        # 定义所有需要创建的注册表位置
        $registryPaths = @(
            @{ Path = $RegistryKeyPath_Background; Arg = '%V'; Desc = '文件夹背景' },
            @{ Path = $RegistryKeyPath_Directory;   Arg = '%1'; Desc = '文件夹' }
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

            Write-Host "[成功] 已创建 $desc 右键菜单" -ForegroundColor Green
        }

        Write-Host "`n[成功] 右键菜单已全部创建！" -ForegroundColor Green
        Write-Host "  菜单名称: $MenuName" -ForegroundColor Gray
        Write-Host "  Git Bash: $BashPath" -ForegroundColor Gray
        if ($IconPath) {
            Write-Host "  图标: $IconPath" -ForegroundColor Gray
        }
    }
    catch {
        Write-Host "[错误] 创建注册表项失败: $_" -ForegroundColor Red
        exit 1
    }
}

function Remove-ClaudeHereRegistry {
    Write-Host "`n[信息] 正在删除注册表项..." -ForegroundColor Cyan

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
            Write-Host "[提示] 已清理旧的驱动器菜单项" -ForegroundColor Yellow
        }

        if ($deletedCount -gt 0) {
            Write-Host "[成功] 已删除 $deletedCount 个右键菜单项！" -ForegroundColor Green
        }
        else {
            Write-Host "[提示] 未找到已安装的右键菜单" -ForegroundColor Yellow
        }
    }
    catch {
        Write-Host "[错误] 删除注册表项失败: $_" -ForegroundColor Red
        exit 1
    }
}

#endregion


#region 安装功能

function Install-ClaudeHere {
    Write-Host "`n=============================================================================" -ForegroundColor Cyan
    Write-Host "                    Claude Here - 安装向导" -ForegroundColor Cyan
    Write-Host "=============================================================================`n" -ForegroundColor Cyan

    # 检查 Windows Terminal
    if (-not (Test-WindowsTerminal)) {
        exit 1
    }

    # 检查 Claude 命令
    if (-not (Test-ClaudeCommand)) {
        exit 1
    }

    # 菜单名称
    $menuName = Read-Host "`n请输入菜单显示名称（直接回车使用默认名称 'Claude Here'）"
    if ([string]::IsNullOrWhiteSpace($menuName)) {
        $menuName = "Claude Here"
    }

    # Git Bash 路径
    Write-Host "`n[信息] 正在检测 Git Bash 安装路径..." -ForegroundColor Cyan
    $bashPath = Find-GitBashPath

    if ($bashPath) {
        Write-Host "[成功] 找到 Git Bash: $bashPath" -ForegroundColor Green
        $useDefault = Read-Host "是否使用此路径？(Y/n)"

        if ($useDefault -ne 'n' -and $useDefault -ne 'N') {
            # 使用检测到的路径
        }
        else {
            $bashPath = Read-Host "请输入 bash.exe 的完整路径"
            if (-not (Test-Path $bashPath)) {
                Write-Host "[错误] 指定的路径不存在" -ForegroundColor Red
                exit 1
            }
        }
    }
    else {
        Write-Host "[警告] 未自动检测到 Git Bash" -ForegroundColor Yellow
        $bashPath = Read-Host "请输入 bash.exe 的完整路径"

        if ([string]::IsNullOrWhiteSpace($bashPath) -or -not (Test-Path $bashPath)) {
            Write-Host "[错误] 必须提供有效的 Git Bash 路径" -ForegroundColor Red
            exit 1
        }
    }

    # 图标路径
    # 自动检测脚本所在目录的 claude.ico
    $defaultIconPath = Join-Path $PSScriptRoot "claude.ico"

    Write-Host "`n[信息] 图标设置（可选）" -ForegroundColor Cyan
    $iconPrompt = "请输入图标文件路径（.ico文件，直接回车"

    if (Test-Path $defaultIconPath) {
        $iconPrompt += "使用默认图标: claude.ico"
        Write-Host "[提示] 在脚本目录找到默认图标: claude.ico" -ForegroundColor Green
    }
    else {
        $iconPrompt += "跳过"
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
            Write-Host "[警告] 图标文件不存在，将不设置图标" -ForegroundColor Yellow
            $iconPath = ''
        }
        elseif (-not $iconPath.EndsWith('.ico')) {
            Write-Host "[警告] 图标文件应该是 .ico 格式，将不设置图标" -ForegroundColor Yellow
            $iconPath = ''
        }
    }

    # 创建注册表项
    Set-ClaudeHereRegistry -MenuName $menuName -BashPath $bashPath -IconPath $iconPath

    # 保存配置到用户注册表
    Save-UserConfig -MenuName $menuName -BashPath $bashPath -IconPath $iconPath

    Write-Host "`n=============================================================================" -ForegroundColor Green
    Write-Host "                          安装完成！" -ForegroundColor Green
    Write-Host "=============================================================================" -ForegroundColor Green
    Write-Host "`n现在你可以在文件夹中右键，选择 '$menuName' 来启动 Claude Code。`n" -ForegroundColor White
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
        Write-Host "[警告] 保存配置失败: $_" -ForegroundColor Yellow
    }
}

#endregion


#region 卸载功能

function Uninstall-ClaudeHere {
    Write-Host "`n=============================================================================" -ForegroundColor Cyan
    Write-Host "                    Claude Here - 卸载向导" -ForegroundColor Cyan
    Write-Host "=============================================================================`n" -ForegroundColor Cyan

    $response = Read-Host "确定要卸载 Claude Here 右键菜单吗？(y/N)"

    if ($response -ne 'y' -and $response -ne 'Y') {
        Write-Host "`n已取消卸载。" -ForegroundColor Yellow
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

    Write-Host "`n=============================================================================" -ForegroundColor Green
    Write-Host "                          卸载完成！" -ForegroundColor Green
    Write-Host "=============================================================================`n" -ForegroundColor Green
}

#endregion


#region 更新功能

function Update-ClaudeHere {
    Write-Host "`n=============================================================================" -ForegroundColor Cyan
    Write-Host "                    Claude Here - 更新配置向导" -ForegroundColor Cyan
    Write-Host "=============================================================================`n" -ForegroundColor Cyan

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
        Write-Host "[错误] 未找到已安装的右键菜单，请先运行安装。" -ForegroundColor Red
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

    Write-Host "[当前配置]" -ForegroundColor Yellow
    Write-Host "  菜单名称: $currentMenuName" -ForegroundColor Gray
    Write-Host "  Git Bash: $currentBashPath" -ForegroundColor Gray
    Write-Host "  图标路径: $(if ($currentIconPath) { $currentIconPath } else { '(未设置)' })" -ForegroundColor Gray

    # 新的菜单名称
    Write-Host "`n[提示] 直接回车保持当前值" -ForegroundColor Cyan
    $newMenuName = Read-Host "`n请输入新的菜单名称"

    if ([string]::IsNullOrWhiteSpace($newMenuName)) {
        $newMenuName = $currentMenuName
    }

    # 新的 Bash 路径
    $newBashPath = Read-Host "`n请输入新的 Git Bash 路径"

    if ([string]::IsNullOrWhiteSpace($newBashPath)) {
        $newBashPath = $currentBashPath
    }
    elseif (-not (Test-Path $newBashPath)) {
        Write-Host "[错误] 指定的路径不存在" -ForegroundColor Red
        exit 1
    }

    # 新的图标路径
    $newIconPath = Read-Host "`n请输入新的图标路径（直接回车保持当前值，输入 'none' 移除图标）"

    if ([string]::IsNullOrWhiteSpace($newIconPath)) {
        $newIconPath = $currentIconPath
    }
    elseif ($newIconPath -eq 'none') {
        $newIconPath = ''
    }
    elseif (-not (Test-Path $newIconPath)) {
        Write-Host "[错误] 图标文件不存在" -ForegroundColor Red
        exit 1
    }

    # 更新注册表
    Set-ClaudeHereRegistry -MenuName $newMenuName -BashPath $newBashPath -IconPath $newIconPath

    # 保存配置
    Save-UserConfig -MenuName $newMenuName -BashPath $newBashPath -IconPath $newIconPath

    Write-Host "`n=============================================================================" -ForegroundColor Green
    Write-Host "                          配置已更新！" -ForegroundColor Green
    Write-Host "=============================================================================`n" -ForegroundColor Green
}

#endregion


# 执行主函数
Main



