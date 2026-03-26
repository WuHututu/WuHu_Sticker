param(
    [switch]$SmokeTest
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName PresentationCore
Add-Type -AssemblyName WindowsBase
Add-Type -AssemblyName System.Xaml
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

if (-not ('StickerHotKeyNative' -as [type])) {
    Add-Type -Language CSharp -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public static class StickerHotKeyNative
{
    [DllImport("user32.dll")]
    public static extern short GetAsyncKeyState(int vKey);
}
"@
}

[System.Windows.Forms.Application]::EnableVisualStyles()

$script:AppRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:SettingsDirectory = Join-Path $script:AppRoot 'data'
$script:SettingsPath = Join-Path $script:SettingsDirectory 'settings.json'
$script:LauncherPath = Join-Path $script:AppRoot 'StickerLauncher.vbs'
$script:StartupRegistryPath = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run'
$script:StartupRegistryName = 'StickerQuickEdit'
$script:HotkeyPollIntervalMs = 20
$script:HotkeyTapWindowMs = 500
$script:TranslationsJson = @'
{
  "en-US": {
    "app_name": "Sticker",
    "main_title": "Sticker",
    "main_copy_all": "Copy All",
    "main_copy_selected": "Copy Selected",
    "main_clear": "Clear",
    "main_settings": "Settings",
    "status_ready": "Ready",
    "status_opened": "Opened. Show: {0}; Hide: {1}",
    "status_hidden": "Editor hidden.",
    "status_nothing_to_copy": "Nothing to copy.",
    "status_all_copied": "All text copied to clipboard.",
    "status_select_first": "Select some text first.",
    "status_selected_copied": "Selected text copied to clipboard.",
    "status_settings_saved": "Settings saved.",
    "status_editor_cleared": "Editor cleared.",
    "status_app_started": "App started. Waiting for hotkey.",
    "settings_title": "Sticker Settings",
    "settings_hint": "Click a box and press the shortcut. Win is not supported.",
    "settings_language": "Language",
    "language_zh": "Chinese",
    "language_en": "English",
    "settings_open_hotkey": "Open Hotkey",
    "settings_double_tap": "Double-tap the last key",
    "settings_same_hotkey": "Use the open hotkey to close the editor",
    "settings_close_hotkey": "Close Hotkey",
    "settings_run_at_login": "Run at login",
    "settings_save": "Save",
    "settings_cancel": "Cancel",
    "tray_show_editor": "Show Editor",
    "tray_settings": "Settings",
    "tray_exit": "Exit",
    "msg_open_hotkey_modifier_required": "The open hotkey must include at least one modifier. Supported modifiers are Ctrl, Alt, and Shift.",
    "msg_open_hotkey_invalid": "The open hotkey is invalid. Use Ctrl, Alt, or Shift as modifiers. The Win key is not supported.",
    "msg_close_hotkey_invalid": "The close hotkey is invalid. Use Ctrl, Alt, or Shift as modifiers. The Win key is not supported.",
    "msg_hotkey_modifier_required": "Hotkeys must include at least one modifier. Supported modifiers are Ctrl, Alt, and Shift.",
    "msg_win_key_not_supported": "The Win key is not supported.",
    "msg_already_running": "Sticker is already running. Use the configured open hotkey to show the editor.",
    "msg_reset_default_hotkey": "The saved open hotkey was invalid. The app has been reset to {0}.",
    "notify_balloon_text": "Running in the background. Open hotkey: {0}",
    "apply_open_hotkey_invalid": "The open hotkey is invalid.",
    "apply_close_hotkey_invalid": "The close hotkey is invalid."
  },
  "zh-CN": {
    "app_name": "\u4fbf\u7b7e\u7f16\u8f91",
    "main_title": "\u4fbf\u7b7e\u7f16\u8f91",
    "main_copy_all": "\u590d\u5236\u5168\u90e8",
    "main_copy_selected": "\u590d\u5236\u9009\u4e2d",
    "main_clear": "\u6e05\u7a7a",
    "main_settings": "\u8bbe\u7f6e",
    "status_ready": "\u5c31\u7eea",
    "status_opened": "\u5df2\u6253\u5f00\u3002\u5524\u51fa\uff1a{0}\uff1b\u5173\u95ed\uff1a{1}",
    "status_hidden": "\u7f16\u8f91\u7a97\u5df2\u9690\u85cf\u3002",
    "status_nothing_to_copy": "\u6ca1\u6709\u53ef\u590d\u5236\u7684\u5185\u5bb9\u3002",
    "status_all_copied": "\u5168\u90e8\u5185\u5bb9\u5df2\u590d\u5236\u5230\u526a\u8d34\u677f\u3002",
    "status_select_first": "\u8bf7\u5148\u9009\u4e2d\u8981\u590d\u5236\u7684\u5185\u5bb9\u3002",
    "status_selected_copied": "\u9009\u4e2d\u5185\u5bb9\u5df2\u590d\u5236\u5230\u526a\u8d34\u677f\u3002",
    "status_settings_saved": "\u8bbe\u7f6e\u5df2\u4fdd\u5b58\u3002",
    "status_editor_cleared": "\u5185\u5bb9\u5df2\u6e05\u7a7a\u3002",
    "status_app_started": "\u7a0b\u5e8f\u5df2\u542f\u52a8\uff0c\u7b49\u5f85\u70ed\u952e\u5524\u51fa\u3002",
    "settings_title": "\u4fbf\u7b7e\u7f16\u8f91\u8bbe\u7f6e",
    "settings_hint": "\u70b9\u51fb\u8f93\u5165\u6846\u540e\u76f4\u63a5\u6309\u4e0b\u5feb\u6377\u952e\u3002\u4e0d\u652f\u6301 Win \u952e\u3002",
    "settings_language": "\u8bed\u8a00",
    "language_zh": "\u4e2d\u6587",
    "language_en": "\u82f1\u6587",
    "settings_open_hotkey": "\u5524\u51fa\u70ed\u952e",
    "settings_double_tap": "\u6700\u540e\u4e00\u952e\u53cc\u51fb\u89e6\u53d1",
    "settings_same_hotkey": "\u4f7f\u7528\u4e0e\u5524\u51fa\u76f8\u540c\u7684\u70ed\u952e\u5173\u95ed\u7f16\u8f91\u7a97",
    "settings_close_hotkey": "\u5173\u95ed\u70ed\u952e",
    "settings_run_at_login": "\u5f00\u673a\u81ea\u542f\u52a8",
    "settings_save": "\u4fdd\u5b58",
    "settings_cancel": "\u53d6\u6d88",
    "tray_show_editor": "\u663e\u793a\u7f16\u8f91\u7a97",
    "tray_settings": "\u8bbe\u7f6e",
    "tray_exit": "\u9000\u51fa",
    "msg_open_hotkey_modifier_required": "\u5524\u51fa\u70ed\u952e\u81f3\u5c11\u9700\u8981\u4e00\u4e2a\u4fee\u9970\u952e\u3002\u53ef\u7528\u4fee\u9970\u952e\u4e3a Ctrl\u3001Alt\u3001Shift\u3002",
    "msg_open_hotkey_invalid": "\u5524\u51fa\u70ed\u952e\u65e0\u6548\u3002\u4fee\u9970\u952e\u53ea\u652f\u6301 Ctrl\u3001Alt\u3001Shift\uff0c\u4e0d\u652f\u6301 Win \u952e\u3002",
    "msg_close_hotkey_invalid": "\u5173\u95ed\u70ed\u952e\u65e0\u6548\u3002\u4fee\u9970\u952e\u53ea\u652f\u6301 Ctrl\u3001Alt\u3001Shift\uff0c\u4e0d\u652f\u6301 Win \u952e\u3002",
    "msg_hotkey_modifier_required": "\u5feb\u6377\u952e\u81f3\u5c11\u9700\u8981\u4e00\u4e2a\u4fee\u9970\u952e\u3002\u53ef\u7528\u4fee\u9970\u952e\u4e3a Ctrl\u3001Alt\u3001Shift\u3002",
    "msg_win_key_not_supported": "\u4e0d\u652f\u6301 Win \u952e\u3002",
    "msg_already_running": "\u7a0b\u5e8f\u5df2\u5728\u540e\u53f0\u8fd0\u884c\uff0c\u8bf7\u76f4\u63a5\u4f7f\u7528\u5f53\u524d\u70ed\u952e\u5524\u51fa\u7f16\u8f91\u7a97\u3002",
    "msg_reset_default_hotkey": "\u5df2\u4fdd\u5b58\u7684\u5524\u51fa\u70ed\u952e\u65e0\u6548\uff0c\u7a0b\u5e8f\u5df2\u6062\u590d\u4e3a {0}\u3002",
    "notify_balloon_text": "\u7a0b\u5e8f\u5df2\u5728\u540e\u53f0\u8fd0\u884c\u3002\u5524\u51fa\u70ed\u952e\uff1a{0}",
    "apply_open_hotkey_invalid": "\u5524\u51fa\u70ed\u952e\u65e0\u6548\u3002",
    "apply_close_hotkey_invalid": "\u5173\u95ed\u70ed\u952e\u65e0\u6548\u3002"
  }
}
'@
$script:Translations = $script:TranslationsJson | ConvertFrom-Json
$script:State = @{
    Application = $null
    MainWindow = $null
    SettingsWindow = $null
    NotifyIcon = $null
    EditorTextBox = $null
    StatusTextBlock = $null
    CopyAllButton = $null
    CopySelectionButton = $null
    ClearButton = $null
    SettingsButton = $null
    InvokeHotkeyTextBox = $null
    CloseHotkeyTextBox = $null
    SameHotkeyCheckBox = $null
    StartupCheckBox = $null
    InvokeDoubleTapCheckBox = $null
    CloseDoubleTapCheckBox = $null
    LanguageLabelTextBlock = $null
    LanguageComboBox = $null
    LanguageZhItem = $null
    LanguageEnItem = $null
    SettingsHintTextBlock = $null
    OpenHotkeyLabel = $null
    CloseHotkeyLabel = $null
    SaveSettingsButton = $null
    CancelSettingsButton = $null
    TrayShowItem = $null
    TraySettingsItem = $null
    TrayExitItem = $null
    WindowVisible = $false
    IsExiting = $false
    Settings = $null
    PendingInvokeHotkey = $null
    PendingCloseHotkey = $null
    Mutex = $null
    HotkeyTimer = $null
    PressedKeys = @{}
    LastStatusKey = 'status_ready'
    LastStatusArgs = @()
    HotkeyTrackers = @{
        Invoke = @{
            Signature = ''
            Count = 0
            Timestamp = [datetime]::MinValue
        }
        Close = @{
            Signature = ''
            Count = 0
            Timestamp = [datetime]::MinValue
        }
    }
}

function Get-DefaultLanguage {
    $cultureName = [System.Globalization.CultureInfo]::CurrentUICulture.Name
    if ($cultureName -like 'zh*') {
        return 'zh-CN'
    }

    return 'en-US'
}

function Get-ActiveLanguage {
    if ($null -ne $script:State.Settings -and -not [string]::IsNullOrWhiteSpace([string]$script:State.Settings.Language)) {
        return [string]$script:State.Settings.Language
    }

    return Get-DefaultLanguage
}

function Get-Text {
    param(
        [string]$Key,
        [object[]]$Args = @(),
        [string]$Language = ''
    )

    if ([string]::IsNullOrWhiteSpace($Language)) {
        $Language = Get-ActiveLanguage
    }

    $languagePack = $script:Translations.PSObject.Properties[$Language].Value
    if ($null -eq $languagePack) {
        $languagePack = $script:Translations.PSObject.Properties['en-US'].Value
    }

    $text = $languagePack.PSObject.Properties[$Key].Value
    if ($null -eq $text) {
        $text = $script:Translations.PSObject.Properties['en-US'].Value.PSObject.Properties[$Key].Value
    }

    if ($Args.Count -gt 0) {
        return [string]::Format($text, $Args)
    }

    return [string]$text
}

function Set-LocalizedStatus {
    param(
        [string]$Key,
        [object[]]$Args = @()
    )

    $script:State.LastStatusKey = $Key
    $script:State.LastStatusArgs = @($Args)
    Set-Status (Get-Text -Key $Key -Args $Args)
}

function Refresh-StatusText {
    if ([string]::IsNullOrWhiteSpace([string]$script:State.LastStatusKey)) {
        return
    }

    Set-Status (Get-Text -Key $script:State.LastStatusKey -Args $script:State.LastStatusArgs)
}

function Show-AppMessage {
    param(
        [string]$Message,
        [string]$Language = ''
    )

    [System.Windows.MessageBox]::Show($Message, (Get-Text -Key 'app_name' -Language $Language)) | Out-Null
}

function New-DefaultHotkey {
    param(
        [string[]]$Modifiers = @('Ctrl'),
        [string]$Key = 'Shift',
        [int]$TapCount = 2
    )

    return @{
        Modifiers = @($Modifiers)
        Key = $Key
        TapCount = $TapCount
    }
}

function New-DefaultSettings {
    return @{
        InvokeHotkey = New-DefaultHotkey
        CloseHotkey = @{
            UseInvokeHotkey = $true
            Modifiers = @('Ctrl')
            Key = 'Shift'
            TapCount = 2
        }
        Language = Get-DefaultLanguage
        StartWithWindows = $false
        Window = @{
            Width = 760
            Height = 460
        }
    }
}

function Copy-Hotkey {
    param(
        [hashtable]$Hotkey
    )

    return @{
        Modifiers = @($Hotkey.Modifiers)
        Key = [string]$Hotkey.Key
        TapCount = if ($Hotkey.TapCount) { [int]$Hotkey.TapCount } else { 1 }
    }
}

function Copy-Settings {
    param(
        [hashtable]$Settings
    )

    return @{
        InvokeHotkey = Copy-Hotkey $Settings.InvokeHotkey
        CloseHotkey = @{
            UseInvokeHotkey = [bool]$Settings.CloseHotkey.UseInvokeHotkey
            Modifiers = @($Settings.CloseHotkey.Modifiers)
            Key = [string]$Settings.CloseHotkey.Key
            TapCount = if ($Settings.CloseHotkey.TapCount) { [int]$Settings.CloseHotkey.TapCount } else { 1 }
        }
        Language = if ([string]::IsNullOrWhiteSpace([string]$Settings.Language)) { Get-DefaultLanguage } else { [string]$Settings.Language }
        StartWithWindows = [bool]$Settings.StartWithWindows
        Window = @{
            Width = [int]$Settings.Window.Width
            Height = [int]$Settings.Window.Height
        }
    }
}

function Get-RegistryStartupCommand {
    $wscriptPath = Join-Path $env:SystemRoot 'System32\wscript.exe'
    return ('"{0}" "{1}"' -f $wscriptPath, $script:LauncherPath)
}

function Test-StartWithWindows {
    if (-not (Test-Path $script:StartupRegistryPath)) {
        return $false
    }

    try {
        $value = (Get-ItemProperty -Path $script:StartupRegistryPath -Name $script:StartupRegistryName -ErrorAction Stop).$script:StartupRegistryName
        return [string]::Equals($value, (Get-RegistryStartupCommand), [System.StringComparison]::OrdinalIgnoreCase)
    }
    catch {
        return $false
    }
}

function Set-StartupState {
    param(
        [bool]$Enabled
    )

    if ($Enabled) {
        New-Item -Path $script:StartupRegistryPath -Force | Out-Null
        Set-ItemProperty -Path $script:StartupRegistryPath -Name $script:StartupRegistryName -Value (Get-RegistryStartupCommand)
        return
    }

    if (-not (Test-Path $script:StartupRegistryPath)) {
        return
    }

    $currentProperties = (Get-ItemProperty -Path $script:StartupRegistryPath).PSObject.Properties.Name
    if ($currentProperties -contains $script:StartupRegistryName) {
        Remove-ItemProperty -Path $script:StartupRegistryPath -Name $script:StartupRegistryName
    }
}

function Resolve-KeyName {
    param(
        [string]$Name
    )

    try {
        return [System.Windows.Input.Key]::$Name
    }
    catch {
        return $null
    }
}

function Normalize-Hotkey {
    param(
        [hashtable]$Hotkey
    )

    $normalizedModifiers = @()
    foreach ($modifier in @('Ctrl', 'Alt', 'Shift')) {
        if (@($Hotkey.Modifiers) -contains $modifier) {
            $normalizedModifiers += $modifier
        }
    }

    $tapCount = 1
    if ($Hotkey.ContainsKey('TapCount') -and ($Hotkey.TapCount -as [int])) {
        $tapCount = [int]$Hotkey.TapCount
    }

    if ($tapCount -lt 1) {
        $tapCount = 1
    }
    if ($tapCount -gt 2) {
        $tapCount = 2
    }

    return @{
        Modifiers = $normalizedModifiers
        Key = [string]$Hotkey.Key
        TapCount = $tapCount
    }
}

function Test-HotkeyValid {
    param(
        [hashtable]$Hotkey
    )

    if (-not $Hotkey) {
        return $false
    }

    if ([string]::IsNullOrWhiteSpace([string]$Hotkey.Key)) {
        return $false
    }

    if (@($Hotkey.Modifiers).Count -eq 0) {
        return $false
    }

    $normalized = Normalize-Hotkey $Hotkey
    if ($normalized.TapCount -lt 1 -or $normalized.TapCount -gt 2) {
        return $false
    }

    if (@('Ctrl', 'Alt', 'Shift') -contains $normalized.Key) {
        return (@($normalized.Modifiers) -notcontains $normalized.Key) -and (@($normalized.Modifiers).Count -ge 1)
    }

    return $null -ne (Resolve-KeyName $normalized.Key)
}

function Test-HotkeyCaptureIncomplete {
    param(
        [hashtable]$Hotkey
    )

    if (-not $Hotkey) {
        return $true
    }

    if ([string]::IsNullOrWhiteSpace([string]$Hotkey.Key)) {
        return $true
    }

    $normalized = Normalize-Hotkey $Hotkey
    if ((Test-IsModifierKeyName $normalized.Key) -and (@($normalized.Modifiers).Count -eq 0)) {
        return $true
    }

    return $false
}

function Get-EffectiveCloseHotkey {
    param(
        [hashtable]$Settings
    )

    if ($Settings.CloseHotkey.UseInvokeHotkey) {
        return Copy-Hotkey $Settings.InvokeHotkey
    }

    return Copy-Hotkey $Settings.CloseHotkey
}

function Get-HotkeySignature {
    param(
        [hashtable]$Hotkey
    )

    $normalized = Normalize-Hotkey $Hotkey
    return ('{0}:{1}:{2}' -f ([string]::Join('+', $normalized.Modifiers)), $normalized.Key, $normalized.TapCount)
}

function Test-HotkeysEqual {
    param(
        [hashtable]$Left,
        [hashtable]$Right
    )

    return (Get-HotkeySignature $Left) -eq (Get-HotkeySignature $Right)
}

function Test-IsModifierKeyName {
    param(
        [string]$KeyName
    )

    return @('Ctrl', 'Alt', 'Shift') -contains $KeyName
}

function Get-LogicalKeyVirtualKeyCode {
    param(
        [string]$KeyName
    )

    switch ($KeyName) {
        'Ctrl' { return 0x11 }
        'Alt' { return 0x12 }
        'Shift' { return 0x10 }
    }

    $key = Resolve-KeyName $KeyName
    if ($null -eq $key) {
        throw "Unknown key: $KeyName"
    }

    return [uint32][System.Windows.Input.KeyInterop]::VirtualKeyFromKey($key)
}

function Format-KeyDisplayName {
    param(
        [string]$Key
    )

    $map = @{
        D0 = '0'
        D1 = '1'
        D2 = '2'
        D3 = '3'
        D4 = '4'
        D5 = '5'
        D6 = '6'
        D7 = '7'
        D8 = '8'
        D9 = '9'
        Return = 'Enter'
        Prior = 'Page Up'
        Next = 'Page Down'
        Capital = 'Caps Lock'
        Escape = 'Esc'
    }

    if ($map.ContainsKey($Key)) {
        return $map[$Key]
    }

    if ($Key -match '^NumPad(\d)$') {
        return "Num $($Matches[1])"
    }

    return $Key
}

function Format-Hotkey {
    param(
        [hashtable]$Hotkey
    )

    $normalized = Normalize-Hotkey $Hotkey
    $parts = @()
    foreach ($modifier in $normalized.Modifiers) {
        switch ($modifier) {
            'Ctrl' { $parts += 'Ctrl' }
            'Alt' { $parts += 'Alt' }
            'Shift' { $parts += 'Shift' }
        }
    }

    $parts += (Format-KeyDisplayName $normalized.Key)
    if ($normalized.TapCount -eq 2) {
        $parts += (Format-KeyDisplayName $normalized.Key)
    }
    return [string]::Join(' + ', $parts)
}

function Get-HotkeyFromKeyEvent {
    param(
        [System.Windows.Input.KeyEventArgs]$EventArgs
    )

    $key = if ($EventArgs.Key -eq [System.Windows.Input.Key]::System) { $EventArgs.SystemKey } else { $EventArgs.Key }
    if ($key -in @([System.Windows.Input.Key]::LWin, [System.Windows.Input.Key]::RWin)) {
        return $null
    }

    $triggerKey = switch ($key) {
        { $_ -in @([System.Windows.Input.Key]::LeftCtrl, [System.Windows.Input.Key]::RightCtrl) } { 'Ctrl'; break }
        { $_ -in @([System.Windows.Input.Key]::LeftAlt, [System.Windows.Input.Key]::RightAlt) } { 'Alt'; break }
        { $_ -in @([System.Windows.Input.Key]::LeftShift, [System.Windows.Input.Key]::RightShift) } { 'Shift'; break }
        default { $key.ToString() }
    }

    $modifiers = @()
    $modifierState = [System.Windows.Input.Keyboard]::Modifiers
    if (($modifierState -band [System.Windows.Input.ModifierKeys]::Control) -ne 0) {
        $modifiers += 'Ctrl'
    }
    if (($modifierState -band [System.Windows.Input.ModifierKeys]::Alt) -ne 0) {
        $modifiers += 'Alt'
    }
    if (($modifierState -band [System.Windows.Input.ModifierKeys]::Shift) -ne 0) {
        $modifiers += 'Shift'
    }
    if (Test-IsModifierKeyName $triggerKey) {
        $modifiers = @($modifiers | Where-Object { $_ -ne $triggerKey })
    }

    return @{
        Modifiers = $modifiers
        Key = $triggerKey
        TapCount = 1
    }
}

function Ensure-SettingsDirectory {
    if (-not (Test-Path $script:SettingsDirectory)) {
        New-Item -ItemType Directory -Path $script:SettingsDirectory -Force | Out-Null
    }
}

function Save-Settings {
    param(
        [hashtable]$Settings
    )

    Ensure-SettingsDirectory
    $payload = @{
        InvokeHotkey = Normalize-Hotkey $Settings.InvokeHotkey
        CloseHotkey = @{
            UseInvokeHotkey = [bool]$Settings.CloseHotkey.UseInvokeHotkey
            Modifiers = @($Settings.CloseHotkey.Modifiers)
            Key = [string]$Settings.CloseHotkey.Key
            TapCount = if ($Settings.CloseHotkey.TapCount) { [int]$Settings.CloseHotkey.TapCount } else { 1 }
        }
        Language = if ([string]::IsNullOrWhiteSpace([string]$Settings.Language)) { Get-DefaultLanguage } else { [string]$Settings.Language }
        StartWithWindows = [bool]$Settings.StartWithWindows
        Window = @{
            Width = [int]$Settings.Window.Width
            Height = [int]$Settings.Window.Height
        }
    }

    $payload | ConvertTo-Json -Depth 5 | Set-Content -Path $script:SettingsPath -Encoding UTF8
}

function Load-Settings {
    $defaults = New-DefaultSettings
    if (-not (Test-Path $script:SettingsPath)) {
        $defaults.StartWithWindows = Test-StartWithWindows
        return $defaults
    }

    try {
        $raw = Get-Content -Path $script:SettingsPath -Raw | ConvertFrom-Json
        $settings = Copy-Settings $defaults

        if ($null -ne $raw.InvokeHotkey) {
            $candidate = @{
                Modifiers = @($raw.InvokeHotkey.Modifiers)
                Key = [string]$raw.InvokeHotkey.Key
                TapCount = if ($null -ne $raw.InvokeHotkey.TapCount) { [int]$raw.InvokeHotkey.TapCount } else { 1 }
            }
            if (Test-HotkeyValid $candidate) {
                $settings.InvokeHotkey = Normalize-Hotkey $candidate
            }
        }

        if ($null -ne $raw.CloseHotkey) {
            $settings.CloseHotkey.UseInvokeHotkey = [bool]$raw.CloseHotkey.UseInvokeHotkey
            $closeCandidate = @{
                Modifiers = @($raw.CloseHotkey.Modifiers)
                Key = [string]$raw.CloseHotkey.Key
                TapCount = if ($null -ne $raw.CloseHotkey.TapCount) { [int]$raw.CloseHotkey.TapCount } else { 1 }
            }
            if (Test-HotkeyValid $closeCandidate) {
                $normalizedClose = Normalize-Hotkey $closeCandidate
                $settings.CloseHotkey.Modifiers = @($normalizedClose.Modifiers)
                $settings.CloseHotkey.Key = [string]$normalizedClose.Key
                $settings.CloseHotkey.TapCount = [int]$normalizedClose.TapCount
            }
        }

        if ($null -ne $raw.Language -and @('zh-CN', 'en-US') -contains [string]$raw.Language) {
            $settings.Language = [string]$raw.Language
        }

        if ($null -ne $raw.Window) {
            if ($raw.Window.Width -as [int]) {
                $settings.Window.Width = [int]$raw.Window.Width
            }
            if ($raw.Window.Height -as [int]) {
                $settings.Window.Height = [int]$raw.Window.Height
            }
        }

        $settings.StartWithWindows = Test-StartWithWindows
        return $settings
    }
    catch {
        return $defaults
    }
}

function Set-Status {
    param(
        [string]$Message
    )

    if ($null -eq $script:State.StatusTextBlock) {
        return
    }

    $timestamp = Get-Date -Format 'HH:mm:ss'
    $script:State.StatusTextBlock.Text = "[${timestamp}] $Message"
}

function Refresh-MainWindowLanguage {
    if ($null -eq $script:State.MainWindow) {
        return
    }

    $script:State.MainWindow.Title = Get-Text 'main_title'
    if ($null -ne $script:State.CopyAllButton) {
        $script:State.CopyAllButton.Content = Get-Text 'main_copy_all'
    }
    if ($null -ne $script:State.CopySelectionButton) {
        $script:State.CopySelectionButton.Content = Get-Text 'main_copy_selected'
    }
    if ($null -ne $script:State.ClearButton) {
        $script:State.ClearButton.Content = Get-Text 'main_clear'
    }
    if ($null -ne $script:State.SettingsButton) {
        $script:State.SettingsButton.Content = Get-Text 'main_settings'
    }
}

function Refresh-SettingsWindowLanguage {
    if ($null -eq $script:State.SettingsWindow) {
        return
    }

    $script:State.SettingsWindow.Title = Get-Text 'settings_title'
    $script:State.SettingsHintTextBlock.Text = Get-Text 'settings_hint'
    $script:State.LanguageLabelTextBlock.Text = Get-Text 'settings_language'
    $script:State.LanguageZhItem.Content = Get-Text 'language_zh'
    $script:State.LanguageEnItem.Content = Get-Text 'language_en'
    $script:State.OpenHotkeyLabel.Text = Get-Text 'settings_open_hotkey'
    $script:State.InvokeDoubleTapCheckBox.Content = Get-Text 'settings_double_tap'
    $script:State.SameHotkeyCheckBox.Content = Get-Text 'settings_same_hotkey'
    $script:State.CloseHotkeyLabel.Text = Get-Text 'settings_close_hotkey'
    $script:State.CloseDoubleTapCheckBox.Content = Get-Text 'settings_double_tap'
    $script:State.StartupCheckBox.Content = Get-Text 'settings_run_at_login'
    $script:State.SaveSettingsButton.Content = Get-Text 'settings_save'
    $script:State.CancelSettingsButton.Content = Get-Text 'settings_cancel'
}

function Refresh-NotifyIconLanguage {
    if ($null -eq $script:State.NotifyIcon) {
        return
    }

    $script:State.NotifyIcon.Text = Get-Text 'app_name'
    if ($null -ne $script:State.TrayShowItem) {
        $script:State.TrayShowItem.Text = Get-Text 'tray_show_editor'
    }
    if ($null -ne $script:State.TraySettingsItem) {
        $script:State.TraySettingsItem.Text = Get-Text 'tray_settings'
    }
    if ($null -ne $script:State.TrayExitItem) {
        $script:State.TrayExitItem.Text = Get-Text 'tray_exit'
    }
}

function Refresh-LocalizedUI {
    Refresh-MainWindowLanguage
    Refresh-SettingsWindowLanguage
    Refresh-NotifyIconLanguage
    Refresh-StatusText
}

function Select-LanguageItem {
    param(
        [string]$Language
    )

    if ($null -eq $script:State.LanguageComboBox) {
        return
    }

    switch ($Language) {
        'zh-CN' { $script:State.LanguageComboBox.SelectedItem = $script:State.LanguageZhItem }
        default { $script:State.LanguageComboBox.SelectedItem = $script:State.LanguageEnItem }
    }
}

function Get-SelectedLanguage {
    if ($null -eq $script:State.LanguageComboBox -or $null -eq $script:State.LanguageComboBox.SelectedItem) {
        return Get-ActiveLanguage
    }

    return [string]$script:State.LanguageComboBox.SelectedItem.Tag
}

function Save-WindowSize {
    if ($null -eq $script:State.MainWindow) {
        return
    }

    if ($script:State.MainWindow.WindowState -ne [System.Windows.WindowState]::Normal) {
        return
    }

    $script:State.Settings.Window.Width = [int][Math]::Round($script:State.MainWindow.Width)
    $script:State.Settings.Window.Height = [int][Math]::Round($script:State.MainWindow.Height)
    Save-Settings $script:State.Settings
}

function Reset-HotkeyTracker {
    param(
        [string]$Name
    )

    $script:State.HotkeyTrackers[$Name].Signature = ''
    $script:State.HotkeyTrackers[$Name].Count = 0
    $script:State.HotkeyTrackers[$Name].Timestamp = [datetime]::MinValue
}

function Reset-AllHotkeyTrackers {
    Reset-HotkeyTracker -Name 'Invoke'
    Reset-HotkeyTracker -Name 'Close'
}

function Test-KeyPressed {
    param(
        [string]$LogicalKey
    )

    if (-not $script:State.PressedKeys.ContainsKey($LogicalKey)) {
        return $false
    }

    return [bool]$script:State.PressedKeys[$LogicalKey]
}

function Get-TrackedLogicalKeys {
    $keys = New-Object 'System.Collections.Generic.HashSet[string]'
    $hotkeys = @($script:State.Settings.InvokeHotkey, (Get-EffectiveCloseHotkey $script:State.Settings))
    foreach ($hotkey in $hotkeys) {
        $normalized = Normalize-Hotkey $hotkey
        foreach ($modifier in $normalized.Modifiers) {
            [void]$keys.Add($modifier)
        }
        [void]$keys.Add($normalized.Key)
    }

    return @($keys)
}

function Test-LogicalKeyDown {
    param(
        [string]$LogicalKey
    )

    $vk = Get-LogicalKeyVirtualKeyCode $LogicalKey
    return (([int][StickerHotKeyNative]::GetAsyncKeyState($vk)) -band 0x8000) -ne 0
}

function Get-ActiveModifiers {
    $modifiers = @()
    foreach ($modifier in @('Ctrl', 'Alt', 'Shift')) {
        if (Test-KeyPressed $modifier) {
            $modifiers += $modifier
        }
    }

    return $modifiers
}

function Reset-InterruptedHotkeyTrackers {
    param(
        [string]$LogicalKey
    )

    $roleHotkeys = @{
        Invoke = Normalize-Hotkey $script:State.Settings.InvokeHotkey
        Close = Normalize-Hotkey (Get-EffectiveCloseHotkey $script:State.Settings)
    }

    foreach ($role in $roleHotkeys.Keys) {
        $hotkey = $roleHotkeys[$role]
        $allowedKeys = @($hotkey.Modifiers + $hotkey.Key)
        if ($allowedKeys -notcontains $LogicalKey) {
            Reset-HotkeyTracker -Name $role
        }
    }
}

function Test-HotkeyMatch {
    param(
        [string]$Role,
        [string]$LogicalKey,
        [hashtable]$Hotkey
    )

    $normalized = Normalize-Hotkey $Hotkey
    if ($normalized.Key -ne $LogicalKey) {
        return $false
    }

    $activeModifiers = Get-ActiveModifiers
    if (Test-IsModifierKeyName $normalized.Key) {
        $activeModifiers = @($activeModifiers | Where-Object { $_ -ne $normalized.Key })
    }

    if (([string]::Join('|', $activeModifiers)) -ne ([string]::Join('|', $normalized.Modifiers))) {
        Reset-HotkeyTracker -Name $Role
        return $false
    }

    if ($normalized.TapCount -eq 1) {
        Reset-HotkeyTracker -Name $Role
        return $true
    }

    $tracker = $script:State.HotkeyTrackers[$Role]
    $signature = Get-HotkeySignature $normalized
    $now = Get-Date
    $elapsed = ($now - [datetime]$tracker.Timestamp).TotalMilliseconds

    if (($tracker.Signature -eq $signature) -and ($elapsed -le $script:HotkeyTapWindowMs)) {
        $tracker.Count++
    }
    else {
        $tracker.Signature = $signature
        $tracker.Count = 1
    }

    $tracker.Timestamp = $now
    if ($tracker.Count -ge $normalized.TapCount) {
        Reset-HotkeyTracker -Name $Role
        return $true
    }

    return $false
}

function Process-HotkeyKeyDown {
    param(
        [string]$LogicalKey
    )

    Reset-InterruptedHotkeyTrackers -LogicalKey $LogicalKey

    $invokeHotkey = Normalize-Hotkey $script:State.Settings.InvokeHotkey
    $closeHotkey = Normalize-Hotkey (Get-EffectiveCloseHotkey $script:State.Settings)
    if (Test-HotkeysEqual $invokeHotkey $closeHotkey) {
        if (Test-HotkeyMatch -Role 'Invoke' -LogicalKey $LogicalKey -Hotkey $invokeHotkey) {
            Toggle-EditorWindow
        }
        return
    }

    if (Test-HotkeyMatch -Role 'Invoke' -LogicalKey $LogicalKey -Hotkey $invokeHotkey) {
        Show-EditorWindow
        return
    }

    if ($script:State.WindowVisible -and (Test-HotkeyMatch -Role 'Close' -LogicalKey $LogicalKey -Hotkey $closeHotkey)) {
        Hide-EditorWindow
    }
}

function Process-HotkeyKeyUp {
    param(
        [string]$LogicalKey
    )

    $roleHotkeys = @{
        Invoke = Normalize-Hotkey $script:State.Settings.InvokeHotkey
        Close = Normalize-Hotkey (Get-EffectiveCloseHotkey $script:State.Settings)
    }

    foreach ($role in $roleHotkeys.Keys) {
        if ($roleHotkeys[$role].Modifiers -contains $LogicalKey) {
            Reset-HotkeyTracker -Name $role
        }
    }
}

function Sync-HotkeyState {
    if ($null -eq $script:State.Settings) {
        return
    }

    $trackedKeys = Get-TrackedLogicalKeys
    foreach ($logicalKey in $trackedKeys) {
        $isDown = Test-LogicalKeyDown $logicalKey
        $wasDown = Test-KeyPressed $logicalKey
        if ($isDown -and -not $wasDown) {
            $script:State.PressedKeys[$logicalKey] = $true
            Process-HotkeyKeyDown -LogicalKey $logicalKey
        }
        elseif (-not $isDown -and $wasDown) {
            $script:State.PressedKeys[$logicalKey] = $false
            Process-HotkeyKeyUp -LogicalKey $logicalKey
        }
    }

    foreach ($key in @($script:State.PressedKeys.Keys)) {
        if ($trackedKeys -notcontains $key) {
            $script:State.PressedKeys.Remove($key)
        }
    }
}

function Capture-PressedKeySnapshot {
    $snapshot = @{}
    if ($null -ne $script:State.Settings) {
        foreach ($logicalKey in Get-TrackedLogicalKeys) {
            $snapshot[$logicalKey] = Test-LogicalKeyDown $logicalKey
        }
    }

    $script:State.PressedKeys = $snapshot
}

function Apply-Hotkeys {
    param(
        [hashtable]$Settings
    )

    if (-not (Test-HotkeyValid $Settings.InvokeHotkey)) {
        return @{
            Success = $false
            Message = Get-Text 'apply_open_hotkey_invalid'
        }
    }

    $closeHotkey = Get-EffectiveCloseHotkey $Settings
    if (-not (Test-HotkeyValid $closeHotkey)) {
        return @{
            Success = $false
            Message = Get-Text 'apply_close_hotkey_invalid'
        }
    }

    Reset-AllHotkeyTrackers
    Capture-PressedKeySnapshot
    return @{
        Success = $true
        Message = ''
    }
}

function Focus-EditorWindow {
    if ($null -eq $script:State.MainWindow) {
        return
    }

    if (-not $script:State.MainWindow.IsVisible) {
        $script:State.MainWindow.Show()
    }

    $script:State.MainWindow.WindowState = [System.Windows.WindowState]::Normal
    $script:State.MainWindow.Activate() | Out-Null
    $script:State.MainWindow.Topmost = $true
    $script:State.MainWindow.Topmost = $false
    $script:State.EditorTextBox.Focus() | Out-Null
    $script:State.EditorTextBox.CaretIndex = $script:State.EditorTextBox.Text.Length
}

function Show-EditorWindow {
    if ($script:State.WindowVisible) {
        Focus-EditorWindow
        return
    }

    $script:State.WindowVisible = $true
    $result = Apply-Hotkeys $script:State.Settings
    if (-not $result.Success) {
        $script:State.WindowVisible = $false
        Apply-Hotkeys $script:State.Settings | Out-Null
        Show-AppMessage $result.Message
        return
    }

    Focus-EditorWindow
    Set-LocalizedStatus 'status_opened' @((Format-Hotkey $script:State.Settings.InvokeHotkey), (Format-Hotkey (Get-EffectiveCloseHotkey $script:State.Settings)))
}

function Hide-EditorWindow {
    if (-not $script:State.WindowVisible) {
        return
    }

    Save-WindowSize
    $script:State.WindowVisible = $false
    if ($null -ne $script:State.SettingsWindow -and $script:State.SettingsWindow.IsVisible) {
        $script:State.SettingsWindow.Close()
    }
    $script:State.MainWindow.Hide()
    Apply-Hotkeys $script:State.Settings | Out-Null
    Set-LocalizedStatus 'status_hidden'
}

function Toggle-EditorWindow {
    $closeHotkey = Get-EffectiveCloseHotkey $script:State.Settings
    if ($script:State.WindowVisible -and (Test-HotkeysEqual $script:State.Settings.InvokeHotkey $closeHotkey)) {
        Hide-EditorWindow
        return
    }

    Show-EditorWindow
}

function Copy-AllText {
    $text = $script:State.EditorTextBox.Text
    if ([string]::IsNullOrEmpty($text)) {
        Set-LocalizedStatus 'status_nothing_to_copy'
        return
    }

    [System.Windows.Clipboard]::SetText($text)
    Set-LocalizedStatus 'status_all_copied'
}

function Copy-SelectedText {
    $text = $script:State.EditorTextBox.SelectedText
    if ([string]::IsNullOrEmpty($text)) {
        Set-LocalizedStatus 'status_select_first'
        return
    }

    [System.Windows.Clipboard]::SetText($text)
    Set-LocalizedStatus 'status_selected_copied'
}

function Update-SettingsInputs {
    $settings = $script:State.Settings
    $effectiveClose = Get-EffectiveCloseHotkey $settings

    $script:State.PendingInvokeHotkey = Copy-Hotkey $settings.InvokeHotkey
    $script:State.PendingCloseHotkey = Copy-Hotkey $effectiveClose
    $script:State.SameHotkeyCheckBox.IsChecked = [bool]$settings.CloseHotkey.UseInvokeHotkey
    $script:State.CloseHotkeyTextBox.IsEnabled = -not [bool]$settings.CloseHotkey.UseInvokeHotkey
    $script:State.InvokeHotkeyTextBox.Text = Format-Hotkey $script:State.PendingInvokeHotkey
    $script:State.CloseHotkeyTextBox.Text = Format-Hotkey $script:State.PendingCloseHotkey
    $script:State.StartupCheckBox.IsChecked = [bool]$settings.StartWithWindows
    $script:State.InvokeDoubleTapCheckBox.IsChecked = ([int]$script:State.PendingInvokeHotkey.TapCount -eq 2)
    $script:State.CloseDoubleTapCheckBox.IsChecked = ([int]$script:State.PendingCloseHotkey.TapCount -eq 2)
    $script:State.CloseDoubleTapCheckBox.IsEnabled = -not [bool]$settings.CloseHotkey.UseInvokeHotkey
    Select-LanguageItem -Language $settings.Language
}

function Save-SettingsFromDialog {
    $newSettings = Copy-Settings $script:State.Settings
    $selectedLanguage = Get-SelectedLanguage

    if (-not (Test-HotkeyValid $script:State.PendingInvokeHotkey)) {
        Show-AppMessage (Get-Text -Key 'msg_open_hotkey_modifier_required' -Language $selectedLanguage) -Language $selectedLanguage
        return
    }

    $invokeHotkey = Copy-Hotkey $script:State.PendingInvokeHotkey
    $invokeHotkey.TapCount = if ($script:State.InvokeDoubleTapCheckBox.IsChecked) { 2 } else { 1 }
    if (-not (Test-HotkeyValid $invokeHotkey)) {
        Show-AppMessage (Get-Text -Key 'msg_open_hotkey_invalid' -Language $selectedLanguage) -Language $selectedLanguage
        return
    }

    $newSettings.InvokeHotkey = Normalize-Hotkey $invokeHotkey
    $newSettings.CloseHotkey.UseInvokeHotkey = [bool]$script:State.SameHotkeyCheckBox.IsChecked

    if ($newSettings.CloseHotkey.UseInvokeHotkey) {
        $newSettings.CloseHotkey.Modifiers = @($newSettings.InvokeHotkey.Modifiers)
        $newSettings.CloseHotkey.Key = [string]$newSettings.InvokeHotkey.Key
        $newSettings.CloseHotkey.TapCount = [int]$newSettings.InvokeHotkey.TapCount
    }
    else {
        $closeHotkey = Copy-Hotkey $script:State.PendingCloseHotkey
        $closeHotkey.TapCount = if ($script:State.CloseDoubleTapCheckBox.IsChecked) { 2 } else { 1 }
        if (-not (Test-HotkeyValid $closeHotkey)) {
            Show-AppMessage (Get-Text -Key 'msg_close_hotkey_invalid' -Language $selectedLanguage) -Language $selectedLanguage
            return
        }

        $normalizedClose = Normalize-Hotkey $closeHotkey
        $newSettings.CloseHotkey.Modifiers = @($normalizedClose.Modifiers)
        $newSettings.CloseHotkey.Key = [string]$normalizedClose.Key
        $newSettings.CloseHotkey.TapCount = [int]$normalizedClose.TapCount
    }

    $newSettings.StartWithWindows = [bool]$script:State.StartupCheckBox.IsChecked
    $newSettings.Language = Get-SelectedLanguage

    $previous = Copy-Settings $script:State.Settings
    $script:State.Settings = $newSettings
    try {
        $result = Apply-Hotkeys $script:State.Settings
        if (-not $result.Success) {
            throw $result.Message
        }

        Set-StartupState -Enabled $newSettings.StartWithWindows
        Save-WindowSize
        Save-Settings $script:State.Settings
        Refresh-LocalizedUI
        Update-SettingsInputs
        Set-LocalizedStatus 'status_settings_saved'
        $script:State.SettingsWindow.Close()
    }
    catch {
        $script:State.Settings = $previous
        Apply-Hotkeys $script:State.Settings | Out-Null
        Refresh-LocalizedUI
        Show-AppMessage ([string]$_.Exception.Message)
    }
}

function Open-SettingsWindow {
    if ($null -ne $script:State.SettingsWindow -and $script:State.SettingsWindow.IsVisible) {
        $script:State.SettingsWindow.Activate() | Out-Null
        return
    }

    New-SettingsWindow
    Update-SettingsInputs
    if ($script:State.MainWindow.IsVisible) {
        $script:State.SettingsWindow.Owner = $script:State.MainWindow
        $script:State.SettingsWindow.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterOwner
    }
    else {
        $script:State.SettingsWindow.WindowStartupLocation = [System.Windows.WindowStartupLocation]::CenterScreen
    }
    $script:State.SettingsWindow.Show()
    $script:State.SettingsWindow.Activate() | Out-Null
}

function Exit-Application {
    if ($script:State.IsExiting) {
        return
    }

    $script:State.IsExiting = $true
    Save-WindowSize
    Save-Settings $script:State.Settings

    if ($null -ne $script:State.HotkeyTimer) {
        $script:State.HotkeyTimer.Stop()
    }

    if ($null -ne $script:State.NotifyIcon) {
        $script:State.NotifyIcon.Visible = $false
        $script:State.NotifyIcon.Dispose()
    }

    if ($null -ne $script:State.SettingsWindow) {
        $script:State.SettingsWindow.Close()
    }

    if ($null -ne $script:State.MainWindow) {
        $script:State.MainWindow.Close()
    }

    if ($null -ne $script:State.Application) {
        $script:State.Application.Shutdown()
    }

    if ($null -ne $script:State.Mutex) {
        $script:State.Mutex.ReleaseMutex()
        $script:State.Mutex.Dispose()
    }
}

function New-MainWindow {
    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Sticker"
        Width="760"
        Height="460"
        MinWidth="520"
        MinHeight="320"
        WindowStartupLocation="CenterScreen"
        ResizeMode="CanResizeWithGrip"
        ShowInTaskbar="False"
        Background="#FFF4EEE2"
        FontFamily="Segoe UI"
        SnapsToDevicePixels="True">
    <Grid Margin="14">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="*" />
            <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>

        <Border Grid.Row="0" Background="#FFFDF9F1" BorderBrush="#E2D5BF" BorderThickness="1" CornerRadius="12" Padding="10" Margin="0,0,0,12">
            <DockPanel LastChildFill="False">
                <StackPanel Orientation="Horizontal" DockPanel.Dock="Left">
                    <Button x:Name="CopyAllButton" Content="Copy All" Width="96" Height="34" Margin="0,0,8,0" Background="#183A37" Foreground="White" BorderBrush="#183A37" />
                    <Button x:Name="CopySelectionButton" Content="Copy Selected" Width="96" Height="34" Margin="0,0,8,0" Background="#35605A" Foreground="White" BorderBrush="#35605A" />
                    <Button x:Name="ClearButton" Content="Clear" Width="72" Height="34" Background="#F1E3C8" Foreground="#4A4033" BorderBrush="#D4C09C" />
                </StackPanel>
                <Button x:Name="SettingsButton" Content="Settings" Width="72" Height="34" Background="#F8F4EA" Foreground="#3C3228" BorderBrush="#D4C09C" />
            </DockPanel>
        </Border>

        <Border Grid.Row="1" Background="#FFFFFCF6" BorderBrush="#E2D5BF" BorderThickness="1" CornerRadius="14">
            <TextBox x:Name="EditorTextBox"
                     Margin="10"
                     BorderThickness="0"
                     Background="Transparent"
                     FontFamily="Consolas"
                     FontSize="15"
                     AcceptsReturn="True"
                     AcceptsTab="True"
                     TextWrapping="Wrap"
                     VerticalScrollBarVisibility="Auto"
                     HorizontalScrollBarVisibility="Disabled"
                     SpellCheck.IsEnabled="False" />
        </Border>

        <Border Grid.Row="2" Background="#FFFDF9F1" BorderBrush="#E2D5BF" BorderThickness="1" CornerRadius="10" Padding="10" Margin="0,12,0,0">
            <TextBlock x:Name="StatusTextBlock" Foreground="#5C5345" FontSize="12" Text="Ready" />
        </Border>
    </Grid>
</Window>
"@

    $reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)
    $window.Width = [double]$script:State.Settings.Window.Width
    $window.Height = [double]$script:State.Settings.Window.Height

    $script:State.MainWindow = $window
    $script:State.EditorTextBox = $window.FindName('EditorTextBox')
    $script:State.StatusTextBlock = $window.FindName('StatusTextBlock')
    $script:State.CopyAllButton = $window.FindName('CopyAllButton')
    $script:State.CopySelectionButton = $window.FindName('CopySelectionButton')
    $script:State.ClearButton = $window.FindName('ClearButton')
    $script:State.SettingsButton = $window.FindName('SettingsButton')

    $script:State.CopyAllButton.Add_Click({ Copy-AllText })
    $script:State.CopySelectionButton.Add_Click({ Copy-SelectedText })
    $script:State.ClearButton.Add_Click({
        $script:State.EditorTextBox.Clear()
        Set-LocalizedStatus 'status_editor_cleared'
    })
    $script:State.SettingsButton.Add_Click({ Open-SettingsWindow })
    Refresh-MainWindowLanguage
    Refresh-StatusText

    $window.Add_Closing({
        param($sender, $eventArgs)
        if (-not $script:State.IsExiting) {
            $eventArgs.Cancel = $true
            Hide-EditorWindow
        }
    })

    $window.Add_PreviewKeyDown({
        param($sender, $eventArgs)
        if ($eventArgs.Key -eq [System.Windows.Input.Key]::Escape) {
            $eventArgs.Handled = $true
            Hide-EditorWindow
        }
    })
}

function New-SettingsWindow {
    $xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Sticker Settings"
        Width="420"
        Height="450"
        ResizeMode="NoResize"
        ShowInTaskbar="False"
        Background="#FFF8F4EA"
        FontFamily="Segoe UI">
    <Grid Margin="18">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="Auto" />
            <RowDefinition Height="*" />
            <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>

        <TextBlock x:Name="SettingsHintTextBlock" Grid.Row="0" Text="Click a box and press the shortcut. Win is not supported." FontSize="13" TextWrapping="Wrap" Foreground="#5C5345" Margin="0,0,0,14" />

        <TextBlock x:Name="LanguageLabelTextBlock" Grid.Row="1" Text="Language" FontWeight="SemiBold" Foreground="#3C3228" />
        <ComboBox x:Name="LanguageComboBox" Grid.Row="2" Height="34" Margin="0,6,0,12" Background="#FFFFFCF6" BorderBrush="#D4C09C">
            <ComboBoxItem x:Name="LanguageZhItem" Tag="zh-CN">Chinese</ComboBoxItem>
            <ComboBoxItem x:Name="LanguageEnItem" Tag="en-US">English</ComboBoxItem>
        </ComboBox>

        <TextBlock x:Name="OpenHotkeyLabelTextBlock" Grid.Row="3" Text="Open Hotkey" FontWeight="SemiBold" Foreground="#3C3228" />
        <TextBox x:Name="InvokeHotkeyTextBox" Grid.Row="4" Height="34" Margin="0,6,0,8" IsReadOnly="True" VerticalContentAlignment="Center" Background="#FFFFFCF6" BorderBrush="#D4C09C" />
        <CheckBox x:Name="InvokeDoubleTapCheckBox" Grid.Row="5" Content="Double-tap the last key" Margin="0,0,0,12" Foreground="#3C3228" />

        <CheckBox x:Name="SameHotkeyCheckBox" Grid.Row="6" Content="Use the open hotkey to close the editor" Margin="0,0,0,10" Foreground="#3C3228" />

        <TextBlock x:Name="CloseHotkeyLabelTextBlock" Grid.Row="7" Text="Close Hotkey" FontWeight="SemiBold" Foreground="#3C3228" />
        <TextBox x:Name="CloseHotkeyTextBox" Grid.Row="8" Height="34" Margin="0,6,0,8" IsReadOnly="True" VerticalContentAlignment="Center" Background="#FFFFFCF6" BorderBrush="#D4C09C" />
        <CheckBox x:Name="CloseDoubleTapCheckBox" Grid.Row="9" Content="Double-tap the last key" Margin="0,0,0,14" Foreground="#3C3228" />

        <StackPanel Grid.Row="10">
            <CheckBox x:Name="StartupCheckBox" Content="Run at login" Margin="0,0,0,16" Foreground="#3C3228" />
            <DockPanel LastChildFill="False">
                <Button x:Name="SaveSettingsButton" Content="Save" Width="80" Height="34" Margin="0,0,8,0" Background="#183A37" Foreground="White" BorderBrush="#183A37" DockPanel.Dock="Right" />
                <Button x:Name="CancelSettingsButton" Content="Cancel" Width="80" Height="34" Background="#F1E3C8" Foreground="#4A4033" BorderBrush="#D4C09C" DockPanel.Dock="Right" />
            </DockPanel>
        </StackPanel>
    </Grid>
</Window>
"@

    $reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)
    $script:State.SettingsWindow = $window
    $script:State.SettingsHintTextBlock = $window.FindName('SettingsHintTextBlock')
    $script:State.LanguageLabelTextBlock = $window.FindName('LanguageLabelTextBlock')
    $script:State.LanguageComboBox = $window.FindName('LanguageComboBox')
    $script:State.LanguageZhItem = $window.FindName('LanguageZhItem')
    $script:State.LanguageEnItem = $window.FindName('LanguageEnItem')
    $script:State.OpenHotkeyLabel = $window.FindName('OpenHotkeyLabelTextBlock')
    $script:State.InvokeHotkeyTextBox = $window.FindName('InvokeHotkeyTextBox')
    $script:State.CloseHotkeyTextBox = $window.FindName('CloseHotkeyTextBox')
    $script:State.SameHotkeyCheckBox = $window.FindName('SameHotkeyCheckBox')
    $script:State.CloseHotkeyLabel = $window.FindName('CloseHotkeyLabelTextBlock')
    $script:State.StartupCheckBox = $window.FindName('StartupCheckBox')
    $script:State.InvokeDoubleTapCheckBox = $window.FindName('InvokeDoubleTapCheckBox')
    $script:State.CloseDoubleTapCheckBox = $window.FindName('CloseDoubleTapCheckBox')
    $script:State.SaveSettingsButton = $window.FindName('SaveSettingsButton')
    $script:State.CancelSettingsButton = $window.FindName('CancelSettingsButton')
    Refresh-SettingsWindowLanguage

    $captureHandler = {
        param($sender, $eventArgs)
        $messageLanguage = Get-SelectedLanguage
        $hotkey = Get-HotkeyFromKeyEvent $eventArgs
        $eventArgs.Handled = $true
        if ($null -eq $hotkey) {
            if ($eventArgs.Key -in @([System.Windows.Input.Key]::LWin, [System.Windows.Input.Key]::RWin)) {
                Show-AppMessage (Get-Text -Key 'msg_win_key_not_supported' -Language $messageLanguage) -Language $messageLanguage
            }
            return
        }

        if (Test-HotkeyCaptureIncomplete $hotkey) {
            return
        }

        if (-not (Test-HotkeyValid $hotkey)) {
            Show-AppMessage (Get-Text -Key 'msg_hotkey_modifier_required' -Language $messageLanguage) -Language $messageLanguage
            return
        }

        if ($sender.Name -eq 'InvokeHotkeyTextBox') {
            $script:State.PendingInvokeHotkey = Normalize-Hotkey $hotkey
            $script:State.PendingInvokeHotkey.TapCount = if ($script:State.InvokeDoubleTapCheckBox.IsChecked) { 2 } else { 1 }
            $script:State.InvokeHotkeyTextBox.Text = Format-Hotkey $script:State.PendingInvokeHotkey
            if ($script:State.SameHotkeyCheckBox.IsChecked) {
                $script:State.PendingCloseHotkey = Copy-Hotkey $script:State.PendingInvokeHotkey
                $script:State.CloseHotkeyTextBox.Text = Format-Hotkey $script:State.PendingCloseHotkey
                $script:State.CloseDoubleTapCheckBox.IsChecked = $script:State.InvokeDoubleTapCheckBox.IsChecked
            }
            return
        }

        $script:State.PendingCloseHotkey = Normalize-Hotkey $hotkey
        $script:State.PendingCloseHotkey.TapCount = if ($script:State.CloseDoubleTapCheckBox.IsChecked) { 2 } else { 1 }
        $script:State.CloseHotkeyTextBox.Text = Format-Hotkey $script:State.PendingCloseHotkey
    }

    $script:State.InvokeHotkeyTextBox.Add_PreviewKeyDown($captureHandler)
    $script:State.CloseHotkeyTextBox.Add_PreviewKeyDown($captureHandler)

    $script:State.SameHotkeyCheckBox.Add_Click({
        $same = [bool]$script:State.SameHotkeyCheckBox.IsChecked
        $script:State.CloseHotkeyTextBox.IsEnabled = -not $same
        $script:State.CloseDoubleTapCheckBox.IsEnabled = -not $same
        if ($same) {
            $script:State.PendingCloseHotkey = Copy-Hotkey $script:State.PendingInvokeHotkey
            $script:State.CloseHotkeyTextBox.Text = Format-Hotkey $script:State.PendingCloseHotkey
            $script:State.CloseDoubleTapCheckBox.IsChecked = $script:State.InvokeDoubleTapCheckBox.IsChecked
        }
    })

    $script:State.InvokeDoubleTapCheckBox.Add_Click({
        $script:State.PendingInvokeHotkey.TapCount = if ($script:State.InvokeDoubleTapCheckBox.IsChecked) { 2 } else { 1 }
        $script:State.InvokeHotkeyTextBox.Text = Format-Hotkey $script:State.PendingInvokeHotkey
        if ($script:State.SameHotkeyCheckBox.IsChecked) {
            $script:State.PendingCloseHotkey = Copy-Hotkey $script:State.PendingInvokeHotkey
            $script:State.CloseHotkeyTextBox.Text = Format-Hotkey $script:State.PendingCloseHotkey
            $script:State.CloseDoubleTapCheckBox.IsChecked = $script:State.InvokeDoubleTapCheckBox.IsChecked
        }
    })

    $script:State.CloseDoubleTapCheckBox.Add_Click({
        $script:State.PendingCloseHotkey.TapCount = if ($script:State.CloseDoubleTapCheckBox.IsChecked) { 2 } else { 1 }
        $script:State.CloseHotkeyTextBox.Text = Format-Hotkey $script:State.PendingCloseHotkey
    })

    $script:State.SaveSettingsButton.Add_Click({ Save-SettingsFromDialog })
    $script:State.CancelSettingsButton.Add_Click({ $script:State.SettingsWindow.Close() })
    $window.Add_Closed({
        $script:State.SettingsWindow = $null
        $script:State.SettingsHintTextBlock = $null
        $script:State.LanguageLabelTextBlock = $null
        $script:State.LanguageComboBox = $null
        $script:State.LanguageZhItem = $null
        $script:State.LanguageEnItem = $null
        $script:State.OpenHotkeyLabel = $null
        $script:State.InvokeHotkeyTextBox = $null
        $script:State.CloseHotkeyTextBox = $null
        $script:State.SameHotkeyCheckBox = $null
        $script:State.CloseHotkeyLabel = $null
        $script:State.StartupCheckBox = $null
        $script:State.InvokeDoubleTapCheckBox = $null
        $script:State.CloseDoubleTapCheckBox = $null
        $script:State.SaveSettingsButton = $null
        $script:State.CancelSettingsButton = $null
    })
}

function Initialize-NotifyIcon {
    $notifyIcon = New-Object System.Windows.Forms.NotifyIcon
    $notifyIcon.Icon = [System.Drawing.SystemIcons]::Application
    $notifyIcon.Text = (Get-Text 'app_name')
    $notifyIcon.Visible = $true

    $menu = New-Object System.Windows.Forms.ContextMenuStrip
    $showItem = $menu.Items.Add((Get-Text 'tray_show_editor'))
    $settingsItem = $menu.Items.Add((Get-Text 'tray_settings'))
    $exitItem = $menu.Items.Add((Get-Text 'tray_exit'))

    $showItem.Add_Click({ Show-EditorWindow })
    $settingsItem.Add_Click({ Open-SettingsWindow })
    $exitItem.Add_Click({ Exit-Application })
    $notifyIcon.ContextMenuStrip = $menu
    $notifyIcon.Add_DoubleClick({ Show-EditorWindow })

    $script:State.NotifyIcon = $notifyIcon
    $script:State.TrayShowItem = $showItem
    $script:State.TraySettingsItem = $settingsItem
    $script:State.TrayExitItem = $exitItem
}

function Initialize-HotkeyPolling {
    $script:State.HotkeyTimer = New-Object System.Windows.Threading.DispatcherTimer
    $script:State.HotkeyTimer.Interval = [TimeSpan]::FromMilliseconds($script:HotkeyPollIntervalMs)
    $script:State.HotkeyTimer.Add_Tick({ Sync-HotkeyState })
    Capture-PressedKeySnapshot
    $script:State.HotkeyTimer.Start()
}

function Initialize-Mutex {
    $createdNew = $false
    $script:State.Mutex = New-Object System.Threading.Mutex($true, 'Global\StickerQuickEdit', [ref]$createdNew)
    if (-not $createdNew) {
        Show-AppMessage (Get-Text 'msg_already_running')
        return $false
    }

    return $true
}

function Start-Sticker {
    if (-not (Initialize-Mutex)) {
        return
    }

    $script:State.Settings = Load-Settings

    $script:State.Application = New-Object System.Windows.Application
    $script:State.Application.ShutdownMode = [System.Windows.ShutdownMode]::OnExplicitShutdown

    New-MainWindow
    Initialize-NotifyIcon

    $initialResult = Apply-Hotkeys $script:State.Settings
    if (-not $initialResult.Success) {
        $fallback = New-DefaultSettings
        $script:State.Settings.InvokeHotkey = Copy-Hotkey $fallback.InvokeHotkey
        $script:State.Settings.CloseHotkey = $fallback.CloseHotkey
        $initialResult = Apply-Hotkeys $script:State.Settings
        if (-not $initialResult.Success) {
            Show-AppMessage $initialResult.Message
            Exit-Application
            return
        }

        Save-Settings $script:State.Settings
        Show-AppMessage (Get-Text 'msg_reset_default_hotkey' @((Format-Hotkey $fallback.InvokeHotkey)))
    }

    Initialize-HotkeyPolling
    Refresh-LocalizedUI
    $script:State.NotifyIcon.BalloonTipTitle = (Get-Text 'app_name')
    $script:State.NotifyIcon.BalloonTipText = (Get-Text 'notify_balloon_text' @((Format-Hotkey $script:State.Settings.InvokeHotkey)))
    $script:State.NotifyIcon.ShowBalloonTip(2000)
    Set-LocalizedStatus 'status_app_started'
    [void]$script:State.Application.Run()
}

function Invoke-SmokeTest {
    $settings = Load-Settings
    $script:State.Settings = $settings
    New-MainWindow
    New-SettingsWindow
    if (-not (Test-HotkeyValid $settings.InvokeHotkey)) {
        throw 'The open hotkey configuration is invalid.'
    }
    if (-not (Test-HotkeyValid (Get-EffectiveCloseHotkey $settings))) {
        throw 'The close hotkey configuration is invalid.'
    }
    'Smoke test passed.'
}

if ($SmokeTest) {
    Invoke-SmokeTest
    return
}

Start-Sticker
