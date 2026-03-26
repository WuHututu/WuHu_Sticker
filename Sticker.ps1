##[Ps1 To Exe]
##
##Kd3HDZOFADWE8uK1
##Nc3NCtDXThU=
##Kd3HFJGZHWLWoLaVvnQnhQ==
##LM/RF4eFHHGZ7/K1
##K8rLFtDXTiW5
##OsHQCZGeTiiZ4dI=
##OcrLFtDXTiW5
##LM/BD5WYTiiZ4tI=
##McvWDJ+OTiiZ4tI=
##OMvOC56PFnzN8u+Vs1Q=
##M9jHFoeYB2Hc8u+Vs1Q=
##PdrWFpmIG2HcofKIo2QX
##OMfRFJyLFzWE8uK1
##KsfMAp/KUzWJ0g==
##OsfOAYaPHGbQvbyVvnQX
##LNzNAIWJGmPcoKHc7Do3uAuO
##LNzNAIWJGnvYv7eVvnQX
##M9zLA5mED3nfu77Q7TV64AuzAgg=
##NcDWAYKED3nfu77Q7TV64AuzAgg=
##OMvRB4KDHmHQvbyVvnQX
##P8HPFJGEFzWE8tI=
##KNzDAJWHD2fS8u+Vgw==
##P8HSHYKDCX3N8u+Vgw==
##LNzLEpGeC3fMu77Ro2k3hQ==
##L97HB5mLAnfMu77Ro2k3hQ==
##P8HPCZWEGmaZ7/K1
##L8/UAdDXTlaDjofG5iZk2RK+Gzp/UuGeqr2zy5GA7P/usSDaXYkoSltzkzHAF1+0WvkXR8kGoNgSXhg4YfcT59I=
##Kc/BRM3KXhU=
##
##
##fd6a9f26a06ea3bc99616d4851b372ba
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
$script:ButtonIconDirectory = Join-Path $script:AppRoot 'assets\button-icons'
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
    "main_pin_on": "Pin",
    "main_pin_off": "Unpin",
    "main_settings": "Settings",
    "main_fit_width": "Fit Longest Line",
    "main_fit_height": "Fit All Lines",
    "status_ready": "Ready",
    "status_opened": "Opened. Show: {0}; Hide: {1}",
    "status_hidden": "Editor hidden.",
    "status_nothing_to_copy": "Nothing to copy.",
    "status_all_copied": "All content copied to clipboard.",
    "status_select_first": "Select some content first.",
    "status_selected_copied": "Selected content copied to clipboard.",
    "status_settings_saved": "Settings saved.",
    "status_editor_cleared": "Editor cleared.",
    "status_topmost_enabled": "Always on top enabled.",
    "status_topmost_disabled": "Always on top disabled.",
    "status_image_pasted": "Image pasted.",
    "status_images_pasted": "{0} images pasted.",
    "status_app_started": "App started. Waiting for hotkey.",
    "settings_title": "Sticker Settings",
    "settings_hint": "Click a box and press the shortcut. Win is not supported.",
    "settings_language": "Language",
    "language_zh": "中文",
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
    "main_pin_on": "\u7f6e\u9876",
    "main_pin_off": "\u53d6\u6d88\u7f6e\u9876",
    "main_settings": "\u8bbe\u7f6e",
    "main_fit_width": "\u9002\u914d\u6700\u957f\u884c",
    "main_fit_height": "\u9002\u914d\u5168\u90e8\u884c",
    "status_ready": "\u5c31\u7eea",
    "status_opened": "\u5df2\u6253\u5f00\u3002\u5524\u51fa\uff1a{0}\uff1b\u5173\u95ed\uff1a{1}",
    "status_hidden": "\u7f16\u8f91\u7a97\u5df2\u9690\u85cf\u3002",
    "status_nothing_to_copy": "\u6ca1\u6709\u53ef\u590d\u5236\u7684\u5185\u5bb9\u3002",
    "status_all_copied": "\u5168\u90e8\u5185\u5bb9\u5df2\u590d\u5236\u5230\u526a\u8d34\u677f\u3002",
    "status_select_first": "\u8bf7\u5148\u9009\u4e2d\u8981\u590d\u5236\u7684\u5185\u5bb9\u3002",
    "status_selected_copied": "\u9009\u4e2d\u5185\u5bb9\u5df2\u590d\u5236\u5230\u526a\u8d34\u677f\u3002",
    "status_settings_saved": "\u8bbe\u7f6e\u5df2\u4fdd\u5b58\u3002",
    "status_editor_cleared": "\u5185\u5bb9\u5df2\u6e05\u7a7a\u3002",
    "status_topmost_enabled": "\u5df2\u5f00\u542f\u7f6e\u9876\u3002",
    "status_topmost_disabled": "\u5df2\u5173\u95ed\u7f6e\u9876\u3002",
    "status_image_pasted": "\u56fe\u7247\u5df2\u7c98\u8d34\u3002",
    "status_images_pasted": "\u5df2\u7c98\u8d34 {0} \u5f20\u56fe\u7247\u3002",
    "status_app_started": "\u7a0b\u5e8f\u5df2\u542f\u52a8\uff0c\u7b49\u5f85\u70ed\u952e\u5524\u51fa\u3002",
    "settings_title": "\u4fbf\u7b7e\u7f16\u8f91\u8bbe\u7f6e",
    "settings_hint": "\u70b9\u51fb\u8f93\u5165\u6846\u540e\u76f4\u63a5\u6309\u4e0b\u5feb\u6377\u952e\u3002\u4e0d\u652f\u6301 Win \u952e\u3002",
    "settings_language": "\u8bed\u8a00",
    "language_zh": "\u4e2d\u6587",
    "language_en": "English",
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
    WindowDragHandle = $null
    EditorRichTextBox = $null
    StatusTextBlock = $null
    CopyAllButton = $null
    CopySelectionButton = $null
    ClearButton = $null
    TopmostButton = $null
    SettingsButton = $null
    FitWidthButton = $null
    FitHeightButton = $null
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
    StartupToastWindow = $null
    StartupToastTimer = $null
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
        [Alias('Args')]
        [object[]]$FormatArgs = @(),
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

    if ($FormatArgs.Count -gt 0) {
        return [string]::Format($text, $FormatArgs)
    }

    return [string]$text
}

function Set-LocalizedStatus {
    param(
        [string]$Key,
        [object[]]$FormatArgs = @()
    )

    $script:State.LastStatusKey = $Key
    $script:State.LastStatusArgs = @($FormatArgs)
    Set-Status (Get-Text -Key $Key -FormatArgs $FormatArgs)
}

function Refresh-StatusText {
    if ([string]::IsNullOrWhiteSpace([string]$script:State.LastStatusKey)) {
        return
    }

    Set-Status (Get-Text -Key $script:State.LastStatusKey -FormatArgs $script:State.LastStatusArgs)
}

function Show-AppMessage {
    param(
        [string]$Message,
        [string]$Language = ''
    )

    [System.Windows.MessageBox]::Show($Message, (Get-Text -Key 'app_name' -Language $Language)) | Out-Null
}

function Close-StartupToast {
    if ($null -ne $script:State.StartupToastTimer) {
        $script:State.StartupToastTimer.Stop()
        $script:State.StartupToastTimer = $null
    }

    if ($null -ne $script:State.StartupToastWindow) {
        try {
            $script:State.StartupToastWindow.Close()
        }
        catch {
        }
        $script:State.StartupToastWindow = $null
    }
}

function Show-StartupToast {
    param(
        [string]$Message
    )

    Close-StartupToast

    $window = New-Object System.Windows.Window
    $window.Width = 320
    $window.Height = 94
    $window.WindowStyle = [System.Windows.WindowStyle]::None
    $window.ResizeMode = [System.Windows.ResizeMode]::NoResize
    $window.AllowsTransparency = $true
    $window.Background = [System.Windows.Media.Brushes]::Transparent
    $window.ShowInTaskbar = $false
    $window.Topmost = $true
    $window.ShowActivated = $false

    $border = New-Object System.Windows.Controls.Border
    $border.CornerRadius = New-Object System.Windows.CornerRadius(14)
    $border.Background = New-Brush '#F8F3E6'
    $border.BorderBrush = New-Brush '#D9C8A9'
    $border.BorderThickness = New-Object System.Windows.Thickness(1)
    $border.Padding = New-Object System.Windows.Thickness(14)

    $stack = New-Object System.Windows.Controls.StackPanel
    $title = New-Object System.Windows.Controls.TextBlock
    $title.Text = Get-Text 'app_name'
    $title.FontSize = 14
    $title.FontWeight = [System.Windows.FontWeights]::SemiBold
    $title.Foreground = New-Brush '#3D3327'
    $title.Margin = New-Object System.Windows.Thickness(0, 0, 0, 6)
    [void]$stack.Children.Add($title)

    $body = New-Object System.Windows.Controls.TextBlock
    $body.Text = $Message
    $body.TextWrapping = [System.Windows.TextWrapping]::Wrap
    $body.FontSize = 12
    $body.Foreground = New-Brush '#5C5345'
    [void]$stack.Children.Add($body)

    $border.Child = $stack
    $window.Content = $border

    $workArea = [System.Windows.SystemParameters]::WorkArea
    $window.Left = $workArea.Right - $window.Width - 18
    $window.Top = $workArea.Bottom - $window.Height - 18

    $script:State.StartupToastWindow = $window
    $window.Add_Closed({
        param($sender, $eventArgs)

        if ($script:State.StartupToastWindow -eq $sender) {
            $script:State.StartupToastWindow = $null
        }
    })
    $window.Show()

    $timer = New-Object System.Windows.Threading.DispatcherTimer
    $timer.Interval = [TimeSpan]::FromSeconds(4)
    $timer.Add_Tick({
        param($sender, $eventArgs)

        $sender.Stop()
        if ($script:State.StartupToastTimer -eq $sender) {
            $script:State.StartupToastTimer = $null
        }
        if ($null -ne $script:State.StartupToastWindow) {
            $script:State.StartupToastWindow.Close()
        }
    })
    $script:State.StartupToastTimer = $timer
    $timer.Start()
}

function New-Brush {
    param(
        [string]$Color
    )

    return [System.Windows.Media.BrushConverter]::new().ConvertFromString($Color)
}

function Get-ButtonIconAssetPath {
    param(
        [string]$Key,
        [bool]$Active = $false
    )

    $searchRoots = @($script:ButtonIconDirectory, $script:AppRoot)
    $names = switch ($Key) {
        'copy_all' { @('copy-all', 'copy_all', 'copyall', 'copy') }
        'copy_selected' { @('copy-selected', 'copy_selected', 'copyselection', 'copy-selection', 'copy_selected_text') }
        'clear' { @('clear', 'erase', 'trash') }
        'fit_width' { @('fit-width', 'fit_width', 'expand-width', 'expand_width', 'width') }
        'fit_height' { @('fit-height', 'fit_height', 'expand-height', 'expand_height', 'height') }
        'pin' {
            if ($Active) {
                @('pin-filled', 'pin_filled', 'pin-on', 'pin_on', 'pinned', 'pin')
            }
            else {
                @('pin', 'pin-off', 'pin_off', 'unpinned')
            }
        }
        'settings' { @('settings', 'setting', 'gear', 'config') }
        default { @($Key) }
    }

    $extensions = @('.ico', '.png', '.bmp', '.gif', '.jpg', '.jpeg')
    foreach ($root in $searchRoots) {
        if (-not (Test-Path $root)) {
            continue
        }

        foreach ($name in $names) {
            foreach ($extension in $extensions) {
                $candidate = Join-Path $root ($name + $extension)
                if (Test-Path $candidate -PathType Leaf) {
                    return $candidate
                }
            }
        }
    }

    return $null
}

function New-IconImageContent {
    param(
        [string]$Path
    )

    $image = New-Object System.Windows.Controls.Image
    $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
    $bitmap.BeginInit()
    $bitmap.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
    $bitmap.UriSource = New-Object System.Uri($Path)
    $bitmap.EndInit()
    $bitmap.Freeze()
    $image.Source = $bitmap
    $image.Width = 18
    $image.Height = 18
    $image.Stretch = [System.Windows.Media.Stretch]::Uniform
    $image.SnapsToDevicePixels = $true
    return $image
}

function Add-CanvasShape {
    param(
        [System.Windows.Controls.Canvas]$Canvas,
        [System.Windows.FrameworkElement]$Element,
        [double]$Left = 0,
        [double]$Top = 0
    )

    [System.Windows.Controls.Canvas]::SetLeft($Element, $Left)
    [System.Windows.Controls.Canvas]::SetTop($Element, $Top)
    [void]$Canvas.Children.Add($Element)
}

function New-VectorToolbarIcon {
    param(
        [string]$Key,
        [bool]$Active = $false
    )

    $stroke = New-Brush '#4A4033'
    $softStroke = New-Brush '#6A6154'
    $accentFill = if ($Active) { New-Brush '#D06B38' } else { [System.Windows.Media.Brushes]::Transparent }
    $lightFill = New-Brush '#EFE3D0'
    $canvas = New-Object System.Windows.Controls.Canvas
    $canvas.Width = 18
    $canvas.Height = 18

    switch ($Key) {
        'copy_all' {
            $back = New-Object System.Windows.Shapes.Rectangle
            $back.Width = 8
            $back.Height = 9
            $back.RadiusX = 1.2
            $back.RadiusY = 1.2
            $back.Stroke = $softStroke
            $back.StrokeThickness = 1.3
            Add-CanvasShape $canvas $back 2 3

            $front = New-Object System.Windows.Shapes.Rectangle
            $front.Width = 8
            $front.Height = 9
            $front.RadiusX = 1.2
            $front.RadiusY = 1.2
            $front.Stroke = $stroke
            $front.StrokeThickness = 1.5
            Add-CanvasShape $canvas $front 7 6
        }
        'copy_selected' {
            $back = New-Object System.Windows.Shapes.Rectangle
            $back.Width = 8
            $back.Height = 9
            $back.RadiusX = 1.2
            $back.RadiusY = 1.2
            $back.Stroke = $softStroke
            $back.StrokeThickness = 1.3
            Add-CanvasShape $canvas $back 2 3

            $front = New-Object System.Windows.Shapes.Rectangle
            $front.Width = 8
            $front.Height = 9
            $front.RadiusX = 1.2
            $front.RadiusY = 1.2
            $front.Stroke = $stroke
            $front.StrokeThickness = 1.5
            $front.Fill = $lightFill
            Add-CanvasShape $canvas $front 7 6

            $selection = New-Object System.Windows.Shapes.Rectangle
            $selection.Width = 4
            $selection.Height = 2.8
            $selection.Fill = New-Brush '#C45A2E'
            Add-CanvasShape $canvas $selection 9 9.2
        }
        'clear' {
            $lid = New-Object System.Windows.Shapes.Line
            $lid.X1 = 4.5
            $lid.Y1 = 5
            $lid.X2 = 13.5
            $lid.Y2 = 5
            $lid.Stroke = $stroke
            $lid.StrokeThickness = 1.4
            [void]$canvas.Children.Add($lid)

            $handle = New-Object System.Windows.Shapes.Line
            $handle.X1 = 7
            $handle.Y1 = 3.2
            $handle.X2 = 11
            $handle.Y2 = 3.2
            $handle.Stroke = $stroke
            $handle.StrokeThickness = 1.3
            $handle.StrokeStartLineCap = [System.Windows.Media.PenLineCap]::Round
            $handle.StrokeEndLineCap = [System.Windows.Media.PenLineCap]::Round
            [void]$canvas.Children.Add($handle)

            $body = New-Object System.Windows.Shapes.Rectangle
            $body.Width = 7.4
            $body.Height = 8.8
            $body.RadiusX = 1.1
            $body.RadiusY = 1.1
            $body.Stroke = $stroke
            $body.StrokeThickness = 1.4
            Add-CanvasShape $canvas $body 5.3 5.4

            foreach ($x in @(7.9, 9.0, 10.1)) {
                $line = New-Object System.Windows.Shapes.Line
                $line.X1 = $x
                $line.Y1 = 7.1
                $line.X2 = $x
                $line.Y2 = 12.5
                $line.Stroke = $stroke
                $line.StrokeThickness = 1.1
                $line.StrokeStartLineCap = [System.Windows.Media.PenLineCap]::Round
                $line.StrokeEndLineCap = [System.Windows.Media.PenLineCap]::Round
                [void]$canvas.Children.Add($line)
            }
        }
        'hide' {
            $line = New-Object System.Windows.Shapes.Line
            $line.X1 = 4
            $line.Y1 = 13
            $line.X2 = 14
            $line.Y2 = 13
            $line.Stroke = $stroke
            $line.StrokeThickness = 1.8
            $line.StrokeStartLineCap = [System.Windows.Media.PenLineCap]::Round
            $line.StrokeEndLineCap = [System.Windows.Media.PenLineCap]::Round
            [void]$canvas.Children.Add($line)
        }
        'fit_width' {
            $left = New-Object System.Windows.Shapes.Line
            $left.X1 = 2.8
            $left.Y1 = 9
            $left.X2 = 15.2
            $left.Y2 = 9
            $left.Stroke = $softStroke
            $left.StrokeThickness = 1.2
            [void]$canvas.Children.Add($left)

            foreach ($coords in @(@(4.2, 9, 7.2, 6), @(4.2, 9, 7.2, 12), @(13.8, 9, 10.8, 6), @(13.8, 9, 10.8, 12))) {
                $line = New-Object System.Windows.Shapes.Line
                $line.X1 = $coords[0]
                $line.Y1 = $coords[1]
                $line.X2 = $coords[2]
                $line.Y2 = $coords[3]
                $line.Stroke = $stroke
                $line.StrokeThickness = 1.5
                $line.StrokeStartLineCap = [System.Windows.Media.PenLineCap]::Round
                $line.StrokeEndLineCap = [System.Windows.Media.PenLineCap]::Round
                [void]$canvas.Children.Add($line)
            }
        }
        'fit_height' {
            $center = New-Object System.Windows.Shapes.Line
            $center.X1 = 9
            $center.Y1 = 2.8
            $center.X2 = 9
            $center.Y2 = 15.2
            $center.Stroke = $softStroke
            $center.StrokeThickness = 1.2
            [void]$canvas.Children.Add($center)

            foreach ($coords in @(@(9, 4.2, 6, 7.2), @(9, 4.2, 12, 7.2), @(9, 13.8, 6, 10.8), @(9, 13.8, 12, 10.8))) {
                $line = New-Object System.Windows.Shapes.Line
                $line.X1 = $coords[0]
                $line.Y1 = $coords[1]
                $line.X2 = $coords[2]
                $line.Y2 = $coords[3]
                $line.Stroke = $stroke
                $line.StrokeThickness = 1.5
                $line.StrokeStartLineCap = [System.Windows.Media.PenLineCap]::Round
                $line.StrokeEndLineCap = [System.Windows.Media.PenLineCap]::Round
                [void]$canvas.Children.Add($line)
            }
        }
        'pin' {
            $head = New-Object System.Windows.Shapes.Polygon
            $head.Points = [System.Windows.Media.PointCollection]::Parse('9,2 14,7 11.4,8.3 11.4,12.2 9,16.2 6.6,12.2 6.6,8.3 4,7')
            $head.Stroke = $stroke
            $head.StrokeThickness = 1.3
            $head.Fill = $accentFill
            [void]$canvas.Children.Add($head)

            $stem = New-Object System.Windows.Shapes.Line
            $stem.X1 = 9
            $stem.Y1 = 8.3
            $stem.X2 = 9
            $stem.Y2 = 17
            $stem.Stroke = $stroke
            $stem.StrokeThickness = 1.4
            $stem.StrokeStartLineCap = [System.Windows.Media.PenLineCap]::Round
            $stem.StrokeEndLineCap = [System.Windows.Media.PenLineCap]::Round
            [void]$canvas.Children.Add($stem)
        }
        'settings' {
            $track1 = New-Object System.Windows.Shapes.Line
            $track1.X1 = 3
            $track1.Y1 = 5
            $track1.X2 = 15
            $track1.Y2 = 5
            $track1.Stroke = $stroke
            $track1.StrokeThickness = 1.3
            [void]$canvas.Children.Add($track1)

            $track2 = New-Object System.Windows.Shapes.Line
            $track2.X1 = 3
            $track2.Y1 = 9
            $track2.X2 = 15
            $track2.Y2 = 9
            $track2.Stroke = $stroke
            $track2.StrokeThickness = 1.3
            [void]$canvas.Children.Add($track2)

            $track3 = New-Object System.Windows.Shapes.Line
            $track3.X1 = 3
            $track3.Y1 = 13
            $track3.X2 = 15
            $track3.Y2 = 13
            $track3.Stroke = $stroke
            $track3.StrokeThickness = 1.3
            [void]$canvas.Children.Add($track3)

            foreach ($knob in @(@(6, 5), @(12, 9), @(8, 13))) {
                $ellipse = New-Object System.Windows.Shapes.Ellipse
                $ellipse.Width = 4
                $ellipse.Height = 4
                $ellipse.Fill = $lightFill
                $ellipse.Stroke = $stroke
                $ellipse.StrokeThickness = 1.2
                Add-CanvasShape $canvas $ellipse ($knob[0] - 2) ($knob[1] - 2)
            }
        }
    }

    $viewbox = New-Object System.Windows.Controls.Viewbox
    $viewbox.Width = 18
    $viewbox.Height = 18
    $viewbox.Stretch = [System.Windows.Media.Stretch]::Uniform
    $viewbox.Child = $canvas
    return $viewbox
}

function New-ToolbarButtonContent {
    param(
        [string]$Key,
        [bool]$Active = $false
    )

    $assetPath = Get-ButtonIconAssetPath -Key $Key -Active $Active
    if ($null -ne $assetPath) {
        try {
            return New-IconImageContent -Path $assetPath
        }
        catch {
        }
    }

    return New-VectorToolbarIcon -Key $Key -Active $Active
}

function Set-ToolbarButtonPresentation {
    param(
        [System.Windows.Controls.Button]$Button,
        [string]$IconKey,
        [string]$TooltipKey,
        [bool]$Active = $false
    )

    if ($null -eq $Button) {
        return
    }

    $Button.Content = New-ToolbarButtonContent -Key $IconKey -Active $Active
    $Button.ToolTip = Get-Text $TooltipKey
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
        AlwaysOnTop = $false
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
        AlwaysOnTop = [bool]$Settings.AlwaysOnTop
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
        AlwaysOnTop = [bool]$Settings.AlwaysOnTop
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

        if ($null -ne $raw.AlwaysOnTop) {
            $settings.AlwaysOnTop = [bool]$raw.AlwaysOnTop
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
        Set-ToolbarButtonPresentation -Button $script:State.CopyAllButton -IconKey 'copy_all' -TooltipKey 'main_copy_all'
    }
    if ($null -ne $script:State.CopySelectionButton) {
        Set-ToolbarButtonPresentation -Button $script:State.CopySelectionButton -IconKey 'copy_selected' -TooltipKey 'main_copy_selected'
    }
    if ($null -ne $script:State.ClearButton) {
        Set-ToolbarButtonPresentation -Button $script:State.ClearButton -IconKey 'clear' -TooltipKey 'main_clear'
    }
    if ($null -ne $script:State.TopmostButton) {
        if ([bool]$script:State.Settings.AlwaysOnTop) {
            Set-ToolbarButtonPresentation -Button $script:State.TopmostButton -IconKey 'pin' -TooltipKey 'main_pin_off' -Active $true
        }
        else {
            Set-ToolbarButtonPresentation -Button $script:State.TopmostButton -IconKey 'pin' -TooltipKey 'main_pin_on'
        }
    }
    if ($null -ne $script:State.SettingsButton) {
        Set-ToolbarButtonPresentation -Button $script:State.SettingsButton -IconKey 'settings' -TooltipKey 'main_settings'
    }
    if ($null -ne $script:State.FitWidthButton) {
        Set-ToolbarButtonPresentation -Button $script:State.FitWidthButton -IconKey 'fit_width' -TooltipKey 'main_fit_width'
    }
    if ($null -ne $script:State.FitHeightButton) {
        Set-ToolbarButtonPresentation -Button $script:State.FitHeightButton -IconKey 'fit_height' -TooltipKey 'main_fit_height'
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

function New-EditorDocument {
    $document = New-Object System.Windows.Documents.FlowDocument
    $document.PagePadding = New-Object System.Windows.Thickness(0)
    $document.LineStackingStrategy = [System.Windows.LineStackingStrategy]::BlockLineHeight
    $document.LineHeight = 18
    $paragraphStyle = New-Object System.Windows.Style([System.Windows.Documents.Paragraph])
    [void]$paragraphStyle.Setters.Add((New-Object System.Windows.Setter([System.Windows.Documents.Block]::MarginProperty, (New-Object System.Windows.Thickness(0)))))
    $document.Resources.Add([System.Windows.Documents.Paragraph], $paragraphStyle)
    [void]$document.Blocks.Add((New-Object System.Windows.Documents.Paragraph))
    return $document
}

function Ensure-EditorDocumentHasParagraph {
    if ($null -eq $script:State.EditorRichTextBox) {
        return
    }

    if ($script:State.EditorRichTextBox.Document.Blocks.Count -eq 0) {
        [void]$script:State.EditorRichTextBox.Document.Blocks.Add((New-Object System.Windows.Documents.Paragraph))
    }
}

function Get-EditorDocumentRange {
    if ($null -eq $script:State.EditorRichTextBox) {
        return $null
    }

    return New-Object System.Windows.Documents.TextRange(
        $script:State.EditorRichTextBox.Document.ContentStart,
        $script:State.EditorRichTextBox.Document.ContentEnd
    )
}

function Get-EditorPlainText {
    $range = Get-EditorDocumentRange
    if ($null -eq $range) {
        return ''
    }

    $text = [string]$range.Text
    if ($text.EndsWith("`r`n")) {
        return $text.Substring(0, $text.Length - 2)
    }
    if ($text.EndsWith("`n")) {
        return $text.Substring(0, $text.Length - 1)
    }

    return $text
}

function Get-EditorTextLines {
    $text = Get-EditorPlainText
    if ($text.Length -eq 0) {
        return @('')
    }

    $lines = [System.Text.RegularExpressions.Regex]::Split($text, "\r?\n")
    if ($lines.Count -eq 0) {
        return @('')
    }

    return $lines
}

function New-EditorTypeface {
    if ($null -eq $script:State.EditorRichTextBox) {
        return $null
    }

    return [System.Windows.Media.Typeface]::new(
        $script:State.EditorRichTextBox.FontFamily,
        $script:State.EditorRichTextBox.FontStyle,
        $script:State.EditorRichTextBox.FontWeight,
        $script:State.EditorRichTextBox.FontStretch
    )
}

function Measure-EditorTextWidth {
    param(
        [string]$Text
    )

    if ($null -eq $script:State.EditorRichTextBox) {
        return 0
    }

    $displayText = if ([string]::IsNullOrEmpty($Text)) { ' ' } else { $Text.Replace("`t", '    ') }
    $formattedText = [System.Windows.Media.FormattedText]::new(
        $displayText,
        [System.Globalization.CultureInfo]::CurrentUICulture,
        [System.Windows.FlowDirection]::LeftToRight,
        (New-EditorTypeface),
        [double]$script:State.EditorRichTextBox.FontSize,
        [System.Windows.Media.Brushes]::Black,
        1.0
    )

    return [Math]::Ceiling($formattedText.WidthIncludingTrailingWhitespace)
}

function Get-EditorLineHeight {
    if ($null -eq $script:State.EditorRichTextBox) {
        return 18
    }

    if ($null -ne $script:State.EditorRichTextBox.Document -and $script:State.EditorRichTextBox.Document.LineHeight -gt 0) {
        return [Math]::Ceiling($script:State.EditorRichTextBox.Document.LineHeight)
    }

    $formattedText = [System.Windows.Media.FormattedText]::new(
        'Ag',
        [System.Globalization.CultureInfo]::CurrentUICulture,
        [System.Windows.FlowDirection]::LeftToRight,
        (New-EditorTypeface),
        [double]$script:State.EditorRichTextBox.FontSize,
        [System.Windows.Media.Brushes]::Black,
        1.0
    )

    return [Math]::Ceiling([Math]::Max($formattedText.Height, ($script:State.EditorRichTextBox.FontSize * 1.15)))
}

function Get-MainWindowChromeSize {
    $fallback = @{
        Width = 72
        Height = 96
    }

    if ($null -eq $script:State.MainWindow -or $null -eq $script:State.EditorRichTextBox) {
        return $fallback
    }

    $script:State.MainWindow.UpdateLayout()
    $script:State.EditorRichTextBox.UpdateLayout()

    if ($script:State.MainWindow.ActualWidth -le 0 -or $script:State.EditorRichTextBox.ActualWidth -le 0) {
        return $fallback
    }

    $chromeWidth = [Math]::Ceiling($script:State.MainWindow.ActualWidth - $script:State.EditorRichTextBox.ActualWidth)
    $chromeHeight = [Math]::Ceiling($script:State.MainWindow.ActualHeight - $script:State.EditorRichTextBox.ActualHeight)

    return @{
        Width = if ($chromeWidth -gt 0) { $chromeWidth } else { $fallback.Width }
        Height = if ($chromeHeight -gt 0) { $chromeHeight } else { $fallback.Height }
    }
}

function Set-MainWindowToWorkArea {
    if ($null -eq $script:State.MainWindow) {
        return
    }

    $workArea = [System.Windows.SystemParameters]::WorkArea
    $script:State.MainWindow.WindowState = [System.Windows.WindowState]::Normal
    $script:State.MainWindow.Left = $workArea.Left
    $script:State.MainWindow.Top = $workArea.Top
    $script:State.MainWindow.Width = $workArea.Width
    $script:State.MainWindow.Height = $workArea.Height
}

function Resize-MainWindowToContent {
    param(
        [switch]$FitWidth,
        [switch]$FitHeight
    )

    if ($null -eq $script:State.MainWindow -or $null -eq $script:State.EditorRichTextBox) {
        return
    }

    try {
        Ensure-EditorDocumentHasParagraph
        $script:State.MainWindow.UpdateLayout()
        $script:State.EditorRichTextBox.UpdateLayout()

        $chrome = Get-MainWindowChromeSize
        $workArea = [System.Windows.SystemParameters]::WorkArea
        $lines = @(Get-EditorTextLines)
        if ($lines.Count -eq 0) {
            $lines = @('')
        }

        $targetWidth = [double]$script:State.MainWindow.Width
        $targetHeight = [double]$script:State.MainWindow.Height

        if ($FitWidth) {
            $longestWidth = 0
            foreach ($line in $lines) {
                $lineWidth = Measure-EditorTextWidth $line
                if ($lineWidth -gt $longestWidth) {
                    $longestWidth = $lineWidth
                }
            }

            $extraCharacterWidth = [Math]::Max((Measure-EditorTextWidth 'W'), 10)
            $targetWidth = $chrome.Width + $longestWidth + $extraCharacterWidth + 10
        }

        if ($FitHeight) {
            $lineHeight = Get-EditorLineHeight
            $targetHeight = $chrome.Height + (($lines.Count + 1) * $lineHeight) + 2
        }

        $targetWidth = [Math]::Max($targetWidth, $script:State.MainWindow.MinWidth)
        $targetHeight = [Math]::Max($targetHeight, $script:State.MainWindow.MinHeight)

        if ($targetWidth -ge $workArea.Width - 1 -or $targetHeight -ge $workArea.Height - 1) {
            Set-MainWindowToWorkArea
            Focus-EditorWindow
            return
        }

        $restoreLeft = if ($script:State.MainWindow.WindowState -eq [System.Windows.WindowState]::Normal) {
            $script:State.MainWindow.Left
        }
        else {
            $script:State.MainWindow.RestoreBounds.Left
        }
        $restoreTop = if ($script:State.MainWindow.WindowState -eq [System.Windows.WindowState]::Normal) {
            $script:State.MainWindow.Top
        }
        else {
            $script:State.MainWindow.RestoreBounds.Top
        }

        if ([double]::IsNaN($restoreLeft) -or [double]::IsInfinity($restoreLeft)) {
            $restoreLeft = $workArea.Left
        }
        if ([double]::IsNaN($restoreTop) -or [double]::IsInfinity($restoreTop)) {
            $restoreTop = $workArea.Top
        }

        $script:State.MainWindow.WindowState = [System.Windows.WindowState]::Normal
        $script:State.MainWindow.Width = [Math]::Round($targetWidth)
        $script:State.MainWindow.Height = [Math]::Round($targetHeight)
        $script:State.MainWindow.Left = [Math]::Max($workArea.Left, [Math]::Min($restoreLeft, ($workArea.Right - $script:State.MainWindow.Width)))
        $script:State.MainWindow.Top = [Math]::Max($workArea.Top, [Math]::Min($restoreTop, ($workArea.Bottom - $script:State.MainWindow.Height)))
        Focus-EditorWindow
    }
    catch {
        Focus-EditorWindow
    }
}

function Convert-TextRangeToXaml {
    param(
        [System.Windows.Documents.TextRange]$Range
    )

    $stream = New-Object System.IO.MemoryStream
    try {
        $Range.Save($stream, [System.Windows.DataFormats]::Xaml)
        $stream.Position = 0
        $reader = New-Object System.IO.StreamReader($stream)
        try {
            return $reader.ReadToEnd()
        }
        finally {
            $reader.Dispose()
        }
    }
    finally {
        $stream.Dispose()
    }
}

function Test-TextRangeHasContent {
    param(
        [System.Windows.Documents.TextRange]$Range
    )

    if ($null -eq $Range) {
        return $false
    }

    if (-not [string]::IsNullOrWhiteSpace(($Range.Text -replace '\s', ''))) {
        return $true
    }

    try {
        $xaml = Convert-TextRangeToXaml $Range
        return $xaml -match 'InlineUIContainer|BlockUIContainer|Image'
    }
    catch {
        return $false
    }
}

function Test-InlineCollectionHasContent {
    param(
        [System.Windows.Documents.InlineCollection]$Inlines
    )

    foreach ($inline in $Inlines) {
        if ($inline -is [System.Windows.Documents.Run]) {
            if (-not [string]::IsNullOrWhiteSpace(($inline.Text -replace '\s', ''))) {
                return $true
            }
            continue
        }

        if ($inline -is [System.Windows.Documents.Span]) {
            if (Test-InlineCollectionHasContent $inline.Inlines) {
                return $true
            }
            continue
        }

        if ($inline -is [System.Windows.Documents.InlineUIContainer]) {
            return $true
        }
    }

    return $false
}

function Test-BlockCollectionHasContent {
    param(
        [System.Windows.Documents.BlockCollection]$Blocks
    )

    foreach ($block in $Blocks) {
        if ($block -is [System.Windows.Documents.Paragraph]) {
            if (Test-InlineCollectionHasContent $block.Inlines) {
                return $true
            }
            continue
        }

        if ($block -is [System.Windows.Documents.Section]) {
            if (Test-BlockCollectionHasContent $block.Blocks) {
                return $true
            }
            continue
        }

        if ($block -is [System.Windows.Documents.List]) {
            foreach ($item in $block.ListItems) {
                if (Test-BlockCollectionHasContent $item.Blocks) {
                    return $true
                }
            }
            continue
        }

        if ($block -is [System.Windows.Documents.BlockUIContainer]) {
            return $true
        }
    }

    return $false
}

function Test-EditorHasContent {
    if ($null -eq $script:State.EditorRichTextBox) {
        return $false
    }

    return Test-BlockCollectionHasContent $script:State.EditorRichTextBox.Document.Blocks
}

function Reset-EditorDocument {
    if ($null -eq $script:State.EditorRichTextBox) {
        return
    }

    $script:State.EditorRichTextBox.Document = New-EditorDocument
}

function Copy-EditorSelection {
    param(
        [switch]$SelectAll
    )

    if ($null -eq $script:State.EditorRichTextBox) {
        return $false
    }

    if ($SelectAll) {
        if (-not (Test-EditorHasContent)) {
            return $false
        }

        $originalStart = $script:State.EditorRichTextBox.Selection.Start
        $originalEnd = $script:State.EditorRichTextBox.Selection.End
        try {
            $script:State.EditorRichTextBox.Selection.Select(
                $script:State.EditorRichTextBox.Document.ContentStart,
                $script:State.EditorRichTextBox.Document.ContentEnd
            )
            $script:State.EditorRichTextBox.Copy()
        }
        finally {
            $script:State.EditorRichTextBox.Selection.Select($originalStart, $originalEnd)
        }

        return $true
    }

    if ($script:State.EditorRichTextBox.Selection.IsEmpty) {
        return $false
    }

    $script:State.EditorRichTextBox.Copy()
    return $true
}

function Get-EditorImageMaxWidth {
    if ($null -eq $script:State.EditorRichTextBox) {
        return 640
    }

    $width = [Math]::Floor($script:State.EditorRichTextBox.ActualWidth - 80)
    if ($width -lt 220) {
        return 220
    }

    return [int]$width
}

function New-BitmapSourceFromFile {
    param(
        [string]$Path
    )

    $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
    $bitmap.BeginInit()
    $bitmap.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
    $bitmap.UriSource = New-Object System.Uri($Path)
    $bitmap.EndInit()
    $bitmap.Freeze()
    return $bitmap
}

function Insert-ImageIntoEditor {
    param(
        [System.Windows.Media.Imaging.BitmapSource]$BitmapSource
    )

    if ($null -eq $script:State.EditorRichTextBox -or $null -eq $BitmapSource) {
        return
    }

    Ensure-EditorDocumentHasParagraph
    $selection = $script:State.EditorRichTextBox.Selection
    if (-not $selection.IsEmpty) {
        $selection.Text = ''
    }

    $image = New-Object System.Windows.Controls.Image
    $image.Source = $BitmapSource
    $image.Stretch = [System.Windows.Media.Stretch]::Uniform
    $image.MaxWidth = Get-EditorImageMaxWidth
    $image.Margin = New-Object System.Windows.Thickness(0, 4, 0, 4)

    if ($BitmapSource.DpiX -gt 0) {
        $desiredWidth = [Math]::Round($BitmapSource.PixelWidth * (96.0 / $BitmapSource.DpiX))
        $image.Width = [Math]::Min($desiredWidth, $image.MaxWidth)
    }

    $insertionPoint = $selection.Start
    $container = New-Object System.Windows.Documents.InlineUIContainer($image, $insertionPoint)
    $script:State.EditorRichTextBox.CaretPosition = $container.ElementEnd
    $script:State.EditorRichTextBox.Focus() | Out-Null
}

function Insert-ImagesFromClipboard {
    $inserted = 0

    if ([System.Windows.Clipboard]::ContainsImage()) {
        $bitmap = [System.Windows.Clipboard]::GetImage()
        if ($null -ne $bitmap) {
            Insert-ImageIntoEditor $bitmap
            $inserted++
        }
        return $inserted
    }

    if ([System.Windows.Clipboard]::ContainsFileDropList()) {
        foreach ($path in [System.Windows.Clipboard]::GetFileDropList()) {
            if (-not (Test-Path $path -PathType Leaf)) {
                continue
            }

            $extension = [System.IO.Path]::GetExtension($path)
            if (@('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tif', '.tiff', '.ico', '.wdp') -notcontains $extension.ToLowerInvariant()) {
                continue
            }

            try {
                Insert-ImageIntoEditor (New-BitmapSourceFromFile $path)
                $inserted++
            }
            catch {
            }
        }
    }

    return $inserted
}

function Apply-DefaultEditorFormattingToRange {
    param(
        [System.Windows.Documents.TextRange]$Range
    )

    if ($null -eq $Range -or $null -eq $script:State.EditorRichTextBox) {
        return
    }

    $Range.ApplyPropertyValue([System.Windows.Documents.TextElement]::FontFamilyProperty, $script:State.EditorRichTextBox.FontFamily)
    $Range.ApplyPropertyValue([System.Windows.Documents.TextElement]::FontSizeProperty, [double]$script:State.EditorRichTextBox.FontSize)
    $Range.ApplyPropertyValue([System.Windows.Documents.TextElement]::FontWeightProperty, [System.Windows.FontWeights]::Normal)
    $Range.ApplyPropertyValue([System.Windows.Documents.TextElement]::FontStyleProperty, [System.Windows.FontStyles]::Normal)
    $Range.ApplyPropertyValue([System.Windows.Documents.TextElement]::FontStretchProperty, [System.Windows.FontStretches]::Normal)
    $Range.ApplyPropertyValue([System.Windows.Documents.TextElement]::ForegroundProperty, [System.Windows.Media.Brushes]::Black)
    $Range.ApplyPropertyValue([System.Windows.Documents.TextElement]::BackgroundProperty, [System.Windows.Media.Brushes]::Transparent)
    $Range.ApplyPropertyValue([System.Windows.Documents.Inline]::TextDecorationsProperty, (New-Object System.Windows.TextDecorationCollection))
}

function Paste-PlainTextIntoEditor {
    if ($null -eq $script:State.EditorRichTextBox) {
        return $false
    }

    if (-not [System.Windows.Clipboard]::ContainsText()) {
        return $false
    }

    Ensure-EditorDocumentHasParagraph
    $text = [System.Windows.Clipboard]::GetText([System.Windows.TextDataFormat]::UnicodeText)
    $selection = $script:State.EditorRichTextBox.Selection
    $start = $selection.Start
    $selection.Text = $text
    $insertedRange = New-Object System.Windows.Documents.TextRange($start, $script:State.EditorRichTextBox.Selection.End)
    Apply-DefaultEditorFormattingToRange -Range $insertedRange
    $script:State.EditorRichTextBox.Focus() | Out-Null
    return $true
}

function Handle-EditorPaste {
    param(
        [switch]$PlainText
    )

    if ($PlainText) {
        [void](Paste-PlainTextIntoEditor)
        return
    }

    $imageCount = Insert-ImagesFromClipboard
    if ($imageCount -gt 0) {
        if ($imageCount -eq 1) {
            Set-LocalizedStatus 'status_image_pasted'
        }
        else {
            Set-LocalizedStatus 'status_images_pasted' @($imageCount)
        }
        return
    }

    $script:State.EditorRichTextBox.Paste()
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

function Apply-TopmostState {
    if ($null -eq $script:State.MainWindow -or $null -eq $script:State.Settings) {
        return
    }

    $script:State.MainWindow.Topmost = [bool]$script:State.Settings.AlwaysOnTop
    Refresh-MainWindowLanguage
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

    if (Test-HotkeyMatch -Role 'Close' -LogicalKey $LogicalKey -Hotkey $closeHotkey) {
        Invoke-CloseHotkeyAction
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

function Test-MainWindowMaximized {
    return (
        ($null -ne $script:State.MainWindow) -and
        ($script:State.MainWindow.WindowState -eq [System.Windows.WindowState]::Maximized)
    )
}

function Toggle-MainWindowMaximized {
    if ($null -eq $script:State.MainWindow) {
        return
    }

    if (Test-MainWindowMaximized) {
        $script:State.MainWindow.WindowState = [System.Windows.WindowState]::Normal
    }
    else {
        $script:State.MainWindow.WindowState = [System.Windows.WindowState]::Maximized
    }

    Refresh-MainWindowLanguage
}

function Start-MainWindowDrag {
    if ($null -eq $script:State.MainWindow) {
        return
    }

    try {
        $script:State.MainWindow.DragMove()
    }
    catch {
    }
}

function Focus-EditorWindow {
    if ($null -eq $script:State.MainWindow) {
        return
    }

    if (-not $script:State.MainWindow.IsVisible) {
        $script:State.MainWindow.Show()
    }

    if ($script:State.MainWindow.WindowState -eq [System.Windows.WindowState]::Minimized) {
        $script:State.MainWindow.WindowState = [System.Windows.WindowState]::Normal
    }
    $script:State.MainWindow.Activate() | Out-Null
    if ([bool]$script:State.Settings.AlwaysOnTop) {
        $script:State.MainWindow.Topmost = $true
    }
    else {
        $script:State.MainWindow.Topmost = $true
        $script:State.MainWindow.Topmost = $false
    }
    $script:State.EditorRichTextBox.Focus() | Out-Null
    $script:State.EditorRichTextBox.CaretPosition = $script:State.EditorRichTextBox.Document.ContentEnd
    $script:State.EditorRichTextBox.ScrollToEnd()
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

function Test-EditorWindowReadyToHide {
    if (-not $script:State.WindowVisible -or $null -eq $script:State.MainWindow) {
        return $false
    }

    if (-not $script:State.MainWindow.IsVisible) {
        return $false
    }

    if ($script:State.MainWindow.WindowState -eq [System.Windows.WindowState]::Minimized) {
        return $false
    }

    return [bool]$script:State.MainWindow.IsActive
}

function Invoke-CloseHotkeyAction {
    if (Test-EditorWindowReadyToHide) {
        Hide-EditorWindow
        return
    }

    Show-EditorWindow
}

function Toggle-EditorWindow {
    $closeHotkey = Get-EffectiveCloseHotkey $script:State.Settings
    if (Test-HotkeysEqual $script:State.Settings.InvokeHotkey $closeHotkey) {
        Invoke-CloseHotkeyAction
        return
    }

    Show-EditorWindow
}

function Toggle-AlwaysOnTop {
    if ($null -eq $script:State.Settings) {
        return
    }

    $script:State.Settings.AlwaysOnTop = -not [bool]$script:State.Settings.AlwaysOnTop
    Apply-TopmostState
    Save-Settings $script:State.Settings

    if ([bool]$script:State.Settings.AlwaysOnTop) {
        Set-LocalizedStatus 'status_topmost_enabled'
    }
    else {
        Set-LocalizedStatus 'status_topmost_disabled'
    }
}

function Copy-AllText {
    if (-not (Copy-EditorSelection -SelectAll)) {
        Set-LocalizedStatus 'status_nothing_to_copy'
        return
    }

    Set-LocalizedStatus 'status_all_copied'
}

function Copy-SelectedText {
    if (-not (Copy-EditorSelection)) {
        Set-LocalizedStatus 'status_select_first'
        return
    }

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
    Close-StartupToast
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
        MinWidth="320"
        MinHeight="150"
        WindowStartupLocation="CenterScreen"
        WindowStyle="None"
        ResizeMode="CanResize"
        ShowInTaskbar="False"
        Background="#FFF4EEE2"
        FontFamily="Segoe UI"
        SnapsToDevicePixels="True">
    <Window.Resources>
        <Style x:Key="ToolbarIconButtonStyle" TargetType="Button">
            <Setter Property="Width" Value="34" />
            <Setter Property="Height" Value="30" />
            <Setter Property="Padding" Value="2" />
            <Setter Property="Background" Value="Transparent" />
            <Setter Property="BorderBrush" Value="Transparent" />
            <Setter Property="BorderThickness" Value="0" />
            <Setter Property="Cursor" Value="Hand" />
            <Setter Property="Focusable" Value="False" />
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="Transparent" Padding="{TemplateBinding Padding}">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center" />
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
        </Style>
    </Window.Resources>
    <Grid Margin="12,8,12,10">
        <Grid.RowDefinitions>
            <RowDefinition Height="8" />
            <RowDefinition Height="*" />
            <RowDefinition Height="Auto" />
        </Grid.RowDefinitions>

        <Border x:Name="WindowDragHandle" Grid.Row="0" Background="Transparent" Cursor="SizeAll" />

        <Border Grid.Row="1" Background="#FFFFFCF6" BorderBrush="#E7DBC8" BorderThickness="1" CornerRadius="14" Padding="12">
            <RichTextBox x:Name="EditorRichTextBox"
                         Margin="0"
                         BorderThickness="0"
                         Background="Transparent"
                         FontFamily="Consolas"
                         FontSize="15"
                         AcceptsTab="True"
                         VerticalScrollBarVisibility="Auto"
                         HorizontalScrollBarVisibility="Auto"
                         SpellCheck.IsEnabled="False">
                <FlowDocument PagePadding="0">
                    <Paragraph />
                </FlowDocument>
            </RichTextBox>
        </Border>

        <Border Grid.Row="2" Background="#FFF4EEE2" BorderBrush="#FFF4EEE2" BorderThickness="1" CornerRadius="10" Margin="0,8,0,0" Padding="2,0">
            <StackPanel Orientation="Horizontal">
                <Button x:Name="CopyAllButton" Style="{StaticResource ToolbarIconButtonStyle}" Margin="0,0,6,0" />
                <Button x:Name="CopySelectionButton" Style="{StaticResource ToolbarIconButtonStyle}" Margin="0,0,6,0" />
                <Button x:Name="ClearButton" Style="{StaticResource ToolbarIconButtonStyle}" Margin="0,0,6,0" />
                <Button x:Name="TopmostButton" Style="{StaticResource ToolbarIconButtonStyle}" Margin="0,0,6,0" />
                <Button x:Name="FitWidthButton" Style="{StaticResource ToolbarIconButtonStyle}" Margin="0,0,6,0" />
                <Button x:Name="FitHeightButton" Style="{StaticResource ToolbarIconButtonStyle}" Margin="0,0,6,0" />
                <Button x:Name="SettingsButton" Style="{StaticResource ToolbarIconButtonStyle}" />
            </StackPanel>
        </Border>
    </Grid>
</Window>
"@

    $reader = New-Object System.Xml.XmlNodeReader ([xml]$xaml)
    $window = [Windows.Markup.XamlReader]::Load($reader)
    $window.Width = [double]$script:State.Settings.Window.Width
    $window.Height = [double]$script:State.Settings.Window.Height

    $script:State.MainWindow = $window
    $script:State.WindowDragHandle = $window.FindName('WindowDragHandle')
    $script:State.EditorRichTextBox = $window.FindName('EditorRichTextBox')
    $script:State.StatusTextBlock = $window.FindName('StatusTextBlock')
    $script:State.CopyAllButton = $window.FindName('CopyAllButton')
    $script:State.CopySelectionButton = $window.FindName('CopySelectionButton')
    $script:State.ClearButton = $window.FindName('ClearButton')
    $script:State.TopmostButton = $window.FindName('TopmostButton')
    $script:State.FitWidthButton = $window.FindName('FitWidthButton')
    $script:State.FitHeightButton = $window.FindName('FitHeightButton')
    $script:State.SettingsButton = $window.FindName('SettingsButton')
    Reset-EditorDocument

    $script:State.CopyAllButton.Add_Click({ Copy-AllText })
    $script:State.CopySelectionButton.Add_Click({ Copy-SelectedText })
    $script:State.ClearButton.Add_Click({
        Reset-EditorDocument
        Set-LocalizedStatus 'status_editor_cleared'
    })
    $script:State.TopmostButton.Add_Click({ Toggle-AlwaysOnTop })
    $script:State.FitWidthButton.Add_Click({ Resize-MainWindowToContent -FitWidth })
    $script:State.FitHeightButton.Add_Click({ Resize-MainWindowToContent -FitHeight })
    $script:State.SettingsButton.Add_Click({ Open-SettingsWindow })
    $script:State.WindowDragHandle.Add_MouseLeftButtonDown({
        param($sender, $eventArgs)

        if ($eventArgs.ChangedButton -ne [System.Windows.Input.MouseButton]::Left) {
            return
        }

        if ($eventArgs.ClickCount -gt 1) {
            Toggle-MainWindowMaximized
            return
        }

        Start-MainWindowDrag
    })
    $script:State.EditorRichTextBox.Add_PreviewKeyDown({
        param($sender, $eventArgs)

        $plainTextPasteModifiers = [System.Windows.Input.ModifierKeys]::Control -bor [System.Windows.Input.ModifierKeys]::Shift
        $isPlainTextPasteGesture = (
            ($eventArgs.Key -eq [System.Windows.Input.Key]::V) -and
            ([System.Windows.Input.Keyboard]::Modifiers -eq $plainTextPasteModifiers)
        )
        if ($isPlainTextPasteGesture) {
            $eventArgs.Handled = $true
            Handle-EditorPaste -PlainText
            return
        }

        $isPasteGesture = (
            (($eventArgs.Key -eq [System.Windows.Input.Key]::V) -and ([System.Windows.Input.Keyboard]::Modifiers -eq [System.Windows.Input.ModifierKeys]::Control)) -or
            (($eventArgs.Key -eq [System.Windows.Input.Key]::Insert) -and ([System.Windows.Input.Keyboard]::Modifiers -eq [System.Windows.Input.ModifierKeys]::Shift))
        )
        if ($isPasteGesture) {
            $eventArgs.Handled = $true
            Handle-EditorPaste
        }
    })
    Apply-TopmostState
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

    $window.Add_StateChanged({
        Refresh-MainWindowLanguage
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

        <StackPanel Grid.Row="11">
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
        Show-AppMessage (Get-Text 'msg_reset_default_hotkey' -FormatArgs @((Format-Hotkey $fallback.InvokeHotkey)))
    }

    Initialize-HotkeyPolling
    Refresh-LocalizedUI
    Show-StartupToast (Get-Text 'notify_balloon_text' -FormatArgs @((Format-Hotkey $script:State.Settings.InvokeHotkey)))
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
