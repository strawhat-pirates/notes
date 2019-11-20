Add-Type -AssemblyName System.Windows.Forms

$actions =
    'window-switch',
    'tab-switch',
    'keystroke',
    'keystroke',
    'pagedown',
    'pagedown',
    'pageup',
    'pageup',
    'mousemove'

$shell = New-Object -com "Wscript.Shell"
$sleep = 2
$screen = [System.Windows.Forms.SystemInformation]::VirtualScreen
$SetForegroundWindow = Add-Type -MemberDefinition '[DllImport("user32.dll")]public static extern bool SetForegroundWindow(IntPtr hWnd);' -Name "Win32SetForegroundWindow" -Namespace Win32Functions -PassThru

While($true) {
    $caps = [console]::CapsLock

    if ($caps -eq $false) {
        Start-Sleep 3
        continue
    }

    $selected = $actions[(Get-Random -Maximum ($actions).count)]
    $secs = Get-Random -Minimum 1 -Maximum $sleep

    switch ($selected) {
        'tab-switch' {
            $shell.sendkeys("^{TAB}")
            continue
        }
        'keystroke' {
            $shell.sendkeys("{F15}")
            continue
        }
        'pagedown' {
            $shell.sendkeys("{PGDN}")
            continue
        }
        'pageup' {
            $shell.sendkeys("{PGUP}")
            continue
        }
        'mousemove' {
            $x = Get-Random -Minimum 1 -Maximum $screen.Width
            $y = Get-Random -Minimum 1 -Maximum $screen.Height
            [System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($x, $y)
        }
        'window-switch' {
            $process = Get-Process | Where-Object {$_.MainWindowHandle -ne 0} | Where-Object {$_.ProcessName -ne "cmd"}  | Where-Object {$_.ProcessName -ne "HubstaffClient"}
            $pindex = Get-Random -Maximum (($process).count)
            [void] $SetForegroundWindow::SetForegroundWindow($process[$pindex].MainWindowHandle)
        }
    }

    Start-Sleep -Seconds $secs
}
