$sig = '
[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
[DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
'
Add-Type -MemberDefinition $sig -name NativeMethods -namespace Win32
Add-Type -AssemblyName Microsoft.VisualBasic
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName UIAutomationClient

function FocusApp($appName)
{
    $p = GetRootProcess $appName
    if ($p)
    {
        FocusProcess $p
        return $true
    }
    else
    {
        return $false
    }
}

function GetRootProcess($appName)
{
    $childCount = @{}
    $processes = Get-CimInstance -class win32_process -filter "Name = '$appName'"
    if (-not $processes)
    {
        return
    }

    foreach ($process in $processes)
    {
        $parent = $process.ParentProcessId
        if ($childCount.ContainsKey($parent))
        {
            $childCount[$parent] += 1
        }
        else
        {
            $childCount[$parent] = 1
        }
    }
    $maxChildCount = 0
    $rootProcess = $processes[0]
    foreach ($process in $processes)
    {
        if ($childCount[$process.ProcessId] -gt $maxChildCount)
        {
            $rootProcess = $process
        }
    }
    Get-Process -Id $rootProcess.ProcessId
}

function FocusProcess($process)
{
    $state = GetWindowState $process
    if ($state -eq "")
    {
        # The process has no window
        ActivateProcess $process
        return
    }

    if ($state -eq "Minimized")
    {
        ShowWindow $process.MainWindowHandle
    }
    elseif (($state -eq "Maximized") -or ($state -eq "Normal"))
    {
        [Win32.NativeMethods]::SetForegroundWindow($process.MainWindowHandle) | Out-Null
    }
}

function GetWindowState($process)
{
    try
    {
        $automationElement = [System.Windows.Automation.AutomationElement]::FromHandle($process.MainWindowHandle)
    }
    catch
    {
        return ""
    }
    $pattern = $automationElement.GetCurrentPattern([System.Windows.Automation.WindowPatternIdentifiers]::Pattern)
    $pattern.Current.WindowVisualState.ToString()
}

function ShowWindow($windowHandle)
{
    try
    {
        $SW_SHOWDEFAULT = 10
        [Win32.NativeMethods]::ShowWindow($windowHandle, $SW_SHOWDEFAULT) | Out-Null
    }
    catch
    {
        Write-Host "ShowWindow failed. [$PSItem]"
    }
}

function ActivateProcess($process)
{
    [Microsoft.VisualBasic.Interaction]::AppActivate($process.id) | Out-Null
}

function SendKeysToActiveApp($keys)
{
    [System.Windows.Forms.SendKeys]::SendWait($keys) | Out-Null
}

