#requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

param(
    [switch]$Resume,
    [switch]$MicrosoftUpdate = $true,
    [ValidateRange(1, 50)]
    [int]$MaxRuns = 12
)

try {
    # ------------------------------------------------------------
    # Quiet / setup-friendly defaults
    # ------------------------------------------------------------
    $env:MG_SHOW_WELCOME_MESSAGE = 'false'
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force | Out-Null
    $ProgressPreference = 'SilentlyContinue'

    # ------------------------------------------------------------
    # Paths / state / logging / cached local runner
    # ------------------------------------------------------------
    $Root        = Join-Path $env:ProgramData 'WBU-WindowsUpdate'
    $StatePath   = Join-Path $Root 'state.json'
    $LogPath     = Join-Path $Root 'update.log'
    $LocalScript = Join-Path $Root 'Run-WindowsUpdate.ps1'
    $TaskName    = 'WBU-WindowsUpdate-Resume'

    New-Item -Path $Root -ItemType Directory -Force | Out-Null

    function Write-Log {
        param([string]$Message)
        $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        $line = "$ts  $Message"
        Add-Content -LiteralPath $LogPath -Value $line -ErrorAction SilentlyContinue
        Write-Output $Message
    }

    function Load-State {
        if (Test-Path -LiteralPath $StatePath) {
            try { return (Get-Content -LiteralPath $StatePath -Raw | ConvertFrom-Json) } catch { }
        }
        return [pscustomobject]@{ RunCount = 0; LastRun = $null }
    }

    function Save-State($state) {
        $state.LastRun = (Get-Date).ToString('o')
        ($state | ConvertTo-Json -Depth 5) | Set-Content -LiteralPath $StatePath -Force
    }

    # ------------------------------------------------------------
    # Bootstrap: NuGet + PSGallery trust + PowerShellGet best effort
    # ------------------------------------------------------------
    function Ensure-NuGetAndGalleryTrust {
        if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force | Out-Null
        }

        $psg = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
        if ($psg -and $psg.InstallationPolicy -ne 'Trusted') {
            Set-PSRepository -Name PSGallery -InstallationPolicy Trusted | Out-Null
        }

        try {
            Install-Module -Name PowerShellGet -Force -AllowClobber -Scope AllUsers -ErrorAction Stop | Out-Null
            Import-Module PowerShellGet -Force -ErrorAction Stop | Out-Null
        }
        catch {
            Import-Module PowerShellGet -ErrorAction SilentlyContinue | Out-Null
        }
    }

    # ------------------------------------------------------------
    # Cache this running script locally (no URL required)
    # ------------------------------------------------------------
    function Ensure-LocalScriptCache {
        if (Test-Path -LiteralPath $LocalScript) { return }

        # When invoked via iex(irm ...), Definition contains the whole script text.
        $scriptText = $MyInvocation.MyCommand.Definition

        if ([string]::IsNullOrWhiteSpace($scriptText) -or $scriptText.Length -lt 200) {
            throw "Unable to cache script locally. Script text was unexpectedly empty/short."
        }

        Set-Content -LiteralPath $LocalScript -Value $scriptText -Encoding UTF8 -Force
        Write-Log "Cached local script: $LocalScript"
    }

    # ------------------------------------------------------------
    # Resume mechanism: startup scheduled task that runs the cached local script
    # ------------------------------------------------------------
    function Ensure-ResumeScheduledTask {
        Import-Module ScheduledTasks -ErrorAction SilentlyContinue | Out-Null

        $existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
        if ($existing) { return }

        Ensure-LocalScriptCache

        Write-Log "Creating Scheduled Task: $TaskName"

        $args = "-NoProfile -ExecutionPolicy Bypass -File `"$LocalScript`" -Resume" +
                ($(if ($MicrosoftUpdate) { " -MicrosoftUpdate" } else { "" })) +
                " -MaxRuns $MaxRuns"

        $action    = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $args
        $trigger   = New-ScheduledTaskTrigger -AtStartup
        $settings  = New-ScheduledTaskSettingsSet `
            -StartWhenAvailable `
            -AllowStartIfOnBatteries `
            -DontStopIfGoingOnBatteries `
            -ExecutionTimeLimit (New-TimeSpan -Hours 6)

        $principal = New-ScheduledTaskPrincipal -UserId 'SYSTEM' -LogonType ServiceAccount -RunLevel Highest

        $task = New-ScheduledTask -Action $action -Trigger $trigger -Settings $settings -Principal $principal
        Register-ScheduledTask -TaskName $TaskName -InputObject $task -Force | Out-Null
    }

    function Remove-ResumeScheduledTask {
        $existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
        if ($existing) {
            Write-Log "Removing Scheduled Task: $TaskName"
            Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        }
    }

    # ------------------------------------------------------------
    # PSWindowsUpdate setup
    # ------------------------------------------------------------
    function Ensure-PSWindowsUpdate {
        if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate -ErrorAction SilentlyContinue)) {
            Write-Log "Installing PSWindowsUpdate module..."
            Install-Module -Name PSWindowsUpdate -Force -SkipPublisherCheck -AllowClobber -Scope AllUsers | Out-Null
        }

        Import-Module PSWindowsUpdate -Force -ErrorAction Stop | Out-Null

        if ($MicrosoftUpdate) {
            try { Add-WUServiceManager -MicrosoftUpdate -ErrorAction SilentlyContinue | Out-Null } catch { }
        }
    }

    function Get-PendingUpdatesCount {
        $params = @{
            IgnoreUserInput = $true
            ErrorAction     = 'SilentlyContinue'
        }
        if ($MicrosoftUpdate) { $params['MicrosoftUpdate'] = $true }

        $updates = Get-WindowsUpdate @params
        if ($updates) { return $updates.Count }
        return 0
    }

    function Install-UpdatesOnce {
        $params = @{
            AcceptAll       = $true
            AutoReboot      = $true
            IgnoreUserInput = $true
            Confirm         = $false
            ErrorAction     = 'Stop'
        }
        if ($MicrosoftUpdate) { $params['MicrosoftUpdate'] = $true }

        Install-WindowsUpdate @params | Out-Null
    }

    # ------------------------------------------------------------
    # Main
    # ------------------------------------------------------------
    Write-Log "Starting Windows Update runner. Resume = $Resume"

    Ensure-NuGetAndGalleryTrust
    Ensure-LocalScriptCache
    Ensure-ResumeScheduledTask
    Ensure-PSWindowsUpdate

    $state = Load-State
    $state.RunCount = [int]$state.RunCount + 1
    Save-State $state

    Write-Log "RunCount = $($state.RunCount) (MaxRuns = $MaxRuns)"

    if ($state.RunCount -gt $MaxRuns) {
        Write-Log "MaxRuns exceeded. Cleaning up scheduled task to prevent looping."
        Remove-ResumeScheduledTask
        throw "Stopped after $MaxRuns runs. Check Windows Update for stuck/failed updates."
    }

    $pending = Get-PendingUpdatesCount
    if ($pending -le 0) {
        Write-Log "No updates found. Cleaning up scheduled task and state."
        Remove-ResumeScheduledTask
        Remove-Item -LiteralPath $StatePath -Force -ErrorAction SilentlyContinue
        Write-Log "Completed."
        exit 0
    }

    Write-Log "Found $pending update(s). Installing with AutoReboot..."
    Install-UpdatesOnce

    # If a reboot happens, this process ends; the scheduled task resumes at next startup.
    Write-Log "Install pass complete. If reboot required, resume will occur automatically."
    exit 0
}
catch {
    try {
        $Root = Join-Path $env:ProgramData 'WBU-WindowsUpdate'
        New-Item -Path $Root -ItemType Directory -Force | Out-Null
        $LogPath = Join-Path $Root 'update.log'
        Add-Content -LiteralPath $LogPath -Value ("{0}  ERROR  {1}" -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $_.Exception.Message) -ErrorAction SilentlyContinue
    } catch { }

    Write-Error $_
    exit 1
}
