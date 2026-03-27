#requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

param(
    # Internal flag used when the scheduled task re-invokes the script
    [switch]$Resume,

    # Optional: install Microsoft Update (Office/other MS products) as well as Windows updates
    [switch]$MicrosoftUpdate = $true,

    # Safety valve to avoid endless loops if an update continually fails
    [ValidateRange(1, 50)]
    [int]$MaxRuns = 12
)

try {
    # ------------------------------------------------------------
    # Quiet / setup-friendly defaults
    # ------------------------------------------------------------
    $env:MG_SHOW_WELCOME_MESSAGE = 'false' # harmless here, keeps console quieter in mixed environments
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force | Out-Null
    $ProgressPreference = 'SilentlyContinue'

    # ------------------------------------------------------------
    # Paths / state / logging
    # ------------------------------------------------------------
    $Root      = Join-Path $env:ProgramData 'WBU-WindowsUpdate'
    $StatePath = Join-Path $Root 'state.json'
    $LogPath   = Join-Path $Root 'update.log'
    $TaskName  = 'WBU-WindowsUpdate-Resume'

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
            # If PowerShellGet can't be updated (common in early setup), fall back to in-box
            Import-Module PowerShellGet -ErrorAction SilentlyContinue | Out-Null
        }
    }

    # ------------------------------------------------------------
    # Resume mechanism: startup scheduled task that re-downloads and runs this script
    # URL is supplied via env var from your USB .bat (not hardcoded here)
    # ------------------------------------------------------------
    function Get-SourceUrl {
        # The .bat sets this before calling iex(irm ...)
        $url = $env:WBU_WU_SOURCE_URL

        if ([string]::IsNullOrWhiteSpace($url)) {
            throw "Missing environment variable WBU_WU_SOURCE_URL. Your USB .bat must set it to your bit.ly URL before launching the script."
        }

        return $url
    }

    function Ensure-ResumeScheduledTask {
        Import-Module ScheduledTasks -ErrorAction SilentlyContinue | Out-Null

        $existing = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
        if ($existing) { return }

        $url = Get-SourceUrl

        Write-Log "Creating Scheduled Task: $TaskName"

        # Startup action:
        # - Wait briefly for networking/services
        # - Set env var for this process
        # - Download and invoke the script again with -Resume
        $cmd = @"
Start-Sleep -Seconds 30;
`$env:WBU_WU_SOURCE_URL = '$url';
iex (irm `$env:WBU_WU_SOURCE_URL) -Resume
"@

        $action  = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-NoProfile -ExecutionPolicy Bypass -Command $([Management.Automation.Language.CodeGeneration]::QuoteArgument($cmd))"
        $trigger = New-ScheduledTaskTrigger -AtStartup

        $settings = New-ScheduledTaskSettingsSet `
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
            # Best effort: enable Microsoft Update catalogue (Office/other MS products)
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

    function Install-Updates {
        $params = @{
