#requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

try {
    # OOBE-safe hardening
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12  # TLS 1.2 guidance for PSWindowsUpdate install. [1](https://github.com/mgajda83/PSWindowsUpdate/blob/main/README.md)[2](https://github.com/mgajda83/PSWindowsUpdate)
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force | Out-Null
    $ConfirmPreference = 'None'

    # Keep bootstrap quiet, but allow visible output later
    $ProgressPreference = 'SilentlyContinue'

    # NuGet + PSGallery trust
    if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force | Out-Null
    }
    $psg = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
    if ($psg -and $psg.InstallationPolicy -ne 'Trusted') {
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted | Out-Null
    }

    # Install/import PSWindowsUpdate
    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate -ErrorAction SilentlyContinue)) {
        Install-Module -Name PSWindowsUpdate -Force -SkipPublisherCheck -AllowClobber -Scope AllUsers | Out-Null
    }
    Import-Module PSWindowsUpdate -Force -ErrorAction Stop | Out-Null

    # Enable Microsoft Update catalogue (Office etc.) without prompting
    try {
        $mu = Get-WUServiceManager -ErrorAction SilentlyContinue | Where-Object { $_.Name -match 'Microsoft Update' }
        if (-not $mu) {
            Add-WUServiceManager -MicrosoftUpdate -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        }
    } catch { }

    # Show progress during update operations
    $ProgressPreference = 'Continue'

    # Scan
    $updates = Get-WindowsUpdate -MicrosoftUpdate -IgnoreUserInput -Verbose -ErrorAction SilentlyContinue

    if (-not $updates -or $updates.Count -eq 0) {
        Write-Output "No updates available."
        exit 0
    }

    Write-Output ("Installing {0} update(s)..." -f $updates.Count)

    # Install (unattended, but verbose)
    Install-WindowsUpdate `
        -MicrosoftUpdate `
        -AcceptAll `
        -AutoReboot `
        -IgnoreUserInput `
        -Confirm:$false `
        -Verbose `
        -ErrorAction Stop

    Write-Output "Update pass completed. Rebooting if required."
    exit 0
}
catch {
    Write-Error $_
    exit 1
}
