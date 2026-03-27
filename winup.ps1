#requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

try {
    # ------------------------------------------------------------
    # OOBE-safe hardening
    # ------------------------------------------------------------

    # TLS 1.2 for module downloads
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    # Execution policy for this process only
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force | Out-Null

    # Suppress progress noise
    $ProgressPreference = 'SilentlyContinue'

    # ------------------------------------------------------------
    # Bootstrap NuGet / PSGallery
    # ------------------------------------------------------------

    if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force | Out-Null
    }

    $psg = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
    if ($psg -and $psg.InstallationPolicy -ne 'Trusted') {
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted | Out-Null
    }

    # ------------------------------------------------------------
    # Install + import PSWindowsUpdate
    # ------------------------------------------------------------

    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
        Install-Module -Name PSWindowsUpdate -Force -SkipPublisherCheck -AllowClobber -Scope AllUsers | Out-Null
    }

    Import-Module PSWindowsUpdate -Force

    # Enable Microsoft Update (Office, etc.) if available
    try { Add-WUServiceManager -MicrosoftUpdate -ErrorAction SilentlyContinue | Out-Null } catch { }

    # ------------------------------------------------------------
    # Fully unattended update install
    # ------------------------------------------------------------

    $updates = Get-WindowsUpdate -MicrosoftUpdate -IgnoreUserInput -ErrorAction SilentlyContinue

    if (-not $updates) {
        Write-Output "No updates available."
        exit 0
    }

    Write-Output "Installing $($updates.Count) update(s)..."

    Install-WindowsUpdate `
        -MicrosoftUpdate `
        -AcceptAll `
        -AutoReboot `
        -IgnoreUserInput `
        -Confirm:$false `
        -ErrorAction Stop

    # If reboot is required, execution ends here automatically
    Write-Output "Update pass completed. Rebooting if required."
    exit 0
}
catch {
    Write-Error $_
    exit 1
}
