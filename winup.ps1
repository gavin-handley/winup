#requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

try {
    # ------------------------------------------------------------
    # Quiet, OOBE-safe defaults
    # ------------------------------------------------------------

    # TLS 1.2 for gallery/module downloads
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    # Execution policy for this process only
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force | Out-Null

    # Suppress noisy progress output
    $ProgressPreference = 'SilentlyContinue'

    # ------------------------------------------------------------
    # Bootstrap PowerShellGet / NuGet / PSGallery trust
    # ------------------------------------------------------------

    if (-not (Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue)) {
        Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force | Out-Null
    }

    $psg = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
    if ($psg -and $psg.InstallationPolicy -ne 'Trusted') {
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted | Out-Null
    }

    try {
        Install-Module -Name PowerShellGet -Force -AllowClobber -Scope AllUsers | Out-Null
        Import-Module PowerShellGet -Force | Out-Null
    }
    catch {
        Import-Module PowerShellGet -ErrorAction SilentlyContinue | Out-Null
    }

    # ------------------------------------------------------------
    # Install + import PSWindowsUpdate
    # ------------------------------------------------------------

    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate)) {
        Install-Module -Name PSWindowsUpdate -Force -SkipPublisherCheck -AllowClobber -Scope AllUsers | Out-Null
    }

    Import-Module PSWindowsUpdate -Force

    # Optional but useful: enable Microsoft Update (Office, etc.)
    try { Add-WUServiceManager -MicrosoftUpdate -ErrorAction SilentlyContinue | Out-Null } catch { }

    # ------------------------------------------------------------
    # Install everything, reboot if needed
    # ------------------------------------------------------------

    $updates = Get-WindowsUpdate -MicrosoftUpdate -IgnoreUserInput

    if (-not $updates) {
        Write-Output "No updates available."
        exit 0
    }

    Write-Output "Installing $($updates.Count) update(s)..."
    Install-WindowsUpdate -MicrosoftUpdate -AcceptAll -AutoReboot -IgnoreUserInput

    # If a reboot is required, execution stops here automatically.
    Write-Output "Update pass completed. Reboot if prompted, then re-run this script to continue."
    exit 0
}
catch {
    Write-Error $_
    exit 1
}
