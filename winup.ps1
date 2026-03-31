#requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

try {
    # ------------------------------------------------------------
    # OOBE-safe hardening / suppression
    # ------------------------------------------------------------
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force | Out-Null
    $ProgressPreference = 'SilentlyContinue'
    $ConfirmPreference  = 'None'   # suppress ShouldProcess confirmation prompts

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
    if (-not (Get-Module -ListAvailable -Name PSWindowsUpdate -ErrorAction SilentlyContinue)) {
        Install-Module -Name PSWindowsUpdate -Force -SkipPublisherCheck -AllowClobber -Scope AllUsers | Out-Null
    }
    Import-Module PSWindowsUpdate -Force -ErrorAction Stop | Out-Null

    # ------------------------------------------------------------
    # Enable Microsoft Update (Office etc.) without prompting
    # ------------------------------------------------------------
    try {
        $mu = Get-WUServiceManager -ErrorAction SilentlyContinue | Where-Object { $_.Name -match 'Microsoft Update' }
        if (-not $mu) {
            Add-WUServiceManager -MicrosoftUpdate -Confirm:$false -ErrorAction SilentlyContinue | Out-Null
        }
    } catch { }

    # ------------------------------------------------------------
    # Fully unattended install
    # ------------------------------------------------------------
    $updates = Get-WindowsUpdate -MicrosoftUpdate -IgnoreUserInput -ErrorAction SilentlyContinue

    if (-not $updates -or $updates.Count -eq 0) {
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
        -ErrorAction Stop | Out-Null

    # If reboot is required, the system will reboot and this process ends.
    Write-Output "Update pass completed. Rebooting if required."
    exit 0
}
catch {
    Write-Error $_
    exit 1
}
