if (-not (Get-Command Write-Log -ErrorAction SilentlyContinue)) {
    . (Join-Path $PSScriptRoot 'utils.ps1')
}

function Test-IsAdmin {
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal($identity)
        return $principal.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
    }
    catch {
        return $false
    }
}

function Get-WindowsVersionInfo {
    try {
        $cv = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop
        $displayVersion = $cv.DisplayVersion
        if (-not $displayVersion) {
            $displayVersion = $cv.ReleaseId
        }

        $productName = $cv.ProductName
        $buildNumber = $cv.CurrentBuildNumber
        $ubr = $cv.UBR
        $edition = $cv.EditionID

        [pscustomobject]@{
            Status         = 'Installed'
            Message        = "$productName $displayVersion (Build $buildNumber.$ubr)"
            ProductName    = $productName
            DisplayVersion = $displayVersion
            BuildNumber    = $buildNumber
            UBR            = $ubr
            Edition        = $edition
            Raw            = $cv
        }
    }
    catch {
        [pscustomobject]@{
            Status      = 'Unknown'
            Message     = 'Unable to read Windows version information.'
            Details     = $_.Exception.Message
            ProductName = ''
            DisplayVersion = ''
            BuildNumber = ''
            UBR = ''
            Edition = ''
        }
    }
}

function Get-PowerShellVersionInfo {
    try {
        $version = $PSVersionTable.PSVersion
        $edition = $PSVersionTable.PSEdition
        [pscustomobject]@{
            Status   = if ($version.Major -ge 5) { 'Installed' } else { 'Unknown' }
            Message  = "$edition $version"
            Version  = $version.ToString()
            Major    = $version.Major
            Minor    = $version.Minor
            Patch    = $version.Build
            Edition  = $edition
            IsPwsh   = ($edition -eq 'Core')
        }
    }
    catch {
        [pscustomobject]@{
            Status   = 'Unknown'
            Message  = 'Unable to read PowerShell version.'
            Details  = $_.Exception.Message
            Version  = ''
            Major    = 0
            Minor    = 0
            Patch    = 0
            Edition  = ''
            IsPwsh   = $false
        }
    }
}

function Test-CommandExists {
    param(
        [Parameter(Mandatory)]
        [string]$Name
    )

    return [bool](Get-Command $Name -ErrorAction SilentlyContinue)
}

function Test-WingetExists {
    return Test-CommandExists -Name 'winget'
}

function Test-InternetConnection {
    param(
        [string]$Target = 'https://www.microsoft.com'
    )

    try {
        $request = [System.Net.WebRequest]::Create($Target)
        $request.Method = 'HEAD'
        $request.Timeout = 2000
        $response = $request.GetResponse()
        $response.Close()
        return $true
    }
    catch {
        return $false
    }
}

function Get-SystemSummary {
    $admin = Test-IsAdmin
    $windows = Get-WindowsVersionInfo
    $ps = Get-PowerShellVersionInfo
    $winget = Test-WingetExists
    $internet = $null

    [pscustomobject]@{
        Admin       = $admin
        Windows     = $windows
        PowerShell  = $ps
        Winget      = $winget
        Internet    = $internet
        Timestamp   = Get-Date
    }
}

