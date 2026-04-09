#requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$appId
)

# ---------------------------[ Config ]---------------------------
# Single package id per deployment (-appId); no catalog blacklist/whitelist.

$logDataRoot = "$env:ProgramData\WinGetNable"

# ---------------------------[ Script Start Timestamp ]---------------------------
$scriptStartTime = Get-Date

function Get-SanitizedPathSegment {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$rawSegment)

    $invalidChars = [System.IO.Path]::GetInvalidFileNameChars()
    $stringBuilder = [System.Text.StringBuilder]::new()
    foreach ($character in $rawSegment.ToCharArray()) {
        if ($invalidChars -contains $character) {
            [void]$stringBuilder.Append('_')
        }
        else {
            [void]$stringBuilder.Append($character)
        }
    }
    $trimmed = $stringBuilder.ToString().Trim()
    if ([string]::IsNullOrWhiteSpace($trimmed)) {
        return 'unknown'
    }
    return $trimmed
}

$sanitizedAppIdFolder = Get-SanitizedPathSegment -rawSegment $appId
$scriptName = "detection | $appId"
$logFileName = 'detection.log'

# ---------------------------[ Logging Setup ]---------------------------
$log = $true
$logDebug = $false
$logGet = $true
$logRun = $true
$enableLogFile = $true

$logFileDirectory = Join-Path -Path $logDataRoot -ChildPath $sanitizedAppIdFolder
$logFile = Join-Path -Path $logFileDirectory -ChildPath $logFileName

if ($enableLogFile -and -not (Test-Path -Path $logFileDirectory)) {
    New-Item -ItemType Directory -Path $logFileDirectory -Force | Out-Null
}

function Write-Log {
    [CmdletBinding()]
    param (
        [string]$Message,
        [string]$Tag = 'Info'
    )

    if (-not $log) { return }

    if (($Tag -eq 'Debug') -and (-not $logDebug)) { return }
    if (($Tag -eq 'Get') -and (-not $logGet)) { return }
    if (($Tag -eq 'Run') -and (-not $logRun)) { return }

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $tagList = @('Start', 'Get', 'Run', 'Info', 'Success', 'Error', 'Debug', 'End')
    $rawTag = $Tag.Trim()

    if ($tagList -contains $rawTag) {
        $rawTag = $rawTag.PadRight(7)
    }
    else {
        $rawTag = 'Error  '
    }

    $color = switch ($rawTag.Trim()) {
        'Start' { 'Cyan' }
        'Get' { 'Blue' }
        'Run' { 'Magenta' }
        'Info' { 'Yellow' }
        'Success' { 'Green' }
        'Error' { 'Red' }
        'Debug' { 'DarkYellow' }
        'End' { 'Cyan' }
        default { 'White' }
    }

    $logMessage = "$timestamp [  $rawTag ] $Message"

    if ($enableLogFile) {
        try {
            Add-Content -Path $logFile -Value $logMessage -Encoding UTF8
        }
        catch {
        }
    }

    Write-Host "$timestamp " -NoNewline
    Write-Host '[  ' -NoNewline -ForegroundColor White
    Write-Host "$rawTag" -NoNewline -ForegroundColor $color
    Write-Host ' ] ' -NoNewline -ForegroundColor White
    Write-Host "$Message"
}

function Complete-Script {
    param([int]$exitCode)

    $scriptEndTime = Get-Date
    $duration = $scriptEndTime - $scriptStartTime

    Write-Log "Runtime $($duration.ToString('hh\:mm\:ss\.ff'))" -Tag 'Info'
    Write-Log "Exit $exitCode" -Tag 'Info'
    Write-Log '==================== End ====================' -Tag 'End'

    exit $exitCode
}

function Get-WingetPath {
    [CmdletBinding()]
    param()

    $wingetBase = "$env:ProgramW6432\WindowsApps"
    $patterns = @(
        'Microsoft.DesktopAppInstaller_*_x64__8wekyb3d8bbwe'
        'Microsoft.DesktopAppInstaller_*_arm64__8wekyb3d8bbwe'
    )

    foreach ($pattern in $patterns) {
        $wingetFolders = Get-ChildItem -Path $wingetBase -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -like $pattern }

        if (-not $wingetFolders) { continue }

        $candidates = foreach ($folder in $wingetFolders) {
            $exePath = Join-Path -Path $folder.FullName -ChildPath 'winget.exe'
            if (-not (Test-Path -LiteralPath $exePath)) { continue }
            try {
                $fileVersion = (Get-Item -LiteralPath $exePath -ErrorAction Stop).VersionInfo.FileVersionRaw
            }
            catch {
                $fileVersion = $null
            }
            [pscustomobject]@{
                path = $exePath
                fileVersion = $fileVersion
                creationTime = $folder.CreationTime
            }
        }

        if (-not $candidates) { continue }

        $latest = $candidates |
            Sort-Object -Property { $_.fileVersion }, creationTime -Descending |
            Select-Object -First 1

        if ($latest.path) {
            return $latest.path
        }
    }

    $userWingetPath = "$env:LocalAppData\Microsoft\WindowsApps\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe\winget.exe"
    if (Test-Path -LiteralPath $userWingetPath) {
        return $userWingetPath
    }

    Write-Log 'Winget: no DesktopAppInstaller / winget.exe' -Tag 'Error'
    throw 'Winget not found in system or user context'
}

function Test-WingetVersion {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$wingetPath)

    $versionOutput = @(& $wingetPath --version 2>&1)
    $exitCode = $LASTEXITCODE
    $versionString = ($versionOutput | Out-String).Trim()
    $isHealthy = ($exitCode -eq 0)
    Write-Log "WinGet --version: exit $exitCode" -Tag 'Debug'
    if ($isHealthy) {
        $versionLine = $versionOutput | Where-Object { $_ -and ($_ -match '\d+\.\d+') } | Select-Object -First 1
        if ($versionLine -and $versionLine -match '(\d+\.\d+(?:\.\d+)?(?:\.\d+)?)') {
            Write-Log "WinGet: v$($matches[1])" -Tag 'Success'
        }
        else {
            Write-Log 'WinGet' -Tag 'Success'
        }
    }
    else {
        Write-Log "Winget failed: exit $exitCode" -Tag 'Error'
        $errorOutput = $versionOutput | Where-Object { $_ -and $_ -notmatch '^\s*$' } | Select-Object -First 3
        if ($errorOutput) {
            Write-Log "Details: $($errorOutput -join '; ')" -Tag 'Debug'
        }
    }
    return @{
        isHealthy = $isHealthy
        version = $versionString
        exitCode = $exitCode
    }
}

function ConvertFrom-WingetUpgradeOutput {
    [CmdletBinding()]
    param(
        [string]$rawOutput,
        [string]$scope
    )

    $updates = [System.Collections.ArrayList]::new()
    if (-not ($rawOutput -match '-----')) {
        return , @()
    }

    $lines = $rawOutput.Split([Environment]::NewLine) | Where-Object { $_ }
    for ($lineIndex = 0; $lineIndex -lt $lines.Count; $lineIndex++) {
        $lines[$lineIndex] = $lines[$lineIndex] -replace '[\u2026]', ' '
    }

    $headerLineIndex = 0
    while ($headerLineIndex -lt $lines.Count -and -not $lines[$headerLineIndex].StartsWith('-----')) { $headerLineIndex++ }
    if ($headerLineIndex -ge $lines.Count) { return , @() }
    $headerLineIndex = $headerLineIndex - 1
    if ($headerLineIndex -lt 0) { return , @() }

    $indexParts = $lines[$headerLineIndex] -split '(?<=\s)(?!\s)'
    if ($indexParts.Count -lt 3) { return , @() }

    $idStart = $($indexParts[0] -replace '[\u4e00-\u9fa5]', '**').Length
    $versionStart = $idStart + $($indexParts[1] -replace '[\u4e00-\u9fa5]', '**').Length
    $availableStart = $versionStart + $($indexParts[2] -replace '[\u4e00-\u9fa5]', '**').Length

    for ($lineIndex = $headerLineIndex + 2; $lineIndex -lt $lines.Count; $lineIndex++) {
        $line = $lines[$lineIndex]
        if ($line.StartsWith('-----')) {
            $headerLineIndex = $lineIndex - 1
            $indexParts = $lines[$headerLineIndex] -split '(?<=\s)(?!\s)'
            $idStart = $($indexParts[0] -replace '[\u4e00-\u9fa5]', '**').Length
            $versionStart = $idStart + $($indexParts[1] -replace '[\u4e00-\u9fa5]', '**').Length
            $availableStart = $versionStart + $($indexParts[2] -replace '[\u4e00-\u9fa5]', '**').Length
            continue
        }
        if ($line -match '\w\.\w') {
            $nameAdjust = $($line.Substring(0, $idStart) -replace '[\u4e00-\u9fa5]', '**').Length - $line.Substring(0, $idStart).Length
            $packageName = $line.Substring(0, $idStart - $nameAdjust).TrimEnd()
            $packageIdRow = $line.Substring($idStart - $nameAdjust, $versionStart - $idStart).TrimEnd()
            $currentVersion = $line.Substring($versionStart - $nameAdjust, $availableStart - $versionStart).TrimEnd()
            $availableVersion = $line.Substring($availableStart - $nameAdjust).TrimEnd()
            if ($currentVersion -eq 'Unknown' -or $availableVersion -eq 'Unknown') {
                continue
            }
            if ($currentVersion -ne $availableVersion) {
                [void]$updates.Add(@{
                        AppId = $packageIdRow
                        AppName = $packageName
                        CurrentVersion = $currentVersion
                        AvailableVersion = $availableVersion
                        Scope = $scope
                    })
            }
        }
    }
    return , @($updates)
}

function Get-AvailableUpdates {
    [CmdletBinding()]
    param(
        [string]$forAppId = ''
    )

    Write-Log 'Upgrades: list (default, user, machine)' -Tag 'Debug'
    $wingetPath = Get-WingetPath

    $previousOutputEncoding = [Console]::OutputEncoding
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
    try {
        $allUpdates = [System.Collections.ArrayList]::new()

        try {
            $upgradeResult = & $wingetPath upgrade --source winget |
                Where-Object { $_ -notlike ' *' } |
                Out-String
            $parsed = ConvertFrom-WingetUpgradeOutput -rawOutput $upgradeResult -scope $null
            foreach ($entry in $parsed) { [void]$allUpdates.Add($entry) }
            Write-Log "Upgrades default: $($parsed.Count)" -Tag 'Debug'
        }
        catch {
            Write-Log "List error (default): $_" -Tag 'Debug'
        }

        try {
            $upgradeResult = & $wingetPath upgrade --source winget --scope user |
                Where-Object { $_ -notlike ' *' } |
                Out-String
            $parsed = ConvertFrom-WingetUpgradeOutput -rawOutput $upgradeResult -scope 'user'
            foreach ($entry in $parsed) { [void]$allUpdates.Add($entry) }
            Write-Log "Upgrades user: $($parsed.Count)" -Tag 'Debug'
        }
        catch {
            Write-Log "List error (user): $_" -Tag 'Debug'
        }

        try {
            $upgradeResult = & $wingetPath upgrade --source winget --scope machine |
                Where-Object { $_ -notlike ' *' } |
                Out-String
            $parsed = ConvertFrom-WingetUpgradeOutput -rawOutput $upgradeResult -scope 'machine'
            foreach ($entry in $parsed) { [void]$allUpdates.Add($entry) }
            Write-Log "Upgrades machine: $($parsed.Count)" -Tag 'Debug'
        }
        catch {
            Write-Log "List error (machine): $_" -Tag 'Debug'
        }

        $seen = @{}
        $updates = [System.Collections.ArrayList]::new()
        foreach ($item in $allUpdates) {
            if (-not $seen.ContainsKey($item.AppId)) {
                $seen[$item.AppId] = $true
                [void]$updates.Add($item)
            }
        }
        Write-Log "Upgrade list (merged): $($updates.Count)" -Tag 'Debug'
        if (-not [string]::IsNullOrWhiteSpace($forAppId)) {
            $hit = @($updates | Where-Object { $_.AppId -eq $forAppId }) | Select-Object -First 1
            if ($hit) {
                Write-Log "$($forAppId): $($hit.CurrentVersion) -> $($hit.AvailableVersion)" -Tag 'Info'
            }
            else {
                Write-Log "$($forAppId): none" -Tag 'Info'
            }
        }
        return , @($updates)
    }
    catch {
        Write-Log "Get updates: $_" -Tag 'Error'
        Write-Log "$($_.ScriptStackTrace)" -Tag 'Debug'
        return , @()
    }
    finally {
        [Console]::OutputEncoding = $previousOutputEncoding
    }
}

Write-Log '==================== Start ====================' -Tag 'Start'
Write-Log "Host $env:COMPUTERNAME | $env:USERNAME | $scriptName" -Tag 'Info'

try {
    $wingetExe = $null
    try {
        $wingetExe = Get-WingetPath
        Write-Log 'WinGet path resolved' -Tag 'Get'
    }
    catch {
        if ($_.Exception.Message -notlike 'Winget not found*') {
            Write-Log "Winget resolve: $_" -Tag 'Error'
        }
        else {
            Write-Log 'Winget: no DesktopAppInstaller / winget.exe' -Tag 'Error'
        }
        Complete-Script -exitCode 1
    }

    $wingetVersionCheck = Test-WingetVersion -wingetPath $wingetExe
    if (-not $wingetVersionCheck.isHealthy) {
        Write-Log 'Winget unavailable.' -Tag 'Error'
        Complete-Script -exitCode 0
    }

    Write-Log "List installed: $appId" -Tag 'Run'
    $previousEncoding = [Console]::OutputEncoding
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
    try {
        $null = & $wingetExe list -e --id $appId --accept-source-agreements
        $listExit = $LASTEXITCODE
    }
    finally {
        [Console]::OutputEncoding = $previousEncoding
    }

    Write-Log "list exit: $listExit" -Tag 'Debug'

    if ($listExit -eq -1978335212) {
        Write-Log 'Not installed' -Tag 'Info'
        Complete-Script -exitCode 1
    }

    if ($listExit -ne 0) {
        Write-Log "List failed: exit $listExit" -Tag 'Error'
        Complete-Script -exitCode 1
    }

    Write-Log "Installed: $appId" -Tag 'Success'

    $availableUpdates = Get-AvailableUpdates -forAppId $appId
    $pendingForThisApp = @($availableUpdates | Where-Object { $_.AppId -eq $appId })
    if ($pendingForThisApp.Count -gt 0) {
        $row = $pendingForThisApp | Select-Object -First 1
        Write-Log "Detect: non-compliant ($appId $($row.CurrentVersion) -> $($row.AvailableVersion))" -Tag 'Success'
        Complete-Script -exitCode 1
    }

    Write-Log 'No upgrades' -Tag 'Success'
    Complete-Script -exitCode 0
}
catch {
    Write-Log "Unhandled: $_" -Tag 'Error'
    Write-Log "$($_.ScriptStackTrace)" -Tag 'Debug'
    Complete-Script -exitCode 1
}
