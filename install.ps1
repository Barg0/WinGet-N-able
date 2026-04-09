#requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$appId,

    [Parameter()]
    [string]$override = ''
)

# ---------------------------[ Config ]---------------------------
# Single package id per deployment (-appId); no catalog blacklist/whitelist.

$logDataRoot = "$env:ProgramData\WinGetNable"

$wingetLocaleWorkaround = 'en-US'
$wingetInProgressMaxRetries = 15
$wingetInProgressWaitSeconds = 120
$wingetDownloadRetryWaitSeconds = 30
$wingetUseInstallVersionFallback = $false
$wingetUseUninstallPrevious = $false
$wingetScopeLadderOrder = @('machine', 'none', 'user')

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
$scriptName = "install | $appId"
$logFileName = 'install.log'

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

    Write-Log 'WinGet executable not found under Program Files WindowsApps or user WindowsApps.' -Tag 'Error'
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

function Format-WingetExitCodeHex {
    [CmdletBinding()]
    param([int]$hresultValue)
    $u = [System.BitConverter]::ToUInt32([System.BitConverter]::GetBytes([int32]$hresultValue), 0)
    return ('0x{0:X8}' -f $u)
}

function Get-WingetExitCodeInfo {
    [CmdletBinding()]
    param(
        [int]$exitCode,
        [switch]$explicitWingetSourceOnCommand
    )
    if ($exitCode -eq -1978335138) {
        if ($explicitWingetSourceOnCommand) {
            return @{ Category = 'RetrySourceRepair'; Description = 'Pinned certificate mismatch (explicit winget source; repair source index)' }
        }
        return @{ Category = 'RetrySource'; Description = 'Pinned certificate mismatch' }
    }
    $codeMap = @{
        0              = @{ Category = 'Success'; Description = 'Success' }
        3010           = @{ Category = 'Success'; Description = 'Success (reboot required to complete)' }
        -1978335135    = @{ Category = 'Success'; Description = 'Package already installed' }
        -1978334965    = @{ Category = 'Success'; Description = 'Reboot initiated to finish installation' }
        -1978334963    = @{ Category = 'Success'; Description = 'Another version already installed' }
        -1978334962    = @{ Category = 'Success'; Description = 'Higher version already installed (downgrade)' }
        -1978335216    = @{ Category = 'RetryScope'; Description = 'No applicable installer for current scope' }
        -1978335212    = @{ Category = 'RetryScope'; Description = 'No packages found' }
        -1978335226    = @{ Category = 'RetryScope'; Description = 'ShellExecute install failed (try other scopes)' }
        -1978335222    = @{ Category = 'RetrySourceRepair'; Description = 'Index is corrupt' }
        -1978335221    = @{ Category = 'RetrySourceRepair'; Description = 'Configured source information is corrupt' }
        -1978335217    = @{ Category = 'RetrySourceRepair'; Description = 'Source data missing' }
        -1978335169    = @{ Category = 'RetrySourceRepair'; Description = 'Source data corrupted or tampered' }
        -1978335163    = @{ Category = 'RetrySourceRepair'; Description = 'Failed to open source' }
        -1978335157    = @{ Category = 'RetrySourceRepair'; Description = 'Failed to open one or more sources' }
        -1978335215    = @{ Category = 'RetryHashRefresh'; Description = 'Installer hash does not match manifest' }
        -1978335224    = @{ Category = 'RetryDownload'; Description = 'Download failed' }
        -1978335186    = @{ Category = 'RetryDownload'; Description = 'Download size mismatch' }
        -1978335098    = @{ Category = 'RetryDownload'; Description = 'Downloaded zero-byte installer' }
        -1978335227    = @{ Category = 'RetryLater'; Description = 'Cancellation signal received' }
        -1978335126    = @{ Category = 'RetryLater'; Description = 'Application shutdown signal received' }
        -1978335125    = @{ Category = 'RetryLater'; Description = 'Failed to download dependencies' }
        -1978335123    = @{ Category = 'RetryLater'; Description = 'Service busy or unavailable' }
        -1978334975    = @{ Category = 'RetryLater'; Description = 'Application is currently running' }
        -1978334974    = @{ Category = 'RetryLater'; Description = 'Another installation in progress' }
        -1978334973    = @{ Category = 'RetryLater'; Description = 'One or more file is in use' }
        -1978334971    = @{ Category = 'RetryLater'; Description = 'Disk full' }
        -1978334970    = @{ Category = 'RetryLater'; Description = 'Insufficient memory' }
        -1978334969    = @{ Category = 'RetryLater'; Description = 'No network connectivity' }
        -1978334967    = @{ Category = 'RetryLater'; Description = 'Reboot required to finish installation' }
        -1978334966    = @{ Category = 'RetryLater'; Description = 'Reboot required then try again' }
        -1978334959    = @{ Category = 'RetryLater'; Description = 'Application in use by another application' }
        -1978335231    = @{ Category = 'Fail'; Description = 'Internal error' }
        -1978335230    = @{ Category = 'Fail'; Description = 'Invalid command line arguments' }
        -1978335229    = @{ Category = 'Fail'; Description = 'Command failed' }
        -1978335228    = @{ Category = 'Fail'; Description = 'Opening manifest failed' }
        -1978335225    = @{ Category = 'Fail'; Description = 'Manifest version higher than supported; update winget' }
        -1978335210    = @{ Category = 'Fail'; Description = 'Multiple packages found matching criteria' }
        -1978335209    = @{ Category = 'Fail'; Description = 'No manifest found matching criteria' }
        -1978335207    = @{ Category = 'Fail'; Description = 'Command requires administrator privileges' }
        -1978335205    = @{ Category = 'Fail'; Description = 'Microsoft Store client blocked by policy' }
        -1978335204    = @{ Category = 'Fail'; Description = 'Microsoft Store app blocked by policy' }
        -1978335189    = @{ Category = 'RetryScope'; Description = 'No applicable upgrade (does not apply to system or scope)' }
        -1978335188    = @{ Category = 'Fail'; Description = 'upgrade --all completed with failures' }
        -1978335187    = @{ Category = 'Fail'; Description = 'Installer failed security check' }
        -1978335174    = @{ Category = 'Fail'; Description = 'Blocked by Group Policy' }
        -1978335159    = @{ Category = 'Fail'; Description = 'MSI install failed' }
        -1978335153    = @{ Category = 'Fail'; Description = 'Upgrade version not newer than installed' }
        -1978335152    = @{ Category = 'Fail'; Description = 'Upgrade version unknown; override not specified' }
        -1978335146    = @{ Category = 'Fail'; Description = 'Installer prohibits elevation' }
        -1978335128    = @{ Category = 'Fail'; Description = 'Package has a pin that prevents upgrade' }
        -1978335127    = @{ Category = 'Fail'; Description = 'Package is a stub; full package needed' }
        -1978335090    = @{ Category = 'Fail'; Description = 'Install technology mismatch (different installer type)' }
        -1978334972    = @{ Category = 'Fail'; Description = 'Missing dependency on system' }
        -1978334968    = @{ Category = 'Fail'; Description = 'Installation error; contact support' }
        -1978334964    = @{ Category = 'Fail'; Description = 'Installation cancelled by user' }
        -1978334961    = @{ Category = 'Fail'; Description = 'Blocked by organization policy' }
        -1978334960    = @{ Category = 'Fail'; Description = 'Failed to install package dependencies' }
        -1978334958    = @{ Category = 'Fail'; Description = 'Invalid parameter' }
        -1978334957    = @{ Category = 'Fail'; Description = 'Package not supported by system' }
        -1978334956    = @{ Category = 'Fail'; Description = 'Installer does not support upgrading existing package' }
        -1978334955    = @{ Category = 'Fail'; Description = 'Installer custom error' }
        -2147023673    = @{ Category = 'RetryLater'; Description = 'Operation cancelled - ERROR_CANCELLED (0x800704C7)' }
        -2147012894    = @{ Category = 'RetryLater'; Description = 'Connection timed out - ERROR_INTERNET_TIMEOUT (0x80072EE2)' }
        -2147012889    = @{ Category = 'RetryLater'; Description = 'DNS name not resolved - ERROR_INTERNET_NAME_NOT_RESOLVED (0x80072EE7)' }
        -2147012867    = @{ Category = 'RetryLater'; Description = 'Cannot connect to server - ERROR_INTERNET_CANNOT_CONNECT (0x80072EFD)' }
        -2147012866    = @{ Category = 'RetryLater'; Description = 'Connection aborted - ERROR_INTERNET_CONNECTION_ABORTED (0x80072EFE)' }
        -2147012465    = @{ Category = 'RetryLater'; Description = 'TLS/SSL error - ERROR_INTERNET_DECRYPTION_FAILED (0x80072F8F)' }
        -2147221003    = @{ Category = 'Fail'; Description = 'Application/uninstaller not found - orphaned ARP entry (0x800401F5)' }
        -2147024891    = @{ Category = 'Fail'; Description = 'Access denied - ERROR_ACCESS_DENIED (0x80070005)' }
        -2147023293    = @{ Category = 'Fail'; Description = 'MSI fatal error - ERROR_INSTALL_FAILURE (0x80070643 / 1603)' }
        -2147023286    = @{ Category = 'RetryLater'; Description = 'Windows Installer busy - ERROR_INSTALL_ALREADY_RUNNING (0x8007064A / 1610)' }
    }
    if ($codeMap.ContainsKey($exitCode)) {
        return $codeMap[$exitCode]
    }
    $hex = Format-WingetExitCodeHex -hresultValue $exitCode
    return @{ Category = 'Unknown'; Description = "Unmapped exit code $exitCode ($hex)" }
}

function Join-ArgumentsForProcess {
    [CmdletBinding()]
    param([string[]]$argumentList)
    $escaped = foreach ($arg in $argumentList) {
        $needsQuoting = $arg.Length -eq 0 -or $arg -match '[\s"]'
        if (-not $needsQuoting) {
            $arg
        }
        else {
            $sb = [System.Text.StringBuilder]::new()
            [void]$sb.Append('"')
            $i = 0
            while ($i -lt $arg.Length) {
                $c = $arg[$i]
                $i++
                if ($c -eq [char]0x5C) {
                    $n = 1
                    while ($i -lt $arg.Length -and $arg[$i] -eq [char]0x5C) { $i++; $n++ }
                    if ($i -eq $arg.Length) {
                        [void]$sb.Append([char]0x5C, $n * 2)
                    }
                    elseif ($arg[$i] -eq '"') {
                        [void]$sb.Append([char]0x5C, $n * 2 + 1)
                        [void]$sb.Append('"')
                        $i++
                    }
                    else {
                        [void]$sb.Append([char]0x5C, $n)
                    }
                    continue
                }
                if ($c -eq '"') {
                    [void]$sb.Append([char]0x5C)
                    [void]$sb.Append('"')
                }
                else {
                    [void]$sb.Append($c)
                }
            }
            [void]$sb.Append('"')
            $sb.ToString()
        }
    }
    $escaped -join ' '
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

function Get-WingetScopeLadderOrderNormalized {
    [CmdletBinding()]
    param()
    $rawList = @($wingetScopeLadderOrder)
    if ($rawList.Count -eq 0) {
        throw 'wingetScopeLadderOrder must be a non-empty array.'
    }
    $tokens = [System.Collections.Generic.List[string]]::new()
    foreach ($raw in $rawList) {
        $t = [string]$raw
        if ([string]::IsNullOrWhiteSpace($t)) { continue }
        foreach ($commaPart in ($t -split ',')) {
            $cp = $commaPart.Trim()
            if ([string]::IsNullOrWhiteSpace($cp)) { continue }
            foreach ($word in ($cp -split '\s+')) {
                $w = $word.Trim()
                if (-not [string]::IsNullOrWhiteSpace($w)) {
                    [void]$tokens.Add($w)
                }
            }
        }
    }
    if ($tokens.Count -eq 0) {
        throw 'wingetScopeLadderOrder resolved to no tokens.'
    }
    $normalized = [System.Collections.Generic.List[string]]::new()
    $seen = @{}
    foreach ($key in $tokens) {
        $mode = switch -Regex ($key.ToLowerInvariant()) {
            '^(machine|system)$' { 'Machine'; break }
            '^(default|none)$' { 'Default'; break }
            '^(user)$' { 'User'; break }
            default { throw "Invalid wingetScopeLadderOrder token: '$key'." }
        }
        if ($seen.ContainsKey($mode)) {
            throw "Duplicate scope in wingetScopeLadderOrder: $mode"
        }
        $seen[$mode] = $true
        [void]$normalized.Add($mode)
    }
    return $normalized.ToArray()
}

function Get-WingetScopeSuccessSuffix {
    param(
        [ValidateSet('Machine', 'Default', 'User')]
        [string]$scopeMode
    )
    switch ($scopeMode) {
        'Machine' { return 'machine' }
        'Default' { return 'default scope' }
        'User' { return 'user' }
    }
}

function Get-WingetScopeUpgradeRetryLog {
    param(
        [ValidateSet('Machine', 'Default', 'User')]
        [string]$scopeMode
    )
    switch ($scopeMode) {
        'Machine' { return 'Retry: --scope machine' }
        'Default' { return 'Retry: no --scope' }
        'User' { return 'Retry: --scope user' }
    }
}

function Get-WingetScopeShortName {
    param(
        [ValidateSet('Machine', 'Default', 'User')]
        [string]$scopeMode
    )
    switch ($scopeMode) {
        'Machine' { return 'machine' }
        'Default' { return 'default' }
        'User' { return 'user' }
    }
}

function Test-WingetUpgradeOutputClaimsNoApplicable {
    [CmdletBinding()]
    param(
        [AllowNull()]
        $lines
    )
    if ($null -eq $lines) { return $false }
    $t = ($lines | Out-String)
    return $t -match '(?i)No applicable upgrade|does not apply to your system or requirements'
}

function Update-WinGetPackage {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$packageId,
        [Parameter(Mandatory)]
        [string]$wingetPath,
        [string]$availableVersion = ''
    )

    function Invoke-UpgradeStep {
        param(
            [ValidateSet('Machine', 'Default', 'User')]
            [string]$scopeMode,
            [string]$locale
        )
        $wingetArgs = @('upgrade', '--id', $packageId, '-e', '--force', '--accept-package-agreements', '--accept-source-agreements',
            '--silent', '--disable-interactivity', '--skip-dependencies', '--source', 'winget')
        if (-not [string]::IsNullOrWhiteSpace($locale)) { $wingetArgs += '--locale', $locale.Trim() }
        if ($scopeMode -eq 'Machine') { $wingetArgs += '--scope', 'machine' }
        elseif ($scopeMode -eq 'User') { $wingetArgs += '--scope', 'user' }
        Write-Log "winget $($wingetArgs -join ' ')" -Tag 'Debug'
        & $wingetPath @wingetArgs 2>&1 | Where-Object { $_ -notlike ' *' }
    }

    function Invoke-UpgradeWithInProgressWait {
        param(
            [ValidateSet('Machine', 'Default', 'User')]
            [string]$scopeMode,
            [string]$locale
        )
        $inProgressCount = 0
        $upgradeOutput = @()
        $exitCodeInner = 0
        do {
            if ($inProgressCount -gt 0) {
                Write-Log "Install busy; wait ${wingetInProgressWaitSeconds}s ($inProgressCount/$wingetInProgressMaxRetries)" -Tag 'Info'
                Start-Sleep -Seconds $wingetInProgressWaitSeconds
            }
            $upgradeOutput = Invoke-UpgradeStep -scopeMode $scopeMode -locale $locale
            $exitCodeInner = $LASTEXITCODE
            if ($null -eq $upgradeOutput) { $upgradeOutput = @() }
            if ($exitCodeInner -ne -1978334974) { break }
            $inProgressCount++
        } while ($inProgressCount -le $wingetInProgressMaxRetries)
        return @{ outputLines = $upgradeOutput; exitCode = $exitCodeInner }
    }

    function Invoke-UpgradeAttemptResolved {
        param(
            [ValidateSet('Machine', 'Default', 'User')]
            [string]$scopeMode,
            [string]$locale
        )
        $result = Invoke-UpgradeWithInProgressWait -scopeMode $scopeMode -locale $locale
        $exitInfoInner = Get-WingetExitCodeInfo -exitCode $result.exitCode -ExplicitWingetSourceOnCommand

        if ($exitInfoInner.Category -eq 'RetryHashRefresh') {
            Write-Log "Hash mismatch ${packageId}: refreshing winget source index" -Tag 'Info'
            & $wingetPath source update --name winget 2>&1 | Out-Null
            $result = Invoke-UpgradeWithInProgressWait -scopeMode $scopeMode -locale $locale
        }
        elseif ($exitInfoInner.Category -eq 'RetryDownload') {
            Write-Log "Download failed ${packageId}: retry in ${wingetDownloadRetryWaitSeconds}s" -Tag 'Info'
            Start-Sleep -Seconds $wingetDownloadRetryWaitSeconds
            $result = Invoke-UpgradeWithInProgressWait -scopeMode $scopeMode -locale $locale
        }
        return $result
    }

    function Test-ShouldAdvanceScopeLadder {
        param(
            $exitInfoRow,
            [int]$code,
            [bool]$outputClaimsNoApplicable
        )
        return ($exitInfoRow.Category -eq 'RetryScope') -or $outputClaimsNoApplicable
    }

    try {
        $scopeOrder = @($script:wingetScopeLadderNormalized)
        if ($null -eq $scopeOrder -or $scopeOrder.Count -eq 0) {
            throw 'wingetScopeLadderNormalized is not set.'
        }

        $localePassMax = if ([string]::IsNullOrWhiteSpace($wingetLocaleWorkaround)) { 1 } else { 2 }
        $sourceRepairDone = $false
        $exitCode = 0
        $exitInfo = $null
        $upgradeOutput = @()
        $outputClaimsNoApplicable = $false

        for ($localePass = 0; $localePass -lt $localePassMax; $localePass++) {
            $localeArg = ''
            if ($localePass -eq 1) {
                $localeArg = $wingetLocaleWorkaround.Trim()
                Write-Log "Retry: --locale $localeArg" -Tag 'Info'
            }

            $deferCategories = @('RetryLater', 'RetryHashRefresh', 'RetryDownload')

            for ($ladderRun = 0; $ladderRun -lt 2; $ladderRun++) {
                if ($ladderRun -eq 1) {
                    Write-Log 'Source repair: winget source reset + update' -Tag 'Info'
                    & $wingetPath source reset --force 2>&1 | Out-Null
                    & $wingetPath source update 2>&1 | Out-Null
                    Write-Log 'Source repaired; retrying upgrade' -Tag 'Info'
                }

                $runNotes = @()
                if ($localePass -eq 1) { $runNotes += 'locale' }
                if ($ladderRun -eq 1) { $runNotes += 'source repair' }
                $successNote = if ($runNotes.Count -gt 0) { " ($($runNotes -join ', '))" } else { '' }

                for ($scopeIndex = 0; $scopeIndex -lt $scopeOrder.Count; $scopeIndex++) {
                    $scopeMode = $scopeOrder[$scopeIndex]

                    if ($scopeIndex -gt 0) {
                        $tryNextScope = (Test-ShouldAdvanceScopeLadder -exitInfoRow $exitInfo -code $exitCode -outputClaimsNoApplicable $outputClaimsNoApplicable) -or ($exitCode -eq -1978335212)
                        if (-not $tryNextScope) { break }
                        Write-Log (Get-WingetScopeUpgradeRetryLog -scopeMode $scopeMode) -Tag 'Info'
                    }

                    $attempt = Invoke-UpgradeAttemptResolved -scopeMode $scopeMode -locale $localeArg
                    $upgradeOutput = $attempt.outputLines
                    $exitCode = $attempt.exitCode

                    if ($exitCode -eq -1978334974) {
                        Write-Log "Defer ${packageId}: install busy (max waits)" -Tag 'Info'
                        return $null
                    }

                    $exitInfo = Get-WingetExitCodeInfo -exitCode $exitCode -ExplicitWingetSourceOnCommand
                    $outputClaimsNoApplicable = Test-WingetUpgradeOutputClaimsNoApplicable -lines $upgradeOutput
                    $treatAsSuccess = ($exitInfo.Category -eq 'Success') -and -not $outputClaimsNoApplicable

                    if ($treatAsSuccess) {
                        $suffix = Get-WingetScopeSuccessSuffix -scopeMode $scopeMode
                        Write-Log "$packageId upgraded ($suffix)$successNote" -Tag 'Success'
                        return $true
                    }

                    if ($exitInfo.Category -in $deferCategories) {
                        Write-Log "Defer ${packageId}: $($exitInfo.Description)" -Tag 'Info'
                        return $null
                    }

                    if ($scopeIndex -eq ($scopeOrder.Count - 1)) {
                        $short = Get-WingetScopeShortName -scopeMode $scopeMode
                        Write-Log "$short scope: exit $exitCode $($exitInfo.Description)" -Tag 'Debug'
                    }
                }

                if ($ladderRun -eq 0 -and -not $sourceRepairDone -and $exitInfo.Category -eq 'RetrySourceRepair') {
                    $sourceRepairDone = $true
                    continue
                }
                break
            }

            if ($localePass -eq 0 -and $exitCode -eq -1978335212 -and -not [string]::IsNullOrWhiteSpace($wingetLocaleWorkaround)) {
                continue
            }

            if ($wingetUseInstallVersionFallback -and $exitCode -eq -1978335212 -and -not [string]::IsNullOrWhiteSpace($availableVersion)) {
                Write-Log "Retry: install fallback (winget install --version $availableVersion)" -Tag 'Info'
                $installScopeIx = 0
                foreach ($scopeMode in $scopeOrder) {
                    if ($installScopeIx -gt 0) {
                        $sn = Get-WingetScopeShortName -scopeMode $scopeMode
                        Write-Log "Retry: ${sn} scope (install)" -Tag 'Info'
                    }
                    $installScopeIx++
                    $installArgs = @('install', '--id', $packageId, '-e', '--version', $availableVersion, '--force',
                        '--accept-package-agreements', '--accept-source-agreements', '--silent', '--disable-interactivity', '--skip-dependencies', '--source', 'winget')
                    if ($scopeMode -eq 'Machine') { $installArgs += '--scope', 'machine' }
                    elseif ($scopeMode -eq 'User') { $installArgs += '--scope', 'user' }
                    Write-Log "winget $($installArgs -join ' ')" -Tag 'Debug'
                    $upgradeOutput = & $wingetPath @installArgs 2>&1 | Where-Object { $_ -notlike ' *' }
                    $exitCode = $LASTEXITCODE
                    $exitInfo = Get-WingetExitCodeInfo -exitCode $exitCode -ExplicitWingetSourceOnCommand
                    $outputClaimsNoApplicable = Test-WingetUpgradeOutputClaimsNoApplicable -lines $upgradeOutput
                    if (($exitInfo.Category -eq 'Success') -and -not $outputClaimsNoApplicable) {
                        $isn = Get-WingetScopeShortName -scopeMode $scopeMode
                        Write-Log "$packageId ($isn, install fallback)" -Tag 'Success'
                        return $true
                    }
                    if ($exitInfo.Category -in $deferCategories) {
                        Write-Log "Defer ${packageId}: $($exitInfo.Description)" -Tag 'Info'
                        return $null
                    }
                    if (-not (Test-ShouldAdvanceScopeLadder -exitInfoRow $exitInfo -code $exitCode -outputClaimsNoApplicable $outputClaimsNoApplicable) -and $exitCode -ne -1978335212) {
                        break
                    }
                }
            }

            if ($wingetUseUninstallPrevious -and -not [string]::IsNullOrWhiteSpace($availableVersion)) {
                Write-Log 'Retry: uninstall-previous' -Tag 'Info'
                $uninstScopeIx = 0
                foreach ($scopeMode in $scopeOrder) {
                    if ($uninstScopeIx -gt 0) {
                        $usn = Get-WingetScopeShortName -scopeMode $scopeMode
                        Write-Log "Retry: ${usn} scope (uninstall-previous)" -Tag 'Info'
                    }
                    $uninstScopeIx++
                    $uninstPrevArgs = @('upgrade', '--id', $packageId, '-e', '--version', $availableVersion, '--force',
                        '--uninstall-previous', '--accept-package-agreements', '--accept-source-agreements', '--silent', '--disable-interactivity', '--skip-dependencies', '--source', 'winget')
                    if ($scopeMode -eq 'Machine') { $uninstPrevArgs += '--scope', 'machine' }
                    elseif ($scopeMode -eq 'User') { $uninstPrevArgs += '--scope', 'user' }
                    Write-Log "winget $($uninstPrevArgs -join ' ')" -Tag 'Debug'
                    $upgradeOutput = & $wingetPath @uninstPrevArgs 2>&1 | Where-Object { $_ -notlike ' *' }
                    $exitCode = $LASTEXITCODE
                    $exitInfo = Get-WingetExitCodeInfo -exitCode $exitCode -ExplicitWingetSourceOnCommand
                    $outputClaimsNoApplicable = Test-WingetUpgradeOutputClaimsNoApplicable -lines $upgradeOutput
                    if (($exitInfo.Category -eq 'Success') -and -not $outputClaimsNoApplicable) {
                        $usn2 = Get-WingetScopeShortName -scopeMode $scopeMode
                        Write-Log "$packageId ($usn2, uninstall-previous)" -Tag 'Success'
                        return $true
                    }
                    if ($exitInfo.Category -in $deferCategories) {
                        Write-Log "Defer ${packageId}: $($exitInfo.Description)" -Tag 'Info'
                        return $null
                    }
                    if (-not (Test-ShouldAdvanceScopeLadder -exitInfoRow $exitInfo -code $exitCode -outputClaimsNoApplicable $outputClaimsNoApplicable)) {
                        break
                    }
                }
            }

            $errorMessages = $upgradeOutput | Where-Object {
                $_ -match 'error|failed|exception|unable|cannot|could not' -or
                ($_ -match '^[A-Z]' -and $_ -notmatch '^Loading|^Found|^Verifying|^Successfully')
            }
            if ($errorMessages) {
                Write-Log "WinGet output for ${packageId}: $($errorMessages -join '; ')" -Tag 'Debug'
            }
            Write-Log "Upgrade failed for ${packageId}: $($exitInfo.Description) ($($exitInfo.Category))" -Tag 'Error'
            return $false
        }
    }
    catch {
        Write-Log "Upgrade ${packageId} error: $_" -Tag 'Error'
        Write-Log "$($_.ScriptStackTrace)" -Tag 'Debug'
        return $false
    }
}

function Invoke-WinGetPackageInstall {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$wingetExe,
        [Parameter(Mandatory)][string]$packageId,
        [string]$installOverride = ''
    )

    $useScope = $true
    $useSource = $false
    $triedNoScope = $false
    $installRepairDone = $false

    if (-not [string]::IsNullOrWhiteSpace($installOverride)) {
        Write-Log "Using installer override string for winget --override (length $($installOverride.Length))." -Tag 'Info'
    }

    while ($true) {
        $currentArgs = @('install', '-e', '--id', $packageId, '--silent', '--skip-dependencies',
            '--accept-package-agreements', '--accept-source-agreements', '--force')
        if ($useScope) { $currentArgs += '--scope', 'machine' }
        if ($useSource) { $currentArgs += '--source', 'winget' }
        if (-not [string]::IsNullOrWhiteSpace($installOverride)) {
            $currentArgs += '--override', $installOverride
        }

        $scopeLabel = if ($useScope) { 'scope machine' } else { 'no scope' }
        $sourceLabel = if ($useSource) { ', source winget' } else { '' }

        $inProgressCount = 0
        $exitCode = 0
        do {
            if ($inProgressCount -gt 0) {
                Write-Log "Install busy; wait ${wingetInProgressWaitSeconds}s ($inProgressCount/$wingetInProgressMaxRetries)" -Tag 'Info'
                Start-Sleep -Seconds $wingetInProgressWaitSeconds
            }
            $runLabel = "Installing ($scopeLabel$sourceLabel)"
            if ($inProgressCount -gt 0) { $runLabel += " [busy retry $inProgressCount/$wingetInProgressMaxRetries]" }
            Write-Log "$runLabel." -Tag 'Run'
            Write-Log "winget $($currentArgs -join ' ')" -Tag 'Debug'

            $psi = [System.Diagnostics.ProcessStartInfo]::new()
            $psi.FileName = $wingetExe
            $psi.UseShellExecute = $false
            $psi.CreateNoWindow = $true
            if ($psi.PSObject.Properties['ArgumentList']) {
                foreach ($arg in $currentArgs) { [void]$psi.ArgumentList.Add($arg) }
            }
            else {
                $psi.Arguments = Join-ArgumentsForProcess -argumentList $currentArgs
            }
            $proc = [System.Diagnostics.Process]::Start($psi)
            $proc.WaitForExit()
            $exitCode = $proc.ExitCode
            $exitInfo = Get-WingetExitCodeInfo -exitCode $exitCode -ExplicitWingetSourceOnCommand:$useSource
            Write-Log ("Exit {0}: {1} ({2})" -f $exitCode, $exitInfo.Description, $exitInfo.Category) -Tag 'Info'

            if ($exitCode -ne -1978334974) { break }
            $inProgressCount++
        } while ($inProgressCount -le $wingetInProgressMaxRetries)

        if ($exitCode -eq -1978334974) {
            Write-Log 'Defer: install busy (max waits)' -Tag 'Error'
            return $null
        }

        if ($exitCode -eq 0 -or $exitInfo.Category -eq 'Success') {
            Write-Log "Installed $packageId ($scopeLabel$sourceLabel)" -Tag 'Success'
            return $true
        }

        if ($exitInfo.Category -eq 'RetryLater') {
            Write-Log "Install deferred (transient): $($exitInfo.Description)." -Tag 'Info'
            return $null
        }

        if ($exitInfo.Category -eq 'RetryHashRefresh') {
            Write-Log 'Hash: source update (winget)' -Tag 'Info'
            & $wingetExe source update --name winget 2>&1 | Out-Null
            continue
        }

        if ($exitInfo.Category -eq 'RetryDownload') {
            Write-Log "Download: retry in ${wingetDownloadRetryWaitSeconds}s" -Tag 'Info'
            Start-Sleep -Seconds $wingetDownloadRetryWaitSeconds
            continue
        }

        if ($exitInfo.Category -eq 'RetrySourceRepair' -and -not $installRepairDone) {
            Write-Log 'Source repair: winget source reset + update' -Tag 'Info'
            & $wingetExe source reset --force 2>&1 | Out-Null
            & $wingetExe source update 2>&1 | Out-Null
            $installRepairDone = $true
            continue
        }

        $workaroundApplied = $false
        if ($exitInfo.Category -eq 'RetryScope' -and -not $triedNoScope) {
            Write-Log 'Retry: no --scope' -Tag 'Info'
            $useScope = $false
            $triedNoScope = $true
            $workaroundApplied = $true
        }

        if ($exitInfo.Category -eq 'RetrySource' -and -not $useSource) {
            Write-Log 'Retry: --source winget' -Tag 'Info'
            $useSource = $true
            $workaroundApplied = $true
        }

        if (-not $workaroundApplied) {
            Write-Log "Fail install: $($exitInfo.Description) ($($exitInfo.Category))" -Tag 'Error'
            return $false
        }
        Write-Log 'Workaround; retry install' -Tag 'Debug'
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

    Write-Log 'WinGet source update' -Tag 'Run'
    & $wingetExe source update 2>&1 | Out-Null

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

    $isNotInstalled = ($listExit -eq -1978335212)
    if ($isNotInstalled) {
        Write-Log 'Not installed' -Tag 'Info'
        $installOutcome = Invoke-WinGetPackageInstall -wingetExe $wingetExe -packageId $appId -installOverride $override
        if ($installOutcome -eq $false) {
            Complete-Script -exitCode 1
        }
        Complete-Script -exitCode 0
    }

    if ($listExit -ne 0) {
        Write-Log "List failed: exit $listExit" -Tag 'Error'
        Complete-Script -exitCode 1
    }

    Write-Log "Installed: $appId" -Tag 'Success'

    $availableUpdates = Get-AvailableUpdates -forAppId $appId
    $pendingForThisApp = @($availableUpdates | Where-Object { $_.AppId -eq $appId })

    if ($pendingForThisApp.Count -eq 0) {
        Write-Log 'No upgrades' -Tag 'Success'
        Complete-Script -exitCode 0
    }

    $pendingRow = $pendingForThisApp | Select-Object -First 1
    Write-Log 'Upgrading' -Tag 'Run'

    try {
        $normalizedScopes = Get-WingetScopeLadderOrderNormalized
        $script:wingetScopeLadderNormalized = $normalizedScopes
    }
    catch {
        Write-Log "Invalid wingetScopeLadderOrder: $_" -Tag 'Error'
        Complete-Script -exitCode 1
    }

    $upgradeOutcome = Update-WinGetPackage -packageId $appId -wingetPath $wingetExe -availableVersion $pendingRow.AvailableVersion
    if ($upgradeOutcome -eq $false) {
        Complete-Script -exitCode 1
    }
    Complete-Script -exitCode 0
}
catch {
    Write-Log "Unhandled: $_" -Tag 'Error'
    Write-Log "$($_.ScriptStackTrace)" -Tag 'Debug'
    Complete-Script -exitCode 1
}
