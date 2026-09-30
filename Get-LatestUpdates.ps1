#Requires -Version 5.1
<#
.SYNOPSIS
  Finds and downloads the latest cumulative updates from the Microsoft Update Catalog.

.EXAMPLE
  .\Get-LatestUpdates.ps1 -FunctionsPath .\MicrosoftUpdateCatalog.ps1 -DestinationPath D:\PatchStaging
  .\Get-LatestUpdates.ps1 -FunctionsPath .\MicrosoftUpdateCatalog.ps1 -DestinationPath D:\PatchStaging -Targets Win11-24H2,Edge -ListOnly

.NOTES
  Must run in Windows PowerShell 5.1 (the catalog functions use the HTMLFile COM object,
  which does not exist in PowerShell 7).
#>
param(
    [Parameter(Mandatory)][string]$FunctionsPath,     # the .ps1 file containing Get-MicrosoftUpdates etc.
    [Parameter(Mandatory)][string]$DestinationPath,
    [string[]]$Targets,                               # default = all
    [switch]$ListOnly                                 # show what would be downloaded, download nothing
)

. $FunctionsPath
$ProgressPreference = 'SilentlyContinue'              # Invoke-WebRequest is MUCH faster without the progress bar
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# ---- Catalog targets --------------------------------------------------------
# Search = words sent to the catalog (AND-matched). Include/Exclude = regex applied to the Title.
$CatalogTargets = [ordered]@{
    'Win11-24H2' = @{
        Search  = 'Cumulative Update Windows 11 24H2 x64'
        Include = 'Cumulative Update for Windows 11,? version 24H2 for x64-based Systems'
        Exclude = 'Preview|\.NET|Dynamic|Hotpatch|Safe OS'
    }
    'Server2019' = @{
        Search  = 'Cumulative Update Windows Server 2019 x64'
        Include = 'Cumulative Update for Windows Server 2019 \(1809\) for x64-based Systems'
        Exclude = 'Preview|\.NET|Dynamic|Safe OS'
    }
    'Server2022' = @{
        Search  = 'Cumulative Update Microsoft server operating system 21H2 x64'
        Include = 'Cumulative Update for Microsoft server operating system,? version 21H2 for x64-based Systems'
        Exclude = 'Preview|\.NET|Dynamic|Safe OS'
    }
}
# Edge is handled separately below (Get-LatestEdge).
$AllTargets = @($CatalogTargets.Keys) + 'Edge'
if (-not $Targets) { $Targets = $AllTargets }

# ---- Helpers ----------------------------------------------------------------
function Test-CatalogDigest {
    param($Path, $Digest, $Algorithm = 'SHA1')
    if (-not $Digest) { return $true }                # nothing to compare against
    $alg = [Security.Cryptography.HashAlgorithm]::Create($Algorithm)
    $fs  = [IO.File]::OpenRead($Path)
    try { $actual = [Convert]::ToBase64String($alg.ComputeHash($fs)) } finally { $fs.Dispose() }
    return $actual -eq $Digest
}

function Save-File {
    param($Url, $OutFile)
    Invoke-WebRequest -UseBasicParsing -Uri $Url -OutFile $OutFile
}

function Get-LatestCatalogUpdate {
    param($Target)
    $rows = @(Get-MicrosoftUpdates -SearchText $Target.Search)
    $rows |
        Where-Object { $_.Title -match $Target.Include -and $_.Title -notmatch $Target.Exclude } |
        Sort-Object @{ Expression = {
                [datetime]::ParseExact(([string]$_.'Last Updated').Trim(), 'M/d/yyyy',
                    [Globalization.CultureInfo]::InvariantCulture) }; Descending = $true },
                    @{ Expression = { $_.Title }; Descending = $true } |
        Select-Object -First 1
}

# Edge isn't well served by the catalog; Microsoft publishes a JSON feed with hashes.
function Get-LatestEdge {
    param([string]$Channel = 'Stable', [string]$Arch = 'x64')
    $products = Invoke-RestMethod -UseBasicParsing -Uri 'https://edgeupdates.microsoft.com/api/products'
    $rel = ($products | Where-Object Product -eq $Channel).Releases |
        Where-Object { $_.Platform -eq 'Windows' -and $_.Architecture -eq $Arch -and $_.Artifacts } |
        Sort-Object { [version]$_.ProductVersion } -Descending | Select-Object -First 1
    $art = $rel.Artifacts | Where-Object ArtifactName -eq 'msi' | Select-Object -First 1
    [pscustomobject]@{
        Version = $rel.ProductVersion
        Url     = $art.Location
        Sha256  = $art.Hash
        File    = "MicrosoftEdgeEnterprise$($Arch)_$($rel.ProductVersion).msi"
    }
}

# ---- Main -------------------------------------------------------------------
New-Item -ItemType Directory -Force -Path $DestinationPath | Out-Null
$summary = @()

foreach ($name in $Targets) {
    try {
        $outDir = Join-Path $DestinationPath $name
        New-Item -ItemType Directory -Force -Path $outDir | Out-Null

        if ($name -eq 'Edge') {
            $e = Get-LatestEdge
            $path = Join-Path $outDir $e.File
            Write-Host "[$name] Edge $($e.Version)" -ForegroundColor Cyan
            if ($ListOnly) { $summary += [pscustomobject]@{ Target=$name; Title="Edge $($e.Version)"; UpdateID=''; File=$e.File; Status='ListOnly' }; continue }
            if (-not (Test-Path $path)) { Save-File $e.Url $path }
            $ok = (Get-FileHash $path -Algorithm SHA256).Hash -eq $e.Sha256
            $summary += [pscustomobject]@{ Target=$name; Title="Edge $($e.Version)"; UpdateID=''; File=$e.File; Status=$(if($ok){'OK'}else{'HASH MISMATCH'}) }
            continue
        }

        $pick = Get-LatestCatalogUpdate $CatalogTargets[$name]
        if (-not $pick) { Write-Warning "[$name] No matching update on page 1 of results. Check Include/Exclude, or use -NextPage."; continue }
        Write-Host "[$name] $($pick.Title)  ($($pick.Id))" -ForegroundColor Cyan

        $info = Get-MicrosoftUpdateFiles -UpdateID $pick.Id
        foreach ($f in $info.Files | Where-Object { $_.URL }) {
            $path = Join-Path $outDir $f.FileName
            if ($ListOnly) { $summary += [pscustomobject]@{ Target=$name; Title=$pick.Title; UpdateID=$pick.Id; File=$f.FileName; Status='ListOnly' }; continue }

            if (-not (Test-Path $path)) { Save-File $f.URL $path }
            $alg = if ($f.PSObject.Properties['DigestAlgorithm']) { $f.DigestAlgorithm } else { 'SHA1' }
            $ok  = Test-CatalogDigest -Path $path -Digest $f.Digest -Algorithm $alg
            if (-not $ok) { Write-Warning "[$name] Hash mismatch for $($f.FileName)" }
            $summary += [pscustomobject]@{ Target=$name; Title=$pick.Title; UpdateID=$pick.Id; File=$f.FileName; Status=$(if($ok){'OK'}else{'HASH MISMATCH'}) }
        }
    }
    catch { Write-Warning "[$name] $($_.Exception.Message)" }
}

$summary | Format-Table -AutoSize
$summary | Export-Csv -NoTypeInformation -Path (Join-Path $DestinationPath 'last-run.csv')
