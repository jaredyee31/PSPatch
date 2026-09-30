#Requires -Version 5.1
<#
.SYNOPSIS
  GUI to download the latest Windows cumulative updates and third-party installers.
  Run in Windows PowerShell 5.1 (powershell.exe), not PowerShell 7.

.EXAMPLE
  powershell.exe -STA -ExecutionPolicy Bypass -File .\Get-LatestUpdatesGUI.ps1
#>
param(
    [string]$ModulePath = (Join-Path $PSScriptRoot 'PSMicrosoftUpdateCatalog.psm1')
)

Add-Type -AssemblyName System.Windows.Forms, System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# =============================================================================
# TARGET DEFINITIONS - add or edit entries here
# =============================================================================
# Catalog targets: Search = words sent to the catalog; Include/Exclude = regex on the Title.
$CatalogTargets = @(
    @{ Key='Win11-24H2';      Label='Windows 11 24H2 x64 - Cumulative Update'
       Search='Cumulative Update Windows 11 24H2 x64'
       Include='Cumulative Update for Windows 11,? version 24H2 for x64-based Systems'
       Exclude='Preview|\.NET|Dynamic|Hotpatch|Safe OS' }
    @{ Key='Server2019-CU';   Label='Windows Server 2019 - Cumulative Update (.msu)'
       Search='Cumulative Update Windows Server 2019 x64'
       Include='Cumulative Update for Windows Server 2019 for x64-based Systems' # ( \(1809\))? took this out of the query
       Exclude='Preview|\.NET|Dynamic|Safe OS' }
    @{ Key='Server2019-NET';  Label='Windows Server 2019 - .NET Framework Cumulative Update'
       Search='Cumulative Update .NET Framework Windows Server 2019 x64'
       Include='Cumulative Update for \.NET Framework.*Windows Server 2019.*x64'
       Exclude='Preview|Dynamic' }
    @{ Key='Server2022-CU';   Label='Windows Server 2022 - Cumulative Update'
       Search='Cumulative Update Microsoft server operating system 21H2 x64'
       Include='Cumulative Update for Microsoft server operating system,? version 21H2 for x64-based Systems'
       Exclude='Preview|\.NET|Dynamic|Safe OS' }
)

# Direct/third-party targets: Resolver maps to a Get-<Resolver>Plan function in the engine below.
$DirectTargets = @(
    @{ Key='Edge';     Label='Microsoft Edge Stable (x64 MSI)';                       Resolver='Edge' }
    @{ Key='Defender'; Label='Defender Antivirus definitions (mpam-fe.exe, x64)';     Resolver='Defender' }
    @{ Key='Splunk';   Label='Splunk Universal Forwarder (x64 MSI)';                  Resolver='Splunk' }
    @{ Key='Acrobat';  Label='Adobe Acrobat DC x64 (latest update .msp)';             Resolver='Acrobat' }
)

# =============================================================================
# ENGINE - runs in a background runspace so the window stays responsive
# =============================================================================
$Engine = {
    param($ModulePath, $Dest, $Selected, $ListOnly, $Log)

    $ErrorActionPreference = 'Stop'
    $ProgressPreference    = 'SilentlyContinue'
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    # added line 58 09302026 1050 ssl bypass - Gemini
    [Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
    function Write-Log([string]$m) { $Log.Enqueue($m) }

    Import-Module $ModulePath -Force

    # ---- Resolvers: each returns @{ Title; Id; Files = @( @{Url;Name;Digest;Algorithm;Encoding;Signature;Overwrite} ) }
    function Get-CatalogPlan($t) {
        $rows = @(Get-MicrosoftUpdates -SearchText $t.Search)
        $pick = $rows |
            Where-Object { $_.Title -match $t.Include -and $_.Title -notmatch $t.Exclude } |
            Sort-Object @{ Expression = { [datetime]::ParseExact(([string]$_.'Last Updated').Trim(), 'M/d/yyyy',
                                [Globalization.CultureInfo]::InvariantCulture) }; Descending = $true },
                        @{ Expression = { $_.Title }; Descending = $true } |
            Select-Object -First 1
        if (-not $pick) { throw 'No matching update on page 1 of catalog results (check Search/Include/Exclude).' }
        $info  = Get-MicrosoftUpdateFiles -UpdateID $pick.Id
        $files = foreach ($f in @($info.Files | Where-Object { $_.URL })) {
            $alg = if ($f.PSObject.Properties['DigestAlgorithm']) { $f.DigestAlgorithm } else { 'SHA1' }
            @{ Url = $f.URL; Name = $f.FileName; Digest = $f.Digest; Algorithm = $alg; Encoding = 'Base64' }
        }
        @{ Title = ([string]$pick.Title).Trim(); Id = $pick.Id; Files = @($files) }
    }

    function Get-EdgePlan {
        $p   = Invoke-RestMethod -Uri 'https://edgeupdates.microsoft.com/api/products'
        $rel = ($p | Where-Object Product -eq 'Stable').Releases |
            Where-Object { $_.Platform -eq 'Windows' -and $_.Architecture -eq 'x64' -and $_.Artifacts } |
            Sort-Object { [version]$_.ProductVersion } -Descending | Select-Object -First 1
        $a = $rel.Artifacts | Where-Object ArtifactName -eq 'msi' | Select-Object -First 1
        @{ Title = "Microsoft Edge Stable $($rel.ProductVersion)"; Id = ''
           Files = @(@{ Url = $a.Location; Name = "MicrosoftEdgeEnterpriseX64_$($rel.ProductVersion).msi"
                        Digest = $a.Hash; Algorithm = 'SHA256'; Encoding = 'Hex' }) }
    }

    function Get-DefenderPlan {
        # Definitions change several times a day, so this one always overwrites. Verified by Authenticode signature.
        @{ Title = 'Microsoft Defender Antivirus definitions (x64)'; Id = ''
           Files = @(@{ Url = 'https://go.microsoft.com/fwlink/?LinkID=121721&arch=x64'; Name = 'mpam-fe.exe'
                        Digest = ''; Algorithm = ''; Encoding = ''; Signature = $true; Overwrite = $true }) }
    }

    function Get-SplunkPlan {
        $html  = (Invoke-WebRequest -UseBasicParsing -Uri 'https://www.splunk.com/en_us/download/universal-forwarder.html').Content
        $links = [regex]::Matches($html, 'data-link="([^"]+)"') | ForEach-Object { $_.Groups[1].Value }
        $url   = $links | Where-Object { $_ -match 'windows-x64\.msi$|x64-release\.msi$' } | Select-Object -First 1
        if (-not $url) { throw 'No x64 MSI link found on the Splunk download page (the page layout may have changed).' }
        $name = [IO.Path]::GetFileName(([uri]$url).AbsolutePath)
        $sha  = ''
        try {
            $c = (Invoke-WebRequest -UseBasicParsing -Uri "$url.sha512").Content
            if ($c -is [byte[]]) { $c = [Text.Encoding]::ASCII.GetString($c) }
            $sha = [regex]::Match([string]$c, '[0-9a-fA-F]{128}').Value
        } catch { }
        @{ Title = "Splunk Universal Forwarder ($name)"; Id = ''
           Files = @(@{ Url = $url; Name = $name; Digest = $sha; Algorithm = 'SHA512'; Encoding = 'Hex' }) }
    }

    function Get-AcrobatPlan {
        # Adobe publishes Acrobat DC x64 updates as .msp patches in per-version folders on its enterprise FTP site.
        $base = 'https://ftp.adobe.com/pub/adobe/acrobat/win/AcrobatDC/'
        $html = (Invoke-WebRequest -UseBasicParsing -Uri $base).Content
        $dirs = [regex]::Matches($html, 'href="(?:[^"]*/)?(\d{9,10})/?"') | ForEach-Object { $_.Groups[1].Value } |
            Sort-Object { [int64]$_ } -Descending -Unique | Select-Object -First 8
        foreach ($d in $dirs) {
            $name = "AcrobatDCx64Upd$d.msp"
            $url  = "$base$d/$name"
            try {
                $r = Invoke-WebRequest -UseBasicParsing -Uri $url -Method Head
                if ($r.StatusCode -eq 200) {
                    $ver = '{0}.{1}.{2}' -f $d.Substring(0,2), $d.Substring(2,3), $d.Substring(5)
                    return @{ Title = "Adobe Acrobat DC x64 update $ver"; Id = ''
                              Files = @(@{ Url = $url; Name = $name; Digest = ''; Algorithm = ''; Encoding = '' }) }
                }
            } catch { }
        }
        throw 'Could not find an AcrobatDCx64Upd*.msp in the newest folders on ftp.adobe.com.'
    }

    # ---- Download + verify ----------------------------------------------------
    function Test-Digest($path, $f) {
        if (-not $f.Digest -or -not $f.Algorithm) { return $true }
        $alg = [Security.Cryptography.HashAlgorithm]::Create($f.Algorithm)
        $fs  = [IO.File]::OpenRead($path)
        try { $bytes = $alg.ComputeHash($fs) } finally { $fs.Dispose(); $alg.Dispose() }
        if ($f.Encoding -eq 'Base64') { return ([Convert]::ToBase64String($bytes) -ceq $f.Digest) }
        return (([BitConverter]::ToString($bytes) -replace '-', '') -ieq $f.Digest)
    }

    $rows = New-Object System.Collections.ArrayList
    foreach ($t in $Selected) {
        Write-Log "[$($t.Key)] resolving..."
        try {
            $plan = if ($t.Resolver) { & "Get-$($t.Resolver)Plan" } else { Get-CatalogPlan $t }
            Write-Log "    $($plan.Title)"
            if ($plan.Id) { Write-Log "    UpdateID: $($plan.Id)" }

            $outDir = Join-Path $Dest $t.Key
            if (-not $ListOnly) { New-Item -ItemType Directory -Force -Path $outDir | Out-Null }

            foreach ($f in $plan.Files) {
                $path = Join-Path $outDir $f.Name
                $status = 'ListOnly'
                if (-not $ListOnly) {
                    if ((Test-Path $path) -and -not $f.Overwrite) {
                        Write-Log "    already present: $($f.Name)"; $status = 'Present'
                    } else {
                        $tmp = Join-Path $outDir ".part-$($f.Name)"
                        Write-Log "    downloading $($f.Name) ..."
                        Invoke-WebRequest -UseBasicParsing -Uri $f.Url -OutFile $tmp
                        if (-not (Test-Digest $tmp $f)) { Remove-Item $tmp -Force; throw "Hash mismatch for $($f.Name)" }
                        if ($f.Signature) {
                            $sig = Get-AuthenticodeSignature $tmp
                            if ($sig.Status -ne 'Valid') { Remove-Item $tmp -Force; throw "Signature check failed for $($f.Name): $($sig.Status)" }
                        }
                        Move-Item $tmp $path -Force
                        $status = 'OK'
                        Write-Log ("    OK  ({0:N1} MB)" -f ((Get-Item $path).Length / 1MB))
                    }
                } else { Write-Log "    would download $($f.Name)" }
                [void]$rows.Add([pscustomobject]@{ Target = $t.Key; Title = $plan.Title; UpdateID = $plan.Id; File = $f.Name; Status = $status })
            }
        } catch {
            Write-Log "    ERROR: $($_.Exception.Message)"
            [void]$rows.Add([pscustomobject]@{ Target = $t.Key; Title = ''; UpdateID = ''; File = ''; Status = "ERROR: $($_.Exception.Message)" })
        }
    }
    if (-not $ListOnly -and $rows.Count) { $rows | Export-Csv -NoTypeInformation -Path (Join-Path $Dest 'last-run.csv') }
    Write-Log 'Done.'
}

# =============================================================================
# GUI
# =============================================================================
$cfgPath = Join-Path $env:APPDATA 'PatchDownloader.json'
$saved   = $null
if (Test-Path $cfgPath) { try { $saved = Get-Content $cfgPath -Raw | ConvertFrom-Json } catch { } }

$form = New-Object Windows.Forms.Form
$form.Text = 'Patch & Installer Downloader'
$form.Size = New-Object Drawing.Size(760, 760)
$form.MinimumSize = New-Object Drawing.Size(760, 640)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object Drawing.Font('Segoe UI', 9)

function Add-Label($text, $y) {
    $l = New-Object Windows.Forms.Label; $l.Text = $text; $l.AutoSize = $true
    $l.Location = New-Object Drawing.Point(12, $y); $form.Controls.Add($l)
}
function Add-PathRow($y, $initial, $onBrowse) {
    $tb = New-Object Windows.Forms.TextBox; $tb.Text = $initial
    $tb.Location = New-Object Drawing.Point(12, $y); $tb.Size = New-Object Drawing.Size(600, 24)
    $tb.Anchor = 'Top,Left,Right'; $form.Controls.Add($tb)
    $b = New-Object Windows.Forms.Button; $b.Text = 'Browse...'
    $b.Location = New-Object Drawing.Point(620, ($y - 1)); $b.Size = New-Object Drawing.Size(108, 26)
    $b.Anchor = 'Top,Right'; $b.Add_Click($onBrowse.GetNewClosure()); $form.Controls.Add($b)
    return $tb
}

Add-Label 'Download folder' 10
$txtDest = Add-PathRow 30 $(if ($saved.Dest) { $saved.Dest } else { 'D:\PatchStaging' }) {
    $d = New-Object Windows.Forms.FolderBrowserDialog
    if ($d.ShowDialog() -eq 'OK') { $txtDest.Text = $d.SelectedPath }
}
Add-Label 'Catalog module (PSMicrosoftUpdateCatalog.psm1)' 62
$txtModule = Add-PathRow 82 $(if ($saved.Module) { $saved.Module } else { $ModulePath }) {
    $d = New-Object Windows.Forms.OpenFileDialog; $d.Filter = 'PowerShell module|*.psm1;*.ps1|All files|*.*'
    if ($d.ShowDialog() -eq 'OK') { $txtModule.Text = $d.FileName }
}

$script:checks = @()
function Add-Group($text, $targets, $top) {
    $gb = New-Object Windows.Forms.GroupBox
    $gb.Text = $text
    $gb.Location = New-Object Drawing.Point(12, $top)
    $gb.Size = New-Object Drawing.Size(716, (30 + 26 * $targets.Count))
    $gb.Anchor = 'Top,Left,Right'
    $i = 0
    foreach ($t in $targets) {
        $cb = New-Object Windows.Forms.CheckBox
        $cb.Text = $t.Label; $cb.AutoSize = $true
        $cb.Location = New-Object Drawing.Point(14, (22 + 26 * $i))
        $cb.Checked = if ($saved -and $saved.Checked) { @($saved.Checked) -contains $t.Key } else { $true }
        $gb.Controls.Add($cb)
        $script:checks += [pscustomobject]@{ Target = $t; Box = $cb }
        $i++
    }
    $form.Controls.Add($gb)
    return ($top + $gb.Height + 8)
}
$y = Add-Group 'Windows Update Catalog' $CatalogTargets 116
$y = Add-Group 'Third-party / direct downloads' $DirectTargets $y

$chkList = New-Object Windows.Forms.CheckBox
$chkList.Text = 'List only (resolve and show what would be downloaded, download nothing)'
$chkList.AutoSize = $true; $chkList.Location = New-Object Drawing.Point(14, $y); $form.Controls.Add($chkList)
$y += 30

$btnAll  = New-Object Windows.Forms.Button; $btnAll.Text = 'Select all';  $btnAll.Location  = New-Object Drawing.Point(12, $y);  $btnAll.Size  = New-Object Drawing.Size(90, 28)
$btnNone = New-Object Windows.Forms.Button; $btnNone.Text = 'Select none'; $btnNone.Location = New-Object Drawing.Point(108, $y); $btnNone.Size = New-Object Drawing.Size(90, 28)
$btnRun    = New-Object Windows.Forms.Button; $btnRun.Text = 'Run';           $btnRun.Location    = New-Object Drawing.Point(432, $y); $btnRun.Size    = New-Object Drawing.Size(90, 28)
$btnCancel = New-Object Windows.Forms.Button; $btnCancel.Text = 'Cancel';     $btnCancel.Location = New-Object Drawing.Point(528, $y); $btnCancel.Size = New-Object Drawing.Size(90, 28); $btnCancel.Enabled = $false
$btnOpen   = New-Object Windows.Forms.Button; $btnOpen.Text = 'Open folder';  $btnOpen.Location   = New-Object Drawing.Point(624, $y); $btnOpen.Size   = New-Object Drawing.Size(104, 28)
foreach ($b in $btnRun, $btnCancel, $btnOpen) { $b.Anchor = 'Top,Right' }
$form.Controls.AddRange(@($btnAll, $btnNone, $btnRun, $btnCancel, $btnOpen))
$y += 38

$bar = New-Object Windows.Forms.ProgressBar
$bar.Location = New-Object Drawing.Point(12, $y); $bar.Size = New-Object Drawing.Size(716, 12); $bar.Anchor = 'Top,Left,Right'
$form.Controls.Add($bar)
$y += 20

$txtLog = New-Object Windows.Forms.TextBox
$txtLog.Multiline = $true; $txtLog.ReadOnly = $true; $txtLog.ScrollBars = 'Vertical'
$txtLog.Font = New-Object Drawing.Font('Consolas', 9)
$txtLog.Location = New-Object Drawing.Point(12, $y)
$txtLog.Size = New-Object Drawing.Size(716, ($form.ClientSize.Height - $y - 12))
$txtLog.Anchor = 'Top,Bottom,Left,Right'
$form.Controls.Add($txtLog)

$timer = New-Object Windows.Forms.Timer
$timer.Interval = 250
$script:ps = $null; $script:async = $null; $script:queue = $null

$btnAll.Add_Click({  foreach ($c in $script:checks) { $c.Box.Checked = $true } })
$btnNone.Add_Click({ foreach ($c in $script:checks) { $c.Box.Checked = $false } })
$btnOpen.Add_Click({ if (Test-Path $txtDest.Text) { Start-Process explorer.exe $txtDest.Text } })

$timer.Add_Tick({
    $done = $script:async -and $script:async.IsCompleted
    $line = $null
    while ($script:queue.TryDequeue([ref]$line)) { $txtLog.AppendText($line + "`r`n") }
    if ($done) {
        $timer.Stop()
        try { [void]$script:ps.EndInvoke($script:async) } catch { $txtLog.AppendText("Stopped: $($_.Exception.Message)`r`n") }
        $script:ps.Runspace.Close(); $script:ps.Dispose(); $script:ps = $null; $script:async = $null
        $bar.Style = 'Blocks'; $bar.Value = 0
        $btnRun.Enabled = $true; $btnCancel.Enabled = $false
    }
})

$btnRun.Add_Click({
    $sel = @($script:checks | Where-Object { $_.Box.Checked } | ForEach-Object { $_.Target })
    if (-not $sel)                          { [void][Windows.Forms.MessageBox]::Show('Select at least one target.'); return }
    if (-not (Test-Path $txtModule.Text))   { [void][Windows.Forms.MessageBox]::Show('Catalog module path not found.'); return }
    if (-not $txtDest.Text.Trim())          { [void][Windows.Forms.MessageBox]::Show('Choose a download folder.'); return }

    @{ Dest = $txtDest.Text; Module = $txtModule.Text
       Checked = @($script:checks | Where-Object { $_.Box.Checked } | ForEach-Object { $_.Target.Key }) } |
        ConvertTo-Json | Set-Content $cfgPath

    $txtLog.Clear()
    $btnRun.Enabled = $false; $btnCancel.Enabled = $true; $bar.Style = 'Marquee'
    $script:queue = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]'

    $rs = [runspacefactory]::CreateRunspace()
    $rs.ApartmentState = 'STA'; $rs.ThreadOptions = 'ReuseThread'; $rs.Open()
    $script:ps = [powershell]::Create()
    $script:ps.Runspace = $rs
    [void]$script:ps.AddScript($Engine.ToString()).
        AddArgument($txtModule.Text).AddArgument($txtDest.Text).AddArgument($sel).
        AddArgument([bool]$chkList.Checked).AddArgument($script:queue)
    $script:async = $script:ps.BeginInvoke()
    $timer.Start()
})

$btnCancel.Add_Click({
    if ($script:ps) { $script:queue.Enqueue('Cancelling...'); [void]$script:ps.BeginStop($null, $null); $btnCancel.Enabled = $false }
})
$form.Add_FormClosing({ if ($script:ps) { [void]$script:ps.BeginStop($null, $null) } })

[void]$form.ShowDialog()
