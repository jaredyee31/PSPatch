Add-Type -AssemblyName System.Web
$Script:MU_NextPage = ""
$Script:MU_NextPage_I = 0
$Script:MU_SearchText = ""

function Get-MicrosoftUpdates{
param($SearchText,[switch]$NextPage)
    if($NextPage.IsPresent){
        if($MU_SearchText -eq $null -or $MU_SearchText -eq ""){
            Write-Host "A previous search query is required before using this parameter." -ForegroundColor Red
            return
        }

        $Script:MU_NextPage_I++
        $Microsoft_Updates = (Invoke-WebRequest -UseBasicParsing -Uri "https://www.catalog.update.microsoft.com/Search.aspx?q=$MU_SearchText&p=$($Script:MU_NextPage_I)" -Method POST -ContentType "application/x-www-form-urlencoded" -Body $NextPage_Query).Content
    }elseif($SearchText){
        $SearchText = [System.Web.HttpUtility]::UrlEncode($SearchText)
        $Script:MU_NextPage_I = 0
        $Script:MU_SearchText = $SearchText
        $Microsoft_Updates = (Invoke-WebRequest -UseBasicParsing -Uri "https://www.catalog.update.microsoft.com/Search.aspx?q=$SearchText" -Method "POST" -ContentType "application/x-www-form-urlencoded" -Body "updateIDs=&updateIDsBlockedForImport=&contentImport=&sku=&serverName=&ssl=&portNumber=&version=&protocol=").Content
    }else{
        Write-Host "This function requires either a SearchText or a NextPage parameter! None were given..." -ForegroundColor Red
        return
    }

    $Microsoft_Updates_HTML = New-Object -Com "HTMLFile"
    $Microsoft_Updates_HTML.IHTMLDocument2_write($Microsoft_Updates)
    $Microsoft_Updates_NextPage = try{(($Microsoft_Updates_HTML.getElementById("ctl00_catalogBody_nextPageLink")).getElementsByTagName("a")).Length -eq 1}catch{}
    $Microsoft_Updates_NextPage_RequestInfo = $null

    if($Microsoft_Updates_NextPage){
        $Microsoft_Updates_NextPage_RequestInfo = [ordered]@{
            '__EVENTTARGET' = 'ctl00$catalogBody$nextPageLinkText'
            '__EVENTARGUMENT' = ''
            '__VIEWSTATE' = $Microsoft_Updates_HTML.getElementById("__VIEWSTATE").value
            '__VIEWSTATEGENERATOR' = $Microsoft_Updates_HTML.getElementById("__VIEWSTATEGENERATOR").value
            '__EVENTVALIDATION' = $Microsoft_Updates_HTML.getElementById("__EVENTVALIDATION").value
            'ctl00$searchTextBox' = $SearchText
            'updateIDs' = ''
            'contentImport' = ''
            'sku' = ''
            'serverName' = ''
            'ssl' = ''
            'portNumber' = ''
            'version' = ''
            'protocol' = ''
        }
    }

    $DownloadTable = $Microsoft_Updates_HTML.getElementById("tableContainer")
    $DownloadTable_Headers = @("Id")
    $DownloadTable_Headers += $DownloadTable.getElementsByTagName("tr")[0].getelementsbytagname("span") | foreach {$_.innertext}
    $DownloadTable_Headers = $DownloadTable_Headers[0..($DownloadTable_Headers.Length-2)]
    $DownloadTable_Headers += "Size2"
    $DownloadTable_Rows =  New-Object System.Collections.ArrayList($null)
    $DownloadTable.getElementsByTagName("tr") | where {$_.id -ne "headerRow"} | foreach{
        $Row = $_
        $RowID = $Row.id.Substring(0,$Row.id.IndexOf("_"))
        $RowData = [ordered]@{$DownloadTable_Headers[0]=$RowID}
        $i=1
        $Row.getElementsByTagName("td") | where {$_.classname -notmatch "resultsIconWidth|resultsButtonWidth"} | foreach {
            if($_.innerhtml -match "SPAN"){
                $_.getElementsByTagName("span") | foreach {$RowData.Add($DownloadTable_Headers[$i],$_.innertext);$i++}
            }else{
                $RowData.Add($DownloadTable_Headers[$i],$_.innertext)
            }
            $i++
        }
        $DownloadTable_Rows.Add([PSCustomObject]$RowData) | Out-Null
    }

    $Script:MU_NextPage = $Microsoft_Updates_NextPage_RequestInfo
    return $DownloadTable_Rows

    <#
    return @{
        "SearchText" = $SearchText
        "Updates" = $DownloadTable_Rows
        "NextPage"=$Microsoft_Updates_NextPage_RequestInfo
    }
    #>
}

function Get-MicrosoftUpdateFiles{
param($UpdateID)
    $Microsoft_UpdateIDs_JSON = ConvertTo-Json -Compress -InputObject @([ordered]@{"size"=0;"languages"="";"uidInfo"=$UpdateID;"updateID"=$UpdateID})
    $Microsoft_Query = [ordered]@{"updateIDs"=$Microsoft_UpdateIDs_JSON;"updateIDsBlockedForImport"="";"wsusApiPresent"="";"contentImport"="";"sku"="";"serverName"="";"ssl"="";"portNumber"="";"version"=""}
    $query = [System.Web.HttpUtility]::ParseQueryString('');$Microsoft_Query.GetEnumerator() | ForEach-Object { $query[$_.Key] = $_.Value };$Microsoft_Query = $query.ToString()

    $Microsoft_Download = Invoke-WebRequest -UseBasicParsing -Uri "https://www.catalog.update.microsoft.com/DownloadDialog.aspx" -Method POST -ContentType "application/x-www-form-urlencoded" -Body $Microsoft_Query
    $Microsoft_Download = ($Microsoft_Download.Content.Trim() -split "`r`n" | where {$_ -match 'downloadInformation\[0\]'})

    $Microsoft_Download_Info = @{'Title'='';'UpdateID'='';'Files'=@()}
    $Microsoft_Download_File = [PSCustomObject]@{'URL'='';'Digest'='';'Architectures'='';'Languages'='';'LongLanguages'='';'FileName'='';'DefaultFileNameLength'=''}
    $Microsoft_Download_File_I = -1

    $Microsoft_Download | foreach {
        if($_ -match [regex]::Escape("downloadInformation[0].enTitle")){
            $Microsoft_Download_Info.Title = $_.Substring($_.IndexOf("'")+1,$_.Length-($_.IndexOf("'")+3))
        }elseif($_ -match [regex]::Escape("downloadInformation[0].updateID")){
            $Microsoft_Download_Info.UpdateID= $_.Substring($_.IndexOf("'")+1,$_.Length-($_.IndexOf("'")+3))
        }elseif($_ -match "$([regex]::Escape("downloadInformation[0].files"))\[\d*\] = new Object"){
            $Microsoft_Download_Info.Files += $Microsoft_Download_File.psobject.Copy()
            $Microsoft_Download_File_I++
        }elseif($_ -match "$([regex]::Escape("downloadInformation[0].files"))\[\d*\]\.defaultFileNameLength ="){
            $Microsoft_Download_Info.Files[$Microsoft_Download_File_I].DefaultFileNameLength = $_.Substring($_.IndexOf(" = ")+3,$_.Length-($_.IndexOf(" = ")+4))
        }elseif($_ -match "$([regex]::Escape("downloadInformation[0].files"))\[\d*\]\.(\w*) ="){
            if (($Microsoft_Download_Info.Files[$Microsoft_Download_File_I] | Get-Member -MemberType NoteProperty).Name -notcontains $Matches[1]){
                $Microsoft_Download_Info.Files[$Microsoft_Download_File_I] | Add-Member -MemberType NoteProperty -Name $Matches[1] -Value ($_.Substring($_.IndexOf("'")+1,$_.Length-($_.IndexOf("'")+3)))
            }else{
                $Microsoft_Download_Info.Files[$Microsoft_Download_File_I].$($Matches[1]) = $_.Substring($_.IndexOf("'")+1,$_.Length-($_.IndexOf("'")+3))
            }
        }
    }

    return [PsCustomObject]$Microsoft_Download_Info
}

function Get-MicrosoftUpdate{
param($UpdateID)
    $WU = Invoke-WebRequest -UseBasicParsing -Uri "https://www.catalog.update.microsoft.com/ScopedViewInline.aspx?updateid=$UpdateID"
    $Microsoft_Update_HTML = New-Object -Com "HTMLFile"
    $Microsoft_Update_HTML.IHTMLDocument2_write($WU.Content)

    $Microsoft_Update_Body_ContentSection = $Microsoft_Update_HTML.getElementById("contentSection").innertext

    $Microsoft_Update_Body = $Microsoft_Update_HTML.getElementById("tabBody")
    $Microsoft_Update_Body_Overview = $Microsoft_Update_Body.outerText
    return $Microsoft_Update_Body_ContentSection + "`n`n" + $Microsoft_Update_Body_Overview
}
