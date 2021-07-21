<#
    .SYNOPSIS
        This script will scrape the Forscan site for the most recent test setup files for specified versions.

    .DESCRIPTION
        A longer description.

    .PARAMETER $TestVersions
        This enables the script to look for test versions of Forscan.

    .PARAMETER $MaxDays
        This is the number of days back to search from todays date to start downloading the files.

    .EXAMPLE
        ScriptName -TestVer1 "2.3.42" -TestVer2 "2.4.3" -MaxDays "90" -SaveFiles

    #>

param (
    [Parameter()]
    [switch]$TestVersions,
    [Parameter()]
    $MaxDays = 90
)

$ProgressPreference = "SilentlyContinue"
$rootUrl = "https://forscan.org/download"
$saveFileRoot = $PSScriptRoot
$saveFileFolder = "FORScanSetup"
$Global:filesFound = @()
$Global:publicVer = $null
$Global:TestVer1 = $null
$Global:TestVer2 = '2.4.3'

Function Format-Size {
    param(
        [Parameter(Mandatory = $True)]    
        $Bytes
    )
    $sizes = 'Bytes,KB,MB,GB,TB,PB,EB,ZB' -split ','
    for ($i = 0; ($Bytes -ge 1kb) -and
        ($i -lt $sizes.Count); $i++) { $Bytes /= 1kb }
    $N = 2; if ($i -eq 0) { $N = 0 }
    "{0:N$($N)} {1}" -f $Bytes, $sizes[$i]
}

Function Get-PublicVersion {
    $url = "https://forscan.org/download.html"
    $WebResponse = Invoke-WebRequest $url -TimeoutSec 15 -ErrorAction SilentlyContinue
    if ($WebResponse.StatusCode -eq 200) {
        $links = $WebResponse.Links | Where-Object { $_.href.StartsWith("download/FORScanSetup") } | Select-Object href
        if ($links.Count -gt 0) {
            $a = $links[0].href.Replace("download/FORScanSetup", '')
            $a = $a.Replace(".exe", '')
            $Global:publicVer = $a
            $b = [version]$a.Replace('.beta', '').Replace(".test", '')
            $b = [version]::New($b.Major, $b.Minor, $b.Build + 1)
            $Global:TestVer1 = $b.ToString()
            return $a
        }
        else {
            Write-Host "Unable to Scrape Release version from FORScan Site | Error: $($_.Exception.Message)" -ForegroundColor Red 
        }
    }
    else {
        Write-Host "Unable to Scrape Release version from FORScan Site | Error: $($_.Exception.Message)" -ForegroundColor Red
    }
    Return $null
}

Function Get-PublicFile {
    if ($Global:publicVer) {
        $fileName = "FORScanSetup$($Global:publicVer).exe"
        if ((Get-SetupFile -url "$rootUrl/$fileName" -file $fileName -ver $null -type "Release") -eq $true) {
            # Write-Host "Found $fileName"
        }
    }
    else {
        Write-Host "Can't download public version because PublicVer variable is missing..." -ForegroundColor Yellow
    }
}

Function Get-TestVersionFiles {
    param(
        [Parameter(Mandatory = $true)]
        [String]$ver1, 
        [Parameter()]
        [String]$ver2, 
        [Parameter(Mandatory = $true)]
        [int]$maxDays
    )
    $versions = @($ver1)
    if ($null -ne $ver2) { $versions += $ver2 }
    Write-Host ""
    Write-Host "****************** TEST VERSION DOWNLOAD *******************" -ForegroundColor Magenta
    Write-Host "Search Date Start: $((get-date).ToString("MM-dd-yyyy"))" -ForegroundColor Magenta
    Write-Host "Search Date End: $((get-date).AddDays(-$maxDays).ToString("MM-dd-yyyy"))" -ForegroundColor Magenta
    Write-Host "******************** RESULTS START ********************" -ForegroundColor White
    foreach ($version in $versions) {
        $fileDate = (get-date).ToString("yyyyMMdd")
        Write-Host "Searching for Setup Files | Version: (v$version)" -ForegroundColor DarkCyan
        for ($i = 0; ($i -lt ($maxDays + 1)); $i++) {
            if ($i -gt 0) { Start-Sleep -s 1.5 }
            $fileName = "FORScanSetup$($version).test$($fileDate).exe"
            $result = (Get-SetupFile -url "$($rootUrl)/$($fileName)" -file $fileName -fileDate $fileDate -ver $version -type "Test")
            $fileDate = (get-date).AddDays(-$i).ToString("yyyyMMdd")
            if ($result -eq $true) {
                break
            }
        }
    }
    Write-Host "******************** RESULTS END ********************" -ForegroundColor Magenta
}

Function Get-SetupFile {
    param(
        [Parameter(Mandatory = $true)]
        [String]$url, 
        [Parameter(Mandatory = $true)]
        [String]$file, 
        [Parameter()]
        [String]$ver,
        [Parameter()]
        [String]$fileDate,
        [Parameter(Mandatory = $true)]
        [String]$type
    )
    # Write-Output "Get-SetupFile | URL: $url | FILE: $file | VER: $ver | TYPE: $type"
    try {
        $req = Invoke-WebRequest -uri $url -Method GET -TimeoutSec 15 -ErrorAction SilentlyContinue
        # Write-Output $req
        if ($req.StatusCode -eq 200) {
            $fileContent = $req.Content
            $fileSize = Format-Size -bytes $req.Headers['Content-Length']
            Write-Host "  ------------------- FILE FOUND ------------------  " -ForegroundColor Green
            Write-Host " | Type: $type" -ForegroundColor Green
            If ($ver) {
                Write-Host " | Version: (v$ver)" -ForegroundColor Green
            }
            
            Write-Host " | URL: $url" -ForegroundColor Green
            if ($type -eq "Test") {
                $Global:filesFound += "found"
            }
            if ($null -ne $fileContent) {
                if (!(Test-Path "$saveFileRoot/$saveFileFolder")) {
                    New-Item -Path $saveFileRoot -Name $saveFileFolder -ItemType "directory" -Force | Out-Null
                }
                if (!(Test-Path -Path "$saveFileRoot/$saveFileFolder/$file")) {
                    try {
                        [io.file]::WriteAllBytes("$saveFileRoot/$saveFileFolder/$file", $fileContent)
                        Write-Host " | File Path: $saveFileRoot/$saveFileFolder" -ForegroundColor Green
                        Write-Host " | File Name: $file" -ForegroundColor Green
                        Write-Host " | File Status: " -ForegroundColor Green -NoNewline
                        Write-Host "Saved Successfully..." -ForegroundColor Green
                    }
                    catch {
                        Write-Host " | File Status: " -ForegroundColor Green -NoNewline
                        Write-Host "Error Saving File $($_.Exception.Message)" -ForegroundColor Red
                    }
                }
                else {
                    Write-Host " | File Path: $saveFileRoot/$saveFileFolder" -ForegroundColor Green
                    Write-Host " | File Name: $file" -ForegroundColor Green
                    Write-Host " | File Status: " -ForegroundColor Green -NoNewline
                    Write-Host "Not Downloaded (File Exists)" -ForegroundColor DarkYellow
                }
                Write-Host " | FileSize: ($fileSize)" -ForegroundColor Green
            }
            Write-Host "  -------------------------------------------------  " -ForegroundColor Green
            return $true
        }
    }
    catch {
        # Write-Host "Get-SetupFile Error: $($_.Exception.Message)"
    }
    return $false
}


Function Start-Execution {
    # Downloads the Current Public Release Version of FORScanSetup
    Get-PublicVersion
    Write-Host "************** SCRIPT PARAMETERS ***************" -ForegroundColor Blue
    Write-Host "*" -NoNewline -ForegroundColor Blue
    Write-Host " Test Version Flag: $($TestVersions)"
    If ($TestVersions -eq $true) {
        Write-Host "*" -NoNewline -ForegroundColor Blue
        Write-Host " Test Version: ($Global:TestVer1)"
        Write-Host "*" -NoNewline -ForegroundColor Blue
        Write-Host " Test Version (2.4.x): ($Global:TestVer2)"
        Write-Host "*" -NoNewline -ForegroundColor Blue
        Write-Host " Days to Search: ($MaxDays)"
    }
    Write-Host "************************************************" -ForegroundColor Blue
    Write-Host ""
    Write-Host "************** RELEASE VERSION DOWNLOAD*******************" -ForegroundColor White
    Write-Host "* Public Version: ($Global:publicVer)"
    Get-PublicFile
    Write-Host "**************************************************" -ForegroundColor White

    if ($TestVersions) {
        Get-TestVersionFiles -ver1 $Global:TestVer1 -ver2 $Global:TestVer2 -maxDays $MaxDays
        if (!($Global:filesFound.Count -gt 0)) {
            Write-Host "No Test Versions Found for these Versions ($Global:TestVer1, $Global:TestVer2) between [$((get-date).ToString("yyyyMMdd")) - $((get-date).AddDays(-$MaxDays).ToString("yyyyMMdd"))]" -ForegroundColor Yellow
        }
    }

    Write-Host ""
    Write-Host "Script Execution is Now Complete..." -ForegroundColor Cyan
}
Start-Execution
