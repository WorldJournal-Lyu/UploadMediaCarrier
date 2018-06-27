<#

UploadMediaCarrier.ps1

    2018-03-19 Initial Creation
    2018-06-19 Modify to GitHub friendly version
    2018-06-27 Modify file search method, query database for actual page info instead of get-childitem

#>

if (!($env:PSModulePath -match 'C:\\PowerShell\\_Modules')) {
    $env:PSModulePath = $env:PSModulePath + ';C:\PowerShell\_Modules\'
}

Import-Module WorldJournal.Ftp -Verbose -Force
Import-Module WorldJournal.Log -Verbose -Force
Import-Module WorldJournal.Email -Verbose -Force
Import-Module WorldJournal.Server -Verbose -Force
Import-Module WorldJournal.Database -Verbose -Force

$scriptPath = $MyInvocation.MyCommand.Path
$scriptName = (($MyInvocation.MyCommand) -Replace ".ps1")
$hasError   = $false

$newlog     = New-Log -Path $scriptPath -LogFormat yyyyMMdd-HHmmss
$log        = $newlog.FullName
$logPath    = $newlog.Directory

$mailFrom   = (Get-WJEmail -Name noreply).MailAddress
$mailPass   = (Get-WJEmail -Name noreply).Password
$mailTo     = (Get-WJEmail -Name lyu).MailAddress
$mailSbj    = $scriptName
$mailMsg    = ""

$localTemp = "C:\temp\" + $scriptName + "\"
if (!(Test-Path($localTemp))) {New-Item $localTemp -Type Directory | Out-Null}

Write-Log -Verb "LOG START" -Noun $log -Path $log -Type Long -Status Normal
Write-Line -Length 100 -Path $log

###################################################################################



# Set up variables

$workDate = ((Get-Date).AddDays(0))
$pubcode  = "'114', '146'"
$section  = "'A', 'B', 'C', 'D'"
$db       = Get-WJDatabase -Name marco6
$ftp      = Get-WJFTP -Name MediaCarrier #WorldJournalNewYork (for testing)
$ePaper   = Get-WJPath -Name epaper
$optimizeda = $ePaper.Path + $workDate.ToString("yyyyMMdd") + "\" + "optimizeda\"

Write-Log -Verb "workDate  " -Noun $workDate.ToString("yyyyMMdd") -Path $log -Type Short -Status Normal
Write-Log -Verb "db        " -Noun $db.Name -Path $log -Type Short -Status Normal
Write-Log -Verb "ftp       " -Noun $ftp.Path -Path $log -Type Short -Status Normal
Write-Log -Verb "ePaper    " -Noun $ePaper.Path -Path $log -Type Short -Status Normal
Write-Log -Verb "optimizeda" -Noun $optimizeda -Path $log -Type Short -Status Normal



# Set up query

[String]$usrnm = $db.Username
[String]$pswrd = $db.Password
[String]$dtsrc = $db.Datasource
[String]$qry = Get-Content -Path ((Split-Path $MyInvocation.MyCommand.Path -Parent)+"\"+($MyInvocation.MyCommand.Name -replace '.ps1', '.Query.txt'))
$qry = $qry.Replace('$workDate', $workDate.ToString("yyyy-MM-dd"))
$qry = $qry.Replace('$pubcode', $pubcode)
$qry = $qry.Replace('$section', $section)
$result = Query-Database -Username $usrnm -Password $pswrd -Datasource $dtsrc -Query $qry



# Process query result

$result.pubid | Select-Object -Unique | ForEach-Object{
    
    $pubid = $_

    switch($pubid){
        "114" { $pubname = "NY"; break; }
        "141" { $pubname = "NJ"; break; }
        "142" { $pubname = "DC"; break; }
        "143" { $pubname = "BO"; break; }
        "144" { $pubname = "AT"; break; }
        "146" { $pubname = "CH"; break; }
    }

    $pubname

    $result | Where-Object{ $_.pubid -eq $pubid } | Select-Object section, page | ForEach-Object{

        $pdfName = $pubname + $workDate.ToString("yyyyMMdd") + $_.section + ($_.page).ToString("00") + ".pdf"
        $copyFrom = $optimizeda + $pdfName
        $copyTo   = $localTemp + $pdfName

        try{

            Write-Log -Verb "COPY FROM" -Noun $copyFrom -Path $log -Type Long -Status Normal
            Copy-Item $copyFrom $copyTo -ErrorAction Stop
            Write-Log -Verb "COPY TO" -Noun $copyTo -Path $log -Type Long -Status Good

        }catch{

            Write-Log -Verb "COPY TO" -Noun $copyTo -Path $log -Type Long -Status Bad
        
        }

    }
}


# Delete temp folder

Write-Log -Verb "REMOVE" -Noun $localTemp -Path $log -Type Long -Status Normal
try{
    $temp = $localTemp
    Remove-Item $localTemp -Recurse -Force -ErrorAction Stop
    $mailMsg = $mailMsg + (Write-Log -Verb "REMOVE" -Noun $temp -Path $log -Type Long -Status Good -Output String) + "`n"
}catch{
    $mailMsg = $mailMsg + (Write-Log -Verb "REMOVE" -Noun $temp -Path $log -Type Long -Status Bad -Output String) + "`n"
    $mailMsg = $mailMsg + (Write-Log -Verb "Exception" -Noun $_.Exception.Message -Path $log -Type Short -Status Bad -Output String) + "`n"
}



# Set hasError status 

if( $false ){
    $hasError = $true
}



###################################################################################

Write-Line -Length 100 -Path $log
Write-Log -Verb "LOG END" -Noun $log -Path $log -Type Long -Status Normal
if($hasError){ $mailSbj = "ERROR " + $scriptName }

$emailParam = @{
    From    = $mailFrom
    Pass    = $mailPass
    To      = $mailTo
    Subject = $mailSbj
    Body    = $scriptName + " completed at " + (Get-Date).ToString("HH:mm:ss") + "`n`n" + $mailMsg
    ScriptPath = $scriptPath
    Attachment = $log
}
#Emailv2 @emailParam