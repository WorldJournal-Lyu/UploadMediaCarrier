<#

UploadMediaCarrier.ps1

    2018-03-19 Initial Creation
    2018-06-19 Modify to GitHub friendly version

#>

if (!($env:PSModulePath -match 'C:\\PowerShell\\_Modules')) {
    $env:PSModulePath = $env:PSModulePath + ';C:\PowerShell\_Modules\'
}

Import-Module WorldJournal.Ftp -Verbose -Force
Import-Module WorldJournal.Log -Verbose -Force
Import-Module WorldJournal.Email -Verbose -Force
Import-Module WorldJournal.Server -Verbose -Force

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

Write-Log -Verb "LOG START" -Noun $log -Path $log -Type Long -Status Normal
Write-Line -Length 100 -Path $log

###################################################################################



# Set up variables

#$ftp      = Get-WJFTP -Name WorldJournalNewYork
$ftp      = Get-WJFTP -Name MediaCarrier
$workDate = ((Get-Date).AddDays(0)).ToString("yyyyMMdd")
$ePaper   = Get-WJPath -Name epaper
$optimizeda = $ePaper.Path + $workDate + "\" + "optimizeda"
$pubcode   = @("NY","CH")
$section   = @("A","B","C","D")

Write-Log -Verb "ftp" -Noun $ftp.Path -Path $log -Type Short -Status Normal
Write-Log -Verb "ePaper" -Noun $ePaper.Path -Path $log -Type Short -Status Normal
Write-Log -Verb "workDate" -Noun $workDate -Path $log -Type Short -Status Normal
Write-Log -Verb "optimizeda" -Noun $optimizeda -Path $log -Type Short -Status Normal

$localTemp = "C:\temp\" + $scriptName + "\"
if (!(Test-Path($localTemp))) {New-Item $localTemp -Type Directory | Out-Null}

foreach ($pub in $pubcode){
    
    # define regex
    
    $regex = $pub + $workDate + "("
    foreach($sec in $section){$regex = $regex + ($sec + "|")}
    $regex = $regex.Substring(0,$regex.LastIndexOf("|")) + ")\d{2}(.pdf)"

    Get-ChildItem $optimizeda | Where-Object{$_.Name -match $regex} | ForEach-Object{

        try{
            Copy-Item $_.FullName ($localTemp + "\" + $_.Name) -ErrorAction Stop
            Write-Log -Verb "COPY" -Noun $_.FullName -Path $log -Type Long -Status Good

        }catch{
            Write-Log -Verb "COPY" -Noun $_.FullName -Path $log -Type Long -Status Bad
        
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