<#

UploadMediaCarrier.ps1

    2018-03-19 Initial Creation

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
Write-Line -Length 50 -Path $log

###################################################################################



# Set up variables

$workDate = ((Get-Date).AddDays(0))
$pubcode  = "'114', '146'"
$section  = "'A', 'B', 'C', 'D'"
$db       = Get-WJDatabase -Name marco6
$ftp      = Get-WJFTP -Name MediaCarrier # -Name MediaCarrier # -Name WorldJournalNewYork
$ePaper   = Get-WJPath -Name epaper
$optimizeda = $ePaper.Path + $workDate.ToString("yyyyMMdd") + "\" + "optimizeda\"

Write-Log -Verb "workDate" -Noun $workDate.ToString("yyyyMMdd") -Path $log -Type Short -Status Normal
Write-Log -Verb "db" -Noun $db.Name -Path $log -Type Short -Status Normal
Write-Log -Verb "ftp" -Noun $ftp.Path -Path $log -Type Short -Status Normal
Write-Log -Verb "ePaper" -Noun $ePaper.Path -Path $log -Type Short -Status Normal
Write-Log -Verb "optimizeda" -Noun $optimizeda -Path $log -Type Short -Status Normal
Write-Line -Length 50 -Path $log



# Set up query

[String]$usrnm = $db.Username
[String]$pswrd = $db.Password
[String]$dtsrc = $db.Datasource
[String]$qry   = Get-Content -Path ((Split-Path $MyInvocation.MyCommand.Path -Parent)+"\"+($MyInvocation.MyCommand.Name -replace '.ps1', '.Query.txt'))
$qry = $qry.Replace('$workDate', $workDate.ToString("yyyy-MM-dd"))
$qry = $qry.Replace('$pubcode', $pubcode)
$qry = $qry.Replace('$section', $section)
$result = Query-Database -Username $usrnm -Password $pswrd -Datasource $dtsrc -Query $qry



# Process query result for each pubid

$result.pubid | Select-Object -Unique | ForEach-Object{
    
    $pubid = $_
    $expectedPage = ($result | Where-Object { $_.pubid -eq $pubid }).Count

    switch($pubid){
        "114" { $pubname = "NY"; break; }
        "141" { $pubname = "NJ"; break; }
        "142" { $pubname = "DC"; break; }
        "143" { $pubname = "BO"; break; }
        "144" { $pubname = "AT"; break; }
        "146" { $pubname = "CH"; break; }
    }

    Write-Log -Verb "pubid" -Noun $pubid -Path $log -Type Short -Status Normal
    Write-Log -Verb "pubname" -Noun $pubname -Path $log -Type Short -Status Normal
    Write-Log -Verb "expectedPage" -Noun $expectedPage -Path $log -Type Short -Status Normal
    Write-Line -Length 50 -Path $log

    $acrobat = New-Object -ComObject AcroExch.AVDoc
    $pdf     = $null
    $pdf2    = $null
    $isFirst = $true
    
    $result | Where-Object{ $_.pubid -eq $pubid } | Select-Object section, page | ForEach-Object{

        $pdfName1 = $pubname + $workDate.ToString("yyyyMMdd") + $_.section + ($_.page).ToString("00") + ".pdf"
        $copyFrom = $optimizeda + $pdfName1
        $copyTo   = $localTemp + $pdfName1
        Write-Log -Verb "pdfName1" -Noun $pdfName1 -Path $log -Type Short -Status Normal
        Write-Log -Verb "copyFrom" -Noun $copyFrom -Path $log -Type Short -Status Normal
        Write-Log -Verb "copyTo" -Noun $copyTo -Path $log -Type Short -Status Normal



        if(Test-Path $copyFrom){

            Write-Log -Verb "FILE EXIST" -Noun $copyFrom -Path $log -Type Long -Status System



            # Copy pdf file to temp folder

            try{

                Write-Log -Verb "COPY FROM" -Noun $copyFrom -Path $log -Type Long -Status Good
                Copy-Item $copyFrom $copyTo -ErrorAction Stop
                Write-Log -Verb "COPY TO" -Noun $copyTo -Path $log -Type Long -Status Good

            }catch{

                Write-Log -Verb "COPY TO" -Noun $copyTo -Path $log -Type Long -Status Bad
        
            }



            # Merge pdf in workpath together

            $pdfSize = ("{0:N2}" -f (((Get-Item $copyTo).Length)/1MB))

	        try{

                if($isFirst) {

		            $isFirst = $false

		            $acrobat.Open($copyTo, "temp") | Out-Null
		            $pdf = $acrobat.GetPDDoc()

	            }else{

		            $acrobat2 = New-Object -ComObject AcroExch.AVDoc
		            $acrobat2.Open($copyTo, "temp") | Out-Null
		            $pdf2 = $acrobat2.GetPDDoc()

		            $pdf.InsertPages(($pdf.GetNumPages()-1), $pdf2, 0, $pdf2.GetNumPages(), 0) | Out-Null
		            $pdf2.Close() | Out-Null
		            $acrobat2.Close(1) | Out-Null

	            }

                Write-Log -Verb "MERGE" -Noun ($copyTo + " (" + $pdfSize + " MB)") -Path $log -Type Long -Status Good

            }catch{

                Write-Log -Verb "MERGE" -Noun ($copyTo + " (" + $pdfSize + " MB)") -Path $log -Type Long -Status Bad
                Write-Log -Verb "Exception" -Noun $_.Exception.Message -Path $log -Type Short -Status Bad

            }

        }else{
        
            Write-Log -Verb "FILE NOT EXIST" -Noun $copyFrom -Path $log -Type Long -Status Bad

        }
    }



    # Save merged pdf file

    Write-Line -Length 50 -Path $log

    $pdfName2 = "WorldJournal_" + $pubname + "_" + $workDate.ToString("yyyyMMdd") + ".pdf"

    $output = $localTemp + $pdfName2
    $pdf.Save(1, $output) | Out-Null
    $outputPage = 0
    $outputPage = $pdf.GetNumPages()
    $pdf.Close() | Out-Null
    $acrobat.Close(1) | Out-Null
    $outputSize = ("{0:N2}" -f (((Get-Item $output).Length)/1MB))
    Write-Log -Verb "SAVE FILE" -Noun ($output + " (" + $outputSize + " MB)") -Path $log -Type Long -Status Normal
    if($expectedPage -eq $outputPage){
        Write-Log -Verb "PAGE CHECK" -Noun ($outputPage.ToString() + " out of " + $expectedPage.ToString()) -Path $log -Type Long -Status Good
    }else{
        Write-Log -Verb "PAGE CHECK" -Noun ($outputPage.ToString() + " out of " + $expectedPage.ToString()) -Path $log -Type Long -Status Bad
        $hasError = $true
    }
    Stop-Process -Name Acrobat



    # Set up upload and download variables

    $uploadFrom  = $localTemp + $pdfName2
    $uploadTo = $ftp.Path + $pdfName2
    $downloadFrom  = $uploadTo
    $downloadTo = $localTemp + (Get-Date).ToString("yyyyMMdd-HHmmss") + ".pdf"

    Write-Log -Verb "pdfName2" -Noun $pdfname2 -Path $log -Type Short -Status Normal
    Write-Log -Verb "uploadFrom" -Noun $uploadFrom -Path $log -Type Short -Status Normal
    Write-Log -Verb "uploadTo" -Noun $uploadTo -Path $log -Type Short -Status Normal
    Write-Log -Verb "downloadFrom" -Noun $downloadFrom -Path $log -Type Short -Status Normal
    Write-Log -Verb "downloadTo" -Noun $downloadTo -Path $log -Type Short -Status Normal



    # Upload file from local to Ftp

    $upload = WebClient-UploadFile -Username $ftp.User -Password $ftp.Pass -RemoteFilePath $uploadTo -LocalFilePath $uploadFrom

    if($upload.Status -eq "Good"){
        Write-Log -Verb $upload.Verb -Noun $upload.Noun -Path $log -Type Long -Status $upload.Status
    }elseif($upload.Status -eq "Bad"){
        $mailMsg = $mailMsg + (Write-Log -Verb $upload.Verb -Noun $upload.Noun -Path $log -Type Long -Status $upload.Status -Output String) + "`n"
        $mailMsg = $mailMsg + (Write-Log -Verb "Exception" -Noun $upload.Exception -Path $log -Type Short -Status $upload.Status -Output String) + "`n"
    }



    # Download file from Ftp to temp folder for verification

    $download = WebClient-DownloadFile -Username $ftp.User -Password $ftp.Pass -RemoteFilePath $downloadFrom -LocalFilePath $downloadTo

    if($download.Status -eq "Good"){
        Write-Log -Verb $download.Verb -Noun $download.Noun -Path $log -Type Long -Status $download.Status
    }elseif($download.Status -eq "Bad"){
        $mailMsg = $mailMsg + (Write-Log -Verb $download.Verb -Noun $download.Noun -Path $log -Type Long -Status $download.Status -Output String) + "`n"
        $mailMsg = $mailMsg + (Write-Log -Verb "Exception" -Noun $download.Exception -Path $log -Type Short -Status $download.Status -Output String) + "`n"
    }



    Write-Line -Length 50 -Path $log

}



# Delete temp folder

Write-Log -Verb "REMOVE" -Noun $localTemp -Path $log -Type Long -Status Normal
try{
    $temp = $localTemp
    Remove-Item $localTemp -Recurse -Force -ErrorAction Stop
    Write-Log -Verb "REMOVE" -Noun $temp -Path $log -Type Long -Status Good
}catch{
    $mailMsg = $mailMsg + (Write-Log -Verb "REMOVE" -Noun $temp -Path $log -Type Long -Status Bad -Output String) + "`n"
    $mailMsg = $mailMsg + (Write-Log -Verb "Exception" -Noun $_.Exception.Message -Path $log -Type Short -Status Bad -Output String) + "`n"
}



# Set hasError status 

if( ($expectedPage -ne $outputPage) -or ($upload.Status -eq "Bad") -or ($download.Status -eq "Bad") ){
    $hasError = $true
}



###################################################################################

Write-Line -Length 50 -Path $log
Write-Log -Verb "LOG END" -Noun $log -Path $log -Type Long -Status Normal
if($hasError){ $mailSbj = "ERROR " + $scriptName }

$emailParam = @{
    From    = $mailFrom
    Pass    = $mailPass
    To      = $mailTo
    Subject = $mailSbj
    Body    = $mailMsg
    ScriptPath = $scriptPath
    Attachment = $log
}
Emailv2 @emailParam