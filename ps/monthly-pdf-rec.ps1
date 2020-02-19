param (
    $symFileLoc,
    $rdbFileLoc,
    $runType,
    $day,
    $projFolder,
    $getlatest = $false,
    $option
)
$now = get-date -format "yyyy_MM_dd_HH"
$now1 = get-date -format "yyyy_MM_dd_HHmm"
$fileName = -join ("monthlyPdfLog_", $now, ".txt")

# $publishFolderPath = "C:\services\Ruffer.PDFRecWebApp\dotnet-app\publish"
$errorLogFile = Join-Path "\\ruffer.local\dfs\Shared\PDFRec\" $fileName
try {
    Write-Host "publishFolderPath: $projFolder"
    Write-Host "inside ps1 symFileLoc: $symFileLoc, rdbFileLoc: $rdbFileLoc"
    
    if ($runType -eq "distribute") {       
        $distribute = $true
        Write-Host "distribute value is $distribute"
    }
    elseif ($runType -eq "archive") {
        $archiveDistribute = $true
    }

    $now = Get-Date
    $pdfLockFile = "\\ruffer.local\dfs\Shared\PDFRec\monthlyPdfLock.txt"
    
    #Create a lock file in the folder before starting the run
    if (Test-Path -Path $pdfLockFile) {
        Remove-Item -path $pdfLockFile -Force
        New-Item -path $pdfLockFile -type "file" -value "Initiated a pdf rec run for monthly at $($now.DateTime)"
    }
    else {
        New-Item -path $pdfLockFile -type "file" -value "Initiated a pdf rec run for monthly at $($now.DateTime)"
    }  
    $username = "ruffer\devadmin"
    $password = ConvertTo-SecureString -String "l0rdw3asel" -AsPlainText -Force 
    $cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, $password
    $PSEmailServer = 'outlook.ruffer.local'
    $servers = "pdcwkstst01", "pdcwkstst02", "pdcwkstst03", "pdcwkstst04"
    $emailTo = 'pprakash@ruffer.co.uk' 
    # $emailAttachmentTo = 'ishaikh@ruffer.co.uk', 'pprakash@ruffer.co.uk', 'FLalande@ruffer.co.uk', 'mcunningham@ruffer.co.uk'
    $emailAttachmentTo = 'pprakash@ruffer.co.uk'
    $numberOfProcesses = 5
    $jobs = [System.Collections.ArrayList]@()
    $serverCounter = 0;
    $startTime = get-date

    $distributeResultCommand = @"
$projFolder\ValuationReport.exe "distribute" "\\ruffer.local\dfs\Symphony\TestRun\Monthly Valuations\$symFileLoc" "\\ruffer.local\dfs\Symphony\TestRun\Monthly Valuations\$rdbFileLoc" "monthly" $day
"@
    $archiveDistributeResultCommand = @"
$projFolder\ValuationReport.exe "archive-distribute" "\\ruffer.local\dfs\Symphony\TestRun\Monthly Valuations\$symFileLoc" "\\ruffer.local\dfs\Symphony\TestRun\Monthly Valuations\$rdbFileLoc" "monthly" $day
"@
    $copyPdfRecApplicationCommand = "Robocopy.exe \\ruffer.local\dfs\Shared\PDFRec\CodeBackup\publish $publishFolderPath /E /w:1 /r:1"

    # $currentLocation = Get-Location

    # $publishDirectory = "$currentLocation\publish"


    # cd $publishDirectory
    #cd "..\dotnet-app\publish"
    # cd $publishFolderPath

    $valReportResult

    if ($archiveDistribute) {
        $valReportResult = invoke-command -ScriptBlock $( [ScriptBlock]::Create($archiveDistributeResultCommand) )
    }
    elseif ($distribute) {
        Write-Host "inside distribute loop"
        $valReportResult = invoke-command -ScriptBlock $( [ScriptBlock]::Create($distributeResultCommand) )
    }

    # $copyPdfRecApplicationSb = [ScriptBlock]::Create($copyPdfRecApplicationCommand)

    # if ($getlatest) {
    #     Invoke-Command -ScriptBlock $copyPdfRecApplicationSb
    # }

    # $distributeResult = invoke-command -ScriptBlock { .\ValuationReport.exe "distribute" "\\ruffer.local\dfs\Symphony\LiveRun\Email\Monthly Valuations\2019\06 - June\Business Day 2\SYMPHONY\2nd Batch Test" "\\ruffer-fs-02\DFS\Symphony\TestRun\Monthly Valuations\2019\JUN" "monthly" "BD2" }

    if (($valReportResult | Out-String).Contains("FATAL")) {
        Write-Host "FATAL ERROR!"
       
        New-Item -path $errorLogFile -type "file" -value `r`n"Timestamp: $valReportResult"
        return $valReportResult        
    }
    Clear-Variable -Name "valReportResult"
    
    foreach ($server in $servers) { 
        # if ($getlatest) {
        #     invoke-command -ComputerName $server -ScriptBlock  $copyPdfRecApplicationSb -Credential $cred -Authentication credssp 
        # }

        for ($counter = 1; $counter -le $numberOfProcesses; $counter++) {
            $partno = ($serverCounter * $numberOfProcesses) + $counter
            $batchfile = "& 'C:\PDFRec\MonthlyPDFRec - PART$partno.bat'"
            [ScriptBlock]$sb = [ScriptBlock]::Create($batchfile) 
            $job = invoke-command -ComputerName $server -ScriptBlock  $sb -Credential $cred -Authentication credssp -AsJob
            $jobs.Add($job)
        }
        $serverCounter++;
    }

    $result = "<html><body><h3>PDF REC RESULTS - $(Get-Date -Format 'f')</h3>"
    $attachmentEmailBody = "<html><body><h3>PDF REC RESULTS - $(Get-Date -Format 'f')</h3>"
    foreach ($job in $jobs) {
        Wait-Job -Job $job 
        # $jobOutput = Receive-Job -Job $job
        $result += "<div style=`"border:1px solid black; margin:20px; padding:10px;`">"
        $result += "<b>Command = $($job.Command.TrimStart('&'))</b><br>"
        $result += "<b>Location = $($job.Location)</b><br>"
        $result += "<p>$($job.ChildJobs[0].Output)</p>"
        $result += "<p style=`"color:red`" >$($job.ChildJobs[0].Error)</p>"
        $result += "</div>"
    }
    $now = get-date
    $result += "<p> start time : $($startTime.DateTime), end time : $($now.DateTime) </p>"
    $result += "<p>Total duration = $([Math]::Round($now.Subtract($startTime).TotalMinutes, 2)) minutes</p>"
    $result += "</body></html>"
    try {

        Write-Host "initiate consolidaion csv program"     
        $consolidationCSVCommand = @"
$projFolder\ValuationReport.exe "consolidateMonthlyCSV" "Monthly"
"@
        $consolidationResult = invoke-command -ScriptBlock $( [ScriptBlock]::Create($consolidationCSVCommand) )
         
    }
    catch {
        Write-Host ""+$error[0]
    }
    Write-Host "Completed run"
    Send-MailMessage -From 'pavan <pprakash@ruffer.co.uk>' -To $emailTo -Subject "Monthly Pdf Rec" -BodyAsHtml $result -Attachments "\\ruffer.local\dfs\Shared\PDFRec\Monthly\consolidated_results_Monthly.csv"
    Send-MailMessage -From 'pavan <pprakash@ruffer.co.uk>' -To $emailAttachmentTo -Subject "Monthly Pdf Rec Results" -BodyAsHtml $attachmentEmailBody -Attachments "\\ruffer.local\dfs\Shared\PDFRec\Monthly\consolidated_results_Monthly.csv"
    
    Write-Host "deleting lock file"
    # delete the monthlyPDFLock.txt file from shared drive
    Remove-Item -path $pdfLockFile -Force
    Write-Host "ps1- returning result"
    return $result
}
catch {
    Write-Host "inside catch block"
    # $errorLogFile = "\\ruffer.local\dfs\Shared\PDFRec\monthlyErrorLog.txt"
    if (Test-Path -Path $errorLogFile) {
        # Remove-Item -path $errorLogFile -Force
        $now1 = get-date -format "yyyy_MM_dd_HHmm"
        Add-Content -path $errorLogFile -value `r`n"Timestamp: $now1"`r`n$error[0]
        # New-Item -path $errorLogFile -type "file" -value $error[0]
    }
    else {
        New-Item -path $errorLogFile -type "file" -value $error[0]
    }
}
