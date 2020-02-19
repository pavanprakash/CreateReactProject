param (
    $symFileLoc,
    $rdbFileLoc,
    $runType,
    $day,
    $option,
    $projFolder,
    $getlatest = $false
)
$now = get-date -format "yyyy_MM_dd_HH"
$now1 = get-date -format "yyyy_MM_dd_HHmm"
$fileName = -join ("quarterlyPdfLog_", $now, ".txt")
# $publishFolderPath = "C:\services\Ruffer.PDFRecWebApp\dotnet-app\publish"

$errorLogFile = Join-Path "\\ruffer.local\dfs\Shared\PDFRec\" $fileName
try {
    if ($runType -eq "distribute") {
        $distribute = $true
    }
    elseif ($runType -eq "archive") {
        $archiveDistribute = $true
    }
    #Create a lock file in the folder before starting the run
    $now = Get-Date
    $pdfLockFile = "\\ruffer.local\dfs\Shared\PDFRec\quarterlyPdfLock.txt"
    if (Test-Path -Path $pdfLockFile) {
        Remove-Item -path $pdfLockFile -Force
        New-Item -path $pdfLockFile -type "file" -value "Initiated a pdf rec run for quarterly at $($now.DateTime)"
    }
    else {
        New-Item -path $pdfLockFile -type "file" -value "Initiated a pdf rec run for quarterly at $($now.DateTime)"
    } 
    
    $username = "ruffer\devadmin"
    $password = ConvertTo-SecureString -String "l0rdw3asel" -AsPlainText -Force 
    $cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $username, $password
    $PSEmailServer = 'outlook.ruffer.local'
    $servers = "pdcwkstst01", "pdcwkstst02", "pdcwkstst03", "pdcwkstst04", "pdcwkstst05"
    $emailTo = 'pprakash@ruffer.co.uk' 
    #$emailAttachmentTo = 'pprakash@ruffer.co.uk', 'FLalande@ruffer.co.uk'
    $emailAttachmentTo = 'pprakash@ruffer.co.uk'
    $numberOfProcesses = 15
    $jobs = [System.Collections.ArrayList]@()
    $serverCounter = 0;
    $startTime = get-date

    $distributeResultCommand = @"
    $projFolder\ValuationReport.exe "distribute" "\\ruffer.local\DFS\Symphony\TestRun\Quarterly Valuations\$symFileLoc" "\\ruffer.local\DFS\Symphony\TestRun\Quarterly Valuations\$rdbFileLoc" "quarterly"
    "@
        $archiveDistributeResultCommand = @"
        $projFolder\ValuationReport.exe "archive-distribute" "\\ruffer.local\dfs\Symphony\TestRun\Quarterly Valuations\$symFileLoc" "\\ruffer.local\dfs\Symphony\TestRun\Quarterly Valuations\$rdbFileLoc" "quarterly"
"@   
    # $copyPdfRecApplicationCommand = "Robocopy.exe \\ruffer.local\dfs\Shared\PDFRec c:\pdfrec\publish /E /w:1 /r:1"

    
        

    # $currentLocation = Get-Location

    # $publishDirectory = "$currentLocation\publish"

    # cd $publishDirectory
    # cd "..\dotnet-app\publish"
    

    $valReportResult

    if ($distribute -and $archiveDistribute) {
        Write-Error -Message "Both parameters: distribute and archive distribute are set to true"
    }
    else {
        

        if ($distribute) {
            $valReportResult = invoke-command -ScriptBlock $( [ScriptBlock]::Create($distributeResultCommand) )
        }
        elseif ($archiveDistribute) {
            $valReportResult = invoke-command -ScriptBlock $( [ScriptBlock]::Create($archiveDistributeResultCommand) )
        }
        $copyPdfRecApplicationSb = [ScriptBlock]::Create($copyPdfRecApplicationCommand)

        if ($getlatest) {
            $valReportResult = Invoke-Command -ScriptBlock $copyPdfRecApplicationSb
        }

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
                $batchfile = "& 'C:\PDFRec\QuarterlyPDFRec - PART$partno.bat'"
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
        $resultFile = "\\ruffer.local\dfs\Shared\PDFRec\Quarterly\consolidated_results_Quarterly.csv"

        Write-Host "initiate consolidaion csv program"     
        $consolidationCSVCommand = @"
$projFolder\ValuationReport.exe "consolidateQuarterlyCSV" "Quarterly"
"@
        $consolidationResult = invoke-command -ScriptBlock $( [ScriptBlock]::Create($consolidationCSVCommand) )


        Write-Host "Completed run"
        Send-MailMessage -From 'pavan <pprakash@ruffer.co.uk>' -To $emailTo -Subject "Quarterly Pdf Rec" -BodyAsHtml $result -Attachments $resultFile
        Send-MailMessage -From 'pavan <pprakash@ruffer.co.uk>' -To $emailAttachmentTo -Subject "Quarterly Pdf Rec Results" -BodyAsHtml $attachmentEmailBody -Attachments $resultFile
        # delete the quarterlyPDFLock.txt file from shared drive
        Remove-Item -path $pdfLockFile -Force
        Write-Host "ps1- returning result"
        return $result
    }
}
catch {
    Write-Host "inside catch block"
    
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
