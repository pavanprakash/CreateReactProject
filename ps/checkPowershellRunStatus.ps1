param(
  $folder
)

if ($folder -eq "monthly") {
  $fileName = "monthlyPdfLock.txt"
}
elseif ($folder -eq "quarterly") {
  $fileName = "quarterlyPdfLock.txt"
}
$pdfRecHomeFolder = "\\ruffer.local\dfs\Shared\PDFRec\$fileName"
$now = Get-Date
$check

if (Test-Path -Path $pdfRecHomeFolder) {
  Add-Content -path $pdfRecHomeFolder -value "User attempted to initiated another run while current one is running at $($now.DateTime)"
  $check = "RUNNING..."
}
else { 
  
  $check = "NOT RUNNING..."
}

return $check