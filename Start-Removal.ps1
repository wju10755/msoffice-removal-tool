# Microsoft Office Removal Tool
$scriptUrl = "https://raw.githubusercontent.com/wju10755/msoffice-removal-tool/main/msoffice-removal-tool2.ps1"
$scriptPath = "C:\temp\remove-office.ps1"
$response = Invoke-WebRequest -Uri $scriptUrl -OutFile $scriptPath
if($response.StatusCode -eq 200){
    Start-Process -FilePath "powershell.exe" -ArgumentList "-File `"$scriptPath`""
} else {
    Write-Host "Failed to download the office removal script"
}