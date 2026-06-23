$XLSX = "./BluestreakActiveUsers.xlsx"
$XLSPath = "./BluestreakActiveUsers.xls"

If (-not (Test-Path $XLSX)) {
    If (Test-Path $XLSPath) {
        ConvertTo-ExcelXlsx -Path $XLSPath
        #Remove-Item -Path $XLSPath
    } Else {
        Write-Host "File" ($XLSPath.Substring(2)) "not found!" -BackgroundColor Red
        Exit 1
    }
}

$BluestreakUsers = Import-Excel -Path $XLSX -ImportColumns 2 | Select-Object -ExpandProperty "Full Name"

$TrimmedBluestreakUsers = @()

ForEach ($BluestreakUser in $BluestreakUsers) {
    $TrimmedBluestreakUsers += ($BluestreakUser -replace '\s+', ' ').Trim()
}

Connect-MgGraph -Scopes "User.Read.All"

$365Users = @(Get-MgUser -All | Select-Object -ExpandProperty DisplayName)

$Trimmed365Users = @()

ForEach ($365User in $365Users) {
    $Trimmed365Users += ($365User -replace '\s+', ' ').Trim()
}

ForEach ($Trimmed365User in $Trimmed365Users) {
    If (-not ($TrimmedBluestreakUsers -contains $Trimmed365User)) {
        Write-Host $Trimmed365User "is present in Microsoft 365 but inconsistent with Bluestreak!" -BackgroundColor Green
    }
}

ForEach ($TrimmedBluestreakUser in $TrimmedBluestreakUsers) {
    If (-not ($Trimmed365Users -contains $TrimmedBluestreakUser)) {
        Write-Host $TrimmedBluestreakUser "is present in Bluestreak but inconsistent with Microsoft 365!" -BackgroundColor Blue
    }
}