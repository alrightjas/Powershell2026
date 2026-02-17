Get-ChildItem -Path C:\ -Recurse -Filter *.log

$session = New-PSSession -ComputerName "WIN-1OCJVG8RKKR.jas-dejesus.local" -Credential jas-dejesus.local\Administrator
Enter-PSSession -ComputerName "WIN-1OCJVG8RKKR.jas-dejesus.local" -Credential jas-dejesus.local\Administrator

$import = Import-Excel "C:\Users\Administrator\Downloads\Lab5-3.xlsx" -WorksheetName "Processes" -ImportColumns @(1,46,12,6) -StartRow 1 -EndRow 7
$import | Export-Excel C:\Users\Administrator\Downloads\Processes.xlsx

Read-Host Input first name:
Read-Host Input last name:

Get-Content "C:\Users\Administrator\Documents\Example.txt" | ForEach-Object {
    Write-Output "Processing line: $_"
}

Get-Process | Select-Object Name, ProcessName, CPU

Enter-PSSession -ComputerName "WIN-1OCJVG8RKKR.jas-dejesus.local" -Credential jas-dejesus.local\Administrator
Get-NetIPAddress -AddressFamily IPv4 -InterfaceAlias "Ethernet"

