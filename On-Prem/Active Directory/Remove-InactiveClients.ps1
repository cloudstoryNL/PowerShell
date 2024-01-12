$oldClients = Get-ADComputer -Filter * -Properties * | ? { (($_.OperatingSystem -match 'Windows 10') -or ($_.OperatingSystem -match 'Windows 11')) -and ($_.LastLogonDate -lt (Get-Date).AddDays(-180)) }

Foreach($computer in $oldClients)
{
    Remove-ADObject -Identity $computer.DistinguishedName -Recursive -Confirm:$false
    'Deleted {0}' -f $computer.Name | Write-Host -ForegroundColor Yellow
}