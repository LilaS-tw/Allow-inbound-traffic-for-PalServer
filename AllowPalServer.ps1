$programToMatch = "*PalServer-Win64-Shipping-Cmd.exe*"

$matchingApp = Get-NetFirewallApplicationFilter -PolicyStore ActiveStore | 
    Where-Object { $_.Program -like $programToMatch } | 
    Select-Object -First 1

if ($matchingApp) {
    $programPath = $matchingApp.Program
    Write-Output "Found program path: $programPath"
    
    $rulesToDelete = Get-NetFirewallRule -Direction Inbound | 
        Where-Object { $_.Enabled -eq 'True' -and (Get-NetFirewallApplicationFilter -PolicyStore ActiveStore -AssociatedNetFirewallRule $_ | Where-Object { $_.Program -eq $programPath }) }

    foreach ($rule in $rulesToDelete) {
        Write-Output "Deleting rule $($rule.DisplayName) with program $($programPath)"
        Remove-NetFirewallRule -Name $rule.Name -Confirm:$false
    }
    
    New-NetFirewallRule -DisplayName "PalServer-Win64-Shipping-Cmd.exe" -Direction Inbound -Program $programPath -Action Allow
} else {
    Write-Output "No matching program found."
}
