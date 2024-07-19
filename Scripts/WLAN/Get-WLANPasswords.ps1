$Profiles = netsh wlan show profile|where {$_ -match "profile "}
Foreach($Profile in $Profiles -replace "    All User Profile     : "){
    $WLAN = netsh wlan show profile name=$Profile key=clear|where{$_ -match "Key Content            : "} 
    Write-Output "$Profile password: $($wlan -replace '    Key Content            : ')"
}
