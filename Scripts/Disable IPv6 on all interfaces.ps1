# Disable IPv6 on all interfaces
Get-NetAdapter | ForEach-Object {
    Disable-NetAdapterBinding -Name $_.Name -ComponentID ms_tcpip6
}
