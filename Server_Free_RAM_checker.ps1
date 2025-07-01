$VMHosts = "Host1", "Host2", "Host3"  # Replace with actual host names



$VMHosts | % { Invoke-Command -ComputerName $_ -ScriptBlock {
    $memory = ((Get-CIMInstance Win32_OperatingSystem).FreePhysicalMemory / 1MB)
    $host_x = $env:computername
    "$memory,$host_x"
  }
}