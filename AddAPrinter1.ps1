#run as Admin
#add the path ;)	

Get-WmiObject -Class Win32_Printer -ComputerName printserver | select Name, PortName, DriverName | Export-Csv -Path \\Path\Drucker.csv -NoClobber -Delimiter ","
# " still need to be removed