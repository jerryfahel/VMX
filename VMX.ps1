# Params #####################

$hostFile = "LOCATION_OF_HOSTS.CSV"
$logDir = "LOCATION_OF_LOGS"
$fromEmail = "Name <name@email.com>"
$toEmail = "Name <name@email.com>"
$smtpServer = "smtp.email.com"

# Main Script Body ###########

$hosts = Import-Csv $hostFile
foreach ($line in $hosts) {
	$guest = $($line.GuestName)
	$source = $($line.SourceDir)
	$dest = $($line.DestDir)
	$waitstart = $($line.WaitStart)
	$waitshutdown = $($line.WaitShutdown)
	$logfile =  $logDir
	$logfile += Split-Path $dest -Leaf
	$logfile += ".log"
	$vm = gwmi -namespace root\virtualization\v2 -query "select * from msvm_computersystem where elementname='$guest'"
	$vmname = $vm.name
	if(!$vmname) {
		write-host "$guest is not hosted here" -foregroundcolor red
		write-host ""
	} else {
		write-host ""
		write-host "$guest is hosted here" -foregroundcolor green
		write-host ""
		write-host "$guest is shutting down"
		$vmshut = gwmi -namespace root\virtualization\v2 -query "SELECT * FROM Msvm_ShutdownComponent WHERE SystemName='$vmname'"
		$result = $vmshut.InitiateShutdown("$true","no comment")
		if ($result.returnvalue -match "0") {
			start-sleep -s $waitshutdown
			write-host ""
			write-host "$guest has been shutdown" -foregroundcolor green
			write-host ""
			Send-MailMessage -To "$toEmail" -From "$fromEmail" -Subject "VM Export - $guest shutdown" -SmtpServer "$smtpServer"
			robocopy $source $dest /MIR /LOG:$logFile
			Send-MailMessage -To "$toEmail" -From "$fromEmail" -Subject "VM Export - $guest exported" -Body "See Log: $logFile" -SmtpServer "$smtpServer"
			write-host "$guest is starting up"
			$result = $vm.requeststatechange(2)
			if ($result.returnvalue -match "0") {
				start-sleep -s $waitstart
				write-host ""
				write-host "$guest has started" -foregroundcolor green
				write-host ""
				Send-MailMessage -To "$toEmail" -From "$fromEmail" -Subject "VM Export - $guest started" -SmtpServer "$smtpServer"
			} else {
				write-host ""
				write-host "$guest unable to start" -foregroundcolor red
				write-host ""
				Send-MailMessage -To "$toEmail" -From "$fromEmail" -Priority High -Subject "VM Export - $guest unable to start" -SmtpServer "$smtpServer"
			}
		} else {
			write-host ""
			write-host "$guest unable to shutdown" -foregroundcolor red
			write-host ""
			Send-MailMessage -To "$toEmail" -From "$fromEmail" -Priority High -Subject "VM Export - $guest unable to shutdown" -SmtpServer "$smtpServer"
		}
	}
}
