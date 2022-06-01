#add to profile.ps1

function coms{
	#This function finds all USB COM Port Devices (RS-232, UART, Arduinos, etc) and gives their name and which physical USB port they are connected to.
	<# 
	Example Output:
		Current COM Ports________________________________________
		COM8      USB2     Arduino Uno
		COM9      USB3     Silicon Labs CP210x USB to UART Bridge
		COM11     USB4     USB-SERIAL CH340
	#>
	#USER MUST CONFIGURE THESE____________
	#Use the Windows device manager to locate the following. Each port has two possible locations depending on the type of device attached to it. 
	#You'll need to put the devices in each port and map them out. Downstream hubs need to be mapped out if desired.
	$usb_ports =@(
		[pscustomobject]@{user_defined_name = "USB0";   port_location = "Port_#00xx.Hub_#00xx"; pci_location = "0000.00xx.0000.0xx.00x.00x.000.000.000";},
		[pscustomobject]@{user_defined_name = "USB1";   port_location = "Port_#00xx.Hub_#00xx"; pci_location = "0000.00xx.0000.0xx.00x.00x.000.000.000";},
		[pscustomobject]@{user_defined_name = "USB2";   port_location = "Port_#00xx.Hub_#00xx"; pci_location = "0000.00xx.0000.0xx.00x.00x.000.000.000";},
		[pscustomobject]@{user_defined_name = "USB3";   port_location = "Port_#00xx.Hub_#00xx"; pci_location = "0000.00xx.0000.0xx.00x.00x.000.000.000";},
		[pscustomobject]@{user_defined_name = "USB4";   port_location = "Port_#00xx.Hub_#00xx"; pci_location = "0000.00xx.0000.0xx.00x.00x.000.000.000";},
		[pscustomobject]@{user_defined_name = "USB5";   port_location = "Port_#00xx.Hub_#00xx"; pci_location = "0000.00xx.0000.0xx.00x.00x.000.000.000";},
		[pscustomobject]@{user_defined_name = "USB6";   port_location = "Port_#00xx.Hub_#00xx"; pci_location = "0000.00xx.0000.0xx.00x.00x.000.000.000";},
		[pscustomobject]@{user_defined_name = "USB1.1"; port_location = "Port_#00xx.Hub_#00xx"; pci_location = "0000.00xx.0000.0xx.00x.00x.000.000.000";},
		[pscustomobject]@{user_defined_name = "USB1.2"; port_location = "Port_#00xx.Hub_#00xx"; pci_location = "0000.00xx.0000.0xx.00x.00x.000.000.000";},
		[pscustomobject]@{user_defined_name = "USB1.3"; port_location = "Port_#00xx.Hub_#00xx"; pci_location = "0000.00xx.0000.0xx.00x.00x.000.000.000";})
	#_____________________________________
	
	$usb_devices = New-Object System.Collections.ArrayList
	
	#get a list of all devices with 'COM' in their name
	$device_query = Get-PnpDevice -PresentOnly | Where-Object { $_.FriendlyName -match '\(COM' }

	ForEach ($device in $device_query) {
		#get the USB port location from the registry (ex: Port_#0003.Hub_#0016 or 0000.0014.0000.004.001.003.000.000.000)
		$registry = 'HKLM:\SYSTEM\CurrentControlSet\Enum\' + $device.InstanceId
		$usb_location = (Get-ItemPropertyValue -Path $registry -Name LocationInformation | Out-String).Trim()
		
		#get the more descriptive bus reported description
		$bus_description = $((Get-WMIObject Win32_PnPEntity | where {$_.name -match $device.description}).GetDeviceProperties("DEVPKEY_Device_BusReportedDeviceDesc").DeviceProperties.Data)
		
		$com_port = $device.FriendlyName -replace '.+?(?=\(COM)' -replace'[\(\)]'
		$usb_port_name = "X" #default unknown port name
		
		#match each USB port location to a user defined USB port name
		ForEach ($usb_port in $usb_ports) {
			If ($usb_location -eq $usb_port.port_location -Or $usb_location -eq $usb_port.pci_location){
				$usb_port_name = $usb_port.user_defined_name	
			}
		}	
		$usb_devices.Add([pscustomobject]@{com_port = $com_port; description = $device.description; bus_description = $bus_description; usb_location = $usb_location; usb_port_name = $usb_port_name; InstanceId = $device.InstanceId}) | out-null
	}

	If ($usb_devices.Count -lt 1) {
		echo "No COM devices found"
		
	} Else {
		echo "Current COM Ports   _______________________________________"
		ForEach ($device in $usb_devices){
		echo ("  " + $device.com_port.PadRight(9,' ')  +  $device.usb_port_name.PadRight(12,' ')  + $device.bus_description)
		}
	}
}