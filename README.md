# List-COM-Ports
Windows PowerShell function to list all COM Port Devices and their physical USB port location. 

# Example Use
<img src="https://i.postimg.cc/mkbdCBXM/comsusb.png" width="50%">

# Requirements
You need to map out your physical USB ports to find their locations.
There are two locations per USB port. Some devices get mapped one way, other devices get mapped the other.
Each port of a downstream hub needs to be mapped.
You need to find a device of each type to fully map the ports.
Use Windows Device Manager to find the locations.
Devices used for mapping for type 1 any usb uart device. For type 2:Arduino Nano BLE/IOT/RP.

<img src="https://i.postimg.cc/9MxghzYM/comdev.png" width="50%">

