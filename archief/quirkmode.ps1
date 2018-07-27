#Find Arduino COM Port
#$PortName = (Get-WmiObject Win32_SerialPort | Where-Object { $_.Name -match "USB Serial Port"}).DeviceID
#if ( $PortName -eq $null ) { throw "XBee Not Found"}
#Create SerialPort and Configure
$port = New-Object System.IO.Ports.SerialPort
$port.PortName = "COM20"
$port.BaudRate = "38400"
$port.Parity = "None"
$port.DataBits = 8
$port.StopBits = 1
$port.ReadTimeout = 1000 #Milliseconds
$port.open() #open serial connection
Start-Sleep -Milliseconds 100 #wait 0.1 seconds
$port.Write("4") #write $byte parameter content to the serial connection
try    {

#Check for response
if (($response = $port.ReadLine()) -gt 0)
{ $response }
}
catch [TimeoutException] {
#Timeout
}
finally    {
$port.Close() #close serial connection
}