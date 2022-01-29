''' Random snips of VBS stuff as I can be bothered to add them here for later reference
'''

''' ----- Add a mapped drive.

ON ERROR RESUME NEXT

Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run "Explorer " & "\\location.abc.com\WORKDATA", 1, false ' Change that to the right path.
wScript.Sleep 3000
''' Above:: Opens the folder path in Explorer first, sometimes this fixes issues 
''' where Windows claims it can't find the path. 

Dim objNetwork, strRemoteShare
Set objNetwork = WScript.CreateObject("WScript.Network")
strRemoteShare = "\\location.abc.com\WORKDATA" ' Change this to the right path. 
objNetwork.RemoveNetworkDrive "W:", True ' Removes the drive first, change this to the right drive letter
wScript.Sleep 1000
objNetwork.MapNetworkDrive "W:", strRemoteShare, True ' Change this to the right drive letter. 
''' Above:: Maps the drive persistently.
''' CHANGE the path above for strRemoteShare!

message = msgbox("Your drive should be mapped. Please close the Explorer window and re-open it, and test. Service Desk: 555 123 456" ,64, "Drives mapped.")

''' --------------------------------------------------------------------

''' ----- Add a printer

ON ERROR RESUME NEXT

Set WSHNetwork = CreateObject("WScript.Network")
WSHNetwork.AddWindowsPrinterConnection "\\location.com\CAN1PR014"
WSHNetwork.SetDefaultPrinter "\\location.com\CAN1PR014"
