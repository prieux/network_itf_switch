computer = "."
cableConnectionName = "Connexion au réseau local"
wifiConnectionName = "Connexion réseau sans fil"
set objWMIService = GetObject("winmgmts:\\" & computer & "\root\cimv2" )
set colAdapters = objWMIService.Execquery("Select * from Win32_NetworkAdapter" )
for each Adapter in colAdapters
    if Adapter.NetConnectionID = cableConnectionName then
        set cableAdapter = Adapter
    end if
    if Adapter.NetConnectionID = wifiConnectionName then
        set wifiAdapter = Adapter
    end if
next
WScript.Echo("/nid:" & cableAdapter.NetConnectionID & "/nom:" & cableAdapter.Name & "/enabled:" & cableAdapter.NetEnabled)
WScript.Echo("/nid:" & wifiAdapter.NetConnectionID & "/nom:" & wifiAdapter.Name & "/enabled:" & wifiAdapter.NetEnabled)
if (cableAdapter.NetEnabled = "Vrai") and (wifiAdapter.NetEnabled = "Faux") then
    wifiAdapter.Enable()
    cableAdapter.Disable()
    WScript.Echo("Wifi connection enabled")
else
    wifiAdapter.Disable()
    cableAdapter.Enable()
    WScript.Echo("Cable connection enabled")
end if
