' 
' WMI Script - ChangeVLKey.vbs
'
' This script changes the product key on the computer
'
'***************************************************************************

ON ERROR RESUME NEXT

if Wscript.arguments.count<1 then
   Wscript.echo "Script can't run without VolumeProductKey argument"
   Wscript.echo "Correct usage: Cscript ChangeVLKey.vbs ABCDE-FGHIJ-KLMNO-PRSTU-WYQZX"
   Wscript.quit
end if

Dim VOL_PROD_KEY
VOL_PROD_KEY = Wscript.arguments.Item(0)
VOL_PROD_KEY = Replace(VOL_PROD_KEY,"-","") 'remove hyphens if any
Dim WshShell
Set WshShell = WScript.CreateObject("WScript.Shell")
'WshShell.RegDelete "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\WPAEvents\OOBETimer" 'delete OOBETimer registry value
for each Obj in GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf ("win32_WindowsProductActivation")

   result = Obj.SetProductKey (VOL_PROD_KEY)

   if err <> 0 then
      WScript.Echo Err.Description, "0x" & Hex(Err.Number)
      Err.Clear
   end if

Next 
