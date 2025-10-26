' Constants for sound
Const SND_SYNC = &H0
Const SND_ASYNC = &H1
Const SND_FILENAME = &H20000

' Connect to the WMI service
set oLocator = CreateObject("WbemScripting.SWbemLocator")
set oServices = oLocator.ConnectServer(".","root\wmi")

' Get the full charged capacity of the battery
set oResults = oServices.ExecQuery("select * from batteryfullchargedcapacity")
for each oResult in oResults
   iFull = oResult.FullChargedCapacity
next

' Start monitoring battery status
while (1)
  set oResults = oServices.ExecQuery("select * from batterystatus")
  for each oResult in oResults
    iRemaining = oResult.RemainingCapacity
    bCharging = oResult.Charging
  next

  iPercent = ((iRemaining / iFull) * 100) mod 100

  ' Check if battery is charging and above 95%
  if bCharging and (iPercent > 95) Then
    ' Play the recorded WAV file using Windows Media Player
    set objShell = CreateObject("WScript.Shell")
    objShell.Run "wmplayer.exe /play /close ""C:\Scripts\BatteryAlert96.wav"""
    
    ' Show message box
    msgbox "Battery is at " & iPercent & "%", vbInformation, "Battery monitor"
  end if

  wscript.sleep 30000 ' 30 seconds
wend
