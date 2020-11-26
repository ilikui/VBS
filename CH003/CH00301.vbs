option Explicit

dim fs, driver

' get access to the ActiveX object:

set fs = CreateObject("Scripting.FileSystemObject")

set driver = fs.GetDrive("C:\")

MsgBox "Aviable Space on C:\:" & FormatNumber(driver.AvailableSpace/1024^3) & "G"

