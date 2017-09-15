dim shell
dim user

Set shell = WScript.CreateObject("WScript.Shell")
 user = shell.ExpandEnvironmentStrings("%USERNAME%")

Set fso = CreateObject("Scripting.FileSystemObject")

windowsDir = fso.GetSpecialFolder(0) 
 wallpaper = "C:\Users\212585590\Desktop\wallpaper.jpg"

shell.RegWrite "HKCU\Control Panel\Desktop\Wallpaper", wallpaper

shell.Run "C:\Windows\System32\rundll32.exe user32.dll,UpdatePerUserSystemParameters", 1, True
