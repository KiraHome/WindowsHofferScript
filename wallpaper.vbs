dim shell
dim user

Set shell = WScript.CreateObject("WScript.Shell")
 user = shell.ExpandEnvironmentStrings("%USERNAME%")

Set fso = CreateObject("Scripting.FileSystemObject")

windowsDir = fso.GetSpecialFolder(0) 
 wallpaper = ".\wallpaper.jpg"

shell.RegWrite "HKCU\Control Panel\Desktop\Wallpaper", wallpaper
shell.RegWrite "HKCU\Control Panel\Desktop\WallpaperStyle", 6
shell.RegWrite "HKCU\Control Panel\Desktop\tilewallpaper", 0

shell.Run "C:\Windows\System32\rundll32.exe user32.dll,UpdatePerUserSystemParameters", 1, True
