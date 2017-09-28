Option Explicit

Dim Obj 

Set Obj = CreateObject("wscript.shell")

Obj.run "notepad.exe"
wscript.sleep 2000
Obj.sendkeys "Hello World!"
Obj.sendkeys "%fs"
wscript.sleep 500
obj.sendkeys "test.vbs"
wscript.sleep 300
obj.sendkeys "{enter}"