Option Explicit
Dim obj, a, b, c

set obj = CreateObject("wscript.shell")

a = MsgBox("Open a Program", vbYesNo+vbQuestion+vbSystemModal)

if a = vbYes  then 

	obj.Run "mspaint.exe"
	b = MsgBox("Open a Folder", vbYesNo+vbQuestion+vbSystemModal)

else

	b = MsgBox("Open a Folder", vbYesNo+vbQuestion+vbSystemModal)

end if

If b = vbYes Then 

	obj.run "C:\Users\nhilt\Desktop\vbscript"
	c = MsgBox("Open a File?", vbYesNo+vbQuestion+vbSystemModal)

Else

	c = MsgBox("Open a File?", vbYesNo+vbQuestion+vbSystemModal)

end If

If c = vbYes Then 

	obj.run "mkfile C:\Users\nhilt\Desktop\cool.txt"

else

	wscript.quit 
end if 