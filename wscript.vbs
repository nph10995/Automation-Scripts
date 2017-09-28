Option Explicit

dim a 
dim b 
dim nameLength

a = wscript.scriptname
b = wscript.scriptfullname
nameLength = len(b)-len(a)

Msgbox (Left(b, nameLength))


