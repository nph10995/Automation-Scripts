Option Explicit

Dim a

a = MsgBox("Guess a Button", vbAbortRetryIgnore)
if a = vbRetry then 
		MsgBox "Correct!"
elseif a = vbIgnore then 
		MsgBox "Wrong"
else
		MsgBox "Wrong!"
end if