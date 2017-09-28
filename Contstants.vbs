dim message

message =Msgbox ("Pandora Radio",_ 
	vbAbortRetryIgnore+vbExclamation+vbDefaultButton3+vbSystemModal,_
	 "Gadget:")

if message = vbAbort then Msgbox "Quit", vbCritical, "Quit"
