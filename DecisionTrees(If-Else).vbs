Option Explicit
Dim birthday 

birthday = Msgbox ("Is it your birthday?", vbYesNo+vbQuestion, "Hello world")
If birthday = vbYes then 
		Msgbox "Happy Birthday", vbInformation
Else
		Msgbox "Sorry, I hope you have a great day!", vbCritical, "Oops"
End If 