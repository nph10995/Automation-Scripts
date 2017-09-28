Option Explicit
'Syntax
'InputBox "message", "title", "InputField", xposition, yposition
Dim input 



input = InputBox( "What is your name?", "Information", "Name Goes Here.")


msgbox input, vbOKOnly, "This is your name."
