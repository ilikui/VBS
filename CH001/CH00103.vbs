option Explicit

Dim Greeting
Dim YourName
Dim TryAgain



Do
	TryAgain = "No"
	
	YourName = InputBox("Hello!What is your name?")
	
	If YourName = " " Then
		MsgBox = "You must enter your name to continue."
		
		TryAgain = "Yes"
			
	Else
		Greeting = "Hello "  & YourName & "! Nice to meet you."
		Msgbox Greeting
	End If


Loop While TryAgain = "Yes"












