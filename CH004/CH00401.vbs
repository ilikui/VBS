option Explicit


dim result

result = MsgBox("Do you want to continue?",vbYesNo)

if result = vbYes  then
	MsgBox "We'll continue..."
else
	MsgBox "Ok,enough..."
end if

