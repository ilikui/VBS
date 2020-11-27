option Explicit
 
 
dim age

dim name

dim birthdate




call  OutPutData(age)


call  OutPutData(Name)

call  OutPutData(birthdate)







'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'
'
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
function OutPutData(data)

    dim  dataflg
	
	do
		data  = InputBox("Please input your information?")
		
		dataflg = not(IsEmpty(data))
		
	
		if len(data) =  0 then
			dataflg = true
			
		else
			if dataflg then
			
			msgbox "You enter data is "  & data
			dataflg = false
			else
			msgbox "datais null,please check it"
			
			end if
		
		end if
		
	loop  while  dataflg



end function 
