echo off

rem  the path of the git command


	set git = D:\Git\bin\git.exe

    set y=%date:~0,2%	
	set mo=%date:~5,2%
	set d=%date:~8,2%
    set wek =%date:~10,6%  
	set h=%time:~0,2%
	set mi=%time:~3,2%	
	set sec = %time:~6,2%	
    set milsec = %time:~9,2%

	

	git add *

	REM delay 3 second
	TIMEOUT /T 3	 
	git commit -m  "update"
	REM delay 5 second
	TIMEOUT /T 5	
	git push

	git status

pause

@echo off

