
:: #==============================================================#
:: #  LOCKDIR.BAT                                                 #
:: #==============================================================#
:: #  Lock and or unlock a specified directory from users access  #
:: #  Must be run as administrator.                               #
:: #                                                              #
:: #            Copyright(C) ZulNs, Yogyakarta, June 22'nd, 2013  #
:: #==============================================================#

@echo off
echo.

::checkAdminPermissions
net SESSION >nul 2>&1
if %ErrorLevel% == 0 goto :Begin
echo.Administrative privileges required!
echo.
exit /B 1

:Begin
echo. #=============================#
echo. #                             #
echo. #  FOLDER  LOCKER / UNLOCKER  #
echo. #                             #
echo. #         Copyright(C) ZulNs  #
echo. #  Yogyakarta, June 12, 2013  #
echo. #                             #
echo. #=============================#
echo.
echo.

if "%~f1" == "" goto :NoFolder
if not exist "%~f1" goto :NoExistingFolder
if "%~2" == "" goto :NoPassword

set "EXT=.{22877a6d-37a1-461a-91b0-dbda5aaebc99}"
set "LockedOwn=SYSTEM"
set "PwdFile=~ZulNs.sec"

if /I "%~x1" == "%EXT%" goto :Unlock

:Lock
takeown.exe /F "%~f1" /A
if %ERRORLEVEL% NEQ 0 goto :LockFail

icacls.exe "%~f1" /reset
if %ERRORLEVEL% NEQ 0 goto :LockFail

call :DelPwdFile "%~1"
echo.%~2> "%~f1\%PwdFile%"
if not exist "%~f1\%PwdFile%" goto :LockFail
call :SetPwdFile "%~1"

ren "%~f1" "%~nx1%EXT%"
if not exist "%~f1%EXT%" (
	call :DelPwdFile "%~1"
	goto :LockFail
)

icacls.exe "%~f1%EXT%" /setowner "%LockedOwn%"
if %ERRORLEVEL% NEQ 0 (
	ren "%~f1%EXT%" "%~nx1"
	call :DelPwdFile "%~1"
	goto :LockFail
)

icacls.exe "%~f1%EXT%" /inheritance:r
if %ERRORLEVEL% NEQ 0 (
	takeown.exe /F "%~f1%EXT%" /A
	ren "%~f1%EXT%" "%~nx1"
	call :DelPwdFile "%~1"
	goto :LockFail
)

echo.
echo. #==========#
echo. #          # 
echo. #  LOCKED  #
echo. #          #
echo. #==========#
echo.
call :ResetVar
exit /B 0

:Unlock
takeown.exe /F "%~f1" /A
if %ERRORLEVEL% NEQ 0 goto :UnlockFail

icacls.exe "%~f1" /reset
if %ERRORLEVEL% NEQ 0 (
	icacls.exe "%~f1" /setowner "%LockedOwn%"
	goto :UnlockFail
)

if not exist "%~f1\%PwdFile%" (
	icacls.exe "%~f1" /setowner "%LockedOwn%"
	icacls.exe "%~f1" /inheritance:r
	goto :NoSavedPassword
)

call :ResetPwdFile "%~1"
set /P SavedPwd=< "%~f1\%PwdFile%"
if "%SavedPwd%" == "" (
	call :SetPwdFile "%~1"
	icacls.exe "%~f1" /setowner "%LockedOwn%"
	icacls.exe "%~f1" /inheritance:r
	goto :NoSavedPassword
)

if "%~2" NEQ "%SavedPwd%" (
	call :SetPwdFile "%~1"
	icacls.exe "%~f1" /setowner "%LockedOwn%"
	icacls.exe "%~f1" /inheritance:r
	goto :WrongPassword
)

ren "%~f1" "%~n1"
if not exist "%~dpn1" (
	call :SetPwdFile "%~1"
	icacls.exe "%~f1" /setowner "%LockedOwn%"
	icacls.exe "%~f1" /inheritance:r
	goto :UnlockFail
)

call :DelPwdFile "%~dpn1"

echo.
echo. #============#
echo. #            #
echo. #  UNLOCKED  #
echo. #            #
echo. #============#
echo.
call :ResetVar
exit /B 0

:UnlockFail
echo.
echo  #=======================#
echo  #                       #
echo  #  UNLOCKING FAILED!!!  #
echo  #                       #
echo  #=======================#
echo.
call :ResetVar
exit /B 8

:LockFail
echo.
echo  #=====================#
echo  #                     #
echo  #  LOCKING FAILED!!!  #
echo  #                     #
echo  #=====================#
echo.
call :ResetVar
exit /B 7

:WrongPassword
echo.
echo. #========================#
echo. #                        #
echo. #  PASSWORD MISMATCH!!!  #
echo. #                        #
echo. #========================#
call :ResetVar
exit /B 6

:NoSavedPassword
echo.
echo. #========================#
echo. #                        #
echo. #  NO SAVED PASSWORD!!!  #
echo. #                        #
echo. #========================#
call :ResetVar
exit /B 5

:NoPassword
echo.
echo. #==================================#
echo. #                                  #
echo. #  YOU MUST PROVIDE A PASSWORD!!!  #
echo. #                                  #
echo. #==================================#
call :Usage
exit /B 4

:NoExistingFolder
echo.
echo. #==================================
echo. #
echo. #  FOLDER NOT FOUND: "%~f1"
echo. #
echo. #==================================
call :Usage
exit /B 3

:NoFolder
echo.
echo. #===========================#
echo. #                           #
echo. #  NO FOLDER TO PROCESS!!!  #
echo. #                           #
echo. #===========================#
call :Usage
exit /B 2

:Usage
echo.
echo.Usage:
echo.
set "UpCase=%~n0"
call :ToUpperCase
echo.%UpCase% ^<file_or_folder_to_lock_or_unlock^> ^<your_password^>
echo.
echo.Note:
echo.   Lock or unlock a specific folder from users access.
echo.   If the entered folder was in unlocked condition,
echo.   then it will be immediately lock and vice versa.
echo.   Requires administrative privileges.
echo.
goto :EOF

:SetPwdFile
attrib.exe +s +h +r "%~f1\%PwdFile%"
icacls.exe "%~f1\%PwdFile%" /setowner "%LockedOwn%"
icacls.exe "%~f1\%PwdFile%" /inheritance:r
goto :EOF

:ResetPwdFile
takeown.exe /F "%~f1\%PwdFile%" /A
icacls.exe "%~f1\%PwdFile%" /reset
attrib.exe -s -h -r "%~f1\%PwdFile%"
goto :EOF

:DelPwdFile
if not exist "%~f1\%PwdFile%" goto :EOF
call :ResetPwdFile "%~1"
del "%~f1\%PwdFile%"
goto :EOF

:ResetVar
set "SavedPwd="
set "PwdFile="
set "LockedOwn="
set "EXT="
goto :EOF

:ToUpperCase
set UpCase=%UpCase:a=A%
set UpCase=%UpCase:b=B%
set UpCase=%UpCase:c=C%
set UpCase=%UpCase:d=D%
set UpCase=%UpCase:e=E%
set UpCase=%UpCase:f=F%
set UpCase=%UpCase:g=G%
set UpCase=%UpCase:h=H%
set UpCase=%UpCase:i=I%
set UpCase=%UpCase:j=J%
set UpCase=%UpCase:k=K%
set UpCase=%UpCase:l=L%
set UpCase=%UpCase:m=M%
set UpCase=%UpCase:n=N%
set UpCase=%UpCase:o=O%
set UpCase=%UpCase:p=P%
set UpCase=%UpCase:q=Q%
set UpCase=%UpCase:r=R%
set UpCase=%UpCase:s=S%
set UpCase=%UpCase:t=T%
set UpCase=%UpCase:u=U%
set UpCase=%UpCase:v=V%
set UpCase=%UpCase:w=W%
set UpCase=%UpCase:x=X%
set UpCase=%UpCase:y=Y%
set UpCase=%UpCase:z=Z%
goto :EOF
