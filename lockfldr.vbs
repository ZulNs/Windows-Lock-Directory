
' #===============================================================#
' #  LOCKFLDR.VBS                                                 #
' #===============================================================#
' #  Lock and or unlock a specified directory from users access.  #
' #  This file will be run with admin rights.                     #
' #                                                               #
' #              Copyright(C) ZulNs, Yogyakarta, July 3'rd, 2013  #
' #===============================================================#

Option Explicit
Dim shell, quote
Set shell = CreateObject("Shell.Application")
quote = Chr(34)

If WScript.Arguments.Count = 0 Then
	shell.ShellExecute "wscript.exe", quote & WScript.ScriptFullName & _
			quote & " ~adm", "", "runas", 1
	Set shell = Nothing
	Wscript.Quit
End If

Const TITLE = "ZulNs: Folder Locker / Unlocker"
Const EXT = "{22877a6d-37a1-461a-91b0-dbda5aaebc99}"
Const PWDFILE = "~ZulNs.sec"
Const LOCKEDOWN = "SYSTEM"

'======================================================================
' Predefine
'======================================================================

If GetResponse("This will be lock or unlock a folder to disallow " _
		& "or allow users for accessing that folder." & vbCrLf & _
		vbCrLf & "Continue?") = vbNo Then
	Set shell = Nothing
	Wscript.Quit
End If

Dim fso, folder, source
Set fso = CreateObject("Scripting.FileSystemObject")

Do
	Set folder = shell.BrowseForFolder(0, _
			"Please choose a folder to lock or unlock:", 0, 0)
	If folder Is Nothing Then
		Set shell = Nothing
		AbortedMsg
		Wscript.Quit
	End If
	
	source = folder.Self.Path
	If Left(source, 3) = "::{" Or _
			( Right(source, 1) = "}" And _
			LCase(fso.GetExtensionName(source)) <> EXT ) Or _
			Right(source, 2) = ":\" Then
		If GetResponse("Can't process this following folder:" & _
				vbCrLf & vbCrLf & _
				quote & source & quote & vbCrLf & vbCrLf & _
				"Browse again?") = vbNo Then
			Set folder = Nothing
			Set shell = Nothing
			AbortedMsg
			Wscript.Quit
		End If
	Else
		Set folder = Nothing
		Set shell = Nothing
		Exit Do
	End If
Loop

Dim toLock, process, password, confirm

If LCase(fso.GetExtensionName(source)) = EXT Then
	toLock = False
	process = "UNLOCK"
Else
	toLock = True
	process = "LOCK"
End If

If GetResponse("Are you sure to " & process & _
		" the following folder?" & vbCrLf & vbCrLf & _
		quote & source & quote) = vbNo Then
	AbortedMsg
	Set fso = Nothing
	Wscript.Quit
End If

Do
	password = InputBox("Enter your password to " & process & _
			" this folder:", TITLE)
	If password = "" Then
		process = ""
		Exit Do
	End If

	confirm = InputBox("Re-enter your password for confirmation:", _
			TITLE)
	If confirm = "" Then
		process = ""
		Exit Do
	End If

	If password = confirm Then
		Exit Do
	End If

	If GetResponse("Mismatch for password and confirmation that " & _
			"you have just entered!!!" & vbCrLf & vbCrLf & _
			"Re-enter again?") = vbNo Then
		process = ""
		Exit Do
	End If
Loop

If process = "" Then
	AbortedMsg
	Set fso = Nothing
	Wscript.Quit
End If

'======================================================================
' Processing...
'======================================================================

Dim wsh, retCode, file, pwdFullName, content, dest
Set wsh = CreateObject("Wscript.Shell")
pwdFullName = source & "\" & PWDFILE

'On Error Resume Next

TakeOwner source
If retCode <> 0 Then
	ProcessFailed "Error taking owner!!!"
End If

ResetACLs source
If retCode <> 0 Then
	ProcessFailed "Error reseting ACLs!!!"
End If

If toLock Then
	' LOCK processing here...
	
	If Not SavePassword(pwdFullName, password) Then
		ProcessFailed "Error saving password!!!"
	End If
	
	DenyFileAccess pwdFullName
	
	dest = source & "." & EXT
	On Error Resume Next
	fso.MoveFolder source, dest
	If Not fso.FolderExists(dest) Then
		DelPwdFile pwdFullName
		ProcessFailed "The folder you want to lock may be opened by another program!!!"
	End If
	On Error Goto 0
	
	SetOwner dest, LOCKEDOWN
	If retCode <> 0 Then
		fso.MoveFolder dest, source
		DelPwdFile pwdFullName
		ProcessFailed "Error setting owner!!!"
	End If
	
	RemoveInheritance dest
	If retCode <> 0 Then
		TakeOwner dest
		fso.MoveFolder dest, source
		DelPwdFile pwdFullName
		ProcessFailed "Error removing inheritance permissions!!!"
	End If
Else
	' UNLOCK processing here...
	
	GrantFileAccess pwdFullName
	
	If Not VerifyPassword(pwdFullName, password) Then
		DenyFileAccess pwdFullName
		SetOwner source, LOCKEDOWN
		RemoveInheritance source
		ProcessFailed "Password mismatch or corrupted!!!"
	End If
	
	dest = fso.GetParentFolderName(source) & "\" & _
			fso.GetBaseName(source)
	On Error Resume Next
	fso.MoveFolder source, dest
	If Not fso.FolderExists(dest) Then
		DenyFileAccess pwdFullName
		SetOwner source, LOCKEDOWN
		RemoveInheritance source
		ProcessFailed "Error renaming folder!!!"
	End If
	On Error Goto 0
	
	DelPwdFile dest & "\" & PWDFILE
End If

InformationMsg "Congratulation." & vbCrLf & vbCrLf & _
		process & "ING process has been done."

Set wsh = Nothing
Set fso = Nothing
Wscript.Quit

'END here...
		
'======================================================================
' Sub-Routines
'======================================================================

Function ProcessFailed(msg)
	Set wsh = Nothing
	Set fso = Nothing
	FailedMsg msg
	Wscript.Quit
End Function

Function TakeOwner(dir)
	retCode = wsh.Run("takeown.exe /F " & quote & dir & quote & _
			" /A", 0, True)
End Function

Function SetOwner(dir, owner)
	retCode = wsh.Run("icacls.exe " & quote & dir & quote & _
			" /setowner " & quote & owner & quote, 0, True)
End Function

Function ResetACLs(dir)
	retCode = wsh.Run("icacls.exe " & quote & dir & quote & _
			" /reset", 0, True)
End Function

Function RemoveInheritance(dir)
	retCode = wsh.Run("icacls.exe " & quote & dir & quote & _
			" /inheritance:r", 0, True)
End Function

Function SetFileAttrib(fName)
	retCode = wsh.Run("attrib.exe +s +h +r " & _
			quote & fName & quote, 0, True)
End Function

Function ResetFileAttrib(fName)
	retCode = wsh.Run("attrib.exe -s -h -r " & _
			quote & fName & quote, 0, True)
End Function

Function SavePassword(fName, strPwd)
	SavePassword = False
	If fso.FileExists(fName) Then
		If Not DelPwdFile(fName) Then Exit Function
	End If
	Dim file
	On Error Resume Next
	Set file = fso.OpenTextFile(fName, 2, True)
	file.Write Encrypt(strPwd)
	file.Close
	set file = Nothing
	If Err.Number = 0 And VerifyPassword(fName, strPwd) Then _
			SavePassword = True
	On Error Goto 0
End Function

Function VerifyPassword(fName, strPwd)
	VerifyPassword = False
	If Not fso.FileExists(fName) Then Exit Function
	Dim file, content
	On Error Resume Next
	Set file = fso.OpenTextFile(fName, 1)
	If file.AtEndOfStream Then
		content = ""
	Else
		content = file.ReadAll
	End If
	file.Close
	Set file = Nothing
	If Err.Number = 0 And Encrypt(strPwd) = content Then _
			VerifyPassword = True
	On Error Goto 0
End Function

Function DelPwdFile(fName)
	On Error Resume Next
	fso.DeleteFile fName, True
	If fso.FileExists(fName) Then
		GrantFileAccess fName
		fso.DeleteFile fName, True
	End If
	DelPwdFile = Not fso.FileExists(fName)
	On Error Goto 0
End Function

Function GrantFileAccess(fName)
	TakeOwner fName
	ResetACLs fName
	ResetFileAttrib fName
End Function

Function DenyFileAccess(fName)
	SetFileAttrib fName
	SetOwner fName, LOCKEDOWN
	RemoveInheritance fName
End Function	

Function Encrypt(str)
	Dim i, c, msd, lsd
	For i = 1 to Len(str)
		c = Asc(Mid(str, i, 1))
		msd = c \ 16
		lsd = c And 15
		If i Mod 2 = 0 Then
			msd = msd Xor 6
			lsd = lsd Xor 9
		Else
			msd = msd Xor 9
			lsd = lsd Xor 6
		End If
		c = (c And 96) Xor 32
		If c <= 32 Then c = c Or 64
		Encrypt = Encrypt & Chr(c Or lsd) & Chr(c Or msd) 
	Next
End Function

Function AbortedMsg()
	InformationMsg "Aborted."
End Function

Function FailedMsg(msg)
	CriticalMsg msg & vbCrLf & _
			"Can't continue " & process & "ING process."
End Function

Function InformationMsg(msg)
	ShowMsg msg, vbInformation
End Function

Function ExclamationMsg(msg)
	ShowMsg msg, vbExclamation
End Function

Function CriticalMsg(msg)
	ShowMsg msg, vbCritical
End Function

Function GetResponse(msg)
	GetResponse = ShowMsg(msg, vbYesNo + vbQuestion)
End Function

Function ShowMsg(msg, intStyle)
	ShowMsg = MsgBox(msg, intStyle, TITLE)
End Function
