
' #===============================================================#
' #  LOCKDIRW.VBS                                                 #
' #===============================================================#
' #  Lock and or unlock a specified directory from users access.  #
' #  Required LOCKDIR.BAT file.                                   #
' #  That's file will be run with admin rights.                   #
' #                                                               #
' #             Copyright(C) ZulNs, Yogyakarta, June 22'nd, 2013  #
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
Const EXT = ".{22877a6d-37a1-461a-91b0-dbda5aaebc99}"
Const batFileName = "lockdir.bat"

'======================================================================
' Predefine
'======================================================================

If GetResponse("This will be lock or unlock a folder to disallow " _
		& "or allow users for accessing that folder." & vbCrLf & _
		vbCrLf & "Continue?") = vbNo Then
	Set shell = Nothing
	Wscript.Quit
End If

Dim folder, target

Do
	Set folder = shell.BrowseForFolder(0, _
			"Please choose a folder to lock or unlock:", 0)
	If folder Is Nothing Then
		Set shell = Nothing
		AbortedMsg
		Wscript.Quit
	End If
	
	target = folder.Self.Path
	If Left(target, 3) = "::{" Or _
			( Right(target, 1) = "}" And _
			LCase(Right(target, 39)) <> EXT ) Or _
			Right(target, 2) = ":\" Then
		If GetResponse("Can't process this following folder:" & _
				vbCrLf & vbCrLf & _
				quote & target & quote & vbCrLf & vbCrLf & _
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

If LCase(Right(target, 39)) = EXT Then
	toLock = False
	process = "UNLOCK"
else
	toLock = True
	process = "LOCK"
End If

If GetResponse("Are you sure to " & process & _
		" the following folder?" & vbCrLf & vbCrLf & _
		quote & target & quote) = vbNo Then
	AbortedMsg
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
	Wscript.Quit
End If

'======================================================================
' Processing...
'======================================================================

Dim fso, path

set fso = CreateObject("Scripting.FileSystemObject")
path = fso.GetParentFolderName(Wscript.ScriptFullName)
If Not fso.FileExists(path & "\" & batFileName) Then
	ShowMsg "Can't find " & quote & batFileName & quote & _
			" file!!!" & vbCrLf & vbCrLf & _
			process & "ING process has been aborted.", vbCritical
	set fso = Nothing
	Wscript.Quit
End If

Dim wsh, retCode

Set wsh = CreateObject("Wscript.Shell")
retCode = wsh.Run("cmd.exe /C " & quote & _
		quote & path & "\" & batFileName & quote & " " & _
		quote & target & quote & " " & _
		quote & Encrypt(password) & quote & quote, 0, True)
Set wsh = Nothing

If retCode = 0 Then
	ShowMsg "Congratulation." & vbCrLf & vbCrLf & _
			process & "ING process has been completed.", vbInformation
ElseIf retCode = 6 Then
	ShowMsg "Password mismatch." & vbCrLf & vbCrLf & _
			process & "ING process has been aborted.", vbCritical
Else
	ShowMsg "Found an error." & vbCrLf & vbCrLf & _
			process & "ING process has been aborted.", vbExclamation
End If

'END here...
		
'======================================================================
' Sub-Routines
'======================================================================

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
	ShowMsg "Aborted.", vbInformation
End Function

Function ShowMsg(strPrompt, intStyle)
	CommonMsg strPrompt, vbOKOnly + intStyle
End Function

Function GetResponse(strPrompt)
	GetResponse = CommonMsg(strPrompt, vbYesNo + vbQuestion)
End Function

Function CommonMsg(strPrompt, intStyle)
	CommonMsg = MsgBox(strPrompt, intStyle, TITLE)
End Function
