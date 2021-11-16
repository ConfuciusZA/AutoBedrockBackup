Dim x : set x = createObject("wscript.shell")
Dim obj : set obj = GetObject("winmgmts:")
Dim FSO : Set FSO = CreateObject("Scripting.FileSystemObject")
Dim num : num = 1

Dim strFolder 
Dim serverFile : serverFile = "bedrock_server.exe"
Dim worldLoc : worldLoc = "worlds"
Dim backLoc : backLoc = "backups"

Set objFile = FSO.GetFile(Wscript.ScriptFullName)
strFolder = FSO.GetParentFolderName(objFile) 
Function ServerRunning 'Is the server running
	
	Dim ProL, pro, pror
	
	set ProL = obj.ExecQuery _
	("Select * from Win32_Process")
	
	For Each pro In ProL
		If pro.name = serverFile Then
			ServerRunning = 1
			Exit Function
		Else
			ServerRunning = 0
		End If
	Next
		
End Function


Function Backup 'copy world folder to the backup folder
	
	If ServerRunning = 1 Then
		Exit Function
	Else
		If Not (FSO.FolderExists(strFolder & "\" & backLoc)) then
			FSO.CreateFolder(strFolder & "\" & backLoc)
		End If
		
		Dim a : a = "\\" & Year(Date) & "-" & Month(Date) & "-" & Day(Date) & "_" & Time()
		a = Replace(a,":","")
		
		FSO.CreateFolder(strFolder & "\" & backLoc & a)

		FSO.CopyFolder worldLoc, (backLoc & a)
		wscript.sleep 1000
	End If
	
End Function

Function Start 'starts the server
	If ServerRunning = 1 Then
		Exit Function
	End If
	
	If num = 4 Then
		MsgBox "Failed to start server", , "Retries Exceeded"
		exit Function
	end If
	If (FSO.FileExists(serverFile)) Then
		x.run serverFile
		num = num + 1
		wscript.sleep 60000
	Else
		MsgBox "Server File Not Found", , "Error Server File"
		wscript.quit
	End If
	
	If ServerRunning = 0 Then
		Start
	End If
	
End Function

Function Announce 'Sends a five minute warning to all players then stops server after 5 minutes
	
	Dim q, t1, t2, t3, t4, t5, t7
	
	q = """"
	t1 = "tellraw @a {{}"
	t2 = "rawtext"
	t3 = ":{[}{{}"
	t4 = "text"
	t5 = ":"
	t7 = "{}}{]}{}}{ENTER}"
	
	x.AppActivate serverFile
	x.sendkeys t1 & q & t2 & q & t3 & q & t4 & q & t5 & q & "Server shutting down in 5 minutes." & q & t7
	wscript.sleep 240000
	If ServerRunning = 1 Then
		x.AppActivate serverFile
		x.sendkeys t1 & q & t2 & q & t3 & q & t4 & q & t5 & q & "Server shutting down in 1 minute." & q & t7
		wscript.sleep 30000
	End If
	If ServerRunning = 1 Then
		x.AppActivate serverFile 
		x.sendkeys t1 & q & t2 & q & t3 & q & t4 & q & t5 & q & "Starting Backup! Goodbye!" & q & t7
		x.sendkeys "stop{ENTER}"
		wscript.sleep 10000
	End If
	
End Function

Function Main 'stops server, copies world, restarts server
	
	num = 1
	
	If ServerRunning = 0 Then
		Backup
		Start
		wscript.quit
	Else
		Announce
		Backup
		Start
		wscript.quit
	End If

End Function

Main
