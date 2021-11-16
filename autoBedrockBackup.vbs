Dim x : set x = createObject("wscript.shell")
Dim obj : set obj = GetObject("winmgmts:")
Dim FSO : Set FSO = CreateObject("Scripting.FileSystemObject")
Dim num : num = 1

Dim strFolder 
Dim serverFile : serverFile = "bedrock_server.exe" 'Server executable file. Default: bedrock_server.exe
Dim worldLoc : worldLoc = "worlds" 'folder name where worlds are
Dim backLoc : backLoc = "backups" 'folder name for backups
Dim backupsToKeep: backupsToKeep = 5 'number of backups to keep

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

		FSO.CopyFolder worldLoc, (strFolder & "\" & backLoc & a)
		
		If (FSO.GetFolder(strFolder & "\" & backLoc).Subfolders.Count) > backupsToKeep then				
			ClearExcessBackups
		End If
		
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

Sub ClearExcessBackups
	Dim f, a(), sf, i, j, p, d

	Set f = FSO.GetFolder(strFolder & "\" & backLoc)
	
	If f.SubFolders.Count <= backupsToKeep Then Exit Sub
	ReDim a(1, f.SubFolders.Count - 1)
	For Each sf In f.SubFolders
	  a(0, i) = sf.Path
	  a(1, i) = sf.DateCreated
	  i = i + 1
	Next
	For i = 0 To UBound(a, 2) - 1
	  For j = i + 1 To UBound(a, 2)
		If a(1, i) < a(1, j) Then
		  p = a(0, j): d = a(1, j)
		  a(0, j) = a(0, i): a(1, j) = a(1, i)
		  a(0, i) = p: a(1, i) = d
		End If
	  Next
	Next
	For i = backupsToKeep To UBound(a, 2)
	  FSO.DeleteFolder a(0, i), True
	Next
End Sub


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
