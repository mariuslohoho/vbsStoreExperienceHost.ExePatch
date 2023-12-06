' Made By Marius '
Dim appTitle : appTitle = "Run Executables"
Dim oShell : Set oShell = CreateObject("WScript.Shell")
Dim pathUIP : pathUIP = ""
Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")

run = MsgBox("Hello, welcome back!,              " & vbnewline & "Wish to continue?", 4, appTitle)

If run = vbNo Then
	WScript.Quit
End If


Do Until fso.FileExists(pathUIP)
	pathUIP = InputBox("Enter the Path of the executable that you want to start", appTitle)
	If IsEmpty(pathUIP) Then
		WScript.Quit
	End If
	pathUIP = Replace(pathUIP,"""","")
Loop

runAsAdminUIP = MsgBox("Run as administrator? (UAC Bypass)" & vbnewline & "Note: This does not fully make it run as administrator",4,appTitle)

If runAsAdminUIP = vbYes Then
	oShell.Environment("Process").Item("__COMPAT_LAYER") = "RunAsInvoker"
End If

oShell.Run "cmd /c taskkill /f /im StoreExperienceHost.exe", 0
oShell.Run "cmd /c start """" """ & pathUIP & """", 0



