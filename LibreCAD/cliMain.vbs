Option Explicit

Sub CmdMain
	Dim fso
	Dim WshShell
	Dim objArgs
	
	Dim scrFile

	Dim startPath
	Dim MyDocPath
	Dim CADDscr

	Set fso = CreateObject("Scripting.FileSystemObject")
	Set WshShell = CreateObject("WScript.Shell")
	Set objArgs = WScript.Arguments
	
	MyDocPath = WshShell.SpecialFolders("MyDocuments")	
	
	'Check Command Line Parameters if Needed
	If objArgs.Count = 1 Then
		CADDscr = objArgs(0)		
	Else
		CADDscr = "DrawBoxLC.scr"
		'Wscript.Echo "Not Enough Parameters"
	End If

	startPath= MyDocPath & "\TestCAD\" & CADDscr
	Wscript.Echo startPath	
    'Create Text File
	Set scrFile = fso.CreateTextFile(startPath, True)

	'Call Main Application
	Call ScriptWriter2(scrFile)

	scrFile.Close
	Wscript.Echo "All Done!"
	
End Sub