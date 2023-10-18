Option Explicit

Const MyDocsSubFolder="\CAD\default"

Const cadAppPath="C:\Program Files (x86)\ProgeCAD\progeCAD 2016 Professional ENG"
Const cadApp ="icad.exe"
Const CADDmode=2
Const delimiter="\"

Dim CADDApplication

Dim CADDscr
Dim CADDdoc
Dim CADDtemplate

Dim isScriptExist
Dim isCADDRunning
Dim nameScript
Dim fname

Dim scrFile


'Paths and file names with spaces needed to be contained between quotes (").
Function formatPath(fullpath ) 
  formatPath = """" & fullpath & """"
End Function

Function FmtAcadPath(fPath)
  Dim i , n 
  
  n = Len(fPath)
  For i = 1 To n
    If Mid(fPath, i, 1) = "\" Then
       Mid(fPath, i, 1) = "/"
    End If
  Next
  
  FmtAcadPath = fPath
End Function

Sub setConfigData()
	Dim fso
	Dim WshShell
	Dim MyDocPath


	Set WshShell = CreateObject("WScript.Shell")
	MyDocPath = WshShell.SpecialFolders("MyDocuments")
	
	nameScript = "default"
	fname = MyDocPath & MyDocsSubFolder & delimiter & nameScript
	If CADDScr = "" Then
		CADDscr = fname & ".scr"
	End If
	CADDtemplate = formatPath(fname & ".dwt")
	CADDdoc = formatPath(fname & ".dwg")
	
	CADDApplication = formatPath(cadAppPath & delimiter & cadApp)

End Sub


Sub cprintConfigData()

    Wscript.Echo "MyDocsSubFolder:: " & MyDocsSubFolder
    Wscript.Echo "nameScript:: " &  nameScript
    Wscript.Echo "fname:: " & fname
    Wscript.Echo "cadApp:: " & cadApp
    Wscript.Echo "cadAppPath:: " & cadAppPath
    Wscript.Echo "CADDmode:: " & CADDmode
    Wscript.Echo "CADDscr:: " & CADDscr
    Wscript.Echo "CADDtemplate:: " & CADDtemplate
    Wscript.Echo "CADDdoc:: ", CADDdoc
    Wscript.Echo "CADDApplication:: " & CADDApplication
    Wscript.Echo "isScriptExist:: " & isScriptExist
    
End Sub


Function CADopen()
	Dim WshShell
    Dim RetVal
    Dim shellCmdStr
    
    On Error Resume Next
	
    Set WshShell = CreateObject("WScript.Shell")
	
    isCADDRunning = False

    Select Case CADDmode
        Case 0
            shellCmdStr = CADDApplication
        Case 1
            shellCmdStr = CADDApplication & " " & CADDdoc
        Case 2
            shellCmdStr = CADDApplication & " /b " & CADDscr
        Case 3
            shellCmdStr = CADDApplication & " /t " & CADDtemplate & " /b " & CADDscr
    End Select

  
    Wscript.Echo shellCmdStr
    
    RetVal = WshShell.Run(shellCmdStr)
    
    If RetVal = 0 Then
        Wscript.Echo "Shell Successful ... "
        CADopen = True
    Else
        If Err.Number <> 0 Then
            CADopen = False
            Wscript.Echo "Shell Failed ... ", Err.Number, Err.Description
            Err.Clear    ' Clear Err object in case error occurred.
        End If
    End If

End Function


Sub CmdMain
	Dim fso
	Dim objArgs

	Set fso = CreateObject("Scripting.FileSystemObject")
	Set objArgs = WScript.Arguments
		
	'Check Command Line Parameters if Needed
	If objArgs.Count = 1 Then
		CADDscr = objArgs(0)		
		Wscript.Echo objArgs(0)
	Else
		Wscript.Echo "No Script File Name Provided Using Default Value"
		CADDscr = ""
		'CADDscr = "default.scr"
		'Wscript.Echo "Not Enough Parameters"
	End If

	setConfigData
	cprintConfigData
		
    'Create Text File
	Set scrFile = fso.CreateTextFile(CADDscr, True)
	Call TextScriptWriter5(scrFile)
	scrFile.Close
	
	If fso.FileExists(CADDscr) then
		Wscript.Echo CADDscr
		If CADopen() Then
			Wscript.Echo "CAD launched ..."
		Else
			Wscript.Echo "Failed to Launch CAD Application!"
		End If
	Else
		Wscript.Echo "No Script File Available!"
	End If
	
	Wscript.Echo "All Done!"
	
End Sub