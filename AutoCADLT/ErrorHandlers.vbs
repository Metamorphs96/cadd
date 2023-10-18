Option Explicit

'Set On Error Resume Next at start of each subroutine
'Then call this error handler where ever suspect a runtime error is occuring
'VBScript does not support on error goto label.
Sub ErrorHandler(errorNote)
	If Err.Number <> 0 Then
		Wscript.Echo "---------------------------------------------------------"
		Wscript.Echo "ERROR: " & errorNote
		Wscript.Echo "Error Number      : " & CStr(Err.Number)
		Wscript.Echo "Error Description : " & Err.Description
		Wscript.Echo "Error Source      : " & Err.Source
		Wscript.Echo "---------------------------------------------------------"
		Err.Clear
		Wscript.Quit
	End If
End Sub


'Testing array initialised based on:
'1) http://developer.rhino3d.com/guides/rhinoscript/testing-for-empty-arrays/
'   Author: Dale Fugier
'2) http://www.cpearson.com/excel/isarrayallocated.aspx
'   Author:  Charles H. Pearson
'If declare using Dim arr : arr = Array() then UBound(arr)=-1
'If declare using Dim arr() Then LBound(arr) and UBound(arr) both generate errors
'From testing: isEmpty(arr) and isNull(arr) and (arr is Nothing) do not assist
Function IsArrayAllocated(arr)
	Err.Clear
	If IsArray(arr) Then
		On Error Resume Next
		Dim ub : ub = UBound(arr) 'Generates an Error if array not allocated memory
		If (Err.Number = 0) And (ub >= 0) Then 
			IsArrayAllocated = True
		Else
			Err.Clear
			IsArrayAllocated = False
		End If
	Else
		IsArrayAllocated = False
	End If
End Function