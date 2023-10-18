Option Explicit
'ProgeCAD Scripting Routines

Public R14Acad 'As Boolean
Public AcadR2000 'As Boolean
Public Const dfLine = "CONTINUOUS"
Public Const CADDFmtStr = "####0.000000"

Function R14Prefix(R14Acad) 
  If R14Acad Then
    R14Prefix = "-"
  Else
    R14Prefix = ""
  End If
End Function

Function R2000Prefix(R2000Acad) 
  If AcadR2000 Then
    R2000Prefix = "-"
  Else
    R2000Prefix = ""
  End If
End Function

Sub drawPoint(fp,pt1)
	fp.WriteLine("Point")
	fp.WriteLine(pt1.ptStr2)
End Sub


Sub startLine(fp)
	fp.WriteLine("line")
End Sub

Sub StartPline(fp)
  fp.WriteLine("polyline")
End Sub

Sub closeLines(fp)
	'fp.WriteLine("close")
	fp.WriteLine("c")
End Sub

Sub drawLine(fp,pt1,pt2)
	fp.WriteLine("line")
	fp.WriteLine(pt1.ptStr)
	fp.WriteLine(pt2.ptStr)
	fp.WriteLine("")
End Sub

Sub drawCircle(fp, pt1, diam)
  fp.WriteLine("CIRCLE")
  fp.WriteLine(ptStr2D(pt1))
  fp.WriteLine("D")
  fp.WriteLine(FormatNumber(diam, 6))
End Sub

Sub zoomExtents(fp)
	fp.WriteLine("ZOOM E")
End Sub

'Vertical assumed allowed
Sub DefTextStyle(fp, sname, fontName)
  fp.WriteLine(R14Prefix(R14Acad) & "STYLE")
  fp.WriteLine(sname)
  fp.WriteLine(fontName)
  fp.WriteLine(" 0 1 0 n n n")
End Sub

Sub write_Text(fp, sname, pt1, Hght, rot, txt)
  fp.WriteLine("TEXT")
  fp.WriteLine("STYLE")
  fp.WriteLine(sname)
  fp.WriteLine( ptStr2D(pt1))
  fp.WriteLine(FormatNumber(Hght, 1))
  fp.WriteLine(FormatNumber(rot, 1))
  fp.WriteLine(txt)
End Sub

Sub write_JText(fp, sname, TxtAlgn, pt1, Hght, rot, txt)
  fp.WriteLine("TEXT")
  fp.WriteLine("STYLE")
  fp.WriteLine(sname)
  fp.WriteLine("J")
  fp.WriteLine(TxtAlgn)
  fp.WriteLine( ptStr2D(pt1))
  fp.WriteLine(FormatNumber(Hght, 1))
  fp.WriteLine(FormatNumber(rot, 1))
  fp.WriteLine(txt) 
End Sub