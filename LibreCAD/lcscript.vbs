Option Explicit
'LibreCAD Scripting Routines
'Note LibreCAD command line only appears to support drawing tools and point input

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

Sub zoomExtents(fp)
	Wscript.Echo "zoomExtents: Not supported by LibreCAD command line"
	'fp.WriteLine("ZOOM E")
End Sub
