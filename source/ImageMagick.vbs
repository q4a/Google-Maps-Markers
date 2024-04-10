'============================================
'Settings
'
'convert base.png -font Roboto-Bold -pointsize 13 -fill "rgb(72,13,14)" -gravity center -draw "text 1,-8 'A'" 1.png
'============================================
strPath  = Replace(WScript.ScriptFullName,Wscript.ScriptName,"")
strFile  = ""
strSave  = "markers\"
strPos   = "text 1,-8"
strBase  = """C:\Program Files\ImageMagick-6.9.13-Q8\convert.exe"" {base} -font Roboto-Bold -pointsize 13 -fill ""{color}"" -gravity center -draw """ & strPos & " '{text}'"" """ & strPath & "{file}"""
intNumer = 99
strAlpha = "О, !"

'============================================
'List Base Files (_red.png, _green.png, etc)
'============================================
Set objFSO    = CreateObject("Scripting.FileSystemObject")
Set objFolder = objFSO.GetFolder(strPath)
Set objFiles  = objFolder.Files

For Each objFile In objFiles
	strFileFull  = objFile.Name
	strFileExt   = LCase(Mid(strFileFull,InStrRev(strFileFull,".")+1))
	strFileColor = Mid(strFileFull,2,InStrRev(strFileFull,".")-2)
	
	If Left(strFileFull,1) = "_" And strFileExt = "png" Then

		' Numeric Markers
		For i = 1 To intNumer
			Select Case i
				Case 1, 2, 3, 5, 6, 7, 9
					strExe = Replace(strBase, strPos, "text 0,-8")
				Case Else
					strExe = strBase
			End Select

			Run InsertParams(strExe, strFileFull, TextColor(strFileColor), i, strSave & strFile & i & "_" & strFileColor  & ".png")
		Next


		' Alphabetic Markers
		arrAlpha = Split(strAlpha, ", ")

		For i = 0 To UBound(arrAlpha)
			strChar = arrAlpha(i)
			strExe = strBase
			Select Case strChar
				Case "@"
					strExe = Replace(strBase, strPos, "text 2,-9")
				Case Else
					strExe = strBase
			End Select

			Run InsertParams(strExe, strFileFull, TextColor(strFileColor), i, strSave & strFile & i & "_" & strFileColor  & ".png")
		Next

	End If
Next

Set objFiles  = Nothing
Set objFolder = Nothing
Set objFSO    = Nothing


'============================================
'Functions
'============================================

' Insert Params
Function InsertParams(strExe, strBase, strColor, strText, strFile)
	strExe = Replace(strExe, "{base}",  strBase)
	strExe = Replace(strExe, "{color}", strColor)
	strExe = Replace(strExe, "{text}",  strText)
	strExe = Replace(strExe, "{file}",  strFile)
	InsertParams = strExe
End Function


' Marker Text Color
Function TextColor(strColor)
	Select Case LCase(strColor)
		Case LCase("000080")
			TextColor = "#FFFFFF"
		Case LCase("0000CD")
			TextColor = "#FFFFFF"
		Case LCase("0000FF")
			TextColor = "#FFFFFF"
		Case LCase("008000")
			TextColor = "#FFFFFF"
		Case LCase("008080")
			TextColor = "#FFFFFF"
		Case LCase("191970")
			TextColor = "#FFFFFF"
		Case LCase("2F4F4F")
			TextColor = "#FFFFFF"
		Case LCase("4169E1")
			TextColor = "#FFFFFF"
		Case LCase("556B2F")
			TextColor = "#FFFFFF"
		Case LCase("800000")
			TextColor = "#FFFFFF"
		Case LCase("800080")
			TextColor = "#FFFFFF"
		Case Else
			TextColor = "#000000"
	End Select
End Function


' Run File
Sub Run(ByVal strFile)
	Dim objShell
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run strFile, 0, True
    Set objShell = Nothing
End Sub
