' https://isvbscriptdead.com/vbs-obfuscator/
Option Explicit

Function Rot47(str)
	Dim i, j, k, r
	j = Len(str)
	r = ""
	For i = 1 to j
		k = Asc(Mid(str, i, 1))
		If k >= 33 And k <= 126 Then
			r = r & Chr(33 + ((k + 14) Mod 94))
		Else
			r = r & Chr(k)
		End If
	Next
	Rot47 = r
End Function

Function Obfuscator(vbs)
	Dim length, s, i, F
	F = "Function l(str):Dim i,j,k,r:j=Len(str):r=" & Chr(34) & Chr(34) & ":For i=1 to j:k=Asc(Mid(str,i,1)):If k>=33 And k<=126 Then:r=r&Chr(33+((k+14)Mod 94)):Else:r=r&Chr(k):End If:Next:l=r:End Function"
	length = Len(vbs)
	s = ""
	For i = 1 To length
		s = s & Rot47(Mid(vbs, i, 1))
	Next	
	Obfuscator = F & vbCrlf & "Execute l(" & Chr(34) & (s)& Chr(34) &")" & vbCrLf 
End Function

If WScript.Arguments.Count = 0 Then
	WScript.Echo "Missing parameter(s): VBScript source file(s)"
	WScript.Quit
End If 

Dim fso, i
Const ForReading = 1
Set fso = CreateObject("Scripting.FileSystemObject")
For i = 0 To WScript.Arguments.Count - 1
	Dim FileName
	FileName = WScript.Arguments(i)
	Dim MyFile
	Set MyFile = fso.OpenTextFile(FileName, ForReading)
	Dim vbs
	vbs = MyFile.ReadAll	
	WScript.Echo Obfuscator(vbs)
	MyFile.Close
Next

Set fso = Nothing
