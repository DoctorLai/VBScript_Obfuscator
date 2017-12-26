' https://isvbscriptdead.com/vbs-obfuscator/
Option Explicit

Function btoa(sourceStr)
    Dim i, j, n, carr, rarr(), a, b, c
    carr = Array("A", "B", "C", "D", "E", "F", "G", "H", _
            "I", "J", "K", "L", "M", "N", "O" ,"P", _
            "Q", "R", "S", "T", "U", "V", "W", "X", _
            "Y", "Z", "a", "b", "c", "d", "e", "f", _
            "g", "h", "i", "j", "k", "l", "m", "n", _
            "o", "p", "q", "r", "s", "t", "u", "v", _
            "w", "x", "y", "z", "0", "1", "2", "3", _
            "4", "5", "6", "7", "8", "9", "+", "/")
    n = Len(sourceStr)-1
    ReDim rarr(n\3)
    For i=0 To n Step 3
        a = AscW(Mid(sourceStr,i+1,1))
        If i < n Then
            b = AscW(Mid(sourceStr,i+2,1))
        Else
            b = 0
        End If
        If i < n-1 Then
            c = AscW(Mid(sourceStr,i+3,1))
        Else
            c = 0
        End If
        rarr(i\3) = carr(a\4) & carr((a And 3) * 16 + b\16) & carr((b And 15) * 4 + c\64) & carr(c And 63)
    Next
    i = UBound(rarr)
    If n Mod 3 = 0 Then
        rarr(i) = Left(rarr(i),2) & "=="
    ElseIf n Mod 3 = 1 Then
        rarr(i) = Left(rarr(i),3) & "="
    End If
    btoa = Join(rarr,"")
End Function

Function char_to_utf8(sChar)
    Dim c, b1, b2, b3
    c = AscW(sChar)
    If c < 0 Then
        c = c + &H10000
    End If
    If c < &H80 Then
        char_to_utf8 = sChar
    ElseIf c < &H800 Then
        b1 = c Mod 64
        b2 = (c - b1) / 64
        char_to_utf8 = ChrW(&HC0 + b2) & ChrW(&H80 + b1)
    ElseIf c < &H10000 Then
        b1 = c Mod 64
        b2 = ((c - b1) / 64) Mod 64
        b3 = (c - b1 - (64 * b2)) / 4096
        char_to_utf8 = ChrW(&HE0 + b3) & ChrW(&H80 + b2) & ChrW(&H80 + b1)
    Else
    End If
End Function

Function str_to_utf8(sSource)
    Dim i, n, rarr()
    n = Len(sSource)
    ReDim rarr(n - 1)
    For i=0 To n-1
        rarr(i) = char_to_utf8(Mid(sSource,i+1,1))
    Next
    str_to_utf8 = Join(rarr,"")
End Function

Function str_to_base64(sSource)
    str_to_base64 = btoa(str_to_utf8(sSource))
End Function

Function Obfuscator(vbs)
	Dim s, F1, F2
	s = str_to_base64(vbs)
	F1 = "Function l(a): With CreateObject(" & Chr(34) & "Msxml2.DOMDocument" & Chr(34) & ").CreateElement(" & Chr(34) & "aux" & Chr(34) & "): .DataType = "& Chr(34) & "bin.base64"& Chr(34) & ": .Text = a: l = r(.NodeTypedValue): End With: End Function"
	F2 = "Function r(b): With CreateObject("& Chr(34) & "ADODB.Stream"& Chr(34) & "): .Type = 1: .Open: .Write b: .Position = 0: .Type = 2: .CharSet = "& Chr(34) & "utf-8"& Chr(34) & ": r = .ReadText: .Close:  End With: End function"
	Obfuscator = F1 & vbCrLf & F2 & vbCrLf & "Execute l(" & Chr(34) & (s)& Chr(34) &")" & vbCrLf 
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
