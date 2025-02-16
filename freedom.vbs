Option Explicit 
Dim a, b, c, d, Ws, strExePath, strURL
a =MsgBox("Ouvrir le Site Freedom.fr ?",vbYesNo+vbQuestion)
If a=vbYes then
Set Ws = CreateObject("WScript.Shell") 
strExePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
strURL = "https://freedom.fr"
Ws.Run Chr(34) & strExePath & Chr(34) & " " & strURL
End if 
b =MsgBox("Ecouter La Radio Freedom 1 ?",vbYesNo+vbQuestion)
If b=vbYes then
Set Ws = CreateObject("WScript.Shell") 
strExePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
strURL = "https://freedom.fr/free-dom-1/"
Ws.Run Chr(34) & strExePath & Chr(34) & " " & strURL
End if 
c =MsgBox("Ecouter La Radio Freedom 2 ?",vbYesNo+vbQuestion)
If c=vbYes then 
Set Ws = CreateObject("WScript.Shell") 
strExePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
strURL = "https://freedom.fr/free-dom-2/"
Ws.Run Chr(34) & strExePath & Chr(34) & " " & strURL
End if
d =MsgBox("Consulter Freedom Club ?",vbYesNo+vbQuestion)
If d=vbYes then
Set Ws = CreateObject("WScript.Shell") 
strExePath = "C:\Program Files\Google\Chrome\Application\chrome.exe"
strURL = "https://freedom.fr/category/freedom-club/"
Ws.Run Chr(34) & strExePath & Chr(34) & " " & strURL
End if
