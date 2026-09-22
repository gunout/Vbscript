Option Explicit 
Dim a, b, c, d, Ws, strURL

Set Ws = CreateObject("WScript.Shell")

a = MsgBox("Ouvrir le Site Freedom.fr ?", vbYesNo + vbQuestion)
If a = vbYes Then
    strURL = "https://freedom.fr"
    Ws.Run strURL
End If

b = MsgBox("Ecouter La Radio Freedom 1 ?", vbYesNo + vbQuestion)
If b = vbYes Then
    strURL = "https://freedom.fr/free-dom-1/"
    Ws.Run strURL
End If

c = MsgBox("Ecouter La Radio Freedom 2 ?", vbYesNo + vbQuestion)
If c = vbYes Then
    strURL = "https://freedom.fr/free-dom-2/"
    Ws.Run strURL
End If

d = MsgBox("Consulter Freedom Club ?", vbYesNo + vbQuestion)
If d = vbYes Then
    strURL = "https://freedom.fr/category/freedom-club/"
    Ws.Run strURL
End If
