Set obiektWMI = GetObject("winmgmts:\\.\root\CIMV2") 
Set ustawEQ = obiektWMI.ExecQuery("SELECT * FROM Win32_Process",,48)
sprawdzaPR = "linkbar64.exe"
For Each licznikPR in ustawEQ 
  If LCase(licznikPR.Name) = LCase(sprawdzaPR) Then istniejePR = True
Next
Set obiektWS = WScript.CreateObject ("WScript.Shell")
If Not(istniejePR = True) Then
  obiektWS.Run sprawdzaPR
Else
  obiektWS.Run "rundll32 user32.dll, MessageBeep"
End If
Set obiektWMI = Nothing: Set obiektWS = Nothing

