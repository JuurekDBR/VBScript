Set obiektWMI = GetObject("winmgmts:\\.\root\CIMV2") 
Set ustawEQ = obiektWMI.ExecQuery("SELECT * FROM Win32_Process",,48)
nazwaPR = "rklauncher.exe"
For Each licznikPR in ustawEQ 
  If LCase(licznikPR.Name) = LCase(nazwaPR) Then testPR = True
Next
Set obiektWS = WScript.CreateObject ("WScript.Shell")
If Not(testPR = True) Then
  obiektWS.Run nazwaPR
Else
  obiektWS.Run "rundll32 user32.dll, MessageBeep"
End If
Set obiektWMI = Nothing: Set obiektWS = Nothing
  
