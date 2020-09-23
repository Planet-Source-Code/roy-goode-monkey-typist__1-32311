<div align="center">

## Monkey Typist


</div>

### Description

This code is basically a random text generator that taps out random strings of characters. It is a Monkey Typist because sometimes it might just make sense! The idea is that you watch the text appear and then see if the 'Monkey' enters real words or even sentences!
 
### More Info
 
The user is not required to input anything, just run the program and watch that monkey type!

1. Start a new Standard EXE project

' 2. Rename Form1 to frmMonkey

' 3. Put a text box onto the form and name it txtMonkey

' 4. Set the text box's 'MultiLine' property to True

' 5. Set the text box's 'ScrollBars' property to 2 - Vertical

' 6. Put a timer onto the form and name it tmrMonkey

' 7. Set the timer's 'Enabled' property to False

' 8. Copy this code into the form's class module


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Roy Goode](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/roy-goode.md)
**Level**          |Advanced
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Jokes/ Humor](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/jokes-humor__1-40.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/roy-goode-monkey-typist__1-32311/archive/master.zip)





### Source Code

```
Option Explicit
Dim intOriginalRnd As Integer
Private Sub Form_Load()
Dim fntA As New StdFont
fntA.Name = "Tahoma"
fntA.Size = 8
Me.Width = 6225
txtMonkey.Font = fntA
txtMonkey.Text = ""
txtMonkey.Locked = True
tmrMonkey.Interval = 100
tmrMonkey.Enabled = True
End Sub
Private Sub Form_Resize()
tmrMonkey.Enabled = False
With frmMonkey
txtMonkey.Height = .ScaleHeight
txtMonkey.Width = .ScaleWidth
txtMonkey.Left = .ScaleLeft
txtMonkey.Top = .ScaleTop
If .WindowState = 0 Then
 .Left = (Screen.Width - .Width) / 2
 .Top = (Screen.Height - .Height) / 2
End If
End With
tmrMonkey.Enabled = True
End Sub
Private Sub tmrMonkey_Timer()
intOriginalRnd = Int(Rnd * 10)
Dim intRnd As Integer
Randomize
frmMonkey.Caption = "Monkey Typist - Click text to stop"
If intOriginalRnd < 1 Then
 intRnd = 32
ElseIf intOriginalRnd < 2 Then
 intRnd = Int(3 * Rnd + 44)
ElseIf intOriginalRnd < 6 Then
 intRnd = Int(26 * Rnd + 97)
ElseIf intOriginalRnd < 10 Then
 intRnd = Int(26 * Rnd + 65)
End If
frmMonkey.Caption = "Monkey Typist - Click text to stop - " & Chr(intRnd)
txtMonkey.Text = txtMonkey.Text & Chr(intRnd)
End Sub
Private Sub txtMonkey_Click()
If tmrMonkey.Enabled = True Then
 tmrMonkey.Enabled = False
 frmMonkey.Caption = "Monkey Typist - Click text to start"
Else
 tmrMonkey.Enabled = True
End If
End Sub
```

