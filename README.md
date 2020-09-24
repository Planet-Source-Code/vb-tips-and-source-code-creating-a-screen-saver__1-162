<div align="center">

## Creating a Screen Saver


</div>

### Description

Create a screen saver in VB!
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[VB Tips and Source Code](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vb-tips-and-source-code.md)
**Level**          |Unknown
**User Rating**    |4.0 (12 globes from 3 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows System Services](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-system-services__1-35.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/vb-tips-and-source-code-creating-a-screen-saver__1-162/archive/master.zip)





### Source Code

```
In order to accomplish this task, start a new Visual Basic project. This example only requires a form - no VBXs or additional modules necessary. On the form, set the following properties:
•Caption = "" •ControlBox = False •MinButton = False •MaxButton = False •BorderStyle = 0 ' None •WindowState = 2 ' Maximized •BackColor = Black
The next order of business is to place a line (shape control) on the form. Draw it to any orientation and color you wish. Set the color by using the BorderColor property.
The last control that you will need to place on the form is a timer control. Set the timer's interval property anywhere from 100 to 500 (1/10 to 1/2 of a second).
In the general declarations section of the form you will need to declare two API functions. The first of these (SetWindowPos) is used to enable the form to stay on top of all other windows. The second (ShowCursor) is used to hide the mouse pointer while the screen saver runs and to restore it when the screen saver ends. The declares look like the following:
For VB3:
   Declare Function SetWindowPos Lib "user" (ByVal h%, ByVal hb%, ByVal x%, ByVal Y%, ByVal cx%, ByVal cy%, ByVal f%) As Integer
   Declare Function ShowCursor Lib "User" (ByVal bShow As Integer) As Integer
For VB4:
   Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
   Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
The first SUB we will write will be the routine that we will call to keep the form always on top. Place this SUB into the general declarations section of the form.
Sub AlwaysOnTop (FrmID As Form, OnTop As Integer)
  ' This function uses an argument to determine whether
  ' to make the specified form always on top or not
  Const SWP_NOMOVE = 2
  Const SWP_NOSIZE = 1
  Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
  Const HWND_TOPMOST = -1
  Const HWND_NOTOPMOST = -2
  If OnTop Then
    OnTop = SetWindowPos(FrmID.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
  Else
    OnTop = SetWindowPos(FrmID.hWnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
  End If
End Sub
The next issue we will take up will be the issue of getting the program started. This is of course the Form_Load event procedure. The actions we will take in this procedure is to randomize the number generator (so that the line moves around differently each time the screen saver is activated). We will also call the AlwaysOnTop SUB so that it will appear over everything else on the screen.
Sub Form_Load ()
  Dim x As Integer   ' Declare variable
  Randomize Timer    ' Variety is the spice of life
  AlwaysOnTop Me, True ' Cover everything else on screen
  x = ShowCursor(False) ' Hide MousePointer while running
End Sub
Now, before we handle the logic of making the line bounce around the screen, let's go ahead and handle shutting the program down. Most screen savers terminate when one of two things happen. Our's will end when the mouse is moved or when a key is pressed on the keyboard. Therefore we will need to trap two event procedures. Since there are no controls on the screen that can generate event procedures, we need to trap them at the form level. We will use the Form_KeyPress and Form_MouseMove event procedures to handle this. They appear as follows:
Sub Form_KeyPress (KeyAscii As Integer)
  Dim x As Integer
  x = ShowCursor(True) ' Restore Mousepointer
  Unload Me
  End
End Sub
Sub Form_MouseMove (Button As Integer, Shift As Integer, x As Single, Y As Single)
  Static Count As Integer
  Count = Count + 1 ' Give enough time for program to run
  If Count > 5 Then
    x = ShowCursor(True) ' Restore Mousepointer
    Unload Me
    End
  End If
End Sub
Finally, we need to handle the logic necessary to cause motion on the screen. I have created two sets of variables. One set DirXX handles the direction (1=Right or Down and 2=Left or Up) of the motion for each of the line control's four coordinates. The other set SpeedXX handles the speed factor for each of the line's four coordinates. These will be generated randomly (hence the Randomize Timer statement in Form_Load). These variables are Static, which of course means that each time the event procedure is called, they will retain their values from the preceeding time. The first time through the procedure they will also be set to zero. Therefore the program will assign these random values the first time through. From that point on, the program checks the direction of movement of each of the four coordinates and relocates them to a new position (the distance governed by the SpeedXX variable). The last section of code simply checks these coordinates to see if they left the visible area of the form and if they did their direction is reversed. This of course goes in the Timer's event procedure.
Sub Timer1_Timer ()
  Static DirX1 As Integer, Speedx1 As Integer
  Static DirX2 As Integer, Speedx2 As Integer
  Static DirY1 As Integer, Speedy1 As Integer
  Static DirY2 As Integer, Speedy2 As Integer
  ' Set initial Direction
  If DirX1 = 0 Then DirX1 = Rnd * 3
  If DirX2 = 0 Then DirX2 = Rnd * 3
  If DirY1 = 0 Then DirY1 = Rnd * 3
  If DirY2 = 0 Then DirY2 = Rnd * 3
  ' Set Speed
  If Speedx1 = 0 Then Speedx1 = 60 * Int(Rnd * 5)
  If Speedx2 = 0 Then Speedx2 = 60 * Int((Rnd * 5))
  If Speedy1 = 0 Then Speedy1 = 60 * Int((Rnd * 5))
  If Speedy2 = 0 Then Speedy2 = 60 * Int((Rnd * 5))
  ' Handle Movement
  ' If X1=1 then moving right else moving left
  ' If X2=1 then moving right else moving left
  ' If Y1=1 then moving down else moving up
  ' If Y2=1 then moving down else moving up
  If DirX1 = 1 Then
    Line1.X1 = Line1.X1 + Speedx1
  Else
    Line1.X1 = Line1.X1 - Speedx1
  End If
  If DirX2 = 1 Then
    Line1.X2 = Line1.X2 + Speedx2
  Else
    Line1.X2 = Line1.X2 - Speedx1
  End If
  If DirY1 = 1 Then
    Line1.Y1 = Line1.Y1 + Speedy1
  Else
    Line1.Y1 = Line1.Y1 - Speedy1
  End If
  If DirY2 = 1 Then
    Line1.Y2 = Line1.Y2 + Speedy2
  Else
    Line1.Y2 = Line1.Y2 - Speedy2
  End If
  ' Handle bouncing (change directions if off screen)
  If Line1.X1 < 0 Then DirX1 = 1
  If Line1.X1 > Me.ScaleWidth Then DirX1 = 2
  If Line1.X2 < 0 Then DirX2 = 1
  If Line1.X2 > Me.ScaleWidth Then DirX2 = 2
  If Line1.Y1 < 0 Then DirY1 = 1
  If Line1.Y1 > Me.ScaleHeight Then DirY1 = 2
  If Line1.Y2 < 0 Then DirY2 = 1
  If Line1.Y2 > Me.ScaleHeight Then DirY2 = 2
End Sub
Once you have entered all the preceeding code you have a nice little program that looks like a screen saver. You can compile it into an EXE and run it anytime you like. However, to make it into a true Windows screen-saver you need to do the following steps:
1.Choose "Make EXE File" from the File menu. 2.In the "Application Title" text box, type in the following: SCRNSAVE:VB4UandME Example 3.Change the extension in the EXE filename to have an SCR extension instead of an EXE. 4.Change the destination directory to your Windows directory (where all screen savers need to reside) 5.Click OK and let the compilation proceed.
At this point, you should be able to bring up the Windows Control Panel and select VB4UandME Example as the new screen saver. For Windows 3.1 this is found in the Desktop icon within Control Panel. For Windows 95, it is found in the Display icon in Control Panel (second tab).
```

