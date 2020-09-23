<div align="center">

## Bezier splines


</div>

### Description

A simple Bezier spline implementation. Allows the user to select control 'points on a picture box and then draw a Bezier curve between them.

NEW!! - User can now move control points!!
 
### More Info
 
I just wrote this to help me with something else so it's not even slightly optimised - in fact its really badly done but it does the job. Implements the explicit x,y functions of the normal parametric equation.

You can move the points by selecting clicking

on them and dragging with the left button

If you want to put multiple points on the same

location add them with the right button

Place the following controls on a form

Command button NAME= cmdReset

Picture Box  NAME=picDisplay

Label     NAME = label1


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mark Roberts](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mark-roberts.md)
**Level**          |Intermediate
**User Rating**    |5.0 (30 globes from 6 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Custom Controls/ Forms/  Menus](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/custom-controls-forms-menus__1-4.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mark-roberts-bezier-splines__1-1871/archive/master.zip)





### Source Code

```
Dim nc As Integer
Dim Cont(100, 1) As Integer
Dim NewLocPoint As Integer
Const Smooth = 0.02
Dim Dragging As Boolean
Function B(k, n, u)
 'Bezier blending function
 B = C(n, k) * (u ^ k) * (1 - u) ^ (n - k)
End Function
Function C(n, r)
 ' Implements c!/r!*(n-r)!
 C = fact(n) / (fact(r) * fact(n - r))
End Function
Function fact(n)
 ' Recursive factorial fucntion
 If n = 1 Or n = 0 Then
 fact = 1
 Else
 fact = n * fact(n - 1)
 End If
End Function
Private Sub AddCont(X, Y)
 Cont(nc, 0) = X: Cont(nc, 1) = Y
 nc = nc + 1
End Sub
Private Sub cmdReset_Click()
 nc = 0
 picDisplay.Cls
End Sub
Private Sub Form_Load()
 Form1.ScaleMode = vbTwips
 Form1.Caption = "Bezier Curves by Mark Roberts"
 Form1.Move 900, 900, 5900, 5200
 picDisplay.Move 120, 120, 5535, 4250
 cmdReset.Move 4640, 4435, 1015, 255
 cmdReset.Caption = "&Reset"
 With Label1
 .BackColor = &HC0FFFF
 .BorderStyle = vbFixedSingle
 .Move 120, 4440, 4435, 255
 .Alignment = vbCenter
 .Caption = "Select new points or drag points to move"
 End With
 picDisplay.ScaleMode = vbPixels
 picDisplay.FontSize = 5
End Sub
Private Sub picDisplay_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 xv = Int(X): yv = Int(Y) 'In case not pixels
 cval = Clicked(xv, yv)
 If cval > -1 And Button = 1 Then ' In case you want multiple points
 Dragging = True
 NewLocPoint = cval
 Label1.Caption = "Dragging point " + Trim$(cval + 1)
 Else
 AddCont xv, yv  'Add the control points
 picDisplay.Circle (xv, yv), 2, 255
 picDisplay.Print nc
 If nc = 1 Then
 PSet (xv, yv)
 Else
 picDisplay.DrawStyle = vbDot
 picDisplay.Line (Cont(nc - 2, 0), Cont(nc - 2, 1))-(Cont(nc - 1, 0), Cont(nc - 1, 1)), 0
 picDisplay.DrawStyle = vbSolid
 End If
 If nc > 1 Then Redraw
 End If
End Sub
Private Sub picDisplay_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Clicked(X, Y) > -1 Then
 MousePointer = vbCrosshair
 Else
 MousePointer = vbDefault
 End If
 If Dragging = True Then
 xv = Int(X): yv = Int(Y)
 Cont(NewLocPoint, 0) = xv: Cont(NewLocPoint, 1) = yv
 Redraw
 End If
End Sub
Private Sub picDisplay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 ' End dragging operation
 If Dragging = True Then
 Dragging = False
 Redraw
 Label1.Caption = "Select new points or drag current ones"
 End If
End Sub
Private Function Clicked(X, Y)
 ' Has the user clicked within the circle
 ' of a current point
 For i = 0 To nc
 xp = Cont(i, 0): yp = Cont(i, 1)
 If Abs(xp - X) < 3 And Abs(yp - Y) < 3 Then
 Clicked = i
 Exit Function
 End If
 Next i
 Clicked = -1
End Function
Sub Redraw()
 'Redraws entire display
 picDisplay.Cls
 For i = 1 To nc
 xv = Cont(i - 1, 0): yv = Cont(i - 1, 1)
 picDisplay.Circle (xv, yv), 2, 255
 picDisplay.Print i
 Next i
 picDisplay.DrawStyle = vbDot
 For i = 0 To nc - 2
 picDisplay.Line (Cont(i, 0), Cont(i, 1))-(Cont(i + 1, 0), Cont(i + 1, 1)), 0
 Next i
 picDisplay.DrawStyle = vbSolid
 DrawBezier Smooth
End Sub
Sub DrawBezier(du)
 ' Draws a Bezier curve using the control points given in
 ' Cont(...). Uses delta u steps
 n = nc - 1 'N = number of control points -1
 If n < 1 Then
 MsgBox "Need more control points", vbInformation
 Exit Sub
 End If
 picDisplay.PSet (Cont(0, 0), Cont(0, 1)) 'Plot the first point
 For u = 0 To 1 Step du
 X = 0: Y = 0
 For k = 0 To n ' For each control point
 bv = B(k, n, u) ' Calculate blending function
 X = X + Cont(k, 0) * bv
 Y = Y + Cont(k, 1) * bv
 Next k
 picDisplay.Line -(X, Y), 65535 ' Draw to the point
 Next u
 picDisplay.Line -(Cont(n, 0), Cont(n, 1)), 65535
End Sub
```

