VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Real Time Alpha Blending"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   3435
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      LargeChange     =   10
      Left            =   120
      Max             =   255
      SmallChange     =   10
      TabIndex        =   6
      Top             =   2760
      Width           =   3255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   3360
      Width           =   2175
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   10
      Left            =   120
      Max             =   255
      SmallChange     =   10
      TabIndex        =   3
      Top             =   2400
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Use Look Up Tables"
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   3120
      Width           =   2175
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   2280
      Left            =   1800
      Picture         =   "Real Time Alpha Blend.frx":0000
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   1
      Top             =   120
      Width           =   1530
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2280
      Left            =   120
      Picture         =   "Real Time Alpha Blend.frx":1E8C
      ScaleHeight     =   150
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   100
      TabIndex        =   0
      Top             =   120
      Width           =   1530
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alpha B -"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   75
      TabIndex        =   9
      Top             =   3360
      Width           =   675
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alpha A -"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   75
      TabIndex        =   8
      Top             =   3120
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   840
      TabIndex        =   7
      Top             =   3360
      Width           =   105
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   840
      TabIndex        =   5
      Top             =   3120
      Width           =   105
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Declare Function GetPixel Lib "gdi32" (ByVal hDc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long



'I set it to double so the colors are more precise when the math is done to it.
'The speed wont change when set to Long if LUT's are done but don't set it on
'Byte or Integer cause I got errors when it exceeded the data type.

Private Type RGBA
Red As Double
Green As Double
Blue As Double
Alpha As Double
End Type

'Throw in any pictures you want, but just make sure you change these:

Const Picture_A_Width As Integer = 100
Const Picture_A_Height As Integer = 150
Const Picture_B_Width As Integer = 100
Const Picture_B_Height As Integer = 150

'You can create Look Up Tables for just about anything. Cause if you work with
'numbers that only go so far rather than having to work with -32768 to 32767 for integers
'(which is possible for addition and subtraction I believe) and -2147483648 to 2147483647 for
'long, etc, then LUT's can come in handy for speeding up your program tremendously. Whats good
'with what I'm working with is that my numbers for the math only go from 0 to 255 for one number
'and 0 to 255 for the other to work with RGB color values. Also the Widths (0 to 100) and Heights
'(0 to 150) to the pictures I'm working with. Little tip on LUT's though which I learned
'on my own. Never have 3 or more arrays for your variable. If for example you have
'Test(255,255,255) and you throw that into a triple For Loop to pre calculate every
'combination, you are looking at 255 * 255 * 255 = 16777215 Combinations!!! It wont crash
'with that often but if you use 4 (255 * 255 * 255 * 255 = Ahhhhh dont even bother trying) the
'computer will just give up on you and you'll get a Project1 (Not Responding) when you try to
'Ctrl + Alt + Delete it. Here is an example of one Look Up Table that you may have used back
'in the ol elementary days:

'For A = 0 to 12
'For B = 0 to 12
'Multiplication_Table(A, B) = A * B
'Next B
'Next A

'If you were to do the multiplication from the top of your head, it would take a moment
'to figure out the answer. But if we were staring at the multiplication table, it would
'be faster to find the answer than to think it. That's how things work here, only the
'computer is doing it like about 15000 times over.

Dim Percent_Look_Up_Table(255, 255) As Double
Dim Multiplication_Look_Up_Table(255, 255) As Double
Dim Addition_Look_Up_Table(255, 255) As Double
Dim Subtraction_Look_Up_Table(255, 255) As Double

Dim Red_A_Look_Up_Table(Picture_A_Width, Picture_A_Height) As Double
Dim Green_A_Look_Up_Table(Picture_A_Width, Picture_A_Height) As Double
Dim Blue_A_Look_Up_Table(Picture_A_Width, Picture_A_Height) As Double

Dim Red_B_Look_Up_Table(Picture_B_Width, Picture_B_Height) As Double
Dim Green_B_Look_Up_Table(Picture_B_Width, Picture_B_Height) As Double
Dim Blue_B_Look_Up_Table(Picture_B_Width, Picture_B_Height) As Double

Dim Color_A As RGBA, Color_B As RGBA, Color_C As RGBA

'This turns your high speed blends on and off, with the default as off.

Dim Look_Up_Table_Activated As Boolean

'Need this to set up the LUT's in the 2 dimensional arrays and to plot pixels in
'a 2D picture.

Dim A As Double, B As Double

Private Sub Pre_Calculate_Percent_Look_Up_Table()
For A = 0 To 255
For B = 0 To 255
Percent_Look_Up_Table(A, B) = ((A / 255) * (B / 255)) * 255
Next B
Next A
End Sub

Private Sub Pre_Calculate_Multiplication_Look_Up_Table()
For A = 0 To 255
For B = 0 To 255
Multiplication_Look_Up_Table(A, B) = (A * B) / 255
Next B
Next A
End Sub

Private Sub Pre_Calculate_Addition_Look_Up_Table()

For A = 0 To 255
For B = 0 To 255
Addition_Look_Up_Table(A, B) = A + B
Next B
Next A
End Sub

Private Sub Pre_Calculate_Subtraction_Look_Up_Table()
For A = 0 To 255
For B = 0 To 255
Subtraction_Look_Up_Table(A, B) = A - B
Next B
Next A
End Sub


Private Sub Pre_Calculate_Color_Look_Up_Table()

'These are the actual formulas for obtaining the RGB values from pixels:

' Blue = Int(Color / 65536)
' Green = Int((Color - (Blue * 65536)) / 256)
' Red = Int(Color - (Blue * 65536) - (Green * 256))

'Problem is that you are looking at 2 divisions, 3 multiplications, and 3 subtractions per pixel!!!
'And if you were to do that on a 100 x 100 picture for example, that's over 10000 pixels!!! That's
'not good if you are wanting to do it in real time to begin with. Probably why people stuck
'with API's, DirectX, and OpenGL. One solution around this without look up tables would be to use
'bitwise comparisons. That way all you have are 2 divisions and 4 bitwise comparisons per pixel,
'which isn't too bad at all. Bitwise operators are extremly fast on Visual Basic and also on C++ cause
'it operates on a binary level. For example, Make a variable = 12 And 10. 12 is 1100 in binary and 10
'is 1010. The And operator works like this with all numbers:

'Expression 1  And  Expression 2  =  Result
'-----------------------------------------
'   True               True          True
'   True               False         False
'   False              True          False
'   False              False         False

'Where True = 1 and False = 0. That means it will do this to 12 and 10 --> 1100 And 1010 = 1000.
'So 12 And 10 = 8. That's why I use these instead of all that slow math. So heres my new and improved
'formula for obtaining RGB values from pixels:

' Red = Int(Color And 255) And 255
' Green = Int (Color / 256) And 255
' Blue = Int(Color / 65536) And 255

For B = 0 To Picture1.ScaleHeight
For A = 0 To Picture1.ScaleWidth
Red_A_Look_Up_Table(A, B) = Int(Picture1.Point(A, B) And 255) And 255
Green_A_Look_Up_Table(A, B) = Int(Picture1.Point(A, B) / 256) And 255
Blue_A_Look_Up_Table(A, B) = Int(Picture1.Point(A, B) / 65536) And 255

Red_B_Look_Up_Table(A, B) = Int(Picture2.Point(A, B) And 255) And 255
Green_B_Look_Up_Table(A, B) = Int(Picture2.Point(A, B) / 256) And 255
Blue_B_Look_Up_Table(A, B) = Int(Picture2.Point(A, B) / 65536) And 255
Next A
Next B

End Sub

Private Sub Command1_Click()
'This will clear the picture, reset the alphas, pre calculate the RGB colors one
'pixel at a time from both pics (only need to do that once throughout the whole program
'really) and change what you are going to use to change the alphas, whether it's LUT's or pure
'math.

Picture2.Cls
HScroll1.Value = 0
HScroll2.Value = 0
Color_A.Alpha = 0
Color_B.Alpha = 0
Pre_Calculate_Color_Look_Up_Table
If Look_Up_Table_Activated = False Then
Command1.Caption = "Do All That Slow Math"
Look_Up_Table_Activated = True
Picture2.AutoRedraw = False
Else
Look_Up_Table_Activated = False
Command1.Caption = "Use Look Up Tables"
Picture2.AutoRedraw = True
End If
End Sub

Private Sub Command2_Click()
'This will just clear the picture and reset the alphas.

Picture2.Cls
HScroll1.Value = 0
HScroll2.Value = 0
Color_A.Alpha = 0
Color_B.Alpha = 0
End Sub




Private Sub Form_Load()

'You can fiddle with these to change the speed of the alpha blend in the scroll bars.

With HScroll1
.LargeChange = 10
.SmallChange = 10
End With

With HScroll2
.LargeChange = 10
.SmallChange = 10
End With

'I don't need to really set this when I have the LUT's on cause it plots pixels so fast, you can't
'see it blend. AutoRedraw slow depending on how you use it, so use the AutoRedraw wisely. I used it on
'my 3D Engine one time and was trying to produce polygon shading, but it was extremly fast without
'AutoRedraw yet I had to see it shade it when I wasn't supposed to cause of all the math done. I'll
'figure out how to Double Buffer in real time someday without using AutoRedraw. But since we need it
'for pure math, I'll enable it from the beginning and turn it off when I use LUT's

Picture2.AutoRedraw = True


'This will pre calculate all the math needed for alpha blending so we never need to do math at all.

Pre_Calculate_Percent_Look_Up_Table
Pre_Calculate_Multiplication_Look_Up_Table
Pre_Calculate_Addition_Look_Up_Table
Pre_Calculate_Subtraction_Look_Up_Table

'One problem with this next LUT is that theres a slight glitch in Visual Basic. You can't obtain
'points within the Form_Load, Form_Activate, Form_Paint, etc cause it will always read 16777215.
'It will be like picture1 is white. Same with using the GetPixel API I had up there to see if
'I could get around that. Turns out the form isn't always fully loaded in the begining. Go ahead
'and see for yourself by enabling the Color LUT and making the Look_Up_Table_Activate = True. If
'you can get around this from the beginning of the program, please let me know how.

Look_Up_Table_Activated = False
'Pre_Calculate_Color_Look_Up_Table

If Look_Up_Table_Activated = True Then
Command1.Caption = "Do All That Slow Math"
Else
Command1.Caption = "Use Look Up Tables"
End If

End Sub


Private Sub HScroll1_Change()

'Well we have to know what that scroll bar is on somehow. Duh!!!

Label1.Caption = HScroll1.Value

'I know there are faster methods for clearing without API's (I'll leave that challenge to you)
'But we don't need it for our LUT's since they already know what the pixels are. That'll speed
'it up more by not clearing the picture at all. But we most definitely need it for the slow way
'since it needs to clear the pixels. Otherwise you get terrible results.

If Look_Up_Table_Activated = False Then
Picture2.Cls
End If

'Did you know that the computer can do pixels faster horizontal than vertical? That's what
'I read in my 3D Game Programming book, so we will plot these pixels in a row by row fashion.

For B = 0 To Picture1.ScaleHeight
For A = 0 To Picture1.ScaleWidth
    
'Here is the slow method for obtaining RGB values from pixels out of both pics with 4
'divisions, 6 multiplications, and 6 subtractions per pixel.
    
If Look_Up_Table_Activated = False Then
Color_A.Blue = Int(Picture1.Point(A, B) / 65536)
Color_A.Green = Int((Picture1.Point(A, B) - (Color_A.Blue * 65536)) / 256)
Color_A.Red = Int(Picture1.Point(A, B) - (Color_A.Blue * 65536) - (Color_A.Green * 256))

Color_B.Blue = Int(Picture2.Point(A, B) / 65536)
Color_B.Green = Int((Picture2.Point(A, B) - (Color_B.Blue * 65536)) / 256)
Color_B.Red = Int(Picture2.Point(A, B) - (Color_B.Blue * 65536) - (Color_B.Green * 256))
End If
    
'But here, the math has already been done for you on these Look Up Tables. So no math is
'performed at all. Alright!!!

If Look_Up_Table_Activated = True Then
Color_A.Red = Red_A_Look_Up_Table(A, B)
Color_A.Green = Green_A_Look_Up_Table(A, B)
Color_A.Blue = Blue_A_Look_Up_Table(A, B)

Color_B.Red = Red_B_Look_Up_Table(A, B)
Color_B.Green = Green_B_Look_Up_Table(A, B)
Color_B.Blue = Blue_B_Look_Up_Table(A, B)
End If

'Here's the alpha that's similar to OpenGL's alpha only mine is within 0 to 255
'to go with the Red, Green, and Blue's 0 to 255's. It won't have effect if any one
'of the alphas are 0 (just like OpenGL) so fiddle with both scroll bars and see
'what happens.

Color_A.Alpha = HScroll1.Value
Color_B.Alpha = HScroll2.Value

'Next we keep the values in range so we can prevent an error.

If Color_A.Red > 255 Then Color_A.Red = 255
If Color_A.Green > 255 Then Color_A.Green = 255
If Color_A.Blue > 255 Then Color_A.Blue = 255
If Color_B.Red > 255 Then Color_B.Red = 255
If Color_B.Green > 255 Then Color_B.Green = 255
If Color_B.Blue > 255 Then Color_B.Blue = 255
If Color_A.Red < 0 Then Color_A.Red = 0
If Color_A.Green < 0 Then Color_A.Green = 0
If Color_A.Blue < 0 Then Color_A.Blue = 0
If Color_B.Red < 0 Then Color_B.Red = 0
If Color_B.Green < 0 Then Color_B.Green = 0
If Color_B.Blue < 0 Then Color_B.Blue = 0

'Oh boy, doesn't this suck for speed. It got the pixel from both pics with 4 divisions,
'6 multiplications, and 6 subtractions. Then it needs to do 12 divisions, 12 multiplications,
'6 subtractions, and 6 additions to blend the pics together. All together it's 16 divisions,
'18 multiplications, 12 subtractions, and 6 additions per pixel out of two (100 x 150) pictures
'(30000 pixels)...Yikes!!! But anyways, here's the formula to alpha blend pics together:

'Red = (Red_B - (Red_B * ((Alpha_A / 255) * (Alpha_B / 255))) + (Red_A * ((Alpha_A / 255) * (Alpha_B / 255))
'Green = (Green_B - (Green_B * ((Alpha_A / 255) * (Alpha_B / 255))) + (Green_A * ((Alpha_A / 255) * (Alpha_B / 255))
'Blue = (Blue_B - (Blue_B * ((Alpha_A / 255) * (Alpha_B / 255))) + (Blue_A * ((Alpha_A / 255) * (Alpha_B / 255))

If Look_Up_Table_Activated = False Then
Color_C.Red = (Color_B.Red - (Color_B.Red * ((Color_A.Alpha / 255) * (Color_B.Alpha / 255)))) + (Color_A.Red * ((Color_A.Alpha / 255) * (Color_B.Alpha / 255)))
Color_C.Green = (Color_B.Green - (Color_B.Green * ((Color_A.Alpha / 255) * (Color_B.Alpha / 255)))) + (Color_A.Green * ((Color_A.Alpha / 255) * (Color_B.Alpha / 255)))
Color_C.Blue = (Color_B.Blue - (Color_B.Blue * ((Color_A.Alpha / 255) * (Color_B.Alpha / 255)))) + (Color_A.Blue * ((Color_A.Alpha / 255) * (Color_B.Alpha / 255)))
End If

'But who needs math right. Math sucks nads. Might as well use precalculated Look Up Tables and
'cheat a bit. Damn I'm good!!! Now it does no math at all. Sweeeeeeeeeeeeeet!

If Look_Up_Table_Activated = True Then
Color_C.Red = Addition_Look_Up_Table(Subtraction_Look_Up_Table(Color_B.Red, Multiplication_Look_Up_Table(Color_B.Red, Percent_Look_Up_Table(Color_A.Alpha, Color_B.Alpha))), Multiplication_Look_Up_Table(Color_A.Red, Percent_Look_Up_Table(Color_A.Alpha, Color_B.Alpha)))
Color_C.Green = Addition_Look_Up_Table(Subtraction_Look_Up_Table(Color_B.Green, Multiplication_Look_Up_Table(Color_B.Green, Percent_Look_Up_Table(Color_A.Alpha, Color_B.Alpha))), Multiplication_Look_Up_Table(Color_A.Green, Percent_Look_Up_Table(Color_A.Alpha, Color_B.Alpha)))
Color_C.Blue = Addition_Look_Up_Table(Subtraction_Look_Up_Table(Color_B.Blue, Multiplication_Look_Up_Table(Color_B.Blue, Percent_Look_Up_Table(Color_A.Alpha, Color_B.Alpha))), Multiplication_Look_Up_Table(Color_A.Blue, Percent_Look_Up_Table(Color_A.Alpha, Color_B.Alpha)))
End If

'This seems like the only way to plot pixels 1 at a time in pure VB but there are faster methods
'out there without API calls and stuff. I'll stick with the PSet for now. It's still a tad slow cause
'of the fact we are plotting one pixel at a time but still it works within real time when using Look
'Up Tables, so it's all good. It will seem slow yet fast as a mofo on my LUT's cause of the limited
'speed the scroll bars can go when you hold the button down on it. So it should be faster if you use
'some better method to change the alpha values, such as Do loops. Do loops work better than timers
'cause of it's amazingly high speed but it's speed isn't as constant as timers are, yet timers are so
'slow. So if you want, you can use a Do loop to add up the alpha values Just dont forget to add DoEvents
'after the Do. I stuck with scroll bars to give you total control of the alpha changes in this program
'just to let you know. If you want, you can enable the SetPixel API one just one of the scroll bars and
'keep PSet on another and see a difference in speed. There really wouldn't be too much difference when
'my LUT's are on. That's pretty damn fast if I had just made my pure VB code almost as fast as this API
'call. And this API call is 5% faster than the SetPixelV API. Wow!!!

Picture2.PSet (A, B), RGB(Color_C.Red, Color_C.Green, Color_C.Blue)
'SetPixel Picture2.hdc, A, B, RGB(Color_C.Red, Color_C.Green, Color_C.Blue)

Next A
Next B
End Sub


Private Sub HScroll2_Change()

'Well we have to know what that scroll bar is on somehow. Duh!!!

Label2 = HScroll2.Value

'I know there are faster methods for clearing without API's (I'll leave that challenge to you)
'But we don't need it for our LUT's since they already know what the pixels are. That'll speed
'it up more by not clearing the picture at all. But we most definitely need it for the slow way
'since it needs to clear the pixels. Otherwise you get terrible results.

If Look_Up_Table_Activated = False Then
Picture2.Cls
End If

'Did you know that the computer can do pixels faster horizontal than vertical? That's what
'I read in my 3D Game Programming book, so we will plot these pixels in a row by row fashion.

For B = 0 To Picture1.ScaleHeight
For A = 0 To Picture1.ScaleWidth

'Here is the slow method for obtaining RGB values from pixels out of both pics with 4
'divisions, 6 multiplications, and 6 subtractions per pixel.
    
If Look_Up_Table_Activated = False Then
Color_A.Blue = Int(Picture1.Point(A, B) / 65536)
Color_A.Green = Int((Picture1.Point(A, B) - (Color_A.Blue * 65536)) / 256)
Color_A.Red = Int(Picture1.Point(A, B) - (Color_A.Blue * 65536) - (Color_A.Green * 256))

Color_B.Blue = Int(Picture2.Point(A, B) / 65536)
Color_B.Green = Int((Picture2.Point(A, B) - (Color_B.Blue * 65536)) / 256)
Color_B.Red = Int(Picture2.Point(A, B) - (Color_B.Blue * 65536) - (Color_B.Green * 256))
End If

'But here, the math has already been done for you on these Look Up Tables. So no math is
'performed at all. Alright!!!
    
If Look_Up_Table_Activated = True Then
Color_A.Red = Red_A_Look_Up_Table(A, B)
Color_A.Green = Green_A_Look_Up_Table(A, B)
Color_A.Blue = Blue_A_Look_Up_Table(A, B)

Color_B.Red = Red_B_Look_Up_Table(A, B)
Color_B.Green = Green_B_Look_Up_Table(A, B)
Color_B.Blue = Blue_B_Look_Up_Table(A, B)
End If

'Here's the alpha that's similar to OpenGL's alpha only mine is within 0 to 255
'to go with the Red, Green, and Blue's 0 to 255's. It won't have effect if any one
'of the alphas are 0 (just like OpenGL) so fiddle with both scroll bars and see
'what happens.

Color_A.Alpha = HScroll1.Value
Color_B.Alpha = HScroll2.Value

'Next we keep the values in range so we can prevent an error.

If Color_A.Red > 255 Then Color_A.Red = 255
If Color_A.Green > 255 Then Color_A.Green = 255
If Color_A.Blue > 255 Then Color_A.Blue = 255
If Color_B.Red > 255 Then Color_B.Red = 255
If Color_B.Green > 255 Then Color_B.Green = 255
If Color_B.Blue > 255 Then Color_B.Blue = 255
If Color_A.Red < 0 Then Color_A.Red = 0
If Color_A.Green < 0 Then Color_A.Green = 0
If Color_A.Blue < 0 Then Color_A.Blue = 0
If Color_B.Red < 0 Then Color_B.Red = 0
If Color_B.Green < 0 Then Color_B.Green = 0
If Color_B.Blue < 0 Then Color_B.Blue = 0

'Oh boy, doesn't this suck for speed. It got the pixel from both pics with 4 divisions,
'6 multiplications, and 6 subtractions. Then it needs to do 12 divisions, 12 multiplications,
'6 subtractions, and 6 additions to blend the pics together. All together it's 16 divisions,
'18 multiplications, 12 subtractions, and 6 additions per pixel out of two (100 x 150) pictures
'(30000 pixels)...Yikes!!! But anyways, here's the formula to alpha blend pics together:

'Red = (Red_B - (Red_B * ((Alpha_A / 255) * (Alpha_B / 255))) + (Red_A * ((Alpha_A / 255) * (Alpha_B / 255))
'Green = (Green_B - (Green_B * ((Alpha_A / 255) * (Alpha_B / 255))) + (Green_A * ((Alpha_A / 255) * (Alpha_B / 255))
'Blue = (Blue_B - (Blue_B * ((Alpha_A / 255) * (Alpha_B / 255))) + (Blue_A * ((Alpha_A / 255) * (Alpha_B / 255))

If Look_Up_Table_Activated = False Then
Color_C.Red = (Color_B.Red - (Color_B.Red * ((Color_A.Alpha / 255) * (Color_B.Alpha / 255)))) + (Color_A.Red * ((Color_A.Alpha / 255) * (Color_B.Alpha / 255)))
Color_C.Green = (Color_B.Green - (Color_B.Green * ((Color_A.Alpha / 255) * (Color_B.Alpha / 255)))) + (Color_A.Green * ((Color_A.Alpha / 255) * (Color_B.Alpha / 255)))
Color_C.Blue = (Color_B.Blue - (Color_B.Blue * ((Color_A.Alpha / 255) * (Color_B.Alpha / 255)))) + (Color_A.Blue * ((Color_A.Alpha / 255) * (Color_B.Alpha / 255)))
End If

'But who needs math right. Math sucks nads. Might as well use precalculated Look Up Tables and
'cheat a bit. Damn I'm good!!! Now it does no math at all. Sweeeeeeeeeeeeeet!

If Look_Up_Table_Activated = True Then
Color_C.Red = Addition_Look_Up_Table(Subtraction_Look_Up_Table(Color_B.Red, Multiplication_Look_Up_Table(Color_B.Red, Percent_Look_Up_Table(Color_A.Alpha, Color_B.Alpha))), Multiplication_Look_Up_Table(Color_A.Red, Percent_Look_Up_Table(Color_A.Alpha, Color_B.Alpha)))
Color_C.Green = Addition_Look_Up_Table(Subtraction_Look_Up_Table(Color_B.Green, Multiplication_Look_Up_Table(Color_B.Green, Percent_Look_Up_Table(Color_A.Alpha, Color_B.Alpha))), Multiplication_Look_Up_Table(Color_A.Green, Percent_Look_Up_Table(Color_A.Alpha, Color_B.Alpha)))
Color_C.Blue = Addition_Look_Up_Table(Subtraction_Look_Up_Table(Color_B.Blue, Multiplication_Look_Up_Table(Color_B.Blue, Percent_Look_Up_Table(Color_A.Alpha, Color_B.Alpha))), Multiplication_Look_Up_Table(Color_A.Blue, Percent_Look_Up_Table(Color_A.Alpha, Color_B.Alpha)))
End If

'This seems like the only way to plot pixels 1 at a time in pure VB but there are faster methods
'out there without API calls and stuff. I'll stick with the PSet for now. It's still a tad slow cause
'of the fact we are plotting one pixel at a time but still it works within real time when using Look
'Up Tables, so it's all good. It will seem slow yet fast as a mofo on my LUT's cause of the limited
'speed the scroll bars can go when you hold the button down on it. So it should be faster if you use
'some better method to change the alpha values, such as Do loops. Do loops work better than timers
'cause of it's amazingly high speed but it's speed isn't as constant as timers are, yet timers are so
'slow. So if you want, you can use a Do loop to add up the alpha values Just dont forget to add DoEvents
'after the Do. I stuck with scroll bars to give you total control of the alpha changes in this program
'just to let you know. If you want, you can enable the SetPixel API one just one of the scroll bars and
'keep PSet on another and see a difference in speed. There really wouldn't be too much difference when
'my LUT's are on. That's pretty damn fast if I had just made my pure VB code almost as fast as this API
'call. And this API call is 5% faster than the SetPixelV API. Wow!!!

Picture2.PSet (A, B), RGB(Color_C.Red, Color_C.Green, Color_C.Blue)
'SetPixel Picture2.hDC, A, B, RGB(Color_C.Red, Color_C.Green, Color_C.Blue)
Next A
Next B
End Sub

'If you understood my code correctly, it's possible to not just blend 3 or more pics in one, but also any
'color! Optimize it any way you can. Add brightness and contrast with it, or if you are really good, add
'blur, mosaic, water effects, psychodelic rainbow effects, etc. Create your own look up tables for those to
'do em in real time. Have fun!!!

