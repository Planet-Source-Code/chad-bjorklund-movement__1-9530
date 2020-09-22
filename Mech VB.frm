VERSION 5.00
Begin VB.Form Mech 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   12960
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   17280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12960
   ScaleWidth      =   17280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   8040
      Top             =   6240
   End
   Begin VB.PictureBox Blank 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   -300
      Width           =   255
   End
End
Attribute VB_Name = "Mech"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Pie = 3.14159265358979
    AngInc = 45
    Ang = 0
    'While statements make sure bitmap loads every time, with the exception
    'of not enough memory.
    While Mech1 = 0
        Mech1 = GenerateDC(App.Path & "\Tank3.bmp")
    Wend
    While Mech2 = 0
        Mech2 = GenerateDC(App.Path & "\TankW2.bmp")
    Wend
    While MechShootB = 0
        MechShootB = GenerateDC(App.Path & "\Tankshoot2.bmp")
    Wend
    While MechShootW = 0
        MechShootW = GenerateDC(App.Path & "\TankShootW.bmp")
    Wend
End Sub

Public Function Game()
'Make sure AutoRedraw is set to True

    Do
        'This is neccesary, I'm not sure why.
        Mech.Refresh
        
        'This command clears everything that is not an object on the
        'form.  Basically it's an eraser. Try commenting it out, it
        'is a pretty neat effect. If you are blitting a background
        'comment this line out
        Mech.Cls
        
        'This checks the state of the key in the ( ).  If it is pressed
        'the variable that is set equal to the function will have a
        'value of -32767 or less. If it isn't, it will have a value of 0
        'You can also use KeyDown in place of this function, just make
        'sure to include the DoEvents command in the loop.
        
        'IF THE KEYS RE BEING READ TOO FAST, which on most of the faster
        'machines will be a problem, remove the < signs in on the "If"
        'statements checking the key variables. Ex. "If Up = -32767 Then"
        
        Up = GetAsyncKeyState(vbKeyUp)
        Down = GetAsyncKeyState(vbKeyDown)
        Lef = GetAsyncKeyState(vbKeyLeft)
        Rig = GetAsyncKeyState(vbKeyRight)
        Shoot = GetAsyncKeyState(vbKeyControl)
        En = GetAsyncKeyState(vbKeyEscape)
         
        
        If Up <= -32767 Then
            'This increments your speed forward, with a max of 3
            If speed < 4 Then speed = speed + 1
        End If
        
        If Down <= -32767 Then
            'This increments your speed backwards with a min of -3
            If speed > -4 Then speed = speed - 1
        End If
        
        If Lef <= -32767 Then
            'This increments the angle at which your traveling
            Ang = Ang - AngInc
            'This increments the position on the bitmap where you are going
            'to blit from(In pixels). If you don't understand blitting look
            'at the bitmap, it might help.
            MechP = MechP - 102
            'Once you have reached the beginning of the bitmap, this sets you
            'at the end.
            If MechP < 0 Then MechP = 714
        End If
        
        If Rig <= -32767 Then
            'This increments the angle at which your traveling
            Ang = Ang + AngInc
            'This increments the position on the bitmap where you are going
            'to blit from(In pixels). If you don't understand blitting look
            'at the bitmap, it might help.
            MechP = MechP + 102
            'Once you have reached the end of the bitmap, this starts you
            'at the beginning again.
            If MechP > 714 Then MechP = 0
        End If
        
        If Shoot <= -32767 Then
            'This switches the DC variables that are going to be blitted to
            'the picture of the tank firing.  Once the loop has gone through
            'and shoot does not = -32767, the else statement will switch it
            'back to the normal tank picture. If you want you can put in a
            'little timer to make the shoot picture last longer.
            MechB = MechShootB
            MechW = MechShootW
            'Once you understand the motion part, this isn't too confusing
            'I just made it so you move in the opposite direction of which
            'your firing, in other words, recoil.  The "5 - speed" part
            'makes it so there is less recoil at highr speeds
            MechX = MechX - Cos(Ang / 180 * Pie) * (5 - speed)
            MechY = MechY - Sin(Ang / 180 * Pie) * (5 - speed)
        Else
            MechB = Mech1
            MechW = Mech2
        End If
        If En = -32767 Then
            'You must delete and DC's you created, otherwise you
            'will have memory loss
             DeleteGeneratedDC Mech1
             DeleteGeneratedDC Mech2
             DeleteGeneratedDC MechB
             DeleteGeneratedDC MechW
             DeleteGeneratedDC MechShootB
             DeleteGeneratedDC MechShootW
            End
        End If
        'Just in case you are doing a lot of turning, this will prevent
        'an overflow on the Ang variable.  For all you trig drop outs,
        '360 and -360 degrees are = to 0 degrees.
        If Ang = 360 Or Ang = -360 Then Ang = 0
        
        'This might seem confusing but it's not.  All I am doing is taking
        'the Sin and Cos of the angle at which your pointing to figure out
        'how far to move you in the X and Y coordinates.  The "/ 180 * Pie"
        'part is just converting from degrees to radians.  The "* speed"
        'part multiplies the answer by a number to make you go faster or
        'slower in that direction.
        MechX = MechX + Cos(Ang / 180 * Pie) * speed
        MechY = MechY + Sin(Ang / 180 * Pie) * speed
        
        'If you have a background you want to use, you can use this line
        'of code. Simply generate a DC like the others using the variable
        'BackGround and uncommenting this line. If you use this, make sure
        'you comment out the Mech.cls line.
        'BitBlt Mech.hdc, 0, 0, Mech.ScaleWidth, Mech.ScaleHeight, BackGround, 0, 0, vbSrcCopy
        
        BitBlt Mech.hdc, MechX, MechY, 102, 72, MechW, MechP, 0, vbSrcAnd
        BitBlt Mech.hdc, MechX, MechY, 102, 72, MechB, MechP, 0, vbSrcPaint
    Loop
End Function

Private Sub Timer1_Timer()
    Game
    'I call the function in a timer because the form needs time to load.
    'If you call it straight from form load, the form will not show unless
    'you use both form.refresh and form.show in the loop.  I use the Timer
    'because I think it make the program run faster.
End Sub
