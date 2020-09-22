Attribute VB_Name = "Vars"
'Sleep I didn't used but it's a great function.  It's basicall a pause.
'Call it like Sleep(1000) where 1000 = 1 second.
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'This is the function that gets the state of a key, pressed or unpressed.
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

'These are the variables used for the keys
'They can be anything you like
Public Up As Integer
Public Down As Integer
Public Lef As Integer
Public Rig As Integer
Public Shoot As Integer
Public En As Integer

Public speed As Single 'Veloity of object
Public Ang As Single 'Angle in degrees in which the object is moving
Public AngInc As Single 'Increment by which angle changes
Public Pie As Single 'Pi

'All DC Variables must be defined as Long
Public MechB As Long
Public MechW As Long
Public Mech1 As Long
Public Mech2 As Long
Public MechShootB As Long
Public MechShootW As Long
Public BackGround As Long

Public MechX As Single 'X position of Tank(was going to be a mechwarrior)
Public MechY As Single 'Y position of Tank
Public MechP As Integer 'Position on bitmap from which you are blitting

