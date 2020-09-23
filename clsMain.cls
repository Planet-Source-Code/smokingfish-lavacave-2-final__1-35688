VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public SnakeSpeed As Integer
Public ItemX As Integer
Public ItemY As Integer
Public EnemyStop As Boolean
Public SnakePositionX As Integer
Public SnakePositionY As Integer
Public EnemyFrame As Integer
Public SnakeOver As Boolean
Public BonusItem As Integer
Public WorldGravity As Integer
Public SnakeDirection As Boolean
Public SnakeAcceleration As Integer
Private EnemyX(0 To 2) As Integer, EnemyY(0 To 2) As Integer
Private LX(0 To 250) As Long, RX(0 To 250) As Long
Public SideMove As Integer
Public WorldWidth As Long

Private Sub Cycle()
    Dim a As Integer
    For a = 250 To 1 Step -1
        LX(a) = LX(a - 1)
        LX(0) = ((150 - WorldWidth) / 2) + SideMove
        RX(a) = RX(a - 1)
        RX(0) = LX(0) + WorldWidth
    Next a
End Sub

Public Sub CheckEnemys()
For i = 0 To 2
If frmMain.Point(EnemyX(i), EnemyY(i)) = vbBlue Then
StopMP3
PlayMP3 App.Path & "\Sound\drop9.mp3"
EnemyY(i) = 100
Randomize Timer
EnemyX(i) = Int(Rnd * 2500) + 300
End If
Next i
End Sub

Public Sub CheckItems()
If frmMain.Point(ItemX, ItemY) = vbBlue Then
StopMP3
PlayMP3 App.Path & "\Sound\drop9.mp3"
Randomize Timer
ItemY = 200
Game.ItemX = Int(Rnd * 2500) + 300
BonusItem = Int(Rnd * 4) + 1
End If
End Sub

Public Sub DrawitemsShadow()
ItemY = ItemY + 4
If BonusItem = 1 Then
For i = 1 To 180 Step 15
frmMain.Circle (ItemX + 150, ItemY + 120), i, RGB(2, 2, 2)
Next i
End If
'-200
If BonusItem = 2 Then
For i = 1 To 180 Step 15
frmMain.Circle (ItemX + 150, ItemY + 120), i, RGB(2, 2, 2)
Next i
End If
'+500
If BonusItem = 3 Then
For i = 1 To 180 Step 15
frmMain.Circle (ItemX + 150, ItemY + 120), i, RGB(2, 2, 2)
Next i
End If
'-Wall
If BonusItem = 4 Then
For i = 1 To 180 Step 15
frmMain.Circle (ItemX + 150, ItemY + 120), i, RGB(2, 2, 2)
Next i
End If
'+Wall
If BonusItem = 5 Then
For i = 1 To 180 Step 15
frmMain.Circle (ItemX + 150, ItemY + 120), i, RGB(2, 2, 2)
Next i
End If
'Dead
If BonusItem = 6 Then
For i = 1 To 180 Step 15
frmMain.Circle (ItemX + 150, ItemY + 120), i, RGB(2, 2, 2)
Next i
End If
'SlowDown
If BonusItem = 7 Then
For i = 1 To 180 Step 15
frmMain.Circle (ItemX + 150, ItemY + 120), i, RGB(2, 2, 2)
Next i
End If
'SpeedUp
If BonusItem = 8 Then
For i = 1 To 180 Step 15
frmMain.Circle (ItemX + 150, ItemY + 120), i, RGB(2, 2, 2)
Next i
End If
'StopEnemy
If BonusItem = 9 Then
For i = 1 To 180 Step 15
frmMain.Circle (ItemX + 150, ItemY + 120), i, RGB(2, 2, 2)
Next i
End If
End Sub

Public Sub DrawItems()
'+200
If BonusItem = 1 Then
BitBlt frmMain.HDC, TwipsToPixelsX(ItemX), TwipsToPixelsY(ItemY), 20, 16, frmMain.pic200pmask.HDC, 0, 0, MERGEPAINT
BitBlt frmMain.HDC, TwipsToPixelsX(ItemX), TwipsToPixelsY(ItemY), 20, 16, frmMain.pic200p.HDC, 0, 0, SRCAND
End If
'-200
If BonusItem = 2 Then
BitBlt frmMain.HDC, TwipsToPixelsX(ItemX), TwipsToPixelsY(ItemY), 20, 16, frmMain.pic200mask.HDC, 0, 0, MERGEPAINT
BitBlt frmMain.HDC, TwipsToPixelsX(ItemX), TwipsToPixelsY(ItemY), 20, 16, frmMain.pic200.HDC, 0, 0, SRCAND
End If
'+500
If BonusItem = 3 Then
BitBlt frmMain.HDC, TwipsToPixelsX(ItemX), TwipsToPixelsY(ItemY), 20, 16, frmMain.pic500pmask.HDC, 0, 0, MERGEPAINT
BitBlt frmMain.HDC, TwipsToPixelsX(ItemX), TwipsToPixelsY(ItemY), 20, 16, frmMain.pic500p.HDC, 0, 0, SRCAND
End If
'+Wall
If BonusItem = 4 Then
BitBlt frmMain.HDC, TwipsToPixelsX(ItemX), TwipsToPixelsY(ItemY), 20, 16, frmMain.picwallbackmask.HDC, 0, 0, MERGEPAINT
BitBlt frmMain.HDC, TwipsToPixelsX(ItemX), TwipsToPixelsY(ItemY), 20, 16, frmMain.picwallback.HDC, 0, 0, SRCAND
End If
'-Wall
If BonusItem = 5 Then
BitBlt frmMain.HDC, TwipsToPixelsX(ItemX), TwipsToPixelsY(ItemY), 20, 16, frmMain.picwallsmask.HDC, 0, 0, MERGEPAINT
BitBlt frmMain.HDC, TwipsToPixelsX(ItemX), TwipsToPixelsY(ItemY), 20, 16, frmMain.picWalls.HDC, 0, 0, SRCAND
End If
'DEAD
If BonusItem = 6 Then
BitBlt frmMain.HDC, TwipsToPixelsX(ItemX), TwipsToPixelsY(ItemY), 20, 16, frmMain.pickillmask.HDC, 0, 0, MERGEPAINT
BitBlt frmMain.HDC, TwipsToPixelsX(ItemX), TwipsToPixelsY(ItemY), 20, 16, frmMain.pickill.HDC, 0, 0, SRCAND
End If
'SlowDown
If BonusItem = 7 Then
BitBlt frmMain.HDC, TwipsToPixelsX(ItemX), TwipsToPixelsY(ItemY), 20, 16, frmMain.picspeedmask.HDC, 0, 0, MERGEPAINT
BitBlt frmMain.HDC, TwipsToPixelsX(ItemX), TwipsToPixelsY(ItemY), 20, 16, frmMain.picSpeed.HDC, 0, 0, SRCAND
End If
'SpeedUP
If BonusItem = 8 Then
BitBlt frmMain.HDC, TwipsToPixelsX(ItemX), TwipsToPixelsY(ItemY), 20, 16, frmMain.picspeedslowmask.HDC, 0, 0, MERGEPAINT
BitBlt frmMain.HDC, TwipsToPixelsX(ItemX), TwipsToPixelsY(ItemY), 20, 16, frmMain.picspeedslow.HDC, 0, 0, SRCAND
End If
'enemy slow
If BonusItem = 9 Then
BitBlt frmMain.HDC, TwipsToPixelsX(ItemX), TwipsToPixelsY(ItemY), 20, 16, frmMain.picenemyslowmask.HDC, 0, 0, MERGEPAINT
BitBlt frmMain.HDC, TwipsToPixelsX(ItemX), TwipsToPixelsY(ItemY), 20, 16, frmMain.picenemyslow.HDC, 0, 0, SRCAND
End If
End Sub

Public Sub DrawEnemys()
If EnemyStop = False Then
EnemyY(0) = EnemyY(0) + 3
EnemyY(1) = EnemyY(1) + 5
EnemyY(2) = EnemyY(2) + 2
End If
'GradientCircle frmMain, EnemyX(0), EnemyY(0), 50, 200, 50, 30, 4, True, False
'GradientCircle frmMain, EnemyX(1), EnemyY(1), 50, 200, 50, 30, 4, True, False
'GradientCircle frmMain , EnemyX(2), EnemyY(2), 50, 200, 50, 30, 4, True, False
If EnemyFrame = 0 Then
BitBlt frmMain.HDC, TwipsToPixelsX(EnemyX(0)), TwipsToPixelsY(EnemyY(0)), 12, 12, frmMain.picenemymask.HDC, 0, 0, MERGEPAINT
BitBlt frmMain.HDC, TwipsToPixelsX(EnemyX(0)), TwipsToPixelsY(EnemyY(0)), 12, 12, frmMain.picenemy.HDC, 0, 0, SRCAND

BitBlt frmMain.HDC, TwipsToPixelsX(EnemyX(1)), TwipsToPixelsY(EnemyY(1)), 12, 12, frmMain.picenemymask.HDC, 0, 0, MERGEPAINT
BitBlt frmMain.HDC, TwipsToPixelsX(EnemyX(1)), TwipsToPixelsY(EnemyY(1)), 12, 12, frmMain.picenemy.HDC, 0, 0, SRCAND

BitBlt frmMain.HDC, TwipsToPixelsX(EnemyX(2)), TwipsToPixelsY(EnemyY(2)), 12, 12, frmMain.picenemymask.HDC, 0, 0, MERGEPAINT
BitBlt frmMain.HDC, TwipsToPixelsX(EnemyX(2)), TwipsToPixelsY(EnemyY(2)), 12, 12, frmMain.picenemy.HDC, 0, 0, SRCAND
End If
If EnemyFrame = 1 Then
BitBlt frmMain.HDC, TwipsToPixelsX(EnemyX(0)), TwipsToPixelsY(EnemyY(0)), 12, 12, frmMain.picenemymask.HDC, 12, 0, MERGEPAINT
BitBlt frmMain.HDC, TwipsToPixelsX(EnemyX(0)), TwipsToPixelsY(EnemyY(0)), 12, 12, frmMain.picenemy.HDC, 12, 0, SRCAND

BitBlt frmMain.HDC, TwipsToPixelsX(EnemyX(1)), TwipsToPixelsY(EnemyY(1)), 12, 12, frmMain.picenemymask.HDC, 12, 0, MERGEPAINT
BitBlt frmMain.HDC, TwipsToPixelsX(EnemyX(1)), TwipsToPixelsY(EnemyY(1)), 12, 12, frmMain.picenemy.HDC, 12, 0, SRCAND

BitBlt frmMain.HDC, TwipsToPixelsX(EnemyX(2)), TwipsToPixelsY(EnemyY(2)), 12, 12, frmMain.picenemymask.HDC, 12, 0, MERGEPAINT
BitBlt frmMain.HDC, TwipsToPixelsX(EnemyX(2)), TwipsToPixelsY(EnemyY(2)), 12, 12, frmMain.picenemy.HDC, 12, 0, SRCAND
End If
If EnemyFrame = 2 Then
BitBlt frmMain.HDC, TwipsToPixelsX(EnemyX(0)), TwipsToPixelsY(EnemyY(0)), 12, 12, frmMain.picenemymask.HDC, 24, 0, MERGEPAINT
BitBlt frmMain.HDC, TwipsToPixelsX(EnemyX(0)), TwipsToPixelsY(EnemyY(0)), 12, 12, frmMain.picenemy.HDC, 24, 0, SRCAND

BitBlt frmMain.HDC, TwipsToPixelsX(EnemyX(1)), TwipsToPixelsY(EnemyY(1)), 12, 12, frmMain.picenemymask.HDC, 24, 0, MERGEPAINT
BitBlt frmMain.HDC, TwipsToPixelsX(EnemyX(1)), TwipsToPixelsY(EnemyY(1)), 12, 12, frmMain.picenemy.HDC, 24, 0, SRCAND

BitBlt frmMain.HDC, TwipsToPixelsX(EnemyX(2)), TwipsToPixelsY(EnemyY(2)), 12, 12, frmMain.picenemymask.HDC, 24, 0, MERGEPAINT
BitBlt frmMain.HDC, TwipsToPixelsX(EnemyX(2)), TwipsToPixelsY(EnemyY(2)), 12, 12, frmMain.picenemy.HDC, 24, 0, SRCAND
End If
If EnemyFrame = 3 Then
BitBlt frmMain.HDC, TwipsToPixelsX(EnemyX(0)), TwipsToPixelsY(EnemyY(0)), 12, 12, frmMain.picenemymask.HDC, 36, 0, MERGEPAINT
BitBlt frmMain.HDC, TwipsToPixelsX(EnemyX(0)), TwipsToPixelsY(EnemyY(0)), 12, 12, frmMain.picenemy.HDC, 36, 0, SRCAND

BitBlt frmMain.HDC, TwipsToPixelsX(EnemyX(1)), TwipsToPixelsY(EnemyY(1)), 12, 12, frmMain.picenemymask.HDC, 36, 0, MERGEPAINT
BitBlt frmMain.HDC, TwipsToPixelsX(EnemyX(1)), TwipsToPixelsY(EnemyY(1)), 12, 12, frmMain.picenemy.HDC, 36, 0, SRCAND

BitBlt frmMain.HDC, TwipsToPixelsX(EnemyX(2)), TwipsToPixelsY(EnemyY(2)), 12, 12, frmMain.picenemymask.HDC, 36, 0, MERGEPAINT
BitBlt frmMain.HDC, TwipsToPixelsX(EnemyX(2)), TwipsToPixelsY(EnemyY(2)), 12, 12, frmMain.picenemy.HDC, 36, 0, SRCAND
End If
End Sub

Private Sub SidesChange()
    SideMove = SideMove + Round((Rnd * 2), 1) - 1
    If SideMove > 100 Then SideMove = 100
    If SideMove < 5 Then SideMove = 5
End Sub

Public Sub DrawSnakeShadow()
If frmMain.lstX.ListCount > 4 Then
BitBlt frmMain.HDC, TwipsToPixelsX(frmMain.lstX.List(0)), TwipsToPixelsY(frmMain.lstY.List(0)), 12, 12, frmMain.picPlayermask.HDC, 60, 0, MERGEPAINT
BitBlt frmMain.HDC, TwipsToPixelsX(frmMain.lstX.List(0)), TwipsToPixelsY(frmMain.lstY.List(0)), 12, 12, frmMain.picPlayer.HDC, 60, 0, SRCAND

BitBlt frmMain.HDC, TwipsToPixelsX(frmMain.lstX.List(1)), TwipsToPixelsY(frmMain.lstY.List(1)), 12, 12, frmMain.picPlayermask.HDC, 48, 0, MERGEPAINT
BitBlt frmMain.HDC, TwipsToPixelsX(frmMain.lstX.List(1)), TwipsToPixelsY(frmMain.lstY.List(1)), 12, 12, frmMain.picPlayer.HDC, 48, 0, SRCAND

BitBlt frmMain.HDC, TwipsToPixelsX(frmMain.lstX.List(2)), TwipsToPixelsY(frmMain.lstY.List(2)), 12, 12, frmMain.picPlayermask.HDC, 36, 0, MERGEPAINT
BitBlt frmMain.HDC, TwipsToPixelsX(frmMain.lstX.List(2)), TwipsToPixelsY(frmMain.lstY.List(2)), 12, 12, frmMain.picPlayer.HDC, 36, 0, SRCAND

BitBlt frmMain.HDC, TwipsToPixelsX(frmMain.lstX.List(3)), TwipsToPixelsY(frmMain.lstY.List(3)), 12, 12, frmMain.picPlayermask.HDC, 24, 0, MERGEPAINT
BitBlt frmMain.HDC, TwipsToPixelsX(frmMain.lstX.List(3)), TwipsToPixelsY(frmMain.lstY.List(3)), 12, 12, frmMain.picPlayer.HDC, 24, 0, SRCAND

BitBlt frmMain.HDC, TwipsToPixelsX(frmMain.lstX.List(4)), TwipsToPixelsY(frmMain.lstY.List(4)), 12, 12, frmMain.picPlayermask.HDC, 12, 0, MERGEPAINT
BitBlt frmMain.HDC, TwipsToPixelsX(frmMain.lstX.List(4)), TwipsToPixelsY(frmMain.lstY.List(4)), 12, 12, frmMain.picPlayer.HDC, 12, 0, SRCAND
 
BitBlt frmMain.HDC, TwipsToPixelsX(Game.SnakePositionX), TwipsToPixelsY(Game.SnakePositionY), 12, 12, frmMain.picPlayermask.HDC, 0, 0, MERGEPAINT
BitBlt frmMain.HDC, TwipsToPixelsX(Game.SnakePositionX), TwipsToPixelsY(Game.SnakePositionY), 12, 12, frmMain.picPlayer.HDC, 0, 0, SRCAND

End If
End Sub

Public Sub DrawSnake()
'GradientCircle frmMain, SnakePositionX, SnakePositionY, 60, 0, 200, 200, 5, True, False
'TransparentBlt frmMain.HDC, TwipsToPixelsX(Game.SnakePositionX), TwipsToPixelsY(Game.SnakePositionY), 12, 12, frmMain.picPlayer.HDC, 0, 0, 12, 12, &HFF00FF
BitBlt frmMain.HDC, TwipsToPixelsX(Game.SnakePositionX), TwipsToPixelsY(Game.SnakePositionY), 12, 12, frmMain.picPlayermask.HDC, 0, 0, MERGEPAINT
BitBlt frmMain.HDC, TwipsToPixelsX(Game.SnakePositionX), TwipsToPixelsY(Game.SnakePositionY), 12, 12, frmMain.picPlayer.HDC, 0, 0, SRCAND
End Sub

Public Sub DummiWorld()
Randomize Timer
EnemyX(0) = Int(Rnd * 2500) + 300
EnemyX(1) = Int(Rnd * 2500) + 300
EnemyX(2) = Int(Rnd * 2500) + 300

EnemyY(0) = 100
EnemyY(1) = 100
EnemyY(2) = 100
Randomize Timer
ItemY = 100
Game.ItemX = Int(Rnd * 2500) + 300
BonusItem = Int(Rnd * 4) + 1
    For a = 1 To 250
        LX(a) = (24 + (Rnd * 1))
        RX(a) = LX(a) + WorldWidth
    Next a
End Sub

Public Sub DrawWorld()
Dim a As Integer, b As Integer
    For a = 1 To 250 Step 1
        RX(a) = LX(a) + WorldWidth
        frmMain.Line (a * 12 + 220, 200)-(a * 12 + 220, 10 + LX(a) * 12), vbRed
        frmMain.Line (a * 12 + 220, RX(a) * 12)-(a * 12 + 220, 2500), vbBlue
    Next a
    Cycle
SidesChange
'GradientLine frmMain, 200, 200, 200, 2500, 100, 200, 100, 5, False
'GradientLine frmMain, 200, 200, 3200, 200, 100, 200, 100, 5, False
'GradientLine frmMain, 200, 2500, 3200, 2500, 100, 200, 100, 5, False
'GradientLine frmMain, 3200, 200, 3200, 2500, 100, 200, 100, 5, False
'GradientCircle frmMain, 200, 200, 50, 100, 200, 200, 5, True, False
'GradientCircle frmMain, 3200, 200, 50, 100, 200, 200, 5, True, False
'GradientCircle frmMain, 200, 2500, 50, 100, 200, 200, 5, True, False
'GradientCircle frmMain, 3200, 2500, 50, 100, 200, 200, 5, True, False
BitBlt frmMain.HDC, 0, 0, frmMain.Width, frmMain.Height, frmMain.Picture4.HDC, 0, 0, MERGEPAINT
BitBlt frmMain.HDC, 0, 0, frmMain.Width, frmMain.Height, frmMain.Picture3.HDC, 0, 0, SRCAND
End Sub

Public Function GradientLine(Obj, Xpos1, Ypos1, xpos2, Ypos2, ColorRed, ColorGreen, ColorBlue, BorderWidth, invert As Boolean)
    TBorder = Obj.DrawWidth
    If BorderWidth < 1 Then BorderWidth = 1
    For q = 1 To BorderWidth
        twidth = BorderWidth + 1 - q
        rval = Val(ColorRed / BorderWidth * q)
        If rval > 255 Then rval = 255
        gval = Val(ColorGreen / BorderWidth * q)
        If gval > 255 Then gval = 255
        bval = Val(ColorBlue / BorderWidth * q)
        If bval > 255 Then bval = 255
        If invert Then
            rval = 255 - rval
            gval = 255 - gval
            bval = 255 - bval
        End If
        Obj.DrawWidth = twidth
        Obj.Line (Xpos1, Ypos1)-(xpos2, Ypos2), RGB(rval, gval, bval)
    Next q
    Obj.DrawWidth = TBorder
End Function

Public Function GradientCircle(Obj, Xpos, Ypos, Radius, ColorRed, ColorGreen, ColorBlue, BorderWidth, Fill As Boolean, invert As Boolean)
    TBorder = Obj.DrawWidth
    tfill = Obj.FillStyle
    tfc = Obj.FillColor
    If Fill Then
        Obj.FillStyle = 0
        If Not invert Then Obj.FillColor = RGB(ColorRed, ColorGreen, ColorBlue) Else Obj.FillColor = RGB(255 - ColorRed, 255 - ColorGreen, 255 - ColorBlue)
    End If
    If BorderWidth < 1 Then BorderWidth = 1
    For q = 1 To BorderWidth
        twidth = BorderWidth + 1 - q
        rval = Val(ColorRed / BorderWidth * q)
        If rval > 255 Then rval = 255
        gval = Val(ColorGreen / BorderWidth * q)
        If gval > 255 Then gval = 255
        bval = Val(ColorBlue / BorderWidth * q)
        If bval > 255 Then bval = 255
        If invert Then
            rval = 255 - rval
            gval = 255 - gval
            bval = 255 - bval
        End If
        Obj.DrawWidth = twidth
        Obj.Circle (Xpos, Ypos), Radius, RGB(rval, gval, bval)
    Next q
    Obj.FillStyle = tfill
    Obj.FillColor = tfc
    Obj.DrawWidth = TBorder
End Function
