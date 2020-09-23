Attribute VB_Name = "modAI"
Type XYPos
    X As Single
    Y As Single
End Type
Public Function AIChase(Chased As Image, Chasing As Image, MaxUnits As Integer) As Boolean
    Dim ChaseDir As Byte, NumOfUnits As Integer
    Randomize
    ChaseDir = Int(Rnd * 2)
    If ChaseDir <= 0 Then
        If Chasing.Left < Chased.Left Then
            NumOfUnits = Int(Rnd * MaxUnits)
            Chasing.Left = Chasing.Left + NumOfUnits
        Else
            NumOfUnits = Int(Rnd * MaxUnits)
            Chasing.Left = Chasing.Left - NumOfUnits
        End If
        GoTo DefineChased
    End If
    If ChaseDir >= 1 Then
        If Chasing.Top < Chased.Top Then
            NumOfUnits = Int(Rnd * MaxUnits)
            Chasing.Top = Chasing.Top + NumOfUnits
        Else
            NumOfUnits = Int(Rnd * MaxUnits)
            Chasing.Top = Chasing.Top - NumOfUnits
        End If
        GoTo DefineChased
    End If
    Exit Function
    
DefineChased:
    AIChase = Collision(Chasing.Left, Chasing.Top, Chased.Left, Chased.Top, Chasing.Width, Chasing.Height, Chased.Width, Chased.Height)
        
End Function
Private Function AIChase1(Chased As XYPos, Chasing As XYPos, MaxUnits As Integer) As XYPos
    Dim ChaseDir As Byte, NumOfUnits As Integer
    Randomize
    ChaseDir = Int(Rnd * 1)
    If ChaseDir = 0 Then
        If Chasing.X < Chased.X Then
            NumOfUnits = Int(Rnd * MaxUnits)
            Chasing.X = Chasing.X + NumOfUnits
        Else
            NumOfUnits = Int(Rnd * MaxUnits)
            Chasing.X = Chasing.X - NumOfUnits
        End If
        GoTo DefineChased
    End If
    If ChaseDir = 1 Then
        If Chasing.Y < Chased.Y Then
            NumOfUnits = Int(Rnd * MaxUnits)
            Chasing.Y = Chasing.Y + NumOfUnits
        Else
            NumOfUnits = Int(Rnd * MaxUnits)
            Chasing.Y = Chasing.Y - NumOfUnits
        End If
        GoTo DefineChased
    End If
    Exit Function
    
DefineChased:
    AIChase1.X = Chasing.X
    AIChase1.Y = Chasing.Y
        
End Function
Public Function AIFollow(ChasedLastPos As XYPos, Chased As Image, Chasing As Image, MaxUnits As Byte) As Boolean
    Dim ChasedXY As XYPos, ChasingXY As XYPos, NewPos As XYPos
    ChasedXY.X = Chased.Left
    ChasedXY.Y = Chased.Top
    ChasingXY.X = Chasing.Left
    ChasingXY.Y = Chasing.Top
    NewPos = AIChase1(ChasedLastPos, Chasing, MaxUnits)
    Chasing.Left = NewPos.X
    Chasing.Top = NewPos.Y
    AIFollow = AIChase(Chased, Chasing, MaxUnits)
End Function

Public Function Collision(XPOSITION1 As Integer, YPOSITION1 As Integer, XPOSITION2 As Integer, YPOSITION2 As Integer, BOXSIZEX1 As Integer, BOXSIZEY1 As Integer, BOXSIZEX2 As Integer, BOXSIZEY2 As Integer) As Boolean
     If XPOSITION1 > XPOSITION2 - ((BOXSIZEX2 / 640) * 639) Then
          If XPOSITION1 < XPOSITION2 + ((BOXSIZEX1 / 640) * 639) Then
               If YPOSITION1 > YPOSITION2 - ((BOXSIZEY2 / 480) * 479) Then
                    If YPOSITION1 < YPOSITION2 + ((BOXSIZEY1 / 480) * 479) Then
                         Collision = True
                    End If
               End If
          End If
     End If
     If XPOSITION1 > XPOSITION2 - ((BOXSIZEX1 / 640) * 639) Then
          If XPOSITION1 < XPOSITION2 + ((BOXSIZEX2 / 640) * 639) Then
               If YPOSITION1 > YPOSITION2 - ((BOXSIZEY1 / 480) * 479) Then
                    If YPOSITION1 < YPOSITION2 + ((BOXSIZEY2 / 480) * 479) Then
                         Collision = True
                    End If
               End If
          End If
     End If
End Function


