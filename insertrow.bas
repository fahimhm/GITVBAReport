Sub insertrow()
  For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
    If Cells(i, 12).Value = "v" Then
      Rows(i).Insert Shift:=xlShiftDown
      For k = 1 To 5
        Cells(i+1, k).Copy Destination:=Cells(i, k)
      Next k
      Cells(i, 6).Value = "Planned"
      Cells(i, 7).Value = "no_wo"
      Cells(i, 8).Value = "Brief_preparation_travel_dll"
      If Cells(i, 5).Value = 1 Then
        Cells(i, 9).Value = Int(Cells(i+1, 9).Value) + 0.25
      ElseIf Cells(i, 5).Value = 2 Then
        If Cells(i, 4).Value = "normal" Then
          Cells(i, 9).Value = Int(Cells(i+1, 9).Value) + 0.58333
        ElseIf Cells(i, 4).Value = "overtime" Then
          Cells(i, 9).Value = Int(Cells(i+1, 9).Value) + 0.5625
        End If
      ElseIf Cells(i, 5).Value = 3 Then
        If Cells(i, 4).Value = "normal" Then
          Cells(i, 9).Value = Int(Cells(i+1, 9).Value) + 0.916666666666667
        ElseIf Cells(i, 4).Value = "overtime" Then
          Cells(i, 9).Value = Int(Cells(i+1, 9).Value) + 0.875
        End If
      ElseIf Cells(i, 5).Value = 4 Then
        If Cells(i, 4).Value = "normal" Then
          Cells(i, 9).Value = Int(Cells(i+1, 9).Value) + 0.3125
        ElseIf Cells(i, 4).Value = "overtime" Then
          Cells(i, 9).Value = Int(Cells(i+1, 9).Value) + 0.29166667
        End If
      ElseIf Cells(i, 5).Value = 5 Then
        Cells(i, 9).Value = Int(Cells(i+1, 9).Value) + 0.5
      End If
      Cells(i+1, 9).Copy Destination:=Cells(i, 10)
      i = i + 1
    ElseIf Cells(i, 12).Value = "vv" Then
      Rows(i+1).Insert Shift:=xlShiftDown
      For j = 1 To 5
        Cells(i, j).Copy Destination:=Cells(i+1,j)
      Next j
      Cells(i+1, 6).Value = "Planned"
      Cells(i+1, 7).Value = "no_wo"
      Cells(i+1, 8).Value = "Brief_preparation_travel_dll"
      Cells(i, 10).Copy Destination:=Cells(i+1, 9)
      If Cells(i+1, 5).Value = 1 Then
        If Cells(i+1, 4).Value = "normal" Then
          Cells(i+1, 10).Value = Int(Cells(i, 10).Value) + 0.60416667
        ElseIf Cells(i+1, 4).Value = "overtime" Then
          Cells(i+1, 10).Value = Int(Cells(i, 10).Value) + 0.54166667
        End If
      ElseIf Cells(i+1, 5).Value = 2 Then
        If Cells(i+1, 4).Value = "normal" Then
          Cells(i+1, 10).Value = Int(Cells(i, 10).Value) + 0.9375
        ElseIf Cells(i+1, 4).Value = "overtime" Then
          Cells(i+1, 10).Value = Int(Cells(i, 10).Value) + 0.875
        End If
      ElseIf Cells(i+1, 5).Value = 3 Then
        If Cells(i+1, 4).Value = "normal" Then
          Cells(i+1, 10).Value = Int(Cells(i, 10).Value) + 0.27083333
        ElseIf Cells(i+1, 4).Value = "overtime" Then
          Cells(i+1, 10).Value = Int(Cells(i, 10).Value) + 0.1875
        End If
      ElseIf Cells(i+1, 5).Value = 4 Then
        If Cells(i+1, 4).Value = "normal" Then
          Cells(i+1, 10).Value = Int(Cells(i, 10).Value) + 0.6875
        ElseIf Cells(i+1, 4).Value = "overtime" Then
          Cells(i+1, 10).Value = Int(Cells(i, 10).Value) + 0.60416667
        End If
      ElseIf Cells(i+1, 5).Value = 5 Then
        Cells(i+1, 10).Value = Int(Cells(i, 10).Value) + 0.8125
      End If
      i = i + 1
    End If
  Next i
End Sub
