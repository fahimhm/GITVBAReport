Sub insertrow()
  For l = 2 To Cells(Rows.Count, 1).End(xlUp).Row
    If Cells(l, 12).Value = "v" Then
      Rows(l).Insert Shift:=xlShiftDown
      For k = 1 To 5
        Cells(l+1, k).Copy Destination:=Cells(l, k)
      Next k
      Cells(l, 6).Value = "Planned"
      Cells(l, 7).Value = "no_wo"
      Cells(l, 8).Value = "Brief_preparation_travel_dll"
      If Cells(l, 5).Value = 1 Then
        Cells(l, 9).Value = Int(Cells(l+1, 9).Value) + 0.25
      ElseIf Cells(l, 5).Value = 2 Then
        If Cells(l, 4).Value = "normal" Then
          Cells(l, 9).Value = Int(Cells(l+1, 9).Value) + 0.58333
        ElseIf Cells(l, 4).Value = "overtime" Then
          Cells(l, 9).Value = Int(Cells(l+1, 9).Value) + 0.5625
        End If
      ElseIf Cells(l, 5).Value = 3 Then
        If Cells(l, 4).Value = "normal" Then
          Cells(l, 9).Value = Int(Cells(l+1, 9).Value) + 0.916666666666667
        ElseIf Cells(l, 4).Value = "overtime" Then
          Cells(l, 9).Value = Int(Cells(l+1, 9).Value) + 0.875
        End If
      ElseIf Cells(l, 5).Value = 4 Then
        If Cells(l, 4).Value = "normal" Then
          Cells(l, 9).Value = Int(Cells(l+1, 9).Value) + 0.3125
        ElseIf Cells(l, 4).Value = "overtime" Then
          Cells(l, 9).Value = Int(Cells(l+1, 9).Value) + 0.29166667
        End If
      ElseIf Cells(l, 5).Value = 5 Then
        Cells(l, 9).Value = Int(Cells(l+1, 9).Value) + 0.5
      End If
      Cells(l+1, 9).Copy Destination:=Cells(l, 10)
      l = l + 1
    ElseIf Cells(l, 12).Value = "vv" Then
      Rows(l+1).Insert Shift:=xlShiftDown
      For j = 1 To 5
        Cells(l, j).Copy Destination:=Cells(l+1,j)
      Next j
      Cells(l+1, 6).Value = "Planned"
      Cells(l+1, 7).Value = "no_wo"
      Cells(l+1, 8).Value = "Brief_preparation_travel_dll"
      Cells(l, 10).Copy Destination:=Cells(l+1, 9)
      If Cells(l+1, 5).Value = 1 Then
        If Cells(l+1, 4).Value = "normal" Then
          Cells(l+1, 10).Value = Int(Cells(l, 10).Value) + 0.60416667
        ElseIf Cells(l+1, 4).Value = "overtime" Then
          Cells(l+1, 10).Value = Int(Cells(l, 10).Value) + 0.54166667
        End If
      ElseIf Cells(l+1, 5).Value = 2 Then
        If Cells(l+1, 4).Value = "normal" Then
          Cells(l+1, 10).Value = Int(Cells(l, 10).Value) + 0.9375
        ElseIf Cells(l+1, 4).Value = "overtime" Then
          Cells(l+1, 10).Value = Int(Cells(l, 10).Value) + 0.875
        End If
      ElseIf Cells(l+1, 5).Value = 3 Then
        If Cells(l+1, 4).Value = "normal" Then
          Cells(l+1, 10).Value = Int(Cells(l, 10).Value) + 0.27083333
        ElseIf Cells(l+1, 4).Value = "overtime" Then
          Cells(l+1, 10).Value = Int(Cells(l, 10).Value) + 0.1875
        End If
      ElseIf Cells(l+1, 5).Value = 4 Then
        If Cells(l+1, 4).Value = "normal" Then
          Cells(l+1, 10).Value = Int(Cells(l, 10).Value) + 0.6875
        ElseIf Cells(l+1, 4).Value = "overtime" Then
          Cells(l+1, 10).Value = Int(Cells(l, 10).Value) + 0.60416667
        End If
      ElseIf Cells(l+1, 5).Value = 5 Then
        Cells(l+1, 10).Value = Int(Cells(l, 10).Value) + 0.8125
      End If
      l = l + 1
    End If
  Next l
End Sub
