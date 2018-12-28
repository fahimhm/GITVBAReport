Sub insertrow()
  For i = 2 To Cells(Rows.Count, 1).End(xlUp).Row
    If Cells(i, 13).Value = "v" Then
      Rows(i).Insert Shift:=xlShiftDown
      For k = 1 To 6
        Cells(i+1, k).Copy Destination:=Cells(i, k)
      Next k
      Cells(i, 7).Value = "Planned"
      Cells(i, 8).Value = "no_wo"
      Cells(i, 9).Value = "Brief_preparation_travel_dll"
      If Cells(i, 5).Value = 1 Then
        Cells(i, 10).Value = Int(Cells(i+1, 10).Value) + 0.25
      End If
      Cells(i+1, 10).Copy Destination:=Cells(i, 11)
      i = i + 1
    End If
  Next i
End Sub
