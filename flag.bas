Sub flag()
  For i = 3 To Cells(Rows.Count, 1).End(xlUp).Row
    If Cells(i, 5) = 1 Then
      If Int(Cells(i - 1, 10).Value) <> Int(Cells(i, 10).Value) And (Cells(i, 10).Value - Int(Cells(i, 10).Value)) > 0.25 Then
        Cells(i, 13).Value = "v"
      End If
    ElseIf Cells(i, 5) = 2 Then
      If Int(Cells(i - 1, 10).Value) <> Int(Cells(i, 10).Value) And (Cells(i, 10).Value - Int(Cells(i, 10).Value)) > 0.586806 Then
        Cells(i, 13).Value = "v"
      End If
    ElseIf Cells(i, 5) = 3 Then
      If (Cells(i, 10).Value - Cells(i - 1, 11).Value > 0.5) And (Cells(i, 10).Value - Int(Cells(i, 10).Value)) > 0.920139 Then
        Cells(i, 13).Value = "v"
      End If
    Else
      If Int(Cells(i - 1, 10).Value) <> Int(Cells(i, 10).Value) And (Cells(i, 10).Value - Int(Cells(i, 10).Value)) > 0.315972 Then
        Cells(i, 13).Value = "v"
      End If
    End If
  Next i
  For j = 2 To Cells(Rows.Count, 1).End(xlUp).Row
    If Cells(j, 5) = 1 Then
      If Int(Cells(j + 1, 11).Value) <> Int(Cells(j, 11).Value) And (Cells(j, 11).Value - Int(Cells(j, 11).Value)) < 0.58333 Then
        Cells(j, 13).Value = "vv"
      End If
    ElseIf Cells(j, 5) = 2 Then
      If Int(Cells(j + 1, 11).Value) <> Int(Cells(j, 11).Value) And (Cells(j, 11).Value - Int(Cells(j, 11).Value)) < 0.913194 Then
        Cells(j, 13).Value = "vv"
      End If
    ElseIf Cells(j, 5) = 3 Then
      If (Cells(j + 1, 10).Value - Cells(j, 11).Value > 0.5) And (Cells(j, 11).Value - Int(Cells(j, 11).Value)) < 0.246528 Then
        Cells(j, 13).Value = "vv"
      End If
    Else
      If Int(Cells(j + 1, 11).Value) <> Int(Cells(j, 11).Value) And (Cells(j, 11).Value - Int(Cells(j, 11).Value)) < 0.684028 Then
        Cells(j, 13).Value = "vv"
      End If
    End If
  Next j
End Sub
