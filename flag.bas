Sub flag()
  For i = 3 To Cells(Rows.Count, 1).End(xlUp).Row
    If Cells(i, 4) = "normal" Then
      If Cells(i, 5) = 1 Then
        If Int(Cells(i - 1, 9).Value) <> Int(Cells(i, 9).Value) And (Cells(i, 9).Value - Int(Cells(i, 9).Value)) > 0.25 Then
          Cells(i, 12).Value = "v"
        End If
      ElseIf Cells(i, 5) = 2 Then
        If Int(Cells(i - 1, 9).Value) <> Int(Cells(i, 9).Value) And (Cells(i, 9).Value - Int(Cells(i, 9).Value)) > 0.586806 Then
          Cells(i, 12).Value = "v"
        End If
      ElseIf Cells(i, 5) = 3 Then
        If (Cells(i, 9).Value - Cells(i - 1, 10).Value > 0.5) And (Cells(i, 9).Value - Int(Cells(i, 9).Value)) > 0.920139 Then
          Cells(i, 12).Value = "v"
        End If
      Else
        If Int(Cells(i - 1, 9).Value) <> Int(Cells(i, 9).Value) And (Cells(i, 9).Value - Int(Cells(i, 9).Value)) > 0.315972 Then
          Cells(i, 12).Value = "v"
        End If
      End If
    Else
      If Cells(i, 5) = 1 Then
        If Int(Cells(i - 1, 9).Value) <> Int(Cells(i, 9).Value) And (Cells(i, 9).Value - Int(Cells(i, 9).Value)) > 0.25 Then
          Cells(i, 12).Value = "v"
        End If
      ElseIf Cells(i, 5) = 2 Then
        If Int(Cells(i - 1, 9).Value) <> Int(Cells(i, 9).Value) And (Cells(i, 9).Value - Int(Cells(i, 9).Value)) > 0.5625 Then
          Cells(i, 12).Value = "v"
        End If
      ElseIf Cells(i, 5) = 3 Then
        If Int(Cells(i - 1, 9).Value) <> Int(Cells(i, 9).Value) And (Cells(i, 9).Value - Int(Cells(i, 9).Value)) > 0.875 Then
          Cells(i, 12).Value = "v"
        End If
      ElseIf Cells(i, 5) = 4 Then
        If Int(Cells(i - 1, 9).Value) <> Int(Cells(i, 9).Value) And (Cells(i, 9).Value - Int(Cells(i, 9).Value)) > 0.291667 Then
          Cells(i, 12).Value = "v"
        End If
      ElseIf Cells(i, 5) = 5 Then
        If Int(Cells(i - 1, 9).Value) <> Int(Cells(i, 9).Value) And (Cells(i, 9).Value - Int(Cells(i, 9).Value)) > 0.5 Then
          Cells(i, 12).Value = "v"
        End If
      End If
    End If
  Next i
  For j = 2 To Cells(Rows.Count, 1).End(xlUp).Row
    If Cells(j, 4) = "normal" Then
      If Cells(j, 5) = 1 Then
        If Int(Cells(j + 1, 10).Value) <> Int(Cells(j, 10).Value) And (Cells(j, 10).Value - Int(Cells(j, 10).Value)) < 0.58333 Then
          Cells(j, 12).Value = "vv"
        End If
      ElseIf Cells(j, 5) = 2 Then
        If Int(Cells(j + 1, 10).Value) <> Int(Cells(j, 10).Value) And (Cells(j, 10).Value - Int(Cells(j, 10).Value)) < 0.913194 Then
          Cells(j, 12).Value = "vv"
        End If
      ElseIf Cells(j, 5) = 3 Then
        If (Cells(j + 1, 9).Value - Cells(j, 10).Value > 0.5) And (Cells(j, 10).Value - Int(Cells(j, 10).Value)) < 0.246528 Then
          Cells(j, 12).Value = "vv"
        End If
      Else
        If Int(Cells(j + 1, 10).Value) <> Int(Cells(j, 10).Value) And (Cells(j, 10).Value - Int(Cells(j, 10).Value)) < 0.684028 Then
          Cells(j, 12).Value = "vv"
        End If
      End If
    Else
      If Cells(j, 5) = 1 Then
        If Int(Cells(j + 1, 10).Value) <> Int(Cells(j, 10).Value) And (Cells(j, 10).Value - Int(Cells(j, 10).Value)) < 0.541667 Then
          Cells(j, 12).Value = "vv"
        End If
      ElseIf Cells(j, 5) = 2 Then
        If Int(Cells(j + 1, 10).Value) <> Int(Cells(j, 10).Value) And (Cells(j, 10).Value - Int(Cells(j, 10).Value)) < 0.875 Then
          Cells(j, 12).Value = "vv"
        End If
      ElseIf Cells(j, 5) = 3 Then
        If (Cells(j + 1, 9).Value - Cells(j, 10).Value > 0.5) And (Cells(j, 10).Value - Int(Cells(j, 10).Value)) < 0.1875 Then
          Cells(j, 12).Value = "vv"
        End If
      ElseIf Cells(j, 5) = 4 Then
        If Int(Cells(j + 1, 10).Value) <> Int(Cells(j, 10).Value) And (Cells(j, 10).Value - Int(Cells(j, 10).Value)) < 0.604167 Then
          Cells(j, 12).Value = "vv"
        End If
      ElseIf Cells(j, 5) = 5 Then
        If Int(Cells(j + 1, 10).Value) <> Int(Cells(j, 10).Value) And (Cells(j, 10).Value - Int(Cells(j, 10).Value)) < 0.8125 Then
          Cells(j, 12).Value = "vv"
        End If
      End If
    End If
  Next j
End Sub
