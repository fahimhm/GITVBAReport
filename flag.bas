Sub flag()
  For i = 3 To Cells(Rows.Count, 1).End(xlUp).Row
    If Cells(i, 4) = "normal" Then
      If Cells(i, 5) = 1 Then
        'Jika normal, shift 1 dan pekerjaan pertama dimulai diatas jam 6, maka bubuhkan flag'
        If Int(Cells(i - 1, 9).Value) <> Int(Cells(i, 9).Value) And (Cells(i, 9).Value - Int(Cells(i, 9).Value)) > 0.25 Then
          Cells(i, 12).Value = "v"
        End If
      ElseIf Cells(i, 5) = 2 Then
        'Jika normal, shift 2 dan pekerjaan pertama dimulai diatas jam 14:00, maka bubuhkan flag'
        If Int(Cells(i - 1, 9).Value) <> Int(Cells(i, 9).Value) And (Cells(i, 9).Value - Int(Cells(i, 9).Value)) > 0.586806 Then
          Cells(i, 12).Value = "v"
        End If
      ElseIf Cells(i, 5) = 3 Then
        'Jika normal, shift 3 dan pekerjaan pertama dimulai diatas jam 22:00, maka bubuhkan flag'
        If (Cells(i, 9).Value - Cells(i - 1, 10).Value > 0.5) And (Cells(i, 9).Value - Int(Cells(i, 9).Value)) > 0.920139 Then
          Cells(i, 12).Value = "v"
        End If
      Else
        'Jika normal, non-shift dan pekerjaan pertama dimulai diatas jam 7:30, maka bubuhkan flag'
        If Int(Cells(i - 1, 9).Value) <> Int(Cells(i, 9).Value) And (Cells(i, 9).Value - Int(Cells(i, 9).Value)) > 0.315972 Then
          Cells(i, 12).Value = "v"
        End If
      End If
    Else 'untuk kasus overtime'
      If Cells(i, 5) = 1 Then
        'Jika overtime di shift 1 (full) dan pekerjaan pertama dimulai diatas jam 6:00, maka bubuhkan flag'
        If Int(Cells(i - 1, 9).Value) <> Int(Cells(i, 9).Value) And (Cells(i, 9).Value - Int(Cells(i, 9).Value)) > 0.25 Then
          Cells(i, 12).Value = "v"
        End If
      ElseIf Cells(i, 5) = 2 Then
        'Jika overtime di shift 2 (full) dan pekerjaan pertama dimulai diatas jam 13:30, maka bubuhkan flag'
        If Int(Cells(i - 1, 9).Value) <> Int(Cells(i, 9).Value) And (Cells(i, 9).Value - Int(Cells(i, 9).Value)) > 0.5625 Then
          Cells(i, 12).Value = "v"
        End If
      ElseIf Cells(i, 5) = 3 Then
        'Jika overtime di shift 3 (full) dan pekerjaan pertama dimulai diatas jam 21:00, maka bubuhkan flag'
        If Int(Cells(i - 1, 9).Value) <> Int(Cells(i, 9).Value) And (Cells(i, 9).Value - Int(Cells(i, 9).Value)) > 0.875 Then
          Cells(i, 12).Value = "v"
        End If
      ElseIf Cells(i, 5) = 4 Then
        'Jika overtime di shift 1 banci (full) dan pekerjaan pertama dimulai diatas jam 7:00, maka bubuhkan flag'
        If Int(Cells(i - 1, 9).Value) <> Int(Cells(i, 9).Value) And (Cells(i, 9).Value - Int(Cells(i, 9).Value)) > 0.291667 Then
          Cells(i, 12).Value = "v"
        End If
      ElseIf Cells(i, 5) = 5 Then
        'Jika overtime di shift 2 banci (full) dan pekerjaan pertama dimulai diatas jam 12::, maka bubuhkan flag'
        If Int(Cells(i - 1, 9).Value) <> Int(Cells(i, 9).Value) And (Cells(i, 9).Value - Int(Cells(i, 9).Value)) > 0.5 Then
          Cells(i, 12).Value = "v"
        End If
      End If
    End If
  Next i
  For j = 2 To Cells(Rows.Count, 1).End(xlUp).Row
    If Cells(j, 4) = "normal" Then
      If Cells(j, 5) = 1 Then
        'Jika normal shift 1 dan pekerjaan terakhir selesai sebelum jam 14:30, maka bubuhkan flag'
        If Int(Cells(j + 1, 10).Value) <> Int(Cells(j, 10).Value) And (Cells(j, 10).Value - Int(Cells(j, 10).Value)) < 0.600694444 Then
          Cells(j, 12).Value = "vv"
        End If
      ElseIf Cells(j, 5) = 2 Then
        'Jika normal shift 2 dan pekerjaan terakhir selesai sebelum jam 22:30, maka bubuhkan flag'
        If Int(Cells(j + 1, 10).Value) <> Int(Cells(j, 10).Value) And (Cells(j, 10).Value - Int(Cells(j, 10).Value)) < 0.9340277778 Then
          Cells(j, 12).Value = "vv"
        End If
      ElseIf Cells(j, 5) = 3 Then
        'Jika normal shift 3 dan pekerjaan terakhir selesai sebelum jam 6:30, maka bubuhkan flag'
        If (Cells(j + 1, 9).Value - Cells(j, 10).Value > 0.5) And (Cells(j, 10).Value - Int(Cells(j, 10).Value)) < 0.26736111 Then
          Cells(j, 12).Value = "vv"
        End If
      Else
        'Jika normal non-shift dan pekerjaan terakhir selesai sebelum jam 16:30, maka bubuhkan flag'
        If Int(Cells(j + 1, 10).Value) <> Int(Cells(j, 10).Value) And (Cells(j, 10).Value - Int(Cells(j, 10).Value)) < 0.684028 Then
          Cells(j, 12).Value = "vv"
        End If
      End If
    Else 'untuk kasus overtime'
      If Cells(j, 5) = 1 Then
        'Jika overtime di shift 1 (full) dan pekerjaan terakhir selesai sebelum jam 13:00, maka bubuhkan flag'
        If Int(Cells(j + 1, 10).Value) <> Int(Cells(j, 10).Value) And (Cells(j, 10).Value - Int(Cells(j, 10).Value)) < 0.53819444 Then
          Cells(j, 12).Value = "vv"
        End If
      ElseIf Cells(j, 5) = 2 Then
        'Jika overtime di shift 2 (full) dan pekerjaan terakhir selesai sebelum jam 21:00, maka bubuhkan flag'
        If Int(Cells(j + 1, 10).Value) <> Int(Cells(j, 10).Value) And (Cells(j, 10).Value - Int(Cells(j, 10).Value)) < 0.87152778 Then
          Cells(j, 12).Value = "vv"
        End If
      ElseIf Cells(j, 5) = 3 Then
        'Jika overtime di shift 3 (full) dan pekerjaan terakhir selesai sebelum jam 4:30, maka bubuhkan flag'
        If (Cells(j + 1, 9).Value - Cells(j, 10).Value > 0.5) And (Cells(j, 10).Value - Int(Cells(j, 10).Value)) < 0.18402778 Then
          Cells(j, 12).Value = "vv"
        End If
      ElseIf Cells(j, 5) = 4 Then
        'Jika overtime di shift 1 banci (full) dan pekerjaan terakhir selesai sebelum jam 14:30, maka bubuhkan flag'
        If Int(Cells(j + 1, 10).Value) <> Int(Cells(j, 10).Value) And (Cells(j, 10).Value - Int(Cells(j, 10).Value)) < 0.60069444 Then
          Cells(j, 12).Value = "vv"
        End If
      ElseIf Cells(j, 5) = 5 Then
        'Jika overtime di shift 2 banci (full) dan pekerjaan terakhir selesai sebelum jam 19:30, maka bubuhkan flag'
        If Int(Cells(j + 1, 10).Value) <> Int(Cells(j, 10).Value) And (Cells(j, 10).Value - Int(Cells(j, 10).Value)) < 0.80902778 Then
          Cells(j, 12).Value = "vv"
        End If
      End If
    End If
  Next j
End Sub
