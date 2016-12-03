Function EnterWord(word As String, column As Integer, row As Integer, Optional direction As Integer = 0) As Boolean
  Dim wordLength As Integer
  wordLength = Len(word)
  Dim counter As Integer
  
  For counter = 1 To wordLength
    Cells(row, column) = Mid(word, counter, 1)
    Cells(row, column).Select
    Selection.Font.Bold = True
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    
    If direction = 1 Then
      row = row - 1
      column = column + 1
    ElseIf direction = 2 Then
      column = column + 1
    ElseIf direction = 3 Then
      row = row + 1
      column = column + 1
    ElseIf direction = 4 Then
      row = row + 1
    ElseIf direction = 5 Then
      row = row + 1
      column = column - 1
    ElseIf direction = 6 Then
      column = column - 1
    ElseIf direction = 7 Then
      row = row - 1
      column = column - 1
    Else
      row = row - 1
    End If
  Next counter
  
EnterWord = True

End Function
