Function WillFit(word As String, column As Integer, maxRows As Integer, maxColumns As Integer, row As Integer, Optional direction As Integer = 0, Optional minRow As Integer = 1, Optional minColumn As Integer = 1) As Boolean
'Where 0 is up, 1 is diagonal up/right, 2 right, 3 right/down, 4 down, 5 down/left, 6 left, and 7 up/left
  Dim wordLength As Integer
  wordLength = Len(word)
  Dim counter As Integer
  WillFit = True
  
  For counter = 1 To wordLength
    If column > maxColumns Or row > maxRows Or column < minColumn Or row < minRow Then
      WillFit = False
    End If
    If column >= minColumn And row >= minRow Then
      If Cells(row, column) <> "" And Cells(row, column) <> Mid(word, counter, 1) Then
        WillFit = False
      End If
    End If
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
    
    
    
    

End Function

