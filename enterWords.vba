Option Explicit
Function enterWords(theWords() As String, maxRows As Integer, maxColumns As Integer, Optional minRow As Integer = 1, Optional minColumn As Integer = 1) As Boolean

Dim startColumn As Integer
Dim startRow As Integer
Dim direction As Integer
Dim ghostRow As Integer
Dim ghostColumn As Integer

Dim aCounter As Integer

Randomize

Dim junk As Boolean
Dim missedWords As Object
Set missedWords = CreateObject("System.Collections.ArrayList")
Dim hider As Variant
Dim isgonna As Boolean


For Each hider In theWords
  aCounter = 0
  Do
    startColumn = CInt(Int(((maxColumns - minRow) * Rnd() + minRow)))
    startRow = CInt(Int(((maxRows - minColumn) * Rnd() + minColumn)))
    ghostColumn = startColumn
    ghostRow = startRow
    direction = CInt(Int(8 * Rnd()))
    isgonna = WillFit(CStr(hider), startColumn, maxRows, maxColumns, startRow, direction, minRow, minColumn)
    aCounter = aCounter + 1
  Loop Until isgonna = True Or aCounter = 1000
  
  If isgonna = True Then
    junk = EnterWord(CStr(hider), ghostColumn, ghostRow, direction)
  Else
    missedWords.Add (hider)
  End If
Next hider

If missedWords.Count <> 0 Then
  Dim vari As Integer
  vari = MsgBox("The following words were skipped:" & masterstring(missedWords) & "Would you like to retry?", 4)
  If vari = 6 Then
    Dim myLength As Integer
    myLength = missedWords.Count - 1
    Dim mynewarray() As String
    ReDim mynewarray(0 To myLength)
    Dim myCount As Integer
    myCount = 0
    Dim one As Variant
    For Each one In missedWords
      mynewarray(myCount) = missedWords(myCount)
      myCount = myCount + 1
      Next one
    Dim myjunk As Boolean
    myjunk = enterWords(mynewarray, maxRows, maxColumns, minRow, minColumn)
  End If
    
End If

enterWords = True
 
End Function

