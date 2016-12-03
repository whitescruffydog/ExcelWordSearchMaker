Option Explicit

Sub myMain()


Dim maxC As Integer
Dim maxR As Integer
Dim minR As Integer
Dim minC As Integer
maxC = 10  'The last column you want populated
maxR = 10  'The last row you want populated
minC = 3   'The first column you want populated
minR = 3   'The first row you want populated

'The words you want to hide
'As I mentioned, this isn't a pretty algorithm, so you'll have to capitalize them if you want it to work right

Dim theWords(0 To 6) As String
theWords(0) = "ONE"
theWords(1) = "TWO"
theWords(2) = "THREE"
theWords(3) = "FOUR"
theWords(4) = "FIVE"
theWords(5) = "SIX"



 
Dim myjunk As Boolean
 
myjunk = enterWords(theWords, maxR, maxC, minR, minC)
myjunk = randomizeBlanks(maxR, maxC, minR, minC, 3)



End Sub

