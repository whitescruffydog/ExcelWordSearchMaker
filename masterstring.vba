Function masterstring(arrayed As Object) As String
  masterstring = vbCrLf
  For Each missing In arrayed
    masterstring = masterstring & missing & vbCrLf
  Next missing
End Function
