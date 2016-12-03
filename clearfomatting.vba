Sub clearformatting()
'
' clearformatting Macro
'

'
    Cells.Select
    Range("A119").Activate
    Selection.Font.Bold = True
    Selection.Font.Bold = False
    With Selection.Font
        .ColorIndex = xlAutomatic
        .TintAndShade = 0
    End With
    Selection.ClearContents
End Sub