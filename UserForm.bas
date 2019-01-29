Private Sub btn3_Click()
     findBestWorst (3)
End Sub

Private Sub btnSelect_Click()
    If (Not IsNumeric(txtN)) Then
        MsgBox ("Please, enter the integer number!")
        Exit Sub
    End If
    
    n = Val(txtN.Text)
    findBestWorst (n)
End Sub

Private Sub btnSingle_Click()
MsgBox (getRandomTour & vbNewLine & "Total distance: " & Round(Range("D50").Value, 2))
End Sub

Private Sub btnSort_Click()
Dim Indexes(48)

Range("A2:C49").Select
    ActiveWorkbook.Worksheets("Location").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Location").Sort.SortFields.Add Key:=Range("A2"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Location").Sort
        .SetRange Range("A2:C49")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    
    'fill collection
    For i = 1 To 48
        Indexes(i) = i
    Next i
    
    allDistances
   
    MsgBox ("tour: " & showArray(Indexes) & vbNewLine & "Total distance: " & Round(Range("D50").Value, 2))
    
End Sub

Public Sub findBestWorst(n As Integer)

Dim tour As String, cost As Double, maz As Double, min As Double

Minim = 1111111111
Maxim = 0
For i = 1 To n
    tour = getRandomTour
    's = i & tour & vbNewLine & "Total distance: " & Round(Range("D50").Value, 2)
    cost = Range("D50").Value
    If cost > Maxim Then Maxim = cost:  tourMax = tour
    If cost < Minim Then Minim = cost:  tourMin = tour
Next

MsgBox "The best " & tourMin & vbTab & "Total distance: " & Round(Minim, 2) & vbNewLine & vbNewLine & "The worst " & tourMax & vbTab & "Total distance: " & Round(Maxim, 2), , "N=" & n

End Sub

