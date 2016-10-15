Sub kos()

Dim i, count, k1, k2, x As Integer

i = 2
count = 0
k1 = 2
lngLastRow = Range("A" & Rows.count).End(xlUp).Row
Do While i <= lngLastRow
    If Cells(k1, 3).Value = Cells(k1, 4).Value Then
    count = count + 1
    End If
    i = i + 1
    k1 = k1 + 1
    
    
Loop
Cells(2, 9).Value = count
End Sub

Sub lkm()

Dim i, count, k1, k2, x As Integer

i = 2
count = 0
k1 = 2
lngLastRow = Range("A" & Rows.count).End(xlUp).Row
Do While i <= lngLastRow
    If Cells(k1, 3).Value > Cells(k1, 4).Value And Cells(k1, 4).Value <> 0 Then
    count = count + 1
    End If
    i = i + 1
    k1 = k1 + 1
    
    
Loop
Cells(2, 10).Value = count
End Sub
Sub zq()

Dim i, count, k1, k2, x As Integer

i = 2
count = 0
k1 = 2
lngLastRow = Range("A" & Rows.count).End(xlUp).Row
Do While i <= lngLastRow
    If Cells(k1, 3).Value < Cells(k1, 4).Value Then
    count = count + 1
    End If
    i = i + 1
    k1 = k1 + 1
    
    
Loop
Cells(2, 11).Value = count
End Sub
Sub ojk()

Dim i, count, k1, k2, x As Integer

i = 2
count = 0
k1 = 2
lngLastRow = Range("A" & Rows.count).End(xlUp).Row
Do While i <= lngLastRow
    If Cells(k1, 3).Value > Cells(k1, 4).Value And Cells(k1, 4).Value = 0 Then
    count = count + 1
    End If
    i = i + 1
    k1 = k1 + 1
    
    
Loop
Cells(2, 12).Value = count
End Sub

