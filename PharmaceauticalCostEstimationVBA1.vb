Sub mjknjknj()

Dim i, count, k1, k2, x As Integer
i = 2
count = 0
k1 = 2
k2 = 3
lngLastRow = Range("A" & Rows.count).End(xlUp).Row
Do While i < lngLastRow
    If Cells(k1, 2).Value = "Med A" And Cells(k2, 2).Value = "Med B" And Cells(k1, 1).Value = Cells(k2, 1).Value And Month(Cells(k2, 3)) = 11 Then
    count = count + 1
    End If
    i = i + 1
    k1 = k1 + 1
    k2 = k2 + 1
    
Loop
Cells(2, 9).Value = count
End Sub

Sub kjl()
Dim i, count, k1, k2, x As Integer

i = 2
count = 0
k1 = 2
k2 = 3
lngLastRow = Range("A" & Rows.count).End(xlUp).Row
Do While i < lngLastRow
    If Cells(k1, 2).Value = "Med A" And Cells(k2, 2).Value = "Med B" And Cells(k1, 1).Value = Cells(k2, 1).Value And Month(Cells(k2, 3)) = 10 Then
    x = x + 1
    End If
    i = i + 1
    k1 = k1 + 1
    k2 = k2 + 1
    
Loop
Cells(2, 8).Value = x
End Sub

Sub trk()
Dim i, count, k1, k2, s As Integer

i = 2
count = 0
k1 = 2
k2 = 3
lngLastRow = Range("A" & Rows.count).End(xlUp).Row
Do While i < lngLastRow
    If Cells(k1, 2).Value = "Med A" And Cells(k2, 2).Value = "Med B" And Cells(k1, 1).Value = Cells(k2, 1).Value And Month(Cells(k2, 3)) = 9 Then
    s = s + 1
    End If
    i = i + 1
    k1 = k1 + 1
    k2 = k2 + 1
    
Loop
Cells(2, 7).Value = s
End Sub
Sub gty()

Dim i, m, k1, k2 As Integer

i = 2
m = 0
k1 = 2
k2 = 3
lngLastRow = Range("A" & Rows.count).End(xlUp).Row
Do While i < lngLastRow
    If Cells(k1, 2).Value = "Med B" And Cells(k2, 2).Value = "Med A" And Cells(k1, 1).Value = Cells(k2, 1).Value And Month(Cells(k2, 3)) = 11 Then
    m = m + 1
    End If
    i = i + 1
    k1 = k1 + 1
    k2 = k2 + 1
    
Loop
Cells(3, 9).Value = m
End Sub

Sub sawr()

Dim i, z, k1, k2, x As Integer

i = 2
z = 0
k1 = 2
k2 = 3
lngLastRow = Range("A" & Rows.count).End(xlUp).Row
Do While i < lngLastRow
    If Cells(k1, 2).Value = "Med B" And Cells(k2, 2).Value = "Med A" And Cells(k1, 1).Value = Cells(k2, 1).Value And Month(Cells(k2, 3)) = 10 Then
    z = z + 1
    End If
    i = i + 1
    k1 = k1 + 1
    k2 = k2 + 1
    
Loop
Cells(3, 8).Value = z
End Sub
Sub skwr()

Dim i, y, k1, k2 As Integer

i = 2
y = 0
k1 = 2
k2 = 3
lngLastRow = Range("A" & Rows.count).End(xlUp).Row
Do While i < lngLastRow
    If Cells(k1, 2).Value = "Med B" And Cells(k2, 2).Value = "Med A" And Cells(k1, 1).Value = Cells(k2, 1).Value And Month(Cells(k2, 3)) = 9 Then
    y = y + 1
    End If
    i = i + 1
    k1 = k1 + 1
    k2 = k2 + 1
    
Loop
Cells(3, 7).Value = y
End Sub


