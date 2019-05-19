Attribute VB_Name = "Module1"
Sub test()
    i = 1
    j = 2
    l = 3
    
    y = 4
    
    Dim arr(1 To 3, 1 To 2)
    arr(1, 1) = 3
    arr(2, 1) = 2
    arr(3, 1) = 6
    arr(1, 2) = 2
    arr(2, 2) = 1
    arr(3, 2) = 3
    Dim arry(1 To 3)
    arry(1) = 3
    arry(2) = 2
    arry(3) = 5
    
    ts = WorksheetFunction.LinEst(arry, WorksheetFunction.Transpose(arr), False, True)
    prediction = ts(1) * 4 + ts(2)
End Sub

Sub test2()
    Dim arr()
    arr = Range("E2:E13").Value
    Dim arr_x()
    order = 3
    pos = order
    ' RECUPERER LA MATRICE  DES X
    ReDim arr_x(1 To order, 1 To order)
    For i = 1 To order
        For j = 1 To order
            arr_x(i, j) = arr(pos - i - j, 1)
        Next j
    Next i
    ' RECUPERER LE VECTEUR DES Y
    Dim arr_y()
    ReDim arr_y(1 To order)
    c = order
    For j = 1 To order
        arr_y(j) = arr(pos - j, 1)
    Next j
    
    coeff = WorksheetFunction.LinEst(arr_y, WorksheetFunction.Transpose(arr_x), True, True)
End Sub
