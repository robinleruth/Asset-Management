VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Portefeuille"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private arr() As Double
Private arr_date() As Date
'arr[x][1] = Monétaire à l'instant x
'arr[x][2] = Actif acheté à l'instant x
'arr[x][3] = VL à l'instant x : monétaire + actif acheté - actif vendu 'faut voir ça
'arr[x][4] = Actif vendu à l'instant x
'arr[x][5] = Valeur de l'actif à l'instant x
Private positionAchat As Boolean
Private positionVente As Boolean

Private Sub calculer_VL(ByVal i As Integer)
    Dim a As Double
    Dim v As Double
    
    a = (arr(i, 5) - arr(i, 2)) + arr(i, 2)
    v = (arr(i, 4) - arr(i, 5)) + arr(i, 4)
    
    arr(i, 3) = arr(i, 1)
    If positionAchat = True Then
        arr(i, 3) = arr(i, 3) + a
    End If
    If positionVente = True Then
        arr(i, 3) = arr(i, 3) + v
    End If
End Sub

Public Sub new_line(ByVal i As Integer)
    If i > 1 Then
        arr(i, 1) = arr(i - 1, 1)
        arr(i, 2) = arr(i - 1, 2)
        arr(i, 3) = arr(i - 1, 3)
        arr(i, 4) = arr(i - 1, 4)
    End If
    Call calculer_VL(i)
End Sub

Public Sub acheter(ByVal i As Integer)
    Call new_line(i)
    If positionAchat = False Then
        If positionVente = False Then
            If arr(i, 1) > arr(i, 5) Then
                arr(i, 1) = arr(i, 1) - arr(i, 5)
                arr(i, 2) = arr(i, 5)
                positionAchat = True
            End If
        Else
            arr(i, 1) = arr(i, 1) + arr(i, 4) - arr(i, 5) + arr(i, 4)
            arr(i, 4) = 0
        End If
        positionVente = False
    End If
    Call calculer_VL(i)
End Sub

Public Sub vendre(ByVal i As Integer)
    Call new_line(i)
    If positionVente = False Then
        If positionAchat = False Then
            If arr(i, 1) > arr(i, 5) Then
                arr(i, 1) = arr(i, 1) - arr(i, 5)
                arr(i, 4) = arr(i, 5)
                positionVente = True
            End If
        Else
            arr(i, 1) = arr(i, 1) + arr(i, 5)
            arr(i, 2) = 0
        End If
        positionAchat = False
    End If
    Call calculer_VL(i)
End Sub

Public Function get_VL() As Variant()
    get_VL = Application.WorksheetFunction.Index(arr, 0, 3) 'get the 3rd column from arr
End Function

Public Property Get get_Date()
    get_Date = arr_date
End Property

Public Function getDateAsVariant() As Variant()
    Dim i As Integer
    Dim ret() As Variant
    ReDim ret(LBound(arr_date) To UBound(arr_date))
    For i = LBound(arr_date) To UBound(arr_date)
        ret(i) = CVar(arr_date(i))
    Next i
    getDateAsVariant = ret
End Function

Public Property Get get_arr()
    get_arr = arr
End Property

Public Function getValueAsVariant() As Variant()
    Dim i As Integer
    Dim ret() As Variant
    ReDim ret(LBound(arr) To UBound(arr))
    For i = LBound(arr) To UBound(arr)
        ret(i) = CVar(arr(i, 5))
    Next i
    getValueAsVariant = ret
End Function

Private Sub Class_Initialize()
    Dim CASH_INITIAL As Double
    CASH_INITIAL = 100000
    Dim i As Integer
    
    Dim temp() As Variant
    Dim temp2() As Variant
    Dim actif_worksheet As Worksheet
    Set actif_worksheet = Sheets("Actif")
    
    actif_worksheet.Activate
    temp = actif_worksheet.Range("B2:B" & Cells(1000000, 1).End(xlUp).Row).Value
    temp2 = actif_worksheet.Range("A2:A" & Cells(1000000, 1).End(xlUp).Row).Value
    
    
    ReDim arr(LBound(temp) To UBound(temp), 1 To 5)
    ReDim arr_date(LBound(temp2) To UBound(temp2))
    arr(1, 1) = CASH_INITIAL
    
    For i = LBound(temp) To UBound(temp)
        arr(i, 5) = temp(i, 1)
        arr_date(i) = temp2(i, 1)
    Next i

    
    positionAchat = False
    positionVente = False
End Sub
