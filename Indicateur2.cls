VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Indicateur2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IIndicateur

Private choix As Decision
Private arr_valeur() As Double

Private Sub Class_Initialize()
    Randomize
    choix = rien
End Sub

Private Function IIndicateur_Decider(ByVal i As Integer) As Decision
    Call IIndicateur_Evaluer(i)
    IIndicateur_Decider = choix
End Function

Private Sub IIndicateur_Evaluer(ByVal i As Integer)
    'Remplir avec la stratégie de l'indicateur
    Dim order As Integer
    order = 3
    Dim prediction As Double
    Dim previous_values() As Double
    ReDim previous_values(1 To order)
    Dim j As Integer
    Dim l As Integer
    Dim coeff() As Variant
    Dim arr_x() As Double
    Dim arr_y() As Double
    
    If i <= order * 2 Then
        choix = rien
        GoTo Fin
    End If
    
     ' RECUPERER LA MATRICE  DES X
    ReDim arr_x(1 To order, 1 To order)
    For j = 1 To order
        For l = 1 To order
            arr_x(j, l) = arr_valeur(i - j - l)
        Next l
    Next j
    ' RECUPERER LE VECTEUR DES Y
    ReDim arr_y(1 To order)
    For j = 1 To order
        arr_y(j) = arr_valeur(i - j)
    Next j
    
    coeff = WorksheetFunction.LinEst(arr_y, WorksheetFunction.Transpose(arr_x), True, True)
    
    prediction = 0
    For j = 1 To order
        prediction = prediction + coeff(1, j) * arr_valeur(i + 1 - j)
    Next j
    prediction = prediction + coeff(1, order + 1)
    
    If prediction > arr_valeur(i) Then
        choix = acheter
    'ElseIf prediction - arr_valeur(i) < 0.5 Then
     '   choix = rien
    Else
        choix = vendre
    End If
    
    'If prediction - arr_valeur(i) < 0.1 Then
     '   choix = rien
     'End If
Fin:
End Sub


Private Sub IIndicateur_getArr(arr() As Variant)
    Dim i As Integer
    ReDim arr_valeur(LBound(arr) To UBound(arr))
    For i = LBound(arr) To UBound(arr)
        arr_valeur(i) = CDbl(arr(i))
    Next i
End Sub
