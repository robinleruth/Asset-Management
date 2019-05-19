VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Reporting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Private arr_VL() As Double
Private arr_performance() As Double
Private graphe As ChartObject
Private perf_periode As Double
Private perf_annu As Double
Private volatilite As Double

Private date_debut As String
Private date_fin As String

Private nb_jour As Integer

Private Sub setNbJour()
    nb_jour = CDate(date_fin) - CDate(date_debut)
End Sub

Public Sub setDateDeb(dateDeb As String)
    date_debut = dateDeb
End Sub

Public Sub setDateFin(dateFin As String)
    date_fin = dateFin
End Sub

Public Property Get getPerfPeriode()
    getPerfPeriode = perf_periode
End Property

Public Property Get getGraphe()
    Set getGraphe = graphe.Chart
End Property

Public Function vol() As Double
    vol = WorksheetFunction.StDev_S(arr_performance) * 252 ^ 0.5
    volatilite = vol
End Function

Public Function max_drawdown() As Double
    max_drawdown = 0
    Dim max As Double
    Dim min As Double
    Dim i As Integer
    Dim j As Integer
    Dim temp As Variant
    For i = LBound(arr_VL) To UBound(arr_VL)
        For j = i + 1 To UBound(arr_VL)
            temp = (arr_VL(j) - arr_VL(i)) / arr_VL(i) ' Calcul de performance
            If temp < max_drawdown Then
                max_drawdown = temp 'Récupération du plus grand drawdown
                max = arr_VL(i)
                min = arr_VL(j)
            End If
        Next j
    Next i
End Function

Public Function var_parametrique() As Double
    var_parametrique = volatilite / (252 ^ 0.5) * 2.33
End Function

Public Function var_historique() As Double
    Dim min As Double, max As Double, pas As Double, i As Integer, somme As Double
    min = WorksheetFunction.min(arr_performance)
    max = WorksheetFunction.max(arr_performance)

    Dim interv(1 To 21) As Double
    pas = (max - min) / 20
    interv(LBound(interv)) = min
    For i = LBound(interv) + 1 To UBound(interv)
        interv(i) = interv(i - 1) + pas
    Next i
    Dim freq() As Variant
    freq = WorksheetFunction.Frequency(arr_performance, interv)
    somme = WorksheetFunction.Sum(freq)
    For i = LBound(freq) To UBound(freq)
        freq(i, 1) = freq(i, 1) / somme
    Next i
    For i = LBound(freq) + 1 To UBound(freq)
        freq(i, 1) = freq(i - 1, 1) + freq(i, 1)
    Next i
    For i = LBound(freq) + 1 To UBound(freq)
        If freq(i, 1) > 0.1 Then
            var_historique = interv(i - 1) - (interv(i) - interv(i - 1)) * (0.1 - freq(i - 1, 1)) / (freq(i, 1) - freq(i - 1, 1))
            GoTo Fin
        End If
    Next i
Fin:
End Function

Public Function ratio_sharpe() As Double
    ratio_sharpe = (perf_annu - 0) / volatilite 'Remplacer 0 par l'EONIA
End Function

Public Function perf_annuelle() As Double
    Dim perf_quotidienne As Double
    
    perf_quotidienne = (1 + perf_periode) ^ (1 / nb_jour) - 1
    perf_annuelle = (1 + perf_quotidienne) ^ 365 - 1
    perf_annu = perf_annuelle
End Function

Private Sub setGraphe(arr_base_100)
    Dim temp As Chart
    Range("C1:C" & UBound(arr_base_100)).Value = WorksheetFunction.Transpose(arr_base_100)
    Set graphe = ActiveSheet.ChartObjects.Add(1, 1, 250, 200)
    With graphe.Chart
        .SetSourceData Source:=Range("C1:C" & UBound(arr_base_100))
    End With
End Sub

Public Function base_100() As Double()
    Dim perf() As Double
    Dim arr_base_100() As Double
    perf = performance(arr_VL) ' Calculer toutes les perfs
    arr_base_100 = b_100(perf) ' Obtenir la base 100
    Call setGraphe(arr_base_100)
    perf_periode = arr_base_100(UBound(arr_base_100)) - arr_base_100(LBound(arr_base_100))
    perf_periode = perf_periode / 100
    base_100 = arr_base_100
    
End Function

Private Function performance(arr() As Double) As Double()
    Dim ret() As Double
    ReDim ret(LBound(arr) To UBound(arr))
    
    Dim i As Integer
    
    For i = LBound(arr) + 1 To UBound(arr)
        ret(i) = 1 + ((arr(i) - arr(i - 1)) / arr(i - 1)) ' Calcul de (1 + perf)
    Next i
    
    performance = ret
End Function

Private Function performance_reel(arr() As Double) As Double()
    Dim ret() As Double
    ReDim ret(LBound(arr) To UBound(arr))
    
    Dim i As Integer
    
    For i = LBound(arr) + 1 To UBound(arr)
        ret(i) = ((arr(i) - arr(i - 1)) / arr(i - 1)) ' Calcul de (perf)
    Next i
    
    performance_reel = ret
End Function

Private Function b_100(arr() As Double) As Double()
    Dim i As Integer
    Dim ret() As Double
    ReDim ret(LBound(arr) To UBound(arr))
    
    ret(1) = 100 'base 100 donc premier nombre est 100
    For i = LBound(arr) + 1 To UBound(arr)
        ret(i) = ret(1) * recur_product(arr, i) '= 100 * produit(1er_nbre:Actuel_nbre)
    Next i
    
    b_100 = ret
End Function

Private Function recur_product(arr() As Double, i As Integer) As Double
    If i <= 1 Then
        recur_product = 1
        GoTo Fin
    End If
    recur_product = arr(i) * recur_product(arr, i - 1)
Fin:
End Function

'Private Sub Class_Initialize()
    'Modifier en prenant compte des dates !!!!!!!!!!!!!!
    'Dim first_cell As Range
    'Set first_cell = Range("B1")
    'Dim temp() As Variant
    'Dim i As Integer
    
    '-----------------------
    'Dim le_test() As Double
    'le_test = get_correct_array(date_debut, date_fin)
    '-----------------------
    
    'temp = Range(first_cell, first_cell.End(xlDown)) ' Récupérer les valeurs brut
    'ReDim arr_VL(LBound(temp) To UBound(temp))
    'For i = LBound(temp) To UBound(temp)
        'arr_VL(i) = temp(i, 1)
    'Next i
    'arr_performance = performance_reel(arr_VL)
'End Sub

Public Sub setArrPerf()
    'Modifier en prenant compte des dates !!!!!!!!!!!!!!
    Dim first_cell As Range
    Set first_cell = Range("B1")
    Dim temp() As Variant
    Dim i As Integer
    
    arr_VL = get_correct_array(date_debut, date_fin)
    Call setNbJour
    
    arr_performance = performance_reel(arr_VL)
End Sub

Private Function get_correct_array(ByVal dateDeb As String, ByVal dateFin As String) As Double()
    Dim temp() As Variant
    Dim first_cell As Range
    Set first_cell = Range("A1")
    temp = Range(first_cell, first_cell.End(xlDown)) 'on récupère toutes les dates
    
    Dim i As Integer
    Dim i_deb As Integer
    Dim i_fin As Integer
    For i = LBound(temp) To UBound(temp)
        If temp(i, 1) = dateDeb & "             je triche pour que les dates s'affichent correctement dans le bon format" Then
            i_deb = i
            GoTo Apres
        End If
    Next i
Apres:
    For i = i_deb To UBound(temp)
        If temp(i, 1) = dateFin & "             je triche pour que les dates s'affichent correctement dans le bon format" Then
            i_fin = i
        End If
    Next i
    
    Set first_cell = Range("B1")
    temp = Range(first_cell, first_cell.End(xlDown))
    
    Dim toRet() As Double
    ReDim toRet(1 To i_fin - i_deb)
    For i = LBound(toRet) To UBound(toRet)
        toRet(i) = temp(i + i_deb - 1, 1)
    Next i
    
    get_correct_array = toRet
End Function
