VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Indicateur1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IIndicateur

Private choix As Decision
Private arr_date() As Date

Private Sub Class_Initialize()
    Randomize
    choix = rien
End Sub

Private Function IIndicateur_Decider(ByVal i As Integer) As Decision
    Call IIndicateur_Evaluer(i)
    IIndicateur_Decider = choix
End Function


Private Sub IIndicateur_Evaluer(ByVal i As Integer)
    
    Dim mois As Integer
    
    mois = CInt(Mid(arr_date(i), 4, 2))
    
    If mois = 11 Then
        choix = acheter
    ElseIf mois = 4 Then
        choix = vendre
    Else
        choix = rien
    End If
       
        
End Sub



Private Sub IIndicateur_getArr(arr() As Variant)
    Dim i As Integer
    ReDim arr_date(LBound(arr) To UBound(arr))
    For i = LBound(arr) To UBound(arr)
        arr_date(i) = CDate(arr(i))
    Next i
End Sub
