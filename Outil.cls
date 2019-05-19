VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Outil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private PTF As Portefeuille
Private idc1 As IIndicateur
Private idc2 As IIndicateur
Private choix As Decision

Private Sub Class_Initialize()
    Set PTF = New Portefeuille
    Set idc1 = New Indicateur1
    Dim a() As Variant
    a = PTF.getDateAsVariant()
    Call idc1.getArr(a)
    Set idc2 = New Indicateur2
    a = PTF.getValueAsVariant()
    Call idc2.getArr(a)
End Sub

Private Function Eval_idc1(ByVal i As Integer) As Decision
    Eval_idc1 = idc1.Decider(i)
End Function

Private Function Eval_idc2(ByVal i As Integer) As Decision
    Eval_idc2 = idc2.Decider(i)
End Function

Private Function Prendre_Decision(ByVal i As Integer) As Decision
    Dim decision1 As Decision
    Dim decision2 As Decision
    
    decision1 = Eval_idc1(i)
    decision2 = Eval_idc2(i)

    
    'A modifier en fonction de la corrélation des indicateurs !
    If decision1 = acheter And decision2 = acheter Then
        Prendre_Decision = acheter
    ElseIf decision1 = vendre And decision2 = vendre Then
        Prendre_Decision = vendre
    ElseIf decision1 = acheter And decision2 = vendre Then
        Prendre_Decision = rien
    ElseIf decision1 = vendre And decision2 = acheter Then
        Prendre_Decision = rien
    ElseIf decision1 = rien And decision2 = rien Then
        Prendre_Decision = rien
    End If
Fin:
End Function

Public Sub Evaluer(ByVal i As Integer)
    Dim decision_temp As Decision
    decision_temp = Prendre_Decision(i)
    If decision_temp = acheter Then
        Call PTF.acheter(i)
    ElseIf decision_temp = vendre Then
        Call PTF.vendre(i)
    Else
        Call PTF.new_line(i)
    End If
End Sub

Public Function get_VL() As Variant()
    get_VL = PTF.get_VL()
End Function

Public Function get_Date() As Date()
    get_Date = PTF.get_Date()
End Function

Public Function get_Detail()
    get_Detail = PTF.get_arr()
End Function
