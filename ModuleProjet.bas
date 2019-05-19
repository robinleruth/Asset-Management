Attribute VB_Name = "ModuleProjet"
Option Explicit

Sub Projet()
    Dim i As Integer
    Dim last_row As Integer
    last_row = Cells(1000000, 1).End(xlUp).Row
    Call Creer_onglet_VL
    Sheets("VL").Range("A1:D100000").Value = ""
    Dim a As Outil
    Set a = New Outil
    'modifier ça
    For i = 1 To last_row - 1
        Call a.Evaluer(i)
    Next i
    Dim temp() As Variant
    Dim arr_date() As String
    arr_date = copying_date_array_to_string(a.get_Date())
    temp = a.get_VL()
    Sheets("VL").Range("A1:A" & (UBound(temp) - LBound(temp))) = WorksheetFunction.Transpose(arr_date)
    Sheets("VL").Range("B1:B" & (UBound(temp) - LBound(temp))) = temp
    Call Creer_onglet_detail(a, UBound(temp) - LBound(temp))
    Sheets("VL").Activate
    Dim the_form As ReportingForm
    Set the_form = New ReportingForm
    the_form.Show
End Sub

Private Sub Creer_onglet_VL()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Sheets("VL")
    If Err.Number <> 0 Then
        Sheets.Add
        ActiveSheet.Name = "VL"
        Err.Number = 0
    End If
End Sub

Sub Creer_onglet_detail(out As Outil, size As Integer)
    Dim detail_sheet As Worksheet
    On Error Resume Next
    Set detail_sheet = Sheets("Detail")
    If Err.Number <> 0 Then
        Set detail_sheet = Sheets.Add
        detail_sheet.Name = "Detail"
        Err.Number = 0
    End If
    
    detail_sheet.Range("A1") = "Dates"
    detail_sheet.Range("B1") = "Monétaire"
    detail_sheet.Range("C1") = "Actif acheté"
    detail_sheet.Range("D1") = "VL"
    detail_sheet.Range("E1") = "Actif vendu"
    detail_sheet.Range("F1") = "Cours de l'actif"
    
    Dim arr_detail() As Double
    Dim arr_date() As String
    
    arr_detail = out.get_Detail()
    arr_date = copying_date_array_to_string(out.get_Date())
    
    detail_sheet.Range("A2:A" & CStr(size + 1)).Value = WorksheetFunction.Transpose(arr_date)
    detail_sheet.Range("B2:F" & CStr(size + 1)).Value = arr_detail
End Sub

Private Function copying_date_array_to_string(arr() As Date) As String()
    Dim temp() As String
    ReDim temp(LBound(arr) To UBound(arr))
    Dim i As Integer
    
    For i = LBound(arr) To UBound(arr)
        temp(i) = CStr(arr(i)) & "             je triche pour que les dates s'affichent correctement dans le bon format"
    Next i
    copying_date_array_to_string = temp
End Function


