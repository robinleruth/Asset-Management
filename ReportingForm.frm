VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ReportingForm 
   Caption         =   "Reporting"
   ClientHeight    =   5655
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10020
   OleObjectBlob   =   "ReportingForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ReportingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private report As Reporting

Private Sub dateDeb_tbox_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If dateDeb_tbox.Value = "a" Then
        Cancel = True
    End If
End Sub

Private Sub UserForm_Initialize()
    Set report = New Reporting
    dateDeb_tbox.Value = "05/01/2000"
    dateFin_tbox.Value = "12/12/2017"
End Sub

Private Sub SaveChart()
    Dim MyChart As Chart
    Dim Fname As String

    Set MyChart = report.getGraphe
    Fname = ThisWorkbook.Path & "\temp1.gif"
    MyChart.Export Filename:=Fname, FilterName:="GIF"
End Sub


Private Sub Enter_button_Click()
    Dim datedebut As String, dateFin As String
    datedebut = CStr(Me.dateDeb_tbox.Value)
    dateFin = CStr(Me.dateFin_tbox.Value)
    
    If verifDate(datedebut) And verifDate(dateFin) Then
        If CDate(datedebut) < CDate(dateFin) Then
            If out_of_bound(datedebut) = True Then
                MsgBox ("La date de dÈbut est hors cadre")
                GoTo Fin
            End If
            If out_of_bound(dateFin) = True Then
                MsgBox ("La date de fin est hors cadre")
                GoTo Fin
            End If
            Call report.setDateDeb(datedebut)
            Call report.setDateFin(dateFin)
            Call report.setArrPerf
            Dim arr_base_100() As Double
            arr_base_100 = report.base_100()
            'perf_annu_box.Value = "Besoin des dates dÈbuts et dates fin"
            perf_annu_box.Value = Format(report.perf_annuelle(), "0.0000")
            perf_periode_box.Value = Format(report.getPerfPeriode, "0.0000")
            vol_box.Value = Format(report.vol(), "0.0000")
            max_dd_box.Value = Format(report.max_drawdown(), "0.0000")
            var_box.Value = Format(report.var_parametrique(), "0.0000")
            ratio_sharpe_box.Value = Format(report.ratio_sharpe(), "0.0000")
            
            report.getGraphe.ChartType = xlLine
            report.getGraphe.Axes(xlValue).MinimumScale = WorksheetFunction.min(arr_base_100) - 1
            report.getGraphe.Axes(xlValue).MaximumScale = WorksheetFunction.max(arr_base_100) + 1
            'report.getGraphe.FullSeriesCollection(1).Name = "test"
            'report.getGraphe.ApplyLayout (5)
            
            Call SaveChart
            Fname = ThisWorkbook.Path & "\temp1.gif"
            Me.Frame1.Picture = LoadPicture(Fname)
            
            report.getGraphe.Parent.Delete
        Else
            MsgBox ("La date de fin est avant la date de dÈbut")
        End If
    Else
        MsgBox ("Une des dates n'est pas valide")
    End If
Fin:
End Sub

Private Function verifDate1(myDate As String)
    Dim bool As Boolean
    bool = True
    
    If Not IsDate(myDate) Then
    End If
    
    
    If myDate <> Format("dd/mm/yyyy") Then
        MsgBox ("Format de date invalide")
        bool = False
        GoTo Fin:
    End If
    
    Dim d As Integer
    d = CInt(Left(myDate, 2))
    
    If d > 31 Then
        MsgBox ("Erreur sur le jour")
        bool = False
        GoTo Fin:
    End If
    
    d = CInt(Mid(myDate, 4, 2))
    
    If d > 12 Then
        MsgBox ("Erreur sur le mois")
        bool = False
        GoTo Fin:
    End If
    
    d = CInt(Right(myDate, 4))
    
    If d < 0 And d > 3000 Then
        MsgBox ("Erreur sur l'année")
        bool = False
        GoTo Fin:
    End If
    
Fin:
    verifDate = bool

End Function

Function verifDate(ByVal myDate As String)
    Dim bool As Boolean
    bool = True
    
    If Not IsDate(myDate) Then
        MsgBox ("Format de date invalide")
        bool = False
        GoTo Fin:
    End If
    
    Dim d As Integer
    d = CInt(Left(myDate, 2))
    
    If d > 31 Then
        MsgBox ("Erreur sur le jour")
        bool = False
        GoTo Fin:
    End If
    
    d = CInt(Mid(myDate, 4, 2))
    
    If d > 12 Then
        MsgBox ("Erreur sur le mois")
        bool = False
        GoTo Fin:
    End If
    
    d = CInt(Right(myDate, 4))
    
    If d < 0 And d > 3000 Then
        MsgBox ("Erreur sur l'année")
        bool = False
        GoTo Fin:
    End If
    
Fin:
    verifDate = bool
End Function

Private Function out_of_bound(my_date As String) As Boolean
    out_of_bound = True
    Dim temp() As Variant
    Dim first_cell As Range
    Set first_cell = Range("A1")
    temp = Range(first_cell, first_cell.End(xlDown))
    Dim i As Integer
    For i = LBound(temp) To UBound(temp)
        If temp(i, 1) = my_date & "             je triche pour que les dates s'affichent correctement dans le bon format" Then
            out_of_bound = False
            GoTo Fin
        End If
    Next i
Fin:
End Function

