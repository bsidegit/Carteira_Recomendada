'Autor: Eduardo Scheffer
'Contato:(42)99950-5555
'Last Edited: 20/12/2022
'-------------------------------------------

Function rebalancear_AA() As Boolean  ' Copies sheets into new PDF file for e-mailing

    Dim modelo As Worksheet
    Set modelo = Worksheets("Portfólios modelo") 'A

    Dim row_A1 As Integer
    Dim row_A2 As Integer
    Dim row_A3 As Integer
    Dim n_prod As Integer
    Dim row_end As Integer
    
    Dim col_A1 As Integer
    Dim col_A2 As Integer
    
    Dim row_perfil As Integer
    Dim col_perfil As Integer
    Dim col_prod_A As Integer
    Dim perfil As String
     
    Application.Calculation = xlCalculationManual
    
    rebalanceamento.Unprotect Password:="Alohomora"
        
    perfil = rebalanceamento.Range("N3")
    col_class_A = 5
    col_x_A = col_class_A - 1
   
    row_A = 5
    
    'Find perfil column
    col_perfil = col_class_A + 7
    row_perfil = row_A
    Do While modelo.Cells(row_perfil, col_perfil).Value <> perfil And col_perfil <= modelo.Cells(row_perfil, Columns.Count).End(xlToLeft).Column
        col_perfil = col_perfil + 2
    Loop
    col_perfil = col_perfil + 1
    
    'Copy strategies of model portfolio
    row_A = modelo.Cells(row_A, col_class_A).End(xlDown).End(xlDown).End(xlDown).Row
    Do While modelo.Cells(row_A, col_x_A - 2).Value <> "xx" Or rebalanceamento.Cells(row_R, col_x_R - 2).Value <> "xx"
        If modelo.Cells(row_A, col_x_A).Value = "x" Then
            If rebalanceamento.Cells(row_R, col_class_R).Value = modelo.Cells(row_A, col_class_A).Value Then
                rebalanceamento.Cells(row_R, col_perc_R).Value = modelo.Cells(row_A, col_perfil).Value
                row_R = rebalanceamento.Cells(row_R, col_x_R).End(xlDown).Row
                row_A = modelo.Cells(row_A, col_x_A).End(xlDown).Row
            Else
                row_R = rebalanceamento.Cells(row_R, col_x_R).End(xlDown).Row
            End If
        Else
            row_A = modelo.Cells(row_A, col_x_A).End(xlDown).Row
        End If
    Loop
     
    rebalanceamento.Range("E9").Select
    Application.CutCopyMode = False
    
    If Application.UserName <> "Eduardo Scheffer" And Application.UserName <> "Marcelo Carrara" Then rebalanceamento.Protect Password:="Alohomora"
    Application.Calculation = xlCalculationAutomatic
    
    VBA.Interaction.MsgBox "Asset Allocation Tático importado.", , "Concluido"

End Function

Function rebalancear_copyPL() As Boolean  ' Copies sheets into new PDF file for e-mailing



End Function

Function rebalancear_copyPercent() As Boolean  ' Copies sheets into new PDF file for e-mailing



End Function


Sub rebalancear()
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    Call lock_sheets
    
    Dim posicao As Worksheet
    Dim rebalanceamento As Worksheet
    Set rebalanceamento = Worksheets("Rebalanceamento") 'R
    
    Dim col_asset_R As Integer
    Dim col_x_R As Integer
    Dim row_R As Integer
    Dim row_Ri As Integer
    Dim row_Rf As Integer
    Dim row_RR As Integer
    Dim col_PL1_R As Integer
    Dim col_perc2_R As Integer
    Dim total As Variant
    Dim percent_PL As Variant

    Application.Calculation = xlCalculationManual

    Call limpa_rebalanceamento
    Call limpa_rebalanceamento
    
    rebalanceamento.Unprotect Password:="Alohomora"
        
    client = rebalanceamento.Range("G3").Value
    col_asset_R = 5
    col_x_R = col_asset_R - 1
    col_PL1_R = col_asset_R + 9
    col_perc2_R = col_PL1_R + 5
    col_vehicle_R = col_PL1_R - 2
    row_Ri = 10
    row_Rf = rebalanceamento.Cells(Rows.Count, 5).End(xlUp).Row

    col_client_P = 4
    col_strategy_P = 6
    col_PL_P = 12
    col_caixaPL_P = 23
    col_asset_P = 31
    row_P = 5
    row_Pf = posicao.Cells(Rows.Count, 1).End(xlUp).Row
    row_P = 2
    
    'Find client row
    Do While row_P <= row_Pf
        If (posicao.Cells(row_P, col_client_P).Value = client And row_Pi = 0) Then
            row_Pi = row_P
        ElseIf (posicao.Cells(row_P, col_client_P).Value <> client And row_Pi > 0) Then
            row_Pf = row_P - 1
        End If
        row_P = row_P + 1
    Loop
    
    'Copy formulas
    classificacao.Range("F24").Copy
    posicao.Range(posicao.Cells(row_Pi, col_asset_P), posicao.Cells(row_Pf, col_asset_P)).PasteSpecial xlPasteFormulas
    classificacao.Range("F25").Copy
    posicao.Range(posicao.Cells(row_Pi, col_strategy_P), posicao.Cells(row_Pf, col_strategy_P)).PasteSpecial xlPasteFormulas

    'Calculate total
    row_P = row_Pi
    total = 0
    Do While row_P <= row_Pf
        total = total + posicao.Cells(row_P, col_PL_P).Value + posicao.Cells(row_P, col_caixaPL_P).Value
        row_P = row_P + 1
    Loop
    rebalanceamento.Range("O3").Value = total


    'Copy assets
    row_P = row_Pi
    Do While row_P <= row_Pf
        asset = posicao.Cells(row_P, col_asset_P).Value
        strategy = posicao.Cells(row_P, col_strategy_P).Value

        'Look for strategy row in rebalanceamento
        row_R = row_Ri
        Do While row_R < row_Rf
            row_R = rebalanceamento.Cells(row_R, col_x_R).End(xlDown).Row
            If rebalanceamento.Cells(row_R, col_x_R).Value = "x" And rebalanceamento.Cells(row_R, col_x_R - 1).Value <> "y" And rebalanceamento.Cells(row_R, col_asset_R).Value = strategy Then
                row_RR = rebalanceamento.Cells(row_R, col_x_R).End(xlDown).Row
                Exit Do
            End If
        Loop

        'Copy asset
        If rebalanceamento.Cells(row_R + 1, col_asset_R) <> "" Then
            row_R = rebalanceamento.Cells(row_R, col_asset_R).End(xlDown).Row
            rebalanceamento.Range(rebalanceamento.Cells(row_R, 1), rebalanceamento.Cells(row_R, 1)).EntireRow.Copy
            rebalanceamento.Range(rebalanceamento.Cells(row_R + 1, 1), rebalanceamento.Cells(row_R + 1, 1)).EntireRow.Insert
            row_Rf = row_Rf + 1
            rebalanceamento.Cells(row_R + 1, col_asset_R).ClearContents
            rebalanceamento.Cells(row_R + 1, col_PL1_R).ClearContents
            rebalanceamento.Cells(row_R + 1, col_perc2_R).ClearContents
            rebalanceamento.Cells(row_R + 1, col_vehicle_R).ClearContents
        End If
            
        row_R = row_R + 1
        posicao.Cells(row_P, col_asset_P).Copy
        rebalanceamento.Cells(row_R, col_asset_R).PasteSpecial xlPasteValues
        PL_asset = posicao.Cells(row_P, col_PL_P).Value + posicao.Cells(row_P, col_caixaPL_P).Value
        percent_PL = PL_asset / total
        rebalanceamento.Cells(row_R, col_PL1_R).Value = PL_asset
        rebalanceamento.Cells(row_R, col_perc2_R).Value = percent_PL

        row_P = row_P + 1
    Loop
    
    row_R = rebalanceamento.Cells(row_Ri, col_x_R).End(xlDown).Row
    
    
    
    Do While rebalanceamento.Cells(row_R, col_x_R).Value <> "xx"
        If rebalanceamento.Cells(row_R, col_x_R).Value = "x" Then
            rebalanceamento.Cells(row_R, col_perc2_R).Value = rebalanceamento.Cells(row_R, col_PL1_R + 1).Value
        End If
        row_R = rebalanceamento.Cells(row_R, col_x_R).End(xlDown).Row
    Loop
    
    
    rebalanceamento.Range("I4").Value = posicao.Range("A2").Value
    rebalanceamento.Range("E9").Select
    Application.CutCopyMode = False
    
    If Application.UserName <> "Eduardo Scheffer" And Application.UserName <> "Marcelo Carrara" Then rebalanceamento.Protect Password:="Alohomora"
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    VBA.Interaction.MsgBox "Posição do cliente importada com sucesso. Confirme o perfil.", , "Concluido"
    
    
End Sub