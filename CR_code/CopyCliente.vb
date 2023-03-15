'Autor: Eduardo Scheffer
'Contato:(42)99950-5555
'Last Edited: 12/13/2022
'-------------------------------------------

Sub posicao_cliente()
    
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False
    
    Call lock_sheets
    
    Dim posicao As Worksheet
    Dim rebalanceamento As Worksheet
    Set posicao = Worksheets("Posição AUM") 'P
    Set rebalanceamento = Worksheets("Rebalanceamento") 'R
    Set classificacao = Worksheets("Classificação") 'R
    
    Dim col_asset_R As Integer
    Dim col_x_R As Integer
    Dim row_R As Integer
    Dim row_Ri As Integer
    Dim row_Rf As Integer
    Dim row_RR As Integer
    Dim col_client_P As Integer
    Dim col_strategy_P As Integer
    Dim col_PL_P As Integer
    Dim col_PL1_R As Integer
    Dim col_perc2_R As Integer
    Dim col_caixaPL_P As Integer
    Dim col_asset_P As Integer
    Dim row_P As Integer
    Dim row_Pi As Integer
    Dim row_Pf As Integer
    Dim total As Variant
    Dim percent_PL As Variant
    Dim client As String
    Dim asset As String
    Dim strategy As String

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
    rebalanceamento.Range("O2").Value = total


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

