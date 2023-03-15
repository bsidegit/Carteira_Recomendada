 'Autor: Eduardo Scheffer
 'Contato:(42)99950-5555
 'Last Edited: 30/11/2022
 
Sub copiar_asset_allocation()
    
    Application.Calculation = xlCalculationManual
    
    Call lock_sheets
    
    Dim modelo As Worksheet
    Dim rebalanceamento As Worksheet
    Set modelo = Worksheets("Portfólios modelo") 'A
    Set rebalanceamento = Worksheets("Rebalanceamento") 'R
    
    Dim row_A1 As Integer
    Dim row_A2 As Integer
    Dim row_A3 As Integer
    Dim row_prod_A As Integer
    Dim row_R1 As Integer
    Dim row_R2 As Integer
    Dim row_prod_R As Integer
    Dim n_prod As Integer
    Dim row_end As Integer
    
    Dim col_A1 As Integer
    Dim col_A2 As Integer
    Dim col_R1 As Integer
    Dim col_R2 As Integer
    
    Dim row_perfil As Integer
    Dim col_perfil As Integer
    Dim col_prod_A As Integer
    Dim col_prod_R As Integer
    Dim col_proxy_R As Integer
    Dim col_perc_R As Integer
    Dim perfil As String
     
    Application.Calculation = xlCalculationManual
    
    rebalanceamento.Unprotect Password:="Alohomora"
        
    perfil = rebalanceamento.Range("N3")
    col_class_R = 5
    col_x_R = col_class_R - 1
    col_perc_R = 19
    col_class_A = 5
    col_x_A = col_class_A - 1
   
    row_A = 5
    row_R = 12
    
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
    
    
End Sub
