Attribute VB_Name = "Excel_SorteioGrupos"
Sub Sorteio_Grupos()
' Definindo e calculando variáveis necessárias para a execução do sorteio
    Linha_Nomes = Cells(1, 1).End(xlDown).Row
    Linha_Grupos = Cells(1, 2).End(xlDown).Row
    N_Participantes = (Linha_Nomes \ Linha_Grupos)
    IntervaloTabela = "$A$1:$B$" & Linha_Nomes
    NomePlanilha = ActiveSheet.Name
    
    Letras = "AÁÂÃBCÇDEÉÊFGHIÍÎJKLMNOÓÔÕPQRSTUÚÛÜVWXYZ"
    
    ' Verificando se o número de participantes é suficiente para formar todos os grupos e encerrando a macro em caso negativo
    If N_Participantes < 1 Then
        MsgBox "A quantidade de grupos excede a quantidade de participantes disponíveis."
        Exit Sub
    End If

    ' Randomizando o valor da variável "Numeros" para garantir resultados diferentes a cada execução
    Randomize
        Numeros = Int(1937554306 * Rnd(1937554306 + 1937554306 * Rnd())) + 1937554306
        Numeros = (Numeros + Numeros + 82098) * (7868787 + Numeros)
    
' Copiando Nomes e Grupos
    Range(Cells(1, 1), Cells(Linha_Nomes, 2)).Copy
    Range("B1").Select
    ActiveSheet.Paste
    
' Loop Reordenação
    Do While Entrada <> "Z"
        Entrada = Right(Left(Letras, N), 1)
        Saida = Right(Left(Numeros, N), 2)
        Range(Cells(1, 1), Cells(Linha_Nomes, 1)).Replace What:=Entrada, Replacement:=Saida, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
        N = N + 1
    Loop
    
' Ordenando as correspondencias com base na redordenação
    Range(Cells(1, 3), Cells(Linha_Grupos, 3)).Cut
    Range("E1").Select
    ActiveSheet.Paste
    
    ActiveSheet.ListObjects.Add(xlSrcRange, Range(IntervaloTabela), , xlNo).Name = "Tabela777777"
    Range("Tabela777777[#All]").Select
        ActiveWorkbook.Worksheets(NomePlanilha).ListObjects("Tabela777777").Sort.SortFields.Clear
        ActiveWorkbook.Worksheets(NomePlanilha).ListObjects("Tabela777777").Sort.SortFields.Add2 Key:=Range("Tabela777777[[#All],[Coluna1]]"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With ActiveWorkbook.Worksheets(NomePlanilha).ListObjects("Tabela777777").Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With

' Preenchendo a tabela final com os nomes dos participantes e seus respectivos grupos
    N = 2
    Do While N <> (Linha_Grupos + 1)
        Cells(N, 5).Cut
        Cells(1, (N + 4)).Select
        ActiveSheet.Paste
        N = N + 1
    Loop
    
    Nome = "A"
    N = 2
    Linha = 2
    Coluna = 5
    Do While Nome <> ""
        Nome = Cells(N, 2).Value
        Cells(Linha, Coluna).Value = Nome
        N = N + 1
        
        Coluna = Coluna + 1
        If Coluna > (Linha_Grupos + 4) Then
            Coluna = 5
            Linha = Linha + 1
        End If
    Loop
    
' Finalizando
    Columns("A:D").Delete Shift:=xlToLeft
    
    Linha = Cells(1, 1).End(xlDown).Row
    Coluna = Cells(1, 1).End(xlToRight).Column
    
    Application.CutCopyMode = False
    
    Set Tabela = Range(Cells(1, 1), Cells(Linha, Coluna))
    ActiveSheet.ListObjects.Add(xlSrcRange, Range(Cells(1, 1), Cells(Linha, Coluna)), , xlYes).Name = "Tabela777777"
    Range("Tabela777777[#All]").Select
    ActiveSheet.ListObjects("Tabela777777").TableStyle = "TableStyleMedium14"
    
    Range(Cells(1, 1), Cells(Linha, Coluna)).EntireColumn.AutoFit
    Sheets(NomePlanilha).Name = "SorteioGrupos"

    Apagar_Linhas = (Cells(1, 1).End(xlDown).Row + 1) & ":" & (Cells(1, 1).End(xlDown).Row + 10)
    Rows(Apagar_Linhas).Delete Shift:=xlUp
End Sub
