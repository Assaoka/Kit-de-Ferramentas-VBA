Attribute VB_Name = "Módulo1"
Sub WordConversor_minuscula_MAIUSCULA()
' Define as letras minúsculas e suas correspondentes maiúsculas
    letrasMai = "AÁÂÃBCÇDEÉÊFGHIÍÎJKLMNOÓÔÕPQRSTUÚÛÜVWXYZ"
    letrasMin = "aáâãbcçdeéêfghiíîjklmnoóôõpqrstuúûüvwxyz"
    N = 1
    
' Desligando atualização de tela para aumentar o desempenho
    Application.ScreenUpdating = False
    
' Loop para converter
    Do While lmin <> "z"
        LMAI = Right(Left(letrasMai, N), 1)
        lmin = Right(Left(letrasMin, N), 1)
        
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = lmin
            .Replacement.Text = LMAI
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
                
        N = N + 1
    Loop
    
' Ligando ativação de tela novamente
    Application.ScreenUpdating = True
End Sub

Sub WordConversor_MAIUSCULA_minuscula()
' Define as letras minúsculas e suas correspondentes maiúsculas
    letrasMai = "AÁÂÃBCÇDEÉÊFGHIÍÎJKLMNOÓÔÕPQRSTUÚÛÜVWXYZ"
    letrasMin = "aáâãbcçdeéêfghiíîjklmnoóôõpqrstuúûüvwxyz"
    N = 1
    
' Desligando atualização de tela para aumentar o desempenho
    Application.ScreenUpdating = False
    
' Loop para converter
    Do While lmin <> "z"
        LMAI = Right(Left(letrasMai, N), 1)
        lmin = Right(Left(letrasMin, N), 1)
        
        Selection.Find.ClearFormatting
        Selection.Find.Replacement.ClearFormatting
        With Selection.Find
            .Text = LMAI
            .Replacement.Text = lmin
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .MatchCase = True
            .MatchWholeWord = False
            .MatchWildcards = False
            .MatchSoundsLike = False
            .MatchAllWordForms = False
        End With
        Selection.Find.Execute Replace:=wdReplaceAll
                
        N = N + 1
    Loop
    
' Ligando ativação de tela novamente
    Application.ScreenUpdating = True
End Sub
