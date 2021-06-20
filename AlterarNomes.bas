Attribute VB_Name = "Módulo1"
Sub alterarNome()

    Dim Codigo      As String
    Dim Cliente     As String
    Dim AbsCod      As String
    Dim Caminho     As String
    Dim NomeAntigo  As String
    Dim NomeNovo    As String
   
    
    Linha = Sheets("Planilha1").Cells(Sheets("Planilha1").Rows.Count, 1).End(xlUp).Row
    AbsCod = Range("E5").Value
    Caminho = Range("E6").Value
    
    On Error Resume Next
    
    Do While Linha >= 2
        Cliente = Range("A" & Linha)
        Codigo = Range("B" & Linha)
        NomeAntigo = Caminho & Cliente & AbsCod & ".pdf"
        NomeNovo = Caminho & Codigo & ".pdf"

        Name NomeAntigo As NomeNovo
    
        Linha = Linha - 1
    Loop
    
    MsgBox "Ufa, acabou!"

End Sub


