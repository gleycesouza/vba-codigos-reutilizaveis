Attribute VB_Name = "geral"
Sub preenc_relatorio()

Dim Word
Dim Doc
Dim Relatorio

tabela_dados = "Gerar Relatório.xlsx"
Windows(tabela_dados).Activate
Range("B1").Select

Do While ActiveCell.Value <> ""
            
    ensaio = ActiveCell.Value
    ActiveCell.Offset(1, 0).Select
    
    relatorio_GMC = ActiveCell.Value
    ActiveCell.Offset(1, 0).Select
    
    relatorio_BR = ActiveCell.Value
    ActiveCell.Offset(1, 0).Select
    
    data_entrega = ActiveCell.Value
    ActiveCell.Offset(1, 0).Select
    
    rev = ActiveCell.Value
    ActiveCell.Offset(1, 0).Select
    
    esp_tec = ActiveCell.Value
    ActiveCell.Offset(1, 0).Select
    
    campo = ActiveCell.Value
    ActiveCell.Offset(1, 0).Select
    
    prioridade = ActiveCell.Value
    ActiveCell.Offset(1, 0).Select
    
    data_amostra = ActiveCell.Value
    ActiveCell.Offset(1, 0).Select
    
Loop

Set Word = CreateObject("Word.Application")
Word.Visible = True

Set Doc = Word.documents.Open("\\server01.geomecanica.com.br\DRIVE_G\Pessoal do Lab\Gleyce\Modelo Tixotropia BR - RL-5714-GT-0XX_00.docx")
'Filename:=ActiveWorkbook.Path & "\" & Desktop & "Dados Brutos " & nome & " - " & furo & "-" & amostra

With Doc
    .Application.Selection.Find.Text = "#NRELATORIOGMC"
    .Application.Selection.Find.Execute
    .Application.Selection.Text = relatorio_GMC
    
    .Application.Selection.Find.Text = "#NRELATORIOBR"
    .Application.Selection.Find.Execute
    .Application.Selection.Text = relatorio_BR
    
    .Application.Selection.Find.Text = "#DATAREL"
    .Application.Selection.Find.Execute
    .Application.Selection.Find.Text = Format(data_entrega, "DD/MM/YYYY")
    
    .Application.Selection.Find.Text = "#REV"
    .Application.Selection.Find.Execute
    .Application.Selection.Text = rev
    
    .Application.Selection.Find.Text = "#NET"
    .Application.Selection.Find.Execute
    .Application.Selection.Text = esp_tec
    
    .Application.Selection.Find.Text = "#CAMPO"
    .Application.Selection.Find.Execute
    .Application.Selection.Text = campo
    
    .Application.Selection.Find.Text = "#PRIORIDADE"
    .Application.Selection.Find.Execute
    .Application.Selection.Text = prioridade
    
    .Application.Selection.Find.Text = "#TDATAAMOSTRAGEM"
    .Application.Selection.Find.Execute
    .Application.Selection.Text = data_amostra
    
    .SaveAs (Filename = "relatorio_GMC")
    
    End With
    
    Set Doc = Nothing
    Set Word = Nothing

MsgBox "Processo finalizado!"

End Sub
Sub inserir_linhas_programacao()

Range("I117").Select

Do While ActiveCell.Value <> ""
    ActiveCell.EntireRow.Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    ActiveCell.Offset(2, 0).Select
    
Loop

MsgBox "Processo finalizado!", , "Inserir linhas em branco"

End Sub

Sub ajustes_Benthic()

Application.ScreenUpdating = False

Dim nome As String
Dim tabela_amostras As String
Dim caminho As String
Dim num_anexo As Integer
Dim profundidade As String

text_anexo = "RL-5718-GT-010_ANX"

    tabela_amostras = "Gerar PDF.xlsm"
    Windows(tabela_amostras).Activate
    ActiveSheet.Range("H3").Select
    
    Do While ActiveCell.Value <> ""
            
            CACO3 = ActiveCell.Value
            ActiveCell.Offset(0, -1).Select
            mat_org = ActiveCell.Value
            ActiveCell.Offset(0, -1).Select
            num_anexo = ActiveCell.Value
            ActiveCell.Offset(0, -1).Select
            nome = ActiveCell.Value
            ActiveCell.Offset(0, -1).Select
            arquivo = ActiveCell.Value
            ActiveCell.Offset(0, -1).Select
            caminho = ActiveCell.Value
            Workbooks.Open caminho, UpdateLinks:=0
            
            Windows(arquivo).Activate
            
            'Colocando local
                'Sheets("Dados amostra").Select
                'Range("B12").Select
                'ActiveCell.Value = profundidade

'            'Colocando umidade e gama
'                Sheets("Dados amostra").Select
'                Range("F33").Select
'                ActiveCell.Value = umidade
'
'                Sheets("Dados amostra").Select
'                Range("F46").Select
'                ActiveCell.Value = peso

            Sheets("LLLP - Apresentação").Select
            Rows("76:76").Select
            Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
            Range("A76:D76").Select
            Selection.Merge
            ActiveCell.FormulaR1C1 = "Teor de Matéria Orgânica*:"
            Range("A76:D76").Select
            With Selection
                .HorizontalAlignment = xlRight
                .VerticalAlignment = xlCenter
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = True
            End With
            
            Range("E75").Select
            ActiveCell.Value = CACO3
            
            Range("E76").Select
            ActiveCell.Value = mat_org
            
            Range("A78:M78").Select
            With Selection
            End With
            Selection.Merge
            ActiveCell.FormulaR1C1 = _
                "(*) Determinações de peso específico natural, teor de umidade e teor de matéria orgânica realizadas pela Benthic"
            Range("A78:M78").Select
            With Selection
                .HorizontalAlignment = xlRight
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = True
                .Font.Size = 8
            End With
            Range("A79").Select
            Selection.ClearContents
            Range("L74:M74").Select
            Selection.NumberFormat = "0.0%"
            
            'Gerando PDF
                Sheets(Array("LLLP - Apresentação", "Curva Granulométrica")).Select
                ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=ActiveWorkbook.Path & "\" & Desktop & num_anexo & " - " & num_anexo + 1 & " " & nome, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
            
            'Salvando e fechando
                ActiveWorkbook.Save
                ActiveWorkbook.Close
                
            'Indo para o próximo registro
                Windows(tabela_amostras).Activate
                ActiveCell.Offset(1, 5).Select
    
    Loop
    
    Application.ScreenUpdating = True
    
    MsgBox "Processo concluído!"
    
End Sub
    

Sub CREG_gráficos()

Application.ScreenUpdating = False

Dim modelo As String
Dim tabela_amostras As String
Dim caminho_ant As String
Dim arquivo_ant As String
Dim nome_ant As String
Dim caminho_atual As String
Dim arquivo_atual As String
Dim nome_atual As String

    tabela_amostras = "Gerar PDF.xlsm"
    Windows(tabela_amostras).Activate
    Sheets("CREG Gráficos").Select
    
    modelo = "CREG Modelo - Todos.xls"
    caminho_modelo = Range("L4").Value & "\" & modelo
    On Error Resume Next
        Workbooks.Open caminho_modelo, UpdateLinks:=0
    
    Windows(tabela_amostras).Activate
    Sheets("CREG Gráficos").Select
    nome_arquivo = Range("J3").Value
    ActiveSheet.Range("C3").Select
    
    Do While ActiveCell.Value <> ""
            
            'Registros Antigos
            caminho_ant = ActiveCell.Value
            ActiveCell.Offset(0, 1).Select
            arquivo_ant = ActiveCell.Value
            ActiveCell.Offset(0, 1).Select
            nome_ant = ActiveCell.Value
            
            'Registros Atuais
            ActiveCell.Offset(0, 3).Select
            caminho_atual = ActiveCell.Value
            ActiveCell.Offset(0, 1).Select
            arquivo_atual = ActiveCell.Value
            ActiveCell.Offset(0, 1).Select
            nome_atual = ActiveCell.Value
            
            If nome_ant <> nome_atual Then
            
                Windows(modelo).Activate
                
                ActiveWorkbook.ChangeLink Name:=caminho_ant, NewName:=caminho_atual, Type:=xlExcelLinks
                
            End If
                        
            Windows(tabela_amostras).Activate
            ActiveCell.Offset(1, -7).Select
    
    Loop
    
    Windows(modelo).Activate
    ActiveWorkbook.SaveAs Filename:=nome_arquivo & " - Todos"
    
    Application.ScreenUpdating = True
    
    MsgBox "Processo finalizado!"
    
End Sub

Sub pdf_adaptavel()

Application.ScreenUpdating = False

Dim nome As String
Dim tabela_amostras As String
Dim caminho As String
Dim num_anexo As Integer

contador = 2

    tabela_amostras = "Gerar PDF.xlsm"
    Windows(tabela_amostras).Activate
    ActiveSheet.Range("E3").Select
    
    Do While ActiveCell.Value <> ""
    
            nome = ActiveCell.Value
            ActiveCell.Offset(0, -1).Select
            arquivo = ActiveCell.Value
            ActiveCell.Offset(0, -1).Select
            caminho = ActiveCell.Value
            Workbooks.Open caminho, UpdateLinks:=0
            
            Windows(arquivo).Activate
            
            Sheets("Dados amostra").Select
            Range("B12").Value = "Campo Farfan"
                               
            ActiveWorkbook.Save
            ActiveWorkbook.Close
                
            Windows(tabela_amostras).Activate
            ActiveCell.Offset(1, 2).Select
    
    Loop
    
    Application.ScreenUpdating = True
    
    MsgBox "O número dos anexos foram alterados com sucesso.", , "Numerador de anexo"
    
End Sub

Sub analisar_adensamento()

    Windows("Adensamento_2''_REG830-20 conferir.xlsx").Activate
    Sheets("Curva de Compressibilidade").Select
    Range("O19:S25").Select
    Selection.Copy
    'mudar registro
    Windows("Adensamento_2''_REG 765-20 falta gs - analisar.xls").Activate
    Sheets("Curva de Compressibilidade").Select
    Range("O19").Select
    ActiveSheet.Paste
    Range("O20").Select
    Windows("Adensamento_2''_REG830-20 conferir.xlsx").Activate
    ActiveSheet.Shapes.Range(Array("Straight Connector 17")).Select
    ActiveSheet.Shapes.Range(Array("Straight Connector 17", _
        "Straight Connector 19")).Select
    ActiveSheet.Shapes.Range(Array("Straight Connector 17", _
        "Straight Connector 19", "Straight Connector 21")).Select
    ActiveSheet.Shapes.Range(Array("Straight Connector 17", _
        "Straight Connector 19", "Straight Connector 21", "Straight Connector 2")). _
        Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("Adensamento_2''_REG 765-20 falta gs - analisar.xls").Activate
    Range("O28").Select
    ActiveSheet.Paste
    ActiveSheet.ChartObjects("Chart 1025").Activate
    ActiveChart.FullSeriesCollection(1).Select
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(5).XValues = _
        "='Curva de Compressibilidade'!$O$20"
    ActiveChart.FullSeriesCollection(5).Values = _
        "='Curva de Compressibilidade'!$O$22"
        
    Sheets("Apresentação").Select
    Range("E26").Select
    Selection.Copy
    Sheets("Curva de Compressibilidade").Select
    Range("O20").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Range("U32").Select
    Selection.ClearContents
    Range("R35").Select
    Selection.ClearContents
    ActiveSheet.ChartObjects("Chart 1027").Activate
    ActiveChart.FullSeriesCollection(1).Select
    ActiveChart.FullSeriesCollection(1).XValues = "=Apresentação!$AC$55:$AC$62"
    ActiveChart.FullSeriesCollection(1).Values = "=Apresentação!$AE$55:$AE$62"
    ActiveChart.FullSeriesCollection(1).Select
    ActiveChart.FullSeriesCollection(1).Smooth = True
    Application.CommandBars("Format Object").Visible = False
    
    ActiveSheet.ChartObjects("Chart 1028").Activate
    ActiveChart.FullSeriesCollection(1).Select
    ActiveChart.FullSeriesCollection(1).XValues = "=Apresentação!$AG$55:$AG$62"
    ActiveChart.FullSeriesCollection(1).Values = "=Apresentação!$AI$55:$AI$62"
      
    
    Sheets("Apresentação").Select
    Range("M35:N43").Select
    ActiveCell.FormulaR1C1 = "=ROUND('Curva de Compressibilidade'!R[-10]C[3],4)"
    Range("O35:P43").Select
    ActiveCell.FormulaR1C1 = "=ROUND('Curva de Compressibilidade'!R[-10]C[2],4)"
    Range("Q35:Q43").Select
    ActiveCell.FormulaR1C1 = "=ROUND('Curva de Compressibilidade'!R[-10]C[1],4)"
    Range("Q44").Select
    
    Windows("Adensamento_2''_REG830-20 conferir.xlsx").Activate
    Sheets("Apresentação").Select
    Range("K23:K24").Select
    Selection.Copy
    
    Windows("Adensamento_2''_REG 765-20 falta gs - analisar.xls").Activate
    Sheets("Apresentação").Select
    Range("K23:K24").Select
    ActiveSheet.Paste
    ActiveCell.Replace What:="[Adensamento_2''''_REG830-20 conferir.xlsx]", _
        Replacement:="", LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:= _
        False, SearchFormat:=False, ReplaceFormat:=False
    

End Sub
Sub benthic()

Application.ScreenUpdating = False

Dim nome As String
Dim umidade As String
Dim peso_esp As String
Dim campo As String
Dim tabela_amostras As String
Dim caminho As String
Dim contador As Integer
    
    text_anexo = "RL-5718-GT-010_ANX"
    
    tabela_amostras = "Gerar PDF.xlsm"
    Windows(tabela_amostras).Activate
    ActiveSheet.Range("F14").Select
    
    
    Do While ActiveCell.Value <> ""
            
'            campo = ActiveCell.Value
'            ActiveCell.Offset(0, -1).Select
'            umidade = ActiveCell.Value
'            ActiveCell.Offset(0, -1).Select
'            peso_espe = ActiveCell.Value
'            ActiveCell.Offset(0, -1).Select
            num_anexo = ActiveCell.Value
            ActiveCell.Offset(0, -1).Select
            nome = ActiveCell.Value
            ActiveCell.Offset(0, -1).Select
            arquivo = ActiveCell.Value
            ActiveCell.Offset(0, -1).Select
            caminho = ActiveCell.Value
            Workbooks.Open caminho, UpdateLinks:=0
            Windows(arquivo).Activate
            
'            Sheets("Dados amostra").Select
'            Range("B12").Select
'            ActiveCell.Value = campo
'            Range("F33").Select
'            ActiveCell.Value = umidade
'            Range("F46").Select
'            ActiveCell.Value = peso_espe
            
            Sheets("LLLP - Apresentação").Select
            Range("M79").Select
            ActiveCell.Value = num_anexo
            Range("H79").Select
            ActiveCell.Value = text_anexo
            
            Windows(arquivo).Activate
            
            'Gerando PDF
            Sheets(Array("LLLP - Apresentação", "Curva Granulométrica")).Select
            ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=ActiveWorkbook.Path & "\" & Desktop & _
            num_anexo & " - " & num_anexo + 1 & " " & nome, _
            Quality:=xlQualityStandard, IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, OpenAfterPublish:=True
            
            'Salvando e fechando
            Windows(arquivo).Activate
            ActiveWorkbook.Save
            ActiveWorkbook.Close
                
            'Indo para o próximo registro
            Windows(tabela_amostras).Activate
            ActiveCell.Offset(1, 3).Select
    
    Loop
    
    Application.ScreenUpdating = True
    
    MsgBox "Arquivos em PDF gerados!", , "Gerador de PDF"
    
End Sub

Sub ajuste_anexo()

Application.ScreenUpdating = False

Dim nome As String
Dim umidade As String
Dim peso_esp As String
Dim campo As String
Dim tabela_amostras As String
Dim caminho As String
Dim contador As Integer
    
    text_anexo = "RL-5714-GT-042_ANX"
    
    tabela_amostras = "Gerar Dados Brutos.xlsm"
    Windows(tabela_amostras).Activate
    ActiveSheet.Range("F3").Select
    
    
    Do While ActiveCell.Value <> ""
            
            num_anexo = ActiveCell.Value
            ActiveCell.Offset(0, -1).Select
            nome = ActiveCell.Value
            ActiveCell.Offset(0, -1).Select
            arquivo = ActiveCell.Value
            ActiveCell.Offset(0, -1).Select
            caminho = ActiveCell.Value
            Workbooks.Open caminho, UpdateLinks:=0
            Windows(arquivo).Activate
            
            Sheets("Apresentação DSS Estático CP1").Select
'            'Sheets("Apresentação").Select
'            Columns(12).Insert
'            Columns("M:M").ColumnWidth = 3.8
'
'            Range("L20:M20").Select
'            With Selection
'                .HorizontalAlignment = xlLeft
'                .VerticalAlignment = xlCenter
'                .WrapText = False
'                .Orientation = 0
'                .AddIndent = False
'                .IndentLevel = 0
'                .ShrinkToFit = False
'                .ReadingOrder = xlContext
'                .MergeCells = True
'            End With
'
'            Range("L21:M21").Select
'            With Selection
'                .HorizontalAlignment = xlLeft
'                .VerticalAlignment = xlCenter
'                .WrapText = False
'                .Orientation = 0
'                .AddIndent = False
'                .IndentLevel = 0
'                .ShrinkToFit = False
'                .ReadingOrder = xlContext
'                .MergeCells = True
'            End With
'
'            Range("L22:M22").Select
'            With Selection
'                .HorizontalAlignment = xlLeft
'                .VerticalAlignment = xlCenter
'                .WrapText = False
'                .Orientation = 0
'                .AddIndent = False
'                .IndentLevel = 0
'                .ShrinkToFit = False
'                .ReadingOrder = xlContext
'                .MergeCells = True
'            End With
'
'            Range("L23:M23").Select
'            With Selection
'                .HorizontalAlignment = xlLeft
'                .VerticalAlignment = xlCenter
'                .WrapText = False
'                .Orientation = 0
'                .AddIndent = False
'                .IndentLevel = 0
'                .ShrinkToFit = False
'                .ReadingOrder = xlContext
'                .MergeCells = True
'            End With
'
'            Range("L24:M24").Select
'            With Selection
'                .HorizontalAlignment = xlLeft
'                .VerticalAlignment = xlCenter
'                .WrapText = False
'                .Orientation = 0
'                .AddIndent = False
'                .IndentLevel = 0
'                .ShrinkToFit = False
'                .ReadingOrder = xlContext
'                .MergeCells = True
'            End With
'
'            Range("L25:M25").Select
'            With Selection
'                .HorizontalAlignment = xlLeft
'                .VerticalAlignment = xlCenter
'                .WrapText = False
'                .Orientation = 0
'                .AddIndent = False
'                .IndentLevel = 0
'                .ShrinkToFit = False
'                .ReadingOrder = xlContext
'                .MergeCells = True
'            End With
'
'            Range("L26:M26").Select
'            With Selection
'                .HorizontalAlignment = xlLeft
'                .VerticalAlignment = xlCenter
'                .WrapText = False
'                .Orientation = 0
'                .AddIndent = False
'                .IndentLevel = 0
'                .ShrinkToFit = False
'                .ReadingOrder = xlContext
'                .MergeCells = True
'            End With
'
'            Range("L27:M27").Select
'            With Selection
'                .HorizontalAlignment = xlLeft
'                .VerticalAlignment = xlCenter
'                .WrapText = False
'                .Orientation = 0
'                .AddIndent = False
'                .IndentLevel = 0
'                .ShrinkToFit = False
'                .ReadingOrder = xlContext
'                .MergeCells = True
'            End With
'
'            ActiveSheet.PageSetup.PrintArea = "$A$1:$M$83"
'            Rows("82:82").RowHeight = 7.5
'            Range("M83").Select
'            ActiveCell.FormulaR1C1 = num_anexo
'
'            Range("M83").Select
'            With Selection
'                .HorizontalAlignment = xlLeft
'                .VerticalAlignment = xlBottom
'                .WrapText = False
'                .Orientation = 0
'                .AddIndent = False
'                .IndentLevel = 0
'                .ShrinkToFit = False
'                .ReadingOrder = xlContext
'                .MergeCells = False
'                .NumberFormat = "000"
'            End With
'
'            Range("L83").Select
'            ActiveCell.FormulaR1C1 = text_anexo
'
'            With Selection
'                .HorizontalAlignment = xlRight
'                .VerticalAlignment = xlBottom
'                .WrapText = False
'                .Orientation = 0
'                .AddIndent = False
'                .IndentLevel = 0
'                .ShrinkToFit = False
'                .ReadingOrder = xlContext
'                .MergeCells = False
'            End With
'
            Application.PrintCommunication = True
            ActiveSheet.PageSetup.PrintArea = "$A$1:$M$85"
            With ActiveSheet.PageSetup
                .LeftHeader = ""
                .CenterHeader = ""
                .RightHeader = ""
                .LeftFooter = ""
                .CenterFooter = ""
                .RightFooter = ""
                .LeftMargin = Application.InchesToPoints(0.590551181102362)
                .RightMargin = Application.InchesToPoints(0.590551181102362)
                .TopMargin = Application.InchesToPoints(0.196850393700787)
                .BottomMargin = Application.InchesToPoints(0.196850393700787)
                .HeaderMargin = Application.InchesToPoints(0.393700787401575)
                .FooterMargin = Application.InchesToPoints(0.393700787401575)
                .PrintHeadings = False
                .PrintGridlines = False
                .PrintComments = xlPrintNoComments
                .CenterHorizontally = True
                .CenterVertically = True
                .Orientation = xlPortrait
                .Draft = False
                .PaperSize = xlPaperA4
                .FirstPageNumber = xlAutomatic
                .Order = xlDownThenOver
                .BlackAndWhite = False
                .Zoom = False
                .PrintErrors = xlPrintErrorsDisplayed
                .OddAndEvenPagesHeaderFooter = False
                .DifferentFirstPageHeaderFooter = False
                .ScaleWithDocHeaderFooter = False
                .AlignMarginsHeaderFooter = True
                .EvenPage.LeftHeader.Text = ""
                .EvenPage.CenterHeader.Text = ""
                .EvenPage.RightHeader.Text = ""
                .EvenPage.LeftFooter.Text = ""
                .EvenPage.CenterFooter.Text = ""
                .EvenPage.RightFooter.Text = ""
                .FirstPage.LeftHeader.Text = ""
                .FirstPage.CenterHeader.Text = ""
                .FirstPage.RightHeader.Text = ""
                .FirstPage.LeftFooter.Text = ""
                .FirstPage.CenterFooter.Text = ""
                .FirstPage.RightFooter.Text = ""
            End With
            Application.PrintCommunication = True
            
            Sheets("Curvas Adensamento CP1").Select
'            Columns(12).Insert
'            Columns("M:M").ColumnWidth = 3.8
'            ActiveSheet.PageSetup.PrintArea = "$A$1:$M$76"
'            Rows("75:75").RowHeight = 7.5
'
'            Range("M76").Select
'            ActiveCell.FormulaR1C1 = "='Apresentação DSS Estático CP1'!R[7]C+1"
'            With Selection
'                .HorizontalAlignment = xlLeft
'                .VerticalAlignment = xlBottom
'                .WrapText = False
'                .Orientation = 0
'                .AddIndent = False
'                .IndentLevel = 0
'                .ShrinkToFit = False
'                .ReadingOrder = xlContext
'                .MergeCells = False
'                .NumberFormat = "000"
'            End With
'
'            Range("L76").Select
'            ActiveCell.FormulaR1C1 = "='Apresentação DSS Estático CP1'!R[7]C"
'            Range("L76").Select
'            With Selection
'                .HorizontalAlignment = xlRight
'                .VerticalAlignment = xlBottom
'                .WrapText = False
'                .Orientation = 0
'                .AddIndent = False
'                .IndentLevel = 0
'                .ShrinkToFit = False
'                .ReadingOrder = xlContext
'                .MergeCells = False
'            End With
'
            Application.PrintCommunication = True
            ActiveSheet.PageSetup.PrintArea = "$A$1:$M$76"
            With ActiveSheet.PageSetup
                .LeftHeader = ""
                .CenterHeader = ""
                .RightHeader = ""
                .LeftFooter = ""
                .CenterFooter = ""
                .RightFooter = ""
                .LeftMargin = Application.InchesToPoints(0.590551181102362)
                .RightMargin = Application.InchesToPoints(0.590551181102362)
                .TopMargin = Application.InchesToPoints(0.196850393700787)
                .BottomMargin = Application.InchesToPoints(0.196850393700787)
                .HeaderMargin = Application.InchesToPoints(0.393700787401575)
                .FooterMargin = Application.InchesToPoints(0.393700787401575)
                .PrintHeadings = False
                .PrintGridlines = False
                .PrintComments = xlPrintNoComments
                .CenterHorizontally = True
                .CenterVertically = True
                .Orientation = xlPortrait
                .Draft = False
                .PaperSize = xlPaperA4
                .FirstPageNumber = xlAutomatic
                .Order = xlDownThenOver
                .BlackAndWhite = False
                .Zoom = False
                .PrintErrors = xlPrintErrorsDisplayed
                .OddAndEvenPagesHeaderFooter = False
                .DifferentFirstPageHeaderFooter = False
                .ScaleWithDocHeaderFooter = False
                .AlignMarginsHeaderFooter = True
                .EvenPage.LeftHeader.Text = ""
                .EvenPage.CenterHeader.Text = ""
                .EvenPage.RightHeader.Text = ""
                .EvenPage.LeftFooter.Text = ""
                .EvenPage.CenterFooter.Text = ""
                .EvenPage.RightFooter.Text = ""
                .FirstPage.LeftHeader.Text = ""
                .FirstPage.CenterHeader.Text = ""
                .FirstPage.RightHeader.Text = ""
                .FirstPage.LeftFooter.Text = ""
                .FirstPage.CenterFooter.Text = ""
                .FirstPage.RightFooter.Text = ""
            End With
            Application.PrintCommunication = True
            
            Sheets(Array("Apresentação DSS Estático CP1", "Curvas Adensamento CP1")).Select
            ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=ActiveWorkbook.Path & "\" & Desktop & num_anexo & " - " & num_anexo + 1 & " - " & nome, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
            
            'Salvando e fechando
            Windows(arquivo).Activate
            ActiveWorkbook.Save
            ActiveWorkbook.Close
                
            'Indo para o próximo registro
            Windows(tabela_amostras).Activate
            ActiveCell.Offset(1, 3).Select
    
    Loop
    
    Application.ScreenUpdating = True
    
    MsgBox "Arquivos em PDF gerados!", , "Gerador de PDF"
    
End Sub

Sub ajuste_anexo_adensamento()

Application.ScreenUpdating = False

Dim nome As String
Dim umidade As String
Dim peso_esp As String
Dim campo As String
Dim tabela_amostras As String
Dim caminho As String
Dim contador As Integer
    
    text_anexo = "RL-5714-GT-068_ANX"
    
    tabela_amostras = "Gerar Dados Brutos.xlsm"
    Windows(tabela_amostras).Activate
    ActiveSheet.Range("F5").Select
    
    
    Do While ActiveCell.Value <> ""
            
            num_anexo = ActiveCell.Value
            ActiveCell.Offset(0, -1).Select
            nome = ActiveCell.Value
            ActiveCell.Offset(0, -1).Select
            arquivo = ActiveCell.Value
            ActiveCell.Offset(0, -1).Select
            caminho = ActiveCell.Value
            Workbooks.Open caminho, UpdateLinks:=0
            Windows(arquivo).Activate
            
            Sheets("Apresentação").Select
            
            ActiveSheet.PageSetup.PrintArea = "$A$1:$l$65"
            Rows("64:64").RowHeight = 7.5
            Columns("l:M").ColumnWidth = 3.8
            Range("l65").Select
            ActiveCell.FormulaR1C1 = num_anexo
            
            Range("l65").Select
            With Selection
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
                .NumberFormat = "000"
            End With
            
            Range("k65").Select
            ActiveCell.FormulaR1C1 = text_anexo
            
            With Selection
                .HorizontalAlignment = xlRight
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            
            Application.PrintCommunication = True
            ActiveSheet.PageSetup.PrintArea = "$A$1:$l$65"
            With ActiveSheet.PageSetup
                .LeftHeader = ""
                .CenterHeader = ""
                .RightHeader = ""
                .LeftFooter = ""
                .CenterFooter = ""
                .RightFooter = ""
                .LeftMargin = Application.InchesToPoints(0.590551181102362)
                .RightMargin = Application.InchesToPoints(0.590551181102362)
                .TopMargin = Application.InchesToPoints(0.196850393700787)
                .BottomMargin = Application.InchesToPoints(0.196850393700787)
                .HeaderMargin = Application.InchesToPoints(0.393700787401575)
                .FooterMargin = Application.InchesToPoints(0.393700787401575)
                .PrintHeadings = False
                .PrintGridlines = False
                .PrintComments = xlPrintNoComments
                .CenterHorizontally = True
                .CenterVertically = True
                .Orientation = xlPortrait
                .Draft = False
                .PaperSize = xlPaperA4
                .FirstPageNumber = xlAutomatic
                .Order = xlDownThenOver
                .BlackAndWhite = False
                .Zoom = False
                .PrintErrors = xlPrintErrorsDisplayed
                .OddAndEvenPagesHeaderFooter = False
                .DifferentFirstPageHeaderFooter = False
                .ScaleWithDocHeaderFooter = False
                .AlignMarginsHeaderFooter = True
                .EvenPage.LeftHeader.Text = ""
                .EvenPage.CenterHeader.Text = ""
                .EvenPage.RightHeader.Text = ""
                .EvenPage.LeftFooter.Text = ""
                .EvenPage.CenterFooter.Text = ""
                .EvenPage.RightFooter.Text = ""
                .FirstPage.LeftHeader.Text = ""
                .FirstPage.CenterHeader.Text = ""
                .FirstPage.RightHeader.Text = ""
                .FirstPage.LeftFooter.Text = ""
                .FirstPage.CenterFooter.Text = ""
                .FirstPage.RightFooter.Text = ""
            End With
            Application.PrintCommunication = True
            
            Sheets("1º Estágio").Select
            Columns("M:M").ColumnWidth = 3.8
            ActiveSheet.PageSetup.PrintArea = "$A$1:$M$74"
            Rows("73:73").RowHeight = 7.5
            
            Range("L74").Select
            ActiveCell.FormulaR1C1 = "=Apresentação!R[-9]C[-1]"
            With Selection
                .HorizontalAlignment = xlRight
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            
            Range("M74").Select
            ActiveCell.FormulaR1C1 = "=Apresentação!R[-9]C[-1]+1"
            
            With Selection
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
                .NumberFormat = "000"
            End With
            
            Application.PrintCommunication = True
            ActiveSheet.PageSetup.PrintArea = "$A$1:$M$74"
            With ActiveSheet.PageSetup
                .LeftHeader = ""
                .CenterHeader = ""
                .RightHeader = ""
                .LeftFooter = ""
                .CenterFooter = ""
                .RightFooter = ""
                .LeftMargin = Application.InchesToPoints(0.590551181102362)
                .RightMargin = Application.InchesToPoints(0.590551181102362)
                .TopMargin = Application.InchesToPoints(0.393700787401575)
                .BottomMargin = Application.InchesToPoints(0.78740157480315)
                .HeaderMargin = Application.InchesToPoints(0)
                .FooterMargin = Application.InchesToPoints(0.511811023622047)
                .PrintHeadings = False
                .PrintGridlines = False
                .PrintComments = xlPrintNoComments
                .CenterHorizontally = True
                .CenterVertically = True
                .Orientation = xlPortrait
                .Draft = False
                .PaperSize = xlPaperA4
                .FirstPageNumber = xlAutomatic
                .Order = xlDownThenOver
                .BlackAndWhite = False
                .Zoom = False
                .FitToPagesWide = 1
                .FitToPagesTall = 1
                .PrintErrors = xlPrintErrorsDisplayed
                .OddAndEvenPagesHeaderFooter = False
                .DifferentFirstPageHeaderFooter = False
                .ScaleWithDocHeaderFooter = False
                .AlignMarginsHeaderFooter = True
                .EvenPage.LeftHeader.Text = ""
                .EvenPage.CenterHeader.Text = ""
                .EvenPage.RightHeader.Text = ""
                .EvenPage.LeftFooter.Text = ""
                .EvenPage.CenterFooter.Text = ""
                .EvenPage.RightFooter.Text = ""
                .FirstPage.LeftHeader.Text = ""
                .FirstPage.CenterHeader.Text = ""
                .FirstPage.RightHeader.Text = ""
                .FirstPage.LeftFooter.Text = ""
                .FirstPage.CenterFooter.Text = ""
                .FirstPage.RightFooter.Text = ""
            End With
                    
            Sheets("2º Estágio").Select
            Columns("M:M").ColumnWidth = 3.8
            ActiveSheet.PageSetup.PrintArea = "$A$1:$M$74"
            Rows("73:73").RowHeight = 7.5
            
            Range("L74").Select
            ActiveCell.FormulaR1C1 = "=Apresentação!R[-9]C[-1]"
            With Selection
                .HorizontalAlignment = xlRight
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            
            Range("M74").Select
            ActiveCell.FormulaR1C1 = "=Apresentação!R[-9]C[-1]+2"
            
            With Selection
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
                .NumberFormat = "000"
            End With
            
            Application.PrintCommunication = True
            ActiveSheet.PageSetup.PrintArea = "$A$1:$M$74"
            With ActiveSheet.PageSetup
                .LeftHeader = ""
                .CenterHeader = ""
                .RightHeader = ""
                .LeftFooter = ""
                .CenterFooter = ""
                .RightFooter = ""
                .LeftMargin = Application.InchesToPoints(0.590551181102362)
                .RightMargin = Application.InchesToPoints(0.590551181102362)
                .TopMargin = Application.InchesToPoints(0.393700787401575)
                .BottomMargin = Application.InchesToPoints(0.78740157480315)
                .HeaderMargin = Application.InchesToPoints(0)
                .FooterMargin = Application.InchesToPoints(0.511811023622047)
                .PrintHeadings = False
                .PrintGridlines = False
                .PrintComments = xlPrintNoComments
                .CenterHorizontally = True
                .CenterVertically = True
                .Orientation = xlPortrait
                .Draft = False
                .PaperSize = xlPaperA4
                .FirstPageNumber = xlAutomatic
                .Order = xlDownThenOver
                .BlackAndWhite = False
                .Zoom = False
                .FitToPagesWide = 1
                .FitToPagesTall = 1
                .PrintErrors = xlPrintErrorsDisplayed
                .OddAndEvenPagesHeaderFooter = False
                .DifferentFirstPageHeaderFooter = False
                .ScaleWithDocHeaderFooter = False
                .AlignMarginsHeaderFooter = True
                .EvenPage.LeftHeader.Text = ""
                .EvenPage.CenterHeader.Text = ""
                .EvenPage.RightHeader.Text = ""
                .EvenPage.LeftFooter.Text = ""
                .EvenPage.CenterFooter.Text = ""
                .EvenPage.RightFooter.Text = ""
                .FirstPage.LeftHeader.Text = ""
                .FirstPage.CenterHeader.Text = ""
                .FirstPage.RightHeader.Text = ""
                .FirstPage.LeftFooter.Text = ""
                .FirstPage.CenterFooter.Text = ""
                .FirstPage.RightFooter.Text = ""
            End With
            
            Sheets("3º Estágio").Select
            Columns("M:M").ColumnWidth = 3.8
            ActiveSheet.PageSetup.PrintArea = "$A$1:$M$74"
            Rows("73:73").RowHeight = 7.5
            
            Range("L74").Select
            ActiveCell.FormulaR1C1 = "=Apresentação!R[-9]C[-1]"
            With Selection
                .HorizontalAlignment = xlRight
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            
            Range("M74").Select
            ActiveCell.FormulaR1C1 = "=Apresentação!R[-9]C[-1]+3"
            
            With Selection
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
                .NumberFormat = "000"
            End With
            
            Application.PrintCommunication = True
            ActiveSheet.PageSetup.PrintArea = "$A$1:$M$74"
            With ActiveSheet.PageSetup
                .LeftHeader = ""
                .CenterHeader = ""
                .RightHeader = ""
                .LeftFooter = ""
                .CenterFooter = ""
                .RightFooter = ""
                .LeftMargin = Application.InchesToPoints(0.590551181102362)
                .RightMargin = Application.InchesToPoints(0.590551181102362)
                .TopMargin = Application.InchesToPoints(0.393700787401575)
                .BottomMargin = Application.InchesToPoints(0.78740157480315)
                .HeaderMargin = Application.InchesToPoints(0)
                .FooterMargin = Application.InchesToPoints(0.511811023622047)
                .PrintHeadings = False
                .PrintGridlines = False
                .PrintComments = xlPrintNoComments
                .CenterHorizontally = True
                .CenterVertically = True
                .Orientation = xlPortrait
                .Draft = False
                .PaperSize = xlPaperA4
                .FirstPageNumber = xlAutomatic
                .Order = xlDownThenOver
                .BlackAndWhite = False
                .Zoom = False
                .FitToPagesWide = 1
                .FitToPagesTall = 1
                .PrintErrors = xlPrintErrorsDisplayed
                .OddAndEvenPagesHeaderFooter = False
                .DifferentFirstPageHeaderFooter = False
                .ScaleWithDocHeaderFooter = False
                .AlignMarginsHeaderFooter = True
                .EvenPage.LeftHeader.Text = ""
                .EvenPage.CenterHeader.Text = ""
                .EvenPage.RightHeader.Text = ""
                .EvenPage.LeftFooter.Text = ""
                .EvenPage.CenterFooter.Text = ""
                .EvenPage.RightFooter.Text = ""
                .FirstPage.LeftHeader.Text = ""
                .FirstPage.CenterHeader.Text = ""
                .FirstPage.RightHeader.Text = ""
                .FirstPage.LeftFooter.Text = ""
                .FirstPage.CenterFooter.Text = ""
                .FirstPage.RightFooter.Text = ""
            End With
            
            Sheets("4º Estágio").Select
            Columns("M:M").ColumnWidth = 3.8
            ActiveSheet.PageSetup.PrintArea = "$A$1:$M$74"
            Rows("73:73").RowHeight = 7.5
            
            Range("L74").Select
            ActiveCell.FormulaR1C1 = "=Apresentação!R[-9]C[-1]"
            With Selection
                .HorizontalAlignment = xlRight
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            
            Range("M74").Select
            ActiveCell.FormulaR1C1 = "=Apresentação!R[-9]C[-1]+4"
            
            With Selection
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
                .NumberFormat = "000"
            End With
            
            Application.PrintCommunication = True
            ActiveSheet.PageSetup.PrintArea = "$A$1:$M$74"
            With ActiveSheet.PageSetup
                .LeftHeader = ""
                .CenterHeader = ""
                .RightHeader = ""
                .LeftFooter = ""
                .CenterFooter = ""
                .RightFooter = ""
                .LeftMargin = Application.InchesToPoints(0.590551181102362)
                .RightMargin = Application.InchesToPoints(0.590551181102362)
                .TopMargin = Application.InchesToPoints(0.393700787401575)
                .BottomMargin = Application.InchesToPoints(0.78740157480315)
                .HeaderMargin = Application.InchesToPoints(0)
                .FooterMargin = Application.InchesToPoints(0.511811023622047)
                .PrintHeadings = False
                .PrintGridlines = False
                .PrintComments = xlPrintNoComments
                .CenterHorizontally = True
                .CenterVertically = True
                .Orientation = xlPortrait
                .Draft = False
                .PaperSize = xlPaperA4
                .FirstPageNumber = xlAutomatic
                .Order = xlDownThenOver
                .BlackAndWhite = False
                .Zoom = False
                .FitToPagesWide = 1
                .FitToPagesTall = 1
                .PrintErrors = xlPrintErrorsDisplayed
                .OddAndEvenPagesHeaderFooter = False
                .DifferentFirstPageHeaderFooter = False
                .ScaleWithDocHeaderFooter = False
                .AlignMarginsHeaderFooter = True
                .EvenPage.LeftHeader.Text = ""
                .EvenPage.CenterHeader.Text = ""
                .EvenPage.RightHeader.Text = ""
                .EvenPage.LeftFooter.Text = ""
                .EvenPage.CenterFooter.Text = ""
                .EvenPage.RightFooter.Text = ""
                .FirstPage.LeftHeader.Text = ""
                .FirstPage.CenterHeader.Text = ""
                .FirstPage.RightHeader.Text = ""
                .FirstPage.LeftFooter.Text = ""
                .FirstPage.CenterFooter.Text = ""
                .FirstPage.RightFooter.Text = ""
            End With
            
            Sheets("5º Estágio").Select
            Columns("M:M").ColumnWidth = 3.8
            ActiveSheet.PageSetup.PrintArea = "$A$1:$M$74"
            Rows("73:73").RowHeight = 7.5
            
            Range("L74").Select
            ActiveCell.FormulaR1C1 = "=Apresentação!R[-9]C[-1]"
            With Selection
                .HorizontalAlignment = xlRight
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            
            Range("M74").Select
            ActiveCell.FormulaR1C1 = "=Apresentação!R[-9]C[-1]+5"
            
            With Selection
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
                .NumberFormat = "000"
            End With
            
            Application.PrintCommunication = True
            ActiveSheet.PageSetup.PrintArea = "$A$1:$M$74"
            With ActiveSheet.PageSetup
                .LeftHeader = ""
                .CenterHeader = ""
                .RightHeader = ""
                .LeftFooter = ""
                .CenterFooter = ""
                .RightFooter = ""
                .LeftMargin = Application.InchesToPoints(0.590551181102362)
                .RightMargin = Application.InchesToPoints(0.590551181102362)
                .TopMargin = Application.InchesToPoints(0.393700787401575)
                .BottomMargin = Application.InchesToPoints(0.78740157480315)
                .HeaderMargin = Application.InchesToPoints(0)
                .FooterMargin = Application.InchesToPoints(0.511811023622047)
                .PrintHeadings = False
                .PrintGridlines = False
                .PrintComments = xlPrintNoComments
                .CenterHorizontally = True
                .CenterVertically = True
                .Orientation = xlPortrait
                .Draft = False
                .PaperSize = xlPaperA4
                .FirstPageNumber = xlAutomatic
                .Order = xlDownThenOver
                .BlackAndWhite = False
                .Zoom = False
                .FitToPagesWide = 1
                .FitToPagesTall = 1
                .PrintErrors = xlPrintErrorsDisplayed
                .OddAndEvenPagesHeaderFooter = False
                .DifferentFirstPageHeaderFooter = False
                .ScaleWithDocHeaderFooter = False
                .AlignMarginsHeaderFooter = True
                .EvenPage.LeftHeader.Text = ""
                .EvenPage.CenterHeader.Text = ""
                .EvenPage.RightHeader.Text = ""
                .EvenPage.LeftFooter.Text = ""
                .EvenPage.CenterFooter.Text = ""
                .EvenPage.RightFooter.Text = ""
                .FirstPage.LeftHeader.Text = ""
                .FirstPage.CenterHeader.Text = ""
                .FirstPage.RightHeader.Text = ""
                .FirstPage.LeftFooter.Text = ""
                .FirstPage.CenterFooter.Text = ""
                .FirstPage.RightFooter.Text = ""
            End With
            
            Sheets("6º Estágio").Select
            Columns("M:M").ColumnWidth = 3.8
            ActiveSheet.PageSetup.PrintArea = "$A$1:$M$74"
            Rows("73:73").RowHeight = 7.5
            
            Range("L74").Select
            ActiveCell.FormulaR1C1 = "=Apresentação!R[-9]C[-1]"
            With Selection
                .HorizontalAlignment = xlRight
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            
            Range("M74").Select
            ActiveCell.FormulaR1C1 = "=Apresentação!R[-9]C[-1]+6"
            
            With Selection
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
                .NumberFormat = "000"
            End With
            
            Application.PrintCommunication = True
            ActiveSheet.PageSetup.PrintArea = "$A$1:$M$74"
            With ActiveSheet.PageSetup
                .LeftHeader = ""
                .CenterHeader = ""
                .RightHeader = ""
                .LeftFooter = ""
                .CenterFooter = ""
                .RightFooter = ""
                .LeftMargin = Application.InchesToPoints(0.590551181102362)
                .RightMargin = Application.InchesToPoints(0.590551181102362)
                .TopMargin = Application.InchesToPoints(0.393700787401575)
                .BottomMargin = Application.InchesToPoints(0.78740157480315)
                .HeaderMargin = Application.InchesToPoints(0)
                .FooterMargin = Application.InchesToPoints(0.511811023622047)
                .PrintHeadings = False
                .PrintGridlines = False
                .PrintComments = xlPrintNoComments
                .CenterHorizontally = True
                .CenterVertically = True
                .Orientation = xlPortrait
                .Draft = False
                .PaperSize = xlPaperA4
                .FirstPageNumber = xlAutomatic
                .Order = xlDownThenOver
                .BlackAndWhite = False
                .Zoom = False
                .FitToPagesWide = 1
                .FitToPagesTall = 1
                .PrintErrors = xlPrintErrorsDisplayed
                .OddAndEvenPagesHeaderFooter = False
                .DifferentFirstPageHeaderFooter = False
                .ScaleWithDocHeaderFooter = False
                .AlignMarginsHeaderFooter = True
                .EvenPage.LeftHeader.Text = ""
                .EvenPage.CenterHeader.Text = ""
                .EvenPage.RightHeader.Text = ""
                .EvenPage.LeftFooter.Text = ""
                .EvenPage.CenterFooter.Text = ""
                .EvenPage.RightFooter.Text = ""
                .FirstPage.LeftHeader.Text = ""
                .FirstPage.CenterHeader.Text = ""
                .FirstPage.RightHeader.Text = ""
                .FirstPage.LeftFooter.Text = ""
                .FirstPage.CenterFooter.Text = ""
                .FirstPage.RightFooter.Text = ""
            End With
                    
            Sheets("7º Estágio").Select
            Columns("M:M").ColumnWidth = 3.8
            ActiveSheet.PageSetup.PrintArea = "$A$1:$M$74"
            Rows("73:73").RowHeight = 7.5
            
            Range("L74").Select
            ActiveCell.FormulaR1C1 = "=Apresentação!R[-9]C[-1]"
            With Selection
                .HorizontalAlignment = xlRight
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            
            Range("M74").Select
            ActiveCell.FormulaR1C1 = "=Apresentação!R[-9]C[-1]+7"
            
            With Selection
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
                .NumberFormat = "000"
            End With
            
            Application.PrintCommunication = True
            ActiveSheet.PageSetup.PrintArea = "$A$1:$M$74"
            With ActiveSheet.PageSetup
                .LeftHeader = ""
                .CenterHeader = ""
                .RightHeader = ""
                .LeftFooter = ""
                .CenterFooter = ""
                .RightFooter = ""
                .LeftMargin = Application.InchesToPoints(0.590551181102362)
                .RightMargin = Application.InchesToPoints(0.590551181102362)
                .TopMargin = Application.InchesToPoints(0.393700787401575)
                .BottomMargin = Application.InchesToPoints(0.78740157480315)
                .HeaderMargin = Application.InchesToPoints(0)
                .FooterMargin = Application.InchesToPoints(0.511811023622047)
                .PrintHeadings = False
                .PrintGridlines = False
                .PrintComments = xlPrintNoComments
                .CenterHorizontally = True
                .CenterVertically = True
                .Orientation = xlPortrait
                .Draft = False
                .PaperSize = xlPaperA4
                .FirstPageNumber = xlAutomatic
                .Order = xlDownThenOver
                .BlackAndWhite = False
                .Zoom = False
                .FitToPagesWide = 1
                .FitToPagesTall = 1
                .PrintErrors = xlPrintErrorsDisplayed
                .OddAndEvenPagesHeaderFooter = False
                .DifferentFirstPageHeaderFooter = False
                .ScaleWithDocHeaderFooter = False
                .AlignMarginsHeaderFooter = True
                .EvenPage.LeftHeader.Text = ""
                .EvenPage.CenterHeader.Text = ""
                .EvenPage.RightHeader.Text = ""
                .EvenPage.LeftFooter.Text = ""
                .EvenPage.CenterFooter.Text = ""
                .EvenPage.RightFooter.Text = ""
                .FirstPage.LeftHeader.Text = ""
                .FirstPage.CenterHeader.Text = ""
                .FirstPage.RightHeader.Text = ""
                .FirstPage.LeftFooter.Text = ""
                .FirstPage.CenterFooter.Text = ""
                .FirstPage.RightFooter.Text = ""
            End With
            
            Sheets("8º Estágio").Select
            Columns("M:M").ColumnWidth = 3.8
            ActiveSheet.PageSetup.PrintArea = "$A$1:$M$74"
            Rows("73:73").RowHeight = 7.5
            
            Range("L74").Select
            ActiveCell.FormulaR1C1 = "=Apresentação!R[-9]C[-1]"
            With Selection
                .HorizontalAlignment = xlRight
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            
            Range("M74").Select
            ActiveCell.FormulaR1C1 = "=Apresentação!R[-9]C[-1]+8"
            
            With Selection
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
                .NumberFormat = "000"
            End With
            
            Application.PrintCommunication = True
            ActiveSheet.PageSetup.PrintArea = "$A$1:$M$74"
            With ActiveSheet.PageSetup
                .LeftHeader = ""
                .CenterHeader = ""
                .RightHeader = ""
                .LeftFooter = ""
                .CenterFooter = ""
                .RightFooter = ""
                .LeftMargin = Application.InchesToPoints(0.590551181102362)
                .RightMargin = Application.InchesToPoints(0.590551181102362)
                .TopMargin = Application.InchesToPoints(0.393700787401575)
                .BottomMargin = Application.InchesToPoints(0.78740157480315)
                .HeaderMargin = Application.InchesToPoints(0)
                .FooterMargin = Application.InchesToPoints(0.511811023622047)
                .PrintHeadings = False
                .PrintGridlines = False
                .PrintComments = xlPrintNoComments
                .CenterHorizontally = True
                .CenterVertically = True
                .Orientation = xlPortrait
                .Draft = False
                .PaperSize = xlPaperA4
                .FirstPageNumber = xlAutomatic
                .Order = xlDownThenOver
                .BlackAndWhite = False
                .Zoom = False
                .FitToPagesWide = 1
                .FitToPagesTall = 1
                .PrintErrors = xlPrintErrorsDisplayed
                .OddAndEvenPagesHeaderFooter = False
                .DifferentFirstPageHeaderFooter = False
                .ScaleWithDocHeaderFooter = False
                .AlignMarginsHeaderFooter = True
                .EvenPage.LeftHeader.Text = ""
                .EvenPage.CenterHeader.Text = ""
                .EvenPage.RightHeader.Text = ""
                .EvenPage.LeftFooter.Text = ""
                .EvenPage.CenterFooter.Text = ""
                .EvenPage.RightFooter.Text = ""
                .FirstPage.LeftHeader.Text = ""
                .FirstPage.CenterHeader.Text = ""
                .FirstPage.RightHeader.Text = ""
                .FirstPage.LeftFooter.Text = ""
                .FirstPage.CenterFooter.Text = ""
                .FirstPage.RightFooter.Text = ""
            End With
            
            Sheets("9º Estágio").Select
            Columns("M:M").ColumnWidth = 3.8
            ActiveSheet.PageSetup.PrintArea = "$A$1:$M$74"
            Rows("73:73").RowHeight = 7.5
            
            Range("L74").Select
            ActiveCell.FormulaR1C1 = "=Apresentação!R[-9]C[-1]"
            With Selection
                .HorizontalAlignment = xlRight
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            
            Range("M74").Select
            ActiveCell.FormulaR1C1 = "=Apresentação!R[-9]C[-1]+9"
            
            With Selection
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
                .NumberFormat = "000"
            End With
            
            Application.PrintCommunication = True
            ActiveSheet.PageSetup.PrintArea = "$A$1:$M$74"
            With ActiveSheet.PageSetup
                .LeftHeader = ""
                .CenterHeader = ""
                .RightHeader = ""
                .LeftFooter = ""
                .CenterFooter = ""
                .RightFooter = ""
                .LeftMargin = Application.InchesToPoints(0.590551181102362)
                .RightMargin = Application.InchesToPoints(0.590551181102362)
                .TopMargin = Application.InchesToPoints(0.393700787401575)
                .BottomMargin = Application.InchesToPoints(0.78740157480315)
                .HeaderMargin = Application.InchesToPoints(0)
                .FooterMargin = Application.InchesToPoints(0.511811023622047)
                .PrintHeadings = False
                .PrintGridlines = False
                .PrintComments = xlPrintNoComments
                .CenterHorizontally = True
                .CenterVertically = True
                .Orientation = xlPortrait
                .Draft = False
                .PaperSize = xlPaperA4
                .FirstPageNumber = xlAutomatic
                .Order = xlDownThenOver
                .BlackAndWhite = False
                .Zoom = False
                .FitToPagesWide = 1
                .FitToPagesTall = 1
                .PrintErrors = xlPrintErrorsDisplayed
                .OddAndEvenPagesHeaderFooter = False
                .DifferentFirstPageHeaderFooter = False
                .ScaleWithDocHeaderFooter = False
                .AlignMarginsHeaderFooter = True
                .EvenPage.LeftHeader.Text = ""
                .EvenPage.CenterHeader.Text = ""
                .EvenPage.RightHeader.Text = ""
                .EvenPage.LeftFooter.Text = ""
                .EvenPage.CenterFooter.Text = ""
                .EvenPage.RightFooter.Text = ""
                .FirstPage.LeftHeader.Text = ""
                .FirstPage.CenterHeader.Text = ""
                .FirstPage.RightHeader.Text = ""
                .FirstPage.LeftFooter.Text = ""
                .FirstPage.CenterFooter.Text = ""
                .FirstPage.RightFooter.Text = ""
            End With
            
            Sheets("10º Estágio ").Select
            Columns("M:M").ColumnWidth = 3.8
            ActiveSheet.PageSetup.PrintArea = "$A$1:$M$74"
            Rows("73:73").RowHeight = 7.5
            
            Range("L74").Select
            ActiveCell.FormulaR1C1 = "=Apresentação!R[-9]C[-1]"
            With Selection
                .HorizontalAlignment = xlRight
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
            End With
            
            Range("M74").Select
            ActiveCell.FormulaR1C1 = "=Apresentação!R[-9]C[-1]+10"
            
            With Selection
                .HorizontalAlignment = xlLeft
                .VerticalAlignment = xlBottom
                .WrapText = False
                .Orientation = 0
                .AddIndent = False
                .IndentLevel = 0
                .ShrinkToFit = False
                .ReadingOrder = xlContext
                .MergeCells = False
                .NumberFormat = "000"
            End With
            
            Application.PrintCommunication = True
            ActiveSheet.PageSetup.PrintArea = "$A$1:$M$74"
            With ActiveSheet.PageSetup
                .LeftHeader = ""
                .CenterHeader = ""
                .RightHeader = ""
                .LeftFooter = ""
                .CenterFooter = ""
                .RightFooter = ""
                .LeftMargin = Application.InchesToPoints(0.590551181102362)
                .RightMargin = Application.InchesToPoints(0.590551181102362)
                .TopMargin = Application.InchesToPoints(0.393700787401575)
                .BottomMargin = Application.InchesToPoints(0.78740157480315)
                .HeaderMargin = Application.InchesToPoints(0)
                .FooterMargin = Application.InchesToPoints(0.511811023622047)
                .PrintHeadings = False
                .PrintGridlines = False
                .PrintComments = xlPrintNoComments
                .CenterHorizontally = True
                .CenterVertically = True
                .Orientation = xlPortrait
                .Draft = False
                .PaperSize = xlPaperA4
                .FirstPageNumber = xlAutomatic
                .Order = xlDownThenOver
                .BlackAndWhite = False
                .Zoom = False
                .FitToPagesWide = 1
                .FitToPagesTall = 1
                .PrintErrors = xlPrintErrorsDisplayed
                .OddAndEvenPagesHeaderFooter = False
                .DifferentFirstPageHeaderFooter = False
                .ScaleWithDocHeaderFooter = False
                .AlignMarginsHeaderFooter = True
                .EvenPage.LeftHeader.Text = ""
                .EvenPage.CenterHeader.Text = ""
                .EvenPage.RightHeader.Text = ""
                .EvenPage.LeftFooter.Text = ""
                .EvenPage.CenterFooter.Text = ""
                .EvenPage.RightFooter.Text = ""
                .FirstPage.LeftHeader.Text = ""
                .FirstPage.CenterHeader.Text = ""
                .FirstPage.RightHeader.Text = ""
                .FirstPage.LeftFooter.Text = ""
                .FirstPage.CenterFooter.Text = ""
                .FirstPage.RightFooter.Text = ""
            End With
            
            Sheets(Array("Apresentação", "1º Estágio", "2º Estágio", "3º Estágio", "4º Estágio", "5º Estágio", "6º Estágio", "7º Estágio", "8º Estágio", "9º Estágio", "10º Estágio ")).Select
            ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=ActiveWorkbook.Path & "\" & Desktop & num_anexo & " - " & num_anexo + 10 & " - " & nome, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
            
            'Salvando e fechando
            Windows(arquivo).Activate
            ActiveWorkbook.Save
            ActiveWorkbook.Close
                
            'Indo para o próximo registro
            Windows(tabela_amostras).Activate
            ActiveCell.Offset(1, 3).Select
    
    Loop
    
    Application.ScreenUpdating = True
    
    MsgBox "Arquivos em PDF gerados!", , "Gerador de PDF"
    
End Sub

Sub pdf_DSS()

Application.ScreenUpdating = False

Dim nome As String
Dim umidade As String
Dim peso_esp As String
Dim campo As String
Dim tabela_amostras As String
Dim caminho As String
Dim contador As Integer
    
    text_anexo = "RL-5714-GT-030_ANX"
    
    tabela_amostras = "Gerar PDF.xlsm"
    Windows(tabela_amostras).Activate
    ActiveSheet.Range("F3").Select
    
    
    Do While ActiveCell.Value <> ""
            
            num_anexo = ActiveCell.Value
            ActiveCell.Offset(0, -1).Select
            nome = ActiveCell.Value
            ActiveCell.Offset(0, -1).Select
            arquivo = ActiveCell.Value
            ActiveCell.Offset(0, -1).Select
            caminho = ActiveCell.Value
            Workbooks.Open caminho, UpdateLinks:=0
            Windows(arquivo).Activate
            
            Sheets("Apresentação DSS Estático CP1").Select

            Range("K27").Select
            ActiveCell.FormulaR1C1 = "Dh, Deslocamento Vertical [cm]:"
            With ActiveCell.Characters(Start:=1, Length:=1).Font
                .Name = "Symbol"
                .FontStyle = "Regular"
                .Size = 9
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = xlUnderlineStyleNone
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .ThemeFont = xlThemeFontNone
            End With
            With ActiveCell.Characters(Start:=2, Length:=30).Font
                .Name = "Arial"
                .FontStyle = "Regular"
                .Size = 9
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = xlUnderlineStyleNone
                .ColorIndex = xlAutomatic
                .TintAndShade = 0
                .ThemeFont = xlThemeFontNone
            End With
            
            Range("D39").Select
            ActiveCell.Value = "Indeformada"
            
            Range("m85").Select
            ActiveCell.Value = num_anexo
            
            Range("l85").Select
            ActiveCell.Value = text_anexo

            Sheets(Array("Apresentação DSS Estático CP1", "Curvas Adensamento CP1")).Select
            ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=ActiveWorkbook.Path _
            & "\" & Desktop & num_anexo & " - " & num_anexo + 1 & " - " & nome, _
            Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, _
            OpenAfterPublish:=True
            
            'Salvando e fechando
            Windows(arquivo).Activate
            ActiveWorkbook.Save
            ActiveWorkbook.Close
                
            'Indo para o próximo registro
            Windows(tabela_amostras).Activate
            ActiveCell.Offset(1, 3).Select
    
    Loop
    
    Application.ScreenUpdating = True
    
    MsgBox "Arquivos em PDF gerados!", , "Gerador de PDF"
    
End Sub
