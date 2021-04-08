Attribute VB_Name = "anexo_pdf"
Sub pdf_adaptavel()

Application.ScreenUpdating = False

Dim nome As String
Dim tabela_amostras As String
Dim caminho As String
Dim num_anexo As Integer


    text_anexo = "RL-5714-GT-076_ANX00"
    tabela_amostras = "Gerar Dados Brutos.xlsm"
    Windows(tabela_amostras).Activate
    ActiveSheet.Range("E3").Select
    num_anexo = 2
    
    Do While ActiveCell.Value <> ""
            
            
            nome = ActiveCell.Value
            ActiveCell.Offset(0, -1).Select
            arquivo = ActiveCell.Value
            ActiveCell.Offset(0, -1).Select
            caminho = ActiveCell.Value
            Workbooks.Open caminho, UpdateLinks:=0
            
            Windows(arquivo).Activate
                      
            Sheets("Dados").Select
            Range("J14").Value = "Leandro Vieira"
            Range("B13").Value = "Campo Barra/Farfan/Muri� (SE/AL)"
            furo = Range("g12").Value
            amostra = Range("g13").Value
            
'            Sheets("Sensibilidade - Apresenta��o").Select
'            Application.PrintCommunication = False
'            With ActiveSheet.PageSetup
'                .PrintTitleRows = ""
'                .PrintTitleColumns = ""
'            End With
'            Application.PrintCommunication = True
'            ActiveSheet.PageSetup.PrintArea = "$A$1:$K$74"
'            Application.PrintCommunication = False
'            With ActiveSheet.PageSetup
'                .LeftHeader = ""
'                .CenterHeader = ""
'                .RightHeader = ""
'                .LeftFooter = " "
'                .CenterFooter = ""
'                .RightFooter = text_anexo & num_anexo
'                .LeftMargin = Application.InchesToPoints(0.590551181102362)
'                .RightMargin = Application.InchesToPoints(0.590551181102362)
'                .TopMargin = Application.InchesToPoints(0.590551181102362)
'                .BottomMargin = Application.InchesToPoints(0.590551181102362)
'                .HeaderMargin = Application.InchesToPoints(0.511811023622047)
'                .FooterMargin = Application.InchesToPoints(0.511811023622047)
'                .PrintHeadings = False
'                .PrintGridlines = False
'                .PrintComments = xlPrintNoComments
'                .PrintQuality = 600
'                .CenterHorizontally = True
'                .CenterVertically = False
'                .Orientation = xlPortrait
'                .Draft = False
'                .PaperSize = xlPaperA4
'                .FirstPageNumber = xlAutomatic
'                .Order = xlDownThenOver
'                .BlackAndWhite = False
'                .Zoom = 100
'                .PrintErrors = xlPrintErrorsDisplayed
'                .OddAndEvenPagesHeaderFooter = False
'                .DifferentFirstPageHeaderFooter = False
'                .ScaleWithDocHeaderFooter = True
'                .AlignMarginsHeaderFooter = True
'                .EvenPage.LeftHeader.Text = ""
'                .EvenPage.CenterHeader.Text = ""
'                .EvenPage.RightHeader.Text = ""
'                .EvenPage.LeftFooter.Text = ""
'                .EvenPage.CenterFooter.Text = ""
'                .EvenPage.RightFooter.Text = ""
'                .FirstPage.LeftHeader.Text = ""
'                .FirstPage.CenterHeader.Text = ""
'                .FirstPage.RightHeader.Text = ""
'                .FirstPage.LeftFooter.Text = ""
'                .FirstPage.CenterFooter.Text = ""
'                .FirstPage.RightFooter.Text = ""
'            End With
'            Application.PrintCommunication = True
                      
            Sheets(Array("Sensibilidade - Apresenta��o")).Select
            'Sheets(Array("GR�FICO 1", "GR�FICO 2", "GR�FICO 3")).Select
            ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=ActiveWorkbook.Path & "\" & Desktop & num_anexo & " " & nome, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
            ActiveWorkbook.Save
            
            Sheets("Dados").Select
            Cells.Select
            Selection.Copy
            
            Set novo_arquivo = Application.Workbooks.Add
            Range("A1").Select
            ActiveSheet.Paste
            
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            With Selection.Font
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            With Selection.Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            ActiveWindow.DisplayGridlines = False
            Range("m7:m17").Select
            Selection.ClearContents
            Range("i109:i134").Select
            Selection.ClearContents
            
            ActiveWorkbook.SaveAs Filename:="Dados Brutos " & nome & " - " & furo & "-" & amostra
            ActiveWorkbook.Close
            
            Windows(arquivo).Activate
            ActiveWorkbook.Save
            ActiveWorkbook.Close
                
            num_anexo = num_anexo + 1
            Windows(tabela_amostras).Activate
            ActiveCell.Offset(1, 2).Select
    
    Loop
    
    Application.ScreenUpdating = True
    
    MsgBox "O n�mero dos anexos foram alterados com sucesso.", , "Numerador de anexo"
    
End Sub

Sub anexo_pdf_granulometria()

Application.ScreenUpdating = False

Dim nome As String
Dim tabela_amostras As String
Dim caminho As String
Dim num_anexo As Integer

text_anexo = "RL-5714-GT-057_ANX"


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
            
            'Inserir Benthic aqui
            
            anexo_i = num_anexo
            anexo_f = anexo_i + 1
            
            Sheets("LLLP - Apresenta��o").Select
            ActiveSheet.Range("M78").Select
            ActiveCell.Value = num_anexo
            ActiveCell.Offset(0, -1).Select
            ActiveCell.Value = text_anexo
            
            Sheets(Array("LLLP - Apresenta��o", "Curva Granulom�trica")).Select
            ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=ActiveWorkbook.Path & "\" & Desktop & anexo_i & " - " & anexo_f & " " & nome, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
            
            ActiveWorkbook.Save
            ActiveWorkbook.Close
                
            Windows(tabela_amostras).Activate
            ActiveCell.Offset(1, 3).Select
    
    Loop
    
    Application.ScreenUpdating = True
    
    MsgBox "O n�mero dos anexos foram alterados com sucesso.", , "Numerador de anexo"
    
End Sub
Sub anexo_pdf_adensamento()

Application.ScreenUpdating = False

Dim nome As String
Dim tabela_amostras As String
Dim caminho As String
Dim contador As Integer
Dim num_anexo As Integer
Dim current As Worksheet

text_anexo = "RL-5713-GT-023_ANX00"
num_anexo = 8

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
            
            Sheets("Apresenta��o").Select
            Application.PrintCommunication = False
            With ActiveSheet.PageSetup
                .PrintTitleRows = ""
                .PrintTitleColumns = ""
            End With
            Application.PrintCommunication = True
            ActiveSheet.PageSetup.PrintArea = "$A$1:$V$52"
            Application.PrintCommunication = False
            With ActiveSheet.PageSetup
                .LeftHeader = ""
                .CenterHeader = ""
                .RightHeader = ""
                .LeftFooter = " "
                .CenterFooter = ""
                .RightFooter = text_anexo & num_anexo
                .LeftMargin = Application.InchesToPoints(0.590551181102362)
                .RightMargin = Application.InchesToPoints(0.590551181102362)
                .TopMargin = Application.InchesToPoints(0.393700787401575)
                .BottomMargin = Application.InchesToPoints(0.78740157480315)
                .HeaderMargin = Application.InchesToPoints(0)
                .FooterMargin = Application.InchesToPoints(0.511811023622047)
                .PrintHeadings = False
                .PrintGridlines = False
                .PrintComments = xlPrintNoComments
                .PrintQuality = 600
                .CenterHorizontally = True
                .CenterVertically = True
                .Orientation = xlLandscape
                .Draft = False
                .PaperSize = xlPaperA4
                .FirstPageNumber = xlAutomatic
                .Order = xlDownThenOver
                .BlackAndWhite = False
                .Zoom = 90
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
            
            If num_anexo < "9" Then
                text_anexo = "RL-5713-GT-023_ANX00"
                
            Else
                text_anexo = "RL-5713-GT-023_ANX0"
             
            End If
            
            anexo_i = num_anexo
            num_anexo = num_anexo + 1
            
            Sheets("Curva de Compressibilidade").Select
            Application.PrintCommunication = False
            With ActiveSheet.PageSetup
                .PrintTitleRows = ""
                .PrintTitleColumns = ""
            End With
            Application.PrintCommunication = True
            ActiveSheet.PageSetup.PrintArea = "$A$1:$M$72"
            Application.PrintCommunication = False
            With ActiveSheet.PageSetup
                .LeftHeader = ""
                .CenterHeader = ""
                .RightHeader = ""
                .LeftFooter = " "
                .CenterFooter = ""
                .RightFooter = text_anexo & num_anexo
                .LeftMargin = Application.InchesToPoints(0.590551181102362)
                .RightMargin = Application.InchesToPoints(0.590551181102362)
                .TopMargin = Application.InchesToPoints(0.31496062992126)
                .BottomMargin = Application.InchesToPoints(0.590551181102362)
                .HeaderMargin = Application.InchesToPoints(0.31496062992126)
                .FooterMargin = Application.InchesToPoints(0.31496062992126)
                .PrintHeadings = False
                .PrintGridlines = False
                .PrintComments = xlPrintNoComments
                .PrintQuality = 600
                .CenterHorizontally = True
                .CenterVertically = False
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
            Application.PrintCommunication = True
            
            If num_anexo < "9" Then
                text_anexo = "RL-5713-GT-023_ANX00"
                
            Else
                text_anexo = "RL-5713-GT-023_ANX0"
             
            End If
            
            num_anexo = num_anexo + 1
            
            Sheets("1� Est�gio").Select
            Application.PrintCommunication = False
            With ActiveSheet.PageSetup
                .PrintTitleRows = ""
                .PrintTitleColumns = ""
            End With
            Application.PrintCommunication = True
            ActiveSheet.PageSetup.PrintArea = "$A$1:$M$72"
            Application.PrintCommunication = False
            With ActiveSheet.PageSetup
                .LeftHeader = ""
                .CenterHeader = ""
                .RightHeader = ""
                .LeftFooter = " "
                .CenterFooter = ""
                .RightFooter = text_anexo & num_anexo
                .LeftMargin = Application.InchesToPoints(0.590551181102362)
                .RightMargin = Application.InchesToPoints(0.590551181102362)
                .TopMargin = Application.InchesToPoints(0.31496062992126)
                .BottomMargin = Application.InchesToPoints(0.590551181102362)
                .HeaderMargin = Application.InchesToPoints(0.31496062992126)
                .FooterMargin = Application.InchesToPoints(0.31496062992126)
                .PrintHeadings = False
                .PrintGridlines = False
                .PrintComments = xlPrintNoComments
                .PrintQuality = 600
                .CenterHorizontally = True
                .CenterVertically = False
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
            Application.PrintCommunication = True
            
            If num_anexo < "9" Then
                text_anexo = "RL-5713-GT-023_ANX00"
                
            Else
                text_anexo = "RL-5713-GT-023_ANX0"
             
            End If
            
            num_anexo = num_anexo + 1
            
            
            Sheets("2� Est�gio").Select
            Application.PrintCommunication = False
            With ActiveSheet.PageSetup
                .PrintTitleRows = ""
                .PrintTitleColumns = ""
            End With
            Application.PrintCommunication = True
            ActiveSheet.PageSetup.PrintArea = "$A$1:$M$72"
            Application.PrintCommunication = False
            With ActiveSheet.PageSetup
                .LeftHeader = ""
                .CenterHeader = ""
                .RightHeader = ""
                .LeftFooter = " "
                .CenterFooter = ""
                .RightFooter = text_anexo & num_anexo
                .LeftMargin = Application.InchesToPoints(0.590551181102362)
                .RightMargin = Application.InchesToPoints(0.590551181102362)
                .TopMargin = Application.InchesToPoints(0.31496062992126)
                .BottomMargin = Application.InchesToPoints(0.590551181102362)
                .HeaderMargin = Application.InchesToPoints(0.31496062992126)
                .FooterMargin = Application.InchesToPoints(0.31496062992126)
                .PrintHeadings = False
                .PrintGridlines = False
                .PrintComments = xlPrintNoComments
                .PrintQuality = 600
                .CenterHorizontally = True
                .CenterVertically = False
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
            Application.PrintCommunication = True

            If num_anexo < "9" Then
                text_anexo = "RL-5713-GT-023_ANX00"
                
            Else
                text_anexo = "RL-5713-GT-023_ANX0"
             
            End If
            
            num_anexo = num_anexo + 1
            
            Sheets("3� Est�gio").Select
            Application.PrintCommunication = False
            With ActiveSheet.PageSetup
                .PrintTitleRows = ""
                .PrintTitleColumns = ""
            End With
            Application.PrintCommunication = True
            ActiveSheet.PageSetup.PrintArea = "$A$1:$M$72"
            Application.PrintCommunication = False
            With ActiveSheet.PageSetup
                .LeftHeader = ""
                .CenterHeader = ""
                .RightHeader = ""
                .LeftFooter = " "
                .CenterFooter = ""
                .RightFooter = text_anexo & num_anexo
                .LeftMargin = Application.InchesToPoints(0.590551181102362)
                .RightMargin = Application.InchesToPoints(0.590551181102362)
                .TopMargin = Application.InchesToPoints(0.31496062992126)
                .BottomMargin = Application.InchesToPoints(0.590551181102362)
                .HeaderMargin = Application.InchesToPoints(0.31496062992126)
                .FooterMargin = Application.InchesToPoints(0.31496062992126)
                .PrintHeadings = False
                .PrintGridlines = False
                .PrintComments = xlPrintNoComments
                .PrintQuality = 600
                .CenterHorizontally = True
                .CenterVertically = False
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
            Application.PrintCommunication = True
     
            If num_anexo < "9" Then
                text_anexo = "RL-5713-GT-023_ANX00"
                
            Else
                text_anexo = "RL-5713-GT-023_ANX0"
             
            End If
            
            num_anexo = num_anexo + 1
            
            Sheets("4� Est�gio").Select
            Application.PrintCommunication = False
            With ActiveSheet.PageSetup
                .PrintTitleRows = ""
                .PrintTitleColumns = ""
            End With
            Application.PrintCommunication = True
            ActiveSheet.PageSetup.PrintArea = "$A$1:$M$72"
            Application.PrintCommunication = False
            With ActiveSheet.PageSetup
                .LeftHeader = ""
                .CenterHeader = ""
                .RightHeader = ""
                .LeftFooter = " "
                .CenterFooter = ""
                .RightFooter = text_anexo & num_anexo
                .LeftMargin = Application.InchesToPoints(0.590551181102362)
                .RightMargin = Application.InchesToPoints(0.590551181102362)
                .TopMargin = Application.InchesToPoints(0.31496062992126)
                .BottomMargin = Application.InchesToPoints(0.590551181102362)
                .HeaderMargin = Application.InchesToPoints(0.31496062992126)
                .FooterMargin = Application.InchesToPoints(0.31496062992126)
                .PrintHeadings = False
                .PrintGridlines = False
                .PrintComments = xlPrintNoComments
                .PrintQuality = 600
                .CenterHorizontally = True
                .CenterVertically = False
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
            Application.PrintCommunication = True

            If num_anexo < "9" Then
                text_anexo = "RL-5713-GT-023_ANX00"
                
            Else
                text_anexo = "RL-5713-GT-023_ANX0"
             
            End If
            
            num_anexo = num_anexo + 1
            
            Sheets("5� Est�gio").Select
            Application.PrintCommunication = False
            With ActiveSheet.PageSetup
                .PrintTitleRows = ""
                .PrintTitleColumns = ""
            End With
            Application.PrintCommunication = True
            ActiveSheet.PageSetup.PrintArea = "$A$1:$M$72"
            Application.PrintCommunication = False
            With ActiveSheet.PageSetup
                .LeftHeader = ""
                .CenterHeader = ""
                .RightHeader = ""
                .LeftFooter = " "
                .CenterFooter = ""
                .RightFooter = text_anexo & num_anexo
                .LeftMargin = Application.InchesToPoints(0.590551181102362)
                .RightMargin = Application.InchesToPoints(0.590551181102362)
                .TopMargin = Application.InchesToPoints(0.31496062992126)
                .BottomMargin = Application.InchesToPoints(0.590551181102362)
                .HeaderMargin = Application.InchesToPoints(0.31496062992126)
                .FooterMargin = Application.InchesToPoints(0.31496062992126)
                .PrintHeadings = False
                .PrintGridlines = False
                .PrintComments = xlPrintNoComments
                .PrintQuality = 600
                .CenterHorizontally = True
                .CenterVertically = False
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
            Application.PrintCommunication = True
     
            If num_anexo < "9" Then
                text_anexo = "RL-5713-GT-023_ANX00"
                
            Else
                text_anexo = "RL-5713-GT-023_ANX0"
             
            End If
            
            num_anexo = num_anexo + 1
            
            Sheets("6� Est�gio").Select
            Application.PrintCommunication = False
            With ActiveSheet.PageSetup
                .PrintTitleRows = ""
                .PrintTitleColumns = ""
            End With
            Application.PrintCommunication = True
            ActiveSheet.PageSetup.PrintArea = "$A$1:$M$72"
            Application.PrintCommunication = False
            With ActiveSheet.PageSetup
                .LeftHeader = ""
                .CenterHeader = ""
                .RightHeader = ""
                .LeftFooter = " "
                .CenterFooter = ""
                .RightFooter = text_anexo & num_anexo
                .LeftMargin = Application.InchesToPoints(0.590551181102362)
                .RightMargin = Application.InchesToPoints(0.590551181102362)
                .TopMargin = Application.InchesToPoints(0.31496062992126)
                .BottomMargin = Application.InchesToPoints(0.590551181102362)
                .HeaderMargin = Application.InchesToPoints(0.31496062992126)
                .FooterMargin = Application.InchesToPoints(0.31496062992126)
                .PrintHeadings = False
                .PrintGridlines = False
                .PrintComments = xlPrintNoComments
                .PrintQuality = 600
                .CenterHorizontally = True
                .CenterVertically = False
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
            Application.PrintCommunication = True

            If num_anexo < "9" Then
                text_anexo = "RL-5713-GT-023_ANX00"
                
            Else
                text_anexo = "RL-5713-GT-023_ANX0"
             
            End If
            
            num_anexo = num_anexo + 1
            
            Sheets("7� Est�gio").Select
            Application.PrintCommunication = False
            With ActiveSheet.PageSetup
                .PrintTitleRows = ""
                .PrintTitleColumns = ""
            End With
            Application.PrintCommunication = True
            ActiveSheet.PageSetup.PrintArea = "$A$1:$M$72"
            Application.PrintCommunication = False
            With ActiveSheet.PageSetup
                .LeftHeader = ""
                .CenterHeader = ""
                .RightHeader = ""
                .LeftFooter = " "
                .CenterFooter = ""
                .RightFooter = text_anexo & num_anexo
                .LeftMargin = Application.InchesToPoints(0.590551181102362)
                .RightMargin = Application.InchesToPoints(0.590551181102362)
                .TopMargin = Application.InchesToPoints(0.31496062992126)
                .BottomMargin = Application.InchesToPoints(0.590551181102362)
                .HeaderMargin = Application.InchesToPoints(0.31496062992126)
                .FooterMargin = Application.InchesToPoints(0.31496062992126)
                .PrintHeadings = False
                .PrintGridlines = False
                .PrintComments = xlPrintNoComments
                .PrintQuality = 600
                .CenterHorizontally = True
                .CenterVertically = False
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
            Application.PrintCommunication = True
     
            If num_anexo < "9" Then
                text_anexo = "RL-5713-GT-023_ANX00"
                
            Else
                text_anexo = "RL-5713-GT-023_ANX0"
             
            End If
            
            num_anexo = num_anexo + 1
            
             Sheets("8� Est�gio").Select
            Application.PrintCommunication = False
            With ActiveSheet.PageSetup
                .PrintTitleRows = ""
                .PrintTitleColumns = ""
            End With
            Application.PrintCommunication = True
            ActiveSheet.PageSetup.PrintArea = "$A$1:$M$72"
            Application.PrintCommunication = False
            With ActiveSheet.PageSetup
                .LeftHeader = ""
                .CenterHeader = ""
                .RightHeader = ""
                .LeftFooter = " "
                .CenterFooter = ""
                .RightFooter = text_anexo & num_anexo
                .LeftMargin = Application.InchesToPoints(0.590551181102362)
                .RightMargin = Application.InchesToPoints(0.590551181102362)
                .TopMargin = Application.InchesToPoints(0.31496062992126)
                .BottomMargin = Application.InchesToPoints(0.590551181102362)
                .HeaderMargin = Application.InchesToPoints(0.31496062992126)
                .FooterMargin = Application.InchesToPoints(0.31496062992126)
                .PrintHeadings = False
                .PrintGridlines = False
                .PrintComments = xlPrintNoComments
                .PrintQuality = 600
                .CenterHorizontally = True
                .CenterVertically = False
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
            Application.PrintCommunication = True
     
            If num_anexo < "9" Then
                text_anexo = "RL-5713-GT-023_ANX00"
                
            Else
                text_anexo = "RL-5713-GT-023_ANX0"
             
            End If
            
            anexo_f = num_anexo
            num_anexo = num_anexo + 1
            
            Sheets(Array("Apresenta��o", "Curva de Compressibilidade", "1� Est�gio", "2� Est�gio", "3� Est�gio", "4� Est�gio", "5� Est�gio", "6� Est�gio", "7� Est�gio", "8� Est�gio")).Select
            ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=ActiveWorkbook.Path & "\" & Desktop & anexo_i & " - " & anexo_f & " " & nome, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
            
            ActiveWorkbook.Save
            ActiveWorkbook.Close
                
            Windows(tabela_amostras).Activate
            ActiveCell.Offset(1, 2).Select
    
    Loop
    
    Application.ScreenUpdating = True
    
    MsgBox "O n�mero dos anexos foram alterados com sucesso.", , "Numerador de anexo"
    
End Sub
Sub anexo_pdf_sensitividade()

Application.ScreenUpdating = False

Dim nome As String
Dim tabela_amostras As String
Dim caminho As String
Dim contador As Integer
Dim num_anexo As Integer
Dim current As Worksheet

text_anexo = "RL-5714-GT-049_ANX00"
num_anexo = 2

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
            
            Sheets("Sensibilidade - Apresenta��o").Select
            Application.PrintCommunication = False
            With ActiveSheet.PageSetup
                .PrintTitleRows = ""
                .PrintTitleColumns = ""
            End With
            Application.PrintCommunication = True
            ActiveSheet.PageSetup.PrintArea = "$A$1:$K$74"
            Application.PrintCommunication = False
            With ActiveSheet.PageSetup
                .LeftHeader = ""
                .CenterHeader = ""
                .RightHeader = ""
                .LeftFooter = " "
                .CenterFooter = ""
                .RightFooter = text_anexo & num_anexo
                .LeftMargin = Application.InchesToPoints(0.590551181102362)
                .RightMargin = Application.InchesToPoints(0.590551181102362)
                .TopMargin = Application.InchesToPoints(0.590551181102362)
                .BottomMargin = Application.InchesToPoints(0.590551181102362)
                .HeaderMargin = Application.InchesToPoints(0.511811023622047)
                .FooterMargin = Application.InchesToPoints(0.511811023622047)
                .PrintHeadings = False
                .PrintGridlines = False
                .PrintComments = xlPrintNoComments
                .PrintQuality = 600
                .CenterHorizontally = True
                .CenterVertically = False
                .Orientation = xlPortrait
                .Draft = False
                .PaperSize = xlPaperA4
                .FirstPageNumber = xlAutomatic
                .Order = xlDownThenOver
                .BlackAndWhite = False
                .Zoom = 100
                .PrintErrors = xlPrintErrorsDisplayed
                .OddAndEvenPagesHeaderFooter = False
                .DifferentFirstPageHeaderFooter = False
                .ScaleWithDocHeaderFooter = True
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
            
            'If num_anexo < "9" Then
            '    text_anexo = "RL-5714-GT-038_ANX00"
                
            'ElseIf num_anexo < "99" Then
            '    text_anexo = "RL-5714-GT-038_ANX0"
                
            'Else
            '    text_anexo = "RL-5714-GT-038_ANX"
                
            'End If
            
            Sheets(Array("Sensibilidade - Apresenta��o")).Select
            ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=ActiveWorkbook.Path & "\" & Desktop & num_anexo & " - " & nome, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
            
            num_anexo = num_anexo + 1
            
            ActiveWorkbook.Save
            ActiveWorkbook.Close
                
            Windows(tabela_amostras).Activate
            ActiveCell.Offset(1, 2).Select
    
    Loop
    
    Application.ScreenUpdating = True
    
    MsgBox "O n�mero dos anexos foram alterados e os PDFs gerados com sucesso!", , "Numerador de anexo"
    
End Sub



Sub pdf_Bender()

Application.ScreenUpdating = False

Dim nome As String
Dim tabela_amostras As String
Dim caminho As String
Dim contador As Integer
    
    tabela_amostras = "Gerar PDF.xlsm"
    Windows(tabela_amostras).Activate
    ActiveSheet.Range("f3").Select
    
    
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
            
            If num_anexo < 9 Then
                text_anexo = "RL-5714-GT-056_ANX00"
            
            ElseIf num_anexo < 99 Then
                text_anexo = "RL-5714-GT-056_ANX0"
                
            Else
                text_anexo = "RL-5714-GT-056_ANX"
                
            End If
            
            Sheets("GR�FICO 1").Select
            Range("I75").Select
            ActiveCell.Value = text_anexo & num_anexo
            
                If num_anexo < 9 Then
                    text_anexo = "RL-5714-GT-056_ANX00"
                
                ElseIf num_anexo < 99 Then
                    text_anexo = "RL-5714-GT-056_ANX0"
                    
                Else
                    text_anexo = "RL-5714-GT-056_ANX"
                    
                End If
                
            num_anexo = num_anexo + 1
            
            
            Sheets("GR�FICO 2").Select
            Range("I75").Select
            ActiveCell.Value = text_anexo & num_anexo
            
                If num_anexo < 9 Then
                    text_anexo = "RL-5714-GT-056_ANX00"
                
                ElseIf num_anexo < 99 Then
                    text_anexo = "RL-5714-GT-056_ANX0"
                    
                Else
                    text_anexo = "RL-5714-GT-056_ANX"
                    
                End If
                
            num_anexo = num_anexo + 1
            
            Sheets("GR�FICO 3").Select
            Range("I75").Select
            ActiveCell.Value = text_anexo & num_anexo
            
            
            Sheets(Array("GR�FICO 1", "GR�FICO 2", "GR�FICO 3")).Select
            ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=ActiveWorkbook.Path & "\" & Desktop & num_anexo - 2 & "-" & num_anexo & " " & nome, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
            
            ActiveWorkbook.Close (savechanges = True)
                
            Windows(tabela_amostras).Activate
            ActiveCell.Offset(1, 3).Select
    
    Loop
    
    Application.ScreenUpdating = True
    
    MsgBox "Conclu�do"
    
End Sub

Sub pdf_DSS()

Application.ScreenUpdating = False

Dim nome As String
Dim tabela_amostras As String
Dim caminho As String
Dim contador As Integer
    
    text_anexo = "RL-5714-GT-044_ANX"
    
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
            
            Sheets(Array("Apresenta��o DSS Est�tico CP1", "Curvas Adensamento CP1")).Select
            ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=ActiveWorkbook.Path & "\" & Desktop & num_anexo & " - " & num_anexo + 1 & " - " & nome, Quality:=xlQualityStandard, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
            
            Windows(arquivo).Activate
            ActiveWorkbook.Close (savechanges = False)
            
            Windows(tabela_amostras).Activate
            ActiveCell.Offset(1, 3).Select
    
    Loop
    
    Application.ScreenUpdating = True
    
    MsgBox "Arquivos em PDF gerados!", , "Gerador de PDF"
    
End Sub
