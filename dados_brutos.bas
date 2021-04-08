Attribute VB_Name = "dados_brutos"

Sub dados_brutos_sensitividade()

Application.ScreenUpdating = False

Dim furo As String
Dim amostra As String
Dim nome As String
Dim tabela_amostras As String
Dim caminho As String
Dim contador As Integer
Dim num_anexo As Integer
Dim current As Worksheet

    tabela_amostras = "Gerar Dados Brutos.xlsm"
    Windows(tabela_amostras).Activate
    ActiveSheet.Range("E3").Select
    
    Do While ActiveCell.Value <> ""
    
            nome = ActiveCell.Value
            ActiveCell.Offset(0, -1).Select
            arquivo = ActiveCell.Value
            ActiveCell.Offset(0, -1).Select
            caminho = ActiveCell.Value
            Workbooks.Open caminho, UpdateLinks:=0, Editable:=True
           
            Windows(arquivo).Activate

            Sheets("Dados amostra").Select
            Cells.Select
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
                
            Cells.Select
            Application.CutCopyMode = False
            With Selection.Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            With Selection.Font
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            
            Range("I109:I132").Select
            Selection.Delete
            Range("M7:M18").Select
            Selection.Delete
            furo = Range("G12").Value
            amostra = Range("G13").Value
            
            Sheets("Tabela").Select
            ActiveWorkbook.Unprotect ("geo@2017")
            Sheets("Tabela").Delete
            Sheets("Sensibilidade - Apresenta��o").Delete
            
            Sheets("Dados").Select
            With ActiveWorkbook.Sheets("Dados").Tab
                .ColorIndex = xlNone
                .TintAndShade = 0
            End With
            Range("K1").Select
            
            
            ActiveWorkbook.SaveAs Filename:="Dados Brutos " & nome & " - " & furo & amostra
            ActiveWorkbook.Close
                
            Windows(tabela_amostras).Activate
            ActiveCell.Offset(1, 2).Select
    
    Loop
    
    Application.ScreenUpdating = True
    
    MsgBox "Processo conclu�do!"
    
End Sub

Sub dados_brutos_adensamento()
Application.ScreenUpdating = False

Dim nome As String
Dim furo As String
Dim amostra As String
Dim tabela_amostras As String
Dim caminho As String
Dim contador As Integer

    contador = 1
    tabela_amostras = "Gerar Dados Brutos.xlsm"
    Windows(tabela_amostras).Activate
    ActiveSheet.Range("E7").Select

    
    Do While ActiveCell.Value <> ""
    
            nome = ActiveCell.Value
            ActiveCell.Offset(0, -1).Select
            arquivo = ActiveCell.Value
            ActiveCell.Offset(0, -1).Select
            caminho = ActiveCell.Value
            Workbooks.Open caminho, UpdateLinks:=0
            Windows(arquivo).Activate

            Sheets("Dados").Select
            Cells.Select
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            With Selection.Font
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            Range("V32:V126").Select
            Selection.ClearContents
            Range("A28:T125").Select
            Selection.ClearComments
            
            Sheets("Apresenta��o").Select
            furo = Range("I12").Value
            amostra = Range("I13").Value
            
            ActiveSheet.Shapes.Range(Array("Bal�o de Fala: Oval 2")).Select
            Selection.Delete
            Cells.Select
            
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
                
            Cells.Select
            With Selection.Font
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            Range("A54:S78").Select
            Range("S78").Activate
            Selection.ClearContents
            Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            Selection.Borders(xlEdgeLeft).LineStyle = xlNone
            Selection.Borders(xlEdgeTop).LineStyle = xlNone
            Selection.Borders(xlEdgeBottom).LineStyle = xlNone
            Selection.Borders(xlEdgeRight).LineStyle = xlNone
            Selection.Borders(xlInsideVertical).LineStyle = xlNone
            Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
            
            Sheets("Curva de Compressibilidade").Delete
'            Cells.Select
'            Range("A5").Activate
'            Selection.Copy
'            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'                :=False, Transpose:=False
'            Range("N22:S27").Delete
'            Range("L11").Select
                
            Sheets("1� Est�gio").Select
            Cells.Select
            Application.CutCopyMode = False
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Range("R35").Select
            Application.CutCopyMode = False
            Selection.ClearContents
            Range("L11").Select
            
            Sheets("2� Est�gio").Select
            Cells.Select
            Application.CutCopyMode = False
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Range("R35").Select
            Application.CutCopyMode = False
            Selection.ClearContents
            Range("L11").Select
            
            Sheets("3� Est�gio").Select
            Cells.Select
            Application.CutCopyMode = False
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Range("R35").Select
            Application.CutCopyMode = False
            Selection.ClearContents
            Range("L11").Select
            
            Sheets("4� Est�gio").Select
            Cells.Select
            Application.CutCopyMode = False
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Range("R35").Select
            Application.CutCopyMode = False
            Selection.ClearContents
            Range("L11").Select
            
            Sheets("5� Est�gio").Select
            Cells.Select
            Application.CutCopyMode = False
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Range("R35").Select
            Application.CutCopyMode = False
            Selection.ClearContents
            Range("L11").Select
            
            Sheets("6� Est�gio").Select
            Cells.Select
            Application.CutCopyMode = False
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Range("R35").Select
            Application.CutCopyMode = False
            Selection.ClearContents
            Range("L11").Select
            
            Sheets("7� Est�gio").Select
            Cells.Select
            Application.CutCopyMode = False
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Range("R35").Select
            Application.CutCopyMode = False
            Selection.ClearContents
            Range("L11").Select
            
            Sheets("8� Est�gio").Select
            Cells.Select
            Application.CutCopyMode = False
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Range("R35").Select
            Application.CutCopyMode = False
            Selection.ClearContents
            Range("L11").Select
            
            'Exclus�o das abas'
            'Sheets("9� Est�gio").Visible = True
            'Sheets("9� Est�gio").Delete
            
            Sheets("CG 1").Visible = True
            Sheets("CG 1").Delete
            
            Sheets(" Taylor 1").Visible = True
            Sheets(" Taylor 1").Delete
            
            Sheets("CG 2").Visible = True
            Sheets("CG 2").Delete
            
            Sheets("Taylor 2").Visible = True
            Sheets("Taylor 2").Delete
            
            Sheets("CG 3").Visible = True
            Sheets("CG 3").Delete
            
            Sheets("Taylor 3").Visible = True
            Sheets("Taylor 3").Delete
            
            Sheets("CG 4").Visible = True
            Sheets("CG 4").Delete
            
            Sheets("Taylor 4").Visible = True
            Sheets("Taylor 4").Delete
            
            Sheets("CG 5").Visible = True
            Sheets("CG 5").Delete
            
            Sheets("Taylor 5").Visible = True
            Sheets("Taylor 5").Delete
            
            Sheets("CG 6").Visible = True
            Sheets("CG 6").Delete
            
            Sheets("Taylor 6").Visible = True
            Sheets("Taylor 6").Delete
            
            Sheets("CG 7").Visible = True
            Sheets("CG 7").Delete
            
            Sheets("Taylor 7").Visible = True
            Sheets("Taylor 7").Delete
            
            Sheets("CG 8").Visible = True
            Sheets("CG 8").Delete
            
            Sheets("Taylor 8").Visible = True
            Sheets("Taylor 8").Delete
            
            Sheets("CG 9").Visible = True
            Sheets("CG 9").Delete
            
            Sheets("Taylor 9").Visible = True
            Sheets("Taylor 9").Delete
            
            Sheets("Apresenta��o Antiga").Visible = True
            Sheets("Apresenta��o Antiga").Delete
            
            Sheets("Curva de Compressibilidade").Visible = True
            Sheets("Curva de Compressibilidade").Delete
            
            'Cor das abas'
            Sheets("Dados").Select
            With ActiveWorkbook.Sheets("Dados").Tab
                .ColorIndex = xlColorIndexNone
                .TintAndShade = 0
            End With
            Sheets("Apresenta��o").Select
            With ActiveWorkbook.Sheets("Apresenta��o").Tab
                .ColorIndex = xlColorIndexNone
                .TintAndShade = 0
            End With
            Sheets("Curva de Compressibilidade").Select
            With ActiveWorkbook.Sheets("Curva de Compressibilidade").Tab
                .ColorIndex = xlColorIndexNone
                .TintAndShade = 0
            End With
            Sheets("1� Est�gio").Select
            With ActiveWorkbook.Sheets("1� Est�gio").Tab
                .ColorIndex = xlColorIndexNone
                .TintAndShade = 0
            End With
            Sheets("2� Est�gio").Select
            With ActiveWorkbook.Sheets("2� Est�gio").Tab
                .ColorIndex = xlColorIndexNone
                .TintAndShade = 0
            End With
            Sheets("3� Est�gio").Select
            With ActiveWorkbook.Sheets("3� Est�gio").Tab
                .ColorIndex = xlColorIndexNone
                .TintAndShade = 0
            End With
            Sheets("4� Est�gio").Select
            With ActiveWorkbook.Sheets("4� Est�gio").Tab
                .ColorIndex = xlColorIndexNone
                .TintAndShade = 0
            End With
            Sheets("5� Est�gio").Select
            With ActiveWorkbook.Sheets("5� Est�gio").Tab
                .ColorIndex = xlColorIndexNone
                .TintAndShade = 0
            End With
            Sheets("6� Est�gio").Select
            With ActiveWorkbook.Sheets("6� Est�gio").Tab
                .ColorIndex = xlColorIndexNone
                .TintAndShade = 0
            End With
            Sheets("7� Est�gio").Select
            With ActiveWorkbook.Sheets("7� Est�gio").Tab
                .ColorIndex = xlColorIndexNone
                .TintAndShade = 0
            End With
            Sheets("8� Est�gio").Select
            With ActiveWorkbook.Sheets("8� Est�gio").Tab
                .ColorIndex = xlColorIndexNone
                .TintAndShade = 0
            End With
            
            Sheets(Array("Apresenta��o", "1� Est�gio", "2� Est�gio", "3� Est�gio", "4� Est�gio", "5� Est�gio", "6� Est�gio", "7� Est�gio", "8� Est�gio", "9� Est�gio", "10� Est�gio ")).Select
            
            
            ActiveWorkbook.SaveAs Filename:="Dados Brutos " & nome & " - " & furo & "-" & amostra
            ActiveWorkbook.Close
                
            Windows(tabela_amostras).Activate
            ActiveCell.Offset(1, 2).Select
    
    Loop
    
    Application.ScreenUpdating = True
    
    MsgBox "Processo Conclu�do"
    
End Sub


Sub dados_brutos_granulometria()

Application.ScreenUpdating = False

Dim amostra As String
Dim nome As String
Dim tabela_amostras As String
Dim caminho As String
Dim anexo As Integer

tabela_amostras = "Gerar PDF.xlsm"
Windows(tabela_amostras).Activate
Range("E3").Select

Do While ActiveCell.Value <> ""

    nome = ActiveCell.Value
    ActiveCell.Offset(0, -1).Select
    
    arquivo = ActiveCell.Value
    ActiveCell.Offset(0, -1).Select
    
    caminho = ActiveCell.Value
    
    Workbooks.Open caminho, UpdateLinks:=0
    
    Windows(arquivo).Activate
    
    'SALVANDO SOMENTE VALORES
        Sheets("Dados amostra").Select
        
        Sheets("Dados Granulometria").Select
            Cells.Select
            Application.CutCopyMode = False
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            furo = Range("C10").Value
            amostra = Range("I8").Value
                
        Sheets("Curva Granulom�trica").Select
            Cells.Select
            Application.CutCopyMode = False
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
                
        Sheets("Dados amostra").Select
            Cells.Select
            Application.CutCopyMode = False
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
        
        Sheets("Dados Massa Espec�fica").Select
            Cells.Select
            Application.CutCopyMode = False
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
        
        Sheets("Dados Limites de Atterberg").Select
            Cells.Select
            Application.CutCopyMode = False
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
    
    
        'DADOS AMOSTRA
            Sheets("Dados amostra").Select
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
            
            Range("A1").Select
            Sheets("Dados amostra").Select
                With ActiveWorkbook.Sheets("Dados amostra").Tab
                    .ColorIndex = xlNone
                    .TintAndShade = 0
                End With
                
            Range("F28:H32").Select
            Selection.ClearContents
            
            Range("F40:G45").Select
            Selection.ClearContents
            
    'DADOS GRANULOMETRIA
        Sheets("Dados Granulometria").Select
            With Selection.Font
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            With Selection.Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            Range("A26:M81").Select
            Range("K82").Activate
            Selection.ClearComments
            Range("A61:M77").Select
            Selection.ClearComments
            Columns("N:AO").Select
            Selection.Delete
            Range("K82:L106").Select
            Selection.ClearContents
            Selection.Borders(xlDiagonalDown).LineStyle = xlNone
            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
            Selection.Borders(xlEdgeLeft).LineStyle = xlNone
            Selection.Borders(xlEdgeTop).LineStyle = xlNone
            Selection.Borders(xlEdgeBottom).LineStyle = xlNone
            Selection.Borders(xlEdgeRight).LineStyle = xlNone
            Selection.Borders(xlInsideVertical).LineStyle = xlNone
            Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
            Range("A1").Select
            Sheets("Dados Granulometria").Select
                With ActiveWorkbook.Sheets("Dados Granulometria").Tab
                    .ColorIndex = xlNone
                    .TintAndShade = 0
                End With
            ActiveWindow.DisplayGridlines = False
            Range("A1").Select
    
    'DADOS MASSA ESPEC�FICA
        Sheets("Dados Massa Espec�fica").Select
            With Selection.Font
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            With Selection.Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            Columns("M:O").Select
            Selection.Delete
            Range("F36").Select
            Selection.ClearContents
            Range("A1").Select
            Sheets("Dados Massa Espec�fica").Select
                With ActiveWorkbook.Sheets("Dados Massa Espec�fica").Tab
                    .ColorIndex = xlNone
                    .TintAndShade = 0
                End With
    
    'DADOS LIMITES DE ATTERBERG
        Sheets("Dados Limites de Atterberg").Select
            With Selection.Font
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            With Selection.Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            Rows("69:96").Select
            Selection.Delete
            Columns("K:AC").Select
            Selection.Delete
            On Error Resume Next
                ActiveSheet.Shapes.Range(Array("Bal�o de Pensamento: Nuvem 1")).Select
            Selection.Delete
            Range("A1").Select
            Sheets("Dados Limites de Atterberg").Select
                With ActiveWorkbook.Sheets("Dados Limites de Atterberg").Tab
                    .ColorIndex = xlNone
                    .TintAndShade = 0
                End With
    
    'CURVA GRANULOM�TRICA
        Sheets("Curva Granulom�trica").Select
            With Selection.Font
                .ThemeColor = xlThemeColorLight1
                .TintAndShade = 0
            End With
            With Selection.Interior
                .Pattern = xlNone
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
            Rows("41:51").Select
            Selection.Delete
            Range("P1").Select
            Columns("Q:AC").Select
            Selection.Delete
            On Error Resume Next
                ActiveSheet.Shapes.Range(Array("Bal�o de Pensamento: Nuvem 1")).Select
            Selection.Delete
            Range("A1").Select
            Sheets("Curva Granulom�trica").Select
                With ActiveWorkbook.Sheets("Curva Granulom�trica").Tab
                    .ColorIndex = xlNone
                    .TintAndShade = 0
                End With
            
    'DELETANDO DEMAIS PLANILHAS
        Sheets("LLLP - Apresenta��o").Select
        ActiveWindow.SelectedSheets.Delete

        Sheets("Equipe").Select
        ActiveWindow.SelectedSheets.Delete
    
    
    'SALVANDO E FECHANDO A PLANILHA
        ActiveWorkbook.SaveAs Filename:=ActiveWorkbook.Path & "\" & Desktop & "Dados Brutos " & nome & " - " & furo & "-" & amostra
        ActiveWorkbook.Close
                
        Windows(tabela_amostras).Activate
        ActiveCell.Offset(1, 2).Select
        
Loop

Application.ScreenUpdating = True

    MsgBox "Os dados brutos foram gerados!"

End Sub
Sub dados_brutos_DSS()

Application.ScreenUpdating = False

Dim nome As String
Dim tabela_amostras As String
Dim caminho As String
Dim modelo As String
Dim amostra As String
Dim novo_arquivo As Workbook

    tabela_amostras = "Gerar Dados Brutos.xlsm"
    modelo = "DSS Modelo.xlsx"
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
            
            Sheets("Dados Brutos").Select
            furo = Range("d3").Value
            amostra = Range("d4").Value
            
            dados_brutos_nome = "Dados Brutos " & nome & " - " & furo & "-" & amostra
            dados_brutos_arquivo = dados_brutos_nome & ".xlsx"

            Sheets("Apresenta��o DSS Est�tico CP1").Select
            Cells.Select
            Selection.Copy
            
            Set novo_arquivo = Application.Workbooks.Add
            Range("A1").Select
            ActiveSheet.Paste
            ActiveWindow.DisplayGridlines = False
            
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
            
            ActiveSheet.Name = "Apresenta��o DSS Est�tico CP1"
            Range("A1").Select
            ActiveWorkbook.SaveAs Filename:=dados_brutos_nome
            
            Windows(arquivo).Activate
            Sheets("Curvas Adensamento CP1").Select
            Cells.Select
            Selection.Copy
            
            Windows(dados_brutos_arquivo).Activate
            Sheets.Add(After:=Sheets("Apresenta��o DSS Est�tico CP1")).Name = "Curvas Adensamento CP1"
            Range("A1").Select
            ActiveSheet.Paste
            Range("M63:N63").ClearContents
            Range("T31:U31").ClearContents
            
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
            Range("A1").Select
            
            Windows(arquivo).Activate
            Sheets("Dados Brutos").Select
            Cells.Select
            Selection.Copy
            
            Windows(dados_brutos_arquivo).Activate
            Sheets.Add(After:=Sheets("Curvas Adensamento CP1")).Name = "Dados Brutos"
            Range("A1").Select
            ActiveSheet.Paste
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            Range("A1").Select
            
            Windows(dados_brutos_arquivo).Activate
            ActiveWorkbook.Save
            ActiveWorkbook.Close
'
            Windows(arquivo).Activate
            ActiveWorkbook.Close (savechanges = False)
            
            Windows(tabela_amostras).Activate
            ActiveCell.Offset(1, 2).Select
    
    Loop
    
    Application.ScreenUpdating = True
    
End Sub

Sub db_vale_CREG_todos()

Application.ScreenUpdating = False

Dim nome As String
Dim tabela_amostras As String
Dim caminho As String
Dim contador As Integer
Dim arquivo As String
Dim amostra As String
Dim modelo As String
Dim caminho_modelo As String
    
    modelo = "CREG 539-540-20 - Todos.xls"
    tabela_amostras = "1. Gerar Dados Brutos Vale.xlsm"
    
    Windows(tabela_amostras).Activate
    ActiveSheet.Range("E4").Select
    
    Do While ActiveCell.Value <> ""
    
            nome = ActiveCell.Value
            ActiveCell.Offset(0, -1).Select
            arquivo = ActiveCell.Value
            ActiveCell.Offset(0, -1).Select
            caminho = ActiveCell.Value
            Workbooks.Open caminho, UpdateLinks:=0
            Windows(arquivo).Activate
            
            'Windows(arquivo).Activate
            'Sheets("GPS").Delete
            
            'Windows(modelo).Activate
            'Sheets("DB GPS").Select
            'Cells.Select
            'Selection.Copy
            
            'Windows(arquivo).Activate
            'Sheets.Add After:=ActiveSheet
            'ActiveSheet.Paste
            'ActiveSheet.Name = "GPS"
            'ActiveWorkbook.ChangeLink Name:=modelo, NewName:=arquivo, Type:=xlExcelLinks
            
            Sheets("GPS").Select
            Cells.Select
            Selection.Copy
            
            Set novo_arquivo = Application.Workbooks.Add
            Range("A1").Select
            ActiveSheet.Paste
                        
            ActiveSheet.Name = "GPS"
            
            Cells.Select
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            
            amostra = Range("B3").Value
            Range("A1").Select
            
            ActiveWorkbook.SaveAs Filename:=nome & " - " & amostra
            ActiveWorkbook.Close
            
            Windows(arquivo).Activate
            ActiveWorkbook.Close (savechanges = False)
            
                
            Windows(tabela_amostras).Activate
            ActiveCell.Offset(1, 2).Select
    
    Loop
    
Application.ScreenUpdating = True

End Sub

Sub dados_triaxial_parte1()

Application.ScreenUpdating = False

Dim nome As String
Dim tabela_amostras As String
Dim caminho As String
Dim arquivo As String
Dim amostra As String
Dim modelo As String
Dim caminho_modelo As String
Dim ensaio As String
    
    'Sobre o ensaio
    ensaio = "CIU"
    modelo = "DB " & ensaio & " Modelo MSA.xlsx"
    tabela_amostras = "1. Gerar Dados Brutos Vale.xlsm"
    Windows(tabela_amostras).Activate
    
    'Abrir o modelo padr�o
    Sheets(ensaio).Select
    ActiveSheet.Range("B3").Select
    caminho_modelo = ActiveCell.Value
    On Error Resume Next
        Workbooks.Open caminho_modelo, UpdateLinks:=0
    
    Windows(tabela_amostras).Activate
    Sheets(ensaio).Select
    ActiveSheet.Range("E4").Select
    
    Do While ActiveCell.Value <> ""
    
            nome = ActiveCell.Value
            ActiveCell.Offset(0, -1).Select
            arquivo = ActiveCell.Value
            ActiveCell.Offset(0, -1).Select
            caminho = ActiveCell.Value
            Workbooks.Open caminho, UpdateLinks:=0
            Windows(arquivo).Activate
        
            Windows(modelo).Activate
            Sheets("Dados Brutos").Select
            Cells.Select
            Selection.Copy
                
            Windows(arquivo).Activate
            Sheets("Dados Brutos").Select
            ActiveSheet.Paste
            ActiveWorkbook.ChangeLink Name:=modelo, NewName:=arquivo, Type:=xlExcelLinks
            
            'Windows(modelo).Activate
            'Sheets("DB CP1").Cells.Copy
                
            'Windows(arquivo).Activate
            'Sheets.Add After:=ActiveSheet
            'ActiveSheet.Paste
            'ActiveSheet.Name = "DB CP1"
            'Sheets("DB CP1").Select
            'ActiveWorkbook.ChangeLink Name:=modelo, NewName:=arquivo, Type:=xlExcelLinks
            
            'Windows(modelo).Activate
            'Sheets("DB CP2").Cells.Copy
            
            'Windows(arquivo).Activate
            'Sheets.Add After:=ActiveSheet
            'ActiveSheet.Paste
            'ActiveSheet.Name = "DB CP2"
            'Sheets("DB CP2").Select
            'ActiveWorkbook.ChangeLink Name:=modelo, NewName:=arquivo, Type:=xlExcelLinks
            
            'Windows(modelo).Activate
            'Sheets("DB CP3").Cells.Copy
                
            'Windows(arquivo).Activate
            'Sheets.Add After:=ActiveSheet
            'ActiveSheet.Paste
            'ActiveSheet.Name = "DB CP3"
            'Sheets("DB CP3").Select
            'ActiveWorkbook.ChangeLink Name:=modelo, NewName:=arquivo, Type:=xlExcelLinks
            
            Windows(modelo).Activate
            Sheets("DB CP1").Cells.Copy
                
            Windows(arquivo).Activate
            Sheets.Add After:=ActiveSheet
            ActiveSheet.Paste
            ActiveSheet.Name = "DB CP1"
            Sheets("DB CP1").Select
            ActiveWorkbook.ChangeLink Name:=modelo, NewName:=arquivo, Type:=xlExcelLinks
            
            Windows(modelo).Activate
            Sheets("DB CP2").Cells.Copy
                
            Windows(arquivo).Activate
            Sheets.Add After:=ActiveSheet
            ActiveSheet.Paste
            ActiveSheet.Name = "DB CP2"
            Sheets("DB CP2").Select
            ActiveWorkbook.ChangeLink Name:=modelo, NewName:=arquivo, Type:=xlExcelLinks
                    
            Windows(modelo).Activate
            Sheets("DB CP3").Cells.Copy
                
            Windows(arquivo).Activate
            Sheets.Add After:=ActiveSheet
            ActiveSheet.Paste
            ActiveSheet.Name = "DB CP3"
            Sheets("DB CP3").Select
            ActiveWorkbook.ChangeLink Name:=modelo, NewName:=arquivo, Type:=xlExcelLinks
                  
                 
            '_________________________________
            
            'APENAS PARA VALE MSA
            
            'Windows(modelo).Activate
            'Sheets("DB CP5").Cells.Copy
                
            'Windows(arquivo).Activate
            'Sheets.Add After:=ActiveSheet
            'ActiveSheet.Paste
            'ActiveSheet.Name = "DB CP5"
            'Sheets("DB CP5").Select
            'ActiveWorkbook.ChangeLink Name:=modelo, NewName:=arquivo, Type:=xlExcelLinks
            '_________________________________
                    
            ActiveWorkbook.Save
            ActiveWorkbook.Close
                    
            Windows(tabela_amostras).Activate
            ActiveCell.Offset(1, 2).Select
            
    
    Loop
    
Windows(modelo).Activate
Workbooks(modelo).Close savechanges:=False
MsgBox "Dados Gerados. Fa�a as verifica��es antes de executar a Parte 2!"

Application.ScreenUpdating = True

End Sub

Sub dados_triaxial_vale_parte2()

Application.ScreenUpdating = False

Dim nome As String
Dim tabela_amostras As String
Dim caminho As String
Dim contador As Integer
Dim arquivo As String
Dim amostra As String
Dim modelo As String
Dim caminho_modelo As String
Dim ensaio As String

    ensaio = "CIU"
    tabela_amostras = "1. Gerar Dados Brutos Vale.xlsm"
    Windows(tabela_amostras).Activate
   
    Windows(tabela_amostras).Activate
    Sheets(ensaio).Select
    ActiveSheet.Range("E4").Select
    
    
    Do While ActiveCell.Value <> ""
    
            nome = ActiveCell.Value
            ActiveCell.Offset(0, -1).Select
            arquivo = ActiveCell.Value
            ActiveCell.Offset(0, -1).Select
            caminho = ActiveCell.Value
            Workbooks.Open caminho, UpdateLinks:=0
            Windows(arquivo).Activate
            
            amostra = Sheets("CP1").Range("F10").Value
            
            'Sheets("DB CP1").Name = ensaio
            
            'Sheets("DB CP2").Activate
            'Range("A3:AK300").Select
            'Selection.Copy
            
            'Sheets(ensaio).Activate
            'Range("A10000").Select
            'Selection.End(xlUp).Select
            'ActiveCell.Offset(1, 0).Select
            'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            
            'Sheets("DB CP3").Activate
            'Range("A3:AK400").Select
            'Selection.Copy
            
            'Sheets(ensaio).Activate
            'Range("A10000").Select
            'Selection.End(xlUp).Select
            'ActiveCell.Offset(1, 0).Select
            'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            ':=False, Transpose:=False
            
            Sheets("DB CP4").Activate
            'Range("A3:AK400").Select
            'Selection.Copy
            
            'Sheets(ensaio).Activate
            'Range("A10000").Select
            'Selection.End(xlUp).Select
            'ActiveCell.Offset(1, 0).Select
            'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            ':=False, Transpose:=False
            
            'Sheets(ensaio).Activate
            'Range("A3:AK3").Select
            'Selection.Copy
            'Range("A3:AK400").Select
            'Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            
            '_____________________________________
            
            'APENAS PARA VALE MSA
            'Sheets("DB CP5").Activate
            'Range("A3:AK400").Select
            'Selection.Copy
            
            'Sheets(ensaio).Activate
            'Range("A10000").Select
            'Selection.End(xlUp).Select
            'ActiveCell.Offset(1, 0).Select
            'Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
            
            'Sheets(ensaio).Activate
            'Range("A3:AK300").Select
            'Selection.Copy
            'Range("A3:AK400").Select
            'Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            '_____________________________________
            
            Cells.Select
            Selection.Copy
        
            Set novo_arquivo = Application.Workbooks.Add
            Range("A1").Select
            ActiveSheet.Paste
            Cells.Select
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
                        
            ActiveSheet.Name = ensaio
                    
            ActiveWorkbook.SaveAs Filename:=nome & " - " & amostra
            ActiveWorkbook.Close
            
            Windows(arquivo).Activate
            ActiveWorkbook.Save
            ActiveWorkbook.Close
                    
            Windows(tabela_amostras).Activate
            ActiveCell.Offset(1, 2).Select
            
    
    Loop

Application.ScreenUpdating = True
MsgBox "Processo Conclu�do!"


End Sub
Sub permeabilidade_vale_DB()

Application.ScreenUpdating = False

Dim nome As String
Dim tabela_amostras As String
Dim caminho As String
Dim contador As Integer
Dim arquivo As String
Dim amostra As String
Dim modelo As String
Dim caminho_modelo As String

    modelo = "DB Permeabilidade Modelo.xls"
    tabela_amostras = "1. Gerar Dados Brutos Vale.xlsm"
    Windows(tabela_amostras).Activate
    
    'Abrir o modelo padr�o
    Sheets("Permeabilidade").Select
    ActiveSheet.Range("B3").Select
    caminho_modelo = ActiveCell.Value
    On Error Resume Next
        Workbooks.Open caminho_modelo, UpdateLinks:=0
    
    Windows(tabela_amostras).Activate
    Sheets("Permeabilidade").Select
    ActiveSheet.Range("E4").Select
    
    
    Do While ActiveCell.Value <> ""
    
            nome = ActiveCell.Value
            ActiveCell.Offset(0, -1).Select
            arquivo = ActiveCell.Value
            ActiveCell.Offset(0, -1).Select
            caminho = ActiveCell.Value
            Workbooks.Open caminho, UpdateLinks:=0
            
            Windows(arquivo).Activate
            
            'Sheets("Permeabilidade").Select
            'Sheets("Permeabilidade").Range("O27").Select
            'ActiveCell.FormulaR1C1 = "Gradiente hidr�ulico"
            'Sheets("Permeabilidade").Range("Q27").Select
            'ActiveCell.FormulaR1C1 = "=(R[2]C[-2]/10)/R[-8]C[-13]"
            'Sheets("Permeabilidade").Range("P27").Select
            'ActiveCell.FormulaR1C1 = "i"
            
            Windows(modelo).Activate
            Sheets("DB Permeabilidade").Cells.Copy
                
            Windows(arquivo).Activate
            
            Sheets("Permeabilidade").Select
            Range("Z19").Select
            ActiveCell.FormulaR1C1 = "=RC[-2]*R22C35"
            Range("Z19").Select
            Selection.AutoFill Destination:=Range("Z19:Z41")
            Range("Z19:Z41").Select
            Range("AA15").Select
            ActiveCell.FormulaR1C1 = "i"
            Range("AA19").Select
            ActiveCell.FormulaR1C1 = "=(RC[-5]/10)/RC[-23]"
            Range("AA19").Select
            ActiveCell.FormulaR1C1 = "=(RC[-5]/10)/R19C4"
            Range("AA19").Select
            Selection.AutoFill Destination:=Range("AA19:AA41")
            Range("AA19:AA41").Select
            
            Sheets.Add After:=ActiveSheet
            ActiveSheet.Paste
            ActiveSheet.Name = "DB Permeabilidade"
            Sheets("DB Permeabilidade").Select
            ActiveWorkbook.ChangeLink Name:=modelo, NewName:=arquivo, Type:=xlExcelLinks
            
            Range("A3").Select
            ActiveCell.FormulaR1C1 = "Usina da Mina de �guas Limpas"
            Range("A3").Select
            Selection.AutoFill Destination:=Range("A3:A26")
            Range("A3:A26").Select
            Columns("A:A").EntireColumn.AutoFit
            Sheets("Sheet1").Select
            Range("Q3").Select
            Sheets("Sheet1").Select
            Range("S3").Select
            Sheets("Sheet1").Select
            ActiveCell.FormulaR1C1 = "=Permeabilidade!R[16]C[7]"
            Range("S3").Select
            Selection.AutoFill Destination:=Range("S3:S26")
            Range("S3:S26").Select
            Range("S3").Select
            ActiveCell.FormulaR1C1 = "=Permeabilidade!R[15]C[7]"
            Range("S3").Select
            Selection.AutoFill Destination:=Range("S3:S26")
            Range("S3:S26").Select
            Selection.NumberFormat = "0.000000000000000000"
            Range("R10").Select
            
            
            amostra = Sheets("DB Permeabilidade").Range("B3").Value
            
            
            
            
            Sheets("DB Permeabilidade").Cells.Select
            Selection.Copy
            
            Set novo_arquivo = Application.Workbooks.Add
            Range("A1").Select
            ActiveSheet.Paste
            Cells.Select
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
                
            ActiveSheet.Name = "Permeabilidade"
            
            ActiveWorkbook.SaveAs Filename:=nome & " - " & amostra
            
            ActiveWorkbook.Close
            
            Windows(arquivo).Activate
            ActiveWorkbook.Save
            ActiveWorkbook.Close
                
            Windows(tabela_amostras).Activate
            ActiveCell.Offset(1, 2).Select
    
    Loop

Windows(modelo).Active
ActiveWorkbook.Close
MsgBox "Dados Gerados!"

Application.ScreenUpdating = True

End Sub


Sub adensamento_vale()

Application.ScreenUpdating = False

Dim nome As String
Dim tabela_amostras As String
Dim caminho As String
Dim contador As Integer
Dim arquivo As String
Dim amostra As String
Dim modelo As String
Dim caminho_modelo As String

    modelo = "DB Adensamento Modelo MSA.xls"
    tabela_amostras = "1. Gerar Dados Brutos Vale.xlsm"
    Windows(tabela_amostras).Activate
    
    'Abrir o modelo padr�o
    Sheets("Adensamento").Select
    ActiveSheet.Range("B3").Select
    caminho_modelo = ActiveCell.Value
    On Error Resume Next
        Workbooks.Open caminho_modelo, UpdateLinks:=0
    
    Windows(tabela_amostras).Activate
    Sheets("Adensamento").Select
    ActiveSheet.Range("E4").Select
    
    
    Do While ActiveCell.Value <> ""
    
            nome = ActiveCell.Value
            ActiveCell.Offset(0, -1).Select
            arquivo = ActiveCell.Value
            ActiveCell.Offset(0, -1).Select
            caminho = ActiveCell.Value
            Workbooks.Open caminho, UpdateLinks:=0
            
            Windows(modelo).Activate
            Sheets("Dados").Select
            Range("V65:V87").Select
            Selection.Copy
            
            Windows(arquivo).Activate
            Sheets("Dados").Select
            Range("V65").Select
            ActiveSheet.Paste
            ActiveWorkbook.ChangeLink Name:=modelo, NewName:=arquivo, Type:=xlExcelLinks
            
            Windows(modelo).Activate
            Sheets("Dados-2").Select
            Range("Y41:Y46").Select
            Selection.Copy
            
            Windows(arquivo).Activate
            Sheets("Dados-2").Select
            Range("Y41").Select
            ActiveSheet.Paste
            ActiveWorkbook.ChangeLink Name:=modelo, NewName:=arquivo, Type:=xlExcelLinks
            Range("Y42").Select
                        
            Windows(modelo).Activate
            Sheets("DB Adensamento").Cells.Copy
                
            Windows(arquivo).Activate
            Sheets.Add After:=ActiveSheet
            ActiveSheet.Paste
            ActiveSheet.Name = "DB Adensamento"
            Sheets("DB Adensamento").Select
            ActiveWorkbook.ChangeLink Name:=modelo, NewName:=arquivo, Type:=xlExcelLinks
            amostra = Sheets("Dados").Range("K9").Value
            
            'Windows(modelo).Activate
            'Sheets("Lista").Select
            'Cells.Copy
                
            'Windows(arquivo).Activate
            'Sheets.Add After:=ActiveSheet
            'ActiveSheet.Paste
            'ActiveSheet.Name = "Lista"
            'ActiveWorkbook.ChangeLink Name:=modelo, NewName:=arquivo, Type:=xlExcelLinks
            
            Sheets("DB Adensamento").Select
            
            Cells.Select
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
            
            Windows(arquivo).Activate
            
            Set novo_arquivo = Application.Workbooks.Add
            Range("A1").Select
            ActiveSheet.Paste
            Cells.Select
            Selection.Copy
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
                :=False, Transpose:=False
                
            ActiveSheet.Name = "Adensamento"
            
            ActiveWorkbook.SaveAs Filename:=nome & " - " & amostra
            
            ActiveWorkbook.Close
            
            Windows(arquivo).Activate
            
            ActiveWorkbook.Save
            ActiveWorkbook.Close
                
            Windows(tabela_amostras).Activate
            ActiveCell.Offset(1, 2).Select
    
    Loop

Windows(modelo).Activate
Workbooks(modelo).Close savechanges:=False

MsgBox "Dados Gerados!"

Application.ScreenUpdating = True

End Sub

Sub dados_brutos_Bender()
Dim nome As String
Dim furo As String
Dim amostra As String
Dim tabela_amostras As String
Dim caminho As String
Dim contador As Integer

    contador = 1
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
    
        Sheets("CP1").Select
        Cells.Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Application.CutCopyMode = False
        With Selection.Interior
            .Pattern = xlNone
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        With Selection.Font
            .ThemeColor = xlThemeColorLight1
            .TintAndShade = 0
        End With
        ActiveWindow.SmallScroll Down:=-18
        Range("A1:K208").Select
        Selection.ClearComments
        ActiveWindow.SmallScroll Down:=-3
        Sheets("GR�FICO 1").Select
        Cells.Select
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
        Sheets("CP2").Select
        Cells.Select
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
        Sheets("GR�FICO 2").Select
        Cells.Select
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
        Sheets("CP3").Select
        Cells.Select
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
        Sheets("GR�FICO 3").Select
        Cells.Select
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
        Range("A25:I25").Select
        Range("O19:P22").Select
        Selection.ClearContents
        Range("Q7").Select
        Sheets("CP3").Select
        Range("L23:M27").Select
        Selection.ClearContents
        Range("L56:P73").Select
        Selection.ClearContents
        Sheets("GR�FICO 2").Select
        Range("N13:Q22").Select
        Selection.ClearContents
        Sheets("CP2").Select
        Range("L11:O96").Select
        Selection.ClearContents
        Sheets("GR�FICO 1").Select
        Range("O18:P26").Select
        Selection.ClearContents
        Sheets("CP1").Select
        Range("L7:Q42").Select
        Range("L8").Activate
        Selection.ClearContents
        Range("M54:P73").Select
        Range("P54").Activate
        Selection.ClearContents
        Sheets("GR�FICO 1").Select
        Sheets("CP2").Select
        Range("A1:K208").Select
        Selection.ClearComments
        Range("G23").Select
        Sheets("CP3").Select
        Range("A1:K208").Select
        Selection.ClearComments
        Range("K13:K18").Select
        Sheets("Equipe").Select
        ActiveWindow.SelectedSheets.Delete
        Sheets("GR�FICO 3").Select
        Sheets("CP4").Visible = True
        Sheets("CP4").Select
        ActiveWindow.SelectedSheets.Delete
        Sheets("GR�FICO 3").Select
        Sheets("GR�FICO 4").Visible = True
        Sheets("GR�FICO 4").Select
        ActiveWindow.SelectedSheets.Delete
        Sheets("GR�FICO 3").Select
        Sheets("Tens�o-Deforma��o").Visible = True
        Sheets("Tens�o-Deforma��o").Select
        ActiveWindow.SelectedSheets.Delete
        Sheets("GR�FICO 3").Select
        Sheets("Instrumentos").Visible = True
        Sheets("Instrumentos").Select
        ActiveWindow.SelectedSheets.Delete
        Range("J21").Select
        Sheets(Array("CP1", "GR�FICO 1", "CP2", "GR�FICO 2", "CP3", "GR�FICO 3")).Select
        Sheets("CP1").Activate
        With ActiveWorkbook.Sheets("CP1").Tab
            .ColorIndex = xlNone
            .TintAndShade = 0
        End With
        With ActiveWorkbook.Sheets("GR�FICO 3").Tab
            .ColorIndex = xlNone
            .TintAndShade = 0
        End With
        With ActiveWorkbook.Sheets("CP3").Tab
            .ColorIndex = xlNone
            .TintAndShade = 0
        End With
        With ActiveWorkbook.Sheets("GR�FICO 2").Tab
            .ColorIndex = xlNone
            .TintAndShade = 0
        End With
        With ActiveWorkbook.Sheets("CP2").Tab
            .ColorIndex = xlNone
            .TintAndShade = 0
        End With
        With ActiveWorkbook.Sheets("GR�FICO 1").Tab
            .ColorIndex = xlNone
            .TintAndShade = 0
        End With
        Sheets("CP1").Select
        furo = Sheets("CP1").Range("F9").Value
        amostra = Sheets("CP1").Range("F10").Value
        
        ActiveWorkbook.SaveAs Filename:="Dados Brutos " & nome & " - " & furo & " " & amostra
        ActiveWorkbook.Close
                    
        Windows(tabela_amostras).Activate
        ActiveCell.Offset(1, 2).Select
        
    Loop

End Sub
