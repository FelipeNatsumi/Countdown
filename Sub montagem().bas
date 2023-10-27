Sub montagem()
    'Ordem das colunas - Dash
        'Bandeira - A
        'Código do fornecedor - B
        'Código - C
        'Drop - E

    Application.ScreenUpdating = False


    'Variáveis main sheet
    Dim wbfn As Workbook
    Dim wkfn As Worksheet, wkfndrop As Worksheet

    Set wbfn = Workbooks("FelipeNatsumi")
    Set wkfn = Workbooks("FelipeNatsumi").Sheets("FN")
    Set wkfndrop = Workbooks("FelipeNatsumi").Sheets("Drop")

    'Limpa a planilha main
    wkfn.Range("A:AZ").Clear


    'Data de hoje no formato dd_mm
    Dim today As Date
    today = Date
    ntoday = Format(Date, "dd_mm")


    'Abre a planilha da relação de pendências
    Dim pendDescricao As String
    pendDescricao = "C:\Cadastro\Pendentes e relatorios\pend_" & ntoday & ".csv"
    Workbooks.Open (pendDescricao)

    'Declara variáveis planilha de pendentes
    Dim wbpend As Workbook
    Dim wkpend As Worksheet

    Set wbpend = Workbooks("pend_" & ntoday)
    Set wkpend = Workbooks("pend_" & ntoday).Sheets(1)

    'Quantidade de linhas de pendentes
    ultimo = wkpend.Cells(Cells.Rows.Count, 1).End(xlUp).Row


    'Trasfere os dados do pendente para a main
    For X = 1 To ultimo
        wkfn.Cells(X, 1).Value = wkpend.Cells(X, 1).Value
        wkfn.Cells(X, 2).Value = wkpend.Cells(X, 2).Value
        wkfn.Cells(X, 3).Value = wkpend.Cells(X, 3).Value
        wkfn.Cells(X, 4).Value = wkpend.Cells(X, 5).Value
    Next X

    'Fecha a pasta pendentes sem salvar
    Application.DisplayAlerts = False
    wbpend.Close SaveChanges:=False
    Application.DisplayAlerts = True


    'Formata a planilha
    wkfn.Cells(1, 1).Value = "Bandeira"
    wkfn.Cells(1, 2).Value = "Código do fornecedor"
    wkfn.Cells(1, 3).Value = "SKU"
    wkfn.Cells(1, 4).Value = "Drop"
    wkfn.Rows("1").RowHeight = 15
    For X = 2 To ultimo
        If wkfn.Cells(X, 4).Value <> "null" Then
            wkfn.Cells(X, 4).Value = "Drop"
        ElseIf wkfn.Cells(X, 4).Value = "null" Then
            wkfn.Cells(X, 4).Value = "Não"
        End If
    Next X


    'Abre a planilha com os relatórios
    Dim relatorios As String
    relatorios = "C:\Cadastro\Pendentes e relatorios\relatorios_" & ntoday & ".xlsx"
    Workbooks.Open (relatorios)

    'Declara as variáves do relatório
    Dim wbrel As Workbook
    Dim wkrelaf As Worksheet, wkrelaw As Worksheet, wkrelmf As Worksheet

    Set wbrel = Workbooks("relatorios_" & ntoday)
    Set wkrelaf = Workbooks("relatorios_" & ntoday).Sheets("AF")
    Set wkrelaw = Workbooks("relatorios_" & ntoday).Sheets("AW")
    Set wkrelmf = Workbooks("relatorios_" & ntoday).Sheets("MF")

    'Traz as informações relevantes da planilha dos relatórios para a main
    'Verifica se a planilha dos relatórios está com as informações e se sim, deleta as duas primeiras linhas
    If wkrelaf.Cells(1, 1).Value <> "" Then
        relaf = sim
        wkrelaf.Range("A1:A2").EntireRow.Delete
    End If
    If wkrelaw.Cells(1, 1).Value <> "" Then
        relaw = sim
        wkrelaw.Range("A1:A2").EntireRow.Delete
    End If
    If wkrelmf.Cells(1, 1).Value <> "" Then
        relmf = sim
        wkrelmf.Range("A1:A2").EntireRow.Delete
    End If

    'Se estiver preenchida, monta o código 11
    If relaf = sim Then
        wkrelaf.Cells(1, 1).Value = "Codigo Pai"
        X = 2
        Do While wkrelaf.Cells(X, 1).Value <> ""
            wkrelaf.Cells(X, 1).Value = wkrelaf.Cells(X, 2).Value & "-" & Right(wkrelaf.Cells(X, 10).Value, 3)
            X = X + 1
        Loop
    End If
    If relaw = sim Then
        wkrelaw.Cells(1, 1).Value = "Codigo Pai"
        X = 2
        Do While wkrelaw.Cells(X, 1).Value <> ""
            wkrelaw.Cells(X, 1).Value = wkrelaw.Cells(X, 2).Value & "-" & Right(wkrelaw.Cells(X, 10).Value, 3)
            X = X + 1
        Loop
    End If
    If relmf = sim Then
        wkrelmf.Cells(1, 1).Value = "Codigo Pai"
        X = 2
        Do While wkrelmf.Cells(X, 1).Value <> ""
            wkrelmf.Cells(X, 1).Value = wkrelmf.Cells(X, 2).Value & "-" & Right(wkrelmf.Cells(X, 10).Value, 3)
            X = X + 1
        Loop
    End If

    'Traz as informações de Marca, Nome, Classe e linha caso a classe seja tênis. Pega também a origem e destino
    wkfn.Range("E1").Value = "Marca"
    wkfn.Range("F1").Value = "Nome"
    wkfn.Range("G1").Value = "Classe"
    wkfn.Range("H1").Value = "Linha"
    wkfn.Range("AC1").Value = "Origem"
    wkfn.Range("AD1").Value = "Destino"
    For X = 2 To ultimo
        bandeira = wkfn.Cells(X, 1).Value
        Select Case bandeira
            Case "AF"
                wkfn.Cells(X, 5).Value = Application.VLookup(wkfn.Cells(X, 3).Value, wkrelaf.Range("A:H"), 8, False)
                wkfn.Cells(X, 6).Value = Application.VLookup(wkfn.Cells(X, 3).Value, wkrelaf.Range("A:E"), 5, False)
                wkfn.Cells(X, 7).Value = Application.VLookup(wkfn.Cells(X, 3).Value, wkrelaf.Range("A:L"), 12, False)
                wkfn.Cells(X, 8).Value = Application.VLookup(wkfn.Cells(X, 3).Value, wkrelaf.Range("A:M"), 13, False)
                wkfn.Cells(X, 29).Value = Application.VLookup(wkfn.Cells(X, 3).Value, wkrelaf.Range("A:R"), 18, False)
                wkfn.Cells(X, 30).Value = Application.VLookup(wkfn.Cells(X, 3).Value, wkrelaf.Range("A:S"), 19, False)
            Case "AW"
                wkfn.Cells(X, 5).Value = Application.VLookup(wkfn.Cells(X, 3).Value, wkrelaw.Range("A:H"), 8, False)
                wkfn.Cells(X, 6).Value = Application.VLookup(wkfn.Cells(X, 3).Value, wkrelaw.Range("A:E"), 5, False)
                wkfn.Cells(X, 7).Value = Application.VLookup(wkfn.Cells(X, 3).Value, wkrelaw.Range("A:L"), 12, False)
                wkfn.Cells(X, 8).Value = Application.VLookup(wkfn.Cells(X, 3).Value, wkrelaw.Range("A:M"), 13, False)
                wkfn.Cells(X, 29).Value = Application.VLookup(wkfn.Cells(X, 3).Value, wkrelaw.Range("A:R"), 18, False)
                wkfn.Cells(X, 30).Value = Application.VLookup(wkfn.Cells(X, 3).Value, wkrelaw.Range("A:S"), 19, False)
            Case "MF"
                wkfn.Cells(X, 5).Value = Application.VLookup(wkfn.Cells(X, 3).Value, wkrelmf.Range("A:H"), 8, False)
                wkfn.Cells(X, 6).Value = Application.VLookup(wkfn.Cells(X, 3).Value, wkrelmf.Range("A:E"), 5, False)
                wkfn.Cells(X, 7).Value = Application.VLookup(wkfn.Cells(X, 3).Value, wkrelmf.Range("A:L"), 12, False)
                wkfn.Cells(X, 8).Value = Application.VLookup(wkfn.Cells(X, 3).Value, wkrelmf.Range("A:M"), 13, False)
                wkfn.Cells(X, 29).Value = Application.VLookup(wkfn.Cells(X, 3).Value, wkrelmf.Range("A:R"), 18, False)
                wkfn.Cells(X, 30).Value = Application.VLookup(wkfn.Cells(X, 3).Value, wkrelmf.Range("A:S"), 19, False)
        End Select
    Next X

    'Fecha planilha dos relatórios sem salvar
    Application.DisplayAlerts = False
    wbrel.Close SaveChanges:=False
    Application.DisplayAlerts = True

    'Se erro:
    'Verifica se tem algum código que não foi encontrado nos relatórios
    For X = 2 To ultimo
        If IsError(wkfn.Cells(X, 5).Value) Then
            wkfn.Cells(X, 9).Value = "erro"
        End If
    Next X


    'Remove as linhas com erro
    For X = 2 To ultimo
        If wkfn.Cells(X, 9).Value = "erro" Then
            wkfn.Cells(X, 9).EntireRow.Delete
            X = X - 1
        End If
    Next X

    'Nova contagem de linhas
    ultimo2 = wkfn.Cells(Cells.Rows.Count, 1).End(xlUp).Row

    'Arruma as classes de acordo com o nome do produto
    For X = 2 To ultimo2
        texto = wkfn.Cells(X, 6).Value
        espaço = InStr(texto, " ")
        classe = Left(texto, espaço - 1)
        wkfn.Cells(X, 7).Value = classe
    Next X

    'Remove as familias das classes <> tênis
    For X = 2 To ultimo2
        If wkfn.Cells(X, 7).Value <> "Tênis" Then
            wkfn.Cells(X, 8).Clear
        End If
    Next
    
    'Classifica por SKU, Bandeira, Linha e marca
    With wkfn.Sort
        .SortFields.Clear
        .SortFields.Add Key:=wkfn.Range("G:G"), SortOn:=xlSortOnValues, Order:=xlAscending
        .SortFields.Add Key:=wkfn.Range("H:H"), SortOn:=xlSortOnValues, Order:=xlAscending
        .SortFields.Add Key:=wkfn.Range("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending
        .SortFields.Add Key:=wkfn.Range("C:C"), SortOn:=xlSortOnValues, Order:=xlAscending
        .SetRange wkfn.UsedRange
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    'Variáveis descrição
    wkfn.Range("I1").Value = "Gênero"
    wkfn.Range("J1").Value = "Descrição Genérica"
    wkfn.Range("K1").Value = "Material"
    wkfn.Range("L1").Value = "Tecnologia"
    wkfn.Range("M1").Value = "Bolso"
    wkfn.Range("N1").Value = "Caimento"
    wkfn.Range("O1").Value = "Dimensões (EQT)"
    wkfn.Range("P1").Value = "Aba (Boné)"
    wkfn.Range("Q1").Value = "Ajuste (Boné)"
    wkfn.Range("AA1").Value = "Descrição genérica"
    wkfn.Range("AB1").Value = "Material genérico"

    'Opção adicional para incluir os drops com o nome do produto - podem ser incluidos na sheet possiveis drops
    ldrop = wkfndrop.Cells(Cells.Rows.Count, 1).End(xlUp).Row

    For X = 2 To ultimo2
        For Y = 2 To ldrop
            If InStr(1, wkfn.Cells(X, 6).Value, wkfndrop.Cells(Y, 1).Value, vbTextCompare) > 0 Then
                wkfn.Cells(X, 9).Value = "drop"
            End If
        Next Y
    Next X

    'Remove sky jordan 1, porque seria considerado como drop pelo filtro jordan 1
    For X = 2 To ultimo2
        For Y = 2 To ldrop
            If InStr(1, wkfn.Cells(X, 6).Value, wkfndrop.Cells(Y, 3).Value, vbTextCompare) > 0 Then
                wkfn.Cells(X, 9).Value = ""
            End If
        Next Y
    Next X

    'Se for drop, incluir na coluna de descrição generica como "drop"
    For X = 2 To ultimo2
        If wkfn.Cells(X, 4).Value = "Drop" Then
            wkfn.Cells(X, 9).Value = "drop"
        End If
    Next X

    'Põe o genero
    For X = 2 To ultimo2
        texto = wkfn.Cells(X, 6).Value
        espacoDireita = InStrRev(texto, " ") + 1
        genero = Right(texto, Len(texto) - espacoDireita + 1)
        wkfn.Cells(X, 9).Value = genero
    Next X
    
    'Ajusta os unissex
    For X = 2 To ultimo2
        If wkfn.Cells(X, 9).Value <> "Masculino" And wkfn.Cells(X, 9).Value <> "Masculina" And wkfn.Cells(X, 9).Value <> "Feminino" And wkfn.Cells(X, 9).Value <> "Feminina" And wkfn.Cells(X, 9).Value <> "Infantil" Then
            wkfn.Cells(X, 9).Value = "Unissex"
        End If
    Next X

    'Formata a planilha
    'Autoajuste das colunas
    wkfn.Columns.AutoFit

    'Deixa dark mode
    wkfn.Range("A1").CurrentRegion.Select
    With Selection
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(24, 43, 53)
        .Font.Bold = False
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(0, 0, 0)
        .RowHeight = 15
    End With
    wkfn.Range("A1:Q1").Select
    With Selection
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Interior.Color = RGB(0, 0, 0)
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = RGB(0, 0, 0)
    End With

    'Aplica a validação de dados para a descrição genérico
    wkfn.Range("J2:J" & ultimo2).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Sim, Não"

    'Aplica a validação de dados no restante, caso descrição generica = nao e caso necessite
    For x = 2 to ultimo2
        generico = wkfn.Cells(x, 10).value
        if generico = "Não" Then
              'tech
            'bolso
            wkfn.range(x, 13).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Frontal, Lateral"
            'caimento
            'aba
            wkfn.range(x, 16).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Curva, Reta"            
            'ajuste (boné) - Strapback - Snapback
            wkfn.range(x, 17).Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="Strapback, Snapback"
        end if
    Next x

    Application.ScreenUpdating = True
End Sub