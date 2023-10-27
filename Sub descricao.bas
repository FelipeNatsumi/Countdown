Sub descricao()

    Application.ScreenUpdating = False


    'Variáveis main sheet
    Dim wbfn As Workbook
    Dim wkfn As Worksheet, wkfnaf As Worksheet, wkfnaw As Worksheet, wkfnmf As Worksheet, wkfndrop As Worksheet

    Set wbfn = Workbooks("FelipeNatsumi")
    Set wkfn = Workbooks("FelipeNatsumi").Sheets("FN")
    Set wkfnaf = Workbooks("FelipeNatsumi").Sheets("AF")
    Set wkfnaw = Workbooks("FelipeNatsumi").Sheets("AW")
    Set wkfnmf = Workbooks("FelipeNatsumi").Sheets("MF")
    Set wkfndrop = Workbooks("FelipeNatsumi").Sheets("Drop")

    'Quantidade de linhas de pendentes
    ultimo = wkfn.Cells(Cells.Rows.Count, 1).End(xlUp).Row

    'Puxa as descrições genéricas
    For X = 2 To ultimo
        If wkfn.Cells(X, 9).Value = "drop" Then
            wkfn.Cells(X, 27).Value = "***" & wkfn.Cells(X, 6).Value & "***"
        Else
            bandeira = wkfn.Cells(X, 1).Value
            classe = wkfn.Cells(X, 7).Value
            Select Case bandeira
                Case "AF"
                    Select Case classe
                        Case "Tênis"
                            wkfn.Cells(X, 27).Value = Application.VLookup(wkfn.Cells(X, 8).Value, wkfnaf.Range("D:E"), 2, False)
                            wkfn.Cells(X, 28).Value = Application.VLookup(wkfn.Cells(X, 8).Value, wkfnaf.Range("D:F"), 3, False)
                        Case Is <> "Tênis"
                            wkfn.Cells(X, 27).Value = Application.VLookup(wkfn.Cells(X, 7).Value, wkfnaf.Range("A:B"), 2, False)
                            wkfn.Cells(X, 28).Value = Application.VLookup(wkfn.Cells(X, 7).Value, wkfnaf.Range("A:C"), 3, False)
                    End Select
                Case "AW"
                    Select Case classe
                        Case "Tênis"
                            wkfn.Cells(X, 27).Value = Application.VLookup(wkfn.Cells(X, 8).Value, wkfnaw.Range("D:E"), 2, False)
                            wkfn.Cells(X, 28).Value = Application.VLookup(wkfn.Cells(X, 8).Value, wkfnaw.Range("D:F"), 3, False)
                        Case Is <> "Tênis"
                            wkfn.Cells(X, 27).Value = Application.VLookup(wkfn.Cells(X, 7).Value, wkfnaw.Range("A:B"), 2, False)
                            wkfn.Cells(X, 28).Value = Application.VLookup(wkfn.Cells(X, 7).Value, wkfnaw.Range("A:C"), 3, False)
                    End Select
                Case "MF"
                    Select Case classe
                        Case "Tênis"
                            wkfn.Cells(X, 27).Value = Application.VLookup(wkfn.Cells(X, 8).Value, wkfnmf.Range("D:E"), 2, False)
                            wkfn.Cells(X, 28).Value = Application.VLookup(wkfn.Cells(X, 8).Value, wkfnmf.Range("D:F"), 3, False)
                        Case Is <> "Tênis"
                            wkfn.Cells(X, 27).Value = Application.VLookup(wkfn.Cells(X, 7).Value, wkfnmf.Range("A:B"), 2, False)
                            wkfn.Cells(X, 28).Value = Application.VLookup(wkfn.Cells(X, 7).Value, wkfnmf.Range("A:C"), 3, False)
                    End Select
            End Select
        End If
    Next X

    'Verifica se alguma descrição não foi encontrada e aponta qual e monta a tabela de não encontrados
    Y = 1
    For X = 2 To ultimo
        Status = wkfn.Cells(X, 27).Value
        classe = wkfn.Cells(x, 7).value
        If IsError(Status) Then
            wkfn.Cells(Y, 35).Value = X
            wkfn.Cells(Y, 36).Value = wkfn.Cells(X, 1).Value
            Select case classe
                case "Tênis"
                    wkfn.Cells(Y, 37).Value = wkfn.Cells(X, 8).Value
                case is <> "Tênis"
                    wkfn.Cells(Y, 37).Value = wkfn.Cells(X, 7).Value
            end Select
            Y = Y + 1
        End If
    Next X

    'Mostra as classes não encontrados e exclui da relação
    if wkfn.range("AI1").value <> "" Then
        mainNd.Show
        wkfn.Range("AI1").CurrentRegion.Delete
        For X = 2 To ultimo
            Status = wkfn.Cells(X, 27).Value
            If IsError(Status) Then
                wkfn.Cells(X, 27).EntireRow.Delete
                X = X - 1
            End If
        Next X
    end if

    'Arrumar as descrições - Nome e Gênero(se unissex, deixa em branco)
    For X = 2 To ultimo
        genero = wkfn.Cells(X, 9).Value
        wkfn.Cells(X, 27).Value = Replace(wkfn.Cells(X, 27).Value, "#NOME#", wkfn.Cells(X, 6).Value)
        If genero = "Unissex" Then
            wkfn.Cells(X, 27).Value = Replace(wkfn.Cells(X, 27).Value, "#GENERO# ", "")
        Else
            wkfn.Cells(X, 27).Value = Replace(wkfn.Cells(X, 27).Value, "#GENERO#", LCase(genero))
        End If
    Next X

    'Passa as descrições genéricas para o lado, e arrumas as não genericas
    wkfn.range("AF1").value = "Descrição final"
    For x = 2 to ultimo
        generico = Ucase(wkfn.Cells(x, 10).value)
        material = ucase(wkfn.Cells(x, 11).value)
        tecnologia = ucase(wkfn.Cells(x, 12).value)
        bolso = ucase(wkfn.Cells(x, 13).value)
        caimento = ucase(wkfn.Cells(x, 14).value)
        dimensoes = ucase(wkfn.Cells(x, 15).value)
        aba = ucase(wkfn.Cells(x, 16).value)
        ajuste = ucase(wkfn.Cells(x, 17).value)
        select case generico
            case "SIM"
                wkfn.Cells(x, 32).value = Replace(wkfn.Cells(x, 27).value, "#MATERIAL#", wkfn.Cells(x, 28).value)
                wkfn.Cells(X, 32).Replace What:="#*#", Replacement:=""
            case "NÃO"
                Select case material
                    case ""
                        wkfn.Cells(x, 32).value = Replace(wkfn.Cells(x, 27).value, "#MATERIAL#", wkfn.Cells(x, 28).value)
                    case <> ""
                        wkfn.Cells(x, 32).value = Replace(wkfn.Cells(x, 27).value, "#MATERIAL#", wkfn.Cells(x, 11).value)
                end select
                Select case tecnologia
                    case ""
                    case <> ""
                end select
                Select case bolso
                    case ""
                    case <> ""
                end select
                Select case caimento
                    case ""
                    case <> ""
                end select
                Select case dimensoes
                    case ""
                    case <> ""
                end select
                Select case aba
                    case ""
                    case <> ""
                end select
                Select case ajuste
                    case ""
                    case <> ""
                end select
        end select
    Next x


End Sub
