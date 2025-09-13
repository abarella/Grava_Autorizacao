Imports System
Imports System.DirectoryServices
Imports System.IO
Imports System.IO.Directory
Imports System.IO.File
Imports System.Xml
Imports System.Diagnostics


Public Class frmGravarAutorizacao

    Private ReadOnly logFilePath As String = "C:\Unimake\Uninfe4\Backup\log.txt" ' Defina o caminho do log
    Private ReadOnly fileProcessingLock As New Object() ' Lock para sincronização de processamento de arquivos

    Private Sub EscreverLogErro(ex As Exception)
        Try
            Dim directoryPath As String = Path.GetDirectoryName(logFilePath)

            If Not Directory.Exists(directoryPath) Then
                Directory.CreateDirectory(directoryPath)
            End If

            Using writer As New StreamWriter(logFilePath, True)
                writer.WriteLine($"Data/Hora: {DateTime.Now}")
                writer.WriteLine($"Mensagem: {ex.Message}")
                writer.WriteLine($"Stack Trace: {ex.StackTrace}")
                writer.WriteLine(New String("-"c, 50))
            End Using
        Catch logEx As Exception
            Console.WriteLine("Erro ao escrever no log: " & logEx.Message)
        End Try
    End Sub

    ' Aguarda até que o arquivo remoto seja removido (processado) ou até o timeout (ms)
    Private Function WaitForRemoteFileProcessing(remoteFilePath As String, timeoutMilliseconds As Integer) As Boolean
        Try
            Dim sw As New Stopwatch()
            sw.Start()

            While sw.ElapsedMilliseconds < timeoutMilliseconds
                If Not System.IO.File.Exists(remoteFilePath) Then
                    Return True
                End If

                Threading.Thread.Sleep(500)
            End While

            Return Not System.IO.File.Exists(remoteFilePath)
        Catch ex As Exception
            EscreverLogErro(ex)
            Return False
        End Try
    End Function

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        Try

            'Call CarregaConfiguracao()
            'Threading.Thread.Sleep(800)

            Call GravarAutorizacaoDIRF()
            Threading.Thread.Sleep(8000)

            Call GravarXMLGeral()
            Threading.Thread.Sleep(800)

            'Call GravarAutorizacaoDIRF()
            'Threading.Thread.Sleep(800)

            Call GeraTXTServicos()
            Threading.Thread.Sleep(800)

            Call GravarXMLEventoCancela()
            Threading.Thread.Sleep(25000)

            Call EventoCancelarNotaGeral()
            Threading.Thread.Sleep(10000)

            'Call GravarAutorizacaoDIRF()
            'Threading.Thread.Sleep(800)

            Call GravarXMLGeralSituacao()
            Threading.Thread.Sleep(8000)

            Call GravarAutorizacaoSituacao()
            Threading.Thread.Sleep(8000)

            Call Erro_NFe_Email()
            Threading.Thread.Sleep(800)

            Call GravarXMLCartaCorrecao()
            Threading.Thread.Sleep(800)

            Call GravarAutorizacaoCartaCorrecao()
            Threading.Thread.Sleep(800)

        Catch ex As Exception

            EscreverLogErro(ex)

        End Try


    End Sub

    Private Sub CarregaConfiguracao()
        Dim Formato As System.Text.UTF8Encoding
        Formato = System.Text.UTF8Encoding.UTF8

        Try
            For Each Arquivo As String In Directory.GetFiles(".", "CONFIGURACAO.xml")
                Dim ArquivoXml As StreamReader = New StreamReader(Arquivo)
                Dim Leitura As XmlTextReader = New XmlTextReader(ArquivoXml)
                Dim no As String

                no = ""

                Do While (Leitura.Read())
                    Select Case Leitura.NodeType
                        Case XmlNodeType.Element
                            no = Leitura.Name
                        Case XmlNodeType.Text  'Exibir o início do elemento.
                            Select Case no
                                Case "Envio"
                                    PastaEnvio = Leitura.Value
                                Case "Retorno"
                                    PastaRetorno = Leitura.Value
                                Case "SemProtocolo"
                                    PastaSemProtocolo = Leitura.Value
                                Case "Gravados"
                                    PastaGravados = Leitura.Value
                                Case "Importacao"
                                    PastaImportacao = Leitura.Value
                                Case "Pedro"
                                    PastaPedro = Leitura.Value
                                Case "Ctr"
                                    PastaCtr = Leitura.Value
                                Case "Dirf"
                                    PastaDirf = Leitura.Value
                                Case "Servicos"
                                    PastaServico = Leitura.Value
                                Case "Erro"
                                    PastaErro = Leitura.Value
                                Case "Denegado"
                                    PastaDenegado = Leitura.Value
                            End Select
                    End Select
                Loop


                Threading.Thread.Sleep(2000)
                Leitura.Close()
                ArquivoXml.Close()
            Next

        Catch ex As Exception

            EscreverLogErro(ex)

        End Try

    End Sub

    Private Sub GravarAutorizacaoDIRF()
        Try
            ' Usar lock para garantir processamento sequencial
            SyncLock fileProcessingLock
                For Each Arquivo As String In Directory.GetFiles("C:\Unimake\Uninfe4\Retorno\", "*-pro-rec.xml")
                    ProcessarArquivoAutorizacao(Arquivo)
                Next
            End SyncLock
        Catch ex As Exception
            EscreverLogErro(ex)
        End Try
    End Sub

    Private Sub ProcessarArquivoAutorizacao(Arquivo As String)
        ' Variáveis isoladas para cada arquivo - evita sobreposição de conteúdo
        Dim executor As New GRAVA_AUTORIZACAO.AcessoDados.Executar
        Dim Tabela As DataSet
        Dim Tabela2 As DataSet
        Dim recibo_numero As String
        Dim protocolo_autorizacao As String
        Dim data_retorno As Date
        Dim envio_data As Date
        Dim data_processamento As Date
        Dim envio_lote As String
        Dim recibo_data As Date
        Dim status_nf As Integer
        Dim chave_acesso As String
        Dim AreaEmissao As String
        Dim Texto As DataRow
        Dim Formato As System.Text.UTF8Encoding
        Formato = System.Text.UTF8Encoding.UTF8

        Try
            Dim ArquivoXml As StreamReader = New StreamReader(Arquivo)
            Dim Leitura As XmlTextReader = New XmlTextReader(ArquivoXml)
            Dim no As String

            no = ""

            recibo_numero = ""
            protocolo_autorizacao = ""
            envio_lote = ""
            status_nf = 0
            chave_acesso = ""

            Do While (Leitura.Read())
                Select Case Leitura.NodeType
                    Case XmlNodeType.Element
                        no = Leitura.Name
                    Case XmlNodeType.Text  'Exibir o início do elemento.
                        If no = "nRec" Then
                            recibo_numero = Leitura.Value
                        End If

                        If no = "nProt" Then
                            protocolo_autorizacao = Leitura.Value
                        End If

                        If no = "dhRecbto" Then
                            data_retorno = Leitura.Value
                            envio_data = Leitura.Value
                            recibo_data = Leitura.Value
                            data_processamento = Leitura.Value
                        End If

                        If no = "chNFe" Then
                            envio_lote = Leitura.Value
                            chave_acesso = Leitura.Value
                        End If

                        If no = "cStat" Then
                            status_nf = Leitura.Value
                        End If
                End Select
            Loop

            If protocolo_autorizacao = "" Then
                If status_nf = "204" Then
                    FileCopy(Arquivo, "C:\Unimake\Uninfe4\SemProtocolo\" & recibo_numero & "-dupl-rec.xml")
                Else
                    FileCopy(Arquivo, "C:\Unimake\Uninfe4\SemProtocolo\" & recibo_numero & "-pro-rec.xml")
                    System.IO.File.AppendAllText(logFilePath, "Nota " & chave_acesso & " com rejeição, verificar na pasta SemProtocolo" & Environment.NewLine)
                End If
                'FileCopy(Arquivo, PastaSemProtocolo & chave_acesso & "-pro-rec.xml")
                Leitura.Close()
                ArquivoXml.Close()

                'Grava informação para envio de SMS em caso de rejeição de NFE
                'executor.SQLComando = "crsa.PSMS_NFE_ERRO_GRAVA"
                'executor.addParametros("DESCRICAO", chave_acesso, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.VarChar, 130)
                'executor.ExecutarOperacao(AcessoDados.Executar.TipoRetorno.eDataSet, "SGCR")
                'executor.Dispose()
            Else
                If status_nf = "301" Or status_nf = "302" Or status_nf = "303" Then  ' NOTA FISCAL DENEGADA
                    executor.Dispose()
                    executor.SQLComando = "dbo.sp_NFS_DENEGADAS"
                    executor.addParametros("chave_acesso", chave_acesso, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.VarChar, 50)
                    executor.ExecutarOperacao(AcessoDados.Executar.TipoRetorno.eDataSet, "IPENFAT")

                    executor.Dispose()
                    Leitura.Close()
                    ArquivoXml.Close()
                    Threading.Thread.Sleep(200)
                    FileCopy(Arquivo, "C:\Unimake\Uninfe4\Gravados\" & chave_acesso & "-pro-rec.xml")
                    'FileCopy(Arquivo, "D:\Unimake\UniNFe4\Gravados\" & chave_acesso & "-pro-rec.xml")
                Else
                    executor.Dispose()
                    'Gravar Autorização na tabela TNFe_IDENTIFICACAO
                    executor.SQLComando = "dbo.PNFE_GRAVA_AUTORIZACAO"
                    executor.addParametros("recibo_numero", recibo_numero, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.VarChar, 20)
                    executor.addParametros("protocolo_autorizacao", protocolo_autorizacao, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.VarChar, 20)
                    executor.addParametros("data_retorno", data_retorno, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.DateTime, 19)
                    executor.addParametros("envio_data", envio_data, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.DateTime, 19)
                    executor.addParametros("data_processamento", data_processamento, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.DateTime, 19)
                    executor.addParametros("envio_lote", envio_lote, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.VarChar, 50)
                    executor.addParametros("recibo_data", recibo_data, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.DateTime, 19)
                    executor.addParametros("status_nf", status_nf, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.Int, 4)
                    executor.addParametros("chave_acesso", chave_acesso, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.VarChar, 50)
                    executor.ExecutarOperacao(AcessoDados.Executar.TipoRetorno.eDataSet, "VendasInternet")

                    Tabela2 = executor.RetornoDataSet
                    If Tabela2.Tables(0).Rows(0).Item("emissor").ToString() = "" Then
                        AreaEmissao = "CR1"
                    Else
                        AreaEmissao = Trim(Tabela2.Tables(0).Rows(0).Item("emissor").ToString())
                    End If

                    'conferir se já foi emitida
                    If CInt(Tabela2.Tables(0).Rows(0).Item("emitido").ToString()) = 1 Then
                        executor.Dispose()
                        Leitura.Close()
                        ArquivoXml.Close()
                        Threading.Thread.Sleep(200)
                        FileCopy(Arquivo, "C:\Unimake\Uninfe4\Gravados\" & chave_acesso & "-pro-rec.xml")
                    Else
                        executor.Dispose()

                        Leitura.Close()
                        ArquivoXml.Close()

                        Dim i As Int64
                        Dim Documento

                        FileCopy(Arquivo, "C:\Unimake\Uninfe4\Gravados\" & chave_acesso & "-pro-rec.xml")

                        Threading.Thread.Sleep(2000)

                        Select Case AreaEmissao
                            Case "IMP"
                                'Imprimir arquivo com a autorização
                                executor.SQLComando = "dbo.P0110_GERA_NOTA_NFE_GERAL"
                                executor.addParametros("recibo_numero", protocolo_autorizacao, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.VarChar, 20)
                            Case "ALM"
                                executor.SQLComando = "dbo.P0110_GERA_NOTA_NFE_GERAL"
                                executor.addParametros("recibo_numero", protocolo_autorizacao, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.VarChar, 20)
                            Case "CTR"
                                executor.SQLComando = "dbo.P0110_GERA_NOTA_NFE_CTR"
                                executor.addParametros("recibo_numero", protocolo_autorizacao, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.VarChar, 20)
                            Case "CR2"
                                executor.SQLComando = "dbo.P0110_GERA_NOTA_NFE_GERAL"
                                executor.addParametros("recibo_numero", protocolo_autorizacao, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.VarChar, 20)
                            Case Else
                                executor.SQLComando = "dbo.P0110_GERA_NOTA_NFE_DIRF"
                                executor.addParametros("recibo_numero", protocolo_autorizacao, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.VarChar, 20)
                        End Select

                        executor.ExecutarOperacao(AcessoDados.Executar.TipoRetorno.eDataSet, "VendasPelicano")
                        Tabela = executor.RetornoDataSet

                        Dim TRANS As String

                        TRANS = ""
                        Documento = ""

                        i = 0
                        For Each Texto In Tabela.Tables(0).Rows
                            Documento = Documento & Tabela.Tables(0).Rows(i).Item(1) & vbCrLf
                            i = i + 1
                        Next

                        Dim contador As Integer = 0

                        If Documento <> "" Then
                            ' Incrementa o contador
                            contador = contador + 1

                            ' Escreve no log a execução
                            Dim logMessage As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " - Execução " & contador & " - Criando arquivo: " & chave_acesso & ".txt - Emissor: " & AreaEmissao

                            ' Grava a informação no log
                            'System.IO.File.AppendAllText(logFilePath, logMessage & Environment.NewLine)

                            Dim ArquivoTxt As System.IO.File
                            Select Case AreaEmissao
                                Case "IMP"
                                    ArquivoTxt.WriteAllText("C:\Unimake\Uninfe4\Importacao\" & chave_acesso & ".txt", Documento, Formato)
                                    Threading.Thread.Sleep(1000)
                                    'FileCopy(PastaImportacao & chave_acesso & ".txt", "\\10.0.10.79\IMPORTACAO")
                                    System.IO.File.Copy("C:\Unimake\Uninfe4\Importacao\" & chave_acesso & ".txt", "\\10.0.10.79\IMPORTACAO")
                                Case "ALM"
                                    ArquivoTxt.WriteAllText("C:\Unimake\Uninfe4\Pedro\" & chave_acesso & ".txt", Documento, Formato)
                                    Threading.Thread.Sleep(1000)
                                    System.IO.File.Copy("C:\Unimake\Uninfe4\Pedro\" & chave_acesso & ".txt", "\\10.0.10.79\PEDRO")
                                Case "CTR"
                                    ArquivoTxt.WriteAllText("C:\Unimake\Uninfe4\CTR\" & chave_acesso & ".txt", Documento, Formato)
                                    Threading.Thread.Sleep(1000)
                                    System.IO.File.Copy("C:\Unimake\Uninfe4\CTR\" & chave_acesso & ".txt", "\\10.0.10.79\CTR")
                                Case Else
                                    TRANS = Trim(Tabela.Tables(1).Rows(0).Item(0).ToString())

                                    If Not File.Exists("C:\Unimake\Uninfe4\Dirf\" & chave_acesso & ".txt") Then
                                        ArquivoTxt.WriteAllText("C:\Unimake\Uninfe4\Dirf\" & chave_acesso & ".txt", Documento, Formato)
                                    End If

                                    Threading.Thread.Sleep(1000)
                                    If TRANS = "BND" Then
                                        System.IO.File.Copy("C:\Unimake\Uninfe4\Dirf\" & chave_acesso & ".txt", "\\10.0.10.79\BND")
                                        System.IO.File.AppendAllText(logFilePath, logMessage & " Transp.: " & TRANS & Environment.NewLine)
                                    ElseIf TRANS = "REM" Then
                                        System.IO.File.Copy("C:\Unimake\Uninfe4\Dirf\" & chave_acesso & ".txt", "\\10.0.10.79\REM")
                                        System.IO.File.AppendAllText(logFilePath, logMessage & " Transp.: " & TRANS & Environment.NewLine)
                                    ElseIf TRANS = "HVL" Then
                                        System.IO.File.Copy("C:\Unimake\Uninfe4\Dirf\" & chave_acesso & ".txt", "\\10.0.10.79\HVL")
                                        System.IO.File.AppendAllText(logFilePath, logMessage & " Transp.: " & TRANS & Environment.NewLine)
                                    Else
                                        System.IO.File.Copy("C:\Unimake\Uninfe4\Dirf\" & chave_acesso & ".txt", "\\10.0.10.79\Entrada")
                                        System.IO.File.AppendAllText(logFilePath, logMessage & " Transp.: " & TRANS & Environment.NewLine)
                                    End If
                            End Select
                        End If
                        executor.Dispose()
                    End If
                End If 'fim de conferir se já foi emitida
            End If 'fim do If protocolo_autorizacao

            Threading.Thread.Sleep(2000)
            File.Delete(Arquivo)
            Threading.Thread.Sleep(2000)

        Catch ex As Exception
            EscreverLogErro(ex)
        End Try
    End Sub




    Private Sub GravarXMLGeral()
        Dim executor As New GRAVA_AUTORIZACAO.AcessoDados.Executar
        Dim Tabela As DataSet
        Dim Tabela2 As DataSet
        Dim Texto2 As DataRow
        Dim Formato As System.Text.UTF8Encoding
        Formato = System.Text.UTF8Encoding.UTF8
        Dim i As Integer
        Dim Documento As String

        Try

            executor.Dispose()

            executor.SQLComando = "dbo.P0110_LISTA_NFE_ENVIAR"
            executor.ExecutarOperacao(AcessoDados.Executar.TipoRetorno.eDataSet, "VendasPelicano")

            Tabela2 = executor.RetornoDataSet

            i = 0
            For Each Texto2 In Tabela2.Tables(0).Rows
                executor.Dispose()
                Dim Emissor As String
                Emissor = ""

                'Gera arquivo da Nota Fiscal Eletrônica (XML)
                executor.SQLComando = "dbo.PNFe_GERAXML_NF"
                executor.addParametros("notafis_oid", CInt(Tabela2.Tables(0).Rows(i).Item(0).ToString()), AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.Int, 4)
                executor.addParametros("nf", CInt(Tabela2.Tables(0).Rows(i).Item(1).ToString()), AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.Int, 4)
                executor.ExecutarOperacao(AcessoDados.Executar.TipoRetorno.eDataSet, "VendasInternet")

                Tabela = executor.RetornoDataSet

                Documento = ""

                Documento = Documento & Tabela.Tables(0).Rows(0).Item(0).ToString()

                If Documento <> "" Then
                    Dim ArquivoXml As System.IO.File
                    ArquivoXml.WriteAllText("C:\Unimake\Uninfe4\Envio\" & Tabela.Tables(1).Rows(0).Item(0).ToString() & "-nfe.xml", Documento, Formato)
                End If
                i = i + 1
            Next

            executor.Dispose()
        Catch ex As Exception

            EscreverLogErro(ex)

        End Try


    End Sub
    Private Sub GravarXMLCancela()
        Dim executor As New GRAVA_AUTORIZACAO.AcessoDados.Executar
        Dim Tabela As DataSet
        Dim Tabela2 As DataSet
        Dim Texto2 As DataRow
        Dim Formato As System.Text.UTF8Encoding
        Formato = System.Text.UTF8Encoding.UTF8
        'Dim Formato As System.Text.Encoding
        'Formato = System.Text.Encoding.Default
        Dim i As Integer
        Dim Documento As String
        Dim nfe_chave As String
        Try

            executor.Dispose()

            executor.SQLComando = "dbo.P0110_LISTA_NFE_CANCELA"
            executor.ExecutarOperacao(AcessoDados.Executar.TipoRetorno.eDataSet, "VendasPelicano")

            Tabela2 = executor.RetornoDataSet

            i = 0
            For Each Texto2 In Tabela2.Tables(0).Rows
                executor.Dispose()
                Dim Emissor As String
                Emissor = ""

                'Gera arquivo da Nota Fiscal Eletrônica (XML)
                executor.SQLComando = "dbo.PNFe_CANCELAMENTO"
                executor.addParametros("notafis_oid", CInt(Tabela2.Tables(0).Rows(i).Item(0).ToString()), AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.Int, 4)
                executor.addParametros("nf", CInt(Tabela2.Tables(0).Rows(i).Item(1).ToString()), AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.Int, 4)
                executor.ExecutarOperacao(AcessoDados.Executar.TipoRetorno.eDataSet, "VendasInternet")

                Tabela = executor.RetornoDataSet

                Documento = ""

                Documento = Documento & Tabela.Tables(0).Rows(0).Item(0).ToString()

                nfe_chave = ""
                nfe_chave = Tabela2.Tables(0).Rows(i).Item(2).ToString()


                If Documento <> "" Then
                    Dim ArquivoXml As System.IO.File
                    ArquivoXml.WriteAllText("C:\Unimake\Uninfe4\Envio\" & nfe_chave & "-ped-can.xml", Documento, Formato)
                End If
                i = i + 1
            Next

            executor.Dispose()
        Catch ex As Exception

            EscreverLogErro(ex)

        End Try


    End Sub

    Private Sub GravarXMLEventoCancela()
        Dim executor As New GRAVA_AUTORIZACAO.AcessoDados.Executar
        Dim Tabela As DataSet
        Dim Tabela2 As DataSet
        Dim Texto2 As DataRow
        Dim Formato As System.Text.UTF8Encoding
        Formato = System.Text.UTF8Encoding.UTF8
        'Dim Formato As System.Text.Encoding
        'Formato = System.Text.Encoding.Default
        Dim i As Integer
        Dim Documento As String
        Dim nfe_chave As String
        Try

            executor.Dispose()

            executor.SQLComando = "dbo.P0110_LISTA_NFE_CANCELA"
            executor.ExecutarOperacao(AcessoDados.Executar.TipoRetorno.eDataSet, "VendasPelicano")

            Tabela2 = executor.RetornoDataSet

            i = 0
            For Each Texto2 In Tabela2.Tables(0).Rows
                executor.Dispose()
                Dim Emissor As String
                Emissor = ""

                'Gera arquivo da Nota Fiscal Eletrônica (XML)
                executor.SQLComando = "dbo.PNFe_EventoCancelamento"
                executor.addParametros("notafis_oid", CInt(Tabela2.Tables(0).Rows(i).Item(0).ToString()), AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.Int, 4)
                executor.addParametros("nf", CInt(Tabela2.Tables(0).Rows(i).Item(1).ToString()), AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.Int, 4)
                executor.ExecutarOperacao(AcessoDados.Executar.TipoRetorno.eDataSet, "VendasInternet")

                Tabela = executor.RetornoDataSet

                Documento = ""

                Documento = Documento & Tabela.Tables(0).Rows(0).Item(0).ToString()

                nfe_chave = ""
                nfe_chave = Tabela2.Tables(0).Rows(i).Item(2).ToString()

                If Documento <> "" Then
                    Dim ArquivoXml As System.IO.File
                    ArquivoXml.WriteAllText("C:\Unimake\Uninfe4\Envio\" & nfe_chave & "-env-canc.xml", Documento, Formato)
                End If
                i = i + 1
            Next

            executor.Dispose()
        Catch ex As Exception

            EscreverLogErro(ex)

        End Try


    End Sub

    Private Sub GeraTXTServicos()
        Dim executor As New GRAVA_AUTORIZACAO.AcessoDados.Executar
        Dim Tabela As DataSet
        Dim Tabela2 As DataSet
        Dim Texto2 As DataRow
        Dim Texto As DataRow
        Dim Formato As System.Text.UTF8Encoding
        Formato = System.Text.UTF8Encoding.UTF8
        Dim i As Integer
        Dim W As Integer
        Dim Documento As String

        Try

            executor.Dispose()

            executor.SQLComando = "dbo.P0110_LISTA_NFE_SERVICOS"
            executor.ExecutarOperacao(AcessoDados.Executar.TipoRetorno.eDataSet, "VendasPelicano")

            Tabela2 = executor.RetornoDataSet

            i = 0
            For Each Texto2 In Tabela2.Tables(0).Rows
                executor.Dispose()
                Dim NumeroNota As Integer
                NumeroNota = CInt(Tabela2.Tables(0).Rows(i).Item(0).ToString())


                'Gera txt Servicos para ser enviado para almoxarifado
                executor.SQLComando = "dbo.P0110_GERA_NOTA_NFE_SERVICOS"
                executor.addParametros("no_nota", NumeroNota, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.Int, 4)
                executor.ExecutarOperacao(AcessoDados.Executar.TipoRetorno.eDataSet, "VendasPelicano")

                Tabela = executor.RetornoDataSet

                Documento = ""

                W = 0
                For Each Texto In Tabela.Tables(0).Rows
                    Documento = Documento & Tabela.Tables(0).Rows(W).Item(1).ToString() & vbCrLf
                    W = W + 1
                Next

                If Documento <> "" Then
                    ' Escreve o arquivo localmente
                    Dim localPath As String = "C:\Unimake\Uninfe4\Servicos\" & NumeroNota & ".txt"
                    System.IO.File.WriteAllText(localPath, Documento, Formato)

                    Threading.Thread.Sleep(200)

                    ' Copia para o compartilhamento remoto usando um nome temporário e depois renomeia
                    Dim remoteDir As String = "\\10.0.10.79\PEDRO"
                    Dim remoteFinal As String = System.IO.Path.Combine(remoteDir, NumeroNota & ".txt")
                    Dim remoteTmp As String = remoteFinal & ".tmp"

                    Try
                        ' Se já existir tmp, tenta remover
                        If System.IO.File.Exists(remoteTmp) Then
                            System.IO.File.Delete(remoteTmp)
                        End If

                        System.IO.File.Copy(localPath, remoteTmp)

                        ' Se existir o arquivo final, remove para permitir o rename
                        If System.IO.File.Exists(remoteFinal) Then
                            System.IO.File.Delete(remoteFinal)
                        End If

                        System.IO.File.Move(remoteTmp, remoteFinal)

                        ' Aguarda até que o arquivo seja processado/excluído pelo servidor de impressão
                        Dim processed As Boolean = WaitForRemoteFileProcessing(remoteFinal, 60000)
                        If Not processed Then
                            System.IO.File.AppendAllText(logFilePath, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " - Aviso: tempo limite aguardando processamento do arquivo " & remoteFinal & Environment.NewLine)
                        End If
                    Catch ex As Exception
                        EscreverLogErro(ex)
                    End Try
                End If

                i = i + 1
            Next

            executor.Dispose()
        Catch ex As Exception

            EscreverLogErro(ex)

        End Try

    End Sub

    Private Sub CancelarNotaGeral()
        Dim executor As New GRAVA_AUTORIZACAO.AcessoDados.Executar
        Dim protocolo_cancelamento As String
        Dim chave_acesso As String
        Dim cStat As String
        Dim xMotivo As String
        Dim x, y As Integer
        Dim Formato As System.Text.UTF8Encoding
        Formato = System.Text.UTF8Encoding.UTF8

        Try
            For Each Arquivo As String In Directory.GetFiles("C:\Unimake\Uninfe4\Retorno\", "*-can.xml")
                Dim ArquivoXml As StreamReader = New StreamReader(Arquivo)
                Dim Leitura As XmlTextReader = New XmlTextReader(ArquivoXml)
                Dim no As String

                no = ""

                protocolo_cancelamento = ""
                chave_acesso = ""

                Do While (Leitura.Read())
                    Select Case Leitura.NodeType
                        Case XmlNodeType.Element
                            no = Leitura.Name
                        Case XmlNodeType.Text  'Exibir o início do elemento.
                            If no = "nProt" Then
                                protocolo_cancelamento = Leitura.Value
                            End If

                            If no = "chNFe" Then
                                chave_acesso = Leitura.Value
                            End If

                            If no = "cStat" Then
                                cStat = Leitura.Value
                            End If

                            If no = "xMotivo" Then
                                xMotivo = Leitura.Value
                            End If

                    End Select
                Loop


                If protocolo_cancelamento = "" Then
                    If cStat = "420" Then
                        x = InStrRev(xMotivo, ":")
                        y = InStrRev(xMotivo, "]")
                        protocolo_cancelamento = xMotivo.Substring(x, y - x - 1)
                    End If
                End If




                If protocolo_cancelamento = "" Then
                    FileCopy(Arquivo, "C:\Unimake\Uninfe4\SemProtocolo\" & chave_acesso & "-can.xml")
                    Leitura.Close()
                    ArquivoXml.Close()

                Else
                    executor.Dispose()
                    'Gravar Autorização na tabela TNFe_IDENTIFICACAO
                    executor.SQLComando = "dbo.PNFE_GRAVA_CANCELAMENTO"
                    executor.addParametros("protocolo_cancelamento", protocolo_cancelamento, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.VarChar, 20)
                    executor.addParametros("chave_acesso", chave_acesso, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.VarChar, 50)
                    executor.ExecutarOperacao(AcessoDados.Executar.TipoRetorno.eDataSet, "VendasInternet")

                    executor.Dispose()

                    Leitura.Close()
                    ArquivoXml.Close()

                    FileCopy(Arquivo, "C:\Unimake\Uninfe4\Gravados\" & chave_acesso & "-can.xml")
                    Threading.Thread.Sleep(100)
                    File.Delete(Arquivo)
                    Threading.Thread.Sleep(100)

                End If

            Next

        Catch ex As Exception

            EscreverLogErro(ex)

        End Try


    End Sub

    Private Sub EventoCancelarNotaGeral()
        Dim executor As New GRAVA_AUTORIZACAO.AcessoDados.Executar
        Dim protocolo_cancelamento As String
        Dim chave_acesso As String
        Dim cStat As String
        Dim xMotivo As String
        Dim x, y As Integer
        Dim Formato As System.Text.UTF8Encoding
        Formato = System.Text.UTF8Encoding.UTF8

        Try
            For Each Arquivo As String In Directory.GetFiles("C:\Unimake\Uninfe4\Retorno\", "*-ret-env-canc.xml")
                Dim ArquivoXml As StreamReader = New StreamReader(Arquivo)
                Dim Leitura As XmlTextReader = New XmlTextReader(ArquivoXml)
                Dim no As String

                no = ""

                protocolo_cancelamento = ""
                chave_acesso = ""

                Do While (Leitura.Read())
                    Select Case Leitura.NodeType
                        Case XmlNodeType.Element
                            no = Leitura.Name
                        Case XmlNodeType.Text  'Exibir o início do elemento.
                            If no = "nProt" Then
                                protocolo_cancelamento = Leitura.Value
                            End If

                            If no = "chNFe" Then
                                chave_acesso = Leitura.Value
                            End If

                            If no = "cStat" Then
                                cStat = Leitura.Value
                            End If

                            If no = "xMotivo" Then
                                xMotivo = Leitura.Value
                            End If

                    End Select
                Loop


                If protocolo_cancelamento = "" Then
                    If cStat = "420" Then
                        x = InStrRev(xMotivo, ":")
                        y = InStrRev(xMotivo, "]")
                        protocolo_cancelamento = xMotivo.Substring(x, y - x - 1)
                    End If
                End If




                If protocolo_cancelamento = "" Then
                    FileCopy(Arquivo, "C:\Unimake\Uninfe4\SemProtocolo\" & chave_acesso & "-ret-env-canc.xml")
                    Leitura.Close()
                    ArquivoXml.Close()

                Else
                    executor.Dispose()
                    'Gravar Autorização na tabela TNFe_IDENTIFICACAO
                    executor.SQLComando = "dbo.PNFE_GRAVA_CANCELAMENTO"
                    executor.addParametros("protocolo_cancelamento", protocolo_cancelamento, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.VarChar, 20)
                    executor.addParametros("chave_acesso", chave_acesso, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.VarChar, 50)
                    executor.ExecutarOperacao(AcessoDados.Executar.TipoRetorno.eDataSet, "VendasInternet")

                    executor.Dispose()

                    Leitura.Close()
                    ArquivoXml.Close()

                    FileCopy(Arquivo, "C:\Unimake\Uninfe4\Gravados\" & chave_acesso & "-ret-env-canc.xml")
                    Threading.Thread.Sleep(100)
                    File.Delete(Arquivo)
                    Threading.Thread.Sleep(100)

                End If

            Next

        Catch ex As Exception

            EscreverLogErro(ex)

        End Try


    End Sub


    Private Sub Erro_NFe_Email()
        Dim executor As New GRAVA_AUTORIZACAO.AcessoDados.Executar
        Dim chave_acesso As String
        Dim Formato As System.Text.UTF8Encoding
        Formato = System.Text.UTF8Encoding.UTF8
        Dim Tabela As DataSet
        Dim emissor As String
        Dim email_to As String
        Dim email_cc As String
        Dim email_bcc As String
        Dim email_assunto As String
        Dim email_corpo As String
        Dim nome_arquivo As String

        Try
            For Each Arquivo As String In Directory.GetFiles("C:\Unimake\Uninfe4\Retorno\", "*-nfe.err")
                Dim x As Integer

                x = InStrRev(Arquivo, "\")
                chave_acesso = Arquivo.Substring(x, 44)
                nome_arquivo = Arquivo

                executor.Dispose()
                'Gravar erro na tabela TNFe_IDENTIFICACAO
                executor.SQLComando = "dbo.PNFE_GRAVA_ERRO2"
                executor.addParametros("chave_acesso", chave_acesso, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.VarChar, 50)
                executor.ExecutarOperacao(AcessoDados.Executar.TipoRetorno.eDataSet, "VendasInternet")

                Tabela = executor.RetornoDataSet

                emissor = Tabela.Tables(0).Rows(0).Item(2).ToString()
                email_to = Tabela.Tables(0).Rows(0).Item(3).ToString()
                email_cc = Tabela.Tables(0).Rows(0).Item(4).ToString()
                email_assunto = Tabela.Tables(0).Rows(0).Item(5).ToString()
                email_corpo = Tabela.Tables(0).Rows(0).Item(6).ToString()
                email_bcc = Tabela.Tables(0).Rows(0).Item(7).ToString()

                If email_to <> "" Then
                    If emissor = "ALM" Then
                        Email.enviaMensagemEmail("nfe@ipen.br", email_to, email_bcc, email_cc, email_assunto, email_corpo, "smtp.ipen.br", chave_acesso)
                    ElseIf emissor = "IMP" Then
                        Email.enviaMensagemEmail("nfe@ipen.br", email_to, email_bcc, email_cc, email_assunto, email_corpo, "smtp.ipen.br", chave_acesso)
                    ElseIf emissor = "CR1" Then
                        Email.enviaMensagemEmail("nfe@ipen.br", email_to, email_bcc, email_cc, email_assunto, email_corpo, "smtp.ipen.br", chave_acesso)
                    ElseIf emissor = "CR2" Then
                        Email.enviaMensagemEmail("nfe@ipen.br", email_to, email_bcc, email_cc, email_assunto, email_corpo, "smtp.ipen.br", chave_acesso)
                    End If
                End If
                executor.Dispose()
                FileCopy(Arquivo, "C:\Unimake\Uninfe4\Erro\" & chave_acesso & "-nfe.err")
                Threading.Thread.Sleep(300)
                File.Delete(Arquivo)
                Threading.Thread.Sleep(100)
            Next
        Catch ex As Exception

        End Try


    End Sub
    'teste para consultar situação das notas fiscais

    Private Sub GravarXMLGeralSituacao()
        Dim executor As New GRAVA_AUTORIZACAO.AcessoDados.Executar
        Dim Tabela As DataSet
        Dim Tabela2 As DataSet
        Dim Texto2 As DataRow
        Dim Formato As System.Text.UTF8Encoding
        Formato = System.Text.UTF8Encoding.UTF8
        Dim i As Integer
        Dim Documento As String

        Try

            executor.Dispose()

            executor.SQLComando = "dbo.P0110_LISTA_NFE_SITUACAO"
            executor.ExecutarOperacao(AcessoDados.Executar.TipoRetorno.eDataSet, "VendasPelicano")

            Tabela2 = executor.RetornoDataSet

            i = 0
            For Each Texto2 In Tabela2.Tables(0).Rows
                executor.Dispose()
                Dim Emissor As String
                Emissor = ""

                'Gera arquivo da Nota Fiscal Eletrônica (XML)
                executor.SQLComando = "dbo.PNFe_GERAXML_NF"
                executor.addParametros("notafis_oid", CInt(Tabela2.Tables(0).Rows(i).Item(0).ToString()), AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.Int, 4)
                executor.addParametros("nf", CInt(Tabela2.Tables(0).Rows(i).Item(1).ToString()), AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.Int, 4)
                executor.ExecutarOperacao(AcessoDados.Executar.TipoRetorno.eDataSet, "VendasInternet")

                Tabela = executor.RetornoDataSet

                Documento = ""

                Documento = "<?xml version='1.0' encoding='utf-8'?><consSitNFe xmlns='http://www.portalfiscal.inf.br/nfe' versao='3.10'><tpAmb>1</tpAmb><xServ>CONSULTAR</xServ><chNFe>" & Tabela.Tables(1).Rows(0).Item(0).ToString() & "</chNFe></consSitNFe>"

                If Documento <> "" Then
                    Dim ArquivoXml As System.IO.File
                    ArquivoXml.WriteAllText("C:\Unimake\Uninfe4\Envio\" & Tabela.Tables(1).Rows(0).Item(0).ToString() & "-ped-sit.xml", Documento, Formato)
                End If
                i = i + 1
            Next

            executor.Dispose()
        Catch ex As Exception

            EscreverLogErro(ex)

        End Try


    End Sub

    Private Sub GravarAutorizacaoSituacao()
        Try
            ' Usar lock para garantir processamento sequencial
            SyncLock fileProcessingLock
                For Each Arquivo As String In Directory.GetFiles("C:\Unimake\Uninfe4\Retorno\", "*-sit.xml")
                    ProcessarArquivoSituacao(Arquivo)
                Next
            End SyncLock
        Catch ex As Exception
            EscreverLogErro(ex)
        End Try
    End Sub

    Private Sub ProcessarArquivoSituacao(Arquivo As String)
        ' Variáveis isoladas para cada arquivo - evita sobreposição de conteúdo
        Dim executor As New GRAVA_AUTORIZACAO.AcessoDados.Executar
        Dim Tabela As DataSet
        Dim Tabela2 As DataSet
        Dim recibo_numero As String
        Dim protocolo_autorizacao As String
        Dim data_retorno As Date
        Dim envio_data As Date
        Dim data_processamento As Date
        Dim envio_lote As String
        Dim recibo_data As Date
        Dim status_nf As Integer
        Dim chave_acesso As String
        Dim AreaEmissao As String
        Dim Texto As DataRow
        Dim Formato As System.Text.UTF8Encoding
        Formato = System.Text.UTF8Encoding.UTF8

        Try
            Dim ArquivoXml As StreamReader = New StreamReader(Arquivo)
            Dim Leitura As XmlTextReader = New XmlTextReader(ArquivoXml)
            Dim no As String

            no = ""

            recibo_numero = ""
            protocolo_autorizacao = ""
            envio_lote = ""
            status_nf = 0
            chave_acesso = ""

            Do While (Leitura.Read())
                Select Case Leitura.NodeType
                    Case XmlNodeType.Element
                        no = Leitura.Name
                    Case XmlNodeType.Text  'Exibir o início do elemento.
                        If no = "nProt" Then
                            protocolo_autorizacao = Leitura.Value
                            recibo_numero = Leitura.Value
                        End If

                        If no = "dhRecbto" Then
                            data_retorno = Leitura.Value
                            envio_data = Leitura.Value
                            recibo_data = Leitura.Value
                            data_processamento = Leitura.Value
                        End If

                        If no = "chNFe" Then
                            envio_lote = Leitura.Value
                            chave_acesso = Leitura.Value
                        End If

                        If no = "cStat" Then
                            status_nf = Leitura.Value
                        End If
                End Select
            Loop



            If protocolo_autorizacao = "" Then
                FileCopy(Arquivo, "C:\Unimake\Uninfe4\SemProtocolo\" & chave_acesso & "-sit.xml")
                Leitura.Close()
                ArquivoXml.Close()

            Else
                executor.Dispose()
                'Gravar Autorização na tabela TNFe_IDENTIFICACAO
                executor.SQLComando = "dbo.PNFE_GRAVA_AUTORIZACAO"
                executor.addParametros("recibo_numero", recibo_numero, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.VarChar, 20)
                executor.addParametros("protocolo_autorizacao", protocolo_autorizacao, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.VarChar, 20)
                executor.addParametros("data_retorno", data_retorno, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.DateTime, 19)
                executor.addParametros("envio_data", envio_data, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.DateTime, 19)
                executor.addParametros("data_processamento", data_processamento, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.DateTime, 19)
                executor.addParametros("envio_lote", envio_lote, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.VarChar, 50)
                executor.addParametros("recibo_data", recibo_data, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.DateTime, 19)
                executor.addParametros("status_nf", status_nf, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.Int, 4)
                executor.addParametros("chave_acesso", chave_acesso, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.VarChar, 50)
                executor.ExecutarOperacao(AcessoDados.Executar.TipoRetorno.eDataSet, "VendasInternet")

                Tabela2 = executor.RetornoDataSet
                If Tabela2.Tables(0).Rows(0).Item(0).ToString() = "" Then
                    AreaEmissao = "CR1"
                Else
                    AreaEmissao = Trim(Tabela2.Tables(0).Rows(0).Item(0).ToString())
                End If

                executor.Dispose()

                Leitura.Close()
                ArquivoXml.Close()

                Dim i As Int64
                Dim Documento


                FileCopy(Arquivo, "C:\Unimake\Uninfe4\Gravados\" & chave_acesso & "-pro-rec.xml")

                Threading.Thread.Sleep(2000)

                Select Case AreaEmissao
                    Case "IMP"
                        'Imprimir arquivo com a autorização
                        executor.SQLComando = "dbo.P0110_GERA_NOTA_NFE_GERAL"
                        executor.addParametros("recibo_numero", recibo_numero, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.VarChar, 20)
                    Case "ALM"
                        executor.SQLComando = "dbo.P0110_GERA_NOTA_NFE_GERAL"
                        executor.addParametros("recibo_numero", recibo_numero, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.VarChar, 20)
                    Case "CTR"
                        executor.SQLComando = "dbo.P0110_GERA_NOTA_NFE_CTR"
                        executor.addParametros("recibo_numero", recibo_numero, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.VarChar, 20)
                    Case "CR2"
                        executor.SQLComando = "dbo.P0110_GERA_NOTA_NFE_GERAL"
                        executor.addParametros("recibo_numero", recibo_numero, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.VarChar, 20)
                    Case Else
                        executor.SQLComando = "dbo.P0110_GERA_NOTA_NFE_DIRF"
                        executor.addParametros("recibo_numero", recibo_numero, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.VarChar, 20)
                End Select

                executor.ExecutarOperacao(AcessoDados.Executar.TipoRetorno.eDataSet, "VendasPelicano")
                Tabela = executor.RetornoDataSet

                Dim TRANS As String

                TRANS = ""
                Documento = ""

                i = 0
                For Each Texto In Tabela.Tables(0).Rows
                    Documento = Documento & Tabela.Tables(0).Rows(i).Item(1).ToString() & vbCrLf
                    i = i + 1
                Next

                Dim contador As Integer = 0

                If Documento <> "" Then

                    contador = contador + 1

                    ' Escreve no log a execução
                    Dim logMessage As String = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " - Execução " & contador & " - Criando arquivo: " & chave_acesso & ".txt - Emissor: " & AreaEmissao

                    Dim ArquivoTxt As System.IO.File
                    Select Case AreaEmissao
                        Case "IMP"
                            ArquivoTxt.WriteAllText("C:\Unimake\Uninfe4\Importacao\" & chave_acesso & ".txt", Documento, Formato)
                            Threading.Thread.Sleep(200)
                            System.IO.File.Copy("C:\Unimake\Uninfe4\Importacao\" & chave_acesso & ".txt", "\\10.0.10.79\IMPORTACAO")
                            System.IO.File.AppendAllText(logFilePath, logMessage & Environment.NewLine)
                        Case "ALM"
                            ArquivoTxt.WriteAllText("C:\Unimake\Uninfe4\Pedro\" & chave_acesso & ".txt", Documento, Formato)
                            Threading.Thread.Sleep(200)
                            System.IO.File.Copy("C:\Unimake\Uninfe4\Pedro\" & chave_acesso & ".txt", "\\10.0.10.79\PEDRO")
                            System.IO.File.AppendAllText(logFilePath, logMessage & Environment.NewLine)
                        Case "CTR"
                            ArquivoTxt.WriteAllText("C:\Unimake\Uninfe4\CTR\" & chave_acesso & ".txt", Documento, Formato)
                            Threading.Thread.Sleep(200)
                            System.IO.File.Copy("C:\Unimake\Uninfe4\CTR\" & chave_acesso & ".txt", "\\10.0.10.79\CTR")
                            System.IO.File.AppendAllText(logFilePath, logMessage & Environment.NewLine)
                        Case "CR2"
                            ArquivoTxt.WriteAllText("C:\Unimake\Uninfe4\Dirf\" & chave_acesso & ".txt", Documento, Formato)
                            Threading.Thread.Sleep(200)
                            System.IO.File.Copy("C:\Unimake\Uninfe4\Dirf\" & chave_acesso & ".txt", "\\10.0.10.79\Entrada")
                            System.IO.File.AppendAllText(logFilePath, logMessage & Environment.NewLine)
                        Case Else
                            TRANS = Trim(Tabela.Tables(1).Rows(0).Item(0).ToString())
                            ArquivoTxt.WriteAllText("C:\Unimake\Uninfe4\Dirf\" & chave_acesso & ".txt", Documento, Formato)
                            Threading.Thread.Sleep(200)
                            If TRANS = "BND" Then
                                System.IO.File.Copy("C:\Unimake\Uninfe4\Dirf\" & chave_acesso & ".txt", "\\10.0.10.79\BND")
                                System.IO.File.AppendAllText(logFilePath, logMessage & Environment.NewLine)
                            ElseIf TRANS = "REM" Then
                                System.IO.File.Copy("C:\Unimake\Uninfe4\Dirf\" & chave_acesso & ".txt", "\\10.0.10.79\REM")
                                System.IO.File.AppendAllText(logFilePath, logMessage & Environment.NewLine)
                            ElseIf TRANS = "HVL" Then
                                System.IO.File.Copy("C:\Unimake\Uninfe4\Dirf\" & chave_acesso & ".txt", "\\10.0.10.79\HVL")
                                System.IO.File.AppendAllText(logFilePath, logMessage & Environment.NewLine)
                            Else
                                System.IO.File.Copy("C:\Unimake\Uninfe4\Dirf\" & chave_acesso & ".txt", "\\10.0.10.79\Entrada")
                                System.IO.File.AppendAllText(logFilePath, logMessage & Environment.NewLine)
                            End If
                    End Select
                End If
                executor.Dispose()
            End If
            Threading.Thread.Sleep(100)
            File.Delete(Arquivo)
            Threading.Thread.Sleep(100)

        Catch ex As Exception
            EscreverLogErro(ex)
        End Try
    End Sub


    Private Sub GravarAutorizacaoCartaCorrecao()
        Try
            ' Usar lock para garantir processamento sequencial
            SyncLock fileProcessingLock
                For Each Arquivo As String In Directory.GetFiles("C:\Unimake\Uninfe4\Retorno\", "*-ret-env-cce.xml")
                    ProcessarArquivoCartaCorrecao(Arquivo)
                Next
            End SyncLock
        Catch ex As Exception
            EscreverLogErro(ex)
        End Try
    End Sub

    Private Sub ProcessarArquivoCartaCorrecao(Arquivo As String)
        ' Variáveis isoladas para cada arquivo - evita sobreposição de conteúdo
        Dim executor As New GRAVA_AUTORIZACAO.AcessoDados.Executar
        Dim Tabela As DataSet
        Dim Tabela2 As DataSet
        Dim recibo_numero As String
        Dim protocolo_autorizacao As String
        Dim data_retorno As Date
        Dim envio_data As Date
        Dim data_processamento As Date
        Dim envio_lote As String
        Dim recibo_data As Date
        Dim status_nf As Integer
        Dim nSeqEvento As Integer
        Dim chave_acesso As String
        Dim chave_acesso_cce As String
        Dim AreaEmissao As String
        Dim Texto As DataRow
        Dim Formato As System.Text.UTF8Encoding
        Formato = System.Text.UTF8Encoding.UTF8

        Try
            Dim ArquivoXml As StreamReader = New StreamReader(Arquivo)
            Dim Leitura As XmlTextReader = New XmlTextReader(ArquivoXml)
            Dim no As String

            no = ""

            recibo_numero = ""
            protocolo_autorizacao = ""
            envio_lote = ""
            status_nf = 0
            chave_acesso = ""
            chave_acesso_cce = ""
            nSeqEvento = 0

            Do While (Leitura.Read())
                Select Case Leitura.NodeType
                    Case XmlNodeType.Element
                        no = Leitura.Name
                    Case XmlNodeType.Text  'Exibir o início do elemento.

                        If no = "nSeqEvento" Then
                            nSeqEvento = Leitura.Value
                        End If

                        If no = "nProt" Then
                            protocolo_autorizacao = Leitura.Value
                        End If

                        If no = "dhRegEvento" Then
                            data_retorno = Leitura.Value
                            envio_data = Leitura.Value
                            recibo_data = Leitura.Value
                            data_processamento = Leitura.Value
                        End If

                        If no = "chNFe" Then
                            envio_lote = Leitura.Value
                            chave_acesso = Leitura.Value
                        End If

                        If no = "cStat" Then
                            status_nf = Leitura.Value
                        End If
                End Select
            Loop

            If nSeqEvento < 10 Then
                chave_acesso_cce = chave_acesso & "0" & nSeqEvento.ToString()
            Else
                chave_acesso_cce = chave_acesso & nSeqEvento.ToString()
            End If

            If protocolo_autorizacao = "" Then
                FileCopy(Arquivo, "C:\Unimake\Uninfe4\SemProtocolo\" & chave_acesso_cce & "-ret-env-cce.xml")
                Leitura.Close()
                ArquivoXml.Close()
            Else
                executor.Dispose()
                executor.SQLComando = "dbo.PNFE_CartaCorrecao_Autorizacao"
                executor.addParametros("protocolo_autorizacao", protocolo_autorizacao, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.VarChar, 20)
                executor.addParametros("data_retorno", data_retorno, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.DateTime, 19)
                executor.addParametros("status_nf", status_nf, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.Int, 4)
                executor.addParametros("nSeqEvento", nSeqEvento, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.Int, 4)
                executor.addParametros("chave_acesso", chave_acesso, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.VarChar, 50)
                executor.ExecutarOperacao(AcessoDados.Executar.TipoRetorno.eDataSet, "VendasInternet")

                Tabela2 = executor.RetornoDataSet

                If Tabela2.Tables(0).Rows(0).Item(0).ToString() = "" Then
                    AreaEmissao = "CR1"
                Else
                    AreaEmissao = Trim(Tabela2.Tables(0).Rows(0).Item(0).ToString())
                End If

                executor.Dispose()

                Leitura.Close()
                ArquivoXml.Close()

                Dim i As Int64
                Dim Documento

                FileCopy(Arquivo, "C:\Unimake\Uninfe4\Gravados\" & chave_acesso_cce & "-ret-env-cce.xml")

                Threading.Thread.Sleep(2000)

                executor.SQLComando = "dbo.PNFe_CartaCorrecao_txt"
                executor.addParametros("nSeqEvento", nSeqEvento, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.Int, 4)
                executor.addParametros("chave_acesso", chave_acesso, AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.VarChar, 50)
                executor.ExecutarOperacao(AcessoDados.Executar.TipoRetorno.eDataSet, "VendasInternet")
                Tabela = executor.RetornoDataSet

                Dim TRANS As String

                TRANS = ""
                Documento = ""

                i = 0
                For Each Texto In Tabela.Tables(0).Rows
                    Documento = Documento & Tabela.Tables(0).Rows(i).Item(1).ToString() & vbCrLf
                    i = i + 1
                Next

                Dim ArquivoTxt As System.IO.File
                Select Case AreaEmissao
                    Case "IMP"
                        ArquivoTxt.WriteAllText("C:\Unimake\Uninfe4\Importacao\" & chave_acesso_cce & ".txt", Documento, Formato)
                        Threading.Thread.Sleep(200)
                        System.IO.File.Copy("C:\Unimake\Uninfe4\Importacao\" & chave_acesso_cce & ".txt", "\\10.0.10.79\IMPORTACAO")
                        executor.Dispose()
                    Case "ALM"
                        ArquivoTxt.WriteAllText("C:\Unimake\Uninfe4\Pedro\" & chave_acesso_cce & ".txt", Documento, Formato)
                        Threading.Thread.Sleep(200)
                        System.IO.File.Copy("C:\Unimake\Uninfe4\Pedro\" & chave_acesso_cce & ".txt", "\\10.0.10.79\PEDRO")
                        executor.Dispose()
                    Case "CR1"
                        ArquivoTxt.WriteAllText("C:\Unimake\Uninfe4\Dirf\" & chave_acesso_cce & ".txt", Documento, Formato)
                        Threading.Thread.Sleep(200)
                        System.IO.File.Copy("C:\Unimake\Uninfe4\Dirf\" & chave_acesso_cce & ".txt", "\\10.0.10.79\ENTRADA")
                        executor.Dispose()
                    Case "CR2"
                        ArquivoTxt.WriteAllText("C:\Unimake\Uninfe4\Dirf\" & chave_acesso_cce & ".txt", Documento, Formato)
                        Threading.Thread.Sleep(200)
                        System.IO.File.Copy("C:\Unimake\Uninfe4\Dirf\" & chave_acesso_cce & ".txt", "\\10.0.10.79\ENTRADA")
                        executor.Dispose()
                    Case Else
                        ArquivoTxt.WriteAllText("C:\Unimake\Uninfe4\Pedro\" & chave_acesso_cce & ".txt", Documento, Formato)
                        Threading.Thread.Sleep(200)
                        System.IO.File.Copy("C:\Unimake\Uninfe4\Pedro\" & chave_acesso_cce & ".txt", "\\10.0.10.79\PEDRO")
                        executor.Dispose()
                End Select

            End If
            Threading.Thread.Sleep(100)
            File.Delete(Arquivo)
            Threading.Thread.Sleep(100)

        Catch ex As Exception
            EscreverLogErro(ex)
        End Try
    End Sub

    Private Sub GravarXMLCartaCorrecao()
        Dim executor As New GRAVA_AUTORIZACAO.AcessoDados.Executar
        Dim Tabela As DataSet
        Dim Tabela2 As DataSet
        Dim Texto As DataRow
        Dim Texto2 As DataRow
        Dim Formato As System.Text.UTF8Encoding
        Formato = System.Text.UTF8Encoding.UTF8
        Dim i As Integer
        Dim Documento As String

        Try

            executor.Dispose()

            executor.SQLComando = "dbo.PNFe_CartaCorrecao_Lista"
            executor.ExecutarOperacao(AcessoDados.Executar.TipoRetorno.eDataSet, "VendasInternet")

            Tabela2 = executor.RetornoDataSet

            i = 0
            For Each Texto2 In Tabela2.Tables(0).Rows
                executor.Dispose()
                Dim Emissor As String
                Emissor = ""

                'Gera arquivo da Carta de Correção (XML)
                executor.SQLComando = "dbo.PNFe_CartaCorrecao_XML"
                executor.addParametros("id_correcao", CInt(Tabela2.Tables(0).Rows(i).Item(0).ToString()), AcessoDados.Executar.TipoDirecao.eTPEntrada, SqlDbType.Int, 4)
                executor.ExecutarOperacao(AcessoDados.Executar.TipoRetorno.eDataSet, "VendasInternet")
                Tabela = executor.RetornoDataSet

                Documento = ""

                Documento = Documento & Tabela.Tables(0).Rows(0).Item(0).ToString()

                If Documento <> "" Then
                    Dim ArquivoXml As System.IO.File
                    ArquivoXml.WriteAllText("C:\Unimake\Uninfe4\Envio\" & Tabela2.Tables(0).Rows(i).Item(1).ToString() & "-env-cce.xml", Documento, Formato)
                End If
                i = i + 1
            Next

            executor.Dispose()
        Catch ex As Exception

            EscreverLogErro(ex)

        End Try


    End Sub


End Class




