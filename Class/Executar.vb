Namespace AcessoDados
    Public Class Executar
        Inherits AcessoDados.Conexao

        Private Comando As New SqlClient.SqlCommand
        Dim varSQLComando As String
        Dim varNomeDataSet As String
        Dim varRetornoDataReader As SqlClient.SqlDataReader
        Dim varRetornoDataSet As DataSet
        Dim varRetornoDataTable As DataTable
        Dim strMostraErro As String
        Dim strMensagem As String

        Public Enum TipoDirecao
            eTPEntrada = ParameterDirection.Input
            eTPEntradaSaida = ParameterDirection.Output
            eTPSaida = ParameterDirection.Output
            eTPRetorno = ParameterDirection.ReturnValue
        End Enum
        Public Enum TipoRetorno
            eDataSet = 1
            eDataTable = 2
            eReader = 3
            eNenhum = 4
        End Enum
        Public Enum TipoComando
            eStoredProcedure = CommandType.StoredProcedure
            eTableDirect = CommandType.TableDirect
            eText = CommandType.Text
        End Enum

        Public ReadOnly Property RetornoDataTable() As DataTable
            Get
                Return varRetornoDataTable
            End Get
        End Property

        Public ReadOnly Property RetornoDataSet() As DataSet
            Get
                Return varRetornoDataSet
            End Get
        End Property

        Public ReadOnly Property RetornoDataReader() As SqlClient.SqlDataReader
            Get
                Return varRetornoDataReader
            End Get
        End Property

        Public WriteOnly Property NomeDataSet() As String
            Set(ByVal value As String)
                varNomeDataSet = value
            End Set
        End Property

        Public WriteOnly Property SQLComando() As String
            Set(ByVal value As String)
                'Debug.AutoFlush = True
                'Debug.WriteLine(value)
                If Comando Is Nothing Then Comando = New SqlClient.SqlCommand
                varSQLComando = value
                Comando.CommandType = CommandType.StoredProcedure
                Comando.CommandText = value
            End Set
        End Property
        Public Sub addParametros(ByVal strParamNome As String, ByVal strParamValor As String, ByVal eParamTipoDirecao As TipoDirecao, ByVal eParamDataType As SqlDbType, ByVal intParamTamanho As Integer)
            'ADICIONA OS PARAMETROS NA STRING DE ACESSO A PROCEDURE
            Dim Parametro As SqlClient.SqlParameter
            Dim nome As String
            Try

                nome = strParamNome
                If strParamNome.IndexOf("@") = -1 Then
                    nome = "@" & nome
                End If
                If Comando Is Nothing Then Comando = New SqlClient.SqlCommand

                'Debug.WriteLine(strParamNome & " " & eParamDataType.ToString & "(" & intParamTamanho & ") = " & strParamValor)

                Parametro = New SqlClient.SqlParameter(nome, eParamDataType, intParamTamanho)
                Parametro.Value = strParamValor
                Parametro.Direction = eParamTipoDirecao
                Comando.Parameters.Add(Parametro)
            Catch ex As Exception
                strMostraErro = ex.Message
                Throw New Exception(strMostraErro)

            End Try

        End Sub
        Public Sub ExecutarOperacao(ByVal eParamTipoRetorno As TipoRetorno, ByVal eConexaoBanco As String)
            'EXECUTA A CONEXÃO COM O BANCO DE DADOS - PASSANDO A PROCEDURE(E SEUS PARAMETROS) QUE DEVERÁ SER EXECUTADA
            Dim conexao As New Conexao
            Dim dataAdapter As New SqlClient.SqlDataAdapter
            'Dim retorno As SqlClient.SqlDataReader

            Try
                Comando.Connection = Me.EfetuarConexao(eConexaoBanco)

                If varSQLComando.IndexOf("_INCLUIR_") > -1 Then
                    dataAdapter.InsertCommand = Comando
                ElseIf varSQLComando.IndexOf("_ALTERAR_") > -1 Then
                    dataAdapter.UpdateCommand = Comando
                ElseIf varSQLComando.IndexOf("_EXCLUIR_") > -1 Then
                    dataAdapter.DeleteCommand = Comando
                Else
                    dataAdapter.SelectCommand = Comando
                End If

                Select Case eParamTipoRetorno
                    Case TipoRetorno.eDataSet
                        If varNomeDataSet = "" Then
                            Dim strNomeDataSet As String
                            Dim intPosicao As Integer
                            intPosicao = varSQLComando.IndexOf("_", varSQLComando.IndexOf("PROC_") + 5)
                            strNomeDataSet = varSQLComando.Substring(intPosicao + 1)
                            varNomeDataSet = strNomeDataSet
                        End If
                        If varRetornoDataSet Is Nothing Then varRetornoDataSet = New DataSet("SGCR_NET")

                        dataAdapter.Fill(varRetornoDataSet, varNomeDataSet)
                    Case TipoRetorno.eDataTable
                        varRetornoDataTable = New DataTable
                        dataAdapter.Fill(varRetornoDataTable)
                    Case TipoRetorno.eReader
                        varRetornoDataReader = Comando.ExecuteReader()
                    Case Else
                        Comando.ExecuteNonQuery()
                End Select

            Catch ex As Exception
                strMostraErro = ex.Message
                'Throw New Exception(strMostraErro)

            End Try

            MyBase.Dispose()

        End Sub
        Public Sub LimparParametros()
            Me.Comando.Parameters.Clear()
        End Sub

        Public Overloads Sub Dispose()
            If Not varRetornoDataReader Is Nothing Then varRetornoDataReader.Close()
            If Not varRetornoDataSet Is Nothing Then varRetornoDataSet = Nothing
            If Not varRetornoDataTable Is Nothing Then varRetornoDataTable = Nothing

            If Not Comando Is Nothing Then
                Comando = Nothing
            End If
            MyBase.Dispose()
            MyBase.Finalize()
        End Sub

    End Class
End Namespace
