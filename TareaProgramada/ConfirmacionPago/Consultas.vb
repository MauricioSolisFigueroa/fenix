
Imports Microsoft.VisualBasic
Imports Oracle.ManagedDataAccess.Client

'Imports Oracle.ManagedDataAccess
'Imports System.Data.SqlClient

Public Class Consultas
#Region "Variables"
    Private mstrCURP As String = ""

    Private mintIdPersona As Integer = Nothing
    Private mintTipoTemporalidad As Integer = Nothing
#End Region

#Region "Propiedades"
    Public Property IdPersona As Integer
        Get
            Return mintIdPersona
        End Get
        Set(value As Integer)
            mintIdPersona = value
        End Set
    End Property
    Public Property TipoTemporalidad As Integer
        Get
            Return mintTipoTemporalidad
        End Get
        Set(value As Integer)
            mintTipoTemporalidad = value
        End Set
    End Property
#End Region

#Region "Consultas"
    ''' <summary>
    ''' Descripcion
    ''' </summary>
    ''' <returns></returns>
    ''' Formulario(s):
    ''' Control(es): 
    ''' Método(s): ConsultaReferencia
    ''' Nombre de quién elaboró: Noé Mauricio Solís Figueroa
    ''' Fecha de creación: 2023/03/14
    ''' Nombre de quién actualizó: 
    ''' Fecha de actualización: 
    Public Function ConsultaFechaServidor() As Date
        Dim connstring As String = Comunes.strConexionSara
        Dim monjConexionOracle As New OracleConnection

        Try
            Dim strFecha As String = ""
            monjConexionOracle.ConnectionString = connstring
            monjConexionOracle.Open()
            Dim cmdComando As New OracleCommand("SPD_CONSULTA_FECHA_SERVIDOR", monjConexionOracle)
            cmdComando.CommandType = CommandType.StoredProcedure
            cmdComando.Parameters.Add("PcFecha", OracleDbType.Date, ParameterDirection.Output)

            'Dim reader As OracleDataReader = cmdComando.ExecuteReader()
            'strFecha = reader.GetString(0)

            strFecha = cmdComando.ExecuteScalar()


            If IsDate(strFecha) Then
                Return CDate(strFecha)
            Else
                ConsultaFechaServidor = Today.Date
            End If

        Catch ex As Exception

        End Try
    End Function
    ''' <summary>
    ''' Descripcion
    ''' </summary>
    ''' <returns></returns>
    ''' Formulario(s):
    ''' Control(es): 
    ''' Método(s): ConsultaReferencia
    ''' Nombre de quién elaboró: Noé Mauricio Solís Figueroa
    ''' Fecha de creación: 2023/03/14
    ''' Nombre de quién actualizó: 
    ''' Fecha de actualización: 
    Public Function ConsultaPersona(ByVal IdPersona As Integer)
        Dim connstring As String = Comunes.strConexionSara
        Dim monjConexionOracle As New OracleConnection

        Try
            Dim objDsPersona As New DataSet
            Dim strFecha As String = Nothing
            monjConexionOracle.ConnectionString = connstring
            monjConexionOracle.Open()
            Dim cmdComando As New OracleCommand("SPD_CONSULTA_DATOS_USUARIO", monjConexionOracle)
            cmdComando.CommandType = CommandType.StoredProcedure

            Dim objParam As New OracleParameter()
            objParam.ParameterName = "pnidPersona"
            objParam.OracleDbType = OracleDbType.Decimal
            objParam.Value = IdPersona
            objParam.Direction = ParameterDirection.Input
            cmdComando.Parameters.Add(objParam)

            cmdComando.Parameters.Add(New OracleParameter("vtPersona", OracleDbType.RefCursor)).Direction = ParameterDirection.Output

            Dim adaptador As New OracleDataAdapter(cmdComando)
            adaptador.Fill(objDsPersona)
            Return objDsPersona
        Catch ex As Exception
            Console.Write("Ocurrio un error en la base de datos")
        End Try
    End Function

    ''' <summary>
    ''' Descripcion
    ''' </summary>
    ''' <returns></returns>
    ''' Formulario(s):
    ''' Control(es): 
    ''' Método(s): ConsultaReferencia
    ''' Nombre de quién elaboró: Noé Mauricio Solís Figueroa
    ''' Fecha de creación: 2023/03/21
    ''' Nombre de quién actualizó: 
    ''' Fecha de actualización: 
    Public Function ConsultaCorreos(ByVal TipoTemporalidad As Integer)
        Dim connstring As String = Comunes.strConexionSara
        Dim monjConexionOracle As New OracleConnection
        Try
            Dim objDsDatosCorreo As New DataSet
            Dim strFecha As String = Nothing
            monjConexionOracle.ConnectionString = connstring
            monjConexionOracle.Open()
            Dim cmdComando As New OracleCommand("SPD_CONSULTA_CORREOS", monjConexionOracle)
            cmdComando.CommandType = CommandType.StoredProcedure

            Dim objParam As New OracleParameter()
            objParam.ParameterName = "PNTIPOTEMPORALIDAD"
            objParam.OracleDbType = OracleDbType.Decimal
            objParam.Value = TipoTemporalidad
            objParam.Direction = ParameterDirection.Input
            cmdComando.Parameters.Add(objParam)

            cmdComando.Parameters.Add(New OracleParameter("VTCORREOS", OracleDbType.RefCursor)).Direction = ParameterDirection.Output

            Dim adaptador As New OracleDataAdapter(cmdComando)
            adaptador.Fill(objDsDatosCorreo)
            Return objDsDatosCorreo
        Catch ex As Exception
            Console.Write("Ocurrio un error en la base de datos")
        End Try
    End Function

#End Region
End Class
