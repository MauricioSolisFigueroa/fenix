Imports Microsoft.VisualBasic
Imports System.Net
Imports System.Data
Imports System.IO

Public Class Comunes
    Public Shared Function strConexionSara() As String
        Dim strIP As String = ConsultaIp()
        Dim strCadenaConexion As String = Nothing

        Select Case strIP
            Case "132.248.180.207"
                strCadenaConexion = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=132.248.180.228)(PORT=1521)))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=BASESDSI)));User Id = saradb; Password = 3db5HmKj;"
            Case "132.248.180.110"
                'strCadenaConexion = "User ID=usr_mercurio_pago;Password=HTDK+341;Initial Catalog=DB_MERCURIO_PAGOS;Data Source=132.248.180.110"
            Case Else
                'strCadenaConexion = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=132.248.180.50)(PORT=1521)))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=BASESRECTOR)));User Id = saradb; Password = saradb;"
                strCadenaConexion = "Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=132.248.180.228)(PORT=1521)))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=BASESDSI)));User Id = saradb; Password = 3db5HmKj;"
        End Select
        Return strCadenaConexion
    End Function

    Public Shared Function ConsultaIp() As String
        Dim myHost As String = Dns.GetHostName
        Dim ipEntry As IPHostEntry = Dns.GetHostEntry(myHost)
        Dim mstrIP As String = ""
        'Se agrego esta parte para que pudiera sacar la ip con maquinas que tienen windows 7 u 8
        For Each tmpIpAddress As IPAddress In ipEntry.AddressList
            If tmpIpAddress.AddressFamily = Sockets.AddressFamily.InterNetwork Then
                Dim ipAddress As String = tmpIpAddress.ToString
                mstrIP = ipAddress
                Exit For
            End If
        Next
        Return mstrIP
    End Function

    Public Shared Sub EnviaABitacora(ByVal pstrFuncionError As String, ByVal pstrError As String)
        Try
            Dim strCadena As String
            Dim Escritor As New StreamWriter(Path.GetPathRoot(Environment.SystemDirectory).ToString & "temp/Fenix_Errores.txt", True, System.Text.ASCIIEncoding.Default)
            Escritor.Flush()
            strCadena = Now & "|FuncionError: " & pstrFuncionError & "|Error: " & pstrError
            Escritor.WriteLine(strCadena)
            Escritor.Close()
        Catch ex As Exception
            Console.Write("Error ...")
        End Try
    End Sub

    'Se agregó 31/01/2023
    Public Shared Function ConsultaTiposDatosDataSet(ByVal dsConsulta As DataSet) As Constantes.TiposDatosDataSet
        Try
            If dsConsulta.Tables(0).Columns.Count >= 1 Then
                If dsConsulta.Tables(0).Rows.Count > 0 Then
                    Return Constantes.TiposDatosDataSet.ConDatos
                Else
                    Return Constantes.TiposDatosDataSet.SinDatos
                End If
            ElseIf dsConsulta.Tables(0).Columns.Count = 1 Then
                Select Case UCase(dsConsulta.Tables(0).Columns(0).ColumnName.ToString)
                    Case "RESPUESTA"
                        If dsConsulta.Tables(0).Rows.Count > 0 Then
                            If dsConsulta.Tables(0).Rows(0).Item(0).ToString.Trim <> "" Then
                                Return Constantes.TiposDatosDataSet.Respuesta
                            Else
                                Return Constantes.TiposDatosDataSet.OcurrioError
                            End If
                        Else
                            Return Constantes.TiposDatosDataSet.OcurrioError
                        End If
                    Case "ERROR"
                        If dsConsulta.Tables(0).Rows.Count > 0 Then
                            If dsConsulta.Tables(0).Rows(0).Item(0).ToString.Trim <> "" Then
                                Return Constantes.TiposDatosDataSet.MensajeError
                            Else
                                Return Constantes.TiposDatosDataSet.OcurrioError
                            End If
                        Else
                            Return Constantes.TiposDatosDataSet.OcurrioError
                        End If
                    Case Else
                        If dsConsulta.Tables(0).Rows.Count > 0 Then
                            If dsConsulta.Tables(0).Rows(0).Item(0).ToString.Trim <> "" Then
                                Return Constantes.TiposDatosDataSet.MensajeError
                            Else
                                Return Constantes.TiposDatosDataSet.OcurrioError
                            End If
                        Else
                            Return Constantes.TiposDatosDataSet.OcurrioError
                        End If
                End Select
            Else
                Return Constantes.TiposDatosDataSet.OcurrioError
            End If
        Catch ex As Exception
            Return Constantes.TiposDatosDataSet.OcurrioError
        End Try
    End Function

    'Se agregó 31/01/2023
    'Public Shared Function fnDsConfiguracionVacaciones() As DataSet
    '    Try
    '        Dim objConsultas As New Consultas
    '        Dim dsConfiguracion As DataSet = objConsultas.fnDsConsultaConfiguracionVacaciones()
    '        'Dim dsConfiguracion As DataSet = fnDsPruebaMensaje()
    '        Dim dtFechaServidor As Date = objConsultas.ConsultaFechaServidor()

    '        Select Case Comunes.ConsultaTiposDatosDataSet(dsConfiguracion)
    '            Case Constantes.TiposDatosDataSet.ConDatos
    '                If CDate(dsConfiguracion.Tables(0).Rows(0).Item("dFechaInicio").ToString) <= dtFechaServidor And dtFechaServidor <= CDate(dsConfiguracion.Tables(0).Rows(0).Item("dFechaFin").ToString) Then
    '                    Return dsConfiguracion
    '                Else
    '                    Return Nothing
    '                End If
    '            Case Else
    '                Return Nothing
    '        End Select
    '    Catch ex As Exception
    '        Return Nothing
    '    End Try
    'End Function
End Class
