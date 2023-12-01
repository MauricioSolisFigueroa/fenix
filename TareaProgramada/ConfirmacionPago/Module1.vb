Imports System.Data
Imports System.Net.Mail
Imports System.IO
Imports CookComputing.XmlRpc

Module Module1
    Public strNombreWebService As String = ""
    Sub Main()
        Try
            Dim strVarForm(14) As String
            'Dim binValidaVac As Boolean
            Dim objDSPagosPendientes As New Data.DataSet
            Dim objConsulta As New Consultas
            'Dim objDsPersona As New DataSet
            Dim fechaServidor As String
            'Dim IdPersona As Integer = 10596
            Dim objDsDatosCorreo As New DataSet
            Dim contador As Integer = 0

            'Dim mintConfirmado As Integer
            'Se agregó 31/01/2023
            'Dim dsConfiguracionVacaciones As Data.DataSet = Comunes.fnDsConfiguracionVacaciones()


            fechaServidor = objConsulta.ConsultaFechaServidor()
            'objDsPersona = objConsulta.ConsultaPersona(IdPersona)
            'If objDsPersona.Tables(0).Rows().Count() < 1 Then
            '    Console.Write("El DataSet Esta vacio, intente de nuevo.")
            '    Exit Sub
            'End If

            'Dim TelefonoDestinatario As String = objDsPersona.Tables(0).Rows(0)(0).ToString()
            'Dim CorreoDestinario As String = objDsPersona.Tables(0).Rows(0)(1).ToString()

            For TipoTemporalidad As Integer = 1 To 3
                objDsDatosCorreo = objConsulta.ConsultaCorreos(TipoTemporalidad)
                contador = objDsDatosCorreo.Tables(0).Rows().Count()
                If objDsDatosCorreo.Tables(0).Rows().Count() > 0 Then
                    For Each dr As DataRow In objDsDatosCorreo.Tables(0).Rows
                        strVarForm(1) = dr.Item("CNOMBRE").ToString
                        strVarForm(2) = dr.Item("CCORREO").ToString
                        'Solo va a enviar el correo en caso de que tenga un correo valido
                        If strVarForm(2) <> "" Then
                            strVarForm(3) = dr.Item("CCONCEPTO").ToString
                            strVarForm(4) = dr.Item("CTIPOCONCEPTO").ToString
                            strVarForm(5) = dr.Item("CIMPORTE").ToString
                            strVarForm(6) = dr.Item("CFECHAVENCIMIENTO").ToString
                            strVarForm(7) = dr.Item("CAREA").ToString
                            strVarForm(8) = dr.Item("CCORREOAREA").ToString
                            strVarForm(9) = TipoTemporalidad.ToString
                            'EnviaCorreoCuerpo1(strVarForm)
                        Else
                            'En caso de no tener un correo válido se enviará a la bitácora
                            Comunes.EnviaABitacora(dr.Item("CNOMBRE").ToString, " Sin correo")
                        End If
                    Next
                End If
            Next
            'Exit Sub
        Catch ex As Exception
            Comunes.EnviaABitacora("SubMain()", ex.Message)
        End Try
    End Sub


    Public Sub EnviaCorreoCuerpo1(ByVal vwsdatosCorreo() As String) 'Envia correo de aviso aportación
        Dim strVarForm(8) As String
        Dim strIP As String = Comunes.ConsultaIp()
        Try
            Dim strURLConfirmacionPago As String = Nothing
            Dim dtmFechaActual As Date
            Dim strFechaDeEnvio As String
            Dim objConsulta As New Consultas
            Dim enviacorreo As New MailMessage
            Dim smtp As New System.Net.Mail.SmtpClient("smtp.gmail.com")
            strVarForm = vwsdatosCorreo
            enviacorreo.From = New System.Net.Mail.MailAddress("staff_fenix@acatlan.unam.mx")
            enviacorreo.Priority = MailPriority.High

            dtmFechaActual = Date.Now
            strFechaDeEnvio = Format(dtmFechaActual, "Long Date")

            If strIP = "132.248.180.207" Then
                enviacorreo.To.Add(strVarForm(2))
                enviacorreo.To.Add("liderproyecto_dsi@acatlan.unam.mx")
                strURLConfirmacionPago = "http://132.248.180.50/Fenix/Formularios/ConfirmacionPago.aspx?f=" & strVarForm(7).ToString.Trim
            Else
                enviacorreo.To.Add(strVarForm(2))
                strURLConfirmacionPago = "http://132.248.180.50/Fenix/Formularios/ConfirmacionPago.aspx?f=" & strVarForm(7).ToString.Trim
                'enviacorreo.To.Add("liderproyecto_dsi@acatlan.unam.mx")
            End If

            enviacorreo.SubjectEncoding = System.Text.Encoding.UTF8
            enviacorreo.Subject = "Notificación de prueba - FÉNIX"
            enviacorreo.Body = "<html>" &
                                "<body>" &
                                    "<table align=center>" &
                                        "<tbody>" &
                                            "<tr bgcolor=#002B74>" &
                                                "<td>" &
                                                    "<p align=center><font color=white size=4>Universidad Nacional Autónoma de México</font></p>" &
                                                    "<p align=center><font color=white size=4>Facultad de Estudios Superiores Acatlán</font></p>" &
                                                    "<p align=center><font color=white size=4>Fénix</font></p>" &
                                                "</td>" &
                                            "</tr>" &
                                            "<tr>" &
                                            "<td>" &
                                            "<br/><font color='black'>" &
                                            "Tipo Temporalidad: " & strVarForm(9) &
                                            "<br/><br/>" &
                                            "Desde la Ip: " & strIP &
                                            "<br/><br/>" &
                                            "Hola, " &
                                            "<br/><br/>" &
                                            strVarForm(1) &
                                            "<br/><br/>" &
                                            "Te invito a realizar tu aportación del " & strVarForm(4) & " de " & strVarForm(3) &
                                            "<br/><br/>" &
                                            "<b>Fecha de vencimiento:</b><font color='gray'> " & strVarForm(6) & "</font><br/><br/>" &
                                            "<b>Monto:</b><font color='gray'> $" & strVarForm(5) & "</font><br/><br/>" &
                                            "Lo puedes realizar en el portal de <a href=" & strURLConfirmacionPago & ">FÉNIX</a>. " &
                                            "Es importante que guardes y/o imprimas tu comprobante de aportación para futuras aclaraciones.</font><br/><br/>" &
                                            "Cualquier detalle sobre tu aportación, comunícate con el área: " & strVarForm(7) & ", al correo: " & strVarForm(8) &
                                            "<br/><br/>" &
                                            "<strong>Esta es una cuenta de correo no monitoreada. Su contenido es un mensaje automático generado por el sistema. " &
                                            "Por favor, no respondas a este correo.</strong>" &
                                            "<br/><br/>" &
                                            "" &
                                            "</td>" &
                                            "</tr>" &
                                            "<tr bgcolor=#BB8800>" &
                                                "<td>" &
                                                    "<p align=center><font color=white size=3>Atentamente</font></p>" &
                                                    "<p align=center><font color=white size=3>""POR MI RAZA HABLARÁ EL ESPÍRITU""</font></p>" &
                                                    "<p align=center><font color=white size=3>Acatlán, Estado de México a " & strFechaDeEnvio & "</font></p>" &
                                                "</td>" &
                                            "</tr>" &
                                        "</tbody>" &
                                    "</table>" &
                                "</body>" &
                                "</br>" &
                                "</html>"
            enviacorreo.IsBodyHtml = True
            'que se va a utilizar para enviar los mensajes de correo electrónico.
            smtp.Host = "smtp.gmail.com"
            'Se establece el usuario y contraseña que se utilizarán para enviar el correo
            smtp.Credentials = New System.Net.NetworkCredential("staff_fenix@acatlan.unam.mx", "F4nix2#!3.")
            'Se establece la seguridad
            smtp.Port = 587
            smtp.EnableSsl = True
            'suministrados en las propiedades de la clase MailMessage.
            smtp.Send(enviacorreo)
        Catch ex As Exception
            Console.Write("Error al enviar correo de prueba")
            Comunes.EnviaABitacora("EnviaCorreoCuerpo1", ex.Message)
        End Try
    End Sub
End Module
