Imports System
Imports System.IO
Imports System.Net
Imports System.Text
Imports System.Data.SqlClient
Imports Newtonsoft.Json.Linq

Module Module1
    Sub Main()
        'Dim msg As String = "'<texto voz=" + Chr(34) + "Carlos" + Chr(34) + ">'  Esta es una prueba basica de envio  '</texto>'"
        Console.WriteLine("El Proceso para enviar bases de REMINDER Masivo inició a las : {0:HH:mm:ss.fff}", DateTime.Now)
        'Console.ReadLine()
        'Abrimos la conexión hacia el servidor
        Dim sConnectionString1 As String = "User ID=usrGestion;Password=G3st10nW0rk$;Initial Catalog=DB_GESTION;Data Source=172.18.2.207; MultipleActiveResultSets = True"
        Dim objConn1 As New SqlConnection(sConnectionString1)
        objConn1.Open()

        Dim sSQL1 As String = "exec [spEvaluaMarcacionReminderMasivo2Proveedor1Envio1]"
        Dim objCmd1 As New SqlCommand(sSQL1, objConn1)
        Dim myReader1 As SqlDataReader = objCmd1.ExecuteReader()

        If myReader1.HasRows Then
            myReader1.Read()
            Dim nIdcontrolbasesreminder = myReader1(0).ToString
            Dim cte = myReader1(1).ToString
            Dim encpwd = myReader1(2).ToString
            Dim email = myReader1(3).ToString
            Dim json = myReader1(4).ToString
            Dim mtipo = myReader1(5).ToString
            Dim mensajeXML = myReader1(6).ToString
            Dim tipoDestino = myReader1(7).ToString
            Dim archivo = myReader1(8).ToString
            Dim TitulosBase = myReader1(9).ToString
            Dim nombreEnvio = myReader1(10).ToString
            Dim fechaInicio = myReader1(11).ToString
            Dim fechaFin = myReader1(12).ToString
            Dim horaInicio = myReader1(13).ToString
            Dim horaFin = myReader1(14).ToString
            Dim intentos = myReader1(15).ToString
            Dim IdsListasNegras = myReader1(16).ToString

            Dim Request As HttpWebRequest = DirectCast(WebRequest.Create(New Uri("https://api1.calixtaondemand.com/Controller.php/__a/sms.extsend.remote.sa")), HttpWebRequest)
            Request.Method = "POST"
            Request.KeepAlive = True
            Dim boundary As String = "-------------------------" + DateTime.Now.Ticks.ToString("x")
            Dim NewLine As String = Environment.NewLine
            Dim header As String = NewLine & "--" + boundary + NewLine
            Dim footer As String = NewLine & "--" + boundary + NewLine
            Request.ContentType = String.Format("multipart/form-data; boundary={0}", boundary)
            Dim contents As StringBuilder = New StringBuilder()
            contents.Append(NewLine)
            contents.Append(header)
            contents.Append("Content-Type: text/plain; charset=us-ascii" & NewLine)
            contents.Append("Content-Transfer-Encoding: 7bit" & NewLine)
            contents.Append("Content-Disposition: form-data; name=""cte""" & NewLine)
            contents.Append(NewLine)
            contents.Append(cte)
            contents.Append(header)
            contents.Append("Content-Type: text/plain; charset=us-ascii" & NewLine)
            contents.Append("Content-Transfer-Encoding: 7bit" & NewLine)
            contents.Append("Content-Disposition: form-data; name=""encpwd""" & NewLine)
            contents.Append(NewLine)
            contents.Append(encpwd)
            contents.Append(header)
            contents.Append("Content-Type: text/plain; charset=us-ascii" & NewLine)
            contents.Append("Content-Transfer-Encoding: 7bit" & NewLine)
            contents.Append("Content-Disposition: form-data; name=""email""" & NewLine)
            contents.Append(NewLine)
            contents.Append(email)
            contents.Append(header)
            contents.Append("Content-Type: text/plain; charset=us-ascii" & NewLine)
            contents.Append("Content-Transfer-Encoding: 7bit" & NewLine)
            contents.Append("Content-Disposition: form-data; name=""json""" & NewLine)
            contents.Append(NewLine)
            contents.Append(json)
            contents.Append(header)
            contents.Append("Content-Type: text/plain; charset=us-ascii" & NewLine)
            contents.Append("Content-Transfer-Encoding: 7bit" & NewLine)
            contents.Append("Content-Disposition: form-data; name=""mtipo""" & NewLine)
            contents.Append(NewLine)
            contents.Append(mtipo)
            contents.Append(header)
            contents.Append("Content-Type: text/plain; charset=Cp1252" & NewLine)
            contents.Append("Content-Transfer-Encoding: quoted-printable" & NewLine)
            contents.Append("Content-Disposition: form-data; name=""mensaje""" & NewLine)
            contents.Append(NewLine)
            contents.Append(mensajeXML)
            contents.Append(header)
            contents.Append("Content-Type: text/plain; charset=us-ascii" & NewLine)
            contents.Append("Content-Transfer-Encoding: 7bit" & NewLine)
            contents.Append("Content-Disposition: form-data; name=""tipoDestino""" & NewLine)
            contents.Append(NewLine)
            contents.Append(tipoDestino)
            contents.Append(header)
            contents.Append("Content-Type: application/octet-stream; name=Baseprueba.csv" & NewLine)
            contents.Append("Content-Transfer-Encoding: binary" & NewLine)
            contents.Append("Content-Disposition: form-data; name=""archivo""; filename=""" + archivo + """" & NewLine)
            contents.Append(NewLine)
            'En caso de simulacion agregamos encabezados de envio de ejemplo basicos
            'contents.Append("TELEFONO,ID_REGISTRO,NOMBRE")
            'En caso de simulacion agregamos encabezados de envio de ejemplo basicos
            'En caso de produccion enviamos encabezados de produccion obtenidos desde la base de datos
            contents.Append(TitulosBase)
            'Terminamos de agregar encabezados de produccion.
            'En caso de simulacion agregamos registros con mi telefono:
            'contents.Append(NewLine)
            'contents.Append("5544712827,00001,OSCAR")
            'contents.Append(NewLine)
            'contents.Append("5544712827,00002,GUSTAVO")
            'contents.Append(NewLine)
            'contents.Append("5574862706,00003,KARINA")
            ''contents.Append(NewLine)
            'contents.Append("5544712827,00003")
            'Terminamos de agregar telefonos en caso de simulacion
            'En caso de Produccion obtenemos la lista de registros que contiene la base y agregamos registro por registro al listado
            Console.WriteLine(contents.ToString())
            Dim sSQL2 As String = "exec spObtenerRegistrosReminderWebMasivo2Proveedor1Envio1 @nIdControlBasesReminder"
            Dim objCmd2 As New SqlCommand(sSQL2, objConn1)
            objCmd2.Parameters.Add("@nIdControlBasesReminder", SqlDbType.Int)
            objCmd2.Parameters.Item("@nIdControlBasesReminder").Value = nIdcontrolbasesreminder
            Dim myReader2 As SqlDataReader = objCmd2.ExecuteReader()
            Do While myReader2.Read()
                contents.Append(NewLine)
                contents.Append(myReader2(0).ToString)
            Loop
            'Terminamos de agregar la lista de registros en caso de produccion
            contents.Append(header)
            contents.Append("Content-Type: text/plain; charset=us-ascii" & NewLine)
            contents.Append("Content-Transfer-Encoding: 7bit" & NewLine)
            contents.Append("Content-Disposition: form-data; name=""nombreEnvio""" & NewLine)
            contents.Append(NewLine)
            'En caso de simulacion enviamos un nombre de envio de prueba
            'contents.Append("Prueba1")
            'En caso de simulacion enviamos un nombre de envio de prueba
            'En caso de produccion enviamos el nombre de envio recuperado desde la base
            contents.Append(nombreEnvio)
            'En caso de produccion enviamos el nombre de envio recuperado desde la base
            contents.Append(header)
            contents.Append("Content-Type: text/plain; charset=us-ascii" & NewLine)
            contents.Append("Content-Transfer-Encoding: 7bit" & NewLine)
            contents.Append("Content-Disposition: form-data; name=""fechaInicio""" & NewLine)
            contents.Append(NewLine)
            contents.Append(fechaInicio)
            contents.Append(header)
            contents.Append("Content-Type: text/plain; charset=us-ascii" & NewLine)
            contents.Append("Content-Transfer-Encoding: 7bit" & NewLine)
            contents.Append("Content-Disposition: form-data; name=""fechaFin""" & NewLine)
            contents.Append(NewLine)
            contents.Append(fechaFin)
            contents.Append(header)
            contents.Append("Content-Type: text/plain; charset=us-ascii" & NewLine)
            contents.Append("Content-Transfer-Encoding: 7bit" & NewLine)
            contents.Append("Content-Disposition: form-data; name=""horaInicio""" & NewLine)
            contents.Append(NewLine)
            'contents.Append("0.3333")
            contents.Append(horaInicio)
            contents.Append(header)
            contents.Append("Content-Type: text/plain; charset=us-ascii" & NewLine)
            contents.Append("Content-Transfer-Encoding: 7bit" & NewLine)
            contents.Append("Content-Disposition: form-data; name=""horaFin""" & NewLine)
            contents.Append(NewLine)
            'contents.Append("0.916")
            contents.Append(horaFin)
            contents.Append(header)
            contents.Append("Content-Type: text/plain; charset=us-ascii" & NewLine)
            contents.Append("Content-Transfer-Encoding: 7bit" & NewLine)
            contents.Append("Content-Disposition: form-data; name=""intentos""" & NewLine)
            contents.Append(NewLine)
            contents.Append(intentos)
            contents.Append(header)
            contents.Append("Content-Type: text/plain; charset=us-ascii" & NewLine)
            contents.Append("Content-Transfer-Encoding: 7bit" & NewLine)
            contents.Append("Content-Disposition: form-data; name=""listasNegras""" & NewLine)
            contents.Append(NewLine)
            contents.Append(IdsListasNegras)

            Console.WriteLine(contents.ToString())
            Dim BodyBytes As Byte() = Encoding.UTF8.GetBytes(contents.ToString())
            Dim footerBytes As Byte() = Encoding.UTF8.GetBytes(footer)
            Request.ContentLength = BodyBytes.Length + footerBytes.Length
            Dim requestStream As Stream = Request.GetRequestStream()
            requestStream.Write(BodyBytes, 0, BodyBytes.Length)
            requestStream.Write(footerBytes, 0, footerBytes.Length)
            requestStream.Flush()
            requestStream.Close()

            Console.WriteLine("Obtenemos la respuesta del Web Service en formato Json:" & vbCrLf)
            Dim ret As String = New StreamReader(Request.GetResponse.GetResponseStream).ReadToEnd
            Console.WriteLine(ret)


            Dim ser As JObject = JObject.Parse(ret)
            Dim data As List(Of JToken) = ser.Children().ToList
            Dim ResultadoEnvio As String
            ResultadoEnvio = ""

            For Each item As JProperty In data
                item.CreateReader()
                If item.Name = "resultadoTransaccion" Or item.Name = "error" Then
                    ResultadoEnvio = item.Value
                End If
            Next
            Console.WriteLine(ResultadoEnvio)
            '            Console.ReadKey()
            Dim sSQL3 As String = "exec spActualizaEstatusSubidaReminderMasivo2Proveedor1Envio1 @nIdControlBasesReminder,@ResultadoSubida"
            Dim objCmd3 As New SqlCommand(sSQL3, objConn1)
            objCmd3.Parameters.Add("@nIdControlBasesReminder", SqlDbType.Int)
            objCmd3.Parameters.Item("@nIdControlBasesReminder").Value = nIdcontrolbasesreminder
            objCmd3.Parameters.Add("@ResultadoSubida", SqlDbType.VarChar, 100)
            objCmd3.Parameters.Item("@ResultadoSubida").Value = ResultadoEnvio
            objCmd3.CommandTimeout = 10
            objCmd3.ExecuteNonQuery()


            myReader1.Close()
            '            myReader2.Close()
            objConn1.Close()
        Else
            Console.WriteLine("NO HAY BASES PENDIENTES DE ENVIAR DE REMINDER")
        End If
    End Sub


End Module
