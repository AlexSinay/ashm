Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.IO
Imports System.Net
Imports System.Net.NetworkInformation
Imports MySql.Data.MySqlClient
Imports System.Text



Public Class FuncionesBasicas

    Public Shared Function Version(Optional ByVal Win As Boolean = True) As String
        Return "SistemSHM.DLL Versión " & My.Application.Info.Version.ToString & " para Uso Libre"
    End Function

    Public Shared Function PrimerDiaDelMes(ByVal Fecha As Date) As Date
        PrimerDiaDelMes = DateSerial(Year(Fecha), Month(Fecha), 1)
    End Function

    Public Shared Function UltimoDiaDelMes(ByVal Fecha As Date) As Date
        UltimoDiaDelMes = DateSerial(Year(Fecha), Month(Fecha) + 1, 0)
    End Function

    Public Shared Function Espacios(Optional ByVal Texto As String = "") As String
        Texto = RTrim(LTrim(Replace(Texto, "'", "''")))
        Return Texto
    End Function

    Public Shared Function Tildes(Texto As String) As String
        If Len(Texto) > 0 Then
            Dim Encw1252 As Encoding = Encoding.GetEncoding("windows-1252")
            Dim EncUTF8 As Encoding = Encoding.GetEncoding("utf-8")
            Dim Str As String = Texto
            Str = Str.Replace("Ã ", "Ã ")
            Str = Encw1252.GetString(Encoding.Convert(EncUTF8, Encw1252, Encoding.Default.GetBytes(Str)))
            Return Str
        Else
            Return ""
        End If
    End Function

    Public Shared Function AgregarEspacios(Optional ByVal Texto As String = "", Optional ByVal Largo As Integer = 0, Optional ByVal Horientacion As Integer = 0) As String
        'Horientacion = 0 los espacios se cargarn a la derecha
        'Horientacion = 1 los espacios se cargarn a la izquierda
        Dim VI As Integer = 0
        Dim Espacios As String = ""
        Try
            For VI = 1 To Largo
                Espacios = Espacios & " "
            Next
            If Horientacion = 0 Then
                Texto = Microsoft.VisualBasic.Left(Texto & Espacios, Largo)
            Else
                Texto = Microsoft.VisualBasic.Right(Espacios & Texto, Largo)
            End If
        Catch ex As Exception
            Texto = ""
        End Try
        Return Texto
    End Function

    Public Shared Function ValidarMail(Correo As String) As Boolean
        If InStr(Correo, "@@", CompareMethod.Text) > 0 Then
            Return False
        Else
            If InStr(Correo, "@", CompareMethod.Text) > 0 Then
                If InStr(Correo, ".", CompareMethod.Text) > 0 Then
                    Return True
                End If
            End If
        End If
        Return False
    End Function

    Public Shared Function OrientarTexto(Optional ByVal Texto As String = "", Optional ByVal Longitud As Integer = 2, Optional ByVal Orientacion As Integer = 1) As String
        Dim I As Integer
        Dim L As Integer
        Dim E As Integer
        Dim textoT As String
        Texto = Trim(Texto)
        Select Case Orientacion
            Case 1  'IZQUIERDA
                'SE ELIMINAN LOS ESPACIOS
            Case 2  'CENTRADO   
                Texto = RTrim(LTrim(Texto))
                L = Len(Texto)
                If L < Longitud Then
                    E = (Longitud - L) / 2
                    textoT = ""
                    For I = 1 To E
                        textoT = textoT & " "
                    Next
                    Texto = textoT & Texto
                End If
            Case 3  'DERECHA
                For I = 1 To Longitud
                    Texto = " " & Texto
                Next
                Texto = Right(RTrim(Texto), Longitud)
        End Select
        Return Texto
    End Function

    Public Shared Function NombreMes(Optional ByVal NumMes As Integer = 1) As String
        Select Case NumMes
            Case 1
                Return "ENERO"
            Case 2
                Return "FEBRERO"
            Case 3
                Return "MARZO"
            Case 4
                Return "ABRIL"
            Case 5
                Return "MAYO"
            Case 6
                Return "JUNIO"
            Case 7
                Return "JULIO"
            Case 8
                Return "AGOSTO"
            Case 9
                Return "SEPTIEMBRE"
            Case 10
                Return "OCTUBRE"
            Case 11
                Return "NOVIEMBRE"
            Case 12
                Return "DICIEMBRE"
        End Select
        Return ""
    End Function

    Public Shared Function HoyEs(Optional ByVal Fecha As String = "01/01/2021") As String
        Dim Res As String
        Res = CInt(Mid(Fecha, 1, 2)).ToString & " de " & iNombreMes(Mid(Fecha, 4, 2)) & " del " & Mid(Fecha, 7, 4)
        Return Res
    End Function

    Public Shared Function NombreDiaSemana(Optional ByVal NumDia As Integer = 0) As String
        Select Case NumDia
            Case 0
                Return "Domingo"
            Case 1
                Return "Lunes"
            Case 2
                Return "Martes"
            Case 3
                Return "Miercoles"
            Case 4
                Return "Jueves"
            Case 5
                Return "Viernes"
            Case 6
                Return "Sabado"
        End Select
        Return ""
    End Function

    Public Shared Function ValidarPuntoDecimal(Numero As String) As String
        Dim I As Integer = 0
        Dim J As Integer = 0
        Dim TX As String = LTrim(Numero)
        For I = 1 To Len(TX)
            If Mid(TX, I, 1) = "." Then
                J = I
                Exit For
            End If
        Next
        If J = 0 Then
            TX = TX & ".00"
        Else
            I = Len(TX) - J
            If I = 1 Then
                TX = TX & "0"
            End If
        End If
        Return TX
    End Function

    Public Shared Function EsNumero(Numero As String) As Boolean
        Dim PAst As Integer = 0
        Dim Punto As Integer = 0
        For I = 1 To Len(Numero)
            If Mid(Numero, I, 1) = "0" Or Mid(Numero, I, 1) = "1" Or Mid(Numero, I, 1) = "2" Or Mid(Numero, I, 1) = "3" Or Mid(Numero, I, 1) = "4" Or Mid(Numero, I, 1) = "5" Or Mid(Numero, I, 1) = "6" Or Mid(Numero, I, 1) = "7" Or Mid(Numero, I, 1) = "8" Or Mid(Numero, I, 1) = "9" Or Mid(Numero, I, 1) = "." Then
                'ES UN NUMERO
                If Mid(Numero, I, 1) = "." Then Punto = Punto + 1
                If Punto > 1 Then   'TIENE MAS DE 1 PUNTO DECIMAL
                    PAst = 1
                    Exit For
                End If
            Else
                PAst = 1
                Exit For
            End If
        Next
        If PAst = 1 Then Return False Else Return True
    End Function

    Public Shared Function ArreglarNumero(Numero As String) As String
        ArreglarNumero = ""
        Numero = LTrim(RTrim(Numero))
        Dim I As Integer
        Dim L As Integer
        Dim J As Integer
        Dim ParteDecimal As String
        Dim ParteEntera As String
        L = Len(Numero)
        If L > 0 Then
            'buscar punto decimal
            ParteDecimal = ".00"
            J = 0
            For I = 1 To L
                If Mid(Numero, I, 1) = "." Then
                    J = I
                End If
            Next
            If J = 0 Then
                ParteEntera = Numero
            Else
                ParteEntera = Mid(Numero, 1, J - 1)
                ParteDecimal = Left(Mid(Numero, J, 3) & "00", 3)
            End If
            Select Case Len(ParteEntera)
                Case 4
                    ParteEntera = Left(ParteEntera, 1) & "," & Right(ParteEntera, 3)
                Case 5
                    ParteEntera = Left(ParteEntera, 2) & "," & Right(ParteEntera, 3)
                Case 6
                    ParteEntera = Left(ParteEntera, 3) & "," & Right(ParteEntera, 3)
            End Select
            Return ParteEntera & ParteDecimal
        End If
    End Function

    Public Shared Function GenerarAleatorio(ByVal Largo As Integer) As String
        Dim guidResult As String = System.Guid.NewGuid().ToString()
        guidResult = guidResult.Replace("-", String.Empty)
        Return Left(guidResult, Largo)
    End Function

    Public Shared Function PasswordAleatorio(Optional ByVal Longitud As Integer = 2, Optional ByVal Minusculas As Boolean = True, Optional ByVal Mayusculas As Boolean = False, Optional ByVal Numeros As Boolean = False, Optional ByVal Caracteres As Boolean = False)
        Dim Cadena As String = ""
        Try
            Dim I As Integer = 0
            Dim Posicion As Integer = 0
            Dim Matriz As String = ""
            If Minusculas = True Then Matriz = Matriz & "abcdefghijklmnopqrstuvwxyz"
            If Mayusculas = True Then Matriz = Matriz & "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
            If Numeros = True Then Matriz = Matriz & "1234567890"
            If Caracteres = True Then Matriz = Matriz & "[]{}!¡¿?#$%&/()"
            If Longitud > 0 Then
                Static Generador As System.Random = New System.Random()
                For I = 1 To Longitud
                    Posicion = Generador.Next(1, Len(Matriz))
                    Cadena = Cadena & Mid(Matriz, Posicion, 1)
                Next
            End If
        Catch ex As Exception

        End Try
        Return Cadena
    End Function

    Public Shared Function Encriptar(Texto As String, Factor As Integer) As String
        Dim Resultado As String = ""
        Dim I As Integer = 0
        Dim J As Integer = 0
        If Len(Texto) > 0 Then
            For I = 1 To Len(Texto)
                J = Asc(Mid(Texto, I, 1)) + Factor
                Resultado = Resultado & Chr(J)
            Next
        End If
        Return Resultado
    End Function

    Public Shared Function DesEncriptar(Texto As String, Factor As Integer) As String
        Dim Resultado As String = ""
        Dim I As Integer = 0
        Dim J As Integer = 0
        If Len(Texto) > 0 Then
            For I = 1 To Len(Texto)
                J = Asc(Mid(Texto, I, 1)) - Factor
                Resultado = Resultado & Chr(J)
            Next
        End If
        Return Resultado
    End Function

    Public Shared Function Retardo(Segundos As Integer) As Boolean
        Dim SegFinal As DateTime = DateAdd(DateInterval.Second, Segundos, Now)
        Dim DifSegundos As Integer = DateDiff(DateInterval.Second, Now, SegFinal)
        Do While DifSegundos >= 0
            DifSegundos = DateDiff(DateInterval.Second, Now, SegFinal)
        Loop
        Return True
    End Function

    Public Shared Function SeparaMillares(Valor As String) As String
        Valor = RTrim(LTrim(Valor))
        Dim Negativo As Boolean = False
        If InStr(Valor, "-", CompareMethod.Text) > 0 Then
            Negativo = True
            Valor = Replace(Valor, "-", "")
        End If
        Select Case Len(Valor)
            Case 1 To 3
                'no es suficientemente grande para separa
            Case 4 To 6
                Valor = "      " & Valor
                Valor = Microsoft.VisualBasic.Right(Valor, 6)
                Valor = Mid(Valor, 1, 3) & "," & Mid(Valor, 4, 3)
                Valor = RTrim(LTrim(Valor))
            Case 7 To 9
                Valor = "         " & Valor
                Valor = Microsoft.VisualBasic.Right(Valor, 9)
                Valor = Mid(Valor, 1, 3) & "," & Mid(Valor, 4, 3) & "," & Mid(Valor, 7, 3)
                Valor = RTrim(LTrim(Valor))
        End Select
        If Negativo = True Then Valor = "-" & Valor
        Return Valor
    End Function

End Class

Public Class Archivo

    Public Shared Function GuardarLog(Arc As String, Optional ByVal Texto As String = "", Optional ByVal Fecha As Boolean = True) As Boolean
        If Fecha = True Then Texto = Now.ToShortDateString & " " & Now.ToShortTimeString & " " & Texto
        Return iTXTAgregar(Arc, Texto)
    End Function


    Public Shared Function TXTCrear(Arc As String, Optional ByVal ErrMSG As Boolean = False) As Boolean
        Return iTXTCrear(Arc)
    End Function

    Public Shared Function TXTAgregar(Arc As String, Texto As String, Optional ByVal ErrMSG As Boolean = False) As Boolean
        Return iTXTAgregar(Arc, Texto)
    End Function

    Public Shared Function TXTLeer(Arc As String) As String
        Dim Texto As String
        Try
            Dim sr As New System.IO.StreamReader(Arc)
            Texto = sr.ReadToEnd()
            sr.Close()
            Return LTrim(RTrim(Texto))
        Catch ex As Exception

        End Try
        Return ""
    End Function

End Class

Public Class BasesDeDatos

    Public Shared Function SQLActualizar(CConexion As String, txSQL As String) As String
        Dim TipoBD As Integer = 0   '1=OLE   2=SQL   3=MySQL
        Dim TipoBDtx As String = ""
        Dim Res As String = "OK"
        Try
            If InStr(UCase(CConexion), "PROVIDER=") > 0 And InStr(UCase(CConexion), "DATA SOURCE=") Then
                TipoBD = 1
                TipoBDtx = "ACCESS"
            Else
                If InStr(UCase(CConexion), "USER ID=") > 0 And InStr(UCase(CConexion), "DATA SOURCE=") Then
                    TipoBD = 2
                    TipoBDtx = "SQL"
                Else
                    If InStr(UCase(CConexion), "DATABASE=") > 0 And InStr(UCase(CConexion), "USERID=") Then
                        TipoBD = 3
                        TipoBDtx = "MySQL"
                    Else
                        TipoBD = 0
                    End If
                End If
            End If

            Dim Loca As Integer = 0

            Select Case TipoBD
                Case 1  'ACCESS
                    Dim SQLComando
                    Using conn As New OleDbConnection(CConexion)
                        conn.Open()
                        SQLComando = New OleDbCommand(txSQL)
                        SQLComando.Connection = conn
                        SQLComando.ExecuteNonQuery()
                        conn.Close()
                    End Using
                    OleDbConnection.ReleaseObjectPool()
                    Return True

                Case 2  'MICROSOFT SQL
                    Dim sql As SqlCommand
                    Using Conn As New SqlClient.SqlConnection(CConexion)
                        Conn.Open()
                        sql = New SqlCommand(txSQL, Conn)
                        Loca = sql.ExecuteNonQuery()
                        Conn.Close()
                    End Using
                    SqlConnection.ClearAllPools()
                    Return True

                Case 3  'MYSQL
                    Dim SQLcmd As MySqlCommand
                    Using Conn As New MySql.Data.MySqlClient.MySqlConnection(CConexion)
                        Conn.Open()
                        SQLcmd = New MySqlCommand(txSQL, Conn)
                        SQLcmd.ExecuteNonQuery()
                        Conn.Close()
                    End Using
                    MySqlConnection.ClearAllPools()
                    Return True

                Case Else
                    Res = "ERROR_BD: TIPO DE CONEXION NO RECONOCIDA"
            End Select

        Catch ex As Exception
            Res = "ERROR: " & ex.Message
        End Try
        Return Res
    End Function

    Public Shared Function SQLLeer(CConexion As String, txSQL As String, Optional ByVal ErrMsg As Boolean = False, Optional ByVal TextoAlerta As String = "") As String
        Dim TipoBD As Integer = 0   '1=OLE   2=SQL   3=MySQL
        Dim TipoBDtx As String = ""
        Dim Res As String = ""

        Try
            If InStr(UCase(CConexion), "PROVIDER=") > 0 And InStr(UCase(CConexion), "DATA SOURCE=") Then
                TipoBD = 1
                TipoBDtx = "ACCESS"
            Else
                If InStr(UCase(CConexion), "USER ID=") > 0 And InStr(UCase(CConexion), "DATA SOURCE=") Then
                    TipoBD = 2
                    TipoBDtx = "SQL"
                Else
                    If InStr(UCase(CConexion), "DATABASE=") > 0 And InStr(UCase(CConexion), "USERID=") Then
                        TipoBD = 3
                        TipoBDtx = "MySQL"
                    Else
                        TipoBD = 0
                    End If
                End If
            End If

            Select Case TipoBD
                Case 1  'ACCESS
                    Dim SQLComando
                    Using conn As New OleDbConnection(CConexion)
                        conn.Open()
                        SQLComando = New OleDbCommand(txSQL)
                        SQLComando.Connection = conn
                        Dim RS As OleDbDataReader = SQLComando.ExecuteReader()
                        While RS.Read()
                            If IsDBNull(RS("dato")) Then Res = "" Else Res = RTrim(RS("dato"))
                        End While
                        conn.Close()
                    End Using
                    OleDbConnection.ReleaseObjectPool()
                    Return Res

                Case 2  'MICROSOFT SQL
                    Dim SQL As SqlCommand
                    Dim RS As SqlDataReader
                    Using Conn As New SqlClient.SqlConnection(CConexion)
                        Conn.Open()
                        SQL = New SqlCommand(txSQL)
                        SQL.Connection = Conn
                        RS = SQL.ExecuteReader
                        While RS.Read
                            If IsDBNull(RS("dato")) Then Return "" Else Return RTrim(RS("dato"))
                        End While
                        Conn.Close()
                    End Using
                    SqlConnection.ClearAllPools()
                    Return Res

                Case 3  'MYSQL
                    Dim SQLcmd As MySqlCommand
                    Dim RS As MySqlDataReader
                    Using Conn As New MySql.Data.MySqlClient.MySqlConnection(CConexion)
                        Conn.Open()
                        SQLcmd = New MySqlCommand(txSQL, Conn)
                        RS = SQLcmd.ExecuteReader
                        While RS.Read
                            If IsDBNull(RS("dato")) Then Return "" Else Return RTrim(RS("dato"))
                        End While
                        Conn.Close()
                    End Using
                    MySqlConnection.ClearAllPools()
                    Return Res

                Case Else
                    Res = "ERROR: [TIPO DE CONEXION NO RECONOCIDA]"
            End Select

        Catch ex As Exception
            Res = "ERROR: " & ex.Message
        End Try
        Return Res
    End Function

    Public Shared Function SQLValidar(CConexion As String, Optional ByVal ErrMsg As Boolean = False, Optional ByVal TextoAlerta As String = "") As String
        Dim TipoBD As Integer = 0   '1=OLE   2=SQL   3=MySQL
        Dim TipoBDtx As String = ""
        Dim Res As Boolean = False

        Try
            If InStr(UCase(CConexion), "PROVIDER=") > 0 And InStr(UCase(CConexion), "DATA SOURCE=") Then
                TipoBD = 1
                TipoBDtx = "ACCESS"
            Else
                If InStr(UCase(CConexion), "USER ID=") > 0 And InStr(UCase(CConexion), "DATA SOURCE=") Then
                    TipoBD = 2
                    TipoBDtx = "SQL"
                Else
                    If InStr(UCase(CConexion), "DATABASE=") > 0 And InStr(UCase(CConexion), "USERID=") Then
                        TipoBD = 3
                        TipoBDtx = "MySQL"
                    Else
                        TipoBD = 0
                    End If
                End If
            End If

            Select Case TipoBD
                Case 1  'ACCESS
                    Using conn As New OleDbConnection(CConexion)
                        conn.Open()
                        Res = True
                        conn.Close()
                    End Using
                    OleDbConnection.ReleaseObjectPool()

                Case 2  'MICROSOFT SQL
                    Using Conn As New SqlClient.SqlConnection(CConexion)
                        Conn.Open()
                        Res = True
                        Conn.Close()
                    End Using
                    SqlConnection.ClearAllPools()

                Case 3  'MYSQL
                    Using Conn As New MySql.Data.MySqlClient.MySqlConnection(CConexion)
                        Conn.Open()
                        Res = True
                        Conn.Close()
                    End Using
                    MySqlConnection.ClearAllPools()

                Case Else
                    Res = "ERROR: [TIPO DE CONEXION NO RECONOCIDA]"
            End Select

        Catch ex As Exception
            Res = "ERROR: " & ex.Message
        End Try
        Return Res
    End Function

    Public Shared Function CrearCCSQL(Instancia As String, BaseDatos As String, Usuario As String, Clave As String) As String
        Instancia = LTrim(RTrim(Instancia))
        BaseDatos = LTrim(RTrim(BaseDatos))
        Usuario = LTrim(RTrim(Usuario))
        Clave = LTrim(RTrim(Clave))
        Dim CC As String = "Data Source=@svr@;Initial Catalog=@bd@;Persist Security Info=True;User ID=@u@;Password=@p@"
        CC = Replace(CC, "@svr@", Instancia)
        CC = Replace(CC, "@bd@", BaseDatos)
        CC = Replace(CC, "@u@", Usuario)
        CC = Replace(CC, "@p@", Clave)
        Return CC
    End Function

    Public Shared Function CrearCCMySQL(Servidor As String, BaseDatos As String, Usuario As String, Clave As String) As String
        Servidor = LTrim(RTrim(Servidor))
        BaseDatos = LTrim(RTrim(BaseDatos))
        Usuario = LTrim(RTrim(Usuario))
        Clave = LTrim(RTrim(Clave))
        Dim CC As String = "DATABASE=@bd@;SERVER=@svr@;USERID=@u@;PASSWORD=@p@"
        CC = Replace(CC, "@svr@", Servidor)
        CC = Replace(CC, "@bd@", BaseDatos)
        CC = Replace(CC, "@u@", Usuario)
        CC = Replace(CC, "@p@", Clave)
        Return CC
    End Function

    Public Shared Function CrearCCAccess(BaseDatos As String, Usuario As String, Clave As String) As String
        BaseDatos = LTrim(RTrim(BaseDatos))
        Usuario = LTrim(RTrim(Usuario))
        Clave = LTrim(RTrim(Clave))
        Dim CC As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=@bd@;Jet OLEDB:Database Password=@p@"
        CC = Replace(CC, "@bd@", BaseDatos)
        CC = Replace(CC, "@u@", Usuario)
        CC = Replace(CC, "@p@", Clave)
        Return CC
    End Function

End Class

Public Class Red

    Public Shared Function HacerPing(ByVal IP As String) As Boolean
        Try
            If My.Computer.Network.IsAvailable = False Then Return False
            Return My.Computer.Network.Ping(IP)
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Shared Function ObtenerPC() As String
        Try
            Return RTrim(LCase(System.Net.Dns.GetHostName()))
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Public Shared Function ObtenerSesion() As String
        Try
            Return RTrim(LCase(My.User.Name))
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Public Shared Function ObtenerIP() As String
        Try
            Dim GetIPv4Address
            GetIPv4Address = String.Empty
            Dim strHostName As String = System.Net.Dns.GetHostName()
            Dim iphe As System.Net.IPHostEntry = System.Net.Dns.GetHostEntry(strHostName)
            For Each ipheal As System.Net.IPAddress In iphe.AddressList
                If ipheal.AddressFamily = System.Net.Sockets.AddressFamily.InterNetwork Then GetIPv4Address = ipheal.ToString()
            Next
            Return GetIPv4Address
        Catch ex As Exception
            Return ""
        End Try
    End Function

    Public Shared Function ObtenerIPPublica() As String
        Try
            Return iObtenerWeb("http://apisgratis.com/api/ip/")
        Catch ex As Exception

        End Try
        Return ""
    End Function

    Public Shared Function ObtenerHTTPWeb(Web As String, Optional ByVal ErrMSG As Boolean = False) As String
        Dim WebResultado As String = ""
        Web = Trim(Web)
        If Len(Web) > 4 Then
            Try
                Dim req As HttpWebRequest = WebRequest.Create(Web)
                Dim res As HttpWebResponse = req.GetResponse()
                Dim Stream As Stream = res.GetResponseStream()
                Dim sr As StreamReader = New StreamReader(Stream)
                Return RTrim(sr.ReadToEnd())
            Catch ex As Exception
                WebResultado = "ERROR: " & ex.Message
            End Try
        End If
        Return WebResultado
    End Function

    Public Shared Function ObtenerPuertaEnlace() As String
        Dim Res As String = ""
        Try
            Dim NetworkAdapters() As NetworkInterface = NetworkInterface.GetAllNetworkInterfaces
            Dim myAdapterProps As IPInterfaceProperties = Nothing
            Dim myGateways As GatewayIPAddressInformationCollection = Nothing
            For Each adapter As NetworkInterface In NetworkAdapters
                myAdapterProps = adapter.GetIPProperties
                myGateways = myAdapterProps.GatewayAddresses
                For Each Gateway As GatewayIPAddressInformation In myGateways
                    Return Gateway.Address.ToString
                Next
            Next
        Catch ex As Exception
            Res = "ERROR: " & ex.Message
        End Try
        Return Res
    End Function

    Public Shared Function FTPEnviar(ArcOrigen As String, TrayDestino As String, Usu As String, Psw As String, Optional ByVal ErrMSG As Boolean = False) As String
        Dim Res As String = ""
        Try
            If File.Exists(ArcOrigen) = True Then
                My.Computer.Network.UploadFile(ArcOrigen, TrayDestino, Usu, Psw)
                'Return True
                Res = "OK"
            Else
                'Return False
                Res = "Archivo Origen No Localizado."
            End If
        Catch ex As Exception
            Res = "ERROR: " & ex.Message
            '            Return False
        End Try
        Return Res
    End Function

    Public Shared Function FTPDescargar(ByVal TrayFTP As String, ByVal Arc As String, ByVal TrayLocal As String, Usu As String, Psw As String, Optional ByVal ErrMSG As Boolean = False) As String
        Dim Res As String = ""
        Try
            Dim ArcLocal As String = ""
            TrayLocal = Replace((TrayLocal & "\"), "\\", "\")
            TrayLocal = RTrim(TrayLocal)
            If TrayLocal = "-" Then ArcLocal = "c:\software\" & Arc Else ArcLocal = TrayLocal & "\" & Arc
            If System.IO.Directory.Exists(TrayLocal) = False Then System.IO.Directory.CreateDirectory(TrayLocal)
            ArcLocal = Replace(ArcLocal, "\\", "\")
            Dim ArcFTP As String = TrayFTP & Arc
            If System.IO.File.Exists(ArcLocal) = True Then System.IO.File.Delete(ArcLocal)
            My.Computer.Network.DownloadFile(ArcFTP, ArcLocal, Usu, Psw)
            ' Return True
            Res = "OK"
        Catch ex As Exception
            Res = "ERROR: " & ex.Message
        End Try
        Return Res
    End Function

    Public Shared Function WEBDescargar(Link As String, Arc As String, Traylocal As String, Optional ByVal ErrMSG As Boolean = False) As String
        Dim Res As String = ""
        Try
            Dim ArcLocal As String = ""
            Traylocal = Replace((Traylocal & "\"), "\\", "\")
            Traylocal = RTrim(Traylocal)
            If Traylocal = "-" Then ArcLocal = "c:\windows\temp\" & Arc Else ArcLocal = Traylocal & "\" & Arc
            If System.IO.Directory.Exists(Traylocal) = False Then System.IO.Directory.CreateDirectory(Traylocal)
            ArcLocal = Replace(ArcLocal, "\\", "\")
            If System.IO.File.Exists(ArcLocal) = True Then System.IO.File.Delete(ArcLocal)
            My.Computer.Network.DownloadFile(Link, ArcLocal)
            Res = "OK"
        Catch ex As Exception
            Res = "ERROR: " & ex.Message
        End Try
        Return Res
    End Function


End Class




