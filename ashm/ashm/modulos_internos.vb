Imports System.Net
Imports System.IO
Imports System.Net.NetworkInformation


Module modulos_internos

    Public Function iObtenerWeb(URL As String) As String
        Dim WebResultado As String = ""
        URL = Trim(URL)
        If Len(URL) > 4 Then
            Try
                Dim req As HttpWebRequest = WebRequest.Create(URL)
                Dim res As HttpWebResponse = req.GetResponse()
                Dim Stream As Stream = res.GetResponseStream()
                Dim sr As StreamReader = New StreamReader(Stream)
                WebResultado = RTrim(sr.ReadToEnd())
            Catch ex As Exception

            End Try
        End If
        Return WebResultado
    End Function

    Public Function iObtenerIP(adapterProperties As IPInterfaceProperties) As String
        Dim Resultado As String = ""
        Dim uniCast As UnicastIPAddressInformationCollection = adapterProperties.UnicastAddresses
        If uniCast IsNot Nothing Then
            For Each uni As UnicastIPAddressInformation In uniCast
                If Len(uni.Address.ToString) > 3 And Len(uni.Address.ToString) < 20 Then Resultado = uni.Address.ToString
            Next
        End If
        Return Resultado
    End Function

    Public Function iNombreMes(Optional ByVal NumMes As String = "01") As String
        Select Case CInt(NumMes)
            Case 1
                Return "Enero"
            Case 2
                Return "Febrero"
            Case 3
                Return "Marzo"
            Case 4
                Return "Abril"
            Case 5
                Return "Mayo"
            Case 6
                Return "Junio"
            Case 7
                Return "Julio"
            Case 8
                Return "Agosto"
            Case 9
                Return "Septiembre"
            Case 10
                Return "Octubre"
            Case 11
                Return "Noviembre"
            Case 12
                Return "Diciembre"
        End Select
        Return ""
    End Function

    Public Function iTXTCrear(Arc As String) As Boolean
        Try
            If File.Exists(Arc) = True Then File.Delete(Arc)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function iTXTAgregar(Arc As String, Texto As String) As Boolean
        Try
            Dim ruta As String = Arc
            Dim escritor As StreamWriter
            escritor = File.AppendText(ruta)
            escritor.Write(Texto & vbCrLf)
            escritor.Flush()
            escritor.Close()
            Return True
        Catch ex As Exception

        End Try
        Return False
    End Function

End Module
