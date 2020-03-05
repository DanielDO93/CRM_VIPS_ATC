Imports System.IO
Imports System.Net


Public Class RepDisponibles
    Public Property login() As String
        Get
            Return m_login
        End Get
        Set
            m_login = Value
        End Set
    End Property
    Private m_login As String
    Public Property nombre() As String
        Get
            Return m_nombre
        End Get
        Set
            m_nombre = Value
        End Set
    End Property
    Private m_nombre As String
    Public Property estatus() As String
        Get
            Return m_estatus
        End Get
        Set
            m_estatus = Value
        End Set
    End Property
    Private m_estatus As String
    Public Property activos() As String
        Get
            Return m_activos
        End Get
        Set
            m_activos = Value
        End Set
    End Property
    Private m_activos As String
    Public Property terminados() As String
        Get
            Return m_terminados
        End Get
        Set
            m_terminados = Value
        End Set
    End Property
    Private m_terminados As String
    Public Property distancia() As String
        Get
            Return m_distancia
        End Get
        Set
            m_distancia = Value
        End Set
    End Property
    Private m_distancia As String
    Public Property eta() As String
        Get
            Return m_eta
        End Get
        Set
            m_eta = Value
        End Set
    End Property
    Private m_eta As String
End Class

Public Class GPSDriver
    Public Property ubicacion() As String
        Get
            Return m_ubicacion
        End Get
        Set
            m_ubicacion = Value
        End Set
    End Property
    Private m_ubicacion As String
    Public Property fecha_ubicacion() As String
        Get
            Return m_fecha_ubicacion
        End Get
        Set
            m_fecha_ubicacion = Value
        End Set
    End Property
    Private m_fecha_ubicacion As String
End Class

Public Class Drive

    Public Property id_request() As Integer
        Get
            Return m_id_request
        End Get
        Set
            m_id_request = Value
        End Set
    End Property
    Private m_id_request As Integer
    Public Property status() As String
        Get
            Return m_status
        End Get
        Set
            m_status = Value
        End Set
    End Property
    Private m_status As String
    Public Property driver() As String
        Get
            Return m_driver
        End Get
        Set
            m_driver = Value
        End Set
    End Property
    Private m_driver As String
    Public Property fecha_captura() As String
        Get
            Return m_fecha_captura
        End Get
        Set
            m_fecha_captura = Value
        End Set
    End Property
    Private m_fecha_captura As String
    Public Property fecha_cancelacion() As String
        Get
            Return m_fecha_cancelacion
        End Get
        Set
            m_fecha_cancelacion = Value
        End Set
    End Property
    Private m_fecha_cancelacion As String
    Public Property fecha_aceptado() As String
        Get
            Return m_fecha_aceptado
        End Get
        Set
            m_fecha_aceptado = Value
        End Set
    End Property
    Private m_fecha_aceptado As String
    Public Property fecha_1er_suc() As String
        Get
            Return m_fecha_1er_suc
        End Get
        Set
            m_fecha_1er_suc = Value
        End Set
    End Property
    Private m_fecha_1er_suc As String
    Public Property fecha_sucursal() As String
        Get
            Return m_fecha_sucursal
        End Get
        Set
            m_fecha_sucursal = Value
        End Set
    End Property
    Private m_fecha_sucursal As String
    Public Property fecha_recoleccion() As String
        Get
            Return m_fecha_recoleccion
        End Get
        Set
            m_fecha_recoleccion = Value
        End Set
    End Property
    Private m_fecha_recoleccion As String
    Public Property gps_recoleccion() As String
        Get
            Return m_gps_recoleccion
        End Get
        Set
            m_gps_recoleccion = Value
        End Set
    End Property
    Private m_gps_recoleccion As String
    Public Property fecha_entrega() As String
        Get
            Return m_fecha_entrega
        End Get
        Set
            m_fecha_entrega = Value
        End Set
    End Property
    Private m_fecha_entrega As String
    Public Property gps_entrega() As String
        Get
            Return m_gps_entrega
        End Get
        Set
            m_gps_entrega = Value
        End Set
    End Property
    Private m_gps_entrega As String
    Public Property DLA() As String
        Get
            Return m_DLA
        End Get
        Set
            m_DLA = Value
        End Set
    End Property
    Private m_DLA As String


    Function RequestStatus(Cliente As String, Token As String, Pedido As String) As String


        Dim postData As String = "<?xml version='1.0' encoding='utf-8'?> <Envio xmlns:xsd='http://www.w3.org/2001/XMLSchema'xmlns:xsi = 'http://www.w3.org/2001/XMLSchema-instance'> <Cliente>" & Cliente & "</Cliente><Token>" & Token & "</Token><Pedido>" & Pedido & "</Pedido></Envio>"

        Dim myHttpWebRequest As HttpWebRequest = DirectCast(HttpWebRequest.Create("http://vips.qp.com.mx/delivery_mobile/ws_consultas_completa/index.php"), HttpWebRequest)
        myHttpWebRequest.Method = "POST"

        Dim data As Byte() = Encoding.ASCII.GetBytes(postData)

        myHttpWebRequest.ContentType = "text/plain"
        myHttpWebRequest.ContentLength = data.Length

        Dim requestStream As Stream = myHttpWebRequest.GetRequestStream()
        requestStream.Write(data, 0, data.Length)
        requestStream.Close()

        Dim myHttpWebResponse As HttpWebResponse = DirectCast(myHttpWebRequest.GetResponse(), HttpWebResponse)

        Dim responseStream As Stream = myHttpWebResponse.GetResponseStream()

        Dim myStreamReader As New StreamReader(responseStream, Encoding.[Default])

        Dim pageContent As String = myStreamReader.ReadToEnd()

        myStreamReader.Close()
        responseStream.Close()

        myHttpWebResponse.Close()

        If myHttpWebResponse.StatusCode = 200 Then
            Return pageContent
        Else
            Return False
        End If



    End Function

    Function SendPedido(Cliente As String, Token As String, Pedido As String, Nombre As String, Telefono As String, Calle As String, Colonia As String, Municipio As String, Estado As String, CP As String, TipoPago As String, DetallePedido As String, Monto As String, Comentarios As String, TipoPedido As String, FechaProgramado As String, Tipo As String, Latitud As String, Longitud As String) As String


        Dim postData As String = "<?xml version='1.0' encoding='utf-8'?><Envio xmlns:xsd='http://www.w3.org/2001/XMLSchema'xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'><Cliente>" & Cliente & "</Cliente><Token>" & Token & "</Token><Pedido><ticket>" & Pedido & "</ticket><nombre>" & Nombre & "</nombre><telefono>" & Telefono & "</telefono><calle>" & Calle & "</calle><colonia>" & Colonia & "</colonia><municipio>" & Municipio & "</municipio><estado>" & Estado & "</estado><cp>" & CP & "</cp><tipo_pago>" & TipoPago & "</tipo_pago><detalle_pedido>" & DetallePedido & "</detalle_pedido><monto>" & Monto & "</monto><comentarios>" & Comentarios & "</comentarios><tipo_pedido>" & TipoPedido & "</tipo_pedido><fecha_programado>" & FechaProgramado & "</fecha_programado><tipo>" & Tipo & "</tipo><gps>" & Latitud & "," & Longitud & " </gps></Pedido></Envio>"

        Dim myHttpWebRequest As HttpWebRequest = DirectCast(HttpWebRequest.Create("http://vips.qp.com.mx/delivery_mobile/ws_peticiones/index.php"), HttpWebRequest)
        myHttpWebRequest.Method = "POST"

        Dim data As Byte() = Encoding.ASCII.GetBytes(postData)

        myHttpWebRequest.ContentType = "text/plain"
        myHttpWebRequest.ContentLength = data.Length

        Dim requestStream As Stream = myHttpWebRequest.GetRequestStream()
        requestStream.Write(data, 0, data.Length)
        requestStream.Close()

        Dim myHttpWebResponse As HttpWebResponse = DirectCast(myHttpWebRequest.GetResponse(), HttpWebResponse)

        Dim responseStream As Stream = myHttpWebResponse.GetResponseStream()

        Dim myStreamReader As New StreamReader(responseStream, Encoding.[Default])

        Dim pageContent As String = myStreamReader.ReadToEnd()

        myStreamReader.Close()
        responseStream.Close()

        myHttpWebResponse.Close()
        If myHttpWebResponse.StatusCode = 200 Then
            Return pageContent.Trim
        Else
            Return False
        End If


    End Function

    Function CancelPedido(Cliente As String, Token As String, Pedido As String) As String


        Dim postData As String = "<?xml version='1.0' encoding='utf-8'?><Envio xmlns:xsd='http://www.w3.org/2001/XMLSchema' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'><Cliente>" & Cliente & "</Cliente><Token>" & Token & "</Token><Pedido>" & Pedido & "</Pedido></Envio>"

        Dim myHttpWebRequest As HttpWebRequest = DirectCast(HttpWebRequest.Create("http://vips.qp.com.mx/delivery_mobile/ws_cancelaciones/index.php"), HttpWebRequest)
        myHttpWebRequest.Method = "POST"

        Dim data As Byte() = Encoding.ASCII.GetBytes(postData)

        myHttpWebRequest.ContentType = "text/plain"
        myHttpWebRequest.ContentLength = data.Length

        Dim requestStream As Stream = myHttpWebRequest.GetRequestStream()
        requestStream.Write(data, 0, data.Length)
        requestStream.Close()

        Dim myHttpWebResponse As HttpWebResponse = DirectCast(myHttpWebRequest.GetResponse(), HttpWebResponse)

        Dim responseStream As Stream = myHttpWebResponse.GetResponseStream()

        Dim myStreamReader As New StreamReader(responseStream, Encoding.[Default])

        Dim pageContent As String = myStreamReader.ReadToEnd()

        myStreamReader.Close()
        responseStream.Close()

        myHttpWebResponse.Close()

        Return pageContent.Trim

    End Function

    Function GetGPS(Cliente As String, Token As String, Login As String) As String


        Dim postData As String = "<?xml version='1.0' encoding='utf-16'?><Envio xmlns:xsd='http://www.w3.org/2001/XMLSchema'xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'><Cliente>" & Cliente & "</Cliente><Token>" & Token & "</Token><login>" & Login & "</login></Envio>"

        Dim myHttpWebRequest As HttpWebRequest = DirectCast(HttpWebRequest.Create("http://vips.qp.com.mx/delivery_mobile/ws_consultas_colaborador/index.php"), HttpWebRequest)
        myHttpWebRequest.Method = "POST"

        Dim data As Byte() = Encoding.ASCII.GetBytes(postData)

        myHttpWebRequest.ContentType = "text/plain"
        myHttpWebRequest.ContentLength = data.Length

        Dim requestStream As Stream = myHttpWebRequest.GetRequestStream()
        requestStream.Write(data, 0, data.Length)
        requestStream.Close()

        Dim myHttpWebResponse As HttpWebResponse = DirectCast(myHttpWebRequest.GetResponse(), HttpWebResponse)

        Dim responseStream As Stream = myHttpWebResponse.GetResponseStream()

        Dim myStreamReader As New StreamReader(responseStream, Encoding.[Default])

        Dim pageContent As String = myStreamReader.ReadToEnd()

        myStreamReader.Close()
        responseStream.Close()

        myHttpWebResponse.Close()

        If myHttpWebResponse.StatusCode = 200 Then
            Return pageContent
        Else
            Return False
        End If

    End Function

    Function GetRepDisponibles(Cliente As String, Token As String, Distancia As String) As String


        Dim postData As String = "<?xml version='1.0' encoding='utf-16'?><Envio xmlns:xsd='http://www.w3.org/2001/XMLSchema'xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'><Cliente>" & Cliente & "</Cliente><Token>" & Token & "</Token><Rango>" & Distancia & "</Rango></Envio>"

        Dim myHttpWebRequest As HttpWebRequest = DirectCast(HttpWebRequest.Create("http://vips.qp.com.mx/delivery_mobile/ws_consultas_zona/index.php"), HttpWebRequest)
        myHttpWebRequest.Method = "POST"

        Dim data As Byte() = Encoding.ASCII.GetBytes(postData)

        myHttpWebRequest.ContentType = "text/plain"
        myHttpWebRequest.ContentLength = data.Length

        Dim requestStream As Stream = myHttpWebRequest.GetRequestStream()
        requestStream.Write(data, 0, data.Length)
        requestStream.Close()

        Dim myHttpWebResponse As HttpWebResponse = DirectCast(myHttpWebRequest.GetResponse(), HttpWebResponse)

        Dim responseStream As Stream = myHttpWebResponse.GetResponseStream()

        Dim myStreamReader As New StreamReader(responseStream, Encoding.[Default])

        Dim pageContent As String = myStreamReader.ReadToEnd()

        myStreamReader.Close()
        responseStream.Close()

        myHttpWebResponse.Close()

        If myHttpWebResponse.StatusCode = 200 Then
            Return pageContent
        Else
            Return False
        End If

    End Function

End Class