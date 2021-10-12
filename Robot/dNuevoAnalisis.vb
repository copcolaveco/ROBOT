Public Class dNuevoAnalisis
#Region "Atributos"
    Private m_id As Long
    Private m_ficha As Long
    Private m_muestra As String
    Private m_detallemuestra As String
    Private m_tipoinforme As Integer
    Private m_analisis As Integer
    Private m_resultado As String
    Private m_resultado2 As String
    Private m_m As Integer
    Private m_metodo As Integer
    Private m_unidad As Integer
    Private m_orden As Integer
    Private m_operador As Integer
    Private m_fechaproceso As String
    Private m_finalizado As Integer
#End Region

#Region "Getters y Setters"
    Public Property ID() As Long
        Get
            Return m_id
        End Get
        Set(ByVal value As Long)
            m_id = value
        End Set
    End Property
    Public Property FICHA() As Long
        Get
            Return m_ficha
        End Get
        Set(ByVal value As Long)
            m_ficha = value
        End Set
    End Property
    Public Property MUESTRA() As String
        Get
            Return m_muestra
        End Get
        Set(ByVal value As String)
            m_muestra = value
        End Set
    End Property
    Public Property DETALLEMUESTRA() As String
        Get
            Return m_detallemuestra
        End Get
        Set(ByVal value As String)
            m_detallemuestra = value
        End Set
    End Property
    Public Property TIPOINFORME() As Integer
        Get
            Return m_tipoinforme
        End Get
        Set(ByVal value As Integer)
            m_tipoinforme = value
        End Set
    End Property
    Public Property ANALISIS() As Integer
        Get
            Return m_analisis
        End Get
        Set(ByVal value As Integer)
            m_analisis = value
        End Set
    End Property
    Public Property RESULTADO() As String
        Get
            Return m_resultado
        End Get
        Set(ByVal value As String)
            m_resultado = value
        End Set
    End Property
    Public Property RESULTADO2() As String
        Get
            Return m_resultado2
        End Get
        Set(ByVal value As String)
            m_resultado2 = value
        End Set
    End Property
    Public Property M() As Integer
        Get
            Return m_m
        End Get
        Set(ByVal value As Integer)
            m_m = value
        End Set
    End Property
    Public Property METODO() As Integer
        Get
            Return m_metodo
        End Get
        Set(ByVal value As Integer)
            m_metodo = value
        End Set
    End Property
    Public Property UNIDAD() As Integer
        Get
            Return m_unidad
        End Get
        Set(ByVal value As Integer)
            m_unidad = value
        End Set
    End Property
    Public Property ORDEN() As Integer
        Get
            Return m_orden
        End Get
        Set(ByVal value As Integer)
            m_orden = value
        End Set
    End Property
    Public Property OPERADOR() As Integer
        Get
            Return m_operador
        End Get
        Set(ByVal value As Integer)
            m_operador = value
        End Set
    End Property
    Public Property FECHAPROCESO() As String
        Get
            Return m_fechaproceso
        End Get
        Set(ByVal value As String)
            m_fechaproceso = value
        End Set
    End Property
    Public Property FINALIZADO() As Integer
        Get
            Return m_finalizado
        End Get
        Set(ByVal value As Integer)
            m_finalizado = value
        End Set
    End Property
#End Region

#Region "Constructores"
    Public Sub New()
        m_id = 0
        m_ficha = 0
        m_muestra = ""
        m_detallemuestra = ""
        m_tipoinforme = 0
        m_analisis = 0
        m_resultado = ""
        m_resultado2 = ""
        m_m = 0
        m_metodo = 0
        m_unidad = 0
        m_orden = 0
        m_operador = 0
        m_fechaproceso = ""
        m_finalizado = 0
    End Sub
    Public Sub New(ByVal id As Long, ByVal ficha As Long, ByVal muestra As String, ByVal detallemuestra As String, ByVal tipoinforme As Integer, ByVal analisis As Integer, ByVal resultado As String, ByVal resultado2 As String, ByVal m As Integer, ByVal metodo As Integer, ByVal unidad As Integer, ByVal orden As Integer, ByVal operador As Integer, ByVal fechaproceso As String, ByVal finalizado As Integer)
        m_id = id
        m_ficha = ficha
        m_muestra = muestra
        m_detallemuestra = detallemuestra
        m_tipoinforme = tipoinforme
        m_analisis = analisis
        m_resultado = resultado
        m_resultado2 = resultado2
        m_m = m
        m_metodo = metodo
        m_unidad = unidad
        m_orden = orden
        m_operador = operador
        m_fechaproceso = fechaproceso
        m_finalizado = finalizado
    End Sub
#End Region

#Region "Métodos ABM"
    Public Function guardar() As Boolean
        Dim e As New pNuevoAnalisis
        Return e.guardar(Me)
    End Function
    Public Function modificar() As Boolean
        Dim e As New pNuevoAnalisis
        Return e.modificar(Me)
    End Function
    Public Function marcarfinalizado() As Boolean
        Dim e As New pNuevoAnalisis
        Return e.marcarfinalizado(Me)
    End Function
    'Public Function asignaroperador() As Boolean
    '    Dim e As New pNuevoAnalisis
    '    Return e.asignaroperador(Me)
    'End Function
    Public Function actualizar_resultado() As Boolean
        Dim e As New pNuevoAnalisis
        Return e.actualizar_resultado(Me)
    End Function
    Public Function actualizar_metodo() As Boolean
        Dim e As New pNuevoAnalisis
        Return e.actualizar_metodo(Me)
    End Function
    Public Function actualizar_detalle() As Boolean
        Dim e As New pNuevoAnalisis
        Return e.actualizar_detalle(Me)
    End Function
    Public Function actualizar_fecha() As Boolean
        Dim e As New pNuevoAnalisis
        Return e.actualizar_fecha(Me)
    End Function
    Public Function actualizar_unidad() As Boolean
        Dim e As New pNuevoAnalisis
        Return e.actualizar_unidad(Me)
    End Function
    Public Function eliminar() As Boolean
        Dim e As New pNuevoAnalisis
        Return e.eliminar(Me)
    End Function
  
    Public Function buscar() As dNuevoAnalisis
        Dim e As New pNuevoAnalisis
        Return e.buscar(Me)
    End Function
    Public Function buscarxficha() As dNuevoAnalisis
        Dim e As New pNuevoAnalisis
        Return e.buscarxficha(Me)
    End Function
    Public Function buscarxfichaxmuestra() As dNuevoAnalisis
        Dim e As New pNuevoAnalisis
        Return e.buscarxfichaxmuestra(Me)
    End Function
    Public Function buscarrepetidas() As dNuevoAnalisis
        Dim e As New pNuevoAnalisis
        Return e.buscarrepetidas(Me)
    End Function
#End Region

    Public Overrides Function ToString() As String

        Return m_muestra

    End Function
    Public Function listar() As ArrayList
        Dim e As New pNuevoAnalisis
        Return e.listar
    End Function
    Public Function listarfichas() As ArrayList
        Dim e As New pNuevoAnalisis
        Return e.listarfichas
    End Function
    Public Function listarfichasnuevas(ByVal tipoinf As Integer) As ArrayList
        Dim e As New pNuevoAnalisis
        Return e.listarfichasnuevas(tipoinf)
    End Function
    Public Function listarporid(ByVal texto As Long) As ArrayList
        Dim e As New pNuevoAnalisis
        Return e.listarporid(texto)
    End Function
    Public Function listarporficha(ByVal texto As Long) As ArrayList
        Dim e As New pNuevoAnalisis
        Return e.listarporficha(texto)
    End Function
    Public Function listarporfichamuestra(ByVal texto As Long) As ArrayList
        Dim e As New pNuevoAnalisis
        Return e.listarporfichamuestra(texto)
    End Function
    Public Function listarporficha2(ByVal texto As Long) As ArrayList
        Dim e As New pNuevoAnalisis
        Return e.listarporficha2(texto)
    End Function
    Public Function listarporficha3(ByVal texto As Long) As ArrayList
        Dim e As New pNuevoAnalisis
        Return e.listarporficha3(texto)
    End Function
    Public Function listarpormuestra(ByVal ficha As Long, ByVal muestra As String) As ArrayList
        Dim e As New pNuevoAnalisis
        Return e.listarpormuestra(ficha, muestra)
    End Function
    Public Function listardistintosanalisis(ByVal ficha As Long) As ArrayList
        Dim e As New pNuevoAnalisis
        Return e.listardistintosanalisis(ficha)
    End Function
    Public Function listarxanalisis(ByVal id As Integer) As ArrayList
        Dim e As New pNuevoAnalisis
        Return e.listarxanalisis(id)
    End Function
    Public Function listarxfichaxanalisis(ByVal ficha As Long, ByVal id As Integer) As ArrayList
        Dim e As New pNuevoAnalisis
        Return e.listarxfichaxanalisis(ficha, id)
    End Function
End Class
