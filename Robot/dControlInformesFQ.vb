﻿Public Class dControlInformesFQ
#Region "Atributos"
    Private m_id As Long
    Private m_fechacontrol As String
    Private m_ficha As Long
    Private m_fecha As String
    Private m_tipo As Integer
    Private m_resultado As Integer
    Private m_coincide As Integer
    Private m_om As Integer
    Private m_nc As Integer
    Private m_observaciones As String
    Private m_controlador As Integer
    Private m_controlado As Integer
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
    Public Property FECHACONTROL() As String
        Get
            Return m_fechacontrol
        End Get
        Set(ByVal value As String)
            m_fechacontrol = value
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
    Public Property FECHA() As String
        Get
            Return m_fecha
        End Get
        Set(ByVal value As String)
            m_fecha = value
        End Set
    End Property
    Public Property TIPO() As Integer
        Get
            Return m_tipo
        End Get
        Set(ByVal value As Integer)
            m_tipo = value
        End Set
    End Property
    Public Property RESULTADO() As Integer
        Get
            Return m_resultado
        End Get
        Set(ByVal value As Integer)
            m_resultado = value
        End Set
    End Property
    Public Property COINCIDE() As Integer
        Get
            Return m_coincide
        End Get
        Set(ByVal value As Integer)
            m_coincide = value
        End Set
    End Property
    Public Property OM() As Integer
        Get
            Return m_om
        End Get
        Set(ByVal value As Integer)
            m_om = value
        End Set
    End Property
    Public Property NC() As Integer
        Get
            Return m_nc
        End Get
        Set(ByVal value As Integer)
            m_nc = value
        End Set
    End Property
    Public Property OBSERVACIONES() As String
        Get
            Return m_observaciones
        End Get
        Set(ByVal value As String)
            m_observaciones = value
        End Set
    End Property
    Public Property CONTROLADOR() As Integer
        Get
            Return m_controlador
        End Get
        Set(ByVal value As Integer)
            m_controlador = value
        End Set
    End Property
    Public Property CONTROLADO() As Integer
        Get
            Return m_controlado
        End Get
        Set(ByVal value As Integer)
            m_controlado = value
        End Set
    End Property

#End Region

#Region "Constructores"
    Public Sub New()
        m_id = 0
        m_fechacontrol = Now
        m_ficha = 0
        m_fecha = Now
        m_tipo = 0
        m_resultado = 0
        m_coincide = 0
        m_om = 0
        m_nc = 0
        m_observaciones = ""
        m_controlador = 0
        m_controlado = 0
    End Sub
    Public Sub New(ByVal id As Long, ByVal fechacontrol As String, ByVal ficha As Long, ByVal fecha As String, ByVal tipo As Integer, ByVal resultado As Integer, ByVal coincide As Integer, ByVal om As Integer, ByVal nc As Integer, ByVal observaciones As String, ByVal controlador As Integer, ByVal controlado As Integer)
        m_id = id
        m_fechacontrol = fechacontrol
        m_ficha = ficha
        m_fecha = fecha
        m_tipo = tipo
        m_resultado = resultado
        m_coincide = coincide
        m_om = om
        m_nc = nc
        m_observaciones = observaciones
        m_controlador = controlador
        m_controlado = controlado
    End Sub
#End Region

#Region "Métodos ABM"
    Public Function guardar() As Boolean
        Dim c As New pControlInformesFQ
        Return c.guardar(Me)
    End Function
    Public Function modificar() As Boolean
        Dim c As New pControlInformesFQ
        Return c.modificar(Me)
    End Function
    Public Function eliminar() As Boolean
        Dim c As New pControlInformesFQ
        Return c.eliminar(Me)
    End Function
    Public Function buscar() As dControlInformesFQ
        Dim c As New pcontrolinformesfq
        Return c.buscar(Me)
    End Function
    Public Function buscarxficha() As dControlInformesFQ
        Dim c As New pControlInformesFQ
        Return c.buscarxficha(Me)
    End Function

#End Region

    Public Overrides Function tostring() As String
        Return m_ficha
    End Function
    Public Function listar() As ArrayList
        Dim c As New pcontrolinformesfq
        Return c.listar
    End Function
    Public Function listarxfecha(ByVal desde As String, ByVal hasta As String) As ArrayList
        Dim c As New pcontrolinformesfq
        Return c.listarxfecha(desde, hasta)
    End Function
    Public Function listarxtipoxfecha(ByVal tipo As Integer, ByVal desde As String, ByVal hasta As String) As ArrayList
        Dim c As New pControlInformesFQ
        Return c.listarxtipoxfecha(tipo, desde, hasta)
    End Function
    Public Function listarxfechanc(ByVal desde As String, ByVal hasta As String) As ArrayList
        Dim c As New pcontrolinformesfq
        Return c.listarxfechanc(desde, hasta)
    End Function

    Public Function listarxfechaom(ByVal desde As String, ByVal hasta As String) As ArrayList
        Dim c As New pcontrolinformesfq
        Return c.listarxfechaom(desde, hasta)
    End Function
    Public Function marcarresultado() As Boolean
        Dim ci As New pControlInformesFQ
        Return ci.marcarresultado(Me)
    End Function
    Public Function desmarcarresultado() As Boolean
        Dim ci As New pControlInformesFQ
        Return ci.desmarcarresultado(Me)
    End Function
    Public Function marcarcoincide() As Boolean
        Dim ci As New pControlInformesFQ
        Return ci.marcarcoincide(Me)
    End Function
    Public Function desmarcarcoincide() As Boolean
        Dim ci As New pControlInformesFQ
        Return ci.desmarcarcoincide(Me)
    End Function
    Public Function marcarom() As Boolean
        Dim ci As New pControlInformesFQ
        Return ci.marcarom(Me)
    End Function
    Public Function desmarcarom() As Boolean
        Dim ci As New pControlInformesFQ
        Return ci.desmarcarom(Me)
    End Function
    Public Function marcarnc() As Boolean
        Dim ci As New pControlInformesFQ
        Return ci.marcarnc(Me)
    End Function
    Public Function desmarcarnc() As Boolean
        Dim ci As New pControlInformesFQ
        Return ci.desmarcarnc(Me)
    End Function
    Public Function marcarcontrolada() As Boolean
        Dim ci As New pControlInformesFQ
        Return ci.marcarcontrolada(Me)
    End Function
    Public Function guardarobservaciones(ByVal obs As String) As Boolean
        Dim ci As New pControlInformesFQ
        Return ci.guardarobservaciones(Me, obs)
    End Function
End Class
