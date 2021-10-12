﻿Public Class dProveedores
#Region "Atributos"
    Private m_id As Integer
    Private m_nombre As String
    Private m_telefono As String
    Private m_direccion As String
    Private m_email As String
    Private m_email2 As String
    Private m_email3 As String
    Private m_contacto As String
    Private m_otrosdatos As String

#End Region

#Region "Getters y Setters"
    Public Property ID() As Integer
        Get
            Return m_id
        End Get
        Set(ByVal value As Integer)
            m_id = value
        End Set
    End Property
    Public Property NOMBRE() As String
        Get
            Return m_nombre
        End Get
        Set(ByVal value As String)
            m_nombre = value
        End Set
    End Property
    Public Property TELEFONO() As String
        Get
            Return m_telefono
        End Get
        Set(ByVal value As String)
            m_telefono = value
        End Set
    End Property
    Public Property DIRECCION() As String
        Get
            Return m_direccion
        End Get
        Set(ByVal value As String)
            m_direccion = value
        End Set
    End Property
    Public Property EMAIL() As String
        Get
            Return m_email
        End Get
        Set(ByVal value As String)
            m_email = value
        End Set
    End Property
    Public Property EMAIL2() As String
        Get
            Return m_email2
        End Get
        Set(ByVal value As String)
            m_email2 = value
        End Set
    End Property
    Public Property EMAIL3() As String
        Get
            Return m_email3
        End Get
        Set(ByVal value As String)
            m_email3 = value
        End Set
    End Property
    Public Property CONTACTO() As String
        Get
            Return m_contacto
        End Get
        Set(ByVal value As String)
            m_contacto = value
        End Set
    End Property
    Public Property OTROSDATOS() As String
        Get
            Return m_otrosdatos
        End Get
        Set(ByVal value As String)
            m_otrosdatos = value
        End Set
    End Property
#End Region

#Region "Constructores"
    Public Sub New()
        m_id = 0
        m_nombre = ""
        m_telefono = ""
        m_direccion = ""
        m_email = ""
        m_email2 = ""
        m_email3 = ""
        m_contacto = ""
       
    End Sub
    Public Sub New(ByVal id As Integer, ByVal nombre As String, ByVal telefono As String, ByVal direccion As String, ByVal email As String, ByVal email2 As String, ByVal email3 As String, ByVal contacto As String, ByVal otrosdatos As String)
        m_id = id
        m_nombre = nombre
        m_telefono = telefono
        m_direccion = direccion
        m_email = email
        m_email2 = email2
        m_email3 = email3
        m_contacto = contacto
        m_otrosdatos = otrosdatos
    End Sub
#End Region

#Region "Métodos ABM"
    Public Function guardar() As Boolean
        Dim p As New pProveedores
        Return p.guardar(Me)
    End Function
    Public Function modificar() As Boolean
        Dim p As New pProveedores
        Return p.modificar(Me)
    End Function
    Public Function eliminar() As Boolean
        Dim p As New pProveedores
        Return p.eliminar(Me)
    End Function
    Public Function buscar() As dProveedores
        Dim p As New pProveedores
        Return p.buscar(Me)
    End Function
   
    Public Function buscarPorNombre(ByVal pnombre As String) As ArrayList
        Dim s As New pProveedores
        Return s.buscarPorNombre(pnombre)
    End Function
   
#End Region


    Public Overrides Function ToString() As String
        Return m_nombre
    End Function
    Public Function listar() As ArrayList
        Dim p As New pProveedores
        Return p.listar
    End Function
End Class
