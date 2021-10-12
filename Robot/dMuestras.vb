﻿Public Class dMuestras
#Region "Atributos"
    Private m_id As Integer
    Private m_nombre As String
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
#End Region

#Region "Constructores"
    Public Sub New()
        m_id = 0
        m_nombre = ""
    End Sub
    Public Sub New(ByVal id As Integer, ByVal nombre As String)
        m_id = id
        m_nombre = nombre
    End Sub
#End Region

#Region "Métodos ABM"
    Public Function guardar() As Boolean
        Dim p As New pMuestras
        Return p.guardar(Me)
    End Function
    Public Function modificar() As Boolean
        Dim p As New pMuestras
        Return p.modificar(Me)
    End Function
    Public Function eliminar() As Boolean
        Dim p As New pMuestras
        Return p.eliminar(Me)
    End Function
    Public Function buscar() As dMuestras
        Dim p As New pMuestras
        Return p.buscar(Me)
    End Function
#End Region

    Public Overrides Function ToString() As String
        Return m_nombre
    End Function

    Public Function listar() As ArrayList
        Dim p As New pMuestras
        Return p.listar
    End Function
End Class
