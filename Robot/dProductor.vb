﻿Public Class dProductor
#Region "Atributos"
    Private m_id As Long
    Private m_nombre As String
    Private m_email1 As String
    Private m_email2 As String
    Private m_email3 As String
    Private m_envio As String
    Private m_usuario_web As String
    Private m_razon_social As String
    Private m_telefono_2 As String
    Private m_telefono_3 As String
    Private m_celular_1 As String
    Private m_celular_2 As String
    Private m_celular_3 As String
    Private m_rut As String
    Private m_codigofigaro As String
    Private m_tipousuario As Integer
    Private m_direccion As String
    Private m_telefono As String
    Private m_fax As String
    Private m_dicose As String
    Private m_iddepartamento As Integer
    Private m_idlocalidad As Integer
    Private m_tecnico As Long
    Private m_tecnico2 As Long
    Private m_tecnico3 As Long
    Private m_idagencia As Integer
    Private m_contrato As Integer
    Private m_socio As Integer
    Private m_nousar As Integer
    Private m_moroso As Integer
    Private m_contado As Integer
    Private m_caravanas As Integer
    Private m_prolesa As Integer

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
    Public Property NOMBRE() As String
        Get
            Return m_nombre
        End Get
        Set(ByVal value As String)
            m_nombre = value
        End Set
    End Property
    Public Property EMAIL1() As String
        Get
            Return m_email1
        End Get
        Set(ByVal value As String)
            m_email1 = value
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
    Public Property ENVIO() As String
        Get
            Return m_envio
        End Get
        Set(ByVal value As String)
            m_envio = value
        End Set
    End Property
    Public Property USUARIO_WEB() As String
        Get
            Return m_usuario_web
        End Get
        Set(ByVal value As String)
            m_usuario_web = value
        End Set
    End Property
    Public Property RAZON_SOCIAL() As String
        Get
            Return m_razon_social
        End Get
        Set(ByVal value As String)
            m_razon_social = value
        End Set
    End Property
    Public Property TELEFONO_2() As String
        Get
            Return m_telefono_2
        End Get
        Set(ByVal value As String)
            m_telefono_2 = value
        End Set
    End Property
    Public Property TELEFONO_3() As String
        Get
            Return m_telefono_3
        End Get
        Set(ByVal value As String)
            m_telefono_3 = value
        End Set
    End Property
    Public Property CELULAR_1() As String
        Get
            Return m_celular_1
        End Get
        Set(ByVal value As String)
            m_celular_1 = value
        End Set
    End Property
    Public Property CELULAR_2() As String
        Get
            Return m_celular_2
        End Get
        Set(ByVal value As String)
            m_celular_2 = value
        End Set
    End Property
    Public Property CELULAR_3() As String
        Get
            Return m_celular_3
        End Get
        Set(ByVal value As String)
            m_celular_3 = value
        End Set
    End Property
    Public Property RUT() As String
        Get
            Return m_rut
        End Get
        Set(ByVal value As String)
            m_rut = value
        End Set
    End Property
    Public Property CODIGOFIGARO() As String
        Get
            Return m_codigofigaro
        End Get
        Set(ByVal value As String)
            m_codigofigaro = value
        End Set
    End Property
    Public Property TIPOUSUARIO() As String
        Get
            Return m_tipousuario
        End Get
        Set(ByVal value As String)
            m_tipousuario = value
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
    Public Property TELEFONO() As String
        Get
            Return m_telefono
        End Get
        Set(ByVal value As String)
            m_telefono = value
        End Set
    End Property
    Public Property FAX() As String
        Get
            Return m_fax
        End Get
        Set(ByVal value As String)
            m_fax = value
        End Set
    End Property
    Public Property DICOSE() As String
        Get
            Return m_dicose
        End Get
        Set(ByVal value As String)
            m_dicose = value
        End Set
    End Property
    Public Property IDDEPARTAMENTO() As Integer
        Get
            Return m_iddepartamento
        End Get
        Set(ByVal value As Integer)
            m_iddepartamento = value
        End Set
    End Property

    Public Property IDLOCALIDAD() As Integer
        Get
            Return m_idlocalidad
        End Get
        Set(ByVal value As Integer)
            m_idlocalidad = value
        End Set
    End Property
    Public Property TECNICO() As Long
        Get
            Return m_tecnico
        End Get
        Set(ByVal value As Long)
            m_tecnico = value
        End Set
    End Property
    Public Property TECNICO2() As Long
        Get
            Return m_tecnico2
        End Get
        Set(ByVal value As Long)
            m_tecnico2 = value
        End Set
    End Property
    Public Property TECNICO3() As Long
        Get
            Return m_tecnico3
        End Get
        Set(ByVal value As Long)
            m_tecnico3 = value
        End Set
    End Property

    Public Property IDAGENCIA() As Integer
        Get
            Return m_idagencia
        End Get
        Set(ByVal value As Integer)
            m_idagencia = value
        End Set
    End Property
    Public Property CONTRATO() As Integer
        Get
            Return m_contrato
        End Get
        Set(ByVal value As Integer)
            m_contrato = value
        End Set
    End Property
    Public Property SOCIO() As Integer
        Get
            Return m_socio
        End Get
        Set(ByVal value As Integer)
            m_socio = value
        End Set
    End Property
    Public Property NOUSAR() As Integer
        Get
            Return m_nousar
        End Get
        Set(ByVal value As Integer)
            m_nousar = value
        End Set
    End Property
    Public Property MOROSO() As Integer
        Get
            Return m_moroso
        End Get
        Set(ByVal value As Integer)
            m_moroso = value
        End Set
    End Property
    Public Property CONTADO() As Integer
        Get
            Return m_contado
        End Get
        Set(ByVal value As Integer)
            m_contado = value
        End Set
    End Property
    Public Property CARAVANAS() As Integer
        Get
            Return m_caravanas
        End Get
        Set(ByVal value As Integer)
            m_caravanas = value
        End Set
    End Property
    Public Property PROLESA() As Integer
        Get
            Return m_prolesa
        End Get
        Set(ByVal value As Integer)
            m_prolesa = value
        End Set
    End Property
#End Region

#Region "Constructores"
    Public Sub New()
        m_id = 0
        m_nombre = ""
        m_email1 = ""
        m_email2 = ""
        m_email3 = ""
        m_envio = ""
        m_usuario_web = ""
        m_razon_social = ""
        m_telefono_2 = ""
        m_telefono_3 = ""
        m_celular_1 = ""
        m_celular_2 = ""
        m_celular_3 = ""
        m_rut = ""
        m_codigofigaro = ""
        m_tipousuario = 0
        m_direccion = ""
        m_telefono = ""
        m_fax = ""
        m_dicose = ""
        m_iddepartamento = 0
        m_idlocalidad = 0
        m_tecnico = 0
        m_tecnico2 = 0
        m_tecnico3 = 0
        m_idagencia = 0
        m_contrato = 0
        m_socio = 0
        m_nousar = 0
        m_moroso = 0
        m_contado = 0
        m_caravanas = 0
        m_prolesa = 0
    End Sub
    Public Sub New(ByVal id As Long, ByVal nombre As String, ByVal email1 As String, ByVal email2 As String, ByVal email3 As String, ByVal envio As String, ByVal usuario_web As String, ByVal razon_social As String, ByVal telefono_2 As String, ByVal telefono_3 As String, ByVal celular_1 As String, ByVal celular_2 As String, ByVal celular_3 As String, ByVal rut As String, ByVal codigofigaro As String, ByVal tipousuario As Integer, ByVal direccion As String, ByVal telefono As String, ByVal fax As String, ByVal dicose As String, ByVal iddepartamento As Integer, ByVal idlocalidad As Integer, ByVal tecnico As Long, ByVal tecnico2 As Long, ByVal tecnico3 As Long, ByVal matricula As String, ByVal idempresa As Integer, ByVal idagencia As Integer, ByVal contrato As Integer, ByVal socio As Integer, ByVal nousar As Integer, ByVal moroso As Integer, ByVal contado As Integer, ByVal caravanas As Integer, ByVal prolesa As Integer)
        m_id = id
        m_nombre = nombre
        m_email1 = email1
        m_email2 = email2
        m_email3 = email3
        m_envio = envio
        m_usuario_web = usuario_web
        m_razon_social = razon_social
        m_telefono_2 = telefono_2
        m_telefono_3 = telefono_3
        m_celular_1 = celular_1
        m_celular_2 = celular_3
        m_rut = rut
        m_codigofigaro = codigofigaro
        m_tipousuario = tipousuario
        m_direccion = direccion
        m_telefono = telefono
        m_fax = fax
        m_dicose = dicose
        m_iddepartamento = iddepartamento
        m_idlocalidad = idlocalidad
        m_tecnico = tecnico
        m_tecnico2 = tecnico2
        m_tecnico3 = tecnico3
        m_idagencia = idagencia
        m_contrato = contrato
        m_socio = socio
        m_nousar = nousar
        m_moroso = moroso
        m_contado = contado
        m_caravanas = caravanas
        m_prolesa = prolesa
    End Sub
#End Region

#Region "Métodos ABM"
    Public Function guardar() As Boolean
        Dim p As New pProductor
        Return p.guardar(Me)
    End Function
    Public Function modificar() As Boolean
        Dim p As New pProductor
        Return p.modificar(Me)
    End Function
    Public Function eliminar() As Boolean
        Dim p As New pProductor
        Return p.eliminar(Me)
    End Function
    Public Function buscar() As dProductor
        Dim p As New pProductor
        Return p.buscar(Me)
    End Function
    Public Function buscarPorNombreTodos(ByVal pnombre As String) As ArrayList
        Dim s As New pProductor
        Return s.buscarPorNombreTodos(pnombre)
    End Function
    Public Function buscarPorNombre(ByVal pnombre As String) As ArrayList
        Dim s As New pProductor
        Return s.buscarPorNombre(pnombre)
    End Function
    Public Function buscarPorNombreEmpresa(ByVal pnombre As String) As ArrayList
        Dim s As New pProductor
        Return s.buscarPorNombreEmpresa(pnombre)
    End Function
    Public Function buscarPorUsuarioWeb() As dProductor
        Dim p As New pProductor
        Return p.buscarPorUsuarioWeb(Me)
    End Function

    Public Function buscarPorId(ByVal pid As Long) As ArrayList
        Dim s As New pProductor
        Return s.buscarPorId(pid)
    End Function
#End Region


    Public Overrides Function ToString() As String
        Return m_nombre
    End Function
    Public Function listar() As ArrayList
        Dim p As New pProductor
        Return p.listar
    End Function
    Public Function listartodos() As ArrayList
        Dim p As New pProductor
        Return p.listartodos
    End Function
    Public Function listarempresa() As ArrayList
        Dim p As New pProductor
        Return p.listarempresa
    End Function
    Public Function listarxtecnico(ByVal id As Long) As ArrayList
        Dim p As New pProductor
        Return p.listarxtecnico(id)
    End Function
    Public Function actualizardireccion(ByVal idProductor As Integer, ByVal direnvio As String) As Boolean
        Dim p As New pProductor
        Return p.actualizardireccion(idProductor, direnvio)
    End Function
    Public Function actualizartelefono(ByVal idProductor As Integer, ByVal tel As String) As Boolean
        Dim p As New pProductor
        Return p.actualizartelefono(idProductor, tel)
    End Function
    Public Function actualizartecnico(ByVal idProductor As Integer, ByVal tec As Long) As Boolean
        Dim p As New pProductor
        Return p.actualizartecnico(idProductor, tec)
    End Function
    Public Function actualizartecnico2(ByVal idProductor As Integer, ByVal tec2 As Long) As Boolean
        Dim p As New pProductor
        Return p.actualizartecnico2(idProductor, tec2)
    End Function
    Public Function actualizartecnico3(ByVal idProductor As Integer, ByVal tec3 As Long) As Boolean
        Dim p As New pProductor
        Return p.actualizartecnico3(idProductor, tec3)
    End Function
    Public Function actualizaragencia(ByVal idProductor As Integer, ByVal age As Integer) As Boolean
        Dim p As New pProductor
        Return p.actualizaragencia(idProductor, age)
    End Function
    Public Function actualizarmail(ByVal idProductor As Integer, ByVal mail As String) As Boolean
        Dim p As New pProductor
        Return p.actualizarmail(idProductor, mail)
    End Function
    Public Function actualizardicose(ByVal idProductor As Integer, ByVal dicose As String) As Boolean
        Dim p As New pProductor
        Return p.actualizardicose(idProductor, dicose)
    End Function
    Public Function marcarmoroso(ByVal idfigaro As String) As Boolean
        Dim p As New pProductor
        Return p.marcarmoroso(idfigaro)
    End Function
    Public Function marcarmoroso2(ByVal productor As Long) As Boolean
        Dim p As New pProductor
        Return p.marcarmoroso2(productor)
    End Function
    Public Function desmarcarmoroso2(ByVal productor As Long) As Boolean
        Dim p As New pProductor
        Return p.desmarcarmoroso2(productor)
    End Function
    Public Function desmarcarmoroso(ByVal idfigaro As String) As Boolean
        Dim p As New pProductor
        Return p.desmarcarmoroso(idfigaro)
    End Function
    Public Function desmarcarmorosos() As Boolean
        Dim p As New pProductor
        Return p.desmarcarmorosos()
    End Function
End Class

