Imports System.Net
Imports System.Net.FtpWebRequest
Imports System.Text
Imports System.Security.Cryptography
Public Class FormClientes
    Private carpeta As Long = 0
    Private clienteweb_com As String = ""
    Private clienteweb_uy As String = ""
    Private password_cifrado As String
    Private idnuevocliente As Long = 0
#Region "Atributos"
    Private _usuario As dUsuario
    Public Property Usuario() As dUsuario
        Get
            Return _usuario
        End Get
        Set(ByVal value As dUsuario)
            _usuario = value
        End Set
    End Property
#End Region
#Region "Constructores"
    Public Sub New(ByVal u As dUsuario, ByVal idclientenuevo As Long)

        ' Llamada necesaria para el Diseñador de Windows Forms.
        InitializeComponent()

        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().
        Usuario = u
        cargarLista()
        cargarComboLocalidad()
        cargarComboDepartamento()
        cargarcomboProlesa()
        cargarComboGiro()
        cargarComboTipoUsuario()
        cargarComboTecnicos()
        cargarComboAgencia()
        limpiar()
        idnuevocliente = idclientenuevo
        If idnuevocliente <> 0 Then
            cargarnuevocliente()
        End If
    End Sub

#End Region
    Private Sub cargarnuevocliente()
        Dim np As New dNuevocliente
        np.ID = idnuevocliente
        np = np.buscar
        If Not np Is Nothing Then
            TextNombre.Text = np.NOMBRE
            TextEmail.Text = np.EMAIL
            TextEnvio.Text = np.DIRECCIONENVIO
            TextRazonSocial.Text = np.RAZON_SOCIAL
            TextCelular.Text = np.CELULAR
            TextRut.Text = np.RUT
            ComboTipoUsuario.SelectedItem = Nothing
            Dim tu As dTipoUsuario
            For Each tu In ComboTipoUsuario.Items
                If tu.ID = np.TIPOUSUARIO Then
                    ComboTipoUsuario.SelectedItem = tu
                    Exit For
                End If
            Next
            TextDireccion.Text = np.DIRECCION
            TextTelefono.Text = np.TELEFONO
            TextDicose.Text = np.DICOSE
            ComboDepartamento.SelectedItem = Nothing
            Dim d As dDepartamento
            For Each d In ComboDepartamento.Items
                If d.ID = np.IDDEPARTAMENTO Then
                    ComboDepartamento.SelectedItem = d
                    Exit For
                End If
            Next
            ComboLocalidad.SelectedItem = Nothing
            Dim l As dLocalidad
            For Each l In ComboLocalidad.Items
                If l.ID = np.IDLOCALIDAD Then
                    ComboLocalidad.SelectedItem = l
                    Exit For
                End If
            Next
            ComboTecnicos.SelectedItem = Nothing
            Dim t As dTecnicos
            For Each t In ComboTecnicos.Items
                If t.ID = np.TECNICO Then
                    ComboTecnicos.SelectedItem = t
                    Exit For
                End If
            Next
            ComboAgencia.SelectedItem = Nothing
            Dim a As dEmpresaT
            For Each a In ComboAgencia.Items
                If a.ID = np.IDAGENCIA Then
                    ComboAgencia.SelectedItem = a
                    Exit For
                End If
            Next
            TextId.Focus()
        End If
    End Sub

    Public Sub cargarLista()
        Dim c As New dCliente
        Dim lista As New ArrayList
        lista = c.listartodos
        ListClientes.Items.Clear()
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each c In lista
                    ListClientes.Items.Add(c)
                Next
            End If
        End If
    End Sub

    Private Sub listproductores_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListProductores.SelectedIndexChanged
        If ListProductores.SelectedItems.Count = 1 Then
            Dim cli As dCliente = CType(listproductores.SelectedItem, dCliente)
            TextId.Text = cli.ID
            TextNombre.Text = cli.NOMBRE
            TextEmail.Text = cli.EMAIL
            TextEnvio.Text = cli.ENVIO
            TextUsuarioWeb.Text = cli.USUARIO_WEB
            TextRazonSocial.Text = cli.RAZON_SOCIAL
            TextCelular.Text = cli.CELULAR
            TextRut.Text = cli.RUT
            ComboTipoUsuario.SelectedItem = Nothing
            Dim tu As dTipoUsuario
            For Each tu In ComboTipoUsuario.Items
                If tu.ID = cli.TIPOUSUARIO Then
                    ComboTipoUsuario.SelectedItem = tu
                    Exit For
                End If
            Next
            TextDireccion.Text = cli.DIRECCION
            TextTelefono.Text = cli.TELEFONO
            TextFax.Text = cli.FAX
            TextDicose.Text = cli.DICOSE
            ComboDepartamento.SelectedItem = Nothing
            Dim d As dDepartamento
            For Each d In ComboDepartamento.Items
                If d.ID = cli.IDDEPARTAMENTO Then
                    ComboDepartamento.SelectedItem = d
                    Exit For
                End If
            Next
            ComboLocalidad.SelectedItem = Nothing
            Dim l As dLocalidad
            For Each l In ComboLocalidad.Items
                If l.ID = cli.IDLOCALIDAD Then
                    ComboLocalidad.SelectedItem = l
                    Exit For
                End If
            Next
            ComboTecnicos.SelectedItem = Nothing
            Dim t As dTecnicos
            For Each t In ComboTecnicos.Items
                If t.ID = cli.TECNICO Then
                    ComboTecnicos.SelectedItem = t
                    Exit For
                End If
            Next
            ComboTecnicos2.SelectedItem = Nothing
            Dim t2 As dTecnicos
            For Each t2 In ComboTecnicos2.Items
                If t2.ID = cli.TECNICO2 Then
                    ComboTecnicos2.SelectedItem = t2
                    Exit For
                End If
            Next
            ComboTecnicos3.SelectedItem = Nothing
            Dim t3 As dTecnicos
            For Each t3 In ComboTecnicos3.Items
                If t3.ID = cli.TECNICO3 Then
                    ComboTecnicos3.SelectedItem = t3
                    Exit For
                End If
            Next
            ComboAgencia.SelectedItem = Nothing
            Dim a As dEmpresaT
            For Each a In ComboAgencia.Items
                If a.ID = cli.IDAGENCIA Then
                    ComboAgencia.SelectedItem = a
                    Exit For
                End If
            Next
            If cli.CONTRATO = 1 Then
                CheckContrato.Checked = True
            Else
                CheckContrato.Checked = False
            End If
            If cli.SOCIO = 1 Then
                CheckSocio.Checked = True
            Else
                CheckSocio.Checked = False
            End If
            If cli.NOUSAR = 1 Then
                CheckNousar.Checked = True
            Else
                CheckNousar.Checked = False
            End If
            If cli.CARAVANAS = 1 Then
                CheckCaravanas.Checked = True
            Else
                CheckCaravanas.Checked = False
            End If
            If cli.PROLESA = 1 Then
                CheckProlesa.Checked = True
            Else
                CheckProlesa.Checked = False
            End If
            TextObservaciones.Text = cli.OBSERVACIONES
            TextId.Focus()
        End If
    End Sub

    Private Sub TextBuscar_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBuscar.TextChanged
        Dim nombre As String = TextBuscar.Text.Trim
        ListClientes.Items.Clear()
        If nombre.Length > 0 Then
            Dim uncli As New dcliente
            Dim lista As New ArrayList
            lista = uncli.buscarPorNombreTodos(nombre)
            If Not lista Is Nothing And lista.Count > 0 Then

                For Each s As dcliente In lista
                    ListClientes.Items.Add(s)
                Next
                ListClientes.Sorted = True
            End If
        Else : ListClientes.Items.Clear()
        End If
    End Sub

    Private Sub ButtonTodos_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonTodos.Click
        TextBuscar.Text = ""
        cargarLista()
        TextBuscar.Focus()
    End Sub
    Public Sub limpiar()
        TextId.Text = ""
        TextNombre.Text = ""
        TextEmail.Text = ""
        TextEnvio.Text = ""
        TextUsuarioWeb.Text = ""
        TextRazonSocial.Text = ""
        TextTelefono.Text = ""
        TextCelular.Text = ""
        TextRut.Text = ""
        ComboTipoUsuario.Text = ""
        TextDireccion.Text = ""
        TextFax.Text = ""
        TextDicose.Text = ""
        ComboDepartamento.Text = ""
        ComboLocalidad.Text = ""
        ComboTecnicos.Text = ""
        ComboTecnicos2.Text = ""
        ComboTecnicos3.Text = ""
        ComboAgencia.Text = ""
        CheckContrato.Checked = False
        CheckSocio.Checked = False
        CheckNousar.Checked = False
        CheckCaravanas.Checked = False
        CheckProlesa.Checked = False
        ComboProlesa.Enabled = False
        TextObservaciones.Text = ""
        TextFigaro.Text = ""
        TextFacNombre.Text = ""
        TextFacRsocial.Text = ""
        TextFacRut.Text = ""
        TextFacDireccion.Text = ""
        TextFacLocalidad.Text = ""
        ComboFacDepartamento.Text = ""
        TextFacCpostal.Text = ""
        ComboFacGiro.Text = ""
        TextFacTelefonos.Text = ""
        TextFacFax.Text = ""
        TextFacEmail.Text = ""
        TextFacContacto.Text = ""
        TextFacObservaciones.Text = ""
        TextId.Focus()
    End Sub

    Private Sub ButtonNuevo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonNuevo.Click
        cargarLista()
        limpiar()
    End Sub
    Private Sub guardar()
        Dim nombre As String = TextNombre.Text.Trim
        Dim email As String = TextEmail.Text.Trim
        Dim envio As String = TextEnvio.Text.Trim
        'If TextUsuarioWeb.Text.Trim.Length = 0 Then MsgBox("Debe ingresar usuario web", MsgBoxStyle.Exclamation, "Atención") : TextUsuarioWeb.Focus() : Exit Sub
        Dim usuarioweb As String = TextUsuarioWeb.Text.Trim
        Dim razonsocial As String = TextRazonSocial.Text.Trim
        Dim celular As String = TextCelular.Text.Trim
        Dim rut As String = TextRut.Text.Trim
        Dim tipousuario As dTipoUsuario = CType(ComboTipoUsuario.SelectedItem, dTipoUsuario)
        Dim direccion As String = TextDireccion.Text.Trim
        Dim telefono As String = TextTelefono.Text.Trim
        Dim fax As String = TextFax.Text.Trim
        Dim dicose As String = TextDicose.Text.Trim
        Dim departamento As dDepartamento = CType(ComboDepartamento.SelectedItem, dDepartamento)
        Dim localidad As dLocalidad = CType(ComboLocalidad.SelectedItem, dLocalidad)
        Dim tecnico As dTecnicos = CType(ComboTecnicos.SelectedItem, dTecnicos)
        Dim tecnico2 As dTecnicos = CType(ComboTecnicos2.SelectedItem, dTecnicos)
        Dim tecnico3 As dTecnicos = CType(ComboTecnicos3.SelectedItem, dTecnicos)
        Dim agencia As dEmpresaT = CType(ComboAgencia.SelectedItem, dEmpresaT)
        Dim contrato As Integer
        If CheckContrato.Checked = True Then
            contrato = 1
        Else
            contrato = 0
        End If
        Dim socio As Integer
        If CheckSocio.Checked = True Then
            socio = 1
        Else
            socio = 0
        End If
        Dim nousar As Integer
        If CheckNousar.Checked = True Then
            nousar = 1
        Else
            nousar = 0
        End If
        Dim caravanas As Integer
        If CheckCaravanas.Checked = True Then
            caravanas = 1
        Else
            caravanas = 0
        End If
        Dim prolesa As Integer
        If CheckProlesa.Checked = True Then
            prolesa = 1
        Else
            prolesa = 0
        End If
        Dim prolesasuc As dProlesa = CType(ComboProlesa.SelectedItem, dProlesa)
        Dim observaciones As String = ""
        If TextObservaciones.Text <> "" Then
            observaciones = TextObservaciones.Text.Trim
        End If
        Dim facnombre As String = ""
        If TextFacNombre.Text <> "" Then
            facnombre = TextFacNombre.Text.Trim
        End If
        Dim facrsocial As String = ""
        If TextFacRsocial.Text <> "" Then
            facrsocial = TextFacRsocial.Text.Trim
        End If
        Dim facrut As String = ""
        If TextFacRut.Text <> "" Then
            facrut = TextFacRut.Text.Trim
        End If
        Dim facdireccion As String = ""
        If TextFacDireccion.Text <> "" Then
            facdireccion = TextFacDireccion.Text.Trim
        End If
        Dim faclocalidad As String = ""
        If TextFacLocalidad.Text <> "" Then
            faclocalidad = TextFacLocalidad.Text.Trim
        End If
        Dim facdepartamento As dDepartamento = CType(ComboFacDepartamento.SelectedItem, dDepartamento)
        Dim faccpostal As String = ""
        If TextFacCpostal.Text <> "" Then
            faccpostal = TextFacCpostal.Text.Trim
        End If
        Dim facgiro As dGiro = CType(ComboFacGiro.SelectedItem, dGiro)
        Dim factelefonos As String = ""
        If TextFacTelefonos.Text <> "" Then
            factelefonos = TextFacTelefonos.Text.Trim
        End If
        Dim facfax As String = ""
        If TextFacFax.Text <> "" Then
            facfax = TextFacFax.Text.Trim
        End If
        Dim facemail As String = ""
        If TextFacEmail.Text <> "" Then
            facemail = TextFacEmail.Text.Trim
        End If
        Dim faccontacto As String = ""
        If TextFacContacto.Text <> "" Then
            faccontacto = TextFacContacto.Text.Trim
        End If
        Dim facobservaciones As String = ""
        If TextFacObservaciones.Text <> "" Then
            facobservaciones = TextFacObservaciones.Text.Trim
        End If
        Dim faclista As Integer = NumericLista.Value
        If Not ListClientes.SelectedItem Is Nothing And TextId.Text.Trim.Length > 0 Then
            If TextNombre.Text.Trim.Length > 0 Then
                Dim cli As New dCliente()
                Dim pw_com As New dClienteWeb_com
                Dim nocargaenweb As Integer = 0
                'Dim pw_uy As New dclienteWeb_uy
                pw_com.USUARIO = TextUsuarioWeb.Text.Trim
                pw_com = pw_com.buscar
                Dim idclienteweb_com As Long
                'Dim idclienteweb_uy As Long

                If Not pw_com Is Nothing Then
                    idclienteweb_com = pw_com.ID
                Else
                    Dim result = MessageBox.Show("No existe el usuario web, desea continuar de todos modos?", "Atención", MessageBoxButtons.YesNo)
                    If result = DialogResult.No Then
                        Exit Sub
                    ElseIf result = DialogResult.Yes Then
                        nocargaenweb = 1
                        'idclienteweb_com = 4119
                    End If

                End If
                'pw_uy.USUARIO = TextUsuarioWeb.Text.Trim
                'pw_uy = pw_uy.buscar
                'If Not pw_uy Is Nothing Then
                'idclienteweb_uy = pw_uy.ID
                'Else
                'MsgBox("No existe el usuario web (.uy)")
                'Exit Sub
                'End If

                'NET*************************************
                Dim id As Long = TextId.Text.Trim
                cli.ID = id
                cli.NOMBRE = nombre
                cli.EMAIL = email
                cli.ENVIO = envio
                cli.USUARIO_WEB = usuarioweb
                cli.RAZON_SOCIAL = razonsocial
                cli.CELULAR = celular
                cli.RUT = rut
                cli.TIPOUSUARIO = tipousuario.ID
                cli.DIRECCION = direccion
                cli.TELEFONO = telefono
                cli.FAX = fax
                cli.DICOSE = dicose
                If departamento Is Nothing Then
                    cli.IDDEPARTAMENTO = 999
                Else
                    cli.IDDEPARTAMENTO = departamento.ID
                End If
                If localidad Is Nothing Then
                    cli.IDLOCALIDAD = 999
                Else
                    cli.IDLOCALIDAD = localidad.ID
                End If
                If tecnico Is Nothing Then
                    cli.TECNICO = 3282
                Else
                    cli.TECNICO = tecnico.ID
                End If
                If tecnico2 Is Nothing Then
                    cli.TECNICO2 = 3282
                Else
                    cli.TECNICO2 = tecnico2.ID
                End If
                If tecnico3 Is Nothing Then
                    cli.TECNICO3 = 3282
                Else
                    cli.TECNICO3 = tecnico3.ID
                End If
                If Not agencia Is Nothing Then
                    cli.IDAGENCIA = agencia.ID
                End If
                cli.CONTRATO = contrato
                cli.SOCIO = socio
                cli.NOUSAR = nousar
                cli.CARAVANAS = caravanas
                cli.PROLESA = prolesa
                If prolesasuc Is Nothing Then
                    cli.PROLESASUC = 0
                Else
                    cli.PROLESASUC = prolesasuc.ID
                End If
                cli.OBSERVACIONES = observaciones
                cli.FAC_NOMBRE = facnombre
                cli.FAC_RSOCIAL = facrsocial
                cli.FAC_RUT = facrut
                cli.FAC_DIRECCION = facdireccion
                cli.FAC_LOCALIDAD = faclocalidad
                If facdepartamento Is Nothing Then
                    cli.FAC_DEPARTAMENTO = 999
                Else
                    cli.FAC_DEPARTAMENTO = facdepartamento.ID
                End If
                cli.FAC_CPOSTAL = faccpostal
                If facgiro Is Nothing Then
                    cli.FAC_GIRO = 16
                Else
                    cli.FAC_GIRO = facgiro.ID
                End If
                cli.FAC_TELEFONOS = factelefonos
                cli.FAC_FAX = facfax
                cli.FAC_EMAIL = facemail
                cli.FAC_CONTACTO = faccontacto
                cli.FAC_OBSERVACIONES = facobservaciones
                cli.FAC_LISTA = faclista
                'COM**********************************
                If nocargaenweb = 0 Then
                    pw_com.ID = idclienteweb_com
                    pw_com.NOMBRE = nombre
                    pw_com.EMAIL_1 = email
                    pw_com.USUARIO = usuarioweb
                    pw_com.PASSWORD = usuarioweb
                    pw_com.RAZON_SOCIAL = razonsocial
                    pw_com.CELULAR_1 = celular
                    pw_com.RUT = rut
                    pw_com.TIPO_USUARIO_ID = tipousuario.ID
                    pw_com.DIRECCION = direccion
                    pw_com.TELEFONO_1 = telefono
                    pw_com.DICOSE = dicose
                    pw_com.VER_CONTROL_LECHERO = 1
                    pw_com.VER_AGUA = 1
                    pw_com.VER_PAL = 1
                    pw_com.VER_SEROLOGIA = 1
                    pw_com.VER_ANTIBIOGRAMA = 1
                    pw_com.VER_PARASITOLOGIA = 1
                    pw_com.VER_PRODUCTOS_SUBPRODUCTOS = 1
                    pw_com.VER_PATOLOGIA = 1
                    pw_com.VER_CALIDAD_DE_LECHE = 1
                End If

                If (cli.modificar(Usuario)) Then
                    If nocargaenweb = 0 Then
                        pw_com.modificar(Usuario)
                    End If
                    'pw_uy.modificar(Usuario)

                    MsgBox("cliente modificado", MsgBoxStyle.Information, "Atención")
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            End If
        Else
            If TextNombre.Text.Trim.Length > 0 Then
                Dim cli As New dCliente()
                Dim cw_com As New dClienteWeb_com
                'Dim pw_uy As New dclienteWeb_uy
                'NET***********************************
                'cli.ID = id
                cli.NOMBRE = nombre
                cli.EMAIL = email
                cli.ENVIO = envio
                cli.USUARIO_WEB = usuarioweb
                clienteweb_com = usuarioweb
                cli.RAZON_SOCIAL = razonsocial
                cli.CELULAR = celular
                cli.RUT = rut
                cli.TIPOUSUARIO = tipousuario.ID
                cli.DIRECCION = direccion
                cli.TELEFONO = telefono
                cli.FAX = fax
                cli.DICOSE = dicose
                If departamento Is Nothing Then
                    cli.IDDEPARTAMENTO = 999
                Else
                    cli.IDDEPARTAMENTO = departamento.ID
                End If
                If localidad Is Nothing Then
                    cli.IDLOCALIDAD = 999
                Else
                    cli.IDLOCALIDAD = localidad.ID
                End If
                If tecnico Is Nothing Then
                    cli.TECNICO = 3282
                Else
                    cli.TECNICO = tecnico.ID
                End If
                If tecnico2 Is Nothing Then
                    cli.TECNICO2 = 3282
                Else
                    cli.TECNICO2 = tecnico2.ID
                End If
                If tecnico3 Is Nothing Then
                    cli.TECNICO3 = 3282
                Else
                    cli.TECNICO3 = tecnico3.ID
                End If
                If Not agencia Is Nothing Then
                    cli.IDAGENCIA = agencia.ID
                End If
                cli.CONTRATO = contrato
                cli.SOCIO = socio
                cli.NOUSAR = nousar
                cli.CARAVANAS = caravanas
                cli.PROLESA = prolesa
                If prolesasuc Is Nothing Then
                    cli.PROLESASUC = 0
                Else
                    cli.PROLESASUC = prolesasuc.ID
                End If
                cli.OBSERVACIONES = observaciones
                cli.FAC_NOMBRE = facnombre
                cli.FAC_RSOCIAL = facrsocial
                cli.FAC_RUT = facrut
                cli.FAC_DIRECCION = facdireccion
                cli.FAC_LOCALIDAD = faclocalidad
                If facdepartamento Is Nothing Then
                    cli.FAC_DEPARTAMENTO = 999
                Else
                    cli.FAC_DEPARTAMENTO = facdepartamento.ID
                End If
                cli.FAC_CPOSTAL = faccpostal
                If facgiro Is Nothing Then
                    cli.FAC_GIRO = 0
                Else
                    cli.FAC_GIRO = facgiro.ID
                End If
                cli.FAC_TELEFONOS = factelefonos
                cli.FAC_FAX = facfax
                cli.FAC_EMAIL = facemail
                cli.FAC_CONTACTO = faccontacto
                cli.FAC_OBSERVACIONES = facobservaciones
                cli.FAC_LISTA = faclista
                'COM**********************************
                'cw_com.ID = idclienteweb_com
                cw_com.NOMBRE = nombre
                cw_com.EMAIL_1 = email
                cw_com.USUARIO = usuarioweb
                cw_com.PASSWORD = usuarioweb
                clienteweb_uy = usuarioweb
                cw_com.RAZON_SOCIAL = razonsocial
                cw_com.CELULAR_1 = celular
                cw_com.RUT = rut
                cw_com.TIPO_USUARIO_ID = tipousuario.ID
                cw_com.DIRECCION = direccion
                cw_com.TELEFONO_1 = telefono
                cw_com.DICOSE = dicose
                cw_com.VER_CONTROL_LECHERO = 1
                cw_com.VER_AGUA = 1
                cw_com.VER_PAL = 1
                cw_com.VER_SEROLOGIA = 1
                cw_com.VER_ANTIBIOGRAMA = 1
                cw_com.VER_PARASITOLOGIA = 1
                cw_com.VER_PRODUCTOS_SUBPRODUCTOS = 1
                cw_com.VER_PATOLOGIA = 1
                cw_com.VER_CALIDAD_DE_LECHE = 1

                If (cli.guardar(Usuario)) Then
                    cw_com.guardar(Usuario)
                    'cw_uy.guardar(Usuario)
                    crearcarpetas_com()
                    'crearcarpetas_uy()
                    MsgBox("cliente guardado", MsgBoxStyle.Information, "Atención")
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            End If
        End If
    End Sub
    Private Sub ButtonGuardar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonGuardar.Click
        guardar()
        limpiar()
        cargarLista()
    End Sub
    Private Sub crearcarpetas_com()
        Dim cw_com As New dclienteWeb_com
        cw_com.USUARIO = clienteweb_com
        cw_com = cw_com.buscar
        If Not cw_com Is Nothing Then
            carpeta = cw_com.ID
        Else
            MsgBox("No existe el usuario web (.com)")
        End If
        crea_carpeta_com()
        crea_agro_nutricion_com()
        crea_agua_com()
        crea_ambiental_com()
        crea_antibiograma_com()
        crea_calidad_de_leche_com()
        crea_control_lechero_com()
        crea_lactometros_chequeos_maquina_com()
        crea_otros_servicios_com()
        crea_pal_com()
        crea_parasitologia_com()
        crea_patologia_com()
        crea_productos_subproductos_com()
        crea_serologia_com()
        crea_agro_suelos_com()
        crea_brucelosis_leche_com()

    End Sub
   
    Public Sub cargarComboDepartamento()
        Dim d As New dDepartamento
        Dim lista As New ArrayList
        lista = d.listar
        ComboDepartamento.Items.Clear()
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each d In lista
                    ComboDepartamento.Items.Add(d)
                    ComboFacDepartamento.Items.Add(d)
                Next
            End If
        End If
    End Sub
    Public Sub cargarcomboProlesa()
        Dim p As New dProlesa
        Dim lista As New ArrayList
        lista = p.listar
        ComboProlesa.Items.Clear()
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each p In lista
                    ComboProlesa.Items.Add(p)
                Next
            End If
        End If
    End Sub
    Public Sub cargarComboGiro()
        Dim g As New dGiro
        Dim lista As New ArrayList
        lista = g.listar
        ComboFacGiro.Items.Clear()
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each g In lista
                    ComboFacGiro.Items.Add(g)
                Next
            End If
        End If
    End Sub
    Public Sub cargarComboLocalidad()
        Dim l As New dLocalidad
        Dim lista As New ArrayList
        lista = l.listar
        ComboLocalidad.Items.Clear()
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each l In lista
                    ComboLocalidad.Items.Add(l)
                Next
            End If
        End If
    End Sub

    Public Sub cargarComboTipoUsuario()
        Dim tu As New dTipoUsuario
        Dim lista As New ArrayList
        lista = tu.listar
        ComboTipoUsuario.Items.Clear()
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each tu In lista
                    ComboTipoUsuario.Items.Add(tu)
                Next
            End If
        End If
    End Sub
    Public Sub cargarComboTecnicos()
        Dim t As New dTecnicos
        Dim lista As New ArrayList
        lista = t.listar
        ComboTecnicos.Items.Clear()
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each t In lista
                    ComboTecnicos.Items.Add(t)
                    ComboTecnicos2.Items.Add(t)
                    ComboTecnicos3.Items.Add(t)
                Next
            End If
        End If
    End Sub
    Public Sub cargarComboAgencia()
        Dim a As New dEmpresaT
        Dim lista As New ArrayList
        lista = a.listar
        ComboAgencia.Items.Clear()
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each a In lista
                    ComboAgencia.Items.Add(a)
                Next
            End If
        End If
    End Sub

    Private Sub ButtonEmpresa_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonEmpresa.Click
        If TextId.Text <> "" Then
            Dim idcliente As Long = TextId.Text.Trim
            Dim v As New FormProductorEmpresa(Usuario, idcliente)
            v.ShowDialog()
        End If
    End Sub
    Public Sub crea_carpeta_com()
        Dim cweb_com As New dclienteWeb_com
        Dim user As String = "colaveco"
        Dim pass As String = "CLV1582782"

        Dim peticionFTP As FtpWebRequest
        'Dim lista As New ArrayList
        'lista = cweb_com.listarxid
        'If Not lista Is Nothing Then
        'If lista.Count > 0 Then
        'carpeta = cweb_com.ID
        Dim dir As String = "ftp://colaveco.com.uy/www/gestor/data_file/" & carpeta & ""

        ' Creamos una peticion FTP con la dirección del directorio que queremos crear
        peticionFTP = CType(WebRequest.Create(New Uri(dir)), FtpWebRequest)

        ' Fijamos el usuario y la contraseña de la petición
        peticionFTP.Credentials = New NetworkCredential(user, pass)

        ' Seleccionamos el comando que vamos a utilizar: Crear un directorio
        peticionFTP.Method = WebRequestMethods.Ftp.MakeDirectory

        Try
            Dim respuesta As FtpWebResponse
            respuesta = CType(peticionFTP.GetResponse(), FtpWebResponse)
            respuesta.Close()
            ' Si todo ha ido bien, se devolverá String.Empty
            'Return String.Empty
        Catch ex As Exception
            ' Si se produce algún fallo, se devolverá el mensaje del error
            'Return ex.Message
        End Try

    End Sub
    Public Sub crea_control_lechero_com()
        Dim cweb_com As New dclienteWeb_com
        Dim user As String = "colaveco"
        Dim pass As String = "CLV1582782"

        Dim peticionFTP As FtpWebRequest
        'Dim lista As New ArrayList
        'lista = cweb_com.listarxid
        'If Not lista Is Nothing Then
        'If lista.Count > 0 Then
        'carpeta = cweb_com.ID
        Dim dir As String = "ftp://colaveco.com.uy/www/gestor/data_file/" & carpeta & "/control_lechero/"

        ' Creamos una peticion FTP con la dirección del directorio que queremos crear
        peticionFTP = CType(WebRequest.Create(New Uri(dir)), FtpWebRequest)

        ' Fijamos el usuario y la contraseña de la petición
        peticionFTP.Credentials = New NetworkCredential(user, pass)

        ' Seleccionamos el comando que vamos a utilizar: Crear un directorio
        peticionFTP.Method = WebRequestMethods.Ftp.MakeDirectory

        Try
            Dim respuesta As FtpWebResponse
            respuesta = CType(peticionFTP.GetResponse(), FtpWebResponse)
            respuesta.Close()
            ' Si todo ha ido bien, se devolverá String.Empty
            'Return String.Empty
        Catch ex As Exception
            ' Si se produce algún fallo, se devolverá el mensaje del error
            'Return ex.Message
        End Try

    End Sub
    Public Sub crea_agua_com()
        Dim cweb_com As New dclienteWeb_com
        Dim user As String = "colaveco"
        Dim pass As String = "CLV1582782"

        Dim peticionFTP As FtpWebRequest
        'Dim lista As New ArrayList
        'lista = cweb_com.listarxid
        'If Not lista Is Nothing Then
        'If lista.Count > 0 Then
        'carpeta = cweb_com.ID
        Dim dir As String = "ftp://colaveco.com.uy/www/gestor/data_file/" & carpeta & "/agua/"

        ' Creamos una peticion FTP con la dirección del directorio que queremos crear
        peticionFTP = CType(WebRequest.Create(New Uri(dir)), FtpWebRequest)

        ' Fijamos el usuario y la contraseña de la petición
        peticionFTP.Credentials = New NetworkCredential(user, pass)

        ' Seleccionamos el comando que vamos a utilizar: Crear un directorio
        peticionFTP.Method = WebRequestMethods.Ftp.MakeDirectory

        Try
            Dim respuesta As FtpWebResponse
            respuesta = CType(peticionFTP.GetResponse(), FtpWebResponse)
            respuesta.Close()
            ' Si todo ha ido bien, se devolverá String.Empty
            'Return String.Empty
        Catch ex As Exception
            ' Si se produce algún fallo, se devolverá el mensaje del error
            'Return ex.Message
        End Try

    End Sub
    Public Sub crea_antibiograma_com()
        Dim cweb_com As New dclienteWeb_com
        Dim user As String = "colaveco"
        Dim pass As String = "CLV1582782"

        Dim peticionFTP As FtpWebRequest
        'Dim lista As New ArrayList
        'lista = cweb_com.listarxid
        'If Not lista Is Nothing Then
        'If lista.Count > 0 Then
        'carpeta = cweb_com.ID
        Dim dir As String = "ftp://colaveco.com.uy/www/gestor/data_file/" & carpeta & "/antibiograma/"

        ' Creamos una peticion FTP con la dirección del directorio que queremos crear
        peticionFTP = CType(WebRequest.Create(New Uri(dir)), FtpWebRequest)

        ' Fijamos el usuario y la contraseña de la petición
        peticionFTP.Credentials = New NetworkCredential(user, pass)

        ' Seleccionamos el comando que vamos a utilizar: Crear un directorio
        peticionFTP.Method = WebRequestMethods.Ftp.MakeDirectory

        Try
            Dim respuesta As FtpWebResponse
            respuesta = CType(peticionFTP.GetResponse(), FtpWebResponse)
            respuesta.Close()
            ' Si todo ha ido bien, se devolverá String.Empty
            'Return String.Empty
        Catch ex As Exception
            ' Si se produce algún fallo, se devolverá el mensaje del error
            'Return ex.Message
        End Try

    End Sub
    Public Sub crea_pal_com()
        Dim cweb_com As New dclienteWeb_com
        Dim user As String = "colaveco"
        Dim pass As String = "CLV1582782"

        Dim peticionFTP As FtpWebRequest
        'Dim lista As New ArrayList
        'lista = cweb_com.listarxid
        'If Not lista Is Nothing Then
        'If lista.Count > 0 Then
        'carpeta = cweb_com.ID
        Dim dir As String = "ftp://colaveco.com.uy/www/gestor/data_file/" & carpeta & "/pal/"

        ' Creamos una peticion FTP con la dirección del directorio que queremos crear
        peticionFTP = CType(WebRequest.Create(New Uri(dir)), FtpWebRequest)

        ' Fijamos el usuario y la contraseña de la petición
        peticionFTP.Credentials = New NetworkCredential(user, pass)

        ' Seleccionamos el comando que vamos a utilizar: Crear un directorio
        peticionFTP.Method = WebRequestMethods.Ftp.MakeDirectory

        Try
            Dim respuesta As FtpWebResponse
            respuesta = CType(peticionFTP.GetResponse(), FtpWebResponse)
            respuesta.Close()
            ' Si todo ha ido bien, se devolverá String.Empty
            'Return String.Empty
        Catch ex As Exception
            ' Si se produce algún fallo, se devolverá el mensaje del error
            'Return ex.Message
        End Try

    End Sub
    Public Sub crea_parasitologia_com()
        Dim cweb_com As New dclienteWeb_com
        Dim user As String = "colaveco"
        Dim pass As String = "CLV1582782"

        Dim peticionFTP As FtpWebRequest
        'Dim lista As New ArrayList
        'lista = cweb_com.listarxid
        'If Not lista Is Nothing Then
        'If lista.Count > 0 Then
        'carpeta = cweb_com.ID
        Dim dir As String = "ftp://colaveco.com.uy/www/gestor/data_file/" & carpeta & "/parasitologia/"

        ' Creamos una peticion FTP con la dirección del directorio que queremos crear
        peticionFTP = CType(WebRequest.Create(New Uri(dir)), FtpWebRequest)

        ' Fijamos el usuario y la contraseña de la petición
        peticionFTP.Credentials = New NetworkCredential(user, pass)

        ' Seleccionamos el comando que vamos a utilizar: Crear un directorio
        peticionFTP.Method = WebRequestMethods.Ftp.MakeDirectory

        Try
            Dim respuesta As FtpWebResponse
            respuesta = CType(peticionFTP.GetResponse(), FtpWebResponse)
            respuesta.Close()
            ' Si todo ha ido bien, se devolverá String.Empty
            'Return String.Empty
        Catch ex As Exception
            ' Si se produce algún fallo, se devolverá el mensaje del error
            'Return ex.Message
        End Try

    End Sub
    Public Sub crea_productos_subproductos_com()
        Dim cweb_com As New dclienteWeb_com
        Dim user As String = "colaveco"
        Dim pass As String = "CLV1582782"

        Dim peticionFTP As FtpWebRequest
        'Dim lista As New ArrayList
        'lista = cweb_com.listarxid
        'If Not lista Is Nothing Then
        'If lista.Count > 0 Then
        'carpeta = cweb_com.ID
        Dim dir As String = "ftp://colaveco.com.uy/www/gestor/data_file/" & carpeta & "/productos_subproductos/"

        ' Creamos una peticion FTP con la dirección del directorio que queremos crear
        peticionFTP = CType(WebRequest.Create(New Uri(dir)), FtpWebRequest)

        ' Fijamos el usuario y la contraseña de la petición
        peticionFTP.Credentials = New NetworkCredential(user, pass)

        ' Seleccionamos el comando que vamos a utilizar: Crear un directorio
        peticionFTP.Method = WebRequestMethods.Ftp.MakeDirectory

        Try
            Dim respuesta As FtpWebResponse
            respuesta = CType(peticionFTP.GetResponse(), FtpWebResponse)
            respuesta.Close()
            ' Si todo ha ido bien, se devolverá String.Empty
            'Return String.Empty
        Catch ex As Exception
            ' Si se produce algún fallo, se devolverá el mensaje del error
            'Return ex.Message
        End Try

    End Sub
    Public Sub crea_serologia_com()
        Dim cweb_com As New dclienteWeb_com
        Dim user As String = "colaveco"
        Dim pass As String = "CLV1582782"

        Dim peticionFTP As FtpWebRequest
        'Dim lista As New ArrayList
        'lista = cweb_com.listarxid
        'If Not lista Is Nothing Then
        'If lista.Count > 0 Then
        'carpeta = cweb_com.ID
        Dim dir As String = "ftp://colaveco.com.uy/www/gestor/data_file/" & carpeta & "/serologia/"

        ' Creamos una peticion FTP con la dirección del directorio que queremos crear
        peticionFTP = CType(WebRequest.Create(New Uri(dir)), FtpWebRequest)

        ' Fijamos el usuario y la contraseña de la petición
        peticionFTP.Credentials = New NetworkCredential(user, pass)

        ' Seleccionamos el comando que vamos a utilizar: Crear un directorio
        peticionFTP.Method = WebRequestMethods.Ftp.MakeDirectory

        Try
            Dim respuesta As FtpWebResponse
            respuesta = CType(peticionFTP.GetResponse(), FtpWebResponse)
            respuesta.Close()
            ' Si todo ha ido bien, se devolverá String.Empty
            'Return String.Empty
        Catch ex As Exception
            ' Si se produce algún fallo, se devolverá el mensaje del error
            'Return ex.Message
        End Try

    End Sub
    Public Sub crea_patologia_com()
        Dim cweb_com As New dclienteWeb_com
        Dim user As String = "colaveco"
        Dim pass As String = "CLV1582782"

        Dim peticionFTP As FtpWebRequest
        'Dim lista As New ArrayList
        'lista = cweb_com.listarxid
        'If Not lista Is Nothing Then
        'If lista.Count > 0 Then
        'carpeta = cweb_com.ID
        Dim dir As String = "ftp://colaveco.com.uy/www/gestor/data_file/" & carpeta & "/patologia/"

        ' Creamos una peticion FTP con la dirección del directorio que queremos crear
        peticionFTP = CType(WebRequest.Create(New Uri(dir)), FtpWebRequest)

        ' Fijamos el usuario y la contraseña de la petición
        peticionFTP.Credentials = New NetworkCredential(user, pass)

        ' Seleccionamos el comando que vamos a utilizar: Crear un directorio
        peticionFTP.Method = WebRequestMethods.Ftp.MakeDirectory

        Try
            Dim respuesta As FtpWebResponse
            respuesta = CType(peticionFTP.GetResponse(), FtpWebResponse)
            respuesta.Close()
            ' Si todo ha ido bien, se devolverá String.Empty
            'Return String.Empty
        Catch ex As Exception
            ' Si se produce algún fallo, se devolverá el mensaje del error
            'Return ex.Message
        End Try

    End Sub
    Public Sub crea_calidad_de_leche_com()
        Dim cweb_com As New dclienteWeb_com
        Dim user As String = "colaveco"
        Dim pass As String = "CLV1582782"

        Dim peticionFTP As FtpWebRequest
        'Dim lista As New ArrayList
        'lista = cweb_com.listarxid
        'If Not lista Is Nothing Then
        'If lista.Count > 0 Then
        'carpeta = cweb_com.ID
        Dim dir As String = "ftp://colaveco.com.uy/www/gestor/data_file/" & carpeta & "/calidad_de_leche/"

        ' Creamos una peticion FTP con la dirección del directorio que queremos crear
        peticionFTP = CType(WebRequest.Create(New Uri(dir)), FtpWebRequest)

        ' Fijamos el usuario y la contraseña de la petición
        peticionFTP.Credentials = New NetworkCredential(user, pass)

        ' Seleccionamos el comando que vamos a utilizar: Crear un directorio
        peticionFTP.Method = WebRequestMethods.Ftp.MakeDirectory

        Try
            Dim respuesta As FtpWebResponse
            respuesta = CType(peticionFTP.GetResponse(), FtpWebResponse)
            respuesta.Close()
            ' Si todo ha ido bien, se devolverá String.Empty
            'Return String.Empty
        Catch ex As Exception
            ' Si se produce algún fallo, se devolverá el mensaje del error
            'Return ex.Message
        End Try

    End Sub
    Public Sub crea_ambiental_com()
        Dim cweb_com As New dclienteWeb_com
        Dim user As String = "colaveco"
        Dim pass As String = "CLV1582782"

        Dim peticionFTP As FtpWebRequest
        'Dim lista As New ArrayList
        'lista = cweb_com.listarxid
        'If Not lista Is Nothing Then
        'If lista.Count > 0 Then
        'carpeta = cweb_com.ID
        Dim dir As String = "ftp://colaveco.com.uy/www/gestor/data_file/" & carpeta & "/ambiental/"

        ' Creamos una peticion FTP con la dirección del directorio que queremos crear
        peticionFTP = CType(WebRequest.Create(New Uri(dir)), FtpWebRequest)

        ' Fijamos el usuario y la contraseña de la petición
        peticionFTP.Credentials = New NetworkCredential(user, pass)

        ' Seleccionamos el comando que vamos a utilizar: Crear un directorio
        peticionFTP.Method = WebRequestMethods.Ftp.MakeDirectory

        Try
            Dim respuesta As FtpWebResponse
            respuesta = CType(peticionFTP.GetResponse(), FtpWebResponse)
            respuesta.Close()
            ' Si todo ha ido bien, se devolverá String.Empty
            'Return String.Empty
        Catch ex As Exception
            ' Si se produce algún fallo, se devolverá el mensaje del error
            'Return ex.Message
        End Try

    End Sub
    Public Sub crea_lactometros_chequeos_maquina_com()
        Dim cweb_com As New dclienteWeb_com
        Dim user As String = "colaveco"
        Dim pass As String = "CLV1582782"

        Dim peticionFTP As FtpWebRequest
        'Dim lista As New ArrayList
        'lista = cweb_com.listarxid
        'If Not lista Is Nothing Then
        'If lista.Count > 0 Then
        'carpeta = cweb_com.ID
        Dim dir As String = "ftp://colaveco.com.uy/www/gestor/data_file/" & carpeta & "/lactometros_chequeos_maquina/"

        ' Creamos una peticion FTP con la dirección del directorio que queremos crear
        peticionFTP = CType(WebRequest.Create(New Uri(dir)), FtpWebRequest)

        ' Fijamos el usuario y la contraseña de la petición
        peticionFTP.Credentials = New NetworkCredential(user, pass)

        ' Seleccionamos el comando que vamos a utilizar: Crear un directorio
        peticionFTP.Method = WebRequestMethods.Ftp.MakeDirectory

        Try
            Dim respuesta As FtpWebResponse
            respuesta = CType(peticionFTP.GetResponse(), FtpWebResponse)
            respuesta.Close()
            ' Si todo ha ido bien, se devolverá String.Empty
            'Return String.Empty
        Catch ex As Exception
            ' Si se produce algún fallo, se devolverá el mensaje del error
            'Return ex.Message
        End Try

    End Sub
    Public Sub crea_agro_nutricion_com()
        Dim cweb_com As New dclienteWeb_com
        Dim user As String = "colaveco"
        Dim pass As String = "CLV1582782"

        Dim peticionFTP As FtpWebRequest
        'Dim lista As New ArrayList
        'lista = cweb_com.listarxid
        'If Not lista Is Nothing Then
        'If lista.Count > 0 Then
        'carpeta = cweb_com.ID
        Dim dir As String = "ftp://colaveco.com.uy/www/gestor/data_file/" & carpeta & "/agro_nutricion/"

        ' Creamos una peticion FTP con la dirección del directorio que queremos crear
        peticionFTP = CType(WebRequest.Create(New Uri(dir)), FtpWebRequest)

        ' Fijamos el usuario y la contraseña de la petición
        peticionFTP.Credentials = New NetworkCredential(user, pass)

        ' Seleccionamos el comando que vamos a utilizar: Crear un directorio
        peticionFTP.Method = WebRequestMethods.Ftp.MakeDirectory

        Try
            Dim respuesta As FtpWebResponse
            respuesta = CType(peticionFTP.GetResponse(), FtpWebResponse)
            respuesta.Close()
            ' Si todo ha ido bien, se devolverá String.Empty
            'Return String.Empty
        Catch ex As Exception
            ' Si se produce algún fallo, se devolverá el mensaje del error
            'Return ex.Message
        End Try

    End Sub
    Public Sub crea_agro_suelos_com()
        Dim cweb_com As New dclienteWeb_com
        Dim user As String = "colaveco"
        Dim pass As String = "CLV1582782"

        Dim peticionFTP As FtpWebRequest
        'Dim lista As New ArrayList
        'lista = cweb_com.listarxid
        'If Not lista Is Nothing Then
        'If lista.Count > 0 Then
        'carpeta = cweb_com.ID
        Dim dir As String = "ftp://colaveco.com.uy/www/gestor/data_file/" & carpeta & "/agro_suelos/"

        ' Creamos una peticion FTP con la dirección del directorio que queremos crear
        peticionFTP = CType(WebRequest.Create(New Uri(dir)), FtpWebRequest)

        ' Fijamos el usuario y la contraseña de la petición
        peticionFTP.Credentials = New NetworkCredential(user, pass)

        ' Seleccionamos el comando que vamos a utilizar: Crear un directorio
        peticionFTP.Method = WebRequestMethods.Ftp.MakeDirectory

        Try
            Dim respuesta As FtpWebResponse
            respuesta = CType(peticionFTP.GetResponse(), FtpWebResponse)
            respuesta.Close()
            ' Si todo ha ido bien, se devolverá String.Empty
            'Return String.Empty
        Catch ex As Exception
            ' Si se produce algún fallo, se devolverá el mensaje del error
            'Return ex.Message
        End Try

    End Sub
    Public Sub crea_brucelosis_leche_com()
        Dim cweb_com As New dclienteWeb_com
        Dim user As String = "colaveco"
        Dim pass As String = "CLV1582782"

        Dim peticionFTP As FtpWebRequest
        'Dim lista As New ArrayList
        'lista = cweb_com.listarxid
        'If Not lista Is Nothing Then
        'If lista.Count > 0 Then
        'carpeta = cweb_com.ID
        Dim dir As String = "ftp://colaveco.com.uy/www/gestor/data_file/" & carpeta & "/brucelosis_leche/"

        ' Creamos una peticion FTP con la dirección del directorio que queremos crear
        peticionFTP = CType(WebRequest.Create(New Uri(dir)), FtpWebRequest)

        ' Fijamos el usuario y la contraseña de la petición
        peticionFTP.Credentials = New NetworkCredential(user, pass)

        ' Seleccionamos el comando que vamos a utilizar: Crear un directorio
        peticionFTP.Method = WebRequestMethods.Ftp.MakeDirectory

        Try
            Dim respuesta As FtpWebResponse
            respuesta = CType(peticionFTP.GetResponse(), FtpWebResponse)
            respuesta.Close()
            ' Si todo ha ido bien, se devolverá String.Empty
            'Return String.Empty
        Catch ex As Exception
            ' Si se produce algún fallo, se devolverá el mensaje del error
            'Return ex.Message
        End Try

    End Sub
    Public Sub crea_otros_servicios_com()
        Dim cweb_com As New dclienteWeb_com
        Dim user As String = "colaveco"
        Dim pass As String = "CLV1582782"

        Dim peticionFTP As FtpWebRequest
        'Dim lista As New ArrayList
        'lista = cweb_com.listarxid
        'If Not lista Is Nothing Then
        'If lista.Count > 0 Then
        'carpeta = cweb_com.ID
        Dim dir As String = "ftp://colaveco.com.uy/www/gestor/data_file/" & carpeta & "/otros_servicios/"

        ' Creamos una peticion FTP con la dirección del directorio que queremos crear
        peticionFTP = CType(WebRequest.Create(New Uri(dir)), FtpWebRequest)

        ' Fijamos el usuario y la contraseña de la petición
        peticionFTP.Credentials = New NetworkCredential(user, pass)

        ' Seleccionamos el comando que vamos a utilizar: Crear un directorio
        peticionFTP.Method = WebRequestMethods.Ftp.MakeDirectory

        Try
            Dim respuesta As FtpWebResponse
            respuesta = CType(peticionFTP.GetResponse(), FtpWebResponse)
            respuesta.Close()
            ' Si todo ha ido bien, se devolverá String.Empty
            'Return String.Empty
        Catch ex As Exception
            ' Si se produce algún fallo, se devolverá el mensaje del error
            'Return ex.Message
        End Try

    End Sub
   

    Function generarClaveSHA1(ByVal cadena As String) As String

        Dim enc As New UTF8Encoding
        Dim data() As Byte = enc.GetBytes(cadena)
        Dim result() As Byte

        Dim sha As New SHA1CryptoServiceProvider

        result = sha.ComputeHash(data)

        Dim sb As New StringBuilder
        Dim max As Int32 = result.Length



        For i As Integer = 0 To max - 1


            'Convertimos los valores en hexadecimal
            'cuando tiene una cifra hay que rellenarlo con cero
            'para que siempre ocupen dos dígitos.
            If (result(i) < 16) Then
                sb.Append("0")
            End If

            sb.Append(result(i).ToString("x"))


        Next


        'Devolvemos la cadena con el hash en mayúsculas para que quede más chuli :)
        generarClaveSHA1 = sb.ToString().ToUpper()
        password_cifrado = sb.ToString().ToUpper()

    End Function


    Private Sub ComboDepartamento_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboDepartamento.SelectedIndexChanged
        cargarLocalidades()
    End Sub
    Public Sub cargarLocalidades()
        Dim l As New dLocalidad
        Dim lista As New ArrayList
        Dim iddepartamento As dDepartamento = CType(ComboDepartamento.SelectedItem, dDepartamento)
        If Not iddepartamento Is Nothing Then
            Dim texto As Long = iddepartamento.ID
            lista = l.listarpordepartamento(texto)
            ComboLocalidad.Items.Clear()
            If Not lista Is Nothing Then
                If lista.Count > 0 Then
                    For Each l In lista
                        ComboLocalidad.Items.Add(l)
                    Next
                End If
            End If
        End If
    End Sub

    Private Sub ButtonFiltrar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonFiltrar.Click
        Dim v As New FormFiltrarProductor(Usuario)
        v.Show()
    End Sub

    Private Sub ButtonBorrar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonBorrar.Click

    End Sub

    Private Sub ButtonSalir_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSalir.Click
        Me.Close()
    End Sub

    Private Sub ButtonNuevo2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonNuevo2.Click
        cargarLista()
        limpiar()
    End Sub

    Private Sub ButtonNuevo3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonNuevo3.Click
        cargarLista()
        limpiar()
    End Sub

    Private Sub ButtonGuardar2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonGuardar2.Click
        guardar()
        limpiar()
        cargarLista()
    End Sub

    Private Sub ButtonGuardar3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonGuardar3.Click
        guardar()
        limpiar()
        cargarLista()
    End Sub

    Private Sub ButtonEmpresa2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonEmpresa2.Click
        If TextId.Text <> "" Then
            Dim idcliente As Long = TextId.Text.Trim
            Dim v As New FormProductorEmpresa(Usuario, idcliente)
            v.ShowDialog()
        End If
    End Sub

    Private Sub ButtonEmpresa3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonEmpresa3.Click
        If TextId.Text <> "" Then
            Dim idcliente As Long = TextId.Text.Trim
            Dim v As New FormProductorEmpresa(Usuario, idcliente)
            v.ShowDialog()
        End If
    End Sub

    Private Sub ButtonSalir2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSalir2.Click
        Me.Close()
    End Sub

    Private Sub ButtonSalir3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonSalir3.Click
        Me.Close()
    End Sub

    Private Sub ButtonFiltrar2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonFiltrar2.Click
        Dim v As New FormFiltrarProductor(Usuario)
        v.Show()
    End Sub

    Private Sub ButtonFiltrar3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonFiltrar3.Click
        Dim v As New FormFiltrarProductor(Usuario)
        v.Show()
    End Sub

    Private Sub ListClientes_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ListClientes_SelectedIndexChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListClientes.SelectedIndexChanged
        If ListClientes.SelectedItems.Count = 1 Then
            Dim cli As dCliente = CType(ListClientes.SelectedItem, dCliente)
            TextId.Text = cli.ID
            TextNombre.Text = cli.NOMBRE
            TextEmail.Text = cli.EMAIL
            TextEnvio.Text = cli.ENVIO
            TextUsuarioWeb.Text = cli.USUARIO_WEB
            TextRazonSocial.Text = cli.RAZON_SOCIAL
            TextCelular.Text = cli.CELULAR
            TextRut.Text = cli.RUT
            TextFigaro.Text = cli.CODIGOFIGARO
            ComboTipoUsuario.SelectedItem = Nothing
            Dim tu As dTipoUsuario
            For Each tu In ComboTipoUsuario.Items
                If tu.ID = cli.TIPOUSUARIO Then
                    ComboTipoUsuario.SelectedItem = tu
                    Exit For
                End If
            Next
            TextDireccion.Text = cli.DIRECCION
            TextTelefono.Text = cli.TELEFONO
            TextFax.Text = cli.FAX
            TextDicose.Text = cli.DICOSE
            ComboDepartamento.SelectedItem = Nothing
            Dim d As dDepartamento
            For Each d In ComboDepartamento.Items
                If d.ID = cli.IDDEPARTAMENTO Then
                    ComboDepartamento.SelectedItem = d
                    Exit For
                End If
            Next
            ComboLocalidad.SelectedItem = Nothing
            Dim l As dLocalidad
            For Each l In ComboLocalidad.Items
                If l.ID = cli.IDLOCALIDAD Then
                    ComboLocalidad.SelectedItem = l
                    Exit For
                End If
            Next
            ComboTecnicos.SelectedItem = Nothing
            Dim t As dTecnicos
            For Each t In ComboTecnicos.Items
                If t.ID = cli.TECNICO Then
                    ComboTecnicos.SelectedItem = t
                    Exit For
                End If
            Next
            ComboTecnicos2.SelectedItem = Nothing
            Dim t2 As dTecnicos
            For Each t2 In ComboTecnicos2.Items
                If t2.ID = cli.TECNICO2 Then
                    ComboTecnicos2.SelectedItem = t2
                    Exit For
                End If
            Next
            ComboTecnicos3.SelectedItem = Nothing
            Dim t3 As dTecnicos
            For Each t3 In ComboTecnicos3.Items
                If t3.ID = cli.TECNICO3 Then
                    ComboTecnicos3.SelectedItem = t3
                    Exit For
                End If
            Next
            ComboAgencia.SelectedItem = Nothing
            Dim a As dEmpresaT
            For Each a In ComboAgencia.Items
                If a.ID = cli.IDAGENCIA Then
                    ComboAgencia.SelectedItem = a
                    Exit For
                End If
            Next
            If cli.CONTRATO = 1 Then
                CheckContrato.Checked = True
            Else
                CheckContrato.Checked = False
            End If
            If cli.SOCIO = 1 Then
                CheckSocio.Checked = True
            Else
                CheckSocio.Checked = False
            End If
            If cli.NOUSAR = 1 Then
                CheckNousar.Checked = True
            Else
                CheckNousar.Checked = False
            End If
            If cli.CARAVANAS = 1 Then
                CheckCaravanas.Checked = True
            Else
                CheckCaravanas.Checked = False
            End If
            If cli.PROLESA = 1 Then
                CheckProlesa.Checked = True
            Else
                CheckProlesa.Checked = False
            End If
            TextObservaciones.Text = cli.OBSERVACIONES
            TextFacNombre.Text = cli.FAC_NOMBRE
            TextFacRsocial.Text = cli.FAC_RSOCIAL
            TextFacRut.Text = cli.FAC_RUT
            TextFacDireccion.Text = cli.FAC_DIRECCION
            TextFacLocalidad.Text = cli.FAC_LOCALIDAD
            ComboFacDepartamento.SelectedItem = Nothing
            Dim dd As dDepartamento
            For Each dd In ComboFacDepartamento.Items
                If dd.ID = cli.FAC_DEPARTAMENTO Then
                    ComboFacDepartamento.SelectedItem = dd
                    Exit For
                End If
            Next
            TextFacCpostal.Text = cli.FAC_CPOSTAL
            ComboFacGiro.SelectedItem = Nothing
            Dim g As dGiro
            For Each g In ComboFacGiro.Items
                If g.ID = cli.FAC_GIRO Then
                    ComboFacGiro.SelectedItem = g
                    Exit For
                End If
            Next
            TextFacTelefonos.Text = cli.FAC_TELEFONOS
            TextFacFax.Text = cli.FAC_FAX
            TextFacEmail.Text = cli.FAC_EMAIL
            TextFacContacto.Text = cli.FAC_CONTACTO
            TextFacObservaciones.Text = cli.FAC_OBSERVACIONES
            NumericLista.Value = cli.FAC_LISTA
            TextId.Focus()
        End If
    End Sub

    Private Sub CheckProlesa_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckProlesa.CheckedChanged
        prolesa()
    End Sub
    Private Sub prolesa()
        If CheckProlesa.Checked = True Then
            ComboProlesa.Enabled = True
        Else
            ComboProlesa.Enabled = False
        End If
    End Sub
End Class