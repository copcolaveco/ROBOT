Public Class pCliente
    Inherits Conectoras.ConexionMySQL
    Public Function guardar(ByVal o As Object) As Boolean
        Dim obj As dCliente = CType(o, dCliente)
        Dim sql As String = "INSERT INTO cliente (id, nombre, email, nombre_email1, email1, nombre_email2, email2, envio, usuario_web, nombre_celular1, celular, nombre_celular2, celular2, codigofigaro, tipousuario, direccion, nombre_telefono1, telefono1, nombre_telefono2, telefono2, fax, dicose, iddepartamento, idlocalidad, tecnico1, tecnico2, idagencia, contrato, socio, nousar, caravanas, prolesa, prolesasuc, prolesamat, observaciones, fac_rsocial, fac_cedula, fac_rut, fac_direccion, fac_localidad, fac_departamento, fac_cpostal, fac_giro, cob_nombre_telefono1, fac_telefonos, cob_nombre_telefono2, cob_telefono2, cob_nombre_celular1, cob_celular1, cob_nombre_celular2, cob_celular2, cob_nombre_email1, cob_email1, cob_nombre_email2, cob_email2, fac_fax, fac_email, fac_contacto, fac_observaciones, fac_lista, fac_contado, not_email_frascos1, not_email_frascos2, not_email_muestras1, not_email_muestras2, not_email_analisis1, not_email_analisis2, not_email_general1,  not_email_general2) VALUES (" & obj.ID & ", '" & obj.NOMBRE & "', '" & obj.EMAIL & "','" & obj.NOMBRE_EMAIL1 & "','" & obj.EMAIL1 & "', '" & obj.NOMBRE_EMAIL2 & "','" & obj.EMAIL2 & "', '" & obj.ENVIO & "','" & obj.USUARIO_WEB & "','" & obj.NOMBRE_CELULAR1 & "', '" & obj.CELULAR & "', '" & obj.NOMBRE_CELULAR2 & "', '" & obj.CELULAR2 & "', '" & obj.CODIGOFIGARO & "'," & obj.TIPOUSUARIO & ",'" & obj.DIRECCION & "','" & obj.NOMBRE_TELEFONO1 & "','" & obj.TELEFONO1 & "','" & obj.NOMBRE_TELEFONO2 & "','" & obj.TELEFONO2 & "','" & obj.FAX & "','" & obj.DICOSE & "'," & obj.IDDEPARTAMENTO & "," & obj.IDLOCALIDAD & "," & obj.TECNICO1 & "," & obj.TECNICO2 & "," & obj.IDAGENCIA & ", " & obj.CONTRATO & ", " & obj.SOCIO & ", " & obj.NOUSAR & ", " & obj.CARAVANAS & ", " & obj.PROLESA & "," & obj.PROLESASUC & "," & obj.PROLESAMAT & ", '" & obj.OBSERVACIONES & "', '" & obj.FAC_RSOCIAL & "', '" & obj.FAC_CEDULA & "','" & obj.FAC_RUT & "', '" & obj.FAC_DIRECCION & "', '" & obj.FAC_LOCALIDAD & "', " & obj.FAC_DEPARTAMENTO & ", '" & obj.FAC_CPOSTAL & "', " & obj.FAC_GIRO & ", '" & obj.COB_NOMBRE_TELEFONO1 & "', '" & obj.FAC_TELEFONOS & "','" & obj.COB_NOMBRE_TELEFONO2 & "', '" & obj.COB_TELEFONO2 & "', '" & obj.COB_NOMBRE_CELULAR1 & "', '" & obj.COB_CELULAR1 & "', '" & obj.COB_NOMBRE_CELULAR2 & "', '" & obj.COB_CELULAR2 & "', '" & obj.COB_NOMBRE_EMAIL1 & "', '" & obj.COB_EMAIL1 & "','" & obj.COB_NOMBRE_EMAIL2 & "', '" & obj.COB_EMAIL2 & "','" & obj.FAC_FAX & "', '" & obj.FAC_EMAIL & "', '" & obj.FAC_CONTACTO & "', '" & obj.FAC_OBSERVACIONES & "', " & obj.FAC_LISTA & ", " & obj.FAC_CONTADO & ", '" & obj.NOT_EMAIL_FRASCOS1 & "', '" & obj.NOT_EMAIL_FRASCOS2 & "', '" & obj.NOT_EMAIL_MUESTRAS1 & "', '" & obj.NOT_EMAIL_MUESTRAS2 & "', '" & obj.NOT_EMAIL_ANALISIS1 & "', '" & obj.NOT_EMAIL_ANALISIS2 & "', '" & obj.NOT_EMAIL_GENERAL1 & "', '" & obj.NOT_EMAIL_GENERAL2 & "')"

        Dim lista As New ArrayList
        lista.Add(sql)

     

        Return EjecutarTransaccion(lista)
    End Function
    Public Function modificar(ByVal o As Object) As Boolean
        Dim obj As dCliente = CType(o, dCliente)
        Dim sql As String = "UPDATE cliente SET nombre= '" & obj.NOMBRE & "', email= '" & obj.EMAIL & "', nombre_email1='" & obj.NOMBRE_EMAIL1 & "',email1='" & obj.EMAIL1 & "', nombre_email2='" & obj.NOMBRE_EMAIL2 & "',email2='" & obj.EMAIL2 & "', envio='" & obj.ENVIO & "',usuario_web='" & obj.USUARIO_WEB & "',nombre_celular1='" & obj.NOMBRE_CELULAR1 & "', celular= '" & obj.CELULAR & "', nombre_celular2= '" & obj.NOMBRE_CELULAR2 & "', celular2= '" & obj.CELULAR2 & "', codigofigaro= '" & obj.CODIGOFIGARO & "',tipousuario= " & obj.TIPOUSUARIO & ",direccion='" & obj.DIRECCION & "',nombre_telefono1='" & obj.NOMBRE_TELEFONO1 & "',telefono1='" & obj.TELEFONO1 & "',nombre_telefono2='" & obj.NOMBRE_TELEFONO2 & "',telefono2='" & obj.TELEFONO2 & "',fax='" & obj.FAX & "',dicose='" & obj.DICOSE & "',iddepartamento=" & obj.IDDEPARTAMENTO & ",idlocalidad=" & obj.IDLOCALIDAD & ",tecnico1=" & obj.TECNICO1 & ",tecnico2=" & obj.TECNICO2 & ",idagencia=" & obj.IDAGENCIA & ", contrato=" & obj.CONTRATO & ", socio=" & obj.SOCIO & ", nousar=" & obj.NOUSAR & ", caravanas=" & obj.CARAVANAS & ", prolesa=" & obj.PROLESA & ",prolesasuc=" & obj.PROLESASUC & ", prolesamat=" & obj.PROLESAMAT & ",observaciones='" & obj.OBSERVACIONES & "', fac_rsocial='" & obj.FAC_RSOCIAL & "', fac_cedula='" & obj.FAC_CEDULA & "',fac_rut='" & obj.FAC_RUT & "', fac_direccion='" & obj.FAC_DIRECCION & "', fac_localidad='" & obj.FAC_LOCALIDAD & "', fac_departamento=" & obj.FAC_DEPARTAMENTO & ", fac_cpostal='" & obj.FAC_CPOSTAL & "', fac_giro=" & obj.FAC_GIRO & ", cob_nombre_telefono1='" & obj.COB_NOMBRE_TELEFONO1 & "', fac_telefonos='" & obj.FAC_TELEFONOS & "',cob_nombre_telefono2='" & obj.COB_NOMBRE_TELEFONO2 & "', cob_telefono2='" & obj.COB_TELEFONO2 & "', cob_nombre_celular1='" & obj.COB_NOMBRE_CELULAR1 & "', cob_celular1='" & obj.COB_CELULAR1 & "', cob_nombre_celular2='" & obj.COB_NOMBRE_CELULAR2 & "', cob_celular2= '" & obj.COB_CELULAR2 & "', cob_nombre_email1='" & obj.COB_NOMBRE_EMAIL1 & "', cob_email1='" & obj.COB_EMAIL1 & "',cob_nombre_email2='" & obj.COB_NOMBRE_EMAIL2 & "', cob_email2='" & obj.COB_EMAIL2 & "',fac_fax='" & obj.FAC_FAX & "', fac_email='" & obj.FAC_EMAIL & "', fac_contacto='" & obj.FAC_CONTACTO & "', fac_observaciones='" & obj.FAC_OBSERVACIONES & "', fac_lista=" & obj.FAC_LISTA & ", fac_contado=" & obj.FAC_CONTADO & ", not_email_frascos1='" & obj.NOT_EMAIL_FRASCOS1 & "', not_email_frascos2= '" & obj.NOT_EMAIL_FRASCOS2 & "', not_email_muestras1='" & obj.NOT_EMAIL_MUESTRAS1 & "', not_email_muestras2='" & obj.NOT_EMAIL_MUESTRAS2 & "', not_email_analisis1='" & obj.NOT_EMAIL_ANALISIS1 & "', not_email_analisis2='" & obj.NOT_EMAIL_ANALISIS2 & "', not_email_general1='" & obj.NOT_EMAIL_GENERAL1 & "', not_email_general2='" & obj.NOT_EMAIL_GENERAL2 & "' WHERE id = " & obj.ID & ""

        Dim lista As New ArrayList
        lista.Add(sql)

      
        Return EjecutarTransaccion(lista)
    End Function
    Public Function modificardesdeweb(ByVal o As Object) As Boolean
        Dim obj As dCliente = CType(o, dCliente)
        Dim sql As String = "UPDATE cliente SET nombre= '" & obj.NOMBRE & "', email= '" & obj.EMAIL & "', nombre_email1='" & obj.NOMBRE_EMAIL1 & "',email1='" & obj.EMAIL1 & "', nombre_email2='" & obj.NOMBRE_EMAIL2 & "',email2='" & obj.EMAIL2 & "', envio='" & obj.ENVIO & "',usuario_web='" & obj.USUARIO_WEB & "',nombre_celular1='" & obj.NOMBRE_CELULAR1 & "', celular= '" & obj.CELULAR & "', nombre_celular2= '" & obj.NOMBRE_CELULAR2 & "', celular2= '" & obj.CELULAR2 & "', direccion='" & obj.DIRECCION & "',nombre_telefono1='" & obj.NOMBRE_TELEFONO1 & "',telefono1='" & obj.TELEFONO1 & "',nombre_telefono2='" & obj.NOMBRE_TELEFONO2 & "',telefono2='" & obj.TELEFONO2 & "',dicose='" & obj.DICOSE & "',idagencia=" & obj.IDAGENCIA & ", fac_rsocial='" & obj.FAC_RSOCIAL & "', fac_cedula='" & obj.FAC_CEDULA & "',fac_rut='" & obj.FAC_RUT & "', fac_direccion='" & obj.FAC_DIRECCION & "', fac_localidad='" & obj.FAC_LOCALIDAD & "', fac_departamento=" & obj.FAC_DEPARTAMENTO & ", cob_nombre_telefono1='" & obj.COB_NOMBRE_TELEFONO1 & "', fac_telefonos='" & obj.FAC_TELEFONOS & "',cob_nombre_telefono2='" & obj.COB_NOMBRE_TELEFONO2 & "', cob_telefono2='" & obj.COB_TELEFONO2 & "', cob_nombre_celular1='" & obj.COB_NOMBRE_CELULAR1 & "', cob_celular1='" & obj.COB_CELULAR1 & "', cob_nombre_celular2='" & obj.COB_NOMBRE_CELULAR2 & "', cob_celular2= '" & obj.COB_CELULAR2 & "', cob_nombre_email1='" & obj.COB_NOMBRE_EMAIL1 & "', cob_email1='" & obj.COB_EMAIL1 & "',cob_nombre_email2='" & obj.COB_NOMBRE_EMAIL2 & "', cob_email2='" & obj.COB_EMAIL2 & "', fac_email='" & obj.FAC_EMAIL & "', not_email_frascos1='" & obj.NOT_EMAIL_FRASCOS1 & "', not_email_frascos2= '" & obj.NOT_EMAIL_FRASCOS2 & "', not_email_muestras1='" & obj.NOT_EMAIL_MUESTRAS1 & "', not_email_muestras2='" & obj.NOT_EMAIL_MUESTRAS2 & "', not_email_analisis1='" & obj.NOT_EMAIL_ANALISIS1 & "', not_email_analisis2='" & obj.NOT_EMAIL_ANALISIS2 & "', not_email_general1='" & obj.NOT_EMAIL_GENERAL1 & "', not_email_general2='" & obj.NOT_EMAIL_GENERAL2 & "' WHERE id = " & obj.ID & ""

        Dim lista As New ArrayList
        lista.Add(sql)


        Return EjecutarTransaccion(lista)
    End Function
    'Public Function modificarsolog2000(ByVal o As Object, ByVal usuario As dUsuario) As Boolean
    '    Dim obj As dCliente = CType(o, dCliente)
    '    Dim sql As String = "UPDATE cliente SET nombre='" & obj.NOMBRE & "', fac_rsocial='" & obj.FAC_RSOCIAL & "', fac_rut='" & obj.FAC_RUT & "',  fac_direccion='" & obj.FAC_DIRECCION & "', fac_localidad='" & obj.FAC_LOCALIDAD & "', fac_departamento=" & obj.FAC_DEPARTAMENTO & ", fac_cpostal='" & obj.FAC_CPOSTAL & "', fac_giro=" & obj.FAC_GIRO & ", fac_telefonos='" & obj.FAC_TELEFONOS & "', fac_fax='" & obj.FAC_FAX & "', fac_email='" & obj.FAC_EMAIL & "', fac_contacto='" & obj.FAC_CONTACTO & "', fac_observaciones='" & obj.FAC_OBSERVACIONES & "', fac_lista=" & obj.FAC_LISTA & " WHERE id = " & obj.ID & ""

    '    Dim lista As New ArrayList
    '    lista.Add(sql)

    '    Dim sqlAccion As String = "INSERT INTO actividad (act_fecha, act_tabla, act_accion, act_registro, u_id) " _
    '                             & "VALUES (now(), 'cliente', 'modificación', " & obj.ID & ", " & usuario.ID & ")"
    '    lista.Add(sqlAccion)

    '    Return EjecutarTransaccion(lista)
    'End Function
    Public Function eliminar(ByVal o As Object) As Boolean
        Dim obj As dCliente = CType(o, dCliente)
        Dim sql As String = "DELETE FROM cliente WHERE ID = " & obj.ID & ""

        Dim lista As New ArrayList
        lista.Add(sql)


        Return EjecutarTransaccion(lista)
    End Function
    Public Function buscar(ByVal o As Object) As dCliente
        Dim obj As dcliente = CType(o, dcliente)
        Dim p As New dcliente
        Try
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL("SELECT id, nombre, ifnull(email,''), ifnull(nombre_email1,''), ifnull(email1,''), ifnull(nombre_email2,''), ifnull(email2,''), ifnull(envio,''), ifnull(usuario_web,''), ifnull(nombre_celular1,''), ifnull(celular,''), ifnull(nombre_celular2,''), ifnull(celular2,''), ifnull(codigofigaro,''), ifnull(tipousuario,1), ifnull(direccion,''), ifnull(nombre_telefono1,''), ifnull(telefono1,''), ifnull(nombre_telefono2,''), ifnull(telefono2,''), ifnull(fax,''), ifnull(dicose,''), ifnull(iddepartamento,999), ifnull(idlocalidad,999), ifnull(tecnico1,3197), ifnull(tecnico2,3197), ifnull(idagencia,8), contrato, socio, nousar, caravanas, prolesa, ifnull(prolesasuc,0), prolesamat, ifnull(observaciones,''), ifnull(fac_rsocial,''), ifnull(fac_cedula,''), ifnull(fac_rut,''), ifnull(fac_direccion,''), ifnull(fac_localidad,''), ifnull(fac_departamento,0), ifnull(fac_cpostal,''), ifnull(fac_giro,0), ifnull(cob_nombre_telefono1,''), ifnull(fac_telefonos,''), ifnull(cob_nombre_telefono2,''), ifnull(cob_telefono2,''), ifnull(cob_nombre_celular1,''), ifnull(cob_celular1,''), ifnull(cob_nombre_celular2,''), ifnull(cob_celular2,''), ifnull(cob_nombre_email1,''), ifnull(cob_email1,''), ifnull(cob_nombre_email2,''), ifnull(cob_email2,''), ifnull(fac_fax,''), ifnull(fac_email,''), ifnull(fac_contacto,''), ifnull(fac_observaciones,''), fac_lista, fac_contado, ifnull(not_email_frascos1,''), ifnull(not_email_frascos2,''), ifnull(not_email_muestras1,''), ifnull(not_email_muestras2,''), ifnull(not_email_analisis1,''), ifnull(not_email_analisis2,''), ifnull(not_email_general1,''), ifnull(not_email_general2,'') FROM cliente WHERE id = " & obj.ID & "")
            'Ds = Me.EjecutarSQL("SELECT id, nombre, ifnull(email,''), ifnull(envio,''), ifnull(usuario_web,''), ifnull(celular,''), ifnull(codigofigaro,''), ifnull(tipousuario,1), ifnull(direccion,''), ifnull(telefono,''), ifnull(fax,''), ifnull(dicose,''), ifnull(iddepartamento,999), ifnull(idlocalidad,999), ifnull(tecnico,3282),ifnull(tecnico2,3282), ifnull(idagencia,8), contrato, socio, nousar, caravanas, prolesa, ifnull(prolesasuc, 0), ifnull(observaciones,''), ifnull(fac_rsocial,''), ifnull(fac_rut,''), ifnull(fac_direccion,''), ifnull(fac_localidad,''), ifnull(fac_departamento,0), ifnull(fac_cpostal,''), ifnull(fac_giro,0), ifnull(fac_telefonos,''), ifnull(fac_fax,''), ifnull(fac_email,''), ifnull(fac_contacto,''), ifnull(fac_observaciones,''), fac_lista, fac_contado FROM cliente WHERE id = " & obj.ID & "")

            If Ds.Tables(0).Rows.Count > 0 Then
                Dim unaFila As DataRow
                unaFila = Ds.Tables(0).Rows(0)
                p.ID = CType(unaFila.Item(0), Long)
                p.NOMBRE = CType(unaFila.Item(1), String)
                p.EMAIL = CType(unaFila.Item(2), String)
                p.NOMBRE_EMAIL1 = CType(unaFila.Item(3), String)
                p.EMAIL1 = CType(unaFila.Item(4), String)
                p.NOMBRE_EMAIL2 = CType(unaFila.Item(5), String)
                p.EMAIL2 = CType(unaFila.Item(6), String)
                p.ENVIO = CType(unaFila.Item(7), String)
                p.USUARIO_WEB = CType(unaFila.Item(8), String)
                p.NOMBRE_CELULAR1 = CType(unaFila.Item(9), String)
                p.CELULAR = CType(unaFila.Item(10), String)
                p.NOMBRE_CELULAR2 = CType(unaFila.Item(11), String)
                p.CELULAR2 = CType(unaFila.Item(12), String)
                p.CODIGOFIGARO = CType(unaFila.Item(13), String)
                p.TIPOUSUARIO = CType(unaFila.Item(14), Integer)
                p.DIRECCION = CType(unaFila.Item(15), String)
                p.NOMBRE_TELEFONO1 = CType(unaFila.Item(16), String)
                p.TELEFONO1 = CType(unaFila.Item(17), String)
                p.NOMBRE_TELEFONO2 = CType(unaFila.Item(18), String)
                p.TELEFONO2 = CType(unaFila.Item(19), String)
                p.FAX = CType(unaFila.Item(20), String)
                p.DICOSE = CType(unaFila.Item(21), String)
                p.IDDEPARTAMENTO = CType(unaFila.Item(22), Integer)
                p.IDLOCALIDAD = CType(unaFila.Item(23), Integer)
                p.TECNICO1 = CType(unaFila.Item(24), Long)
                p.TECNICO2 = CType(unaFila.Item(25), Long)
                p.IDAGENCIA = CType(unaFila.Item(26), Integer)
                p.CONTRATO = CType(unaFila.Item(27), Integer)
                p.SOCIO = CType(unaFila.Item(28), Integer)
                p.NOUSAR = CType(unaFila.Item(29), Integer)
                p.CARAVANAS = CType(unaFila.Item(30), Integer)
                p.PROLESA = CType(unaFila.Item(31), Integer)
                p.PROLESASUC = CType(unaFila.Item(32), Integer)
                p.PROLESAMAT = CType(unaFila.Item(33), Long)
                p.OBSERVACIONES = CType(unaFila.Item(34), String)
                p.FAC_RSOCIAL = CType(unaFila.Item(35), String)
                p.FAC_CEDULA = CType(unaFila.Item(36), String)
                p.FAC_RUT = CType(unaFila.Item(37), String)
                p.FAC_DIRECCION = CType(unaFila.Item(38), String)
                p.FAC_LOCALIDAD = CType(unaFila.Item(39), String)
                p.FAC_DEPARTAMENTO = CType(unaFila.Item(40), Integer)
                p.FAC_CPOSTAL = CType(unaFila.Item(41), String)
                p.FAC_GIRO = CType(unaFila.Item(42), Integer)
                p.COB_NOMBRE_TELEFONO1 = CType(unaFila.Item(43), String)
                p.FAC_TELEFONOS = CType(unaFila.Item(44), String)
                p.COB_NOMBRE_TELEFONO2 = CType(unaFila.Item(45), String)
                p.COB_TELEFONO2 = CType(unaFila.Item(46), String)
                p.COB_NOMBRE_CELULAR1 = CType(unaFila.Item(47), String)
                p.COB_CELULAR1 = CType(unaFila.Item(48), String)
                p.COB_NOMBRE_CELULAR2 = CType(unaFila.Item(49), String)
                p.COB_CELULAR2 = CType(unaFila.Item(50), String)
                p.COB_NOMBRE_EMAIL1 = CType(unaFila.Item(51), String)
                p.COB_EMAIL1 = CType(unaFila.Item(52), String)
                p.COB_NOMBRE_EMAIL2 = CType(unaFila.Item(53), String)
                p.COB_EMAIL2 = CType(unaFila.Item(54), String)
                p.FAC_FAX = CType(unaFila.Item(55), String)
                p.FAC_EMAIL = CType(unaFila.Item(56), String)
                p.FAC_CONTACTO = CType(unaFila.Item(57), String)
                p.FAC_OBSERVACIONES = CType(unaFila.Item(58), String)
                p.FAC_LISTA = CType(unaFila.Item(59), Integer)
                p.FAC_CONTADO = CType(unaFila.Item(60), Integer)
                p.NOT_EMAIL_FRASCOS1 = CType(unaFila.Item(61), String)
                p.NOT_EMAIL_FRASCOS2 = CType(unaFila.Item(62), String)
                p.NOT_EMAIL_MUESTRAS1 = CType(unaFila.Item(63), String)
                p.NOT_EMAIL_MUESTRAS2 = CType(unaFila.Item(64), String)
                p.NOT_EMAIL_ANALISIS1 = CType(unaFila.Item(65), String)
                p.NOT_EMAIL_ANALISIS2 = CType(unaFila.Item(66), String)
                p.NOT_EMAIL_GENERAL1 = CType(unaFila.Item(67), String)
                p.NOT_EMAIL_GENERAL2 = CType(unaFila.Item(68), String)
                Return p
            End If
            Return Nothing
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function buscarultimo(ByVal o As Object) As dCliente
        Dim obj As dCliente = CType(o, dCliente)
        Dim p As New dCliente
        Try
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL("SELECT id, nombre, ifnull(email,''), ifnull(nombre_email1,''), ifnull(email1,''), ifnull(nombre_email2,''), ifnull(email2,''), ifnull(envio,''), ifnull(usuario_web,''), ifnull(nombre_celular1,''), ifnull(celular,''), ifnull(nombre_celular2,''), ifnull(celular2,''), ifnull(codigofigaro,''), ifnull(tipousuario,1), ifnull(direccion,''), ifnull(nombre_telefono1,''), ifnull(telefono1,''), ifnull(nombre_telefono2,''), ifnull(telefono2,''), ifnull(fax,''), ifnull(dicose,''), ifnull(iddepartamento,999), ifnull(idlocalidad,999), ifnull(tecnico1,3197), ifnull(tecnico2,3197), ifnull(idagencia,8), contrato, socio, nousar, caravanas, prolesa, ifnull(prolesasuc,0), prolesamat, ifnull(observaciones,''), ifnull(fac_rsocial,''), ifnull(fac_cedula,''), ifnull(fac_rut,''), ifnull(fac_direccion,''), ifnull(fac_localidad,''), ifnull(fac_departamento,0), ifnull(fac_cpostal,''), ifnull(fac_giro,0), ifnull(cob_nombre_telefono1,''), ifnull(fac_telefonos,''), ifnull(cob_nombre_telefono2,''), ifnull(cob_telefono2,''), ifnull(cob_nombre_celular1,''), ifnull(cob_celular1,''), ifnull(cob_nombre_celular2,''), ifnull(cob_celular2,''), ifnull(cob_nombre_email1,''), ifnull(cob_email1,''), ifnull(cob_nombre_email2,''), ifnull(cob_email2,''), ifnull(fac_fax,''), ifnull(fac_email,''), ifnull(fac_contacto,''), ifnull(fac_observaciones,''), fac_lista, fac_contado, ifnull(not_email_frascos1,''), ifnull(not_email_frascos2,''), ifnull(not_email_muestras1,''), ifnull(not_email_muestras2,''), ifnull(not_email_analisis1,''), ifnull(not_email_analisis2,''), ifnull(not_email_general1,''), ifnull(not_email_general2,'') FROM cliente WHERE id = (SELECT MAX(id) FROM cliente)")

            If Ds.Tables(0).Rows.Count > 0 Then
                Dim unaFila As DataRow
                unaFila = Ds.Tables(0).Rows(0)
                p.ID = CType(unaFila.Item(0), Long)
                p.NOMBRE = CType(unaFila.Item(1), String)
                p.EMAIL = CType(unaFila.Item(2), String)
                p.NOMBRE_EMAIL1 = CType(unaFila.Item(3), String)
                p.EMAIL1 = CType(unaFila.Item(4), String)
                p.NOMBRE_EMAIL2 = CType(unaFila.Item(5), String)
                p.EMAIL2 = CType(unaFila.Item(6), String)
                p.ENVIO = CType(unaFila.Item(7), String)
                p.USUARIO_WEB = CType(unaFila.Item(8), String)
                p.NOMBRE_CELULAR1 = CType(unaFila.Item(9), String)
                p.CELULAR = CType(unaFila.Item(10), String)
                p.NOMBRE_CELULAR2 = CType(unaFila.Item(11), String)
                p.CELULAR2 = CType(unaFila.Item(12), String)
                p.CODIGOFIGARO = CType(unaFila.Item(13), String)
                p.TIPOUSUARIO = CType(unaFila.Item(14), Integer)
                p.DIRECCION = CType(unaFila.Item(15), String)
                p.NOMBRE_TELEFONO1 = CType(unaFila.Item(16), String)
                p.TELEFONO1 = CType(unaFila.Item(17), String)
                p.NOMBRE_TELEFONO2 = CType(unaFila.Item(18), String)
                p.TELEFONO2 = CType(unaFila.Item(19), String)
                p.FAX = CType(unaFila.Item(20), String)
                p.DICOSE = CType(unaFila.Item(21), String)
                p.IDDEPARTAMENTO = CType(unaFila.Item(22), Integer)
                p.IDLOCALIDAD = CType(unaFila.Item(23), Integer)
                p.TECNICO1 = CType(unaFila.Item(24), Long)
                p.TECNICO2 = CType(unaFila.Item(25), Long)
                p.IDAGENCIA = CType(unaFila.Item(26), Integer)
                p.CONTRATO = CType(unaFila.Item(27), Integer)
                p.SOCIO = CType(unaFila.Item(28), Integer)
                p.NOUSAR = CType(unaFila.Item(29), Integer)
                p.CARAVANAS = CType(unaFila.Item(30), Integer)
                p.PROLESA = CType(unaFila.Item(31), Integer)
                p.PROLESASUC = CType(unaFila.Item(32), Integer)
                p.PROLESAMAT = CType(unaFila.Item(33), Long)
                p.OBSERVACIONES = CType(unaFila.Item(34), String)
                p.FAC_RSOCIAL = CType(unaFila.Item(35), String)
                p.FAC_CEDULA = CType(unaFila.Item(36), String)
                p.FAC_RUT = CType(unaFila.Item(37), String)
                p.FAC_DIRECCION = CType(unaFila.Item(38), String)
                p.FAC_LOCALIDAD = CType(unaFila.Item(39), String)
                p.FAC_DEPARTAMENTO = CType(unaFila.Item(40), Integer)
                p.FAC_CPOSTAL = CType(unaFila.Item(41), String)
                p.FAC_GIRO = CType(unaFila.Item(42), Integer)
                p.COB_NOMBRE_TELEFONO1 = CType(unaFila.Item(43), String)
                p.FAC_TELEFONOS = CType(unaFila.Item(44), String)
                p.COB_NOMBRE_TELEFONO2 = CType(unaFila.Item(45), String)
                p.COB_TELEFONO2 = CType(unaFila.Item(46), String)
                p.COB_NOMBRE_CELULAR1 = CType(unaFila.Item(47), String)
                p.COB_CELULAR1 = CType(unaFila.Item(48), String)
                p.COB_NOMBRE_CELULAR2 = CType(unaFila.Item(49), String)
                p.COB_CELULAR2 = CType(unaFila.Item(50), String)
                p.COB_NOMBRE_EMAIL1 = CType(unaFila.Item(51), String)
                p.COB_EMAIL1 = CType(unaFila.Item(52), String)
                p.COB_NOMBRE_EMAIL2 = CType(unaFila.Item(53), String)
                p.COB_EMAIL2 = CType(unaFila.Item(54), String)
                p.FAC_FAX = CType(unaFila.Item(55), String)
                p.FAC_EMAIL = CType(unaFila.Item(56), String)
                p.FAC_CONTACTO = CType(unaFila.Item(57), String)
                p.FAC_OBSERVACIONES = CType(unaFila.Item(58), String)
                p.FAC_LISTA = CType(unaFila.Item(59), Integer)
                p.FAC_CONTADO = CType(unaFila.Item(60), Integer)
                p.NOT_EMAIL_FRASCOS1 = CType(unaFila.Item(61), String)
                p.NOT_EMAIL_FRASCOS2 = CType(unaFila.Item(62), String)
                p.NOT_EMAIL_MUESTRAS1 = CType(unaFila.Item(63), String)
                p.NOT_EMAIL_MUESTRAS2 = CType(unaFila.Item(64), String)
                p.NOT_EMAIL_ANALISIS1 = CType(unaFila.Item(65), String)
                p.NOT_EMAIL_ANALISIS2 = CType(unaFila.Item(66), String)
                p.NOT_EMAIL_GENERAL1 = CType(unaFila.Item(67), String)
                p.NOT_EMAIL_GENERAL2 = CType(unaFila.Item(68), String)
                Return p
            End If
            Return Nothing
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function buscarxCodigoFigaro(ByVal o As Object) As dCliente
        Dim obj As dCliente = CType(o, dCliente)
        Dim p As New dCliente
        Try
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL("SELECT id, nombre, ifnull(email,''), ifnull(nombre_email1,''), ifnull(email1,''), ifnull(nombre_email2,''), ifnull(email2,''), ifnull(envio,''), ifnull(usuario_web,''), ifnull(nombre_celular1,''), ifnull(celular,''), ifnull(nombre_celular2,''), ifnull(celular2,''), ifnull(codigofigaro,''), ifnull(tipousuario,1), ifnull(direccion,''), ifnull(nombre_telefono1,''), ifnull(telefono1,''), ifnull(nombre_telefono2,''), ifnull(telefono2,''), ifnull(fax,''), ifnull(dicose,''), ifnull(iddepartamento,999), ifnull(idlocalidad,999), ifnull(tecnico1,3197), ifnull(tecnico2,3197), ifnull(idagencia,8), contrato, socio, nousar, caravanas, prolesa, ifnull(prolesasuc,0), prolesamat, ifnull(observaciones,''), ifnull(fac_rsocial,''), ifnull(fac_cedula,''), ifnull(fac_rut,''), ifnull(fac_direccion,''), ifnull(fac_localidad,''), ifnull(fac_departamento,0), ifnull(fac_cpostal,''), ifnull(fac_giro,0), ifnull(cob_nombre_telefono1,''), ifnull(fac_telefonos,''), ifnull(cob_nombre_telefono2,''), ifnull(cob_telefono2,''), ifnull(cob_nombre_celular1,''), ifnull(cob_celular1,''), ifnull(cob_nombre_celular2,''), ifnull(cob_celular2,''), ifnull(cob_nombre_email1,''), ifnull(cob_email1,''), ifnull(cob_nombre_email2,''), ifnull(cob_email2,''), ifnull(fac_fax,''), ifnull(fac_email,''), ifnull(fac_contacto,''), ifnull(fac_observaciones,''), fac_lista, fac_contado, ifnull(not_email_frascos1,''), ifnull(not_email_frascos2,''), ifnull(not_email_muestras1,''), ifnull(not_email_muestras2,''), ifnull(not_email_analisis1,''), ifnull(not_email_analisis2,''), ifnull(not_email_general1,''), ifnull(not_email_general2,'') FROM cliente WHERE codigofigaro = '" & obj.CODIGOFIGARO & "'")

            If Ds.Tables(0).Rows.Count > 0 Then
                Dim unaFila As DataRow
                unaFila = Ds.Tables(0).Rows(0)
                p.ID = CType(unaFila.Item(0), Long)
                p.NOMBRE = CType(unaFila.Item(1), String)
                p.EMAIL = CType(unaFila.Item(2), String)
                p.NOMBRE_EMAIL1 = CType(unaFila.Item(3), String)
                p.EMAIL1 = CType(unaFila.Item(4), String)
                p.NOMBRE_EMAIL2 = CType(unaFila.Item(5), String)
                p.EMAIL2 = CType(unaFila.Item(6), String)
                p.ENVIO = CType(unaFila.Item(7), String)
                p.USUARIO_WEB = CType(unaFila.Item(8), String)
                p.NOMBRE_CELULAR1 = CType(unaFila.Item(9), String)
                p.CELULAR = CType(unaFila.Item(10), String)
                p.NOMBRE_CELULAR2 = CType(unaFila.Item(11), String)
                p.CELULAR2 = CType(unaFila.Item(12), String)
                p.CODIGOFIGARO = CType(unaFila.Item(13), String)
                p.TIPOUSUARIO = CType(unaFila.Item(14), Integer)
                p.DIRECCION = CType(unaFila.Item(15), String)
                p.NOMBRE_TELEFONO1 = CType(unaFila.Item(16), String)
                p.TELEFONO1 = CType(unaFila.Item(17), String)
                p.NOMBRE_TELEFONO2 = CType(unaFila.Item(18), String)
                p.TELEFONO2 = CType(unaFila.Item(19), String)
                p.FAX = CType(unaFila.Item(20), String)
                p.DICOSE = CType(unaFila.Item(21), String)
                p.IDDEPARTAMENTO = CType(unaFila.Item(22), Integer)
                p.IDLOCALIDAD = CType(unaFila.Item(23), Integer)
                p.TECNICO1 = CType(unaFila.Item(24), Long)
                p.TECNICO2 = CType(unaFila.Item(25), Long)
                p.IDAGENCIA = CType(unaFila.Item(26), Integer)
                p.CONTRATO = CType(unaFila.Item(27), Integer)
                p.SOCIO = CType(unaFila.Item(28), Integer)
                p.NOUSAR = CType(unaFila.Item(29), Integer)
                p.CARAVANAS = CType(unaFila.Item(30), Integer)
                p.PROLESA = CType(unaFila.Item(31), Integer)
                p.PROLESASUC = CType(unaFila.Item(32), Integer)
                p.PROLESAMAT = CType(unaFila.Item(33), Long)
                p.OBSERVACIONES = CType(unaFila.Item(34), String)
                p.FAC_RSOCIAL = CType(unaFila.Item(35), String)
                p.FAC_CEDULA = CType(unaFila.Item(36), String)
                p.FAC_RUT = CType(unaFila.Item(37), String)
                p.FAC_DIRECCION = CType(unaFila.Item(38), String)
                p.FAC_LOCALIDAD = CType(unaFila.Item(39), String)
                p.FAC_DEPARTAMENTO = CType(unaFila.Item(40), Integer)
                p.FAC_CPOSTAL = CType(unaFila.Item(41), String)
                p.FAC_GIRO = CType(unaFila.Item(42), Integer)
                p.COB_NOMBRE_TELEFONO1 = CType(unaFila.Item(43), String)
                p.FAC_TELEFONOS = CType(unaFila.Item(44), String)
                p.COB_NOMBRE_TELEFONO2 = CType(unaFila.Item(45), String)
                p.COB_TELEFONO2 = CType(unaFila.Item(46), String)
                p.COB_NOMBRE_CELULAR1 = CType(unaFila.Item(47), String)
                p.COB_CELULAR1 = CType(unaFila.Item(48), String)
                p.COB_NOMBRE_CELULAR2 = CType(unaFila.Item(49), String)
                p.COB_CELULAR2 = CType(unaFila.Item(50), String)
                p.COB_NOMBRE_EMAIL1 = CType(unaFila.Item(51), String)
                p.COB_EMAIL1 = CType(unaFila.Item(52), String)
                p.COB_NOMBRE_EMAIL2 = CType(unaFila.Item(53), String)
                p.COB_EMAIL2 = CType(unaFila.Item(54), String)
                p.FAC_FAX = CType(unaFila.Item(55), String)
                p.FAC_EMAIL = CType(unaFila.Item(56), String)
                p.FAC_CONTACTO = CType(unaFila.Item(57), String)
                p.FAC_OBSERVACIONES = CType(unaFila.Item(58), String)
                p.FAC_LISTA = CType(unaFila.Item(59), Integer)
                p.FAC_CONTADO = CType(unaFila.Item(60), Integer)
                p.NOT_EMAIL_FRASCOS1 = CType(unaFila.Item(61), String)
                p.NOT_EMAIL_FRASCOS2 = CType(unaFila.Item(62), String)
                p.NOT_EMAIL_MUESTRAS1 = CType(unaFila.Item(63), String)
                p.NOT_EMAIL_MUESTRAS2 = CType(unaFila.Item(64), String)
                p.NOT_EMAIL_ANALISIS1 = CType(unaFila.Item(65), String)
                p.NOT_EMAIL_ANALISIS2 = CType(unaFila.Item(66), String)
                p.NOT_EMAIL_GENERAL1 = CType(unaFila.Item(67), String)
                p.NOT_EMAIL_GENERAL2 = CType(unaFila.Item(68), String)
                Return p
            End If
            Return Nothing
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function buscarPorUsuarioWeb(ByVal o As Object) As dCliente
        Dim obj As dcliente = CType(o, dcliente)
        Dim p As New dcliente
        Try
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL("SELECT id, nombre, ifnull(email,''), ifnull(nombre_email1,''), ifnull(email1,''), ifnull(nombre_email2,''), ifnull(email2,''), ifnull(envio,''), ifnull(usuario_web,''), ifnull(nombre_celular1,''), ifnull(celular,''), ifnull(nombre_celular2,''), ifnull(celular2,''), ifnull(codigofigaro,''), ifnull(tipousuario,1), ifnull(direccion,''), ifnull(nombre_telefono1,''), ifnull(telefono1,''), ifnull(nombre_telefono2,''), ifnull(telefono2,''), ifnull(fax,''), ifnull(dicose,''), ifnull(iddepartamento,999), ifnull(idlocalidad,999), ifnull(tecnico1,3197), ifnull(tecnico2,3197), ifnull(idagencia,8), contrato, socio, nousar, caravanas, prolesa, ifnull(prolesasuc,0), prolesamat, ifnull(observaciones,''), ifnull(fac_rsocial,''), ifnull(fac_cedula,''), ifnull(fac_rut,''), ifnull(fac_direccion,''), ifnull(fac_localidad,''), ifnull(fac_departamento,0), ifnull(fac_cpostal,''), ifnull(fac_giro,0), ifnull(cob_nombre_telefono1,''), ifnull(fac_telefonos,''), ifnull(cob_nombre_telefono2,''), ifnull(cob_telefono2,''), ifnull(cob_nombre_celular1,''), ifnull(cob_celular1,''), ifnull(cob_nombre_celular2,''), ifnull(cob_celular2,''), ifnull(cob_nombre_email1,''), ifnull(cob_email1,''), ifnull(cob_nombre_email2,''), ifnull(cob_email2,''), ifnull(fac_fax,''), ifnull(fac_email,''), ifnull(fac_contacto,''), ifnull(fac_observaciones,''), fac_lista, fac_contado, ifnull(not_email_frascos1,''), ifnull(not_email_frascos2,''), ifnull(not_email_muestras1,''), ifnull(not_email_muestras2,''), ifnull(not_email_analisis1,''), ifnull(not_email_analisis2,''), ifnull(not_email_general1,''), ifnull(not_email_general2,'') FROM cliente WHERE usuario_web = '" & obj.USUARIO_WEB & "'")

            If Ds.Tables(0).Rows.Count > 0 Then
                Dim unaFila As DataRow
                unaFila = Ds.Tables(0).Rows(0)
                p.ID = CType(unaFila.Item(0), Long)
                p.NOMBRE = CType(unaFila.Item(1), String)
                p.EMAIL = CType(unaFila.Item(2), String)
                p.NOMBRE_EMAIL1 = CType(unaFila.Item(3), String)
                p.EMAIL1 = CType(unaFila.Item(4), String)
                p.NOMBRE_EMAIL2 = CType(unaFila.Item(5), String)
                p.EMAIL2 = CType(unaFila.Item(6), String)
                p.ENVIO = CType(unaFila.Item(7), String)
                p.USUARIO_WEB = CType(unaFila.Item(8), String)
                p.NOMBRE_CELULAR1 = CType(unaFila.Item(9), String)
                p.CELULAR = CType(unaFila.Item(10), String)
                p.NOMBRE_CELULAR2 = CType(unaFila.Item(11), String)
                p.CELULAR2 = CType(unaFila.Item(12), String)
                p.CODIGOFIGARO = CType(unaFila.Item(13), String)
                p.TIPOUSUARIO = CType(unaFila.Item(14), Integer)
                p.DIRECCION = CType(unaFila.Item(15), String)
                p.NOMBRE_TELEFONO1 = CType(unaFila.Item(16), String)
                p.TELEFONO1 = CType(unaFila.Item(17), String)
                p.NOMBRE_TELEFONO2 = CType(unaFila.Item(18), String)
                p.TELEFONO2 = CType(unaFila.Item(19), String)
                p.FAX = CType(unaFila.Item(20), String)
                p.DICOSE = CType(unaFila.Item(21), String)
                p.IDDEPARTAMENTO = CType(unaFila.Item(22), Integer)
                p.IDLOCALIDAD = CType(unaFila.Item(23), Integer)
                p.TECNICO1 = CType(unaFila.Item(24), Long)
                p.TECNICO2 = CType(unaFila.Item(25), Long)
                p.IDAGENCIA = CType(unaFila.Item(26), Integer)
                p.CONTRATO = CType(unaFila.Item(27), Integer)
                p.SOCIO = CType(unaFila.Item(28), Integer)
                p.NOUSAR = CType(unaFila.Item(29), Integer)
                p.CARAVANAS = CType(unaFila.Item(30), Integer)
                p.PROLESA = CType(unaFila.Item(31), Integer)
                p.PROLESASUC = CType(unaFila.Item(32), Integer)
                p.PROLESAMAT = CType(unaFila.Item(33), Long)
                p.OBSERVACIONES = CType(unaFila.Item(34), String)
                p.FAC_RSOCIAL = CType(unaFila.Item(35), String)
                p.FAC_CEDULA = CType(unaFila.Item(36), String)
                p.FAC_RUT = CType(unaFila.Item(37), String)
                p.FAC_DIRECCION = CType(unaFila.Item(38), String)
                p.FAC_LOCALIDAD = CType(unaFila.Item(39), String)
                p.FAC_DEPARTAMENTO = CType(unaFila.Item(40), Integer)
                p.FAC_CPOSTAL = CType(unaFila.Item(41), String)
                p.FAC_GIRO = CType(unaFila.Item(42), Integer)
                p.COB_NOMBRE_TELEFONO1 = CType(unaFila.Item(43), String)
                p.FAC_TELEFONOS = CType(unaFila.Item(44), String)
                p.COB_NOMBRE_TELEFONO2 = CType(unaFila.Item(45), String)
                p.COB_TELEFONO2 = CType(unaFila.Item(46), String)
                p.COB_NOMBRE_CELULAR1 = CType(unaFila.Item(47), String)
                p.COB_CELULAR1 = CType(unaFila.Item(48), String)
                p.COB_NOMBRE_CELULAR2 = CType(unaFila.Item(49), String)
                p.COB_CELULAR2 = CType(unaFila.Item(50), String)
                p.COB_NOMBRE_EMAIL1 = CType(unaFila.Item(51), String)
                p.COB_EMAIL1 = CType(unaFila.Item(52), String)
                p.COB_NOMBRE_EMAIL2 = CType(unaFila.Item(53), String)
                p.COB_EMAIL2 = CType(unaFila.Item(54), String)
                p.FAC_FAX = CType(unaFila.Item(55), String)
                p.FAC_EMAIL = CType(unaFila.Item(56), String)
                p.FAC_CONTACTO = CType(unaFila.Item(57), String)
                p.FAC_OBSERVACIONES = CType(unaFila.Item(58), String)
                p.FAC_LISTA = CType(unaFila.Item(59), Integer)
                p.FAC_CONTADO = CType(unaFila.Item(60), Integer)
                p.NOT_EMAIL_FRASCOS1 = CType(unaFila.Item(61), String)
                p.NOT_EMAIL_FRASCOS2 = CType(unaFila.Item(62), String)
                p.NOT_EMAIL_MUESTRAS1 = CType(unaFila.Item(63), String)
                p.NOT_EMAIL_MUESTRAS2 = CType(unaFila.Item(64), String)
                p.NOT_EMAIL_ANALISIS1 = CType(unaFila.Item(65), String)
                p.NOT_EMAIL_ANALISIS2 = CType(unaFila.Item(66), String)
                p.NOT_EMAIL_GENERAL1 = CType(unaFila.Item(67), String)
                p.NOT_EMAIL_GENERAL2 = CType(unaFila.Item(68), String)
                Return p
            End If
            Return Nothing
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listar() As ArrayList
        Dim sql As String = "SELECT id, nombre, ifnull(email,''), ifnull(nombre_email1,''), ifnull(email1,''), ifnull(nombre_email2,''), ifnull(email2,''), ifnull(envio,''), ifnull(usuario_web,''), ifnull(nombre_celular1,''), ifnull(celular,''), ifnull(nombre_celular2,''), ifnull(celular2,''), ifnull(codigofigaro,''), ifnull(tipousuario,1), ifnull(direccion,''), ifnull(nombre_telefono1,''), ifnull(telefono1,''), ifnull(nombre_telefono2,''), ifnull(telefono2,''), ifnull(fax,''), ifnull(dicose,''), ifnull(iddepartamento,999), ifnull(idlocalidad,999), ifnull(tecnico1,3197), ifnull(tecnico2,3197), ifnull(idagencia,8), contrato, socio, nousar, caravanas, prolesa, ifnull(prolesasuc,0), prolesamat, ifnull(observaciones,''), ifnull(fac_rsocial,''), ifnull(fac_cedula,''), ifnull(fac_rut,''), ifnull(fac_direccion,''), ifnull(fac_localidad,''), ifnull(fac_departamento,0), ifnull(fac_cpostal,''), ifnull(fac_giro,0), ifnull(cob_nombre_telefono1,''), ifnull(fac_telefonos,''), ifnull(cob_nombre_telefono2,''), ifnull(cob_telefono2,''), ifnull(cob_nombre_celular1,''), ifnull(cob_celular1,''), ifnull(cob_nombre_celular2,''), ifnull(cob_celular2,''), ifnull(cob_nombre_email1,''), ifnull(cob_email1,''), ifnull(cob_nombre_email2,''), ifnull(cob_email2,''), ifnull(fac_fax,''), ifnull(fac_email,''), ifnull(fac_contacto,''), ifnull(fac_observaciones,''), fac_lista, fac_contado, ifnull(not_email_frascos1,''), ifnull(not_email_frascos2,''), ifnull(not_email_muestras1,''), ifnull(not_email_muestras2,''), ifnull(not_email_analisis1,''), ifnull(not_email_analisis2,''), ifnull(not_email_general1,''), ifnull(not_email_general2,'') FROM cliente WHERE nousar= 0 order by Nombre asc"
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim p As New dcliente
                    p.ID = CType(unaFila.Item(0), Long)
                    p.NOMBRE = CType(unaFila.Item(1), String)
                    p.EMAIL = CType(unaFila.Item(2), String)
                    p.NOMBRE_EMAIL1 = CType(unaFila.Item(3), String)
                    p.EMAIL1 = CType(unaFila.Item(4), String)
                    p.NOMBRE_EMAIL2 = CType(unaFila.Item(5), String)
                    p.EMAIL2 = CType(unaFila.Item(6), String)
                    p.ENVIO = CType(unaFila.Item(7), String)
                    p.USUARIO_WEB = CType(unaFila.Item(8), String)
                    p.NOMBRE_CELULAR1 = CType(unaFila.Item(9), String)
                    p.CELULAR = CType(unaFila.Item(10), String)
                    p.NOMBRE_CELULAR2 = CType(unaFila.Item(11), String)
                    p.CELULAR2 = CType(unaFila.Item(12), String)
                    p.CODIGOFIGARO = CType(unaFila.Item(13), String)
                    p.TIPOUSUARIO = CType(unaFila.Item(14), Integer)
                    p.DIRECCION = CType(unaFila.Item(15), String)
                    p.NOMBRE_TELEFONO1 = CType(unaFila.Item(16), String)
                    p.TELEFONO1 = CType(unaFila.Item(17), String)
                    p.NOMBRE_TELEFONO2 = CType(unaFila.Item(18), String)
                    p.TELEFONO2 = CType(unaFila.Item(19), String)
                    p.FAX = CType(unaFila.Item(20), String)
                    p.DICOSE = CType(unaFila.Item(21), String)
                    p.IDDEPARTAMENTO = CType(unaFila.Item(22), Integer)
                    p.IDLOCALIDAD = CType(unaFila.Item(23), Integer)
                    p.TECNICO1 = CType(unaFila.Item(24), Long)
                    p.TECNICO2 = CType(unaFila.Item(25), Long)
                    p.IDAGENCIA = CType(unaFila.Item(26), Integer)
                    p.CONTRATO = CType(unaFila.Item(27), Integer)
                    p.SOCIO = CType(unaFila.Item(28), Integer)
                    p.NOUSAR = CType(unaFila.Item(29), Integer)
                    p.CARAVANAS = CType(unaFila.Item(30), Integer)
                    p.PROLESA = CType(unaFila.Item(31), Integer)
                    p.PROLESASUC = CType(unaFila.Item(32), Integer)
                    p.PROLESAMAT = CType(unaFila.Item(33), Long)
                    p.OBSERVACIONES = CType(unaFila.Item(34), String)
                    p.FAC_RSOCIAL = CType(unaFila.Item(35), String)
                    p.FAC_CEDULA = CType(unaFila.Item(36), String)
                    p.FAC_RUT = CType(unaFila.Item(37), String)
                    p.FAC_DIRECCION = CType(unaFila.Item(38), String)
                    p.FAC_LOCALIDAD = CType(unaFila.Item(39), String)
                    p.FAC_DEPARTAMENTO = CType(unaFila.Item(40), Integer)
                    p.FAC_CPOSTAL = CType(unaFila.Item(41), String)
                    p.FAC_GIRO = CType(unaFila.Item(42), Integer)
                    p.COB_NOMBRE_TELEFONO1 = CType(unaFila.Item(43), String)
                    p.FAC_TELEFONOS = CType(unaFila.Item(44), String)
                    p.COB_NOMBRE_TELEFONO2 = CType(unaFila.Item(45), String)
                    p.COB_TELEFONO2 = CType(unaFila.Item(46), String)
                    p.COB_NOMBRE_CELULAR1 = CType(unaFila.Item(47), String)
                    p.COB_CELULAR1 = CType(unaFila.Item(48), String)
                    p.COB_NOMBRE_CELULAR2 = CType(unaFila.Item(49), String)
                    p.COB_CELULAR2 = CType(unaFila.Item(50), String)
                    p.COB_NOMBRE_EMAIL1 = CType(unaFila.Item(51), String)
                    p.COB_EMAIL1 = CType(unaFila.Item(52), String)
                    p.COB_NOMBRE_EMAIL2 = CType(unaFila.Item(53), String)
                    p.COB_EMAIL2 = CType(unaFila.Item(54), String)
                    p.FAC_FAX = CType(unaFila.Item(55), String)
                    p.FAC_EMAIL = CType(unaFila.Item(56), String)
                    p.FAC_CONTACTO = CType(unaFila.Item(57), String)
                    p.FAC_OBSERVACIONES = CType(unaFila.Item(58), String)
                    p.FAC_LISTA = CType(unaFila.Item(59), Integer)
                    p.FAC_CONTADO = CType(unaFila.Item(60), Integer)
                    p.NOT_EMAIL_FRASCOS1 = CType(unaFila.Item(61), String)
                    p.NOT_EMAIL_FRASCOS2 = CType(unaFila.Item(62), String)
                    p.NOT_EMAIL_MUESTRAS1 = CType(unaFila.Item(63), String)
                    p.NOT_EMAIL_MUESTRAS2 = CType(unaFila.Item(64), String)
                    p.NOT_EMAIL_ANALISIS1 = CType(unaFila.Item(65), String)
                    p.NOT_EMAIL_ANALISIS2 = CType(unaFila.Item(66), String)
                    p.NOT_EMAIL_GENERAL1 = CType(unaFila.Item(67), String)
                    p.NOT_EMAIL_GENERAL2 = CType(unaFila.Item(68), String)
                    Lista.Add(p)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listarxtecnico(ByVal id As Long) As ArrayList
        Dim sql As String = "SELECT id, nombre, ifnull(email,''), ifnull(nombre_email1,''), ifnull(email1,''), ifnull(nombre_email2,''), ifnull(email2,''), ifnull(envio,''), ifnull(usuario_web,''), ifnull(nombre_celular1,''), ifnull(celular,''), ifnull(nombre_celular2,''), ifnull(celular2,''), ifnull(codigofigaro,''), ifnull(tipousuario,1), ifnull(direccion,''), ifnull(nombre_telefono1,''), ifnull(telefono1,''), ifnull(nombre_telefono2,''), ifnull(telefono2,''), ifnull(fax,''), ifnull(dicose,''), ifnull(iddepartamento,999), ifnull(idlocalidad,999), ifnull(tecnico1,3197), ifnull(tecnico2,3197), ifnull(idagencia,8), contrato, socio, nousar, caravanas, prolesa, ifnull(prolesasuc,0), prolesamat, ifnull(observaciones,''), ifnull(fac_rsocial,''), ifnull(fac_cedula,''), ifnull(fac_rut,''), ifnull(fac_direccion,''), ifnull(fac_localidad,''), ifnull(fac_departamento,0), ifnull(fac_cpostal,''), ifnull(fac_giro,0), ifnull(cob_nombre_telefono1,''), ifnull(fac_telefonos,''), ifnull(cob_nombre_telefono2,''), ifnull(cob_telefono2,''), ifnull(cob_nombre_celular1,''), ifnull(cob_celular1,''), ifnull(cob_nombre_celular2,''), ifnull(cob_celular2,''), ifnull(cob_nombre_email1,''), ifnull(cob_email1,''), ifnull(cob_nombre_email2,''), ifnull(cob_email2,''), ifnull(fac_fax,''), ifnull(fac_email,''), ifnull(fac_contacto,''), ifnull(fac_observaciones,''), fac_lista, fac_contado, ifnull(not_email_frascos1,''), ifnull(not_email_frascos2,''), ifnull(not_email_muestras1,''), ifnull(not_email_muestras2,''), ifnull(not_email_analisis1,''), ifnull(not_email_analisis2,''), ifnull(not_email_general1,''), ifnull(not_email_general2,'') FROM cliente WHERE tecnico= " & id & " or tecnico2= " & id & " or tecnico3= " & id & " and nousar= 0 order by Nombre asc"
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim p As New dcliente
                    p.ID = CType(unaFila.Item(0), Long)
                    p.NOMBRE = CType(unaFila.Item(1), String)
                    p.EMAIL = CType(unaFila.Item(2), String)
                    p.NOMBRE_EMAIL1 = CType(unaFila.Item(3), String)
                    p.EMAIL1 = CType(unaFila.Item(4), String)
                    p.NOMBRE_EMAIL2 = CType(unaFila.Item(5), String)
                    p.EMAIL2 = CType(unaFila.Item(6), String)
                    p.ENVIO = CType(unaFila.Item(7), String)
                    p.USUARIO_WEB = CType(unaFila.Item(8), String)
                    p.NOMBRE_CELULAR1 = CType(unaFila.Item(9), String)
                    p.CELULAR = CType(unaFila.Item(10), String)
                    p.NOMBRE_CELULAR2 = CType(unaFila.Item(11), String)
                    p.CELULAR2 = CType(unaFila.Item(12), String)
                    p.CODIGOFIGARO = CType(unaFila.Item(13), String)
                    p.TIPOUSUARIO = CType(unaFila.Item(14), Integer)
                    p.DIRECCION = CType(unaFila.Item(15), String)
                    p.NOMBRE_TELEFONO1 = CType(unaFila.Item(16), String)
                    p.TELEFONO1 = CType(unaFila.Item(17), String)
                    p.NOMBRE_TELEFONO2 = CType(unaFila.Item(18), String)
                    p.TELEFONO2 = CType(unaFila.Item(19), String)
                    p.FAX = CType(unaFila.Item(20), String)
                    p.DICOSE = CType(unaFila.Item(21), String)
                    p.IDDEPARTAMENTO = CType(unaFila.Item(22), Integer)
                    p.IDLOCALIDAD = CType(unaFila.Item(23), Integer)
                    p.TECNICO1 = CType(unaFila.Item(24), Long)
                    p.TECNICO2 = CType(unaFila.Item(25), Long)
                    p.IDAGENCIA = CType(unaFila.Item(26), Integer)
                    p.CONTRATO = CType(unaFila.Item(27), Integer)
                    p.SOCIO = CType(unaFila.Item(28), Integer)
                    p.NOUSAR = CType(unaFila.Item(29), Integer)
                    p.CARAVANAS = CType(unaFila.Item(30), Integer)
                    p.PROLESA = CType(unaFila.Item(31), Integer)
                    p.PROLESASUC = CType(unaFila.Item(32), Integer)
                    p.PROLESAMAT = CType(unaFila.Item(33), Long)
                    p.OBSERVACIONES = CType(unaFila.Item(34), String)
                    p.FAC_RSOCIAL = CType(unaFila.Item(35), String)
                    p.FAC_CEDULA = CType(unaFila.Item(36), String)
                    p.FAC_RUT = CType(unaFila.Item(37), String)
                    p.FAC_DIRECCION = CType(unaFila.Item(38), String)
                    p.FAC_LOCALIDAD = CType(unaFila.Item(39), String)
                    p.FAC_DEPARTAMENTO = CType(unaFila.Item(40), Integer)
                    p.FAC_CPOSTAL = CType(unaFila.Item(41), String)
                    p.FAC_GIRO = CType(unaFila.Item(42), Integer)
                    p.COB_NOMBRE_TELEFONO1 = CType(unaFila.Item(43), String)
                    p.FAC_TELEFONOS = CType(unaFila.Item(44), String)
                    p.COB_NOMBRE_TELEFONO2 = CType(unaFila.Item(45), String)
                    p.COB_TELEFONO2 = CType(unaFila.Item(46), String)
                    p.COB_NOMBRE_CELULAR1 = CType(unaFila.Item(47), String)
                    p.COB_CELULAR1 = CType(unaFila.Item(48), String)
                    p.COB_NOMBRE_CELULAR2 = CType(unaFila.Item(49), String)
                    p.COB_CELULAR2 = CType(unaFila.Item(50), String)
                    p.COB_NOMBRE_EMAIL1 = CType(unaFila.Item(51), String)
                    p.COB_EMAIL1 = CType(unaFila.Item(52), String)
                    p.COB_NOMBRE_EMAIL2 = CType(unaFila.Item(53), String)
                    p.COB_EMAIL2 = CType(unaFila.Item(54), String)
                    p.FAC_FAX = CType(unaFila.Item(55), String)
                    p.FAC_EMAIL = CType(unaFila.Item(56), String)
                    p.FAC_CONTACTO = CType(unaFila.Item(57), String)
                    p.FAC_OBSERVACIONES = CType(unaFila.Item(58), String)
                    p.FAC_LISTA = CType(unaFila.Item(59), Integer)
                    p.FAC_CONTADO = CType(unaFila.Item(60), Integer)
                    p.NOT_EMAIL_FRASCOS1 = CType(unaFila.Item(61), String)
                    p.NOT_EMAIL_FRASCOS2 = CType(unaFila.Item(62), String)
                    p.NOT_EMAIL_MUESTRAS1 = CType(unaFila.Item(63), String)
                    p.NOT_EMAIL_MUESTRAS2 = CType(unaFila.Item(64), String)
                    p.NOT_EMAIL_ANALISIS1 = CType(unaFila.Item(65), String)
                    p.NOT_EMAIL_ANALISIS2 = CType(unaFila.Item(66), String)
                    p.NOT_EMAIL_GENERAL1 = CType(unaFila.Item(67), String)
                    p.NOT_EMAIL_GENERAL2 = CType(unaFila.Item(68), String)
                    Lista.Add(p)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listarxdepartamento(ByVal iddepto As Integer) As ArrayList
        Dim sql As String = "SELECT id, nombre, ifnull(email,''), ifnull(nombre_email1,''), ifnull(email1,''), ifnull(nombre_email2,''), ifnull(email2,''), ifnull(envio,''), ifnull(usuario_web,''), ifnull(nombre_celular1,''), ifnull(celular,''), ifnull(nombre_celular2,''), ifnull(celular2,''), ifnull(codigofigaro,''), ifnull(tipousuario,1), ifnull(direccion,''), ifnull(nombre_telefono1,''), ifnull(telefono1,''), ifnull(nombre_telefono2,''), ifnull(telefono2,''), ifnull(fax,''), ifnull(dicose,''), ifnull(iddepartamento,999), ifnull(idlocalidad,999), ifnull(tecnico1,3197), ifnull(tecnico2,3197), ifnull(idagencia,8), contrato, socio, nousar, caravanas, prolesa, ifnull(prolesasuc,0), prolesamat, ifnull(observaciones,''), ifnull(fac_rsocial,''), ifnull(fac_cedula,''), ifnull(fac_rut,''), ifnull(fac_direccion,''), ifnull(fac_localidad,''), ifnull(fac_departamento,0), ifnull(fac_cpostal,''), ifnull(fac_giro,0), ifnull(cob_nombre_telefono1,''), ifnull(fac_telefonos,''), ifnull(cob_nombre_telefono2,''), ifnull(cob_telefono2,''), ifnull(cob_nombre_celular1,''), ifnull(cob_celular1,''), ifnull(cob_nombre_celular2,''), ifnull(cob_celular2,''), ifnull(cob_nombre_email1,''), ifnull(cob_email1,''), ifnull(cob_nombre_email2,''), ifnull(cob_email2,''), ifnull(fac_fax,''), ifnull(fac_email,''), ifnull(fac_contacto,''), ifnull(fac_observaciones,''), fac_lista, fac_contado, ifnull(not_email_frascos1,''), ifnull(not_email_frascos2,''), ifnull(not_email_muestras1,''), ifnull(not_email_muestras2,''), ifnull(not_email_analisis1,''), ifnull(not_email_analisis2,''), ifnull(not_email_general1,''), ifnull(not_email_general2,'') FROM cliente WHERE iddepartamento= " & iddepto & "  and nousar= 0 order by Nombre asc"
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim p As New dcliente
                    p.ID = CType(unaFila.Item(0), Long)
                    p.NOMBRE = CType(unaFila.Item(1), String)
                    p.EMAIL = CType(unaFila.Item(2), String)
                    p.NOMBRE_EMAIL1 = CType(unaFila.Item(3), String)
                    p.EMAIL1 = CType(unaFila.Item(4), String)
                    p.NOMBRE_EMAIL2 = CType(unaFila.Item(5), String)
                    p.EMAIL2 = CType(unaFila.Item(6), String)
                    p.ENVIO = CType(unaFila.Item(7), String)
                    p.USUARIO_WEB = CType(unaFila.Item(8), String)
                    p.NOMBRE_CELULAR1 = CType(unaFila.Item(9), String)
                    p.CELULAR = CType(unaFila.Item(10), String)
                    p.NOMBRE_CELULAR2 = CType(unaFila.Item(11), String)
                    p.CELULAR2 = CType(unaFila.Item(12), String)
                    p.CODIGOFIGARO = CType(unaFila.Item(13), String)
                    p.TIPOUSUARIO = CType(unaFila.Item(14), Integer)
                    p.DIRECCION = CType(unaFila.Item(15), String)
                    p.NOMBRE_TELEFONO1 = CType(unaFila.Item(16), String)
                    p.TELEFONO1 = CType(unaFila.Item(17), String)
                    p.NOMBRE_TELEFONO2 = CType(unaFila.Item(18), String)
                    p.TELEFONO2 = CType(unaFila.Item(19), String)
                    p.FAX = CType(unaFila.Item(20), String)
                    p.DICOSE = CType(unaFila.Item(21), String)
                    p.IDDEPARTAMENTO = CType(unaFila.Item(22), Integer)
                    p.IDLOCALIDAD = CType(unaFila.Item(23), Integer)
                    p.TECNICO1 = CType(unaFila.Item(24), Long)
                    p.TECNICO2 = CType(unaFila.Item(25), Long)
                    p.IDAGENCIA = CType(unaFila.Item(26), Integer)
                    p.CONTRATO = CType(unaFila.Item(27), Integer)
                    p.SOCIO = CType(unaFila.Item(28), Integer)
                    p.NOUSAR = CType(unaFila.Item(29), Integer)
                    p.CARAVANAS = CType(unaFila.Item(30), Integer)
                    p.PROLESA = CType(unaFila.Item(31), Integer)
                    p.PROLESASUC = CType(unaFila.Item(32), Integer)
                    p.PROLESAMAT = CType(unaFila.Item(33), Long)
                    p.OBSERVACIONES = CType(unaFila.Item(34), String)
                    p.FAC_RSOCIAL = CType(unaFila.Item(35), String)
                    p.FAC_CEDULA = CType(unaFila.Item(36), String)
                    p.FAC_RUT = CType(unaFila.Item(37), String)
                    p.FAC_DIRECCION = CType(unaFila.Item(38), String)
                    p.FAC_LOCALIDAD = CType(unaFila.Item(39), String)
                    p.FAC_DEPARTAMENTO = CType(unaFila.Item(40), Integer)
                    p.FAC_CPOSTAL = CType(unaFila.Item(41), String)
                    p.FAC_GIRO = CType(unaFila.Item(42), Integer)
                    p.COB_NOMBRE_TELEFONO1 = CType(unaFila.Item(43), String)
                    p.FAC_TELEFONOS = CType(unaFila.Item(44), String)
                    p.COB_NOMBRE_TELEFONO2 = CType(unaFila.Item(45), String)
                    p.COB_TELEFONO2 = CType(unaFila.Item(46), String)
                    p.COB_NOMBRE_CELULAR1 = CType(unaFila.Item(47), String)
                    p.COB_CELULAR1 = CType(unaFila.Item(48), String)
                    p.COB_NOMBRE_CELULAR2 = CType(unaFila.Item(49), String)
                    p.COB_CELULAR2 = CType(unaFila.Item(50), String)
                    p.COB_NOMBRE_EMAIL1 = CType(unaFila.Item(51), String)
                    p.COB_EMAIL1 = CType(unaFila.Item(52), String)
                    p.COB_NOMBRE_EMAIL2 = CType(unaFila.Item(53), String)
                    p.COB_EMAIL2 = CType(unaFila.Item(54), String)
                    p.FAC_FAX = CType(unaFila.Item(55), String)
                    p.FAC_EMAIL = CType(unaFila.Item(56), String)
                    p.FAC_CONTACTO = CType(unaFila.Item(57), String)
                    p.FAC_OBSERVACIONES = CType(unaFila.Item(58), String)
                    p.FAC_LISTA = CType(unaFila.Item(59), Integer)
                    p.FAC_CONTADO = CType(unaFila.Item(60), Integer)
                    p.NOT_EMAIL_FRASCOS1 = CType(unaFila.Item(61), String)
                    p.NOT_EMAIL_FRASCOS2 = CType(unaFila.Item(62), String)
                    p.NOT_EMAIL_MUESTRAS1 = CType(unaFila.Item(63), String)
                    p.NOT_EMAIL_MUESTRAS2 = CType(unaFila.Item(64), String)
                    p.NOT_EMAIL_ANALISIS1 = CType(unaFila.Item(65), String)
                    p.NOT_EMAIL_ANALISIS2 = CType(unaFila.Item(66), String)
                    p.NOT_EMAIL_GENERAL1 = CType(unaFila.Item(67), String)
                    p.NOT_EMAIL_GENERAL2 = CType(unaFila.Item(68), String)
                    Lista.Add(p)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listartodos() As ArrayList
        Dim sql As String = "SELECT id, nombre, ifnull(email,''), ifnull(nombre_email1,''), ifnull(email1,''), ifnull(nombre_email2,''), ifnull(email2,''), ifnull(envio,''), ifnull(usuario_web,''), ifnull(nombre_celular1,''), ifnull(celular,''), ifnull(nombre_celular2,''), ifnull(celular2,''), ifnull(codigofigaro,''), ifnull(tipousuario,1), ifnull(direccion,''), ifnull(nombre_telefono1,''), ifnull(telefono1,''), ifnull(nombre_telefono2,''), ifnull(telefono2,''), ifnull(fax,''), ifnull(dicose,''), ifnull(iddepartamento,999), ifnull(idlocalidad,999), ifnull(tecnico1,3197), ifnull(tecnico2,3197), ifnull(idagencia,8), contrato, socio, nousar, caravanas, prolesa, ifnull(prolesasuc,0), prolesamat, ifnull(observaciones,''), ifnull(fac_rsocial,''), ifnull(fac_cedula,''), ifnull(fac_rut,''), ifnull(fac_direccion,''), ifnull(fac_localidad,''), ifnull(fac_departamento,0), ifnull(fac_cpostal,''), ifnull(fac_giro,0), ifnull(cob_nombre_telefono1,''), ifnull(fac_telefonos,''), ifnull(cob_nombre_telefono2,''), ifnull(cob_telefono2,''), ifnull(cob_nombre_celular1,''), ifnull(cob_celular1,''), ifnull(cob_nombre_celular2,''), ifnull(cob_celular2,''), ifnull(cob_nombre_email1,''), ifnull(cob_email1,''), ifnull(cob_nombre_email2,''), ifnull(cob_email2,''), ifnull(fac_fax,''), ifnull(fac_email,''), ifnull(fac_contacto,''), ifnull(fac_observaciones,''), fac_lista, fac_contado, ifnull(not_email_frascos1,''), ifnull(not_email_frascos2,''), ifnull(not_email_muestras1,''), ifnull(not_email_muestras2,''), ifnull(not_email_analisis1,''), ifnull(not_email_analisis2,''), ifnull(not_email_general1,''), ifnull(not_email_general2,'') FROM cliente order by Nombre asc"
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim p As New dcliente
                    p.ID = CType(unaFila.Item(0), Long)
                    p.NOMBRE = CType(unaFila.Item(1), String)
                    p.EMAIL = CType(unaFila.Item(2), String)
                    p.NOMBRE_EMAIL1 = CType(unaFila.Item(3), String)
                    p.EMAIL1 = CType(unaFila.Item(4), String)
                    p.NOMBRE_EMAIL2 = CType(unaFila.Item(5), String)
                    p.EMAIL2 = CType(unaFila.Item(6), String)
                    p.ENVIO = CType(unaFila.Item(7), String)
                    p.USUARIO_WEB = CType(unaFila.Item(8), String)
                    p.NOMBRE_CELULAR1 = CType(unaFila.Item(9), String)
                    p.CELULAR = CType(unaFila.Item(10), String)
                    p.NOMBRE_CELULAR2 = CType(unaFila.Item(11), String)
                    p.CELULAR2 = CType(unaFila.Item(12), String)
                    p.CODIGOFIGARO = CType(unaFila.Item(13), String)
                    p.TIPOUSUARIO = CType(unaFila.Item(14), Integer)
                    p.DIRECCION = CType(unaFila.Item(15), String)
                    p.NOMBRE_TELEFONO1 = CType(unaFila.Item(16), String)
                    p.TELEFONO1 = CType(unaFila.Item(17), String)
                    p.NOMBRE_TELEFONO2 = CType(unaFila.Item(18), String)
                    p.TELEFONO2 = CType(unaFila.Item(19), String)
                    p.FAX = CType(unaFila.Item(20), String)
                    p.DICOSE = CType(unaFila.Item(21), String)
                    p.IDDEPARTAMENTO = CType(unaFila.Item(22), Integer)
                    p.IDLOCALIDAD = CType(unaFila.Item(23), Integer)
                    p.TECNICO1 = CType(unaFila.Item(24), Long)
                    p.TECNICO2 = CType(unaFila.Item(25), Long)
                    p.IDAGENCIA = CType(unaFila.Item(26), Integer)
                    p.CONTRATO = CType(unaFila.Item(27), Integer)
                    p.SOCIO = CType(unaFila.Item(28), Integer)
                    p.NOUSAR = CType(unaFila.Item(29), Integer)
                    p.CARAVANAS = CType(unaFila.Item(30), Integer)
                    p.PROLESA = CType(unaFila.Item(31), Integer)
                    p.PROLESASUC = CType(unaFila.Item(32), Integer)
                    p.PROLESAMAT = CType(unaFila.Item(33), Long)
                    p.OBSERVACIONES = CType(unaFila.Item(34), String)
                    p.FAC_RSOCIAL = CType(unaFila.Item(35), String)
                    p.FAC_CEDULA = CType(unaFila.Item(36), String)
                    p.FAC_RUT = CType(unaFila.Item(37), String)
                    p.FAC_DIRECCION = CType(unaFila.Item(38), String)
                    p.FAC_LOCALIDAD = CType(unaFila.Item(39), String)
                    p.FAC_DEPARTAMENTO = CType(unaFila.Item(40), Integer)
                    p.FAC_CPOSTAL = CType(unaFila.Item(41), String)
                    p.FAC_GIRO = CType(unaFila.Item(42), Integer)
                    p.COB_NOMBRE_TELEFONO1 = CType(unaFila.Item(43), String)
                    p.FAC_TELEFONOS = CType(unaFila.Item(44), String)
                    p.COB_NOMBRE_TELEFONO2 = CType(unaFila.Item(45), String)
                    p.COB_TELEFONO2 = CType(unaFila.Item(46), String)
                    p.COB_NOMBRE_CELULAR1 = CType(unaFila.Item(47), String)
                    p.COB_CELULAR1 = CType(unaFila.Item(48), String)
                    p.COB_NOMBRE_CELULAR2 = CType(unaFila.Item(49), String)
                    p.COB_CELULAR2 = CType(unaFila.Item(50), String)
                    p.COB_NOMBRE_EMAIL1 = CType(unaFila.Item(51), String)
                    p.COB_EMAIL1 = CType(unaFila.Item(52), String)
                    p.COB_NOMBRE_EMAIL2 = CType(unaFila.Item(53), String)
                    p.COB_EMAIL2 = CType(unaFila.Item(54), String)
                    p.FAC_FAX = CType(unaFila.Item(55), String)
                    p.FAC_EMAIL = CType(unaFila.Item(56), String)
                    p.FAC_CONTACTO = CType(unaFila.Item(57), String)
                    p.FAC_OBSERVACIONES = CType(unaFila.Item(58), String)
                    p.FAC_LISTA = CType(unaFila.Item(59), Integer)
                    p.FAC_CONTADO = CType(unaFila.Item(60), Integer)
                    p.NOT_EMAIL_FRASCOS1 = CType(unaFila.Item(61), String)
                    p.NOT_EMAIL_FRASCOS2 = CType(unaFila.Item(62), String)
                    p.NOT_EMAIL_MUESTRAS1 = CType(unaFila.Item(63), String)
                    p.NOT_EMAIL_MUESTRAS2 = CType(unaFila.Item(64), String)
                    p.NOT_EMAIL_ANALISIS1 = CType(unaFila.Item(65), String)
                    p.NOT_EMAIL_ANALISIS2 = CType(unaFila.Item(66), String)
                    p.NOT_EMAIL_GENERAL1 = CType(unaFila.Item(67), String)
                    p.NOT_EMAIL_GENERAL2 = CType(unaFila.Item(68), String)
                    Lista.Add(p)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listarsocios() As ArrayList
        Dim sql As String = "SELECT id, nombre, ifnull(email,''), ifnull(nombre_email1,''), ifnull(email1,''), ifnull(nombre_email2,''), ifnull(email2,''), ifnull(envio,''), ifnull(usuario_web,''), ifnull(nombre_celular1,''), ifnull(celular,''), ifnull(nombre_celular2,''), ifnull(celular2,''), ifnull(codigofigaro,''), ifnull(tipousuario,1), ifnull(direccion,''), ifnull(nombre_telefono1,''), ifnull(telefono1,''), ifnull(nombre_telefono2,''), ifnull(telefono2,''), ifnull(fax,''), ifnull(dicose,''), ifnull(iddepartamento,999), ifnull(idlocalidad,999), ifnull(tecnico1,3197), ifnull(tecnico2,3197), ifnull(idagencia,8), contrato, socio, nousar, caravanas, prolesa, ifnull(prolesasuc,0), prolesamat, ifnull(observaciones,''), ifnull(fac_rsocial,''), ifnull(fac_cedula,''), ifnull(fac_rut,''), ifnull(fac_direccion,''), ifnull(fac_localidad,''), ifnull(fac_departamento,0), ifnull(fac_cpostal,''), ifnull(fac_giro,0), ifnull(cob_nombre_telefono1,''), ifnull(fac_telefonos,''), ifnull(cob_nombre_telefono2,''), ifnull(cob_telefono2,''), ifnull(cob_nombre_celular1,''), ifnull(cob_celular1,''), ifnull(cob_nombre_celular2,''), ifnull(cob_celular2,''), ifnull(cob_nombre_email1,''), ifnull(cob_email1,''), ifnull(cob_nombre_email2,''), ifnull(cob_email2,''), ifnull(fac_fax,''), ifnull(fac_email,''), ifnull(fac_contacto,''), ifnull(fac_observaciones,''), fac_lista, fac_contado, ifnull(not_email_frascos1,''), ifnull(not_email_frascos2,''), ifnull(not_email_muestras1,''), ifnull(not_email_muestras2,''), ifnull(not_email_analisis1,''), ifnull(not_email_analisis2,''), ifnull(not_email_general1,''), ifnull(not_email_general2,'') FROM cliente where socio = 1 order by Nombre asc"
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim p As New dCliente
                    p.ID = CType(unaFila.Item(0), Long)
                    p.NOMBRE = CType(unaFila.Item(1), String)
                    p.EMAIL = CType(unaFila.Item(2), String)
                    p.NOMBRE_EMAIL1 = CType(unaFila.Item(3), String)
                    p.EMAIL1 = CType(unaFila.Item(4), String)
                    p.NOMBRE_EMAIL2 = CType(unaFila.Item(5), String)
                    p.EMAIL2 = CType(unaFila.Item(6), String)
                    p.ENVIO = CType(unaFila.Item(7), String)
                    p.USUARIO_WEB = CType(unaFila.Item(8), String)
                    p.NOMBRE_CELULAR1 = CType(unaFila.Item(9), String)
                    p.CELULAR = CType(unaFila.Item(10), String)
                    p.NOMBRE_CELULAR2 = CType(unaFila.Item(11), String)
                    p.CELULAR2 = CType(unaFila.Item(12), String)
                    p.CODIGOFIGARO = CType(unaFila.Item(13), String)
                    p.TIPOUSUARIO = CType(unaFila.Item(14), Integer)
                    p.DIRECCION = CType(unaFila.Item(15), String)
                    p.NOMBRE_TELEFONO1 = CType(unaFila.Item(16), String)
                    p.TELEFONO1 = CType(unaFila.Item(17), String)
                    p.NOMBRE_TELEFONO2 = CType(unaFila.Item(18), String)
                    p.TELEFONO2 = CType(unaFila.Item(19), String)
                    p.FAX = CType(unaFila.Item(20), String)
                    p.DICOSE = CType(unaFila.Item(21), String)
                    p.IDDEPARTAMENTO = CType(unaFila.Item(22), Integer)
                    p.IDLOCALIDAD = CType(unaFila.Item(23), Integer)
                    p.TECNICO1 = CType(unaFila.Item(24), Long)
                    p.TECNICO2 = CType(unaFila.Item(25), Long)
                    p.IDAGENCIA = CType(unaFila.Item(26), Integer)
                    p.CONTRATO = CType(unaFila.Item(27), Integer)
                    p.SOCIO = CType(unaFila.Item(28), Integer)
                    p.NOUSAR = CType(unaFila.Item(29), Integer)
                    p.CARAVANAS = CType(unaFila.Item(30), Integer)
                    p.PROLESA = CType(unaFila.Item(31), Integer)
                    p.PROLESASUC = CType(unaFila.Item(32), Integer)
                    p.PROLESAMAT = CType(unaFila.Item(33), Long)
                    p.OBSERVACIONES = CType(unaFila.Item(34), String)
                    p.FAC_RSOCIAL = CType(unaFila.Item(35), String)
                    p.FAC_CEDULA = CType(unaFila.Item(36), String)
                    p.FAC_RUT = CType(unaFila.Item(37), String)
                    p.FAC_DIRECCION = CType(unaFila.Item(38), String)
                    p.FAC_LOCALIDAD = CType(unaFila.Item(39), String)
                    p.FAC_DEPARTAMENTO = CType(unaFila.Item(40), Integer)
                    p.FAC_CPOSTAL = CType(unaFila.Item(41), String)
                    p.FAC_GIRO = CType(unaFila.Item(42), Integer)
                    p.COB_NOMBRE_TELEFONO1 = CType(unaFila.Item(43), String)
                    p.FAC_TELEFONOS = CType(unaFila.Item(44), String)
                    p.COB_NOMBRE_TELEFONO2 = CType(unaFila.Item(45), String)
                    p.COB_TELEFONO2 = CType(unaFila.Item(46), String)
                    p.COB_NOMBRE_CELULAR1 = CType(unaFila.Item(47), String)
                    p.COB_CELULAR1 = CType(unaFila.Item(48), String)
                    p.COB_NOMBRE_CELULAR2 = CType(unaFila.Item(49), String)
                    p.COB_CELULAR2 = CType(unaFila.Item(50), String)
                    p.COB_NOMBRE_EMAIL1 = CType(unaFila.Item(51), String)
                    p.COB_EMAIL1 = CType(unaFila.Item(52), String)
                    p.COB_NOMBRE_EMAIL2 = CType(unaFila.Item(53), String)
                    p.COB_EMAIL2 = CType(unaFila.Item(54), String)
                    p.FAC_FAX = CType(unaFila.Item(55), String)
                    p.FAC_EMAIL = CType(unaFila.Item(56), String)
                    p.FAC_CONTACTO = CType(unaFila.Item(57), String)
                    p.FAC_OBSERVACIONES = CType(unaFila.Item(58), String)
                    p.FAC_LISTA = CType(unaFila.Item(59), Integer)
                    p.FAC_CONTADO = CType(unaFila.Item(60), Integer)
                    p.NOT_EMAIL_FRASCOS1 = CType(unaFila.Item(61), String)
                    p.NOT_EMAIL_FRASCOS2 = CType(unaFila.Item(62), String)
                    p.NOT_EMAIL_MUESTRAS1 = CType(unaFila.Item(63), String)
                    p.NOT_EMAIL_MUESTRAS2 = CType(unaFila.Item(64), String)
                    p.NOT_EMAIL_ANALISIS1 = CType(unaFila.Item(65), String)
                    p.NOT_EMAIL_ANALISIS2 = CType(unaFila.Item(66), String)
                    p.NOT_EMAIL_GENERAL1 = CType(unaFila.Item(67), String)
                    p.NOT_EMAIL_GENERAL2 = CType(unaFila.Item(68), String)
                    Lista.Add(p)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listarempresa() As ArrayList
        Dim sql As String = "SELECT SELECT id, nombre, ifnull(email,''), ifnull(nombre_email1,''), ifnull(email1,''), ifnull(nombre_email2,''), ifnull(email2,''), ifnull(envio,''), ifnull(usuario_web,''), ifnull(nombre_celular1,''), ifnull(celular,''), ifnull(nombre_celular2,''), ifnull(celular2,''), ifnull(codigofigaro,''), ifnull(tipousuario,1), ifnull(direccion,''), ifnull(nombre_telefono1,''), ifnull(telefono1,''), ifnull(nombre_telefono2,''), ifnull(telefono2,''), ifnull(fax,''), ifnull(dicose,''), ifnull(iddepartamento,999), ifnull(idlocalidad,999), ifnull(tecnico1,3197), ifnull(tecnico2,3197), ifnull(idagencia,8), contrato, socio, nousar, caravanas, prolesa, ifnull(prolesasuc,0), prolesamat, ifnull(observaciones,''), ifnull(fac_rsocial,''), ifnull(fac_cedula,''), ifnull(fac_rut,''), ifnull(fac_direccion,''), ifnull(fac_localidad,''), ifnull(fac_departamento,0), ifnull(fac_cpostal,''), ifnull(fac_giro,0), ifnull(cob_nombre_telefono1,''), ifnull(fac_telefonos,''), ifnull(cob_nombre_telefono2,''), ifnull(cob_telefono2,''), ifnull(cob_nombre_celular1,''), ifnull(cob_celular1,''), ifnull(cob_nombre_celular2,''), ifnull(cob_celular2,''), ifnull(cob_nombre_email1,''), ifnull(cob_email1,''), ifnull(cob_nombre_email2,''), ifnull(cob_email2,''), ifnull(fac_fax,''), ifnull(fac_email,''), ifnull(fac_contacto,''), ifnull(fac_observaciones,''), fac_lista, fac_contado, ifnull(not_email_frascos1,''), ifnull(not_email_frascos2,''), ifnull(not_email_muestras1,''), ifnull(not_email_muestras2,''), ifnull(not_email_analisis1,''), ifnull(not_email_analisis2,''), ifnull(not_email_general1,''), ifnull(not_email_general2,'') FROM cliente WHERE tipousuario = 2 AND nousar = 0 order by Nombre asc"
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim p As New dcliente
                    p.ID = CType(unaFila.Item(0), Long)
                    p.NOMBRE = CType(unaFila.Item(1), String)
                    p.EMAIL = CType(unaFila.Item(2), String)
                    p.NOMBRE_EMAIL1 = CType(unaFila.Item(3), String)
                    p.EMAIL1 = CType(unaFila.Item(4), String)
                    p.NOMBRE_EMAIL2 = CType(unaFila.Item(5), String)
                    p.EMAIL2 = CType(unaFila.Item(6), String)
                    p.ENVIO = CType(unaFila.Item(7), String)
                    p.USUARIO_WEB = CType(unaFila.Item(8), String)
                    p.NOMBRE_CELULAR1 = CType(unaFila.Item(9), String)
                    p.CELULAR = CType(unaFila.Item(10), String)
                    p.NOMBRE_CELULAR2 = CType(unaFila.Item(11), String)
                    p.CELULAR2 = CType(unaFila.Item(12), String)
                    p.CODIGOFIGARO = CType(unaFila.Item(13), String)
                    p.TIPOUSUARIO = CType(unaFila.Item(14), Integer)
                    p.DIRECCION = CType(unaFila.Item(15), String)
                    p.NOMBRE_TELEFONO1 = CType(unaFila.Item(16), String)
                    p.TELEFONO1 = CType(unaFila.Item(17), String)
                    p.NOMBRE_TELEFONO2 = CType(unaFila.Item(18), String)
                    p.TELEFONO2 = CType(unaFila.Item(19), String)
                    p.FAX = CType(unaFila.Item(20), String)
                    p.DICOSE = CType(unaFila.Item(21), String)
                    p.IDDEPARTAMENTO = CType(unaFila.Item(22), Integer)
                    p.IDLOCALIDAD = CType(unaFila.Item(23), Integer)
                    p.TECNICO1 = CType(unaFila.Item(24), Long)
                    p.TECNICO2 = CType(unaFila.Item(25), Long)
                    p.IDAGENCIA = CType(unaFila.Item(26), Integer)
                    p.CONTRATO = CType(unaFila.Item(27), Integer)
                    p.SOCIO = CType(unaFila.Item(28), Integer)
                    p.NOUSAR = CType(unaFila.Item(29), Integer)
                    p.CARAVANAS = CType(unaFila.Item(30), Integer)
                    p.PROLESA = CType(unaFila.Item(31), Integer)
                    p.PROLESASUC = CType(unaFila.Item(32), Integer)
                    p.PROLESAMAT = CType(unaFila.Item(33), Long)
                    p.OBSERVACIONES = CType(unaFila.Item(34), String)
                    p.FAC_RSOCIAL = CType(unaFila.Item(35), String)
                    p.FAC_CEDULA = CType(unaFila.Item(36), String)
                    p.FAC_RUT = CType(unaFila.Item(37), String)
                    p.FAC_DIRECCION = CType(unaFila.Item(38), String)
                    p.FAC_LOCALIDAD = CType(unaFila.Item(39), String)
                    p.FAC_DEPARTAMENTO = CType(unaFila.Item(40), Integer)
                    p.FAC_CPOSTAL = CType(unaFila.Item(41), String)
                    p.FAC_GIRO = CType(unaFila.Item(42), Integer)
                    p.COB_NOMBRE_TELEFONO1 = CType(unaFila.Item(43), String)
                    p.FAC_TELEFONOS = CType(unaFila.Item(44), String)
                    p.COB_NOMBRE_TELEFONO2 = CType(unaFila.Item(45), String)
                    p.COB_TELEFONO2 = CType(unaFila.Item(46), String)
                    p.COB_NOMBRE_CELULAR1 = CType(unaFila.Item(47), String)
                    p.COB_CELULAR1 = CType(unaFila.Item(48), String)
                    p.COB_NOMBRE_CELULAR2 = CType(unaFila.Item(49), String)
                    p.COB_CELULAR2 = CType(unaFila.Item(50), String)
                    p.COB_NOMBRE_EMAIL1 = CType(unaFila.Item(51), String)
                    p.COB_EMAIL1 = CType(unaFila.Item(52), String)
                    p.COB_NOMBRE_EMAIL2 = CType(unaFila.Item(53), String)
                    p.COB_EMAIL2 = CType(unaFila.Item(54), String)
                    p.FAC_FAX = CType(unaFila.Item(55), String)
                    p.FAC_EMAIL = CType(unaFila.Item(56), String)
                    p.FAC_CONTACTO = CType(unaFila.Item(57), String)
                    p.FAC_OBSERVACIONES = CType(unaFila.Item(58), String)
                    p.FAC_LISTA = CType(unaFila.Item(59), Integer)
                    p.FAC_CONTADO = CType(unaFila.Item(60), Integer)
                    p.NOT_EMAIL_FRASCOS1 = CType(unaFila.Item(61), String)
                    p.NOT_EMAIL_FRASCOS2 = CType(unaFila.Item(62), String)
                    p.NOT_EMAIL_MUESTRAS1 = CType(unaFila.Item(63), String)
                    p.NOT_EMAIL_MUESTRAS2 = CType(unaFila.Item(64), String)
                    p.NOT_EMAIL_ANALISIS1 = CType(unaFila.Item(65), String)
                    p.NOT_EMAIL_ANALISIS2 = CType(unaFila.Item(66), String)
                    p.NOT_EMAIL_GENERAL1 = CType(unaFila.Item(67), String)
                    p.NOT_EMAIL_GENERAL2 = CType(unaFila.Item(68), String)
                    Lista.Add(p)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function buscarPorNombreTodos(ByVal pNombre As String) As ArrayList
        Dim listaResultado As New ArrayList

        Try
            Dim Ds As New DataSet
            Dim sql As String = "SELECT id, nombre, ifnull(email,''), ifnull(nombre_email1,''), ifnull(email1,''), ifnull(nombre_email2,''), ifnull(email2,''), ifnull(envio,''), ifnull(usuario_web,''), ifnull(nombre_celular1,''), ifnull(celular,''), ifnull(nombre_celular2,''), ifnull(celular2,''), ifnull(codigofigaro,''), ifnull(tipousuario,1), ifnull(direccion,''), ifnull(nombre_telefono1,''), ifnull(telefono1,''), ifnull(nombre_telefono2,''), ifnull(telefono2,''), ifnull(fax,''), ifnull(dicose,''), ifnull(iddepartamento,999), ifnull(idlocalidad,999), ifnull(tecnico1,3197), ifnull(tecnico2,3197), ifnull(idagencia,8), contrato, socio, nousar, caravanas, prolesa, ifnull(prolesasuc,0), prolesamat, ifnull(observaciones,''), ifnull(fac_rsocial,''), ifnull(fac_cedula,''), ifnull(fac_rut,''), ifnull(fac_direccion,''), ifnull(fac_localidad,''), ifnull(fac_departamento,0), ifnull(fac_cpostal,''), ifnull(fac_giro,0), ifnull(cob_nombre_telefono1,''), ifnull(fac_telefonos,''), ifnull(cob_nombre_telefono2,''), ifnull(cob_telefono2,''), ifnull(cob_nombre_celular1,''), ifnull(cob_celular1,''), ifnull(cob_nombre_celular2,''), ifnull(cob_celular2,''), ifnull(cob_nombre_email1,''), ifnull(cob_email1,''), ifnull(cob_nombre_email2,''), ifnull(cob_email2,''), ifnull(fac_fax,''), ifnull(fac_email,''), ifnull(fac_contacto,''), ifnull(fac_observaciones,''), fac_lista, fac_contado, ifnull(not_email_frascos1,''), ifnull(not_email_frascos2,''), ifnull(not_email_muestras1,''), ifnull(not_email_muestras2,''), ifnull(not_email_analisis1,''), ifnull(not_email_analisis2,''), ifnull(not_email_general1,''), ifnull(not_email_general2,'') FROM cliente WHERE Nombre LIKE '%" & pNombre & "%'"

            Ds = Me.EjecutarSQL(sql)

            If Ds.Tables(0).Rows.Count > 0 Then
                For Each unaFila As DataRow In Ds.Tables(0).Rows
                    Dim p As New dcliente()
                    p.ID = CType(unaFila.Item(0), Long)
                    p.NOMBRE = CType(unaFila.Item(1), String)
                    p.EMAIL = CType(unaFila.Item(2), String)
                    p.NOMBRE_EMAIL1 = CType(unaFila.Item(3), String)
                    p.EMAIL1 = CType(unaFila.Item(4), String)
                    p.NOMBRE_EMAIL2 = CType(unaFila.Item(5), String)
                    p.EMAIL2 = CType(unaFila.Item(6), String)
                    p.ENVIO = CType(unaFila.Item(7), String)
                    p.USUARIO_WEB = CType(unaFila.Item(8), String)
                    p.NOMBRE_CELULAR1 = CType(unaFila.Item(9), String)
                    p.CELULAR = CType(unaFila.Item(10), String)
                    p.NOMBRE_CELULAR2 = CType(unaFila.Item(11), String)
                    p.CELULAR2 = CType(unaFila.Item(12), String)
                    p.CODIGOFIGARO = CType(unaFila.Item(13), String)
                    p.TIPOUSUARIO = CType(unaFila.Item(14), Integer)
                    p.DIRECCION = CType(unaFila.Item(15), String)
                    p.NOMBRE_TELEFONO1 = CType(unaFila.Item(16), String)
                    p.TELEFONO1 = CType(unaFila.Item(17), String)
                    p.NOMBRE_TELEFONO2 = CType(unaFila.Item(18), String)
                    p.TELEFONO2 = CType(unaFila.Item(19), String)
                    p.FAX = CType(unaFila.Item(20), String)
                    p.DICOSE = CType(unaFila.Item(21), String)
                    p.IDDEPARTAMENTO = CType(unaFila.Item(22), Integer)
                    p.IDLOCALIDAD = CType(unaFila.Item(23), Integer)
                    p.TECNICO1 = CType(unaFila.Item(24), Long)
                    p.TECNICO2 = CType(unaFila.Item(25), Long)
                    p.IDAGENCIA = CType(unaFila.Item(26), Integer)
                    p.CONTRATO = CType(unaFila.Item(27), Integer)
                    p.SOCIO = CType(unaFila.Item(28), Integer)
                    p.NOUSAR = CType(unaFila.Item(29), Integer)
                    p.CARAVANAS = CType(unaFila.Item(30), Integer)
                    p.PROLESA = CType(unaFila.Item(31), Integer)
                    p.PROLESASUC = CType(unaFila.Item(32), Integer)
                    p.PROLESAMAT = CType(unaFila.Item(33), Long)
                    p.OBSERVACIONES = CType(unaFila.Item(34), String)
                    p.FAC_RSOCIAL = CType(unaFila.Item(35), String)
                    p.FAC_CEDULA = CType(unaFila.Item(36), String)
                    p.FAC_RUT = CType(unaFila.Item(37), String)
                    p.FAC_DIRECCION = CType(unaFila.Item(38), String)
                    p.FAC_LOCALIDAD = CType(unaFila.Item(39), String)
                    p.FAC_DEPARTAMENTO = CType(unaFila.Item(40), Integer)
                    p.FAC_CPOSTAL = CType(unaFila.Item(41), String)
                    p.FAC_GIRO = CType(unaFila.Item(42), Integer)
                    p.COB_NOMBRE_TELEFONO1 = CType(unaFila.Item(43), String)
                    p.FAC_TELEFONOS = CType(unaFila.Item(44), String)
                    p.COB_NOMBRE_TELEFONO2 = CType(unaFila.Item(45), String)
                    p.COB_TELEFONO2 = CType(unaFila.Item(46), String)
                    p.COB_NOMBRE_CELULAR1 = CType(unaFila.Item(47), String)
                    p.COB_CELULAR1 = CType(unaFila.Item(48), String)
                    p.COB_NOMBRE_CELULAR2 = CType(unaFila.Item(49), String)
                    p.COB_CELULAR2 = CType(unaFila.Item(50), String)
                    p.COB_NOMBRE_EMAIL1 = CType(unaFila.Item(51), String)
                    p.COB_EMAIL1 = CType(unaFila.Item(52), String)
                    p.COB_NOMBRE_EMAIL2 = CType(unaFila.Item(53), String)
                    p.COB_EMAIL2 = CType(unaFila.Item(54), String)
                    p.FAC_FAX = CType(unaFila.Item(55), String)
                    p.FAC_EMAIL = CType(unaFila.Item(56), String)
                    p.FAC_CONTACTO = CType(unaFila.Item(57), String)
                    p.FAC_OBSERVACIONES = CType(unaFila.Item(58), String)
                    p.FAC_LISTA = CType(unaFila.Item(59), Integer)
                    p.FAC_CONTADO = CType(unaFila.Item(60), Integer)
                    p.NOT_EMAIL_FRASCOS1 = CType(unaFila.Item(61), String)
                    p.NOT_EMAIL_FRASCOS2 = CType(unaFila.Item(62), String)
                    p.NOT_EMAIL_MUESTRAS1 = CType(unaFila.Item(63), String)
                    p.NOT_EMAIL_MUESTRAS2 = CType(unaFila.Item(64), String)
                    p.NOT_EMAIL_ANALISIS1 = CType(unaFila.Item(65), String)
                    p.NOT_EMAIL_ANALISIS2 = CType(unaFila.Item(66), String)
                    p.NOT_EMAIL_GENERAL1 = CType(unaFila.Item(67), String)
                    p.NOT_EMAIL_GENERAL2 = CType(unaFila.Item(68), String)
                    listaResultado.Add(p)
                Next
                Return listaResultado
            End If
            Return listaResultado
        Catch ex As Exception
            Return listaResultado
        End Try
    End Function
    Public Function buscarPorNombre(ByVal pNombre As String) As ArrayList
        Dim listaResultado As New ArrayList

        Try
            Dim Ds As New DataSet
            Dim sql As String = "SELECT id, nombre, ifnull(email,''), ifnull(nombre_email1,''), ifnull(email1,''), ifnull(nombre_email2,''), ifnull(email2,''), ifnull(envio,''), ifnull(usuario_web,''), ifnull(nombre_celular1,''), ifnull(celular,''), ifnull(nombre_celular2,''), ifnull(celular2,''), ifnull(codigofigaro,''), ifnull(tipousuario,1), ifnull(direccion,''), ifnull(nombre_telefono1,''), ifnull(telefono1,''), ifnull(nombre_telefono2,''), ifnull(telefono2,''), ifnull(fax,''), ifnull(dicose,''), ifnull(iddepartamento,999), ifnull(idlocalidad,999), ifnull(tecnico1,3197), ifnull(tecnico2,3197), ifnull(idagencia,8), contrato, socio, nousar, caravanas, prolesa, ifnull(prolesasuc,0), prolesamat, ifnull(observaciones,''), ifnull(fac_rsocial,''), ifnull(fac_cedula,''), ifnull(fac_rut,''), ifnull(fac_direccion,''), ifnull(fac_localidad,''), ifnull(fac_departamento,0), ifnull(fac_cpostal,''), ifnull(fac_giro,0), ifnull(cob_nombre_telefono1,''), ifnull(fac_telefonos,''), ifnull(cob_nombre_telefono2,''), ifnull(cob_telefono2,''), ifnull(cob_nombre_celular1,''), ifnull(cob_celular1,''), ifnull(cob_nombre_celular2,''), ifnull(cob_celular2,''), ifnull(cob_nombre_email1,''), ifnull(cob_email1,''), ifnull(cob_nombre_email2,''), ifnull(cob_email2,''), ifnull(fac_fax,''), ifnull(fac_email,''), ifnull(fac_contacto,''), ifnull(fac_observaciones,''), fac_lista, fac_contado, ifnull(not_email_frascos1,''), ifnull(not_email_frascos2,''), ifnull(not_email_muestras1,''), ifnull(not_email_muestras2,''), ifnull(not_email_analisis1,''), ifnull(not_email_analisis2,''), ifnull(not_email_general1,''), ifnull(not_email_general2,'') FROM cliente WHERE Nombre LIKE '%" & pNombre & "%' AND nousar =0"

            Ds = Me.EjecutarSQL(sql)

            If Ds.Tables(0).Rows.Count > 0 Then
                For Each unaFila As DataRow In Ds.Tables(0).Rows
                    Dim p As New dcliente()
                    p.ID = CType(unaFila.Item(0), Long)
                    p.NOMBRE = CType(unaFila.Item(1), String)
                    p.EMAIL = CType(unaFila.Item(2), String)
                    p.NOMBRE_EMAIL1 = CType(unaFila.Item(3), String)
                    p.EMAIL1 = CType(unaFila.Item(4), String)
                    p.NOMBRE_EMAIL2 = CType(unaFila.Item(5), String)
                    p.EMAIL2 = CType(unaFila.Item(6), String)
                    p.ENVIO = CType(unaFila.Item(7), String)
                    p.USUARIO_WEB = CType(unaFila.Item(8), String)
                    p.NOMBRE_CELULAR1 = CType(unaFila.Item(9), String)
                    p.CELULAR = CType(unaFila.Item(10), String)
                    p.NOMBRE_CELULAR2 = CType(unaFila.Item(11), String)
                    p.CELULAR2 = CType(unaFila.Item(12), String)
                    p.CODIGOFIGARO = CType(unaFila.Item(13), String)
                    p.TIPOUSUARIO = CType(unaFila.Item(14), Integer)
                    p.DIRECCION = CType(unaFila.Item(15), String)
                    p.NOMBRE_TELEFONO1 = CType(unaFila.Item(16), String)
                    p.TELEFONO1 = CType(unaFila.Item(17), String)
                    p.NOMBRE_TELEFONO2 = CType(unaFila.Item(18), String)
                    p.TELEFONO2 = CType(unaFila.Item(19), String)
                    p.FAX = CType(unaFila.Item(20), String)
                    p.DICOSE = CType(unaFila.Item(21), String)
                    p.IDDEPARTAMENTO = CType(unaFila.Item(22), Integer)
                    p.IDLOCALIDAD = CType(unaFila.Item(23), Integer)
                    p.TECNICO1 = CType(unaFila.Item(24), Long)
                    p.TECNICO2 = CType(unaFila.Item(25), Long)
                    p.IDAGENCIA = CType(unaFila.Item(26), Integer)
                    p.CONTRATO = CType(unaFila.Item(27), Integer)
                    p.SOCIO = CType(unaFila.Item(28), Integer)
                    p.NOUSAR = CType(unaFila.Item(29), Integer)
                    p.CARAVANAS = CType(unaFila.Item(30), Integer)
                    p.PROLESA = CType(unaFila.Item(31), Integer)
                    p.PROLESASUC = CType(unaFila.Item(32), Integer)
                    p.PROLESAMAT = CType(unaFila.Item(33), Long)
                    p.OBSERVACIONES = CType(unaFila.Item(34), String)
                    p.FAC_RSOCIAL = CType(unaFila.Item(35), String)
                    p.FAC_CEDULA = CType(unaFila.Item(36), String)
                    p.FAC_RUT = CType(unaFila.Item(37), String)
                    p.FAC_DIRECCION = CType(unaFila.Item(38), String)
                    p.FAC_LOCALIDAD = CType(unaFila.Item(39), String)
                    p.FAC_DEPARTAMENTO = CType(unaFila.Item(40), Integer)
                    p.FAC_CPOSTAL = CType(unaFila.Item(41), String)
                    p.FAC_GIRO = CType(unaFila.Item(42), Integer)
                    p.COB_NOMBRE_TELEFONO1 = CType(unaFila.Item(43), String)
                    p.FAC_TELEFONOS = CType(unaFila.Item(44), String)
                    p.COB_NOMBRE_TELEFONO2 = CType(unaFila.Item(45), String)
                    p.COB_TELEFONO2 = CType(unaFila.Item(46), String)
                    p.COB_NOMBRE_CELULAR1 = CType(unaFila.Item(47), String)
                    p.COB_CELULAR1 = CType(unaFila.Item(48), String)
                    p.COB_NOMBRE_CELULAR2 = CType(unaFila.Item(49), String)
                    p.COB_CELULAR2 = CType(unaFila.Item(50), String)
                    p.COB_NOMBRE_EMAIL1 = CType(unaFila.Item(51), String)
                    p.COB_EMAIL1 = CType(unaFila.Item(52), String)
                    p.COB_NOMBRE_EMAIL2 = CType(unaFila.Item(53), String)
                    p.COB_EMAIL2 = CType(unaFila.Item(54), String)
                    p.FAC_FAX = CType(unaFila.Item(55), String)
                    p.FAC_EMAIL = CType(unaFila.Item(56), String)
                    p.FAC_CONTACTO = CType(unaFila.Item(57), String)
                    p.FAC_OBSERVACIONES = CType(unaFila.Item(58), String)
                    p.FAC_LISTA = CType(unaFila.Item(59), Integer)
                    p.FAC_CONTADO = CType(unaFila.Item(60), Integer)
                    p.NOT_EMAIL_FRASCOS1 = CType(unaFila.Item(61), String)
                    p.NOT_EMAIL_FRASCOS2 = CType(unaFila.Item(62), String)
                    p.NOT_EMAIL_MUESTRAS1 = CType(unaFila.Item(63), String)
                    p.NOT_EMAIL_MUESTRAS2 = CType(unaFila.Item(64), String)
                    p.NOT_EMAIL_ANALISIS1 = CType(unaFila.Item(65), String)
                    p.NOT_EMAIL_ANALISIS2 = CType(unaFila.Item(66), String)
                    p.NOT_EMAIL_GENERAL1 = CType(unaFila.Item(67), String)
                    p.NOT_EMAIL_GENERAL2 = CType(unaFila.Item(68), String)
                    listaResultado.Add(p)
                Next
                Return listaResultado
            End If
            Return listaResultado
        Catch ex As Exception
            Return listaResultado
        End Try
    End Function

    Public Function buscarPorNombreEmpresa(ByVal pNombre As String) As ArrayList
        Dim listaResultado As New ArrayList

        Try
            Dim Ds As New DataSet
            Dim sql As String = "SELECT SELECT id, nombre, ifnull(email,''), ifnull(nombre_email1,''), ifnull(email1,''), ifnull(nombre_email2,''), ifnull(email2,''), ifnull(envio,''), ifnull(usuario_web,''), ifnull(nombre_celular1,''), ifnull(celular,''), ifnull(nombre_celular2,''), ifnull(celular2,''), ifnull(codigofigaro,''), ifnull(tipousuario,1), ifnull(direccion,''), ifnull(nombre_telefono1,''), ifnull(telefono1,''), ifnull(nombre_telefono2,''), ifnull(telefono2,''), ifnull(fax,''), ifnull(dicose,''), ifnull(iddepartamento,999), ifnull(idlocalidad,999), ifnull(tecnico1,3197), ifnull(tecnico2,3197), ifnull(idagencia,8), contrato, socio, nousar, caravanas, prolesa, ifnull(prolesasuc,0), prolesamat, ifnull(observaciones,''), ifnull(fac_rsocial,''), ifnull(fac_cedula,''), ifnull(fac_rut,''), ifnull(fac_direccion,''), ifnull(fac_localidad,''), ifnull(fac_departamento,0), ifnull(fac_cpostal,''), ifnull(fac_giro,0), ifnull(cob_nombre_telefono1,''), ifnull(fac_telefonos,''), ifnull(cob_nombre_telefono2,''), ifnull(cob_telefono2,''), ifnull(cob_nombre_celular1,''), ifnull(cob_celular1,''), ifnull(cob_nombre_celular2,''), ifnull(cob_celular2,''), ifnull(cob_nombre_email1,''), ifnull(cob_email1,''), ifnull(cob_nombre_email2,''), ifnull(cob_email2,''), ifnull(fac_fax,''), ifnull(fac_email,''), ifnull(fac_contacto,''), ifnull(fac_observaciones,''), fac_lista, fac_contado, ifnull(not_email_frascos1,''), ifnull(not_email_frascos2,''), ifnull(not_email_muestras1,''), ifnull(not_email_muestras2,''), ifnull(not_email_analisis1,''), ifnull(not_email_analisis2,''), ifnull(not_email_general1,''), ifnull(not_email_general2,'') FROM cliente WHERE Nombre LIKE '%" & pNombre & "%' AND tipousuario = 2 AND nousar= 0"

            Ds = Me.EjecutarSQL(sql)

            If Ds.Tables(0).Rows.Count > 0 Then
                For Each unaFila As DataRow In Ds.Tables(0).Rows
                    Dim p As New dcliente()
                    p.ID = CType(unaFila.Item(0), Long)
                    p.NOMBRE = CType(unaFila.Item(1), String)
                    p.EMAIL = CType(unaFila.Item(2), String)
                    p.NOMBRE_EMAIL1 = CType(unaFila.Item(3), String)
                    p.EMAIL1 = CType(unaFila.Item(4), String)
                    p.NOMBRE_EMAIL2 = CType(unaFila.Item(5), String)
                    p.EMAIL2 = CType(unaFila.Item(6), String)
                    p.ENVIO = CType(unaFila.Item(7), String)
                    p.USUARIO_WEB = CType(unaFila.Item(8), String)
                    p.NOMBRE_CELULAR1 = CType(unaFila.Item(9), String)
                    p.CELULAR = CType(unaFila.Item(10), String)
                    p.NOMBRE_CELULAR2 = CType(unaFila.Item(11), String)
                    p.CELULAR2 = CType(unaFila.Item(12), String)
                    p.CODIGOFIGARO = CType(unaFila.Item(13), String)
                    p.TIPOUSUARIO = CType(unaFila.Item(14), Integer)
                    p.DIRECCION = CType(unaFila.Item(15), String)
                    p.NOMBRE_TELEFONO1 = CType(unaFila.Item(16), String)
                    p.TELEFONO1 = CType(unaFila.Item(17), String)
                    p.NOMBRE_TELEFONO2 = CType(unaFila.Item(18), String)
                    p.TELEFONO2 = CType(unaFila.Item(19), String)
                    p.FAX = CType(unaFila.Item(20), String)
                    p.DICOSE = CType(unaFila.Item(21), String)
                    p.IDDEPARTAMENTO = CType(unaFila.Item(22), Integer)
                    p.IDLOCALIDAD = CType(unaFila.Item(23), Integer)
                    p.TECNICO1 = CType(unaFila.Item(24), Long)
                    p.TECNICO2 = CType(unaFila.Item(25), Long)
                    p.IDAGENCIA = CType(unaFila.Item(26), Integer)
                    p.CONTRATO = CType(unaFila.Item(27), Integer)
                    p.SOCIO = CType(unaFila.Item(28), Integer)
                    p.NOUSAR = CType(unaFila.Item(29), Integer)
                    p.CARAVANAS = CType(unaFila.Item(30), Integer)
                    p.PROLESA = CType(unaFila.Item(31), Integer)
                    p.PROLESASUC = CType(unaFila.Item(32), Integer)
                    p.PROLESAMAT = CType(unaFila.Item(33), Long)
                    p.OBSERVACIONES = CType(unaFila.Item(34), String)
                    p.FAC_RSOCIAL = CType(unaFila.Item(35), String)
                    p.FAC_CEDULA = CType(unaFila.Item(36), String)
                    p.FAC_RUT = CType(unaFila.Item(37), String)
                    p.FAC_DIRECCION = CType(unaFila.Item(38), String)
                    p.FAC_LOCALIDAD = CType(unaFila.Item(39), String)
                    p.FAC_DEPARTAMENTO = CType(unaFila.Item(40), Integer)
                    p.FAC_CPOSTAL = CType(unaFila.Item(41), String)
                    p.FAC_GIRO = CType(unaFila.Item(42), Integer)
                    p.COB_NOMBRE_TELEFONO1 = CType(unaFila.Item(43), String)
                    p.FAC_TELEFONOS = CType(unaFila.Item(44), String)
                    p.COB_NOMBRE_TELEFONO2 = CType(unaFila.Item(45), String)
                    p.COB_TELEFONO2 = CType(unaFila.Item(46), String)
                    p.COB_NOMBRE_CELULAR1 = CType(unaFila.Item(47), String)
                    p.COB_CELULAR1 = CType(unaFila.Item(48), String)
                    p.COB_NOMBRE_CELULAR2 = CType(unaFila.Item(49), String)
                    p.COB_CELULAR2 = CType(unaFila.Item(50), String)
                    p.COB_NOMBRE_EMAIL1 = CType(unaFila.Item(51), String)
                    p.COB_EMAIL1 = CType(unaFila.Item(52), String)
                    p.COB_NOMBRE_EMAIL2 = CType(unaFila.Item(53), String)
                    p.COB_EMAIL2 = CType(unaFila.Item(54), String)
                    p.FAC_FAX = CType(unaFila.Item(55), String)
                    p.FAC_EMAIL = CType(unaFila.Item(56), String)
                    p.FAC_CONTACTO = CType(unaFila.Item(57), String)
                    p.FAC_OBSERVACIONES = CType(unaFila.Item(58), String)
                    p.FAC_LISTA = CType(unaFila.Item(59), Integer)
                    p.FAC_CONTADO = CType(unaFila.Item(60), Integer)
                    p.NOT_EMAIL_FRASCOS1 = CType(unaFila.Item(61), String)
                    p.NOT_EMAIL_FRASCOS2 = CType(unaFila.Item(62), String)
                    p.NOT_EMAIL_MUESTRAS1 = CType(unaFila.Item(63), String)
                    p.NOT_EMAIL_MUESTRAS2 = CType(unaFila.Item(64), String)
                    p.NOT_EMAIL_ANALISIS1 = CType(unaFila.Item(65), String)
                    p.NOT_EMAIL_ANALISIS2 = CType(unaFila.Item(66), String)
                    p.NOT_EMAIL_GENERAL1 = CType(unaFila.Item(67), String)
                    p.NOT_EMAIL_GENERAL2 = CType(unaFila.Item(68), String)
                    listaResultado.Add(p)
                Next
                Return listaResultado
            End If
            Return listaResultado
        Catch ex As Exception
            Return listaResultado
        End Try
    End Function
    Public Function buscarPorId(ByVal pId As Long) As ArrayList
        Dim listaResultado As New ArrayList

        Try
            Dim Ds As New DataSet
            Dim sql As String = "SELECT id, nombre, ifnull(email,''), ifnull(nombre_email1,''), ifnull(email1,''), ifnull(nombre_email2,''), ifnull(email2,''), ifnull(envio,''), ifnull(usuario_web,''), ifnull(nombre_celular1,''), ifnull(celular,''), ifnull(nombre_celular2,''), ifnull(celular2,''), ifnull(codigofigaro,''), ifnull(tipousuario,1), ifnull(direccion,''), ifnull(nombre_telefono1,''), ifnull(telefono1,''), ifnull(nombre_telefono2,''), ifnull(telefono2,''), ifnull(fax,''), ifnull(dicose,''), ifnull(iddepartamento,999), ifnull(idlocalidad,999), ifnull(tecnico1,3197), ifnull(tecnico2,3197), ifnull(idagencia,8), contrato, socio, nousar, caravanas, prolesa, ifnull(prolesasuc,0), prolesamat, ifnull(observaciones,''), ifnull(fac_rsocial,''), ifnull(fac_cedula,''), ifnull(fac_rut,''), ifnull(fac_direccion,''), ifnull(fac_localidad,''), ifnull(fac_departamento,0), ifnull(fac_cpostal,''), ifnull(fac_giro,0), ifnull(cob_nombre_telefono1,''), ifnull(fac_telefonos,''), ifnull(cob_nombre_telefono2,''), ifnull(cob_telefono2,''), ifnull(cob_nombre_celular1,''), ifnull(cob_celular1,''), ifnull(cob_nombre_celular2,''), ifnull(cob_celular2,''), ifnull(cob_nombre_email1,''), ifnull(cob_email1,''), ifnull(cob_nombre_email2,''), ifnull(cob_email2,''), ifnull(fac_fax,''), ifnull(fac_email,''), ifnull(fac_contacto,''), ifnull(fac_observaciones,''), fac_lista, fac_contado, ifnull(not_email_frascos1,''), ifnull(not_email_frascos2,''), ifnull(not_email_muestras1,''), ifnull(not_email_muestras2,''), ifnull(not_email_analisis1,''), ifnull(not_email_analisis2,''), ifnull(not_email_general1,''), ifnull(not_email_general2,'') FROM cliente WHERE id = pId"

            Ds = Me.EjecutarSQL(sql)

            If Ds.Tables(0).Rows.Count > 0 Then
                For Each unaFila As DataRow In Ds.Tables(0).Rows
                    Dim p As New dcliente()
                    p.ID = CType(unaFila.Item(0), Long)
                    p.NOMBRE = CType(unaFila.Item(1), String)
                    p.EMAIL = CType(unaFila.Item(2), String)
                    p.NOMBRE_EMAIL1 = CType(unaFila.Item(3), String)
                    p.EMAIL1 = CType(unaFila.Item(4), String)
                    p.NOMBRE_EMAIL2 = CType(unaFila.Item(5), String)
                    p.EMAIL2 = CType(unaFila.Item(6), String)
                    p.ENVIO = CType(unaFila.Item(7), String)
                    p.USUARIO_WEB = CType(unaFila.Item(8), String)
                    p.NOMBRE_CELULAR1 = CType(unaFila.Item(9), String)
                    p.CELULAR = CType(unaFila.Item(10), String)
                    p.NOMBRE_CELULAR2 = CType(unaFila.Item(11), String)
                    p.CELULAR2 = CType(unaFila.Item(12), String)
                    p.CODIGOFIGARO = CType(unaFila.Item(13), String)
                    p.TIPOUSUARIO = CType(unaFila.Item(14), Integer)
                    p.DIRECCION = CType(unaFila.Item(15), String)
                    p.NOMBRE_TELEFONO1 = CType(unaFila.Item(16), String)
                    p.TELEFONO1 = CType(unaFila.Item(17), String)
                    p.NOMBRE_TELEFONO2 = CType(unaFila.Item(18), String)
                    p.TELEFONO2 = CType(unaFila.Item(19), String)
                    p.FAX = CType(unaFila.Item(20), String)
                    p.DICOSE = CType(unaFila.Item(21), String)
                    p.IDDEPARTAMENTO = CType(unaFila.Item(22), Integer)
                    p.IDLOCALIDAD = CType(unaFila.Item(23), Integer)
                    p.TECNICO1 = CType(unaFila.Item(24), Long)
                    p.TECNICO2 = CType(unaFila.Item(25), Long)
                    p.IDAGENCIA = CType(unaFila.Item(26), Integer)
                    p.CONTRATO = CType(unaFila.Item(27), Integer)
                    p.SOCIO = CType(unaFila.Item(28), Integer)
                    p.NOUSAR = CType(unaFila.Item(29), Integer)
                    p.CARAVANAS = CType(unaFila.Item(30), Integer)
                    p.PROLESA = CType(unaFila.Item(31), Integer)
                    p.PROLESASUC = CType(unaFila.Item(32), Integer)
                    p.PROLESAMAT = CType(unaFila.Item(33), Long)
                    p.OBSERVACIONES = CType(unaFila.Item(34), String)
                    p.FAC_RSOCIAL = CType(unaFila.Item(35), String)
                    p.FAC_CEDULA = CType(unaFila.Item(36), String)
                    p.FAC_RUT = CType(unaFila.Item(37), String)
                    p.FAC_DIRECCION = CType(unaFila.Item(38), String)
                    p.FAC_LOCALIDAD = CType(unaFila.Item(39), String)
                    p.FAC_DEPARTAMENTO = CType(unaFila.Item(40), Integer)
                    p.FAC_CPOSTAL = CType(unaFila.Item(41), String)
                    p.FAC_GIRO = CType(unaFila.Item(42), Integer)
                    p.COB_NOMBRE_TELEFONO1 = CType(unaFila.Item(43), String)
                    p.FAC_TELEFONOS = CType(unaFila.Item(44), String)
                    p.COB_NOMBRE_TELEFONO2 = CType(unaFila.Item(45), String)
                    p.COB_TELEFONO2 = CType(unaFila.Item(46), String)
                    p.COB_NOMBRE_CELULAR1 = CType(unaFila.Item(47), String)
                    p.COB_CELULAR1 = CType(unaFila.Item(48), String)
                    p.COB_NOMBRE_CELULAR2 = CType(unaFila.Item(49), String)
                    p.COB_CELULAR2 = CType(unaFila.Item(50), String)
                    p.COB_NOMBRE_EMAIL1 = CType(unaFila.Item(51), String)
                    p.COB_EMAIL1 = CType(unaFila.Item(52), String)
                    p.COB_NOMBRE_EMAIL2 = CType(unaFila.Item(53), String)
                    p.COB_EMAIL2 = CType(unaFila.Item(54), String)
                    p.FAC_FAX = CType(unaFila.Item(55), String)
                    p.FAC_EMAIL = CType(unaFila.Item(56), String)
                    p.FAC_CONTACTO = CType(unaFila.Item(57), String)
                    p.FAC_OBSERVACIONES = CType(unaFila.Item(58), String)
                    p.FAC_LISTA = CType(unaFila.Item(59), Integer)
                    p.FAC_CONTADO = CType(unaFila.Item(60), Integer)
                    p.NOT_EMAIL_FRASCOS1 = CType(unaFila.Item(61), String)
                    p.NOT_EMAIL_FRASCOS2 = CType(unaFila.Item(62), String)
                    p.NOT_EMAIL_MUESTRAS1 = CType(unaFila.Item(63), String)
                    p.NOT_EMAIL_MUESTRAS2 = CType(unaFila.Item(64), String)
                    p.NOT_EMAIL_ANALISIS1 = CType(unaFila.Item(65), String)
                    p.NOT_EMAIL_ANALISIS2 = CType(unaFila.Item(66), String)
                    p.NOT_EMAIL_GENERAL1 = CType(unaFila.Item(67), String)
                    p.NOT_EMAIL_GENERAL2 = CType(unaFila.Item(68), String)
                    listaResultado.Add(p)
                Next
                Return listaResultado
            End If
            Return listaResultado
        Catch ex As Exception
            Return listaResultado
        End Try
    End Function
  
    Public Function actualizardireccion(ByVal idcliente As Integer, ByVal direnvio As String) As Boolean
        Dim sql As String = "UPDATE cliente SET envio = '" & direnvio & "' WHERE id = " & idcliente

       

        Dim lista As New ArrayList : lista.Add(sql)
        Return EjecutarTransaccion(lista)
    End Function
    'Public Function actualizartelefono(ByVal idcliente As Integer, ByVal tel As String, ByVal usuario As dUsuario) As Boolean
    '    Dim sql As String = "UPDATE cliente SET telefono = '" & tel & "' WHERE id = " & idcliente

    '    Dim sqlAccion As String = "INSERT INTO actividad (act_fecha, act_tabla, act_accion, act_registro, u_id) " _
    '                                & "VALUES (now(), 'pedidos', 'actualizar teléfono', " & idcliente & ", " & usuario.ID & ")"

    '    Dim lista As New ArrayList : lista.Add(sql) : lista.Add(sqlAccion)
    '    Return EjecutarTransaccion(lista)
    'End Function
    Public Function actualizartecnico1(ByVal idcliente As Integer, ByVal tec As Long) As Boolean
        Dim sql As String = "UPDATE cliente SET tecnico1 = '" & tec & "' WHERE id = " & idcliente & ""

       

        Dim lista As New ArrayList : lista.Add(sql)
        Return EjecutarTransaccion(lista)
    End Function
    Public Function actualizartecnico2(ByVal idcliente As Integer, ByVal tec2 As Long) As Boolean
        Dim sql As String = "UPDATE cliente SET tecnico2 = '" & tec2 & "' WHERE id = " & idcliente & ""

       

        Dim lista As New ArrayList : lista.Add(sql)
        Return EjecutarTransaccion(lista)
    End Function
    Public Function actualizaragencia(ByVal idcliente As Integer, ByVal age As Long) As Boolean
        Dim sql As String = "UPDATE cliente SET idagencia = '" & age & "' WHERE id = " & idcliente

      

        Dim lista As New ArrayList : lista.Add(sql)
        Return EjecutarTransaccion(lista)
    End Function
    'Public Function actualizarmail(ByVal idcliente As Integer, ByVal mail As String, ByVal usuario As dUsuario) As Boolean
    '    Dim sql As String = "UPDATE cliente SET Email1 = '" & mail & "' WHERE id = " & idcliente

    '    Dim sqlAccion As String = "INSERT INTO actividad (act_fecha, act_tabla, act_accion, act_registro, u_id) " _
    '                                & "VALUES (now(), 'pedidos', 'actualizar email', " & idcliente & ", " & usuario.ID & ")"

    '    Dim lista As New ArrayList : lista.Add(sql) : lista.Add(sqlAccion)
    '    Return EjecutarTransaccion(lista)
    'End Function
    Public Function actualizardicose(ByVal idcliente As Integer, ByVal dicose As String) As Boolean
        Dim sql As String = "UPDATE cliente SET dicose = '" & dicose & "' WHERE id = " & idcliente

      

        Dim lista As New ArrayList : lista.Add(sql)
        Return EjecutarTransaccion(lista)
    End Function
   
  
End Class
