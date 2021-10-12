Public Class pProveedores
    Inherits Conectoras.ConexionMySQL
    Public Function guardar(ByVal o As Object) As Boolean
        Dim obj As dProveedores = CType(o, dProveedores)
        Dim sql As String = "INSERT INTO proveedores (id, nombre, telefono, direccion, email, email2, email3, contacto, otrosdatos) VALUES (" & obj.ID & ", '" & obj.NOMBRE & "', '" & obj.TELEFONO & "','" & obj.DIRECCION & "','" & obj.EMAIL & "','" & obj.EMAIL2 & "','" & obj.EMAIL3 & "', '" & obj.CONTACTO & "', '" & obj.OTROSDATOS & "')"

        Dim lista As New ArrayList
        lista.Add(sql)



        Return EjecutarTransaccion(lista)
    End Function
    Public Function modificar(ByVal o As Object) As Boolean
        Dim obj As dProveedores = CType(o, dProveedores)
        Dim sql As String = "UPDATE proveedores SET nombre ='" & obj.NOMBRE & "',telefono ='" & obj.TELEFONO & "',direccion ='" & obj.DIRECCION & "',email ='" & obj.EMAIL & "',email2 ='" & obj.EMAIL2 & "',email3 ='" & obj.EMAIL3 & "',contacto ='" & obj.CONTACTO & "', otrosdatos = '" & obj.OTROSDATOS & "' WHERE id = " & obj.ID & ""

        Dim lista As New ArrayList
        lista.Add(sql)



        Return EjecutarTransaccion(lista)
    End Function
    Public Function eliminar(ByVal o As Object) As Boolean
        Dim obj As dProveedores = CType(o, dProveedores)
        Dim sql As String = "DELETE FROM proveedores WHERE id = " & obj.ID & ""

        Dim lista As New ArrayList
        lista.Add(sql)



        Return EjecutarTransaccion(lista)
    End Function
    Public Function buscar(ByVal o As Object) As dProveedores
        Dim obj As dProveedores = CType(o, dProveedores)
        Dim p As New dProveedores
        Try
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL("SELECT id, nombre, ifnull(telefono,''), ifnull(direccion,''), ifnull(email,''), ifnull(email2,''), ifnull(email3,''), ifnull(contacto,''), ifnull(otrosdatos,'') FROM proveedores WHERE id = " & obj.ID & "")

            If Ds.Tables(0).Rows.Count > 0 Then
                Dim unaFila As DataRow
                unaFila = Ds.Tables(0).Rows(0)
                p.ID = CType(unaFila.Item(0), Integer)
                p.NOMBRE = CType(unaFila.Item(1), String)
                p.TELEFONO = CType(unaFila.Item(2), String)
                p.DIRECCION = CType(unaFila.Item(3), String)
                p.EMAIL = CType(unaFila.Item(4), String)
                p.EMAIL2 = CType(unaFila.Item(5), String)
                p.EMAIL3 = CType(unaFila.Item(6), String)
                p.CONTACTO = CType(unaFila.Item(7), String)
                p.OTROSDATOS = CType(unaFila.Item(8), String)
                Return p
            End If
            Return Nothing
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listar() As ArrayList
        Dim sql As String = "SELECT id, nombre, ifnull(telefono,''), ifnull(direccion,''), ifnull(email,''), ifnull(email2,''), ifnull(email3,''), ifnull(contacto,''), ifnull(otrosdatos,'') FROM proveedores ORDER BY nombre ASC"
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim p As New dProveedores
                    p.ID = CType(unaFila.Item(0), Integer)
                    p.NOMBRE = CType(unaFila.Item(1), String)
                    p.TELEFONO = CType(unaFila.Item(2), String)
                    p.DIRECCION = CType(unaFila.Item(3), String)
                    p.EMAIL = CType(unaFila.Item(4), String)
                    p.EMAIL2 = CType(unaFila.Item(5), String)
                    p.EMAIL3 = CType(unaFila.Item(6), String)
                    p.CONTACTO = CType(unaFila.Item(7), String)
                    p.OTROSDATOS = CType(unaFila.Item(8), String)
                    Lista.Add(p)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Function buscarPorNombre(ByVal pNombre As String) As ArrayList
        Dim listaResultado As New ArrayList

        Try
            Dim Ds As New DataSet
            Dim sql As String = "SELECT id, nombre, ifnull(telefono,''), ifnull(direccion,''), ifnull(email,''), ifnull(email2,''), ifnull(email3,''), ifnull(contacto,''), ifnull(otrosdatos,'') FROM proveedores WHERE Nombre LIKE '%" & pNombre & "%' "

            Ds = Me.EjecutarSQL(sql)

            If Ds.Tables(0).Rows.Count > 0 Then
                For Each unaFila As DataRow In Ds.Tables(0).Rows
                    Dim p As New dProveedores()
                    p.ID = CType(unaFila.Item(0), Integer)
                    p.NOMBRE = CType(unaFila.Item(1), String)
                    p.TELEFONO = CType(unaFila.Item(2), String)
                    p.DIRECCION = CType(unaFila.Item(3), String)
                    p.EMAIL = CType(unaFila.Item(4), String)
                    p.EMAIL2 = CType(unaFila.Item(5), String)
                    p.EMAIL3 = CType(unaFila.Item(6), String)
                    p.CONTACTO = CType(unaFila.Item(7), String)
                    p.OTROSDATOS = CType(unaFila.Item(8), String)
                    listaResultado.Add(p)
                Next
                Return listaResultado
            End If
            Return listaResultado
        Catch ex As Exception
            Return listaResultado
        End Try
    End Function
End Class
