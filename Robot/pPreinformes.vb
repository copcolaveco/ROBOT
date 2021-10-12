Public Class pPreinformes
    Inherits Conectoras.ConexionMySQL
    Public Function guardar(ByVal o As Object) As Boolean
        Dim obj As dPreinformes = CType(o, dPreinformes)
        Dim sql As String = "INSERT INTO preinformes (id, ficha, tipo, creado, abonado, comentario, copia, parasubir, subido, fecha, control) VALUES (" & obj.ID & ", " & obj.FICHA & ", " & obj.TIPO & ", " & obj.CREADO & ", " & obj.ABONADO & ", '" & obj.COMENTARIO & "', '" & obj.COPIA & "', " & obj.PARASUBIR & ", " & obj.SUBIDO & ", '" & obj.FECHA & "', " & obj.CONTROL & ")"

        Dim lista As New ArrayList
        lista.Add(sql)

        'Dim sqlAccion As String = "INSERT INTO actividad (act_fecha, act_tabla, act_accion, act_registro, u_id) " _
        '                         & "VALUES (now(), 'preinformes', 'alta', last_insert_id(), " & usuario.ID & ")"
        'lista.Add(sqlAccion)

        Return EjecutarTransaccion(lista)
    End Function
    Public Function modificar(ByVal o As Object) As Boolean
        Dim obj As dPreinformes = CType(o, dPreinformes)
        Dim sql As String = "UPDATE preinformes SET ficha = " & obj.FICHA & ",tipo = " & obj.TIPO & ", creado= " & obj.CREADO & ", abonado=" & obj.ABONADO & ", comentario = '" & obj.COMENTARIO & "', copia = '" & obj.COPIA & "', parasubir=" & obj.PARASUBIR & ", subido=" & obj.SUBIDO & ", fecha='" & obj.FECHA & "', control=" & obj.CONTROL & " WHERE id = " & obj.ID & ""

        Dim lista As New ArrayList
        lista.Add(sql)

        'Dim sqlAccion As String = "INSERT INTO actividad (act_fecha, act_tabla, act_accion, act_registro, u_id) " _
        '                         & "VALUES (now(), 'preinformes', 'creado', " & obj.ID & ", " & usuario.ID & ")"
        'lista.Add(sqlAccion)

        Return EjecutarTransaccion(lista)
    End Function
    Public Function modificar2(ByVal o As Object) As Boolean
        Dim obj As dPreinformes = CType(o, dPreinformes)
        Dim sql As String = "UPDATE preinformes SET abonado=" & obj.ABONADO & ", comentario = '" & obj.COMENTARIO & "', copia = '" & obj.COPIA & "', parasubir = " & obj.PARASUBIR & ", fecha = '" & obj.FECHA & "' WHERE ficha = " & obj.FICHA & ""

        Dim lista As New ArrayList
        lista.Add(sql)

        'Dim sqlAccion As String = "INSERT INTO actividad (act_fecha, act_tabla, act_accion, act_registro, u_id) " _
        '                         & "VALUES (now(), 'preinformes', 'creado', " & obj.ID & ", " & usuario.ID & ")"
        'lista.Add(sqlAccion)

        Return EjecutarTransaccion(lista)
    End Function
    Public Function marcarcreado(ByVal o As Object) As Boolean
        Dim obj As dPreinformes = CType(o, dPreinformes)
        Dim sql As String = "UPDATE preinformes SET creado= 1 WHERE ficha = " & obj.FICHA & ""

        Dim lista As New ArrayList
        lista.Add(sql)

        'Dim sqlAccion As String = "INSERT INTO actividad (act_fecha, act_tabla, act_accion, act_registro, u_id) " _
        '                         & "VALUES (now(), 'preinformes', 'creado', " & obj.ID & ", " & usuario.ID & ")"
        'lista.Add(sqlAccion)

        Return EjecutarTransaccion(lista)
    End Function
    Public Function desmarcarcreado(ByVal o As Object) As Boolean
        Dim obj As dPreinformes = CType(o, dPreinformes)
        Dim sql As String = "UPDATE preinformes SET creado= 0 WHERE ficha = " & obj.FICHA & ""

        Dim lista As New ArrayList
        lista.Add(sql)

        'Dim sqlAccion As String = "INSERT INTO actividad (act_fecha, act_tabla, act_accion, act_registro, u_id) " _
        '                         & "VALUES (now(), 'preinformes', 'creado', " & obj.ID & ", " & usuario.ID & ")"
        'lista.Add(sqlAccion)

        Return EjecutarTransaccion(lista)
    End Function
    Public Function marcarsubido(ByVal o As Object, ByVal fecha As String) As Boolean
        Dim obj As dPreinformes = CType(o, dPreinformes)
        Dim sql As String = "UPDATE preinformes SET subido= 1, fecha= '" & fecha & "' WHERE ficha = " & obj.FICHA & ""

        Dim lista As New ArrayList
        lista.Add(sql)

        'Dim sqlAccion As String = "INSERT INTO actividad (act_fecha, act_tabla, act_accion, act_registro, u_id) " _
        '                         & "VALUES (now(), 'preinformes', 'creado', " & obj.ID & ", " & usuario.ID & ")"
        'lista.Add(sqlAccion)

        Return EjecutarTransaccion(lista)
    End Function
    Public Function marcarcontrolado(ByVal o As Object) As Boolean
        Dim obj As dPreinformes = CType(o, dPreinformes)
        Dim sql As String = "UPDATE preinformes SET control= 0 WHERE ficha = " & obj.FICHA & ""

        Dim lista As New ArrayList
        lista.Add(sql)

      

        Return EjecutarTransaccion(lista)
    End Function
    Public Function eliminar(ByVal o As Object) As Boolean
        Dim obj As dPreinformes = CType(o, dPreinformes)
        Dim sql As String = "DELETE FROM preinformes WHERE id = " & obj.ID & ""

        Dim lista As New ArrayList
        lista.Add(sql)

       

        Return EjecutarTransaccion(lista)
    End Function
    Public Function buscar(ByVal o As Object) As dPreinformes
        Dim obj As dPreinformes = CType(o, dPreinformes)
        Dim p As New dPreinformes
        Try
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL("SELECT id, ficha, tipo, creado, abonado, comentario, copia, parasubir, subido, fecha, control FROM preinformes WHERE ficha = " & obj.FICHA & "")

            If Ds.Tables(0).Rows.Count > 0 Then
                Dim unaFila As DataRow
                unaFila = Ds.Tables(0).Rows(0)
                p.ID = CType(unaFila.Item(0), Long)
                p.FICHA = CType(unaFila.Item(1), Long)
                p.TIPO = CType(unaFila.Item(2), Integer)
                p.CREADO = CType(unaFila.Item(3), Integer)
                p.ABONADO = CType(unaFila.Item(4), Integer)
                p.COMENTARIO = CType(unaFila.Item(5), String)
                p.COPIA = CType(unaFila.Item(6), String)
                p.PARASUBIR = CType(unaFila.Item(7), Integer)
                p.SUBIDO = CType(unaFila.Item(8), Integer)
                p.FECHA = CType(unaFila.Item(9), String)
                p.CONTROL = CType(unaFila.Item(10), Integer)
                Return p
            End If
            Return Nothing
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listar() As ArrayList
        Dim sql As String = "SELECT id, ficha, tipo, creado, abonado, comentario, copia, parasubir, subido, fecha, control FROM preinformes"
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim p As New dPreinformes
                    p.ID = CType(unaFila.Item(0), Long)
                    p.FICHA = CType(unaFila.Item(1), Long)
                    p.TIPO = CType(unaFila.Item(2), Integer)
                    p.CREADO = CType(unaFila.Item(3), Integer)
                    p.ABONADO = CType(unaFila.Item(4), Integer)
                    p.COMENTARIO = CType(unaFila.Item(5), String)
                    p.COPIA = CType(unaFila.Item(6), String)
                    p.PARASUBIR = CType(unaFila.Item(7), Integer)
                    p.SUBIDO = CType(unaFila.Item(8), Integer)
                    p.FECHA = CType(unaFila.Item(9), String)
                    p.CONTROL = CType(unaFila.Item(10), Integer)
                    Lista.Add(p)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listarsinmarcar() As ArrayList
        Dim sql As String = "SELECT id, ficha, tipo, creado, abonado, ifnull(comentario,''),ifnull(copia,''), parasubir, subido, fecha, control FROM preinformes WHERE creado =0 and ficha > 0 "
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim p As New dPreinformes
                    p.ID = CType(unaFila.Item(0), Long)
                    p.FICHA = CType(unaFila.Item(1), Long)
                    p.TIPO = CType(unaFila.Item(2), Integer)
                    p.CREADO = CType(unaFila.Item(3), Integer)
                    p.ABONADO = CType(unaFila.Item(4), Integer)
                    p.COMENTARIO = CType(unaFila.Item(5), String)
                    p.COPIA = CType(unaFila.Item(6), String)
                    p.PARASUBIR = CType(unaFila.Item(7), Integer)
                    p.SUBIDO = CType(unaFila.Item(8), Integer)
                    p.FECHA = CType(unaFila.Item(9), String)
                    p.CONTROL = CType(unaFila.Item(10), Integer)
                    Lista.Add(p)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listarsinmarcarcalidad() As ArrayList
        Dim sql As String = "SELECT id, ficha, tipo, creado, abonado, ifnull(comentario,''),ifnull(copia,''), parasubir, subido, fecha, control FROM preinformes WHERE creado =0 AND tipo = 10"
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim p As New dPreinformes
                    p.ID = CType(unaFila.Item(0), Long)
                    p.FICHA = CType(unaFila.Item(1), Long)
                    p.TIPO = CType(unaFila.Item(2), Integer)
                    p.CREADO = CType(unaFila.Item(3), Integer)
                    p.ABONADO = CType(unaFila.Item(4), Integer)
                    p.COMENTARIO = CType(unaFila.Item(5), String)
                    p.COPIA = CType(unaFila.Item(6), String)
                    p.PARASUBIR = CType(unaFila.Item(7), Integer)
                    p.SUBIDO = CType(unaFila.Item(8), Integer)
                    p.FECHA = CType(unaFila.Item(9), String)
                    p.CONTROL = CType(unaFila.Item(10), Integer)
                    Lista.Add(p)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listarsinmarcarcontrol() As ArrayList
        Dim sql As String = "SELECT id, ficha, tipo, creado, abonado, ifnull(comentario,''),ifnull(copia,''), parasubir, subido, fecha, control FROM preinformes WHERE creado =0 AND tipo = 1"
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim p As New dPreinformes
                    p.ID = CType(unaFila.Item(0), Long)
                    p.FICHA = CType(unaFila.Item(1), Long)
                    p.TIPO = CType(unaFila.Item(2), Integer)
                    p.CREADO = CType(unaFila.Item(3), Integer)
                    p.ABONADO = CType(unaFila.Item(4), Integer)
                    p.COMENTARIO = CType(unaFila.Item(5), String)
                    p.COPIA = CType(unaFila.Item(6), String)
                    p.PARASUBIR = CType(unaFila.Item(7), Integer)
                    p.SUBIDO = CType(unaFila.Item(8), Integer)
                    p.FECHA = CType(unaFila.Item(9), String)
                    p.CONTROL = CType(unaFila.Item(10), Integer)
                    Lista.Add(p)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listarsincontrolclechero(ByVal desde As String, ByVal hasta As String) As ArrayList
        Dim sql As String = "SELECT id, ficha, tipo, creado, abonado, ifnull(comentario,''),ifnull(copia,''), parasubir, subido, fecha, control FROM preinformes WHERE tipo = 1 AND parasubir = 1 AND control = 0 AND fecha BETWEEN '" & desde & "' AND '" & hasta & "' ORDER BY RAND() LIMIT 5"
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim p As New dPreinformes
                    p.ID = CType(unaFila.Item(0), Long)
                    p.FICHA = CType(unaFila.Item(1), Long)
                    p.TIPO = CType(unaFila.Item(2), Integer)
                    p.CREADO = CType(unaFila.Item(3), Integer)
                    p.ABONADO = CType(unaFila.Item(4), Integer)
                    p.COMENTARIO = CType(unaFila.Item(5), String)
                    p.COPIA = CType(unaFila.Item(6), String)
                    p.PARASUBIR = CType(unaFila.Item(7), Integer)
                    p.SUBIDO = CType(unaFila.Item(8), Integer)
                    p.FECHA = CType(unaFila.Item(9), String)
                    p.CONTROL = CType(unaFila.Item(10), Integer)
                    Lista.Add(p)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listarsincontrolcalidad(ByVal desde As String, ByVal hasta As String) As ArrayList
        Dim sql As String = "SELECT id, ficha, tipo, creado, abonado, ifnull(comentario,''),ifnull(copia,''), parasubir, subido, fecha, control FROM preinformes WHERE tipo = 10 AND parasubir = 1 AND control = 0 AND fecha BETWEEN '" & desde & "' AND '" & hasta & "' ORDER BY RAND() LIMIT 5"
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim p As New dPreinformes
                    p.ID = CType(unaFila.Item(0), Long)
                    p.FICHA = CType(unaFila.Item(1), Long)
                    p.TIPO = CType(unaFila.Item(2), Integer)
                    p.CREADO = CType(unaFila.Item(3), Integer)
                    p.ABONADO = CType(unaFila.Item(4), Integer)
                    p.COMENTARIO = CType(unaFila.Item(5), String)
                    p.COPIA = CType(unaFila.Item(6), String)
                    p.PARASUBIR = CType(unaFila.Item(7), Integer)
                    p.SUBIDO = CType(unaFila.Item(8), Integer)
                    p.FECHA = CType(unaFila.Item(9), String)
                    p.CONTROL = CType(unaFila.Item(10), Integer)
                    Lista.Add(p)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listarsincontrolnutricion(ByVal desde As String, ByVal hasta As String) As ArrayList
        Dim sql As String = "SELECT id, ficha, tipo, creado, abonado, ifnull(comentario,''),ifnull(copia,''), parasubir, subido, fecha, control FROM preinformes WHERE tipo = 13 AND parasubir = 1 AND control = 0 AND fecha BETWEEN '" & desde & "' AND '" & hasta & "' ORDER BY RAND() LIMIT 5"
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim p As New dPreinformes
                    p.ID = CType(unaFila.Item(0), Long)
                    p.FICHA = CType(unaFila.Item(1), Long)
                    p.TIPO = CType(unaFila.Item(2), Integer)
                    p.CREADO = CType(unaFila.Item(3), Integer)
                    p.ABONADO = CType(unaFila.Item(4), Integer)
                    p.COMENTARIO = CType(unaFila.Item(5), String)
                    p.COPIA = CType(unaFila.Item(6), String)
                    p.PARASUBIR = CType(unaFila.Item(7), Integer)
                    p.SUBIDO = CType(unaFila.Item(8), Integer)
                    p.FECHA = CType(unaFila.Item(9), String)
                    p.CONTROL = CType(unaFila.Item(10), Integer)
                    Lista.Add(p)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listarsincontrolsuelos(ByVal desde As String, ByVal hasta As String) As ArrayList
        Dim sql As String = "SELECT id, ficha, tipo, creado, abonado, ifnull(comentario,''),ifnull(copia,''), parasubir, subido, fecha, control FROM preinformes WHERE tipo = 14 AND parasubir = 1 AND control = 0 AND fecha BETWEEN '" & desde & "' AND '" & hasta & "' ORDER BY RAND() LIMIT 5"
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim p As New dPreinformes
                    p.ID = CType(unaFila.Item(0), Long)
                    p.FICHA = CType(unaFila.Item(1), Long)
                    p.TIPO = CType(unaFila.Item(2), Integer)
                    p.CREADO = CType(unaFila.Item(3), Integer)
                    p.ABONADO = CType(unaFila.Item(4), Integer)
                    p.COMENTARIO = CType(unaFila.Item(5), String)
                    p.COPIA = CType(unaFila.Item(6), String)
                    p.PARASUBIR = CType(unaFila.Item(7), Integer)
                    p.SUBIDO = CType(unaFila.Item(8), Integer)
                    p.FECHA = CType(unaFila.Item(9), String)
                    p.CONTROL = CType(unaFila.Item(10), Integer)
                    Lista.Add(p)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listarsincontrolagua(ByVal desde As String, ByVal hasta As String) As ArrayList
        Dim sql As String = "SELECT id, ficha, tipo, creado, abonado, ifnull(comentario,''),ifnull(copia,''), parasubir, subido, fecha, control FROM preinformes WHERE tipo = 3 AND parasubir = 1 AND control = 0 AND fecha BETWEEN '" & desde & "' AND '" & hasta & "' ORDER BY RAND() LIMIT 5"
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim p As New dPreinformes
                    p.ID = CType(unaFila.Item(0), Long)
                    p.FICHA = CType(unaFila.Item(1), Long)
                    p.TIPO = CType(unaFila.Item(2), Integer)
                    p.CREADO = CType(unaFila.Item(3), Integer)
                    p.ABONADO = CType(unaFila.Item(4), Integer)
                    p.COMENTARIO = CType(unaFila.Item(5), String)
                    p.COPIA = CType(unaFila.Item(6), String)
                    p.PARASUBIR = CType(unaFila.Item(7), Integer)
                    p.SUBIDO = CType(unaFila.Item(8), Integer)
                    p.FECHA = CType(unaFila.Item(9), String)
                    p.CONTROL = CType(unaFila.Item(10), Integer)
                    Lista.Add(p)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listarsincontrolsubproductos(ByVal desde As String, ByVal hasta As String) As ArrayList
        Dim sql As String = "SELECT id, ficha, tipo, creado, abonado, ifnull(comentario,''),ifnull(copia,''), parasubir, subido, fecha, control FROM preinformes WHERE tipo = 7 AND parasubir = 1 AND control = 0 AND fecha BETWEEN '" & desde & "' AND '" & hasta & "' ORDER BY RAND() LIMIT 5"
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim p As New dPreinformes
                    p.ID = CType(unaFila.Item(0), Long)
                    p.FICHA = CType(unaFila.Item(1), Long)
                    p.TIPO = CType(unaFila.Item(2), Integer)
                    p.CREADO = CType(unaFila.Item(3), Integer)
                    p.ABONADO = CType(unaFila.Item(4), Integer)
                    p.COMENTARIO = CType(unaFila.Item(5), String)
                    p.COPIA = CType(unaFila.Item(6), String)
                    p.PARASUBIR = CType(unaFila.Item(7), Integer)
                    p.SUBIDO = CType(unaFila.Item(8), Integer)
                    p.FECHA = CType(unaFila.Item(9), String)
                    p.CONTROL = CType(unaFila.Item(10), Integer)
                    Lista.Add(p)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listarparasubir() As ArrayList
        Dim sql As String = "SELECT id, ficha, tipo, creado, abonado, ifnull(comentario,''),ifnull(copia,''), parasubir, subido, fecha, control FROM preinformes WHERE parasubir =1 AND subido = 0"
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim p As New dPreinformes
                    p.ID = CType(unaFila.Item(0), Long)
                    p.FICHA = CType(unaFila.Item(1), Long)
                    p.TIPO = CType(unaFila.Item(2), Integer)
                    p.CREADO = CType(unaFila.Item(3), Integer)
                    p.ABONADO = CType(unaFila.Item(4), Integer)
                    p.COMENTARIO = CType(unaFila.Item(5), String)
                    p.COPIA = CType(unaFila.Item(6), String)
                    p.PARASUBIR = CType(unaFila.Item(7), Integer)
                    p.SUBIDO = CType(unaFila.Item(8), Integer)
                    p.FECHA = CType(unaFila.Item(9), String)
                    p.CONTROL = CType(unaFila.Item(10), Integer)
                    Lista.Add(p)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listarparasubircontrol() As ArrayList
        Dim sql As String = "SELECT id, ficha, tipo, creado, abonado, ifnull(comentario,''),ifnull(copia,''), parasubir, subido, fecha, control FROM preinformes WHERE tipo = 1 AND parasubir =1 AND subido = 0"
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim p As New dPreinformes
                    p.ID = CType(unaFila.Item(0), Long)
                    p.FICHA = CType(unaFila.Item(1), Long)
                    p.TIPO = CType(unaFila.Item(2), Integer)
                    p.CREADO = CType(unaFila.Item(3), Integer)
                    p.ABONADO = CType(unaFila.Item(4), Integer)
                    p.COMENTARIO = CType(unaFila.Item(5), String)
                    p.COPIA = CType(unaFila.Item(6), String)
                    p.PARASUBIR = CType(unaFila.Item(7), Integer)
                    p.SUBIDO = CType(unaFila.Item(8), Integer)
                    p.FECHA = CType(unaFila.Item(9), String)
                    p.CONTROL = CType(unaFila.Item(10), Integer)
                    Lista.Add(p)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listarparasubiragua() As ArrayList
        Dim sql As String = "SELECT id, ficha, tipo, creado, abonado, ifnull(comentario,''),ifnull(copia,''), parasubir, subido, fecha, control FROM preinformes WHERE tipo = 3 AND parasubir =1 AND subido = 0"
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim p As New dPreinformes
                    p.ID = CType(unaFila.Item(0), Long)
                    p.FICHA = CType(unaFila.Item(1), Long)
                    p.TIPO = CType(unaFila.Item(2), Integer)
                    p.CREADO = CType(unaFila.Item(3), Integer)
                    p.ABONADO = CType(unaFila.Item(4), Integer)
                    p.COMENTARIO = CType(unaFila.Item(5), String)
                    p.COPIA = CType(unaFila.Item(6), String)
                    p.PARASUBIR = CType(unaFila.Item(7), Integer)
                    p.SUBIDO = CType(unaFila.Item(8), Integer)
                    p.FECHA = CType(unaFila.Item(9), String)
                    p.CONTROL = CType(unaFila.Item(10), Integer)
                    Lista.Add(p)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listarparasubirefluentes() As ArrayList
        Dim sql As String = "SELECT id, ficha, tipo, creado, abonado, ifnull(comentario,''),ifnull(copia,''), parasubir, subido, fecha, control FROM preinformes WHERE tipo = 16 AND parasubir =1 AND subido = 0"
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim p As New dPreinformes
                    p.ID = CType(unaFila.Item(0), Long)
                    p.FICHA = CType(unaFila.Item(1), Long)
                    p.TIPO = CType(unaFila.Item(2), Integer)
                    p.CREADO = CType(unaFila.Item(3), Integer)
                    p.ABONADO = CType(unaFila.Item(4), Integer)
                    p.COMENTARIO = CType(unaFila.Item(5), String)
                    p.COPIA = CType(unaFila.Item(6), String)
                    p.PARASUBIR = CType(unaFila.Item(7), Integer)
                    p.SUBIDO = CType(unaFila.Item(8), Integer)
                    p.FECHA = CType(unaFila.Item(9), String)
                    p.CONTROL = CType(unaFila.Item(10), Integer)
                    Lista.Add(p)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listarparasubiratb() As ArrayList
        Dim sql As String = "SELECT id, ficha, tipo, creado, abonado, ifnull(comentario,''),ifnull(copia,''), parasubir, subido, fecha, control FROM preinformes WHERE tipo = 4 AND parasubir =1 AND subido = 0"
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim p As New dPreinformes
                    p.ID = CType(unaFila.Item(0), Long)
                    p.FICHA = CType(unaFila.Item(1), Long)
                    p.TIPO = CType(unaFila.Item(2), Integer)
                    p.CREADO = CType(unaFila.Item(3), Integer)
                    p.ABONADO = CType(unaFila.Item(4), Integer)
                    p.COMENTARIO = CType(unaFila.Item(5), String)
                    p.COPIA = CType(unaFila.Item(6), String)
                    p.PARASUBIR = CType(unaFila.Item(7), Integer)
                    p.SUBIDO = CType(unaFila.Item(8), Integer)
                    p.FECHA = CType(unaFila.Item(9), String)
                    p.CONTROL = CType(unaFila.Item(10), Integer)
                    Lista.Add(p)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listarparasubirbact(ByVal tipo As Integer) As ArrayList
        Dim sql As String
        If tipo = 17 Then
            sql = "SELECT id, ficha, tipo, creado, abonado, ifnull(comentario,''),ifnull(copia,''), parasubir, subido, fecha, control FROM preinformes WHERE tipo = 17 AND parasubir =1 AND subido = 0"
        ElseIf tipo = 18 Then
            sql = "SELECT id, ficha, tipo, creado, abonado, ifnull(comentario,''),ifnull(copia,''), parasubir, subido, fecha, control FROM preinformes WHERE tipo = 18 AND parasubir =1 AND subido = 0"

        End If
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim p As New dPreinformes
                    p.ID = CType(unaFila.Item(0), Long)
                    p.FICHA = CType(unaFila.Item(1), Long)
                    p.TIPO = CType(unaFila.Item(2), Integer)
                    p.CREADO = CType(unaFila.Item(3), Integer)
                    p.ABONADO = CType(unaFila.Item(4), Integer)
                    p.COMENTARIO = CType(unaFila.Item(5), String)
                    p.COPIA = CType(unaFila.Item(6), String)
                    p.PARASUBIR = CType(unaFila.Item(7), Integer)
                    p.SUBIDO = CType(unaFila.Item(8), Integer)
                    p.FECHA = CType(unaFila.Item(9), String)
                    p.CONTROL = CType(unaFila.Item(10), Integer)
                    Lista.Add(p)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listarparasubirparasitologia() As ArrayList
        Dim sql As String = "SELECT id, ficha, tipo, creado, abonado, ifnull(comentario,''),ifnull(copia,''), parasubir, subido, fecha, control FROM preinformes WHERE tipo = 6 AND parasubir =1 AND subido = 0"
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim p As New dPreinformes
                    p.ID = CType(unaFila.Item(0), Long)
                    p.FICHA = CType(unaFila.Item(1), Long)
                    p.TIPO = CType(unaFila.Item(2), Integer)
                    p.CREADO = CType(unaFila.Item(3), Integer)
                    p.ABONADO = CType(unaFila.Item(4), Integer)
                    p.COMENTARIO = CType(unaFila.Item(5), String)
                    p.COPIA = CType(unaFila.Item(6), String)
                    p.PARASUBIR = CType(unaFila.Item(7), Integer)
                    p.SUBIDO = CType(unaFila.Item(8), Integer)
                    p.FECHA = CType(unaFila.Item(9), String)
                    p.CONTROL = CType(unaFila.Item(10), Integer)
                    Lista.Add(p)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listarparasubirpatologia() As ArrayList
        Dim sql As String = "SELECT id, ficha, tipo, creado, abonado, ifnull(comentario,''),ifnull(copia,''), parasubir, subido, fecha, control FROM preinformes WHERE tipo = 9 AND parasubir =1 AND subido = 0"
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim p As New dPreinformes
                    p.ID = CType(unaFila.Item(0), Long)
                    p.FICHA = CType(unaFila.Item(1), Long)
                    p.TIPO = CType(unaFila.Item(2), Integer)
                    p.CREADO = CType(unaFila.Item(3), Integer)
                    p.ABONADO = CType(unaFila.Item(4), Integer)
                    p.COMENTARIO = CType(unaFila.Item(5), String)
                    p.COPIA = CType(unaFila.Item(6), String)
                    p.PARASUBIR = CType(unaFila.Item(7), Integer)
                    p.SUBIDO = CType(unaFila.Item(8), Integer)
                    p.FECHA = CType(unaFila.Item(9), String)
                    p.CONTROL = CType(unaFila.Item(10), Integer)
                    Lista.Add(p)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listarparasubirtoxicologia() As ArrayList
        Dim sql As String = "SELECT id, ficha, tipo, creado, abonado, ifnull(comentario,''),ifnull(copia,''), parasubir, subido, fecha, control FROM preinformes WHERE tipo = 20 AND parasubir =1 AND subido = 0"
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim p As New dPreinformes
                    p.ID = CType(unaFila.Item(0), Long)
                    p.FICHA = CType(unaFila.Item(1), Long)
                    p.TIPO = CType(unaFila.Item(2), Integer)
                    p.CREADO = CType(unaFila.Item(3), Integer)
                    p.ABONADO = CType(unaFila.Item(4), Integer)
                    p.COMENTARIO = CType(unaFila.Item(5), String)
                    p.COPIA = CType(unaFila.Item(6), String)
                    p.PARASUBIR = CType(unaFila.Item(7), Integer)
                    p.SUBIDO = CType(unaFila.Item(8), Integer)
                    p.FECHA = CType(unaFila.Item(9), String)
                    p.CONTROL = CType(unaFila.Item(10), Integer)
                    Lista.Add(p)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listarparasubirambiental() As ArrayList
        Dim sql As String = "SELECT id, ficha, tipo, creado, abonado, ifnull(comentario,''),ifnull(copia,''), parasubir, subido, fecha, control FROM preinformes WHERE tipo = 11 AND parasubir =1 AND subido = 0"
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim p As New dPreinformes
                    p.ID = CType(unaFila.Item(0), Long)
                    p.FICHA = CType(unaFila.Item(1), Long)
                    p.TIPO = CType(unaFila.Item(2), Integer)
                    p.CREADO = CType(unaFila.Item(3), Integer)
                    p.ABONADO = CType(unaFila.Item(4), Integer)
                    p.COMENTARIO = CType(unaFila.Item(5), String)
                    p.COPIA = CType(unaFila.Item(6), String)
                    p.PARASUBIR = CType(unaFila.Item(7), Integer)
                    p.SUBIDO = CType(unaFila.Item(8), Integer)
                    p.FECHA = CType(unaFila.Item(9), String)
                    p.CONTROL = CType(unaFila.Item(10), Integer)
                    Lista.Add(p)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listarparasubirserologia() As ArrayList
        Dim sql As String = "SELECT id, ficha, tipo, creado, abonado, ifnull(comentario,''),ifnull(copia,''), parasubir, subido, fecha, control FROM preinformes WHERE tipo = 8 AND parasubir =1 AND subido = 0"
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim p As New dPreinformes
                    p.ID = CType(unaFila.Item(0), Long)
                    p.FICHA = CType(unaFila.Item(1), Long)
                    p.TIPO = CType(unaFila.Item(2), Integer)
                    p.CREADO = CType(unaFila.Item(3), Integer)
                    p.ABONADO = CType(unaFila.Item(4), Integer)
                    p.COMENTARIO = CType(unaFila.Item(5), String)
                    p.COPIA = CType(unaFila.Item(6), String)
                    p.PARASUBIR = CType(unaFila.Item(7), Integer)
                    p.SUBIDO = CType(unaFila.Item(8), Integer)
                    p.FECHA = CType(unaFila.Item(9), String)
                    p.CONTROL = CType(unaFila.Item(10), Integer)
                    Lista.Add(p)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listarparasubirsubproductos() As ArrayList
        Dim sql As String = "SELECT id, ficha, tipo, creado, abonado, ifnull(comentario,''),ifnull(copia,''), parasubir, subido, fecha, control FROM preinformes WHERE tipo = 7 AND parasubir =1 AND subido = 0"
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim p As New dPreinformes
                    p.ID = CType(unaFila.Item(0), Long)
                    p.FICHA = CType(unaFila.Item(1), Long)
                    p.TIPO = CType(unaFila.Item(2), Integer)
                    p.CREADO = CType(unaFila.Item(3), Integer)
                    p.ABONADO = CType(unaFila.Item(4), Integer)
                    p.COMENTARIO = CType(unaFila.Item(5), String)
                    p.COPIA = CType(unaFila.Item(6), String)
                    p.PARASUBIR = CType(unaFila.Item(7), Integer)
                    p.SUBIDO = CType(unaFila.Item(8), Integer)
                    p.FECHA = CType(unaFila.Item(9), String)
                    p.CONTROL = CType(unaFila.Item(10), Integer)
                    Lista.Add(p)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listarparasubircalidad() As ArrayList
        Dim sql As String = "SELECT id, ficha, tipo, creado, abonado, ifnull(comentario,''),ifnull(copia,''), parasubir, subido, fecha, control FROM preinformes WHERE tipo = 10 AND parasubir =1 AND subido = 0"
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim p As New dPreinformes
                    p.ID = CType(unaFila.Item(0), Long)
                    p.FICHA = CType(unaFila.Item(1), Long)
                    p.TIPO = CType(unaFila.Item(2), Integer)
                    p.CREADO = CType(unaFila.Item(3), Integer)
                    p.ABONADO = CType(unaFila.Item(4), Integer)
                    p.COMENTARIO = CType(unaFila.Item(5), String)
                    p.COPIA = CType(unaFila.Item(6), String)
                    p.PARASUBIR = CType(unaFila.Item(7), Integer)
                    p.SUBIDO = CType(unaFila.Item(8), Integer)
                    p.FECHA = CType(unaFila.Item(9), String)
                    p.CONTROL = CType(unaFila.Item(10), Integer)
                    Lista.Add(p)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listarparasubirnutricion() As ArrayList
        Dim sql As String = "SELECT id, ficha, tipo, creado, abonado, ifnull(comentario,''),ifnull(copia,''), parasubir, subido, fecha, control FROM preinformes WHERE tipo = 13 AND parasubir =1 AND subido = 0"
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim p As New dPreinformes
                    p.ID = CType(unaFila.Item(0), Long)
                    p.FICHA = CType(unaFila.Item(1), Long)
                    p.TIPO = CType(unaFila.Item(2), Integer)
                    p.CREADO = CType(unaFila.Item(3), Integer)
                    p.ABONADO = CType(unaFila.Item(4), Integer)
                    p.COMENTARIO = CType(unaFila.Item(5), String)
                    p.COPIA = CType(unaFila.Item(6), String)
                    p.PARASUBIR = CType(unaFila.Item(7), Integer)
                    p.SUBIDO = CType(unaFila.Item(8), Integer)
                    p.FECHA = CType(unaFila.Item(9), String)
                    p.CONTROL = CType(unaFila.Item(10), Integer)
                    Lista.Add(p)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listarparasubirsuelos() As ArrayList
        Dim sql As String = "SELECT id, ficha, tipo, creado, abonado, ifnull(comentario,''),ifnull(copia,''), parasubir, subido, fecha, control FROM preinformes WHERE tipo = 14 AND parasubir =1 AND subido = 0"
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim p As New dPreinformes
                    p.ID = CType(unaFila.Item(0), Long)
                    p.FICHA = CType(unaFila.Item(1), Long)
                    p.TIPO = CType(unaFila.Item(2), Integer)
                    p.CREADO = CType(unaFila.Item(3), Integer)
                    p.ABONADO = CType(unaFila.Item(4), Integer)
                    p.COMENTARIO = CType(unaFila.Item(5), String)
                    p.COPIA = CType(unaFila.Item(6), String)
                    p.PARASUBIR = CType(unaFila.Item(7), Integer)
                    p.SUBIDO = CType(unaFila.Item(8), Integer)
                    p.FECHA = CType(unaFila.Item(9), String)
                    p.CONTROL = CType(unaFila.Item(10), Integer)
                    Lista.Add(p)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listarparasubirfoliar() As ArrayList
        Dim sql As String = "SELECT id, ficha, tipo, creado, abonado, ifnull(comentario,''),ifnull(copia,''), parasubir, subido, fecha, control FROM preinformes WHERE tipo = 19 AND parasubir =1 AND subido = 0"
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim p As New dPreinformes
                    p.ID = CType(unaFila.Item(0), Long)
                    p.FICHA = CType(unaFila.Item(1), Long)
                    p.TIPO = CType(unaFila.Item(2), Integer)
                    p.CREADO = CType(unaFila.Item(3), Integer)
                    p.ABONADO = CType(unaFila.Item(4), Integer)
                    p.COMENTARIO = CType(unaFila.Item(5), String)
                    p.COPIA = CType(unaFila.Item(6), String)
                    p.PARASUBIR = CType(unaFila.Item(7), Integer)
                    p.SUBIDO = CType(unaFila.Item(8), Integer)
                    p.FECHA = CType(unaFila.Item(9), String)
                    p.CONTROL = CType(unaFila.Item(10), Integer)
                    Lista.Add(p)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listarparasubirbrucelosis() As ArrayList
        Dim sql As String = "SELECT id, ficha, tipo, creado, abonado, ifnull(comentario,''),ifnull(copia,''), parasubir, subido, fecha, control FROM preinformes WHERE tipo = 15 AND parasubir =1 AND subido = 0"
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim p As New dPreinformes
                    p.ID = CType(unaFila.Item(0), Long)
                    p.FICHA = CType(unaFila.Item(1), Long)
                    p.TIPO = CType(unaFila.Item(2), Integer)
                    p.CREADO = CType(unaFila.Item(3), Integer)
                    p.ABONADO = CType(unaFila.Item(4), Integer)
                    p.COMENTARIO = CType(unaFila.Item(5), String)
                    p.COPIA = CType(unaFila.Item(6), String)
                    p.PARASUBIR = CType(unaFila.Item(7), Integer)
                    p.SUBIDO = CType(unaFila.Item(8), Integer)
                    p.FECHA = CType(unaFila.Item(9), String)
                    p.CONTROL = CType(unaFila.Item(10), Integer)
                    Lista.Add(p)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listarparasubirmicro() As ArrayList
        Dim sql As String = "SELECT id, ficha, tipo, creado, abonado, ifnull(comentario,''),ifnull(copia,''), parasubir, subido, fecha, control FROM preinformes WHERE tipo = 3 OR tipo = 7 OR tipo = 10 AND parasubir =1 AND subido = 0"
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim p As New dPreinformes
                    p.ID = CType(unaFila.Item(0), Long)
                    p.FICHA = CType(unaFila.Item(1), Long)
                    p.TIPO = CType(unaFila.Item(2), Integer)
                    p.CREADO = CType(unaFila.Item(3), Integer)
                    p.ABONADO = CType(unaFila.Item(4), Integer)
                    p.COMENTARIO = CType(unaFila.Item(5), String)
                    p.COPIA = CType(unaFila.Item(6), String)
                    p.PARASUBIR = CType(unaFila.Item(7), Integer)
                    p.SUBIDO = CType(unaFila.Item(8), Integer)
                    p.FECHA = CType(unaFila.Item(9), String)
                    p.CONTROL = CType(unaFila.Item(10), Integer)
                    Lista.Add(p)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listarsubidas() As ArrayList
        Dim sql As String = "SELECT id, ficha, tipo, creado, abonado, ifnull(comentario,''),ifnull(copia,''), parasubir, subido, fecha, control FROM preinformes WHERE parasubir =1 AND subido = 1 ORDER BY FECHA DESC LIMIT 100"
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim p As New dPreinformes
                    p.ID = CType(unaFila.Item(0), Long)
                    p.FICHA = CType(unaFila.Item(1), Long)
                    p.TIPO = CType(unaFila.Item(2), Integer)
                    p.CREADO = CType(unaFila.Item(3), Integer)
                    p.ABONADO = CType(unaFila.Item(4), Integer)
                    p.COMENTARIO = CType(unaFila.Item(5), String)
                    p.COPIA = CType(unaFila.Item(6), String)
                    p.PARASUBIR = CType(unaFila.Item(7), Integer)
                    p.SUBIDO = CType(unaFila.Item(8), Integer)
                    p.FECHA = CType(unaFila.Item(9), String)
                    p.CONTROL = CType(unaFila.Item(10), Integer)
                    Lista.Add(p)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
End Class
