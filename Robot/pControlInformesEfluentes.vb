﻿Public Class pControlInformesEfluentes
    Inherits Conectoras.ConexionMySQL
    Public Function guardar(ByVal o As Object) As Boolean
        Dim obj As dControlInformesEfluentes = CType(o, dControlInformesEfluentes)
        Dim sql As String = "INSERT INTO controlinformesefluentes (id, fechacontrol, ficha, fecha, tipo, resultado, coincide, opcionmejora, noconformidad, observaciones,  controlador, controlado) VALUES (" & obj.ID & ",'" & obj.FECHACONTROL & "'," & obj.FICHA & ", '" & obj.FECHA & "'," & obj.TIPO & ", " & obj.RESULTADO & "," & obj.COINCIDE & "," & obj.OM & "," & obj.NC & ",'" & obj.OBSERVACIONES & "'," & obj.CONTROLADOR & "," & obj.CONTROLADO & ")"

        Dim lista As New ArrayList
        lista.Add(sql)


        Return EjecutarTransaccion(lista)
    End Function
    Public Function modificar(ByVal o As Object) As Boolean
        Dim obj As dControlInformesEfluentes = CType(o, dControlInformesEfluentes)
        Dim sql As String = "UPDATE controlinformesefluentes SET fechacontrol ='" & obj.FECHACONTROL & "',ficha=" & obj.FICHA & ",fecha ='" & obj.FECHA & "',tipo =" & obj.TIPO & ",resultado =" & obj.RESULTADO & ", coincide=" & obj.COINCIDE & ",opcionmejora=" & obj.OM & ",noconformidad=" & obj.NC & ",  observaciones ='" & obj.OBSERVACIONES & "', controlador =" & obj.CONTROLADOR & ",  controlado =" & obj.CONTROLADO & " WHERE id = " & obj.ID & ""

        Dim lista As New ArrayList
        lista.Add(sql)


        Return EjecutarTransaccion(lista)
    End Function
    Public Function eliminar(ByVal o As Object) As Boolean
        Dim obj As dControlInformesEfluentes = CType(o, dControlInformesEfluentes)
        Dim sql As String = "DELETE FROM controlinformesefluentes WHERE id = " & obj.ID & ""
        Dim lista As New ArrayList
        lista.Add(sql)


        Return EjecutarTransaccion(lista)
    End Function
    Public Function buscar(ByVal o As Object) As dControlInformesEfluentes
        Dim obj As dControlInformesEfluentes = CType(o, dControlInformesEfluentes)
        Dim c As New dControlInformesEfluentes
        Try
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL("SELECT id, fechacontrol, ficha, fecha, tipo, resultado, coincide,opcionmejora, noconformidad, observaciones,  controlador, controlado FROM controlinformesefluentes WHERE id = " & obj.ID & "")

            If Ds.Tables(0).Rows.Count > 0 Then
                Dim unaFila As DataRow
                unaFila = Ds.Tables(0).Rows(0)
                c.ID = CType(unaFila.Item(0), Long)
                c.FECHACONTROL = CType(unaFila.Item(1), String)
                c.FICHA = CType(unaFila.Item(2), Long)
                c.FECHA = CType(unaFila.Item(3), String)
                c.TIPO = CType(unaFila.Item(4), Integer)
                c.RESULTADO = CType(unaFila.Item(5), Integer)
                c.COINCIDE = CType(unaFila.Item(6), Integer)
                c.OM = CType(unaFila.Item(7), Integer)
                c.NC = CType(unaFila.Item(8), Integer)
                c.OBSERVACIONES = CType(unaFila.Item(9), String)
                c.CONTROLADOR = CType(unaFila.Item(10), Integer)
                c.CONTROLADO = CType(unaFila.Item(11), Integer)
                Return c
            End If
            Return Nothing
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function buscarxficha(ByVal o As Object) As dControlInformesEfluentes
        Dim obj As dControlInformesEfluentes = CType(o, dControlInformesEfluentes)
        Dim c As New dControlInformesEfluentes
        Try
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL("SELECT id, fechacontrol, ficha, fecha, tipo, resultado, coincide,opcionmejora, noconformidad, observaciones,  controlador, controlado FROM controlinformesefluentes WHERE ficha = " & obj.FICHA & "")

            If Ds.Tables(0).Rows.Count > 0 Then
                Dim unaFila As DataRow
                unaFila = Ds.Tables(0).Rows(0)
                c.ID = CType(unaFila.Item(0), Long)
                c.FECHACONTROL = CType(unaFila.Item(1), String)
                c.FICHA = CType(unaFila.Item(2), Long)
                c.FECHA = CType(unaFila.Item(3), String)
                c.TIPO = CType(unaFila.Item(4), Integer)
                c.RESULTADO = CType(unaFila.Item(5), Integer)
                c.COINCIDE = CType(unaFila.Item(6), Integer)
                c.OM = CType(unaFila.Item(7), Integer)
                c.NC = CType(unaFila.Item(8), Integer)
                c.OBSERVACIONES = CType(unaFila.Item(9), String)
                c.CONTROLADOR = CType(unaFila.Item(10), Integer)
                c.CONTROLADO = CType(unaFila.Item(11), Integer)
                Return c
            End If
            Return Nothing
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listar() As ArrayList
        Dim sql As String = ("select id, fechacontrol, ficha, fecha, tipo, resultado, coincide,opcionmejora, noconformidad, observaciones,  controlador, controlado FROM controlinformesefluentes WHERE controlado = 0 order by fecha asc")
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim c As New dControlInformesEfluentes
                    c.ID = CType(unaFila.Item(0), Long)
                    c.FECHACONTROL = CType(unaFila.Item(1), String)
                    c.FICHA = CType(unaFila.Item(2), Long)
                    c.FECHA = CType(unaFila.Item(3), String)
                    c.TIPO = CType(unaFila.Item(4), Integer)
                    c.RESULTADO = CType(unaFila.Item(5), Integer)
                    c.COINCIDE = CType(unaFila.Item(6), Integer)
                    c.OM = CType(unaFila.Item(7), Integer)
                    c.NC = CType(unaFila.Item(8), Integer)
                    c.OBSERVACIONES = CType(unaFila.Item(9), String)
                    c.CONTROLADOR = CType(unaFila.Item(10), Integer)
                    c.CONTROLADO = CType(unaFila.Item(11), Integer)
                    Lista.Add(c)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listarxfecha(ByVal desde As String, ByVal hasta As String) As ArrayList
        Dim sql As String = ("select id, fechacontrol, ficha, fecha, tipo, resultado, coincide,opcionmejora, noconformidad, observaciones,  controlador, controlado FROM controlinformesefluentes WHERE fecha >='" & desde & "' and fecha <='" & hasta & "' AND controlado = 1 order by tipo asc")
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim c As New dControlInformesEfluentes
                    c.ID = CType(unaFila.Item(0), Long)
                    c.FECHACONTROL = CType(unaFila.Item(1), String)
                    c.FICHA = CType(unaFila.Item(2), Long)
                    c.FECHA = CType(unaFila.Item(3), String)
                    c.TIPO = CType(unaFila.Item(4), Integer)
                    c.RESULTADO = CType(unaFila.Item(5), Integer)
                    c.COINCIDE = CType(unaFila.Item(6), Integer)
                    c.OM = CType(unaFila.Item(7), Integer)
                    c.NC = CType(unaFila.Item(8), Integer)
                    c.OBSERVACIONES = CType(unaFila.Item(9), String)
                    c.CONTROLADOR = CType(unaFila.Item(10), Integer)
                    c.CONTROLADO = CType(unaFila.Item(11), Integer)
                    Lista.Add(c)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listarxtipoxfecha(ByVal tipo As Integer, ByVal desde As String, ByVal hasta As String) As ArrayList
        Dim sql As String = ("select id, fechacontrol, ficha, fecha, tipo, resultado, coincide,opcionmejora, noconformidad, observaciones,  controlador, controlado FROM controlinformesefluentes WHERE tipo= " & tipo & " AND fecha >='" & desde & "' and fecha <='" & hasta & "'")
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim c As New dControlInformesFQ
                    c.ID = CType(unaFila.Item(0), Long)
                    c.FECHACONTROL = CType(unaFila.Item(1), String)
                    c.FICHA = CType(unaFila.Item(2), Long)
                    c.FECHA = CType(unaFila.Item(3), String)
                    c.TIPO = CType(unaFila.Item(4), Integer)
                    c.RESULTADO = CType(unaFila.Item(5), Integer)
                    c.COINCIDE = CType(unaFila.Item(6), Integer)
                    c.OM = CType(unaFila.Item(7), Integer)
                    c.NC = CType(unaFila.Item(8), Integer)
                    c.OBSERVACIONES = CType(unaFila.Item(9), String)
                    c.CONTROLADOR = CType(unaFila.Item(10), Integer)
                    c.CONTROLADO = CType(unaFila.Item(11), Integer)
                    Lista.Add(c)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listarxfechanc(ByVal desde As String, ByVal hasta As String) As ArrayList
        Dim sql As String = ("select id, fechacontrol, ficha, fecha, tipo, resultado, coincide,opcionmejora, noconformidad, observaciones,  controlador, controlado FROM controlinformesefluentes WHERE fecha >='" & desde & "' and fecha <='" & hasta & "' AND controlado = 1 AND noconformidad=1 order by tipo asc")
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim c As New dControlInformesEfluentes
                    c.ID = CType(unaFila.Item(0), Long)
                    c.FECHACONTROL = CType(unaFila.Item(1), String)
                    c.FICHA = CType(unaFila.Item(2), Long)
                    c.FECHA = CType(unaFila.Item(3), String)
                    c.TIPO = CType(unaFila.Item(4), Integer)
                    c.RESULTADO = CType(unaFila.Item(5), Integer)
                    c.COINCIDE = CType(unaFila.Item(6), Integer)
                    c.OM = CType(unaFila.Item(7), Integer)
                    c.NC = CType(unaFila.Item(8), Integer)
                    c.OBSERVACIONES = CType(unaFila.Item(9), String)
                    c.CONTROLADOR = CType(unaFila.Item(10), Integer)
                    c.CONTROLADO = CType(unaFila.Item(11), Integer)
                    Lista.Add(c)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listarxfechaom(ByVal desde As String, ByVal hasta As String) As ArrayList
        Dim sql As String = ("select id, fechacontrol, ficha, fecha, tipo, resultado, coincide,opcionmejora, noconformidad, observaciones,  controlador, controlado FROM controlinformesefluentes WHERE fecha >='" & desde & "' and fecha <='" & hasta & "' AND controlado = 1 AND opcionmejora=1 order by tipo asc")
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim c As New dControlInformesEfluentes
                    c.ID = CType(unaFila.Item(0), Long)
                    c.FECHACONTROL = CType(unaFila.Item(1), String)
                    c.FICHA = CType(unaFila.Item(2), Long)
                    c.FECHA = CType(unaFila.Item(3), String)
                    c.TIPO = CType(unaFila.Item(4), Integer)
                    c.RESULTADO = CType(unaFila.Item(5), Integer)
                    c.COINCIDE = CType(unaFila.Item(6), Integer)
                    c.OM = CType(unaFila.Item(7), Integer)
                    c.NC = CType(unaFila.Item(8), Integer)
                    c.OBSERVACIONES = CType(unaFila.Item(9), String)
                    c.CONTROLADOR = CType(unaFila.Item(10), Integer)
                    c.CONTROLADO = CType(unaFila.Item(11), Integer)
                    Lista.Add(c)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function marcarresultado(ByVal o As Object) As Boolean
        Dim obj As dControlInformesEfluentes = CType(o, dControlInformesEfluentes)
        Dim sql As String = "UPDATE controlinformesefluentes SET resultado=1 WHERE resultado=0 and ID = " & obj.ID & ""

        Dim lista As New ArrayList
        lista.Add(sql)

        Return EjecutarTransaccion(lista)
    End Function
    Public Function desmarcarresultado(ByVal o As Object) As Boolean
        Dim obj As dControlInformesEfluentes = CType(o, dControlInformesEfluentes)
        Dim sql As String = "UPDATE controlinformesefluentes SET resultado=0 WHERE resultado=1 and ID = " & obj.ID & ""

        Dim lista As New ArrayList
        lista.Add(sql)

        Return EjecutarTransaccion(lista)
    End Function
    Public Function marcarcoincide(ByVal o As Object) As Boolean
        Dim obj As dControlInformesEfluentes = CType(o, dControlInformesEfluentes)
        Dim sql As String = "UPDATE controlinformesefluentes SET coincide=1 WHERE coincide=0 and ID = " & obj.ID & ""

        Dim lista As New ArrayList
        lista.Add(sql)

        Return EjecutarTransaccion(lista)
    End Function
    Public Function desmarcarcoincide(ByVal o As Object) As Boolean
        Dim obj As dControlInformesEfluentes = CType(o, dControlInformesEfluentes)
        Dim sql As String = "UPDATE controlinformesefluentes SET coincide=0 WHERE coincide=1 and ID = " & obj.ID & ""

        Dim lista As New ArrayList
        lista.Add(sql)

        Return EjecutarTransaccion(lista)
    End Function
    Public Function marcarom(ByVal o As Object) As Boolean
        Dim obj As dControlInformesEfluentes = CType(o, dControlInformesEfluentes)
        Dim sql As String = "UPDATE controlinformesefluentes SET opcionmejora=1 WHERE opcionmejora=0 and ID = " & obj.ID & ""

        Dim lista As New ArrayList
        lista.Add(sql)

        Return EjecutarTransaccion(lista)
    End Function
    Public Function desmarcarom(ByVal o As Object) As Boolean
        Dim obj As dControlInformesEfluentes = CType(o, dControlInformesEfluentes)
        Dim sql As String = "UPDATE controlinformesefluentes SET opcionmejora=0 WHERE opcionmejora=1 and ID = " & obj.ID & ""

        Dim lista As New ArrayList
        lista.Add(sql)

        Return EjecutarTransaccion(lista)
    End Function
    Public Function marcarnc(ByVal o As Object) As Boolean
        Dim obj As dControlInformesEfluentes = CType(o, dControlInformesEfluentes)
        Dim sql As String = "UPDATE controlinformesefluentes SET noconformidad=1 WHERE noconformidad=0 and ID = " & obj.ID & ""

        Dim lista As New ArrayList
        lista.Add(sql)

        Return EjecutarTransaccion(lista)
    End Function
    Public Function desmarcarnc(ByVal o As Object) As Boolean
        Dim obj As dControlInformesEfluentes = CType(o, dControlInformesEfluentes)
        Dim sql As String = "UPDATE controlinformesefluentes SET noconformidad=0 WHERE noconformidad=1 and ID = " & obj.ID & ""

        Dim lista As New ArrayList
        lista.Add(sql)

        Return EjecutarTransaccion(lista)
    End Function

    Public Function marcarcontrolada(ByVal o As Object) As Boolean
        Dim obj As dControlInformesEfluentes = CType(o, dControlInformesEfluentes)
        Dim sql As String = "UPDATE controlinformesefluentes SET controlado=1  WHERE controlado=0 and ID = " & obj.ID & ""

        Dim lista As New ArrayList
        lista.Add(sql)

        Return EjecutarTransaccion(lista)
    End Function
    Public Function guardarobservaciones(ByVal o As Object, ByVal obs As String) As Boolean
        Dim obj As dControlInformesEfluentes = CType(o, dControlInformesEfluentes)
        Dim sql As String = "UPDATE controlinformesefluentes SET observaciones='" & obs & "' WHERE ID = " & obj.ID & ""

        Dim lista As New ArrayList
        lista.Add(sql)

        Return EjecutarTransaccion(lista)
    End Function
End Class
