﻿Public Class pParasitologiaSolicitud
    Inherits Conectoras.ConexionMySQL
    Public Function guardar(ByVal o As Object) As Boolean
        Dim obj As dParasitologiaSolicitud = CType(o, dParasitologiaSolicitud)
        Dim sql As String = "INSERT INTO parasitologia_solicitud (id, ficha, gastrointestinales, fasciola, coccidias) VALUES (" & obj.ID & ", " & obj.FICHA & "," & obj.GASTROINTESTINALES & ", " & obj.FASCIOLA & ", " & obj.COCCIDIAS & ")"

        Dim lista As New ArrayList
        lista.Add(sql)


        Return EjecutarTransaccion(lista)
    End Function
    Public Function modificar(ByVal o As Object) As Boolean
        Dim obj As dParasitologiaSolicitud = CType(o, dParasitologiaSolicitud)
        Dim sql As String = "UPDATE parasitologia_solicitud SET gastrointestinales=" & obj.GASTROINTESTINALES & ",fasciola=" & obj.FASCIOLA & ",coccidias=" & obj.COCCIDIAS & " WHERE ficha = " & obj.FICHA

        Dim lista As New ArrayList
        lista.Add(sql)


        Return EjecutarTransaccion(lista)
    End Function
    Public Function eliminar(ByVal o As Object) As Boolean
        Dim obj As dParasitologiaSolicitud = CType(o, dParasitologiaSolicitud)
        Dim sql As String = "DELETE FROM parasitologia_solicitud WHERE ID = " & obj.ID

        Dim lista As New ArrayList
        lista.Add(sql)

     

        Return EjecutarTransaccion(lista)
    End Function
    Public Function buscar(ByVal o As Object) As dParasitologiaSolicitud
        Dim obj As dParasitologiaSolicitud = CType(o, dParasitologiaSolicitud)
        Dim c As New dParasitologiaSolicitud
        Try
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL("SELECT id, ficha, gastrointestinales, fasciola, coccidias FROM parasitologia_solicitud WHERE ficha = " & obj.ficha & "")

            If Ds.Tables(0).Rows.Count > 0 Then
                Dim unaFila As DataRow
                unaFila = Ds.Tables(0).Rows(0)
                c.ID = CType(unaFila.Item(0), Long)
                c.ficha = CType(unaFila.Item(1), Long)
                c.GASTROINTESTINALES = CType(unaFila.Item(2), Integer)
                c.FASCIOLA = CType(unaFila.Item(3), Integer)
                c.COCCIDIAS = CType(unaFila.Item(4), Integer)

                Return c
            End If
            Return Nothing
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Function listar() As ArrayList
        Dim sql As String = "SELECT id, ficha, gastrointestinales, fasciola, coccidias FROM parasitologia_solicitud WHERE marca = 0 order by id desc"
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim c As New dParasitologiaSolicitud
                    c.ID = CType(unaFila.Item(0), Long)
                    c.ficha = CType(unaFila.Item(1), Long)
                    c.GASTROINTESTINALES = CType(unaFila.Item(2), Integer)
                    c.FASCIOLA = CType(unaFila.Item(3), Integer)
                    c.COCCIDIAS = CType(unaFila.Item(4), Integer)
                    Lista.Add(c)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listarfichas() As ArrayList
        Dim sql As String = "SELECT DISTINCT ficha FROM parasitologia_solicitud WHERE marca = 0 order by ficha asc"
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim c As New dParasitologiaSolicitud
                    c.ficha = CType(unaFila.Item(0), Long)
                    Lista.Add(c)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    Public Function listarporid(ByVal texto As Long) As ArrayList
        Dim sql As String = ("SELECT id, ficha, gastrointestinales, fasciola, coccidias FROM parasitologia_solicitud where ficha = " & texto)
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim c As New dParasitologiaSolicitud
                    c.ID = CType(unaFila.Item(0), Long)
                    c.ficha = CType(unaFila.Item(1), Long)
                    c.GASTROINTESTINALES = CType(unaFila.Item(2), Integer)
                    c.FASCIOLA = CType(unaFila.Item(3), Integer)
                    c.COCCIDIAS = CType(unaFila.Item(4), Integer)
                    Lista.Add(c)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function


    Public Function listarporsolicitud(ByVal texto As Long) As ArrayList
        Dim sql As String = ("SELECT id, ficha, gastrointestinales, fasciola, coccidias FROM parasitologia_solicitud where ficha = " & texto)
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim c As New dParasitologiaSolicitud
                    c.ID = CType(unaFila.Item(0), Long)
                    c.ficha = CType(unaFila.Item(1), Long)
                    c.GASTROINTESTINALES = CType(unaFila.Item(2), Integer)
                    c.FASCIOLA = CType(unaFila.Item(3), Integer)
                    c.COCCIDIAS = CType(unaFila.Item(4), Integer)
                    Lista.Add(c)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listarporsolicitud2(ByVal texto As Long) As ArrayList
        Dim sql As String = ("SELECT id, ficha, gastrointestinales, fasciola, coccidias FROM parasitologia_solicitud where marca = 1 and ficha = " & texto)
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim c As New dParasitologiaSolicitud
                    c.ID = CType(unaFila.Item(0), Long)
                    c.ficha = CType(unaFila.Item(1), Long)
                    c.GASTROINTESTINALES = CType(unaFila.Item(2), Integer)
                    c.FASCIOLA = CType(unaFila.Item(3), Integer)
                    c.COCCIDIAS = CType(unaFila.Item(4), Integer)
                    Lista.Add(c)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listarporfecha(ByVal fechadesde As String, ByVal fechahasta As String) As ArrayList
        Dim sql As String = ("SELECT id, ficha, gastrointestinales, fasciola, coccidias FROM parasitologia_solicitud where fechaingreso BETWEEN '" & fechadesde & "' And '" & fechahasta & "'")
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim c As New dParasitologiaSolicitud
                    c.ID = CType(unaFila.Item(0), Long)
                    c.ficha = CType(unaFila.Item(1), Long)
                    c.GASTROINTESTINALES = CType(unaFila.Item(2), Integer)
                    c.FASCIOLA = CType(unaFila.Item(3), Integer)
                    c.COCCIDIAS = CType(unaFila.Item(4), Integer)
                    Lista.Add(c)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
End Class
