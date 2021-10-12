﻿Public Class pDepartamento
    Inherits Conectoras.ConexionMySQL
    Public Function guardar(ByVal o As Object) As Boolean
        Dim obj As dDepartamento = CType(o, dDepartamento)
        Dim sql As String = "INSERT INTO departamento (id, nombre) VALUES (" & obj.ID & ", '" & obj.NOMBRE & "')"

        Dim lista As New ArrayList
        lista.Add(sql)

     

        Return EjecutarTransaccion(lista)
    End Function
    Public Function modificar(ByVal o As Object) As Boolean
        Dim obj As dDepartamento = CType(o, dDepartamento)
        Dim sql As String = "UPDATE departamento SET nombre = '" & obj.NOMBRE & "' WHERE id = " & obj.ID & ""

        Dim lista As New ArrayList
        lista.Add(sql)

    

        Return EjecutarTransaccion(lista)
    End Function
    Public Function eliminar(ByVal o As Object) As Boolean
        Dim obj As dDepartamento = CType(o, dDepartamento)
        Dim sql As String = "DELETE FROM departamento WHERE id = " & obj.ID & ""

        Dim lista As New ArrayList
        lista.Add(sql)

     

        Return EjecutarTransaccion(lista)
    End Function
    Public Function buscar(ByVal o As Object) As dDepartamento
        Dim obj As dDepartamento = CType(o, dDepartamento)
        Dim l As New dDepartamento
        Try
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL("SELECT id, nombre FROM departamento WHERE id = " & obj.ID & "")

            If Ds.Tables(0).Rows.Count > 0 Then
                Dim unaFila As DataRow
                unaFila = Ds.Tables(0).Rows(0)
                l.ID = CType(unaFila.Item(0), Integer)
                l.NOMBRE = CType(unaFila.Item(1), String)
                Return l
            End If
            Return Nothing
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listar() As ArrayList
        Dim sql As String = "SELECT id, nombre FROM departamento"
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim l As New dDepartamento
                    l.ID = CType(unaFila.Item(0), Integer)
                    l.NOMBRE = CType(unaFila.Item(1), String)
                    Lista.Add(l)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
End Class
