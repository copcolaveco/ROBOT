Public Class pTipoInforme
    Inherits Conectoras.ConexionMySQL
    Public Function guardar(ByVal o As Object) As Boolean
        Dim obj As dTipoInforme = CType(o, dTipoInforme)
        Dim sql As String = "INSERT INTO tipoinforme (id, nombre) VALUES (" & obj.ID & ", '" & obj.NOMBRE & "')"

        Dim lista As New ArrayList
        lista.Add(sql)

     

        Return EjecutarTransaccion(lista)
    End Function
    Public Function modificar(ByVal o As Object) As Boolean
        Dim obj As dTipoInforme = CType(o, dTipoInforme)
        Dim sql As String = "UPDATE tipoinforme SET nombre = '" & obj.NOMBRE & "' WHERE id = " & obj.ID

        Dim lista As New ArrayList
        lista.Add(sql)

     

        Return EjecutarTransaccion(lista)
    End Function
    Public Function eliminar(ByVal o As Object) As Boolean
        Dim obj As dTipoInforme = CType(o, dTipoInforme)
        Dim sql As String = "DELETE FROM tipoinforme WHERE id = " & obj.ID

        Dim lista As New ArrayList
        lista.Add(sql)

    

        Return EjecutarTransaccion(lista)
    End Function
    Public Function buscar(ByVal o As Object) As dTipoInforme
        Dim obj As dTipoInforme = CType(o, dTipoInforme)
        Dim t As New dTipoInforme
        Try
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL("SELECT id, nombre FROM tipoinforme WHERE id = " & obj.ID & "")

            If Ds.Tables(0).Rows.Count > 0 Then
                Dim unaFila As DataRow
                unaFila = Ds.Tables(0).Rows(0)
                t.ID = CType(unaFila.Item(0), Integer)
                t.NOMBRE = CType(unaFila.Item(1), String)
                Return t
            End If
            Return Nothing
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
    Public Function listar() As ArrayList
        Dim sql As String = "SELECT id, nombre FROM tipoinforme"
        Try
            Dim Lista As New ArrayList
            Dim Ds As New DataSet
            Ds = Me.EjecutarSQL(sql)
            If Ds.Tables(0).Rows.Count = 0 Then
                Return Nothing
            Else
                Dim unaFila As DataRow
                For Each unaFila In Ds.Tables(0).Rows
                    Dim t As New dTipoInforme
                    t.ID = CType(unaFila.Item(0), Integer)
                    t.NOMBRE = CType(unaFila.Item(1), String)
                    Lista.Add(t)
                Next
                Return Lista
            End If
        Catch ex As Exception
            Return Nothing
        End Try
    End Function
End Class
