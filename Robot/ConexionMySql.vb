﻿Imports mysql.Data.MySqlClient
Namespace Conectoras
    Public Class ConexionMySQL
        'Enumeración de errores.
        Public Enum RetornosError
            ok
            OtroError
        End Enum
        'String que contiene la cadena de conexión a BD.

        Private Shared mCadenaConexion As String = "Server=192.168.1.20;Database=colaveco;Uid=root;Pwd=root;Convert Zero Datetime=True"
        'Private Shared mCadenaConexion As String = "Server=localhost;Database=colaveco;Uid=root;Pwd=root;Convert Zero Datetime=True"

        Public ReadOnly Property CadenaConexion() As String
            Get
                Return mCadenaConexion
            End Get
        End Property
        'Función que brinda conexión a BD.
        Public Function Conectar() As MySqlConnection
            Dim c As MySqlConnection
            Try
                c = New MySqlConnection(CadenaConexion)
                Return c
            Catch
                MessageBox.Show(Err.Description)
                Return Nothing
            End Try
        End Function
        'Función que realiza consulta a BD.
        Public Function EjecutarSQL(ByVal CadenaConsulta As String) As DataSet
            If Not CadenaConsulta = String.Empty Then
                Dim unaC As MySqlConnection = Me.Conectar
                Dim unDs As New DataSet, unDA As MySqlDataAdapter
                Try
                    unaC.Open()
                    unDA = New MySqlDataAdapter(CadenaConsulta, unaC)
                    Dim unCB As MySqlCommandBuilder = New MySqlCommandBuilder(unDA)

                    unDA.Fill(unDs)
                    unaC.Close()
                    Return unDs
                Catch
                    MessageBox.Show(Err.Description & " - " & Err.Source & " Línea: " & Err.Erl)
                End Try
            End If
            Return Nothing
        End Function
        Public Function EjecutarTransaccion(ByVal sentencias As ArrayList) As Boolean
            Dim hecho As Boolean = False

            Dim myConnection As New MySqlConnection(Me.CadenaConexion)
            myConnection.Open()

            Dim myCommand As MySqlCommand = myConnection.CreateCommand()
            Dim myTrans As MySqlTransaction

            ' Inicia transacción local.
            myTrans = myConnection.BeginTransaction()
            ' Must assign both transaction object and connection
            ' to Command object for a pending local transaction
            myCommand.Connection = myConnection
            myCommand.Transaction = myTrans

            Try
                Dim sentencia As String
                For Each sentencia In sentencias
                    myCommand.CommandText = sentencia
                    myCommand.ExecuteNonQuery()
                Next
                myTrans.Commit()
                hecho = True
            Catch e As Exception
                'MsgBox(e.Message)
                Try
                    myTrans.Rollback()
                Catch ex As MySqlException
                    If Not myTrans.Connection Is Nothing Then
                        Console.WriteLine("An exception of type " + ex.GetType().ToString() + _
                                          " was encountered while attempting to roll back the transaction.")
                    End If
                End Try

                Console.WriteLine("An exception of type " + e.GetType().ToString() + _
                                "was encountered while inserting the data.")
                Console.WriteLine("Neither record was written to database.")
            Finally
                myConnection.Close()
            End Try
            Return hecho
        End Function
    End Class

End Namespace

