﻿Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel
Imports System
Imports System.IO
Imports System.Collections
Imports System.Net.FtpWebRequest
Imports System.Net
Imports Newtonsoft.Json
Public Class FormInicio
    Public nombre_pc As String = ""
    Private productorweb_com As String = ""
    Private copiaproductorweb_com As String = ""
    Private idproductorweb_com As Long = 0
    Private copiaidproductorweb_com As Long = 0
    Private idficha As Long = 0
    Private id_ficha_ As Long = 0
    Private tipoinforme As Integer = 0
    Private email As String = ""
    Private email2 As String = ""
    Private celular As String = ""
    Private nficha As Long = 0
    Private mensaje As String = ""
    Private excel As Integer = 0
    Private pdf As Integer = 0
    Private csv As Integer = 0
    Private enviar_copia As String = ""
    Private _abonado As Integer = 0
    Private _comentarios As String = ""
    Private _moroso As Integer = 0
    Private _tipoinforme As Long = 0
    Private subido As Integer = 0
    Private carpeta As Long = 0
    Private compraid As Long = 0
    Private clienteid As Long = 0
    Private tipoinformeweb As String
    Private nmuestrasweb As Integer
    Private muestraweb As String
    Private subtipoinformeweb As String
    Private observacionesweb As String
    Private nombreproductorweb As String
    Private factura_origen As String
    Private userid As Integer
    Private tipo As Integer
    Public Sub New()
        ' Llamada necesaria para el Diseñador de Windows Forms.
        InitializeComponent()
        ' Agregue cualquier inicialización después de la llamada a InitializeComponent().
        nombre_pc = My.Computer.Name
        If nombre_pc = "ROBOT" Then
            Timer1.Enabled = True
        End If
    End Sub

#Region "TIMER`s"

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        ListBox1.Items.Clear()
        ListBox1.Items.Add(Now)
        ListBox1.Items.Add("Revisando cuentas corrientes")
        revisar_cuentas_corrientes()
        actualizar_cajas()
        Timer1.Enabled = False
        Timer2.Enabled = True
    End Sub
    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        ListBox1.Items.Clear()
        ListBox1.Items.Add(Now)
        ListBox1.Items.Add("Subiendo cuentas corrientes")
        subir_ctacte()
        subir_ctacte2()
        Timer2.Enabled = False
        Timer3.Enabled = True
    End Sub
    Private Sub Timer3_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer3.Tick
        'moverarchivossubidos()
        DateFecha.Value = Now
        If nombre_pc = "ROBOT" Then
            ListBox1.Items.Clear()
            ListBox1.Items.Add(Now)
            ListBox1.Items.Add("Importando")
            importar()
            Dim pi As New dPreinformes
            Dim lista As New ArrayList
            Dim creapreinformecalidad As Integer = 1
            lista = pi.listarsinmarcar

            If Not lista Is Nothing Then
                If lista.Count > 0 Then
                    For Each pi In lista
                        If pi.TIPO = 17 Or pi.TIPO = 18 Then
                            tipo = pi.TIPO
                        End If

                        barra = barra + 1
                        ProgressBar1.Value = barra
                        If barra = 100 Then
                            barra = 0
                        End If
                        If pi.TIPO = 10 Then
                            Dim csm As New dCalidadSolicitudMuestra
                            Dim listacsm As New ArrayList
                            Dim ficha As Long = pi.FICHA
                            listacsm = csm.listarporsolicitud(ficha)
                            If Not listacsm Is Nothing Then
                                creapreinformecalidad = 1
                                For Each csm In listacsm
                                    If csm.RB = 1 Then
                                        Dim ibc As New dIbc
                                        ibc.FICHA = csm.FICHA
                                        ibc.MUESTRA = csm.MUESTRA
                                        ibc = ibc.buscarxfichaxmuestra
                                        If Not ibc Is Nothing Then
                                        Else
                                            creapreinformecalidad = 0
                                            Exit For
                                        End If
                                        ibc = Nothing
                                    End If
                                    If csm.PSICROTROFOS = 1 Then
                                        Dim psi As New dPsicrotrofos
                                        psi.FICHA = csm.FICHA
                                        psi.MUESTRA = csm.MUESTRA
                                        psi = psi.buscarxfichaxmuestra
                                        If Not psi Is Nothing Then
                                        Else
                                            creapreinformecalidad = 0
                                            Exit For
                                        End If
                                        psi = Nothing
                                    End If
                                    If csm.ESPORULADOS = 1 Then
                                        Dim esp As New dEsporulados
                                        esp.FICHA = csm.FICHA
                                        esp.MUESTRA = csm.MUESTRA
                                        esp = esp.buscarxfichaxmuestra
                                        If Not esp Is Nothing Then
                                        Else
                                            creapreinformecalidad = 0
                                            Exit For
                                        End If
                                        esp = Nothing
                                    End If
                                    If csm.INHIBIDORES = 1 Then
                                        Dim inh As New dInhibidores
                                        inh.FICHA = csm.FICHA
                                        inh.MUESTRA = csm.MUESTRA
                                        inh = inh.buscarxfichaxmuestra
                                        If Not inh Is Nothing Then
                                            If inh.MARCA = 0 Then
                                                creapreinformecalidad = 0
                                                Exit For
                                            End If
                                        Else
                                            creapreinformecalidad = 0
                                            Exit For
                                        End If
                                        inh = Nothing
                                    End If
                                Next
                            End If

                            If creapreinformecalidad = 1 Then
                                ListBox1.Items.Clear()
                                ListBox1.Items.Add(Now)
                                ListBox1.Items.Add("Creando preinforme de calidad")
                                preinforme_calidad(ficha)
                            End If
                        End If
                    Next
                End If
            End If
            pre_informe_control()
            ListBox1.Items.Clear()
            ListBox1.Items.Add(Now)
            ListBox1.Items.Add("Subiendo informes de control lechero")
            subir_informes_control() ' CONTROL LECHERO
            ListBox1.Items.Clear()
            ListBox1.Items.Add(Now)
            ListBox1.Items.Add("Subiendo informes de agua")
            subir_informes_agua() ' AGUA
            ListBox1.Items.Clear()
            ListBox1.Items.Add(Now)
            ListBox1.Items.Add("Subiendo informes de efluentes")
            subir_informes_efluentes() ' EFLUENTES
            ListBox1.Items.Clear()
            ListBox1.Items.Add(Now)
            ListBox1.Items.Add("Subiendo informes de ATB")
            subir_informes_atb() ' ATB
            ListBox1.Items.Clear()
            ListBox1.Items.Add(Now)
            ListBox1.Items.Add("Subiendo informes de parasitología")
            subir_informes_parasitologia() ' PARASITOLOGIA
            ListBox1.Items.Clear()
            ListBox1.Items.Add(Now)
            ListBox1.Items.Add("Subiendo informes de patología")
            subir_informes_patologia() ' PATOLOGIA
            ListBox1.Items.Clear()
            ListBox1.Items.Add(Now)
            ListBox1.Items.Add("Subiendo informes de toxicología")
            subir_informes_toxicologia() ' TOXICOLOGIA
            ListBox1.Items.Clear()
            ListBox1.Items.Add(Now)
            ListBox1.Items.Add("Subiendo informes de alimentos")
            subir_informes_subproductos() ' ALIMENTOS
            ListBox1.Items.Clear()
            ListBox1.Items.Add(Now)
            ListBox1.Items.Add("Subiendo informes de serología")
            subir_informes_serologia() ' SEROLOGIA
            ListBox1.Items.Clear()
            ListBox1.Items.Add(Now)
            ListBox1.Items.Add("Subiendo informes de calidad de leche")
            subir_informes_calidad() ' CALIDAD
            ListBox1.Items.Clear()
            ListBox1.Items.Add(Now)
            ListBox1.Items.Add("Subiendo informes de nutrición")
            subir_informes_nutricion() ' NUTRICION
            ListBox1.Items.Clear()
            ListBox1.Items.Add(Now)
            ListBox1.Items.Add("Subiendo informes de suelos")
            subir_informes_suelos() ' SUELOS
            ListBox1.Items.Clear()
            ListBox1.Items.Add(Now)
            ListBox1.Items.Add("Subiendo informes de brucelosis")
            subir_informes_brucelosis() ' BRUCELOSIS EN LECHE
            ListBox1.Items.Clear()
            ListBox1.Items.Add(Now)
            ListBox1.Items.Add("Subiendo informes de ambiental")
            subir_informes_ambiental() ' AMBIENTAL
            ListBox1.Items.Add(Now)
            ListBox1.Items.Add("Subiendo informes de foliar")
            subir_informes_foliar() ' FOLIARES
            ListBox1.Items.Clear()
            ListBox1.Items.Add(Now)
            ListBox1.Items.Add("Subiendo informes de Bacteriología")
            subir_informes_bacteriologia() ' BACTERIOLOGIA

            'Enviar correos a proveedores
            Dim fecha As Date
            fecha = DateFecha.Value.ToString("yyyy-MM-dd")
            Dim hoy As Date = Now
            Dim hora As Integer = 0
            Dim minuto As Integer = 0
            Dim dia As Integer = 0
            hora = hoy.Hour
            minuto = hoy.Minute
            dia = hoy.Day
            If fecha.DayOfWeek = DayOfWeek.Thursday Then
                If dia < 27 Then
                    If hora > 9 Then
                        enviarcompras()
                        'enviarMailCajas()
                    End If
                End If
            End If
            If hora = 18 And minuto > 30 And minuto < 35 Then
                calcularsaldos()
            End If
            'mover archivos subidos
            Dim hoy2 As Date = Now
            Dim hora2 As Integer = 0
            hora2 = hoy2.Hour
            If hora2 = 20 Then
                moverarchivossubidos()
            End If
        End If
        Timer3.Enabled = False
        Timer4.Enabled = True
    End Sub
    Private Sub Timer4_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer4.Tick
        If nombre_pc = "ROBOT" Then
            solicitud_frascos()
            actualizacion_clientes()
            ListBox1.Items.Clear()
            ListBox1.Items.Add(Now)
            ListBox1.Items.Add("Cargando pedidos web")
            cargarpedidosweb()
            ListBox1.Items.Clear()
            ListBox1.Items.Add(Now)
            ListBox1.Items.Add("Cargando solicitudes web")
            cargarsolicitudesweb()
            enviaravisoenviocajas()
        End If
        DateFecha.Value = Now
        Timer4.Enabled = False
        Timer1.Enabled = True
    End Sub

#End Region

#Region "IMPORTADOR"

    Private Sub importar()
        borrarimportacionescalidad()
        borrarimportacionescontrol()
        borrarimportacionescalidad2()
        borrarimportacionescontrol2()
        borrarimportacionescalidad3()
        borrarimportacionescontrol3()
        calidadcsv()
        calidadcsv2()
        calidadxls()
        calidadfat()
        ibc()
        controllecherocsv()
        controllecherocsv2()
        controllecheroxls()
        controllecherofat()
    End Sub
    Private Sub borrarimportacionescalidad()
        Dim nombrearchivo As String = ""
        Dim folder As New DirectoryInfo("\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Calidad de leche")
        Dim _ficheros() As String
        _ficheros = Directory.GetFiles("\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Calidad de leche")
        'If (_ficheros.Length > 0) Then
        '    For Each file As FileInfo In folder.GetFiles("*.csv")
        '        nombrearchivo = file.Name
        '        Dim ficha As String = ""
        '        Dim _calidad As New dCalidad
        '        Dim pi As New dPreinformes
        '        ficha = Mid(file.Name, 1, Len(file.Name) - 4)
        '        _calidad.FICHA = ficha
        '        _calidad.eliminarxficha()
        '        pi.FICHA = ficha
        '        pi.desmarcarcreado()
        '    Next
        'End If
        If (_ficheros.Length > 0) Then
            For Each file As FileInfo In folder.GetFiles("*.csv")
                nombrearchivo = file.Name
                Dim ficha As String = ""
                Dim ficha2 As String = ""
                Dim _calidad As New dCalidad
                Dim pi As New dPreinformes
                If file.Name.Length < 15 Then
                    ficha = Mid(file.Name, 1, Len(file.Name) - 4)
                ElseIf file.Name.Length = 27 Then
                    ficha = Mid(file.Name, Len(file.Name) - 26, 6)
                ElseIf file.Name.Length = 28 Then
                    ficha = Mid(file.Name, Len(file.Name) - 27, 6)
                End If
                _calidad.FICHA = ficha
                _calidad.eliminarxficha()
                ficha2 = Mid(file.Name, Len(file.Name) - 4, 1)

                If ficha2 = "a" Or ficha2 = "A" Or ficha2 = "b" Or ficha2 = "B" Or ficha2 = "c" Or ficha2 = "C" Or ficha2 = "d" Or ficha2 = "D" Or ficha2 = "e" Or ficha2 = "E" Or ficha2 = "f" Or ficha2 = "F" Or ficha2 = "g" Or ficha2 = "G" Or ficha2 = "h" Or ficha2 = "H" Or ficha2 = "i" Or ficha2 = "I" Or ficha2 = "j" Or ficha2 = "J" Then
                    ficha = Mid(file.Name, 1, Len(file.Name) - 5)
                End If
                pi.FICHA = ficha
                pi.desmarcarcreado()
            Next
        End If
    End Sub
    Private Sub borrarimportacionescalidad2()
        Dim nombrearchivo As String = ""
        Dim folder As New DirectoryInfo("\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Calidad de leche")
        Dim _ficheros() As String
        _ficheros = Directory.GetFiles("\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Calidad de leche")
        'If (_ficheros.Length > 0) Then
        '    For Each file As FileInfo In folder.GetFiles("*.fat")
        '        nombrearchivo = file.Name
        '        Dim ficha As String = ""
        '        Dim _calidad As New dCalidad
        '        Dim pi As New dPreinformes
        '        ficha = Mid(file.Name, 1, Len(file.Name) - 4)
        '        _calidad.FICHA = ficha
        '        _calidad.eliminarxficha()
        '        pi.FICHA = ficha
        '        pi.desmarcarcreado()
        '    Next
        'End If
        If (_ficheros.Length > 0) Then
            For Each file As FileInfo In folder.GetFiles("*.fat")
                nombrearchivo = file.Name
                Dim ficha As String = ""
                Dim ficha2 As String = ""
                Dim _control As New dCalidad
                Dim pi As New dPreinformes
                ficha = Mid(file.Name, 1, Len(file.Name) - 4)
                _control.FICHA = ficha
                _control.eliminarxficha()
                ficha2 = Mid(file.Name, Len(file.Name) - 4, 1)
                If ficha2 = "a" Or ficha2 = "A" Or ficha2 = "b" Or ficha2 = "B" Or ficha2 = "c" Or ficha2 = "C" Or ficha2 = "d" Or ficha2 = "D" Or ficha2 = "e" Or ficha2 = "E" Or ficha2 = "f" Or ficha2 = "F" Or ficha2 = "g" Or ficha2 = "G" Or ficha2 = "h" Or ficha2 = "H" Or ficha2 = "i" Or ficha2 = "I" Or ficha2 = "j" Or ficha2 = "J" Then
                    ficha = Mid(file.Name, 1, Len(file.Name) - 5)
                End If
                pi.FICHA = ficha
                pi.desmarcarcreado()
            Next
        End If
    End Sub
    Private Sub borrarimportacionescalidad3()
        Dim nombrearchivo As String = ""
        Dim folder As New DirectoryInfo("\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Calidad de leche")
        Dim _ficheros() As String
        _ficheros = Directory.GetFiles("\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Calidad de leche")
        'If (_ficheros.Length > 0) Then
        '    For Each file As FileInfo In folder.GetFiles("*.xls")
        '        nombrearchivo = file.Name
        '        Dim ficha As String = ""
        '        Dim _calidad As New dCalidad
        '        Dim pi As New dPreinformes
        '        ficha = Mid(file.Name, 1, Len(file.Name) - 4)
        '        _calidad.FICHA = ficha
        '        _calidad.eliminarxficha()
        '        pi.FICHA = ficha
        '        pi.desmarcarcreado()
        '    Next
        'End If
        If (_ficheros.Length > 0) Then
            For Each file As FileInfo In folder.GetFiles("*.xls")
                nombrearchivo = file.Name
                Dim ficha As String = ""
                Dim ficha2 As String = ""
                Dim _control As New dCalidad
                Dim pi As New dPreinformes
                ficha = Mid(file.Name, 1, Len(file.Name) - 4)
                _control.FICHA = ficha
                _control.eliminarxficha()
                ficha2 = Mid(file.Name, Len(file.Name) - 4, 1)
                If ficha2 = "a" Or ficha2 = "A" Or ficha2 = "b" Or ficha2 = "B" Or ficha2 = "c" Or ficha2 = "C" Or ficha2 = "d" Or ficha2 = "D" Or ficha2 = "e" Or ficha2 = "E" Or ficha2 = "f" Or ficha2 = "F" Or ficha2 = "g" Or ficha2 = "G" Or ficha2 = "h" Or ficha2 = "H" Or ficha2 = "i" Or ficha2 = "I" Or ficha2 = "j" Or ficha2 = "J" Then
                    ficha = Mid(file.Name, 1, Len(file.Name) - 5)
                End If
                pi.FICHA = ficha
                pi.desmarcarcreado()
            Next
        End If
    End Sub
    Private Sub borrarimportacionescontrol()
        Dim nombrearchivo As String = ""
        Dim folder As New DirectoryInfo("\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Control lechero")
        Dim _ficheros() As String
        _ficheros = Directory.GetFiles("\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Control lechero")
        If (_ficheros.Length > 0) Then
            For Each file As FileInfo In folder.GetFiles("*.csv")
                nombrearchivo = file.Name
                Dim ficha As String = ""
                Dim ficha2 As String = ""
                Dim _control As New dControl
                Dim pi As New dPreinformes
                If file.Name.Length < 15 Then
                    ficha = Mid(file.Name, 1, Len(file.Name) - 4)
                ElseIf file.Name.Length = 27 Then
                    ficha = Mid(file.Name, Len(file.Name) - 26, 6)
                ElseIf file.Name.Length = 28 Then
                    ficha = Mid(file.Name, Len(file.Name) - 27, 6)
                End If
                _control.FICHA = ficha
                _control.eliminarxficha()

                ficha2 = Mid(file.Name, Len(file.Name) - 4, 1)
                If ficha2 = "a" Or ficha2 = "A" Or ficha2 = "b" Or ficha2 = "B" Or ficha2 = "c" Or ficha2 = "C" Or ficha2 = "d" Or ficha2 = "D" Or ficha2 = "e" Or ficha2 = "E" Or ficha2 = "f" Or ficha2 = "F" Or ficha2 = "g" Or ficha2 = "G" Or ficha2 = "h" Or ficha2 = "H" Or ficha2 = "i" Or ficha2 = "I" Or ficha2 = "j" Or ficha2 = "J" Then
                    ficha = Mid(file.Name, 1, Len(file.Name) - 5)
                End If
                pi.FICHA = ficha
                pi.desmarcarcreado()
            Next
        End If
    End Sub
    Private Sub borrarimportacionescontrol2()
        Dim nombrearchivo As String = ""
        Dim folder As New DirectoryInfo("\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Control lechero")
        Dim _ficheros() As String
        _ficheros = Directory.GetFiles("\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Control lechero")
        If (_ficheros.Length > 0) Then
            For Each file As FileInfo In folder.GetFiles("*.fat")
                nombrearchivo = file.Name
                Dim ficha As String = ""
                Dim ficha2 As String = ""
                Dim _control As New dControl
                Dim pi As New dPreinformes
                ficha = Mid(file.Name, 1, Len(file.Name) - 4)
                _control.FICHA = ficha
                _control.eliminarxficha()
                ficha2 = Mid(file.Name, Len(file.Name) - 4, 1)
                If ficha2 = "a" Or ficha2 = "A" Or ficha2 = "b" Or ficha2 = "B" Or ficha2 = "c" Or ficha2 = "C" Or ficha2 = "d" Or ficha2 = "D" Or ficha2 = "e" Or ficha2 = "E" Or ficha2 = "f" Or ficha2 = "F" Or ficha2 = "g" Or ficha2 = "G" Or ficha2 = "h" Or ficha2 = "H" Or ficha2 = "i" Or ficha2 = "I" Or ficha2 = "j" Or ficha2 = "J" Then
                    ficha = Mid(file.Name, 1, Len(file.Name) - 5)
                End If
                pi.FICHA = ficha
                pi.desmarcarcreado()
            Next
        End If
    End Sub
    Private Sub borrarimportacionescontrol3()
        Dim nombrearchivo As String = ""
        Dim folder As New DirectoryInfo("\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Control lechero")
        Dim _ficheros() As String
        _ficheros = Directory.GetFiles("\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Control lechero")
        If (_ficheros.Length > 0) Then
            For Each file As FileInfo In folder.GetFiles("*.xls")
                nombrearchivo = file.Name
                Dim ficha As String = ""
                Dim ficha2 As String = ""
                Dim _control As New dControl
                Dim pi As New dPreinformes
                ficha = Mid(file.Name, 1, Len(file.Name) - 4)
                _control.FICHA = ficha
                _control.eliminarxficha()
                ficha2 = Mid(file.Name, Len(file.Name) - 4, 1)
                If ficha2 = "a" Or ficha2 = "A" Or ficha2 = "b" Or ficha2 = "B" Or ficha2 = "c" Or ficha2 = "C" Or ficha2 = "d" Or ficha2 = "D" Or ficha2 = "e" Or ficha2 = "E" Or ficha2 = "f" Or ficha2 = "F" Or ficha2 = "g" Or ficha2 = "G" Or ficha2 = "h" Or ficha2 = "H" Or ficha2 = "i" Or ficha2 = "I" Or ficha2 = "j" Or ficha2 = "J" Then
                    ficha = Mid(file.Name, 1, Len(file.Name) - 5)
                End If
                pi.FICHA = ficha
                pi.desmarcarcreado()
            Next
        End If
    End Sub
    Private Sub calidadcsv()
        Dim extension As String
        Dim nombrearchivo As String = ""
        Dim linea As Integer
        Dim folder As New DirectoryInfo("\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Calidad de leche")
        Dim _ficheros() As String
        _ficheros = Directory.GetFiles("\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Calidad de leche")
        If Not (_ficheros.Length > 0) Then
        Else
            For Each file As FileInfo In folder.GetFiles("*.csv")
                nombrearchivo = file.Name
                If nombrearchivo.Length < 12 Then 'controlo si el archivo es de delta 400
                    linea = 1
                    extension = Microsoft.VisualBasic.Right(file.Name, 3)
                    Dim objReader As New StreamReader("\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Calidad de leche\" & file.Name)
                    Dim sLine As String = ""
                    Dim arraytext() As String
                    Dim matricula As String = ""
                    Dim grasa As Double = 0
                    Dim proteina As Double = 0
                    Dim lactosa As Double = 0
                    Dim st As Double = 0
                    Dim rc As Integer = 0
                    Dim ficha As String = ""
                    Dim ficha2 As String = ""
                    Dim ficha3 As String = ""
                    Dim equipo As String = ""
                    Dim producto As String = ""
                    Dim crioscopia As Integer = 0
                    Dim urea As Integer = 0
                    Dim proteinav As Double = 0
                    Dim caseina As Double = 0
                    Dim densidad As Double = 0
                    Dim ph As Double = 0
                    ' *** SI EL ARCHIVO ES CSV **************************************************************************************
                    If extension = "csv" Or extension = "CSV" Then
                        Dim c As New dImpCalidad()
                        Do
                            sLine = objReader.ReadLine()
                            If sLine <> " " Then
                                If linea = 3 Then
                                    arraytext = Split(sLine, ";")
                                    If arraytext.Length < 11 Then
                                        arraytext = Split(sLine, ",")
                                    End If
                                    producto = Trim(arraytext(10))
                                End If
                                If Not sLine Is Nothing Then
                                    If linea >= 8 Then
                                        arraytext = Split(sLine, ";")
                                        If arraytext.Length < 39 Then
                                            arraytext = Split(sLine, ";")
                                        End If
                                        matricula = Trim(arraytext(1))
                                        If arraytext.Length <= 5 Then
                                            grasa = -1
                                            proteina = -1
                                            lactosa = -1
                                            st = -1
                                            If Trim(arraytext(3)) <> "" And Trim(arraytext(3)) <> "-" Then
                                                Try
                                                    rc = arraytext(3)
                                                Catch ex As Exception
                                                    MsgBox("Error en archivo: " & file.Name & ", línea: " & linea)
                                                End Try
                                            Else
                                                rc = -1
                                            End If
                                            crioscopia = -1
                                            urea = -1
                                            proteinav = -1
                                            caseina = -1
                                            densidad = -1
                                            ph = -1
                                        Else
                                            If Trim(arraytext(5)) = "" Or Trim(arraytext(5)) = "-" Then
                                                grasa = -1
                                            Else
                                                Try
                                                    grasa = arraytext(5)
                                                Catch ex As Exception
                                                    MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: Grasa")
                                                    Exit Sub
                                                End Try
                                            End If
                                            If Trim(arraytext(6)) = "" Or Trim(arraytext(6)) = "-" Then
                                                proteina = -1
                                            Else
                                                Try
                                                    proteina = arraytext(6)
                                                Catch ex As Exception
                                                    MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: Proteína")
                                                    Exit Sub
                                                End Try
                                            End If
                                            If Trim(arraytext(7)) = "" Or Trim(arraytext(7)) = "-" Then
                                                lactosa = -1
                                            Else
                                                Try
                                                    lactosa = arraytext(7)
                                                Catch ex As Exception
                                                    MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: Lactosa")
                                                    Exit Sub
                                                End Try
                                            End If
                                            If Trim(arraytext(8)) = "" Or Trim(arraytext(8)) = "-" Then
                                                st = -1
                                            Else
                                                Try
                                                    st = arraytext(8)
                                                Catch ex As Exception
                                                    MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: Sólidos totales")
                                                    Exit Sub
                                                End Try
                                            End If
                                            If Trim(arraytext(3)) = "" Or Trim(arraytext(3)) = "-" Then
                                                rc = -1
                                            Else
                                                Try
                                                    rc = arraytext(3)
                                                Catch ex As Exception
                                                    MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: RC")
                                                    Exit Sub
                                                End Try
                                            End If
                                            '** IMPORTAR CRIOSCOPIA **************************************************************************
                                            If Trim(arraytext(9)) = "" Or Trim(arraytext(9)) = "-" Then
                                                crioscopia = -1
                                            Else
                                                Try
                                                    crioscopia = arraytext(9)
                                                Catch ex As Exception
                                                    MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: Crioscopía")
                                                    Exit Sub
                                                End Try
                                            End If
                                            '***************************************************************************************************
                                            If Trim(arraytext(10)) = "" Or Trim(arraytext(10)) = "-" Then
                                                urea = -1
                                            Else
                                                Try
                                                    urea = arraytext(10)
                                                Catch ex As Exception
                                                    MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: Urea")
                                                    Exit Sub
                                                End Try
                                            End If
                                            'If Trim(arraytext(28)) = "" Or Trim(arraytext(28)) = "-" Then
                                            '    proteinav = -1
                                            'Else
                                            '    Try
                                            '        proteinav = arraytext(28)

                                            '    Catch ex As Exception
                                            '        MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: Proteína verdadera")
                                            '        Exit Sub
                                            '    End Try
                                            'End If
                                            If Trim(arraytext(11)) = "" Or Trim(arraytext(11)) = "-" Then
                                                caseina = -1
                                            Else
                                                Try
                                                    caseina = arraytext(11)
                                                Catch ex As Exception
                                                    MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: Caseína")
                                                    Exit Sub
                                                End Try
                                            End If
                                            'If Trim(arraytext(30)) = "" Or Trim(arraytext(30)) = "-" Then
                                            '    densidad = -1
                                            'Else
                                            '    Try
                                            '        densidad = arraytext(30)
                                            '    Catch ex As Exception
                                            '        MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: Densidad")
                                            '        Exit Sub
                                            '    End Try
                                            'End If
                                            If Trim(arraytext(12)) = "" Or Trim(arraytext(12)) = "-" Then
                                                ph = -1
                                            Else
                                                Try
                                                    ph = arraytext(12)
                                                Catch ex As Exception
                                                    MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: pH")
                                                    Exit Sub
                                                End Try
                                            End If
                                        End If
                                        ficha2 = Mid(file.Name, Len(file.Name) - 4, 1)
                                        ficha3 = Mid(file.Name, 1, 1)
                                        If ficha2 = "a" Or ficha2 = "A" Or ficha2 = "b" Or ficha2 = "B" Or ficha2 = "c" Or ficha2 = "C" Or ficha2 = "d" Or ficha2 = "D" Or ficha2 = "e" Or ficha2 = "E" Or ficha2 = "f" Or ficha2 = "F" Or ficha2 = "g" Or ficha2 = "G" Or ficha2 = "h" Or ficha2 = "H" Or ficha2 = "i" Or ficha2 = "I" Or ficha2 = "j" Or ficha2 = "J" Or ficha2 = "k" Or ficha2 = "K" Then
                                            ficha = Mid(file.Name, 1, Len(file.Name) - 5)
                                        Else
                                            ficha = Mid(file.Name, 1, Len(file.Name) - 4)
                                        End If
                                        If Mid(ficha, 1, 1) = "l" Or Mid(ficha, 1, 1) = "L" Then
                                            Dim MyString As String = ficha
                                            Dim MyChar As Char() = {"l"c, "L"c}
                                            Dim NewString As String = MyString.TrimStart(MyChar)
                                            ficha3 = NewString
                                        Else
                                            ficha3 = ficha
                                        End If
                                        '**CONTROL DE CRIOSCOPIA *************************************************************************
                                        Dim cc As New dCrioscopia_Control
                                        Dim ficha_cc As Long = 0
                                        Dim muestra_cc As String = ""
                                        Dim res_delta As Integer = 0
                                        Dim res_crioscopo As Integer = 0
                                        Dim diferencia_cc As Integer = 0
                                        ficha_cc = ficha3
                                        muestra_cc = matricula
                                        cc.FICHA = ficha_cc
                                        cc.MUESTRA = muestra_cc
                                        cc = cc.buscarxfichaxmuestra
                                        If Not cc Is Nothing Then
                                            res_delta = cc.DELTA
                                            res_crioscopo = cc.CRIOSCOPO
                                            If res_delta > res_crioscopo Then
                                                diferencia_cc = res_delta - res_crioscopo
                                            Else
                                                diferencia_cc = res_crioscopo - res_delta
                                            End If
                                            If diferencia_cc > 5 Then
                                                crioscopia = res_crioscopo
                                            End If
                                        End If
                                        cc = Nothing
                                        ficha_cc = Nothing
                                        muestra_cc = Nothing
                                        res_delta = Nothing
                                        res_crioscopo = Nothing
                                        diferencia_cc = Nothing
                                        '*************************************************************************************************
                                        Dim fechaoriginal As Date = Now()
                                        Dim fecha As String
                                        fecha = Format(fechaoriginal, "yyyy-MM-dd")
                                        c.FICHA = ficha3
                                        c.FECHA = fecha
                                        c.EQUIPO = "delta"
                                        c.PRODUCTO = producto
                                        c.MUESTRA = matricula
                                        c.RC = rc
                                        c.GRASA = grasa
                                        c.PROTEINA = proteina
                                        c.LACTOSA = lactosa
                                        c.ST = st
                                        c.CRIOSCOPIA = crioscopia
                                        c.UREA = urea
                                        c.PROTEINAV = proteinav
                                        c.CASEINA = caseina
                                        c.DENSIDAD = densidad
                                        c.PH = ph
                                        c.guardar()
                                        Dim csm As New dCalidadSolicitudMuestra
                                        Dim listacsm As New ArrayList
                                        Dim noactualizafecha As Integer = 0
                                        listacsm = csm.listarporsolicitud(ficha3)
                                        If Not listacsm Is Nothing Then
                                            For Each csm In listacsm
                                                If csm.RB = 1 Then
                                                    noactualizafecha = 1
                                                End If
                                            Next
                                        End If
                                        If noactualizafecha = 0 Then
                                            Dim sa As New dSolicitudAnalisis
                                            sa.ID = ficha3
                                            sa.actualizarfechaproceso(fecha)
                                            sa = Nothing
                                        End If
                                    End If
                                End If
                            End If
                            linea = linea + 1
                        Loop Until sLine Is Nothing
                        objReader.Close()
                    End If
                    '*** MOVER ARCHIVO ***********************************************************************
                    Dim sArchivoOrigen As String = "\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Calidad de leche\" & nombrearchivo
                    Dim sRutaDestino As String = "Y:\documentos\secretaria\analisis\leche\bentley-delta\pasados\calidad de leche\pasados NET\" & nombrearchivo
                    Try
                        ' Mover el fichero.si existe lo sobreescribe  
                        My.Computer.FileSystem.MoveFile(sArchivoOrigen, _
                                                        sRutaDestino, _
                                                        True)
                        'MsgBox("Ok.", MsgBoxStyle.Information, "Mover archivo")
                        ' errores  
                    Catch ex As Exception
                        MsgBox(ex.Message.ToString, MsgBoxStyle.Critical)
                    End Try
                    '***********************************
                    'Insert tabla preinforme_calidad
                    Dim pi As New dPreinformes
                    Dim fechaactual As Date = Now()
                    Dim _fecha As String
                    _fecha = Format(fechaactual, "yyyy-MM-dd")
                    pi.FICHA = ficha3
                    pi = pi.buscar
                    If Not pi Is Nothing Then
                    Else
                        Dim pi2 As New dPreinformes
                        Dim sa As New dSolicitudAnalisis
                        Dim tipof As Integer = 0
                        sa.ID = ficha3
                        sa = sa.buscar
                        If Not sa Is Nothing Then
                            tipof = sa.IDTIPOINFORME
                        End If
                        pi2.FICHA = ficha3
                        pi2.TIPO = tipof
                        pi2.CREADO = 0
                        pi2.FECHA = _fecha
                        pi2.guardar()
                        pi2 = Nothing
                        sa = Nothing
                    End If
                    pi = Nothing
                    ' Grabar estado de la ficha
                    Dim est As New dEstados
                    est.FICHA = ficha3
                    est.ESTADO = 4
                    est.FECHA = _fecha
                    est.guardar2()
                    est = Nothing
                End If 'fin de control si archivoes de delta 400
            Next
        End If
    End Sub
    Private Sub calidadcsv2()
        Dim extension As String
        Dim nombrearchivo As String = ""
        Dim linea As Integer
        Dim folder As New DirectoryInfo("\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Calidad de leche")
        Dim _ficheros() As String
        _ficheros = Directory.GetFiles("\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Calidad de leche")
        If Not (_ficheros.Length > 0) Then
        Else
            For Each file As FileInfo In folder.GetFiles("*.csv")
                nombrearchivo = file.Name
                If nombrearchivo.Length > 12 Then 'controlo si el archivo es de delta nuevo
                    linea = 1
                    extension = Microsoft.VisualBasic.Right(file.Name, 3)
                    Dim objReader As New StreamReader("\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Calidad de leche\" & file.Name)
                    Dim sLine As String = ""
                    Dim arraytext() As String
                    Dim matricula As String = ""
                    Dim grasa As Double = 0
                    Dim proteina As Double = 0
                    Dim lactosa As Double = 0
                    Dim st As Double = 0
                    Dim rc As Integer = 0
                    Dim ficha As String = ""
                    Dim ficha2 As String = ""
                    Dim ficha3 As String = ""
                    Dim equipo As String = ""
                    Dim producto As String = ""
                    Dim crioscopia As Integer = 0
                    Dim urea As Integer = 0
                    Dim proteinav As Double = 0
                    Dim caseina As Double = 0
                    Dim densidad As Double = 0
                    Dim ph As Double = 0
                    Dim grasa_b As Double = 0
                    Dim grasa_a As Double = 0
                    Dim cit As Integer = 0
                    Dim agl As Double = 0
                    Dim sng As Double = 0
                    Dim sfa As Double = 0
                    Dim ufa As Double = 0
                    Dim mufa As Double = 0
                    Dim pufa As Double = 0
                    Dim c16 As Double = 0
                    Dim c180 As Double = 0
                    Dim c181 As Double = 0
                    Dim bhb As Double = 0
                    Dim acetone As Double = 0
                    Dim cisfat As Double = 0
                    Dim transfat As Double = 0
                    Dim denovofa As Double = 0
                    Dim mixedfa As Double = 0
                    Dim preformedfa As Double = 0
                    Dim denovofa2 As Double = 0
                    Dim mixedfa2 As Double = 0
                    Dim preformedfa2 As Double = 0
                    Dim nefa As Double = 0
                    ' *** SI EL ARCHIVO ES CSV **************************************************************************************
                    If extension = "csv" Or extension = "CSV" Then
                        Dim c As New dImpCalidad()
                        Do
                            sLine = objReader.ReadLine()
                            If sLine <> " " And sLine <> "" Then
                                If linea = 3 Then
                                    arraytext = Split(sLine, ";")
                                    If arraytext.Length < 11 Then
                                        arraytext = Split(sLine, ";")
                                    End If
                                    producto = Trim(arraytext(10))
                                End If
                                If Not sLine Is Nothing Then
                                    If linea >= 8 Then
                                        arraytext = Split(sLine, ";")
                                        If arraytext.Length < 39 Then
                                            arraytext = Split(sLine, ";")
                                        End If

                                        If arraytext.Count >= 5 Then
                                            matricula = Trim(arraytext(5))
                                        End If


                                        If arraytext.Length <= 13 Then
                                            grasa = -1
                                            proteina = -1
                                            lactosa = -1
                                            st = -1
                                            If Trim(arraytext(11)) <> "" And Trim(arraytext(11)) <> "-" Then
                                                Try
                                                    rc = arraytext(11)
                                                Catch ex As Exception
                                                    MsgBox("Error en archivo: " & file.Name & ", línea: " & linea)
                                                End Try
                                            Else
                                                rc = -1
                                            End If
                                            crioscopia = -1
                                            urea = -1
                                            proteinav = -1
                                            caseina = -1
                                            densidad = -1
                                            ph = -1
                                        Else
                                            If Trim(arraytext(11)) = "" Or Trim(arraytext(11)) = "-" Then
                                                grasa = -1
                                            Else
                                                Try
                                                    grasa = arraytext(11)
                                                Catch ex As Exception
                                                    MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: Grasa")
                                                    Exit Sub
                                                End Try
                                            End If
                                            If Trim(arraytext(12)) = "" Or Trim(arraytext(12)) = "-" Then
                                                proteina = -1
                                            Else
                                                Try
                                                    proteina = arraytext(12)
                                                Catch ex As Exception
                                                    MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: Proteína")
                                                    Exit Sub
                                                End Try
                                            End If
                                            If Trim(arraytext(13)) = "" Or Trim(arraytext(13)) = "-" Then
                                                lactosa = -1
                                            Else
                                                Try
                                                    lactosa = arraytext(13)
                                                Catch ex As Exception
                                                    MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: Lactosa")
                                                    Exit Sub
                                                End Try
                                            End If
                                            If Trim(arraytext(14)) = "" Or Trim(arraytext(14)) = "-" Then
                                                st = -1
                                            Else
                                                Try
                                                    st = arraytext(14)
                                                Catch ex As Exception
                                                    MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: Sólidos totales")
                                                    Exit Sub
                                                End Try
                                            End If
                                            If Trim(arraytext(9)) = "" Or Trim(arraytext(9)) = "-" Then
                                                rc = -1
                                            Else
                                                Try
                                                    rc = arraytext(9)
                                                Catch ex As Exception
                                                    MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: RC")
                                                    Exit Sub
                                                End Try
                                            End If
                                            '** IMPORTAR CRIOSCOPIA **************************************************************************
                                            If Trim(arraytext(15)) = "" Or Trim(arraytext(15)) = "-" Then
                                                crioscopia = -1
                                            Else
                                                Try
                                                    crioscopia = arraytext(15)
                                                Catch ex As Exception
                                                    MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: Crioscopía")
                                                    Exit Sub
                                                End Try
                                            End If
                                            '***************************************************************************************************
                                            If Trim(arraytext(16)) = "" Or Trim(arraytext(16)) = "-" Then
                                                urea = -1
                                            Else
                                                Try
                                                    urea = arraytext(16)
                                                Catch ex As Exception
                                                    MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: Urea")
                                                    Exit Sub
                                                End Try
                                            End If
                                            If Trim(arraytext(18)) = "" Or Trim(arraytext(18)) = "-" Then
                                                proteinav = -1
                                            Else
                                                Try
                                                    proteinav = arraytext(18)
                                                Catch ex As Exception
                                                    MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: Proteína verdadera")
                                                    Exit Sub
                                                End Try
                                            End If
                                            If Trim(arraytext(19)) = "" Or Trim(arraytext(19)) = "-" Then
                                                caseina = -1
                                            Else
                                                Try
                                                    caseina = arraytext(19)
                                                Catch ex As Exception
                                                    MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: Caseína")
                                                    Exit Sub
                                                End Try
                                            End If
                                            'If Trim(arraytext(30)) = "" Or Trim(arraytext(30)) = "-" Then
                                            densidad = -1
                                            'Else
                                            'Try
                                            '    densidad = arraytext(30)
                                            'Catch ex As Exception
                                            '    MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: Densidad")
                                            '    Exit Sub
                                            'End Try
                                            'End If
                                            If arraytext.Length > 22 Then
                                                If Trim(arraytext(22)) = "" Or Trim(arraytext(22)) = "-" Then
                                                    ph = -1
                                                Else
                                                    Try
                                                        ph = arraytext(22)
                                                    Catch ex As Exception
                                                        MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: pH")
                                                        Exit Sub
                                                    End Try

                                                End If
                                            End If
                                        End If
                                        'ficha2 = Mid(file.Name, Len(file.Name) - 4, 1)
                                        ficha2 = Mid(file.Name, Len(file.Name) - 21, 1)
                                        ficha3 = Mid(file.Name, 1, 1)
                                        If ficha2 = "a" Or ficha2 = "A" Or ficha2 = "b" Or ficha2 = "B" Or ficha2 = "c" Or ficha2 = "C" Or ficha2 = "d" Or ficha2 = "D" Or ficha2 = "e" Or ficha2 = "E" Or ficha2 = "f" Or ficha2 = "F" Or ficha2 = "g" Or ficha2 = "G" Or ficha2 = "h" Or ficha2 = "H" Or ficha2 = "i" Or ficha2 = "I" Or ficha2 = "j" Or ficha2 = "J" Or ficha2 = "k" Or ficha2 = "K" Then
                                            'ficha = Mid(file.Name, 1, Len(file.Name) - 5)
                                            ficha = Mid(file.Name, 1, Len(file.Name) - 22)
                                        Else
                                            'ficha = Mid(file.Name, 1, Len(file.Name) - 4)
                                            ficha = Mid(file.Name, 1, Len(file.Name) - 21)
                                        End If
                                        If Mid(ficha, 1, 1) = "l" Or Mid(ficha, 1, 1) = "L" Then
                                            Dim MyString As String = ficha
                                            Dim MyChar As Char() = {"l"c, "L"c}
                                            Dim NewString As String = MyString.TrimStart(MyChar)
                                            ficha3 = NewString
                                        Else
                                            ficha3 = ficha
                                        End If
                                        '**CONTROL DE CRIOSCOPIA *************************************************************************
                                        Dim cc As New dCrioscopia_Control
                                        Dim ficha_cc As Long = 0
                                        Dim muestra_cc As String = ""
                                        Dim res_delta As Integer = 0
                                        Dim res_crioscopo As Integer = 0
                                        Dim diferencia_cc As Integer = 0
                                        ficha_cc = ficha3
                                        muestra_cc = matricula
                                        cc.FICHA = ficha_cc
                                        cc.MUESTRA = muestra_cc
                                        cc = cc.buscarxfichaxmuestra
                                        If Not cc Is Nothing Then
                                            res_delta = cc.DELTA
                                            res_crioscopo = cc.CRIOSCOPO
                                            If res_delta > res_crioscopo Then
                                                diferencia_cc = res_delta - res_crioscopo
                                            Else
                                                diferencia_cc = res_crioscopo - res_delta
                                            End If
                                            If diferencia_cc > 5 Then
                                                crioscopia = res_crioscopo
                                            End If
                                        End If
                                        cc = Nothing
                                        ficha_cc = Nothing
                                        muestra_cc = Nothing
                                        res_delta = Nothing
                                        res_crioscopo = Nothing
                                        diferencia_cc = Nothing
                                        '*************************************************************************************************
                                        Dim fechaoriginal As Date = Now()
                                        Dim fecha As String
                                        fecha = Format(fechaoriginal, "yyyy-MM-dd")
                                        c.FICHA = ficha3
                                        c.FECHA = fecha
                                        c.EQUIPO = "delta2"
                                        c.PRODUCTO = producto
                                        c.MUESTRA = matricula
                                        c.RC = rc
                                        c.GRASA = grasa
                                        c.PROTEINA = proteina
                                        c.LACTOSA = lactosa
                                        c.ST = st
                                        c.CRIOSCOPIA = crioscopia
                                        c.UREA = urea
                                        c.PROTEINAV = proteinav
                                        c.CASEINA = caseina
                                        c.DENSIDAD = densidad
                                        c.PH = ph
                                        c.guardar()
                                        Dim csm As New dCalidadSolicitudMuestra
                                        Dim listacsm As New ArrayList
                                        Dim noactualizafecha As Integer = 0
                                        listacsm = csm.listarporsolicitud(ficha3)
                                        If Not listacsm Is Nothing Then
                                            For Each csm In listacsm
                                                If csm.RB = 1 Then
                                                    noactualizafecha = 1
                                                End If
                                            Next
                                        End If
                                        If noactualizafecha = 0 Then
                                            Dim sa As New dSolicitudAnalisis
                                            sa.ID = ficha3
                                            sa.actualizarfechaproceso(fecha)
                                            sa = Nothing
                                        End If
                                    End If
                                End If
                            End If
                            linea = linea + 1
                        Loop Until sLine Is Nothing
                        objReader.Close()
                    End If
                    '*** MOVER ARCHIVO ***********************************************************************
                    Dim sArchivoOrigen As String = "\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Calidad de leche\" & nombrearchivo
                    Dim sRutaDestino As String = "Y:\documentos\secretaria\analisis\leche\bentley-delta\pasados\calidad de leche\pasados NET\" & nombrearchivo
                    Try
                        ' Mover el fichero.si existe lo sobreescribe  
                        My.Computer.FileSystem.MoveFile(sArchivoOrigen, _
                                                        sRutaDestino, _
                                                        True)
                        'MsgBox("Ok.", MsgBoxStyle.Information, "Mover archivo")
                        ' errores  
                    Catch ex As Exception
                        MsgBox(ex.Message.ToString, MsgBoxStyle.Critical)
                    End Try
                    '***********************************
                    'Insert tabla preinforme_calidad
                    Dim pi As New dPreinformes
                    Dim fechaactual As Date = Now()
                    Dim _fecha As String
                    _fecha = Format(fechaactual, "yyyy-MM-dd")
                    pi.FICHA = ficha3
                    pi = pi.buscar
                    If Not pi Is Nothing Then
                    Else
                        Dim pi2 As New dPreinformes
                        pi2.FICHA = ficha3
                        pi2.TIPO = 10
                        pi2.CREADO = 0
                        pi2.FECHA = _fecha
                        pi2.guardar()
                        pi2 = Nothing
                    End If
                    pi = Nothing
                    ' Grabar estado de la ficha
                    Dim est As New dEstados
                    est.FICHA = ficha3
                    est.ESTADO = 4
                    est.FECHA = _fecha
                    est.guardar2()
                    est = Nothing
                    '****************************
                End If 'fin de control archivo delta nuevo
            Next
        End If
    End Sub
    Private Sub controllecherocsv2()
        Dim extension As String
        Dim nombrearchivo As String = ""
        Dim linea As Integer
        Dim folder As New DirectoryInfo("\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Control lechero")
        Dim _ficheros() As String
        _ficheros = Directory.GetFiles("\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Control lechero")
        If Not (_ficheros.Length > 0) Then
        Else
            For Each file As FileInfo In folder.GetFiles("*.csv")
                nombrearchivo = file.Name
                If nombrearchivo.Length > 12 Then 'controlo si el archivo es de delta nuevo
                    linea = 1
                    extension = Microsoft.VisualBasic.Right(file.Name, 3)
                    Dim objReader As New StreamReader("\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Control lechero\" & file.Name)
                    Dim sLine As String = ""
                    Dim arraytext() As String
                    Dim matricula As String = ""
                    Dim grasa As Double = 0
                    Dim proteina As Double = 0
                    Dim lactosa As Double = 0
                    Dim st As Double = 0
                    Dim rc As Integer = 0
                    Dim ficha As String = ""
                    Dim ficha2 As String = ""
                    Dim ficha3 As String = ""
                    Dim equipo As String = ""
                    Dim producto As String = ""
                    Dim crioscopia As Integer = 0
                    Dim urea As Integer = 0
                    Dim proteinav As Double = 0
                    Dim caseina As Double = 0
                    Dim densidad As Double = 0
                    Dim ph As Double = 0
                    Dim grasa_b As Double = 0
                    Dim grasa_a As Double = 0
                    Dim cit As Integer = 0
                    Dim agl As Double = 0
                    Dim sng As Double = 0
                    Dim sfa As Double = 0
                    Dim ufa As Double = 0
                    Dim mufa As Double = 0
                    Dim pufa As Double = 0
                    Dim c16 As Double = 0
                    Dim c180 As Double = 0
                    Dim c181 As Double = 0
                    Dim bhb As Double = 0
                    Dim acetone As Double = 0
                    Dim cisfat As Double = 0
                    Dim transfat As Double = 0
                    Dim denovofa As Double = 0
                    Dim mixedfa As Double = 0
                    Dim preformedfa As Double = 0
                    Dim denovofa2 As Double = 0
                    Dim mixedfa2 As Double = 0
                    Dim preformedfa2 As Double = 0
                    Dim nefa As Double = 0
                    ' *** SI EL ARCHIVO ES CSV **************************************************************************************
                    If extension = "csv" Or extension = "CSV" Then
                        Dim c As New dImpControl()
                        'Dim ca As New dControlAux
                        Dim sa As New dSolicitudAnalisis
                        Dim p As New dCliente
                        Do
                            sLine = objReader.ReadLine()
                            If sLine <> " " Then
                                If linea = 3 Then
                                    arraytext = Split(sLine, ";")
                                    If arraytext.Length < 11 Then
                                        arraytext = Split(sLine, ",")
                                    End If
                                    producto = Trim(arraytext(10))
                                End If
                                If Not sLine Is Nothing Then
                                    If sLine <> ";" And sLine <> ";;" And sLine <> ";;;" And sLine <> ";;;;" And sLine <> ";;;;;" And sLine <> ";;;;;;" And sLine <> ";;;;;;;" And sLine <> ";;;;;;;;" And sLine <> ";;;;;;;;;" And sLine <> ";;;;;;;;;;" And sLine <> ";;;;;;;;;;;" And sLine <> ";;;;;;;;;;;;" And sLine <> ";;;;;;;;;;;;;" And sLine <> ";;;;;;;;;;;;;;" And sLine <> ";;;;;;;;;;;;;;;" And sLine <> ";;;;;;;;;;;;;;;;" And sLine <> ";;;;;;;;;;;;;;;;;" And sLine <> ";;;;;;;;;;;;;;;;;;" And sLine <> ";;;;;;;;;;;;;;;;;;" And sLine <> ";;;;;;;;;;;;;;;;;;;" And sLine <> ";;;;;;;;;;;;;;;;;;;;" And sLine <> ";;;;;;;;;;;;;;;;;;;;;" And sLine <> ";;;;;;;;;;;;;;;;;;;;;;" And sLine <> ";;;;;;;;;;;;;;;;;;;;;;;" And sLine <> ";;;;;;;;;;;;;;;;;;;;;;;;" Then
                                        If linea >= 8 Then
                                            arraytext = Split(sLine, ";")
                                            If arraytext.Length < 39 Then
                                                arraytext = Split(sLine, ",")
                                            End If
                                            matricula = Trim(arraytext(5))
                                            If arraytext.Length <= 13 Then
                                                grasa = -1
                                                proteina = -1
                                                lactosa = -1
                                                st = -1
                                                If Trim(arraytext(11)) <> "" And Trim(arraytext(11)) <> "-" Then
                                                    Try
                                                        rc = arraytext(11)
                                                    Catch ex As Exception
                                                        MsgBox("Error en archivo: " & file.Name & ", línea: " & linea)
                                                    End Try
                                                Else
                                                    rc = -1
                                                End If
                                                crioscopia = -1
                                                urea = -1
                                                proteinav = -1
                                                caseina = -1
                                                densidad = -1
                                                ph = -1
                                            Else
                                                If Trim(arraytext(11)) = "" Or Trim(arraytext(11)) = "-" Then
                                                    grasa = -1
                                                Else
                                                    Try
                                                        grasa = arraytext(11)
                                                    Catch ex As Exception
                                                        MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: Grasa")
                                                        Exit Sub
                                                    End Try
                                                End If
                                                If Trim(arraytext(12)) = "" Or Trim(arraytext(12)) = "-" Then
                                                    proteina = -1
                                                Else
                                                    Try
                                                        proteina = arraytext(12)
                                                    Catch ex As Exception
                                                        MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: Proteína")
                                                        Exit Sub
                                                    End Try
                                                End If
                                                If Trim(arraytext(13)) = "" Or Trim(arraytext(13)) = "-" Then
                                                    lactosa = -1
                                                Else
                                                    Try
                                                        lactosa = arraytext(13)
                                                    Catch ex As Exception
                                                        MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: Lactosa")
                                                        Exit Sub
                                                    End Try
                                                End If
                                                If Trim(arraytext(14)) = "" Or Trim(arraytext(14)) = "-" Then
                                                    st = -1
                                                Else
                                                    Try
                                                        st = arraytext(14)
                                                    Catch ex As Exception
                                                        MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: Sólidos totales")
                                                        Exit Sub
                                                    End Try
                                                End If
                                                If Trim(arraytext(9)) = "" Or Trim(arraytext(9)) = "-" Then
                                                    rc = -1
                                                Else
                                                    Try
                                                        rc = arraytext(9)
                                                    Catch ex As Exception
                                                        MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: RC")
                                                        Exit Sub
                                                    End Try
                                                End If
                                                '** IMPORTAR CRIOSCOPIA **************************************************************************
                                                If Trim(arraytext(15)) = "" Or Trim(arraytext(15)) = "-" Then
                                                    crioscopia = -1
                                                Else
                                                    Try
                                                        crioscopia = arraytext(15)
                                                    Catch ex As Exception
                                                        MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: Crioscopía")
                                                        Exit Sub
                                                    End Try
                                                End If
                                                '***************************************************************************************************
                                                If Trim(arraytext(16)) = "" Or Trim(arraytext(16)) = "-" Then
                                                    urea = -1
                                                Else
                                                    Try
                                                        urea = arraytext(16)
                                                    Catch ex As Exception
                                                        MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: Urea")
                                                        Exit Sub
                                                    End Try
                                                End If
                                                If Trim(arraytext(18)) = "" Or Trim(arraytext(18)) = "-" Then
                                                    proteinav = -1
                                                Else
                                                    Try
                                                        proteinav = arraytext(18)
                                                    Catch ex As Exception
                                                        MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: Proteína verdadera")
                                                        Exit Sub
                                                    End Try
                                                End If
                                                If Trim(arraytext(19)) = "" Or Trim(arraytext(19)) = "-" Then
                                                    caseina = -1
                                                Else
                                                    Try
                                                        caseina = arraytext(19)
                                                    Catch ex As Exception
                                                        MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: Caseína")
                                                        Exit Sub
                                                    End Try
                                                End If
                                                'If Trim(arraytext(30)) = "" Or Trim(arraytext(30)) = "-" Then
                                                densidad = -1
                                                '    Else
                                                '    Try
                                                '        densidad = arraytext(30)
                                                '    Catch ex As Exception
                                                '        MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: Densidad")
                                                '        Exit Sub
                                                '    End Try
                                                'End If
                                                If Trim(arraytext(22)) = "" Or Trim(arraytext(22)) = "-" Then
                                                    ph = -1
                                                Else
                                                    Try
                                                        ph = arraytext(22)
                                                    Catch ex As Exception
                                                        MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: pH")
                                                        Exit Sub
                                                    End Try
                                                End If
                                                If Trim(arraytext(30)) = "" Or Trim(arraytext(30)) = "-" Then
                                                    bhb = -1
                                                Else
                                                    Try
                                                        bhb = arraytext(30)
                                                    Catch ex As Exception
                                                        MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: BHB")
                                                        Exit Sub
                                                    End Try
                                                End If
                                            End If
                                            'ficha2 = Mid(file.Name, Len(file.Name) - 4, 1)
                                            ficha2 = Mid(file.Name, Len(file.Name) - 21, 1)
                                            ficha3 = Mid(file.Name, 1, 1)
                                            If ficha2 = "a" Or ficha2 = "A" Or ficha2 = "b" Or ficha2 = "B" Or ficha2 = "c" Or ficha2 = "C" Or ficha2 = "d" Or ficha2 = "D" Or ficha2 = "e" Or ficha2 = "E" Or ficha2 = "f" Or ficha2 = "F" Or ficha2 = "g" Or ficha2 = "G" Or ficha2 = "h" Or ficha2 = "H" Or ficha2 = "i" Or ficha2 = "I" Or ficha2 = "j" Or ficha2 = "J" Or ficha2 = "k" Or ficha2 = "K" Then
                                                'ficha = Mid(file.Name, 1, Len(file.Name) - 5)
                                                ficha = Mid(file.Name, 1, Len(file.Name) - 22)
                                            Else
                                                ficha = Mid(file.Name, 1, Len(file.Name) - 21)
                                            End If
                                            If Mid(ficha, 1, 1) = "l" Or Mid(ficha, 1, 1) = "L" Then
                                                Dim MyString As String = ficha
                                                Dim MyChar As Char() = {"l"c, "L"c}
                                                Dim NewString As String = MyString.TrimStart(MyChar)
                                                ficha3 = NewString
                                            Else
                                                ficha3 = ficha
                                            End If
                                            '**CONTROL DE CRIOSCOPIA *************************************************************************
                                            Dim cc As New dCrioscopia_Control
                                            Dim ficha_cc As Long = 0
                                            Dim muestra_cc As String = ""
                                            Dim res_delta As Integer = 0
                                            Dim res_crioscopo As Integer = 0
                                            Dim diferencia_cc As Integer = 0
                                            ficha_cc = ficha3
                                            muestra_cc = matricula
                                            cc.FICHA = ficha_cc
                                            cc.MUESTRA = muestra_cc
                                            cc = cc.buscarxfichaxmuestra
                                            If Not cc Is Nothing Then
                                                res_delta = cc.DELTA
                                                res_crioscopo = cc.CRIOSCOPO
                                                If res_delta > res_crioscopo Then
                                                    diferencia_cc = res_delta - res_crioscopo
                                                Else
                                                    diferencia_cc = res_crioscopo - res_delta
                                                End If
                                                If diferencia_cc > 5 Then
                                                    crioscopia = res_crioscopo
                                                End If
                                            End If
                                            cc = Nothing
                                            ficha_cc = Nothing
                                            muestra_cc = Nothing
                                            res_delta = Nothing
                                            res_crioscopo = Nothing
                                            diferencia_cc = Nothing
                                            '*************************************************************************************************
                                            Dim fechaoriginal As Date = Now()
                                            Dim fecha As String
                                            fecha = Format(fechaoriginal, "yyyy-MM-dd")
                                            c.FICHA = ficha3
                                            c.FECHA = fecha
                                            c.EQUIPO = "delta"
                                            c.PRODUCTO = producto
                                            c.MUESTRA = matricula
                                            c.RC = rc
                                            c.GRASA = grasa
                                            c.PROTEINA = proteina
                                            c.LACTOSA = lactosa
                                            c.ST = st
                                            c.CRIOSCOPIA = crioscopia
                                            c.UREA = urea
                                            c.PROTEINAV = proteinav
                                            c.CASEINA = caseina
                                            c.DENSIDAD = densidad
                                            c.PH = ph
                                            c.BHB = bhb
                                            c.guardar()
                                            'ca.FICHA = ficha3
                                            'ca.FECHA = fecha
                                            'sa.ID = ficha3
                                            'sa = sa.buscar
                                            'ca.PRODUCTOR = sa.IDPRODUCTOR
                                            'ca.MUESTRA = matricula
                                            'ca.RC = rc
                                            'ca.TAMBO = sa.TAMBO
                                            'ca.guardar()
                                            Dim sa2 As New dSolicitudAnalisis
                                            sa2.ID = ficha3
                                            sa2.actualizarfechaproceso(fecha)
                                            sa2 = Nothing
                                        End If
                                    End If
                                End If
                            End If
                            linea = linea + 1
                        Loop Until sLine Is Nothing
                        objReader.Close()
                    End If
                    '*** MOVER ARCHIVO ***********************************************************************
                    Dim sArchivoOrigen As String = "\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Control lechero\" & nombrearchivo
                    Dim sRutaDestino As String = "Y:\documentos\secretaria\analisis\leche\bentley-delta\pasados\control lechero\pasados NET\" & nombrearchivo
                    Try
                        ' Mover el fichero.si existe lo sobreescribe  
                        My.Computer.FileSystem.MoveFile(sArchivoOrigen, _
                                                        sRutaDestino, _
                                                        True)
                        'MsgBox("Ok.", MsgBoxStyle.Information, "Mover archivo")
                        ' errores  
                    Catch ex As Exception
                        MsgBox(ex.Message.ToString, MsgBoxStyle.Critical)
                    End Try
                    '***********************************
                    'Insert tabla preinforme_calidad
                    Dim pi As New dPreinformes
                    Dim fechaactual As Date = Now()
                    Dim _fecha As String
                    _fecha = Format(fechaactual, "yyyy-MM-dd")
                    pi.FICHA = ficha3
                    pi = pi.buscar
                    If Not pi Is Nothing Then
                    Else
                        Dim pi2 As New dPreinformes
                        pi2.FICHA = ficha3
                        pi2.TIPO = 1
                        pi2.CREADO = 0
                        pi2.FECHA = _fecha
                        pi2.guardar()
                        pi2 = Nothing
                    End If
                    pi = Nothing
                    ' Grabar estado de la ficha
                    Dim est As New dEstados
                    est.FICHA = ficha3
                    est.ESTADO = 4
                    est.FECHA = _fecha
                    est.guardar2()
                    est = Nothing
                    '****************************
                End If 'fin de control si archivoes de delta 400
            Next
        End If
    End Sub
    Private Sub calidadxls()
        Dim extension As String
        Dim nombrearchivo As String = ""
        Dim linea As Integer
        'Dim folder As New DirectoryInfo("d:\documentos\secretaria\analisis\leche\bentley-delta\pasados\calidad de leche")
        Dim folder As New DirectoryInfo("\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Calidad de leche")
        'Dim folder As New DirectoryInfo("\\192.168.1.10\e\documentos\secretaria\resultados informes 2012")
        ' For Each file As FileInfo In folder.GetFiles("*.xls, *.CSV, *.fat")
        For Each file As FileInfo In folder.GetFiles("*.xls")
            'ListBox1.Items.Add(file.Name)
            nombrearchivo = file.Name
            linea = 1
            extension = Microsoft.VisualBasic.Right(file.Name, 3)
            'Dim objReader As New StreamReader("d:\documentos\secretaria\analisis\leche\bentley-delta\pasados\calidad de leche\" & file.Name)
            Dim objReader As New StreamReader("\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Calidad de leche\" & file.Name)
            Dim sLine As String = ""
            'Dim arrText As New ArrayList()
            ''Dim arraytext() As String
            Dim matricula As String = ""
            Dim grasa As Double = 0
            Dim proteina As Double = 0
            Dim lactosa As Double = 0
            Dim st As Double = 0
            Dim rc As Integer = 0
            Dim ficha As String = ""
            Dim ficha2 As String = ""
            Dim ficha3 As String = ""
            Dim equipo As String = ""
            Dim producto As String = ""
            Dim crioscopia As Integer = 0
            Dim urea As Integer = 0
            Dim proteinav As Double = 0
            Dim caseina As Double = 0
            Dim densidad As Double = 0
            Dim ph As Double = 0
            ' *** SI EL ARCHIVO ES XLS **************************************************************************************
            If extension = "xls" Or extension = "XLS" Then
                Dim c As New dImpCalidad()
                Dim Arch As String, CantFilas As Integer
                Arch = "\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Calidad de leche\" & file.Name
                Dim x1app As Microsoft.Office.Interop.Excel.Application
                Dim x1libro As Microsoft.Office.Interop.Excel.Workbook
                Dim x1hoja As Microsoft.Office.Interop.Excel.Worksheet
                x1app = CType(CreateObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)
                x1libro = CType(x1app.Workbooks.Open(Arch), Microsoft.Office.Interop.Excel.Workbook)
                x1hoja = CType(x1libro.Worksheets(1), Microsoft.Office.Interop.Excel.Worksheet)
                Dim bandera As Integer = 0
                CantFilas = x1app.Range("a1").CurrentRegion.Rows.Count
                ficha2 = Mid(file.Name, Len(file.Name) - 4, 1)
                ficha3 = Mid(file.Name, 1, 1)
                If ficha2 = "a" Or ficha2 = "A" Or ficha2 = "b" Or ficha2 = "B" Or ficha2 = "c" Or ficha2 = "C" Or ficha2 = "d" Or ficha2 = "D" Then
                    ficha = Mid(file.Name, 1, Len(file.Name) - 5)
                Else
                    ficha = Mid(file.Name, 1, Len(file.Name) - 4)
                End If
                If Mid(ficha, 1, 1) = "l" Or Mid(ficha, 1, 1) = "L" Then
                    Dim MyString As String = ficha
                    Dim MyChar As Char() = {"l"c, "L"c}
                    Dim NewString As String = MyString.TrimStart(MyChar)
                    ficha3 = NewString
                Else
                    ficha3 = ficha
                End If
                Dim fechaoriginal As Date = Now()
                Dim fecha As String
                fecha = Format(fechaoriginal, "yyyy-MM-dd")
                For i = 1 To CantFilas
                    If Trim(x1hoja.Cells(i, 2).formula) <> "" Then
                        matricula = Trim(x1hoja.Cells(i, 2).value)
                    Else
                        matricula = -1
                    End If
                    If Trim(x1hoja.Cells(i, 3).formula) <> "" Then
                        Try
                            grasa = x1hoja.Cells(i, 3).value
                        Catch ex As Exception
                            MsgBox("Error en archivo: " & file.Name & ", línea: " & i & ", valor: Grasa")
                            Exit Sub
                        End Try
                    Else
                        grasa = -1
                    End If
                    If Trim(x1hoja.Cells(i, 4).formula) <> "" Then
                        Try
                            proteina = x1hoja.Cells(i, 4).value
                            bandera = 1
                        Catch ex As Exception
                            MsgBox("Error en archivo: " & file.Name & ", línea: " & i & ", valor: Proteína")
                            Exit Sub
                        End Try
                    Else
                        proteina = -1
                    End If
                    If Trim(x1hoja.Cells(i, 5).formula) <> "" Then
                        Try
                            lactosa = x1hoja.Cells(i, 5).value
                        Catch ex As Exception
                            MsgBox("Error en archivo: " & file.Name & ", línea: " & i & ", valor: Lactosa")
                            Exit Sub
                        End Try
                    Else
                        lactosa = -1
                    End If
                    If Trim(x1hoja.Cells(i, 6).formula) <> "" Then
                        Try
                            st = x1hoja.Cells(i, 6).value
                        Catch ex As Exception
                            MsgBox("Error en archivo: " & file.Name & ", línea: " & i & ", valor: Sólidos totales")
                            Exit Sub
                        End Try
                    Else
                        st = -1
                    End If
                    If Trim(x1hoja.Cells(i, 7).formula) <> "" Then
                        Try
                            rc = x1hoja.Cells(i, 7).value
                        Catch ex As Exception
                            MsgBox("Error en archivo: " & file.Name & ", línea: " & i & ", valor: RC")
                            Exit Sub
                        End Try
                    Else
                        rc = -1
                    End If
                    If bandera = 0 Then
                        c.FICHA = ficha3
                        c.FECHA = fecha
                        c.EQUIPO = "bentley"
                        c.PRODUCTO = "leche"
                        c.MUESTRA = matricula
                        c.RC = grasa
                        c.GRASA = -1
                        c.PROTEINA = proteina
                        c.LACTOSA = lactosa
                        c.ST = st
                        c.CRIOSCOPIA = -1
                        c.UREA = -1
                        c.PROTEINAV = -1
                        c.CASEINA = -1
                        c.DENSIDAD = -1
                        c.PH = -1
                        c.guardar()

                        Dim csm As New dCalidadSolicitudMuestra
                        Dim listacsm As New ArrayList
                        Dim noactualizafecha As Integer = 0
                        listacsm = csm.listarporsolicitud(ficha3)
                        If Not listacsm Is Nothing Then
                            For Each csm In listacsm
                                If csm.RB = 1 Then
                                    noactualizafecha = 1
                                End If
                            Next
                        End If
                        If noactualizafecha = 0 Then
                            Dim sa As New dSolicitudAnalisis
                            sa.ID = ficha3
                            sa.actualizarfechaproceso(fecha)
                            sa = Nothing
                        End If

                    ElseIf bandera = 1 Then
                        c.FICHA = ficha3
                        c.FECHA = fecha
                        c.EQUIPO = "bentley"
                        c.PRODUCTO = "leche"
                        c.MUESTRA = matricula
                        c.RC = rc
                        c.GRASA = grasa
                        c.PROTEINA = proteina
                        c.LACTOSA = lactosa
                        c.ST = st
                        c.CRIOSCOPIA = -1
                        c.UREA = -1
                        c.PROTEINAV = -1
                        c.CASEINA = -1
                        c.DENSIDAD = -1
                        c.PH = -1
                        c.guardar()

                        Dim csm As New dCalidadSolicitudMuestra
                        Dim listacsm As New ArrayList
                        Dim noactualizafecha As Integer = 0
                        listacsm = csm.listarporsolicitud(ficha3)
                        If Not listacsm Is Nothing Then
                            For Each csm In listacsm
                                If csm.RB = 1 Then
                                    noactualizafecha = 1
                                End If
                            Next
                        End If
                        If noactualizafecha = 0 Then
                            Dim sa As New dSolicitudAnalisis
                            sa.ID = ficha3
                            sa.actualizarfechaproceso(fecha)
                            sa = Nothing
                        End If

                    End If
                Next
                ' Cierro Excel
                x1libro.Close()
                x1app = Nothing
                x1libro = Nothing
                x1hoja = Nothing
                objReader.Close()
                Dim proceso As System.Diagnostics.Process()
                proceso = System.Diagnostics.Process.GetProcessesByName("EXCEL")
                For Each opro As System.Diagnostics.Process In proceso
                    'antes de iniciar el proceso obtengo la fecha en que inicie el 
                    'proceso para detener todos los procesos que excel que inicio
                    'mi código durante el proceso
                    opro.Kill()
                Next
            End If
            '*** MOVER ARCHIVO ***********************************************************************
            'Dim sArchivoOrigen As String = "d:\documentos\secretaria\analisis\leche\bentley-delta\pasados\calidad de leche\" & nombrearchivo
            Dim sArchivoOrigen As String = "\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Calidad de leche\" & nombrearchivo
            'Dim sRutaDestino As String = "d:\documentos\secretaria\analisis\leche\bentley-delta\pasados\calidad de leche\pasados NET\" & nombrearchivo
            Dim sRutaDestino As String = "Y:\documentos\secretaria\analisis\leche\bentley-delta\pasados\calidad de leche\pasados NET\" & nombrearchivo
            Try
                ' Mover el fichero.si existe lo sobreescribe  
                My.Computer.FileSystem.MoveFile(sArchivoOrigen, _
                                                sRutaDestino, _
                                                True)
                'MsgBox("Ok.", MsgBoxStyle.Information, "Mover archivo")
                ' errores  
            Catch ex As Exception
                MsgBox(ex.Message.ToString, MsgBoxStyle.Critical)
            End Try
            '***********************************
            'Insert tabla preinforme_calidad
            Dim pi As New dPreinformes
            Dim fechaactual As Date = Now()
            Dim _fecha As String
            _fecha = Format(fechaactual, "yyyy-MM-dd")
            pi.FICHA = ficha3
            pi = pi.buscar
            If Not pi Is Nothing Then
            Else
                Dim pi2 As New dPreinformes
                Dim sa As New dSolicitudAnalisis
                Dim tipof As Integer = 0
                sa.ID = ficha3
                sa = sa.buscar
                If Not sa Is Nothing Then
                    tipof = sa.IDTIPOINFORME
                End If
                pi2.FICHA = ficha3
                pi2.TIPO = tipof
                pi2.CREADO = 0
                pi2.FECHA = _fecha
                pi2.guardar()
                pi2 = Nothing
                sa = Nothing
            End If
            pi = Nothing
            ' Grabar estado de la ficha
            Dim est As New dEstados
            est.FICHA = ficha3
            est.ESTADO = 4
            est.FECHA = _fecha
            est.guardar2()
            est = Nothing
        Next
    End Sub
    Private Sub calidadfat()
        Dim extension As String
        Dim nombrearchivo As String = ""
        Dim linea As Integer
        'Dim folder As New DirectoryInfo("d:\documentos\secretaria\analisis\leche\bentley-delta\pasados\calidad de leche")
        Dim folder As New DirectoryInfo("\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Calidad de leche")
        'Dim folder As New DirectoryInfo("\\192.168.1.10\e\documentos\secretaria\resultados informes 2012")
        ' For Each file As FileInfo In folder.GetFiles("*.xls, *.CSV, *.fat")
        For Each file As FileInfo In folder.GetFiles("*.fat")
            'ListBox1.Items.Add(file.Name)
            nombrearchivo = file.Name
            linea = 1
            extension = Microsoft.VisualBasic.Right(file.Name, 3)
            'Dim objReader As New StreamReader("d:\documentos\secretaria\analisis\leche\bentley-delta\pasados\calidad de leche\" & file.Name)
            Dim objReader As New StreamReader("\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Calidad de leche\" & file.Name)
            Dim sLine As String = ""
            'Dim arrText As New ArrayList()
            'Dim arraytext() As String
            Dim matricula As String = ""
            Dim grasa As Double = 0
            Dim proteina As Double = 0
            Dim lactosa As Double = 0
            Dim st As Double = 0
            Dim rc As Integer = 0
            Dim ficha As String = ""
            Dim ficha2 As String = ""
            Dim ficha3 As String = ""
            Dim equipo As String = ""
            Dim producto As String = ""
            Dim crioscopia As Integer = 0
            Dim urea As Integer = 0
            Dim proteinav As Double = 0
            Dim caseina As Double = 0
            Dim densidad As Double = 0
            Dim ph As Double = 0
            ' *** SI EL ARCHIVO ES FAT **************************************************************************************
            If extension = "fat" Or extension = "FAT" Then
                Dim c As New dImpCalidad()
                Dim cuentalinea As Long = 1
                Do
                    sLine = objReader.ReadLine()
                    If Not sLine Is Nothing Then
                        Dim Texto As String
                        Dim id As Integer = 0
                        ficha2 = Mid(file.Name, Len(file.Name) - 4, 1)
                        ficha3 = Mid(file.Name, 1, 1)
                        If ficha2 = "a" Or ficha2 = "A" Or ficha2 = "b" Or ficha2 = "B" Or ficha2 = "c" Or ficha2 = "C" Or ficha2 = "d" Or ficha2 = "D" Then
                            ficha = Mid(file.Name, 1, Len(file.Name) - 5)
                        Else
                            ficha = Mid(file.Name, 1, Len(file.Name) - 4)
                        End If
                        If Mid(ficha, 1, 1) = "l" Or Mid(ficha, 1, 1) = "L" Then
                            Dim MyString As String = ficha
                            Dim MyChar As Char() = {"l"c, "L"c}
                            Dim NewString As String = MyString.TrimStart(MyChar)
                            ficha3 = NewString
                        Else
                            ficha3 = ficha
                        End If
                        Dim fechaoriginal As Date = Now()
                        Dim fecha As String
                        fecha = Format(fechaoriginal, "yyyy-MM-dd")
                        Texto = sLine
                        id = Trim(Mid(Texto, 1, 8))
                        matricula = Trim(Mid(Texto, 9, 9))
                        If Trim(Mid(Texto, 18, 9)) <> "" Then
                            Try
                                grasa = Trim(Mid(Texto, 18, 9))
                            Catch ex As Exception
                                MsgBox("Error en archivo: " & file.Name & ", línea: " & cuentalinea & ", valor: Grasa")
                                Exit Sub
                            End Try
                        Else
                            grasa = -1
                        End If
                        If Trim(Mid(Texto, 27, 9)) <> "" Then
                            Try
                                proteina = Trim(Mid(Texto, 27, 9))
                            Catch ex As Exception
                                MsgBox("Error en archivo: " & file.Name & ", línea: " & cuentalinea & ", valor: Proteína")
                                Exit Sub
                            End Try
                        Else
                            proteina = -1
                        End If
                        If Trim(Mid(Texto, 36, 9)) <> "" Then
                            Try
                                lactosa = Trim(Mid(Texto, 36, 9))
                            Catch ex As Exception
                                MsgBox("Error en archivo: " & file.Name & ", línea: " & cuentalinea & ", valor: Lactosa")
                                Exit Sub
                            End Try
                        Else
                            lactosa = -1
                        End If
                        If Trim(Mid(Texto, 45, 9)) <> "" Then
                            Try
                                st = Trim(Mid(Texto, 45, 9))
                            Catch ex As Exception
                                MsgBox("Error en archivo: " & file.Name & ", línea: " & cuentalinea & ", valor: Sólidos totales")
                                Exit Sub
                            End Try
                        Else
                            st = -1
                        End If
                        If Trim(Mid(Texto, 54, 10)) <> "" Then
                            Try
                                rc = Trim(Mid(Texto, 54, 10))
                            Catch ex As Exception
                                MsgBox("Error en archivo: " & file.Name & ", línea: " & cuentalinea & ", valor: RC")
                                Exit Sub
                            End Try
                        Else
                            rc = -1
                        End If
                        c.FICHA = ficha3
                        c.FECHA = fecha
                        c.EQUIPO = "bentley"
                        c.PRODUCTO = "leche"
                        c.MUESTRA = matricula
                        c.RC = rc
                        c.GRASA = grasa
                        c.PROTEINA = proteina
                        c.LACTOSA = lactosa
                        c.ST = st
                        c.CRIOSCOPIA = -1
                        c.UREA = -1
                        c.PROTEINAV = -1
                        c.CASEINA = -1
                        c.DENSIDAD = -1
                        c.PH = -1
                        c.guardar()
                        cuentalinea = cuentalinea + 1
                        Dim csm As New dCalidadSolicitudMuestra
                        Dim listacsm As New ArrayList
                        Dim noactualizafecha As Integer = 0
                        listacsm = csm.listarporsolicitud(ficha3)
                        If Not listacsm Is Nothing Then
                            For Each csm In listacsm
                                If csm.RB = 1 Then
                                    noactualizafecha = 1
                                End If
                            Next
                        End If
                        If noactualizafecha = 0 Then
                            Dim sa As New dSolicitudAnalisis
                            sa.ID = ficha3
                            sa.actualizarfechaproceso(fecha)
                            sa = Nothing
                        End If
                    End If
                Loop Until sLine Is Nothing
                objReader.Close()
            End If
            '*** MOVER ARCHIVO ***********************************************************************
            'Dim sArchivoOrigen As String = "d:\documentos\secretaria\analisis\leche\bentley-delta\pasados\calidad de leche\" & nombrearchivo
            Dim sArchivoOrigen As String = "\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Calidad de leche\" & nombrearchivo
            'Dim sRutaDestino As String = "d:\documentos\secretaria\analisis\leche\bentley-delta\pasados\calidad de leche\pasados NET\" & nombrearchivo
            Dim sRutaDestino As String = "Y:\documentos\secretaria\analisis\leche\bentley-delta\pasados\calidad de leche\pasados NET\" & nombrearchivo
            Try
                ' Mover el fichero.si existe lo sobreescribe  
                My.Computer.FileSystem.MoveFile(sArchivoOrigen, _
                                                sRutaDestino, _
                                                True)
                'MsgBox("Ok.", MsgBoxStyle.Information, "Mover archivo")
                ' errores  
            Catch ex As Exception
                MsgBox(ex.Message.ToString, MsgBoxStyle.Critical)
            End Try
            '***********************************
            'Insert tabla preinforme_calidad
            Dim pi As New dPreinformes
            Dim fechaactual As Date = Now()
            Dim _fecha As String
            _fecha = Format(fechaactual, "yyyy-MM-dd")
            pi.FICHA = ficha3
            pi = pi.buscar
            If Not pi Is Nothing Then
            Else
                Dim pi2 As New dPreinformes
                Dim sa As New dSolicitudAnalisis
                Dim tipof As Integer = 0
                sa.ID = ficha3
                sa = sa.buscar
                If Not sa Is Nothing Then
                    tipof = sa.IDTIPOINFORME
                End If
                pi2.FICHA = ficha3
                pi2.TIPO = tipof
                pi2.CREADO = 0
                pi2.FECHA = _fecha
                pi2.guardar()
                pi2 = Nothing
                sa = Nothing
            End If
            pi = Nothing
            ' Grabar estado de la ficha
            Dim est As New dEstados
            est.FICHA = ficha3
            est.ESTADO = 4
            est.FECHA = _fecha
            est.guardar2()
            est = Nothing
            '**********************************
        Next
    End Sub
    Private Sub ibc()
        Dim extension As String
        Dim nombrearchivo As String = ""
        Dim linea As Integer
        Dim ficha As String = ""
        Dim ficha2 As String = ""
        Dim ficha3 As String = "0"
        Dim fecha2 As String = ""
        '**********************************************************************************
        Dim El_Ping As Boolean
        Dim eco As New System.Net.NetworkInformation.Ping
        Dim res As System.Net.NetworkInformation.PingReply
        Dim ip As Net.IPAddress
        ip = Net.IPAddress.Parse("192.168.1.50")
        res = eco.Send(ip)
        If res.Status = System.Net.NetworkInformation.IPStatus.Success Then
            El_Ping = (My.Computer.Network.Ping("ibc1123"))
        End If
        'Acá mandamos los mensajes para las 2 posibilidades
        If El_Ping = False Then
            'si no se pudo acceder ,avisamos
            'MsgBox("El servidor no está disponible.", MsgBoxStyle.Critical, "Error")
        Else
            'MsgBox("Servidor disponible.", MsgBoxStyle.Information, "Aviso")
            Dim folder As New DirectoryInfo("\\Ibc1123\Carol")
            For Each file As FileInfo In folder.GetFiles("*.csv")
                'ListBox1.Items.Add(file.Name)
                nombrearchivo = file.Name
                linea = 1
                extension = Microsoft.VisualBasic.Right(file.Name, 3)
                Dim objReader2 As New StreamReader("\\Ibc1123\Carol\" & file.Name)
                Dim sLine As String = ""
                Dim arraytext() As String
                Dim muestra As String = ""
                Dim idibc As Integer = 0
                Dim ibc As Long = 0
                Dim rb As Integer = 0
                ' *** SI EL ARCHIVO ES CSV **************************************************************************************
                If extension = "csv" Or extension = "CSV" Then
                    Dim c As New dImpIbc()
                    Do
                        sLine = objReader2.ReadLine()
                        If Not sLine Is Nothing Then
                            arraytext = Split(sLine, ",")
                            Dim muestra2 As String
                            Dim muestrax As String
                            If Trim(arraytext(1)) <> "" Then
                                muestra = arraytext(1)
                                muestrax = Replace(muestra, Chr(34), "")
                                If muestrax <> "" Then
                                    muestra2 = muestrax
                                Else
                                    muestra = arraytext(7)
                                    muestrax = Replace(muestra, Chr(34), "")
                                    If muestrax <> "" Then
                                        muestra2 = muestrax
                                    Else
                                        muestra2 = "error"
                                    End If
                                End If
                            Else
                                If arraytext.Length > 7 Then
                                    muestra = arraytext(7)
                                    muestrax = Replace(muestra, Chr(34), "")
                                    If muestrax <> "" Then
                                        muestra2 = muestrax
                                    Else
                                        muestra2 = "error"
                                    End If
                                Else
                                    muestra2 = "error"
                                End If
                            End If
                            If Trim(arraytext(2)) <> "" Then
                                idibc = arraytext(2)
                            Else
                                idibc = -1
                            End If
                            If Trim(arraytext(4)) <> "" Then
                                ibc = arraytext(4)
                            Else
                                ibc = -1
                            End If
                            If Trim(arraytext(5)) <> "" Then
                                rb = arraytext(5)
                            Else
                                rb = -1
                            End If
                            ficha2 = Mid(file.Name, Len(file.Name) - 4, 1)
                            ficha3 = Mid(file.Name, 1, 1)
                            If ficha2 = "a" Or ficha2 = "A" Or ficha2 = "b" Or ficha2 = "B" Or ficha2 = "c" Or ficha2 = "C" Or ficha2 = "d" Or ficha2 = "D" Then
                                ficha = Mid(file.Name, 1, Len(file.Name) - 5)
                            Else
                                ficha = Mid(file.Name, 1, Len(file.Name) - 4)
                            End If
                            If Mid(ficha, 1, 1) = "l" Or Mid(ficha, 1, 1) = "L" Then
                                Dim MyString As String = ficha
                                Dim MyChar As Char() = {"l"c, "L"c}
                                Dim NewString As String = MyString.TrimStart(MyChar)
                                ficha3 = NewString
                            Else
                                ficha3 = ficha
                            End If
                            Dim fechaoriginal As Date = Now()
                            Dim fecha As String
                            fecha = Format(fechaoriginal, "yyyy-MM-dd")
                            fecha2 = Format(fechaoriginal, "yyyy-MM-dd")
                            c.FICHA = ficha3
                            c.MUESTRA = muestra2
                            c.IDIBC = idibc
                            c.IBC = ibc
                            c.RB = rb
                            c.FECHA = fecha
                            c.guardar()
                            Dim sa As New dSolicitudAnalisis
                            sa.ID = ficha3
                            sa.actualizarfechaproceso(fecha)
                            sa = Nothing
                        End If
                        linea = linea + 1
                    Loop Until sLine Is Nothing
                    objReader2.Close()
                    ' Grabar estado de la ficha
                    Dim est As New dEstados
                    est.FICHA = ficha3
                    est.ESTADO = 3
                    est.FECHA = fecha2
                    est.guardar2()
                    est = Nothing

                    Dim sa2 As New dSolicitudAnalisis
                    Dim pi2 As New dPreinformes
                    Dim tipof As Integer = 0
                    sa2.ID = ficha3
                    sa2 = sa2.buscar
                    If Not sa2 Is Nothing Then
                        tipof = sa2.IDTIPOINFORME
                    End If
                    pi2.FICHA = ficha3
                    pi2.TIPO = tipof
                    pi2.CREADO = 0
                    pi2.FECHA = fecha2
                    pi2.guardar()
                    pi2 = Nothing
                    sa2 = Nothing
                    '****************************

                End If
                '*** MOVER ARCHIVO ***********************************************************************
                Dim sArchivoOrigen As String = "\\Ibc1123\Carol\" & nombrearchivo
                'Dim sRutaDestino1 As String = "d:\documentos\secretaria\analisis\leche\ibc\" & nombrearchivo
                Dim sRutaDestino1 As String = "Y:\documentos\secretaria\analisis\leche\ibc\" & nombrearchivo
                Dim sRutaDestino As String = "\\Ibc1123\Carol\pasados\" & nombrearchivo
                Try
                    ' Mover el fichero.si existe lo sobreescribe  
                    My.Computer.FileSystem.CopyFile(sArchivoOrigen, _
                                                    sRutaDestino1, _
                                                    True)
                    My.Computer.FileSystem.MoveFile(sArchivoOrigen, _
                                                    sRutaDestino, _
                                                    True)
                    'MsgBox("Ok.", MsgBoxStyle.Information, "Mover archivo")
                    ' errores  
                Catch ex As Exception
                    MsgBox(ex.Message.ToString, MsgBoxStyle.Critical)
                End Try
            Next
            '***********************************
            'Insert tabla preinforme_calidad
            Dim pi As New dPreinformes
            Dim fechaactual As Date = Now()
            Dim _fecha As String
            _fecha = Format(fechaactual, "yyyy-MM-dd")
            pi.FICHA = ficha3
            pi = pi.buscar
            If Not pi Is Nothing Then
            Else
                Dim pi2 As New dPreinformes
                pi2.FICHA = ficha3
                pi2.TIPO = 10
                pi2.CREADO = 0
                pi2.FECHA = _fecha
                pi2.guardar()
                pi2 = Nothing
            End If
            pi = Nothing
        End If
        '***********************************************************************************
    End Sub
    Private Sub controllecherocsv()
        Dim extension As String
        Dim nombrearchivo As String = ""
        Dim linea As Integer
        Dim folder As New DirectoryInfo("\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Control lechero")
        For Each file As FileInfo In folder.GetFiles("*.csv")
            nombrearchivo = file.Name
            If nombrearchivo.Length < 12 Then 'controlo si el archivo es de delta nuevo
                linea = 1
                extension = Microsoft.VisualBasic.Right(file.Name, 3)
                Dim objReader As New StreamReader("\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Control lechero\" & file.Name)
                Dim sLine As String = ""
                Dim arraytext() As String
                Dim matricula As String = ""
                Dim grasa As Double = 0
                Dim proteina As Double = 0
                Dim lactosa As Double = 0
                Dim st As Double = 0
                Dim rc As Integer = 0
                Dim ficha As String = ""
                Dim ficha2 As String = ""
                Dim ficha3 As String = ""
                Dim equipo As String = ""
                Dim producto As String = ""
                Dim crioscopia As Integer = 0
                Dim urea As Integer = 0
                Dim proteinav As Double = 0
                Dim caseina As Double = 0
                Dim densidad As Double = 0
                Dim ph As Double = 0
                ' *** SI EL ARCHIVO ES CSV **************************************************************************************
                If extension = "csv" Or extension = "CSV" Then
                    Dim c As New dImpControl()
                    'Dim ca As New dControlAux
                    Dim sa As New dSolicitudAnalisis
                    Dim p As New dCliente
                    Do
                        sLine = objReader.ReadLine()
                        If linea = 3 Then
                            arraytext = Split(sLine, ";")
                            If arraytext.Length < 11 Then
                                arraytext = Split(sLine, ",")
                            End If
                            If Trim(arraytext(10)) <> "" Then
                                producto = Trim(arraytext(10))
                            End If
                        End If
                        If Not sLine Is Nothing Then
                            If Mid(sLine, 1, 2) <> ";;" Then ' controla fin de linea
                                If linea >= 8 Then
                                    arraytext = Split(sLine, ";")
                                    If arraytext.Length < 2 Then
                                        arraytext = Split(sLine, ",")
                                    End If
                                    If Trim(arraytext(1)) <> "" Then
                                        matricula = Trim(arraytext(1))
                                        If arraytext.Length <= 5 Then
                                            grasa = -1
                                            proteina = -1
                                            lactosa = -1
                                            st = -1
                                            If Trim(arraytext(3)) <> "" And Trim(arraytext(3)) <> "-" Then
                                                Try
                                                    rc = arraytext(3)
                                                Catch ex As Exception
                                                    MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: RC")
                                                    Exit Sub
                                                End Try
                                            Else
                                                rc = -1
                                            End If
                                            crioscopia = -1
                                            urea = -1
                                            proteinav = -1
                                            caseina = -1
                                            densidad = -1
                                            ph = -1
                                        Else
                                            Dim prueba As String
                                            prueba = Trim(arraytext(5))
                                            If prueba <> "-" Then
                                                If Trim(arraytext(5)) <> "" Then
                                                    Try
                                                        grasa = arraytext(5)
                                                    Catch ex As Exception
                                                        MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: Grasa")
                                                        Exit Sub
                                                    End Try
                                                Else
                                                    grasa = -1
                                                End If
                                                If Trim(arraytext(6)) <> "" Then
                                                    Try
                                                        proteina = arraytext(6)
                                                    Catch ex As Exception
                                                        MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: Proteína")
                                                        Exit Sub
                                                    End Try
                                                Else
                                                    proteina = -1
                                                End If
                                                If Trim(arraytext(7)) <> "" Then
                                                    Try
                                                        lactosa = arraytext(7)
                                                    Catch ex As Exception
                                                        MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: Lactosa")
                                                        Exit Sub
                                                    End Try
                                                Else
                                                    lactosa = -1
                                                End If
                                                If Trim(arraytext(8)) <> "" Then
                                                    Try
                                                        st = arraytext(8)
                                                    Catch ex As Exception
                                                        MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: Sólidos totales")
                                                        Exit Sub
                                                    End Try
                                                Else
                                                    st = -1
                                                End If
                                                If Trim(arraytext(3)) <> "" And Trim(arraytext(3)) <> "-" Then
                                                    Try
                                                        rc = arraytext(3)
                                                    Catch ex As Exception
                                                        MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: RC")
                                                        Exit Sub
                                                    End Try
                                                Else
                                                    rc = -1
                                                End If
                                                If Trim(arraytext(9)) <> "" Then
                                                    Try
                                                        crioscopia = arraytext(9)
                                                    Catch ex As Exception
                                                        MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: Crioscopía")
                                                        Exit Sub
                                                    End Try
                                                Else
                                                    crioscopia = -1
                                                End If
                                                If Trim(arraytext(10)) <> "" Then
                                                    Dim verifica As String = Trim(arraytext(10))
                                                    If Mid(verifica, 1, 1) = "-" Then
                                                        urea = -1
                                                    Else
                                                        Try
                                                            urea = arraytext(10)
                                                        Catch ex As Exception
                                                            MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: Urea")
                                                            Exit Sub
                                                        End Try
                                                    End If
                                                Else
                                                    urea = -1
                                                End If
                                                'If Trim(arraytext(28)) <> "" Then
                                                '    Try
                                                '        proteinav = arraytext(28)
                                                '    Catch ex As Exception
                                                '        MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: Proteína verdadera")
                                                '        Exit Sub
                                                '    End Try
                                                'Else
                                                '    proteinav = -1
                                                'End If
                                                If Trim(arraytext(11)) <> "" Then
                                                    Try
                                                        caseina = arraytext(11)
                                                    Catch ex As Exception
                                                        MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: Caseína")
                                                        Exit Sub
                                                    End Try
                                                Else
                                                    caseina = -1
                                                End If
                                                'If Trim(arraytext(30)) <> "" Then
                                                '    Try
                                                '        densidad = arraytext(30)
                                                '    Catch ex As Exception
                                                '        MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: Densidad")
                                                '        Exit Sub
                                                '    End Try
                                                'Else
                                                '    densidad = -1
                                                'End If
                                                If Trim(arraytext(12)) <> "" Then
                                                    Try
                                                        ph = arraytext(12)
                                                    Catch ex As Exception
                                                        MsgBox("Error en archivo: " & file.Name & ", línea: " & linea & ", valor: pH")
                                                        Exit Sub
                                                    End Try
                                                Else
                                                    ph = -1
                                                End If
                                            Else
                                                grasa = -1
                                                proteina = -1
                                                lactosa = -1
                                                st = -1
                                                rc = -1
                                                crioscopia = -1
                                                urea = -1
                                                proteinav = -1
                                                caseina = -1
                                                densidad = -1
                                                ph = -1
                                            End If
                                        End If
                                        ficha2 = Mid(file.Name, Len(file.Name) - 4, 1)
                                        ficha3 = Mid(file.Name, 1, 1)
                                        If ficha2 = "a" Or ficha2 = "A" Or ficha2 = "b" Or ficha2 = "B" Or ficha2 = "c" Or ficha2 = "C" Or ficha2 = "d" Or ficha2 = "D" Or ficha2 = "e" Or ficha2 = "E" Or ficha2 = "f" Or ficha2 = "F" Or ficha2 = "g" Or ficha2 = "G" Or ficha2 = "h" Or ficha2 = "H" Or ficha2 = "i" Or ficha2 = "I" Or ficha2 = "j" Or ficha2 = "J" Then
                                            ficha = Mid(file.Name, 1, Len(file.Name) - 5)
                                        Else
                                            ficha = Mid(file.Name, 1, Len(file.Name) - 4)
                                            'ficha = Mid(file.Name, 1, Len(file.Name) - 21)
                                        End If
                                        If Mid(ficha, 1, 1) = "l" Or Mid(ficha, 1, 1) = "L" Then
                                            Dim MyString As String = ficha
                                            Dim MyChar As Char() = {"l"c, "L"c}
                                            Dim NewString As String = MyString.TrimStart(MyChar)
                                            ficha3 = NewString
                                        Else
                                            ficha3 = ficha
                                        End If
                                        Dim fechaoriginal As Date = Now()
                                        Dim fecha As String
                                        fecha = Format(fechaoriginal, "yyyy-MM-dd")
                                        c.FICHA = ficha3
                                        c.FECHA = fecha
                                        c.EQUIPO = "delta"
                                        c.PRODUCTO = producto
                                        c.MUESTRA = matricula
                                        c.RC = rc
                                        c.GRASA = grasa
                                        c.PROTEINA = proteina
                                        c.LACTOSA = lactosa
                                        c.ST = st
                                        c.CRIOSCOPIA = crioscopia
                                        c.UREA = urea
                                        c.PROTEINAV = proteinav
                                        c.CASEINA = caseina
                                        c.DENSIDAD = densidad
                                        c.PH = ph
                                        c.guardar()
                                        'ca.FICHA = ficha3
                                        'ca.FECHA = fecha
                                        'sa.ID = ficha3
                                        'sa = sa.buscar
                                        'ca.PRODUCTOR = sa.IDPRODUCTOR
                                        'ca.MUESTRA = matricula
                                        'ca.RC = rc
                                        'ca.TAMBO = sa.TAMBO
                                        'ca.guardar()
                                        Dim sa2 As New dSolicitudAnalisis
                                        sa2.ID = ficha3
                                        sa2.actualizarfechaproceso(fecha)
                                        sa2 = Nothing
                                    End If
                                End If
                            End If 'controla fin de linea
                        End If
                        linea = linea + 1
                    Loop Until sLine Is Nothing
                    objReader.Close()
                End If
                '*** MOVER ARCHIVO ***********************************************************************
                Dim sArchivoOrigen As String = "\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Control lechero\" & nombrearchivo
                Dim sRutaDestino As String = "Y:\documentos\secretaria\analisis\leche\bentley-delta\pasados\control lechero\pasados NET\" & nombrearchivo
                Try
                    ' Mover el fichero.si existe lo sobreescribe  
                    My.Computer.FileSystem.MoveFile(sArchivoOrigen, _
                                                    sRutaDestino, _
                                                    True)
                    'MsgBox("Ok.", MsgBoxStyle.Information, "Mover archivo")
                    ' errores  
                Catch ex As Exception
                    MsgBox(ex.Message.ToString, MsgBoxStyle.Critical)
                End Try
                '***********************************
                'Insert tabla preinforme_calidad
                Dim pi As New dPreinformes
                Dim fechaactual As Date = Now()
                Dim _fecha As String
                _fecha = Format(fechaactual, "yyyy-MM-dd")
                pi.FICHA = ficha3
                pi = pi.buscar
                If Not pi Is Nothing Then
                Else
                    Dim pi2 As New dPreinformes
                    Dim sa As New dSolicitudAnalisis
                    Dim tipof As Integer = 0
                    sa.ID = ficha3
                    sa = sa.buscar
                    If Not sa Is Nothing Then
                        tipof = sa.IDTIPOINFORME
                    End If
                    pi2.FICHA = ficha3
                    pi2.TIPO = tipof
                    pi2.CREADO = 0
                    pi2.FECHA = _fecha
                    pi2.guardar()
                    pi2 = Nothing
                    sa = Nothing
                End If
                pi = Nothing
                ' Grabar estado de la ficha
                Dim est As New dEstados
                est.FICHA = ficha3
                est.ESTADO = 4
                est.FECHA = _fecha
                est.guardar2()
                est = Nothing
            End If 'fin de control si archivoes de delta 400
        Next
    End Sub
    Private Sub controllecheroxls()
        Dim extension As String
        Dim nombrearchivo As String = ""
        Dim linea As Integer
        'Dim folder As New DirectoryInfo("d:\documentos\secretaria\analisis\leche\bentley-delta\pasados\control lechero")
        Dim folder As New DirectoryInfo("\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Control lechero")
        For Each file As FileInfo In folder.GetFiles("*.xls")
            'ListBox1.Items.Add(file.Name)
            nombrearchivo = file.Name
            linea = 1
            extension = Microsoft.VisualBasic.Right(file.Name, 3)
            'Dim objReader As New StreamReader("d:\documentos\secretaria\analisis\leche\bentley-delta\pasados\control lechero\" & file.Name)
            Dim objReader As New StreamReader("\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Control lechero\" & file.Name)
            Dim sLine As String = ""
            Dim matricula As String = ""
            Dim grasa As Double = 0
            Dim proteina As Double = 0
            Dim lactosa As Double = 0
            Dim st As Double = 0
            Dim rc As Integer = 0
            Dim ficha As String = ""
            Dim ficha2 As String = ""
            Dim ficha3 As String = ""
            Dim equipo As String = ""
            Dim producto As String = ""
            Dim crioscopia As Integer = 0
            Dim urea As Integer = 0
            Dim proteinav As Double = 0
            Dim caseina As Double = 0
            Dim densidad As Double = 0
            Dim ph As Double = 0
            ' *** SI EL ARCHIVO ES XLS **************************************************************************************
            If extension = "xls" Or extension = "XLS" Then
                Dim c As New dImpControl()
                'Dim ca As New dControlAux
                Dim sa As New dSolicitudAnalisis
                Dim p As New dCliente
                Dim Arch As String, CantFilas As Integer
                'Arch = "d:\documentos\secretaria\analisis\leche\bentley-delta\pasados\control lechero\" & file.Name
                Arch = "\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Control lechero\" & file.Name
                Dim x1app As Microsoft.Office.Interop.Excel.Application
                Dim x1libro As Microsoft.Office.Interop.Excel.Workbook
                Dim x1hoja As Microsoft.Office.Interop.Excel.Worksheet
                x1app = CType(CreateObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)
                x1libro = CType(x1app.Workbooks.Open(Arch), Microsoft.Office.Interop.Excel.Workbook)
                x1hoja = CType(x1libro.Worksheets(1), Microsoft.Office.Interop.Excel.Worksheet)
                Dim bandera As Integer = 0
                CantFilas = x1app.Range("a1").CurrentRegion.Rows.Count
                ficha2 = Mid(file.Name, Len(file.Name) - 4, 1)
                ficha3 = Mid(file.Name, 1, 1)
                If ficha2 = "a" Or ficha2 = "A" Or ficha2 = "b" Or ficha2 = "B" Or ficha2 = "c" Or ficha2 = "C" Or ficha2 = "d" Or ficha2 = "D" Or ficha2 = "e" Or ficha2 = "E" Or ficha2 = "f" Or ficha2 = "F" Or ficha2 = "g" Or ficha2 = "G" Or ficha2 = "h" Or ficha2 = "H" Or ficha2 = "i" Or ficha2 = "I" Or ficha2 = "j" Or ficha2 = "J" Then
                    ficha = Mid(file.Name, 1, Len(file.Name) - 5)
                Else
                    ficha = Mid(file.Name, 1, Len(file.Name) - 4)
                End If
                If Mid(ficha, 1, 1) = "l" Or Mid(ficha, 1, 1) = "L" Then
                    'Dim MyString As String = ficha
                    'ficha3 = MyString.Remove(1, 1)
                    Dim MyString As String = ficha
                    Dim MyChar As Char() = {"l"c, "L"c}
                    Dim NewString As String = MyString.TrimStart(MyChar)
                    ficha3 = NewString
                Else
                    ficha3 = ficha
                End If
                Dim fechaoriginal As Date = Now()
                Dim fecha As String
                fecha = Format(fechaoriginal, "yyyy-MM-dd")
                Dim id As String = ""
                For i = 1 To CantFilas
                    'If linea > 7 Then
                    If Trim(x1hoja.Cells(i, 1).formula) <> "" Then
                        id = Trim(x1hoja.Cells(i, 1).value)
                    Else
                        id = -1
                    End If
                    If Trim(x1hoja.Cells(i, 2).formula) <> "" Then
                        matricula = Trim(x1hoja.Cells(i, 2).value)
                    Else
                        matricula = -1
                    End If
                    If Trim(x1hoja.Cells(i, 3).formula) <> "" Then
                        Try
                            grasa = x1hoja.Cells(i, 3).value
                        Catch ex As Exception
                            MsgBox("Error en archivo: " & file.Name & ", línea: " & i & ", valor: Grasa")
                            Exit Sub
                        End Try
                    Else
                        grasa = -1
                    End If
                    If Trim(x1hoja.Cells(i, 4).formula) <> "" Then
                        Try
                            proteina = x1hoja.Cells(i, 4).value
                            bandera = 1
                        Catch ex As Exception
                            MsgBox("Error en archivo: " & file.Name & ", línea: " & i & ", valor: Proteína")
                            Exit Sub
                        End Try
                    Else
                        proteina = -1
                    End If
                    If Trim(x1hoja.Cells(i, 5).formula) <> "" Then
                        Try
                            lactosa = x1hoja.Cells(i, 5).value
                        Catch ex As Exception
                            MsgBox("Error en archivo: " & file.Name & ", línea: " & i & ", valor: Lactosa")
                            Exit Sub
                        End Try
                    Else
                        lactosa = -1
                    End If
                    If Trim(x1hoja.Cells(i, 6).formula) <> "" Then
                        Try
                            st = x1hoja.Cells(i, 6).value
                        Catch ex As Exception
                            MsgBox("Error en archivo: " & file.Name & ", línea: " & i & ", valor: Sólidos totales")
                            Exit Sub
                        End Try
                    Else
                        st = -1
                    End If
                    If Trim(x1hoja.Cells(i, 7).formula) <> "" Then
                        Try
                            rc = x1hoja.Cells(i, 7).value
                        Catch ex As Exception
                            MsgBox("Error en archivo: " & file.Name & ", línea: " & i & ", valor: RC")
                            Exit Sub
                        End Try
                    Else
                        rc = -1
                        bandera = 2
                    End If
                    If bandera = 0 Then
                        c.FICHA = ficha3
                        c.FECHA = fecha
                        c.EQUIPO = "bentley"
                        c.PRODUCTO = "leche"
                        c.MUESTRA = matricula
                        c.RC = grasa
                        c.GRASA = -1
                        c.PROTEINA = proteina
                        c.LACTOSA = lactosa
                        c.ST = st
                        c.CRIOSCOPIA = -1
                        c.UREA = -1
                        c.PROTEINAV = -1
                        c.CASEINA = -1
                        c.DENSIDAD = -1
                        c.PH = -1
                        c.guardar()
                        'ca.FICHA = ficha3
                        'ca.FECHA = fecha
                        'sa.ID = ficha3
                        'sa = sa.buscar
                        'ca.PRODUCTOR = sa.IDPRODUCTOR
                        'ca.MUESTRA = matricula
                        'ca.RC = grasa
                        'ca.TAMBO = sa.TAMBO
                        'ca.guardar()
                        Dim sa2 As New dSolicitudAnalisis
                        sa2.ID = ficha3
                        sa2.actualizarfechaproceso(fecha)
                        sa2 = Nothing
                    ElseIf bandera = 1 Then
                        c.FICHA = ficha3
                        c.FECHA = fecha
                        c.EQUIPO = "bentley"
                        c.PRODUCTO = "leche"
                        c.MUESTRA = matricula
                        c.RC = rc
                        c.GRASA = grasa
                        c.PROTEINA = proteina
                        c.LACTOSA = lactosa
                        c.ST = st
                        c.CRIOSCOPIA = -1
                        c.UREA = -1
                        c.PROTEINAV = -1
                        c.CASEINA = -1
                        c.DENSIDAD = -1
                        c.PH = -1
                        c.guardar()
                        'ca.FICHA = ficha3
                        'ca.FECHA = fecha
                        'sa.ID = ficha3
                        'sa = sa.buscar
                        'ca.PRODUCTOR = sa.IDPRODUCTOR
                        'ca.MUESTRA = matricula
                        'ca.RC = rc
                        'ca.TAMBO = sa.TAMBO
                        'ca.guardar()
                        Dim sa2 As New dSolicitudAnalisis
                        sa2.ID = ficha3
                        sa2.actualizarfechaproceso(fecha)
                        sa2 = Nothing
                    ElseIf bandera = 2 Then
                        c.FICHA = ficha3
                        c.FECHA = fecha
                        c.EQUIPO = "bentley"
                        c.PRODUCTO = "leche"
                        c.MUESTRA = id
                        c.RC = st
                        c.GRASA = matricula
                        c.PROTEINA = grasa
                        c.LACTOSA = proteina
                        c.ST = lactosa
                        c.CRIOSCOPIA = -1
                        c.UREA = -1
                        c.PROTEINAV = -1
                        c.CASEINA = -1
                        c.DENSIDAD = -1
                        c.PH = -1
                        c.guardar()
                        'ca.FICHA = ficha3
                        'ca.FECHA = fecha
                        'sa.ID = ficha3
                        'sa = sa.buscar
                        'ca.PRODUCTOR = sa.IDPRODUCTOR
                        'ca.MUESTRA = id
                        'ca.RC = st
                        'ca.TAMBO = sa.TAMBO
                        'ca.guardar()
                        Dim sa2 As New dSolicitudAnalisis
                        sa2.ID = ficha3
                        sa2.actualizarfechaproceso(fecha)
                        sa2 = Nothing
                    End If
                Next
                ' Cierro Excel
                x1libro.Close()
                x1app = Nothing
                x1libro = Nothing
                x1hoja = Nothing
                objReader.Close()
                Dim proceso As System.Diagnostics.Process()
                proceso = System.Diagnostics.Process.GetProcessesByName("EXCEL")
                For Each opro As System.Diagnostics.Process In proceso
                    'antes de iniciar el proceso obtengo la fecha en que inicie el 
                    'proceso para detener todos los procesos que excel que inicio
                    'mi código durante el proceso
                    opro.Kill()
                Next
            End If
            '*** MOVER ARCHIVO ***********************************************************************
            'Dim sArchivoOrigen As String = "d:\documentos\secretaria\analisis\leche\bentley-delta\pasados\control lechero\" & nombrearchivo
            Dim sArchivoOrigen As String = "\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Control lechero\" & nombrearchivo
            'Dim sRutaDestino As String = "d:\documentos\secretaria\analisis\leche\bentley-delta\pasados\control lechero\pasados NET\" & nombrearchivo
            Dim sRutaDestino As String = "Y:\documentos\secretaria\analisis\leche\bentley-delta\pasados\control lechero\pasados NET\" & nombrearchivo
            Try
                ' Mover el fichero.si existe lo sobreescribe  
                My.Computer.FileSystem.MoveFile(sArchivoOrigen, _
                                                sRutaDestino, _
                                                True)
                'MsgBox("Ok.", MsgBoxStyle.Information, "Mover archivo")
                ' errores  
            Catch ex As Exception
                MsgBox(ex.Message.ToString, MsgBoxStyle.Critical)
            End Try
            '***********************************
            'Insert tabla preinforme_calidad
            Dim pi As New dPreinformes
            Dim fechaactual As Date = Now()
            Dim _fecha As String
            _fecha = Format(fechaactual, "yyyy-MM-dd")
            pi.FICHA = ficha3
            pi = pi.buscar
            If Not pi Is Nothing Then
            Else
                Dim pi2 As New dPreinformes
                Dim sa As New dSolicitudAnalisis
                Dim tipof As Integer = 0
                sa.ID = ficha3
                sa = sa.buscar
                If Not sa Is Nothing Then
                    tipof = sa.IDTIPOINFORME
                End If
                pi2.FICHA = ficha3
                pi2.TIPO = tipof
                pi2.CREADO = 0
                pi2.FECHA = _fecha
                pi2.guardar()
                pi2 = Nothing
                sa = Nothing
            End If
            pi = Nothing
            ' Grabar estado de la ficha
            Dim est As New dEstados
            est.FICHA = ficha3
            est.ESTADO = 4
            est.FECHA = _fecha
            est.guardar2()
            est = Nothing
            '****************************
        Next
    End Sub
    Private Sub controllecherofat()
        Dim extension As String
        Dim nombrearchivo As String = ""
        Dim linea As Integer
        'Dim folder As New DirectoryInfo("d:\documentos\secretaria\analisis\leche\bentley-delta\pasados\control lechero")
        Dim folder As New DirectoryInfo("\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Control lechero")
        For Each file As FileInfo In folder.GetFiles("*.fat")
            nombrearchivo = file.Name
            linea = 1
            extension = Microsoft.VisualBasic.Right(file.Name, 3)
            'Dim objReader As New StreamReader("d:\documentos\secretaria\analisis\leche\bentley-delta\pasados\control lechero\" & file.Name)
            Dim objReader As New StreamReader("\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Control lechero\" & file.Name)
            Dim sLine As String = ""
            Dim matricula As String = ""
            Dim grasa As Double = 0
            Dim proteina As Double = 0
            Dim lactosa As Double = 0
            Dim st As Double = 0
            Dim rc As Integer = 0
            Dim ficha As String = ""
            Dim ficha2 As String = ""
            Dim ficha3 As String = ""
            Dim equipo As String = ""
            Dim producto As String = ""
            Dim crioscopia As Integer = 0
            Dim urea As Integer = 0
            Dim proteinav As Double = 0
            Dim caseina As Double = 0
            Dim densidad As Double = 0
            Dim ph As Double = 0
            ' *** SI EL ARCHIVO ES FAT **************************************************************************************
            If extension = "fat" Or extension = "FAT" Then
                Dim c As New dImpControl()
                'Dim ca As New dControlAux
                Dim sa As New dSolicitudAnalisis
                Dim p As New dCliente
                Dim cuentalinea As Long = 1
                Do
                    sLine = objReader.ReadLine()
                    Dim fechaoriginal As Date = Now()
                    Dim fecha As String = ""
                    If Not sLine Is Nothing Then
                        Dim Texto As String
                        Dim id As String = 0
                        ficha2 = Mid(file.Name, Len(file.Name) - 4, 1)
                        ficha3 = Mid(file.Name, 1, 1)
                        If ficha2 = "a" Or ficha2 = "A" Or ficha2 = "b" Or ficha2 = "B" Or ficha2 = "c" Or ficha2 = "C" Or ficha2 = "d" Or ficha2 = "D" Or ficha2 = "e" Or ficha2 = "E" Or ficha2 = "f" Or ficha2 = "F" Or ficha2 = "g" Or ficha2 = "G" Or ficha2 = "h" Or ficha2 = "H" Or ficha2 = "i" Or ficha2 = "I" Or ficha2 = "j" Or ficha2 = "J" Then
                            ficha = Mid(file.Name, 1, Len(file.Name) - 5)
                        Else
                            ficha = Mid(file.Name, 1, Len(file.Name) - 4)
                        End If
                        If Mid(ficha, 1, 1) = "l" Or Mid(ficha, 1, 1) = "L" Then
                            Dim MyString As String = ficha
                            Dim MyChar As Char() = {"l"c, "L"c}
                            Dim NewString As String = MyString.TrimStart(MyChar)
                            ficha3 = NewString
                        Else
                            ficha3 = ficha
                        End If
                        fecha = Format(fechaoriginal, "yyyy-MM-dd")
                        Texto = sLine
                        id = Trim(Mid(Texto, 1, 8))
                        If Trim(Mid(Texto, 9, 9)) <> "" Then
                            matricula = Trim(Mid(Texto, 9, 9))
                        Else
                            matricula = id
                        End If
                        If Trim(Mid(Texto, 18, 9)) <> "" Then
                            Try
                                grasa = Trim(Mid(Texto, 18, 9))
                            Catch ex As Exception
                                MsgBox("Error en archivo: " & file.Name & ", línea: " & cuentalinea & ", valor: Grasa", MsgBoxStyle.Critical)
                                Exit Sub
                            End Try
                        Else
                            grasa = -1
                        End If
                        If Trim(Mid(Texto, 27, 9)) <> "" Then
                            Try
                                proteina = Trim(Mid(Texto, 27, 9))
                            Catch ex As Exception
                                MsgBox("Error en archivo: " & file.Name & ", línea: " & cuentalinea & ", valor: Proteína")
                                Exit Sub
                            End Try
                        Else
                            proteina = -1
                        End If
                        If Trim(Mid(Texto, 36, 9)) <> "" Then
                            Try
                                lactosa = Trim(Mid(Texto, 36, 9))
                            Catch ex As Exception
                                MsgBox("Error en archivo: " & file.Name & ", línea: " & cuentalinea & ", valor: Lactosa")
                                Exit Sub
                            End Try
                        Else
                            lactosa = -1
                        End If
                        If Trim(Mid(Texto, 45, 9)) <> "" Then
                            Try
                                st = Trim(Mid(Texto, 45, 9))
                            Catch ex As Exception
                                MsgBox("Error en archivo: " & file.Name & ", línea: " & cuentalinea & ", valor: Sólidos totales")
                                Exit Sub
                            End Try
                        Else
                            st = -1
                        End If
                        If Trim(Mid(Texto, 54, 10)) <> "" Then
                            Try
                                rc = Trim(Mid(Texto, 54, 10))
                            Catch ex As Exception
                                MsgBox("Error en archivo: " & file.Name & ", línea: " & cuentalinea & ", valor: RC")
                                Exit Sub
                            End Try
                        Else
                            rc = -1
                        End If
                    End If
                    If fecha <> "" Then
                        c.FICHA = ficha3
                        c.FECHA = fecha
                        c.EQUIPO = "bentley"
                        c.PRODUCTO = "leche"
                        c.MUESTRA = matricula
                        c.RC = rc
                        c.GRASA = grasa
                        c.PROTEINA = proteina
                        c.LACTOSA = lactosa
                        c.ST = st
                        c.CRIOSCOPIA = -1
                        c.UREA = -1
                        c.PROTEINAV = -1
                        c.CASEINA = -1
                        c.DENSIDAD = -1
                        c.PH = -1
                        c.guardar()
                        'ca.FICHA = ficha3
                        'ca.FECHA = fecha
                        'sa.ID = ficha3
                        'sa = sa.buscar
                        'ca.PRODUCTOR = sa.IDPRODUCTOR
                        'ca.MUESTRA = matricula
                        'ca.RC = rc
                        'ca.TAMBO = sa.TAMBO
                        'ca.guardar()
                        Dim sa2 As New dSolicitudAnalisis
                        sa2.ID = ficha3
                        sa2.actualizarfechaproceso(fecha)
                        sa2 = Nothing
                    End If
                    cuentalinea = cuentalinea + 1
                Loop Until sLine Is Nothing
                objReader.Close()
            End If
            '*** MOVER ARCHIVO ***********************************************************************
            'Dim sArchivoOrigen As String = "d:\documentos\secretaria\analisis\leche\bentley-delta\pasados\control lechero\" & nombrearchivo
            Dim sArchivoOrigen As String = "\\192.168.1.20\d\Documentos\SECRETARIA\Analisis\Leche\Bentley-delta\Control lechero\" & nombrearchivo
            'Dim sRutaDestino As String = "d:\documentos\secretaria\analisis\leche\bentley-delta\pasados\control lechero\pasados NET\" & nombrearchivo
            Dim sRutaDestino As String = "Y:\documentos\secretaria\analisis\leche\bentley-delta\pasados\control lechero\pasados NET\" & nombrearchivo
            Try
                ' Mover el fichero.si existe lo sobreescribe  
                My.Computer.FileSystem.MoveFile(sArchivoOrigen, _
                                                sRutaDestino, _
                                                True)
                'MsgBox("Ok.", MsgBoxStyle.Information, "Mover archivo")
                ' errores  
            Catch ex As Exception
                MsgBox(ex.Message.ToString, MsgBoxStyle.Critical)
            End Try
            '***********************************
            'Insert tabla preinforme_calidad
            Dim pi As New dPreinformes
            Dim fechaactual As Date = Now()
            Dim _fecha As String
            _fecha = Format(fechaactual, "yyyy-MM-dd")
            pi.FICHA = ficha3
            pi = pi.buscar
            If Not pi Is Nothing Then
            Else
                Dim pi2 As New dPreinformes
                Dim sa As New dSolicitudAnalisis
                Dim tipof As Integer = 0
                sa.ID = ficha3
                sa = sa.buscar
                If Not sa Is Nothing Then
                    tipof = sa.IDTIPOINFORME
                End If
                pi2.FICHA = ficha3
                pi2.TIPO = tipof
                pi2.CREADO = 0
                pi2.FECHA = _fecha
                pi2.guardar()
                pi2 = Nothing
                sa = Nothing
            End If
            pi = Nothing
            ' Grabar estado de la ficha
            Dim est As New dEstados
            est.FICHA = ficha3
            est.ESTADO = 4
            est.FECHA = _fecha
            est.guardar2()
            est = Nothing
            '**********************************
        Next
    End Sub

#End Region

#Region "PRE-INFORMES"

    Private Sub preinforme_control(ByVal id_sol As Long)
        Dim x1app As Microsoft.Office.Interop.Excel.Application
        Dim x1libro As Microsoft.Office.Interop.Excel.Workbook
        Dim x1hoja As Microsoft.Office.Interop.Excel.Worksheet
        x1app = CType(CreateObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)
        x1libro = CType(x1app.Workbooks.Add, Microsoft.Office.Interop.Excel.Workbook)
        x1hoja = CType(x1libro.Worksheets(1), Microsoft.Office.Interop.Excel.Worksheet)
        Dim c As New dControl
        Dim i As New dIbc
        Dim sa As New dSolicitudAnalisis
        Dim pro As New dCliente
        Dim tec As New dTecnicos
        Dim lista As New ArrayList
        Dim lista2 As New ArrayList
        Dim lista3 As New ArrayList
        '*****************************
        Dim idsol As Long = id_sol 'ficha
        sa.ID = idsol
        sa = sa.buscar
        '*****************************
        Dim cc As New dClienteConvenio
        Dim listacc As New ArrayList
        Dim bhb As Integer = 0
        listacc = cc.listarporcliente(sa.IDPRODUCTOR)
        If Not listacc Is Nothing Then
            For Each cc In listacc
                If cc.CONVENIO = 2 Then
                    bhb = 1
                End If
            Next
        End If
        '************************************************
        Dim fila As Integer
        Dim columna As Integer
        fila = 1
        columna = 2
        '*** ENCABEZADO ********************************************************************************
        '***********************************************************************************************
        x1hoja.Cells(1, 1).columnwidth = 5
        x1hoja.Cells(1, 2).columnwidth = 5
        x1hoja.Cells(1, 3).columnwidth = 5
        x1hoja.Cells(1, 4).columnwidth = 5
        x1hoja.Cells(1, 5).columnwidth = 5
        x1hoja.Cells(1, 6).columnwidth = 5
        x1hoja.Cells(1, 7).columnwidth = 5
        x1hoja.Cells(1, 8).columnwidth = 3
        x1hoja.Cells(1, 9).columnwidth = 5
        x1hoja.Cells(1, 10).columnwidth = 5
        x1hoja.Cells(1, 11).columnwidth = 5
        x1hoja.Cells(1, 12).columnwidth = 5
        x1hoja.Cells(1, 13).columnwidth = 5
        x1hoja.Cells(1, 14).columnwidth = 5
        x1hoja.Cells(1, 15).columnwidth = 5
        x1hoja.Range("A1", "D1").Merge()
        fila = 15
        columna = 1
        '*** FIN DEL ENCABEZADO ***********************************************************************************
        '**********************************************************************************************************
        lista = c.listarporsolicitud(idsol)
        lista2 = c.listarporrc(idsol)
        x1hoja.Cells(fila, columna).Formula = "Listado ordenado por identificación"
        x1hoja.Cells(fila, columna).Font.Bold = True
        x1hoja.Cells(fila, columna).Font.Size = 8
        columna = columna + 8
        x1hoja.Cells(fila, columna).Formula = "Listado ordenado decreciente por Recuento celular"
        x1hoja.Cells(fila, columna).Font.Bold = True
        x1hoja.Cells(fila, columna).Font.Size = 8
        fila = fila + 1
        columna = 1
        Dim filaguia As Integer = fila
        x1hoja.Cells(fila, columna).Formula = "Ident."
        x1hoja.Cells(fila, columna).Font.Bold = True
        x1hoja.Cells(fila, columna).Font.Size = 8
        x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
        x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        columna = columna + 1
        x1hoja.Cells(fila, columna).Formula = "RCS"
        x1hoja.Cells(fila, columna).Font.Bold = True
        x1hoja.Cells(fila, columna).Font.Size = 8
        x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
        x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        columna = columna + 1
        x1hoja.Cells(fila, columna).Formula = "Gr"
        x1hoja.Cells(fila, columna).Font.Bold = True
        x1hoja.Cells(fila, columna).Font.Size = 8
        x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
        x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        columna = columna + 1
        x1hoja.Cells(fila, columna).Formula = "Pr"
        x1hoja.Cells(fila, columna).Font.Bold = True
        x1hoja.Cells(fila, columna).Font.Size = 8
        x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
        x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        columna = columna + 1
        x1hoja.Cells(fila, columna).Formula = "Lac*"
        x1hoja.Cells(fila, columna).Font.Bold = True
        x1hoja.Cells(fila, columna).Font.Size = 8
        x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
        x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        columna = columna + 1
        x1hoja.Cells(fila, columna).Formula = "MUN*"
        x1hoja.Cells(fila, columna).Font.Bold = True
        x1hoja.Cells(fila, columna).Font.Size = 8
        x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
        x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        columna = columna + 1
        If bhb = 0 Then
            x1hoja.Cells(fila, columna).Formula = "Cas*"
            x1hoja.Cells(fila, columna).Font.Bold = True
            x1hoja.Cells(fila, columna).Font.Size = 8
            x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
            x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
            columna = 1
        Else
            x1hoja.Cells(fila, columna).Formula = "BHB"
            x1hoja.Cells(fila, columna).Font.Bold = True
            x1hoja.Cells(fila, columna).Font.Size = 8
            x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
            x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
            columna = 1
        End If

        fila = fila + 1
        x1hoja.Cells(1, 1).RowHeight = 18
        x1hoja.Range("A17", "A18").Merge()
        x1hoja.Range("A17", "A18").Borders.Color = RGB(0, 0, 0)
        x1hoja.Range("A17", "A18").Interior.Color = RGB(192, 192, 192)
        x1hoja.Range("A17", "A18").WrapText = True
        x1hoja.Cells(fila, columna).formula = ""
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
        x1hoja.Cells(fila, columna).Font.Size = 6
        columna = columna + 1
        x1hoja.Cells(1, 1).RowHeight = 18
        x1hoja.Range("B17", "B18").Merge()
        x1hoja.Range("B17", "B18").Borders.Color = RGB(0, 0, 0)
        x1hoja.Range("B17", "B18").Interior.Color = RGB(192, 192, 192)
        x1hoja.Range("B17", "B18").WrapText = True
        x1hoja.Cells(fila, columna).formula = "x 1.000 cel/mL"
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
        x1hoja.Cells(fila, columna).Font.Size = 6
        columna = columna + 1
        x1hoja.Cells(1, 1).RowHeight = 18
        x1hoja.Range("C17", "C18").Merge()
        x1hoja.Range("C17", "C18").Borders.Color = RGB(0, 0, 0)
        x1hoja.Range("C17", "C18").Interior.Color = RGB(192, 192, 192)
        x1hoja.Range("C17", "C18").WrapText = True
        x1hoja.Cells(fila, columna).formula = "g/100mL"
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
        x1hoja.Cells(fila, columna).Font.Size = 6
        columna = columna + 1
        x1hoja.Cells(1, 1).RowHeight = 18
        x1hoja.Range("D17", "D18").Merge()
        x1hoja.Range("D17", "D18").Borders.Color = RGB(0, 0, 0)
        x1hoja.Range("D17", "D18").Interior.Color = RGB(192, 192, 192)
        x1hoja.Range("D17", "D18").WrapText = True
        x1hoja.Cells(fila, columna).formula = "g/100mL"
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
        x1hoja.Cells(fila, columna).Font.Size = 6
        columna = columna + 1
        x1hoja.Cells(1, 1).RowHeight = 18
        x1hoja.Range("E17", "E18").Merge()
        x1hoja.Range("E17", "E18").Borders.Color = RGB(0, 0, 0)
        x1hoja.Range("E17", "E18").Interior.Color = RGB(192, 192, 192)
        x1hoja.Range("E17", "E18").WrapText = True
        x1hoja.Cells(fila, columna).formula = "g/100mL"
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
        x1hoja.Cells(fila, columna).Font.Size = 6
        columna = columna + 1
        x1hoja.Cells(1, 1).RowHeight = 18
        x1hoja.Range("F17", "F18").Merge()
        x1hoja.Range("F17", "F18").Borders.Color = RGB(0, 0, 0)
        x1hoja.Range("F17", "F18").Interior.Color = RGB(192, 192, 192)
        x1hoja.Range("F17", "F18").WrapText = True
        x1hoja.Cells(fila, columna).formula = "mg/dL"
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
        x1hoja.Cells(fila, columna).Font.Size = 6
        columna = columna + 1
        If bhb = 0 Then
            x1hoja.Cells(1, 1).RowHeight = 18
            x1hoja.Range("G17", "G18").Merge()
            x1hoja.Range("G17", "G18").Borders.Color = RGB(0, 0, 0)
            x1hoja.Range("G17", "G18").Interior.Color = RGB(192, 192, 192)
            x1hoja.Range("G17", "G18").WrapText = True
            x1hoja.Cells(fila, columna).formula = "g/100mL"
            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
            x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
            x1hoja.Cells(fila, columna).Font.Size = 6
            columna = 1
        Else
            x1hoja.Cells(1, 1).RowHeight = 18
            x1hoja.Range("G17", "G18").Merge()
            x1hoja.Range("G17", "G18").Borders.Color = RGB(0, 0, 0)
            x1hoja.Range("G17", "G18").Interior.Color = RGB(192, 192, 192)
            x1hoja.Range("G17", "G18").WrapText = True
            x1hoja.Cells(fila, columna).formula = "mmol/L"
            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
            x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
            x1hoja.Cells(fila, columna).Font.Size = 6
            columna = 1
        End If
        fila = fila + 2
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each c In lista
                    If c.MUESTRA <> "" Then
                        x1hoja.Cells(fila, columna).formula = Trim(c.MUESTRA)
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    Else
                        x1hoja.Cells(fila, columna).formula = "-"
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    End If
                    If c.RC = -1 Then
                        x1hoja.Cells(fila, columna).formula = ""
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    Else
                        If c.RC < 30 Then
                            x1hoja.Cells(fila, columna).formula = "30"
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        Else
                            x1hoja.Cells(fila, columna).formula = c.RC
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        End If
                    End If
                    If c.GRASA = -1 Or c.GRASA = 0 Then
                        columna = columna - 1
                        x1hoja.Cells(fila, columna).formula = ""
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                        x1hoja.Cells(fila, columna).formula = "MUESTRA NO APTA **"
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    Else
                        Dim valgrasa = Val(c.GRASA)
                        If valgrasa < 2.5 Or valgrasa > 5.0 Then
                            x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
                        End If
                        x1hoja.Cells(fila, columna).formula = c.GRASA.ToString("##,##0.00")
                        x1hoja.Cells(fila, columna).NumberFormat = "#.00"
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    End If
                    If c.PROTEINA = -1 Or c.PROTEINA = 0 Then
                        x1hoja.Cells(fila, columna).formula = ""
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    Else
                        Dim valproteina = Val(c.PROTEINA)
                        If valproteina < 2.5 Or valproteina > 4.0 Then
                            x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
                        End If
                        x1hoja.Cells(fila, columna).formula = c.PROTEINA.ToString("##,##0.00")
                        x1hoja.Cells(fila, columna).NumberFormat = "#.00"
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    End If
                    If c.LACTOSA = -1 Or c.LACTOSA = 0 Then
                        x1hoja.Cells(fila, columna).formula = ""
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    Else
                        x1hoja.Cells(fila, columna).formula = c.LACTOSA.ToString("##,##0.00")
                        x1hoja.Cells(fila, columna).NumberFormat = "#.00"
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    End If
                    'Dim cs As New dControlSolicitud
                    'cs.FICHA = idsol
                    'cs = cs.buscar
                    Dim na As New dNuevoAnalisis
                    Dim listana As New ArrayList
                    Dim _urea As Integer = 0
                    Dim _caseina As Integer = 0
                    listana = na.listarporficha2(idsol)
                    If Not listana Is Nothing Then
                        For Each na In listana
                            If na.ANALISIS = 117 Then
                                _urea = 1
                            End If
                            If na.ANALISIS = 157 Then
                                _caseina = 1
                            End If
                            If na.ANALISIS = 158 Then
                                _caseina = 1
                                _urea = 1
                            End If

                        Next
                    End If
                    'If Not cs Is Nothing Then
                    If _urea = 1 Then
                        If c.UREA = -1 Or c.UREA = 0 Then
                            x1hoja.Cells(fila, columna).formula = "-"
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        Else
                            Dim valorurea As Integer
                            valorurea = c.UREA * 0.466
                            If valorurea > 20 Or valorurea < 9 Then
                                x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
                            End If
                            x1hoja.Cells(fila, columna).formula = FormatNumber(valorurea, 0)
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        End If
                    Else
                        x1hoja.Cells(fila, columna).formula = "-"
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    End If
                    If _caseina = 1 Then
                        If c.CASEINA = -1 Or c.UREA = 0 Then
                            x1hoja.Cells(fila, columna).formula = "-"
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        Else
                            x1hoja.Cells(fila, columna).formula = c.CASEINA.ToString("##,##0.00")
                            x1hoja.Cells(fila, columna).NumberFormat = "#.00"
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        End If
                    Else
                        If bhb = 1 Then
                            Dim valbhb As String = ""
                            valbhb = c.BHB.ToString("##,##0.00")
                            If c.BHB < 0.07 Then
                                valbhb = "<0.07"
                            End If
                            If c.BHB > 0.27 Then
                                valbhb = ">0.27"
                            End If
                            x1hoja.Cells(fila, columna).formula = valbhb 'c.BHB.ToString("##,##0.00")
                            x1hoja.Cells(fila, columna).NumberFormat = "#.00"
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        End If
                        x1hoja.Cells(fila, columna).formula = "-"
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    End If
                    na = Nothing
                    columna = 1
                    fila = fila + 1
                Next
                'Referencias
                fila = fila + 1
                columna = 1
            End If
            '****** ORDENADO POR RC ************************************************************************
            Dim libreinfeccion As Integer = 0
            Dim posibleinfeccion As Integer = 0
            Dim probableinfeccion As Integer = 0
            fila = filaguia
            columna = 9
            x1hoja.Cells(fila, columna).Formula = "Ident."
            x1hoja.Cells(fila, columna).Font.Bold = True
            x1hoja.Cells(fila, columna).Font.Size = 8
            x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
            x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
            columna = columna + 1
            x1hoja.Cells(fila, columna).Formula = "RCS"
            x1hoja.Cells(fila, columna).Font.Bold = True
            x1hoja.Cells(fila, columna).Font.Size = 8
            x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
            x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
            columna = columna + 1
            x1hoja.Cells(fila, columna).Formula = "Gr"
            x1hoja.Cells(fila, columna).Font.Bold = True
            x1hoja.Cells(fila, columna).Font.Size = 8
            x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
            x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
            columna = columna + 1
            x1hoja.Cells(fila, columna).Formula = "Pr"
            x1hoja.Cells(fila, columna).Font.Bold = True
            x1hoja.Cells(fila, columna).Font.Size = 8
            x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
            x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
            columna = columna + 1
            x1hoja.Cells(fila, columna).Formula = "Lac*"
            x1hoja.Cells(fila, columna).Font.Bold = True
            x1hoja.Cells(fila, columna).Font.Size = 8
            x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
            x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
            columna = columna + 1
            x1hoja.Cells(fila, columna).Formula = "MUN*"
            x1hoja.Cells(fila, columna).Font.Bold = True
            x1hoja.Cells(fila, columna).Font.Size = 8
            x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
            x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
            columna = columna + 1
            If bhb = 0 Then
                x1hoja.Cells(fila, columna).Formula = "Cas*"
                x1hoja.Cells(fila, columna).Font.Bold = True
                x1hoja.Cells(fila, columna).Font.Size = 8
                x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
                x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
                x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                columna = 1
            Else
                x1hoja.Cells(fila, columna).Formula = "BHB"
                x1hoja.Cells(fila, columna).Font.Bold = True
                x1hoja.Cells(fila, columna).Font.Size = 8
                x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
                x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
                x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                columna = 1
            End If
            columna = 9
            fila = fila + 1
            x1hoja.Cells(1, 1).RowHeight = 18
            x1hoja.Range("I17", "I18").Merge()
            x1hoja.Range("I17", "I18").Borders.Color = RGB(0, 0, 0)
            x1hoja.Range("I17", "I18").Interior.Color = RGB(192, 192, 192)
            x1hoja.Range("I17", "I18").WrapText = True
            x1hoja.Cells(fila, columna).formula = ""
            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
            x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
            x1hoja.Cells(fila, columna).Font.Size = 6
            columna = columna + 1
            x1hoja.Cells(1, 1).RowHeight = 18
            x1hoja.Range("J17", "J18").Merge()
            x1hoja.Range("J17", "J18").Borders.Color = RGB(0, 0, 0)
            x1hoja.Range("J17", "J18").Interior.Color = RGB(192, 192, 192)
            x1hoja.Range("J17", "J18").WrapText = True
            x1hoja.Cells(fila, columna).formula = "x 1.000 cel/mL"
            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
            x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
            x1hoja.Cells(fila, columna).Font.Size = 6
            columna = columna + 1
            x1hoja.Cells(1, 1).RowHeight = 18
            x1hoja.Range("K17", "K18").Merge()
            x1hoja.Range("K17", "K18").Borders.Color = RGB(0, 0, 0)
            x1hoja.Range("K17", "K18").Interior.Color = RGB(192, 192, 192)
            x1hoja.Range("K17", "K18").WrapText = True
            x1hoja.Cells(fila, columna).formula = "g/100mL"
            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
            x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
            x1hoja.Cells(fila, columna).Font.Size = 6
            columna = columna + 1
            x1hoja.Cells(1, 1).RowHeight = 18
            x1hoja.Range("L17", "L18").Merge()
            x1hoja.Range("L17", "L18").Borders.Color = RGB(0, 0, 0)
            x1hoja.Range("L17", "L18").Interior.Color = RGB(192, 192, 192)
            x1hoja.Range("L17", "L18").WrapText = True
            x1hoja.Cells(fila, columna).formula = "g/100mL"
            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
            x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
            x1hoja.Cells(fila, columna).Font.Size = 6
            columna = columna + 1
            x1hoja.Cells(1, 1).RowHeight = 18
            x1hoja.Range("M17", "M18").Merge()
            x1hoja.Range("M17", "M18").Borders.Color = RGB(0, 0, 0)
            x1hoja.Range("M17", "M18").Interior.Color = RGB(192, 192, 192)
            x1hoja.Range("M17", "M18").WrapText = True
            x1hoja.Cells(fila, columna).formula = "g/100mL"
            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
            x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
            x1hoja.Cells(fila, columna).Font.Size = 6
            columna = columna + 1
            x1hoja.Cells(1, 1).RowHeight = 18
            x1hoja.Range("N17", "N18").Merge()
            x1hoja.Range("N17", "N18").Borders.Color = RGB(0, 0, 0)
            x1hoja.Range("N17", "N18").Interior.Color = RGB(192, 192, 192)
            x1hoja.Range("N17", "N18").WrapText = True
            x1hoja.Cells(fila, columna).formula = "mg/dL"
            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
            x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
            x1hoja.Cells(fila, columna).Font.Size = 6
            columna = columna + 1
            If bhb = 0 Then
                x1hoja.Cells(1, 1).RowHeight = 18
                x1hoja.Range("O17", "O18").Merge()
                x1hoja.Range("O17", "O18").Borders.Color = RGB(0, 0, 0)
                x1hoja.Range("O17", "O18").Interior.Color = RGB(192, 192, 192)
                x1hoja.Range("O17", "O18").WrapText = True
                x1hoja.Cells(fila, columna).formula = "g/100mL"
                x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
                x1hoja.Cells(fila, columna).Font.Size = 6
                columna = 1
            Else
                x1hoja.Cells(1, 1).RowHeight = 18
                x1hoja.Range("O17", "O18").Merge()
                x1hoja.Range("O17", "O18").Borders.Color = RGB(0, 0, 0)
                x1hoja.Range("O17", "O18").Interior.Color = RGB(192, 192, 192)
                x1hoja.Range("O17", "O18").WrapText = True
                x1hoja.Cells(fila, columna).formula = "mmol/L"
                x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
                x1hoja.Cells(fila, columna).Font.Size = 6
                columna = 1
            End If
            columna = 9
            fila = fila + 2
            If Not lista2 Is Nothing Then
                If lista2.Count > 0 Then
                    For Each c In lista2
                        If c.MUESTRA <> "" Then
                            x1hoja.Cells(fila, columna).formula = Trim(c.MUESTRA)
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        Else
                            x1hoja.Cells(fila, columna).formula = "-"
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        End If
                        If c.RC = -1 Then
                            x1hoja.Cells(fila, columna).formula = ""
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        Else
                            If c.RC < 30 Then
                                x1hoja.Cells(fila, columna).formula = "30"
                                x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                x1hoja.Cells(fila, columna).Font.Size = 8
                                columna = columna + 1
                            Else
                                x1hoja.Cells(fila, columna).formula = c.RC
                                x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                x1hoja.Cells(fila, columna).Font.Size = 8
                                columna = columna + 1
                            End If
                        End If
                        If c.GRASA = -1 Or c.GRASA = 0 Then
                            columna = columna - 1
                            x1hoja.Cells(fila, columna).formula = ""
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                            x1hoja.Cells(fila, columna).formula = "MUESTRA NO APTA **"
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        Else
                            Dim valgrasa = Val(c.GRASA)
                            If valgrasa < 2.5 Or valgrasa > 5.0 Then
                                x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
                            End If
                            x1hoja.Cells(fila, columna).formula = c.GRASA.ToString("##,##0.00")
                            x1hoja.Cells(fila, columna).NumberFormat = "#.00"
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        End If
                        If c.PROTEINA = -1 Or c.PROTEINA = 0 Then
                            x1hoja.Cells(fila, columna).formula = ""
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        Else
                            Dim valproteina = Val(c.PROTEINA)
                            If valproteina < 2.5 Or valproteina > 4.0 Then
                                x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
                            End If
                            x1hoja.Cells(fila, columna).formula = c.PROTEINA.ToString("##,##0.00")
                            x1hoja.Cells(fila, columna).NumberFormat = "#.00"
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        End If
                        If c.LACTOSA = -1 Or c.LACTOSA = 0 Then
                            x1hoja.Cells(fila, columna).formula = ""
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        Else
                            x1hoja.Cells(fila, columna).formula = c.LACTOSA.ToString("##,##0.00")
                            x1hoja.Cells(fila, columna).NumberFormat = "#.00"
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        End If
                        'Dim cs As New dControlSolicitud
                        'cs.FICHA = idsol
                        'cs = cs.buscar
                        Dim na2 As New dNuevoAnalisis
                        Dim listana2 As New ArrayList
                        Dim _urea As Integer = 0
                        Dim _caseina As Integer = 0
                        listana2 = na2.listarporficha2(idsol)
                        If Not listana2 Is Nothing Then
                            For Each na2 In listana2
                                If na2.ANALISIS = 117 Then
                                    _urea = 1
                                End If
                                If na2.ANALISIS = 157 Then
                                    _caseina = 1
                                End If
                                If na2.ANALISIS = 158 Then
                                    _caseina = 1
                                    _urea = 1
                                End If

                            Next
                        End If
                        'If Not cs Is Nothing Then
                        If _urea = 1 Then
                            If c.UREA = -1 Or c.UREA = 0 Then
                                x1hoja.Cells(fila, columna).formula = "-"
                                x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                x1hoja.Cells(fila, columna).Font.Size = 8
                                columna = columna + 1
                            Else
                                Dim valorurea As Integer
                                valorurea = c.UREA * 0.466
                                If valorurea > 20 Or valorurea < 9 Then
                                    x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
                                End If
                                x1hoja.Cells(fila, columna).formula = FormatNumber(valorurea, 0)
                                x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                x1hoja.Cells(fila, columna).Font.Size = 8
                                columna = columna + 1
                            End If
                        Else
                            x1hoja.Cells(fila, columna).formula = "-"
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        End If
                        If _caseina = 1 Then
                            If c.CASEINA = -1 Or c.CASEINA = 0 Then
                                x1hoja.Cells(fila, columna).formula = "-"
                                x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                x1hoja.Cells(fila, columna).Font.Size = 8
                                columna = columna + 1
                            Else
                                x1hoja.Cells(fila, columna).formula = c.CASEINA.ToString("##,##0.00")
                                x1hoja.Cells(fila, columna).NumberFormat = "#.00"
                                x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                x1hoja.Cells(fila, columna).Font.Size = 8
                                columna = columna + 1
                            End If
                        Else
                            If bhb = 1 Then
                                Dim valbhb As String = ""
                                valbhb = c.BHB.ToString("##,##0.00")
                                If c.BHB < 0.07 Then
                                    valbhb = "<0.07"
                                End If
                                If c.BHB > 0.27 Then
                                    valbhb = ">0.27"
                                End If
                                x1hoja.Cells(fila, columna).formula = valbhb 'c.BHB.ToString("##,##0.00")
                                x1hoja.Cells(fila, columna).NumberFormat = "#.00"
                                x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                x1hoja.Cells(fila, columna).Font.Size = 8
                                columna = columna + 1
                            End If
                            x1hoja.Cells(fila, columna).formula = "-"
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        End If
                        'Else
                        '    x1hoja.Cells(fila, columna).formula = "-"
                        '    x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        '    x1hoja.Cells(fila, columna).Font.Size = 8
                        '    columna = columna + 1
                        'End If
                        na2 = Nothing
                        columna = 9
                        fila = fila + 1
                    Next
                    'Referencias
                    fila = fila + 1
                    columna = 1
                End If
            End If
        End If
        'GUARDA EL ARCHIVO DE EXCEL
        x1app.DisplayAlerts = False 'NO PREGUNTA SI EL ARCHIVO EXISTE
        x1hoja.PageSetup.CenterFooter = "Página &P"
        x1hoja.SaveAs("C:\PREINFORMES\CONTROL\" & "p" & idsol & ".xls")
        x1hoja.SaveAs("C:\PREINFORMES\CONTROL\" & idsol & ".xls")
        'Marcar como creado
        Dim preinf As New dPreinformes
        preinf.FICHA = idsol
        preinf.marcarcreado()
        preinf = Nothing
        x1app.Visible = False
        x1libro.Close()
        x1app = Nothing
        x1libro = Nothing
        x1hoja = Nothing

        Dim proceso As System.Diagnostics.Process()
        proceso = System.Diagnostics.Process.GetProcessesByName("EXCEL")
        For Each opro As System.Diagnostics.Process In proceso
            'antes de iniciar el proceso obtengo la fecha en que inicie el 
            'proceso para detener todos los procesos que excel que inicio
            'mi código durante el proceso
            opro.Kill()
        Next
        Dim v As New FormGraficasRC(idsol)
        v.Show()
    End Sub
    Private Sub preinforme_calidad(ByVal id_sol As Long)
        Dim x1app As Microsoft.Office.Interop.Excel.Application
        Dim x1libro As Microsoft.Office.Interop.Excel.Workbook
        Dim x1hoja As Microsoft.Office.Interop.Excel.Worksheet
        x1app = CType(CreateObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)
        x1libro = CType(x1app.Workbooks.Add, Microsoft.Office.Interop.Excel.Workbook)
        x1hoja = CType(x1libro.Worksheets(1), Microsoft.Office.Interop.Excel.Worksheet)
        Dim csm As New dCalidadSolicitudMuestra
        Dim i As New dIbc
        Dim sa As New dSolicitudAnalisis
        Dim pro As New dCliente
        Dim tec As New dTecnicos
        Dim lista As New ArrayList
        '*****************************
        Dim idsol As Long = id_sol 'TextFicha.Text.Trim
        sa.ID = idsol
        sa = sa.buscar
        '*****************************
        Dim fila As Integer
        Dim columna As Integer
        fila = 1
        columna = 2
        '*********************** ENCABEZADO ************************************************************************
        '***********************************************************************************************************
        x1hoja.Cells(1, 1).columnwidth = 6 '7
        x1hoja.Cells(1, 2).columnwidth = 5
        x1hoja.Cells(1, 3).columnwidth = 5
        x1hoja.Cells(1, 4).columnwidth = 5
        x1hoja.Cells(1, 5).columnwidth = 5
        x1hoja.Cells(1, 6).columnwidth = 5
        x1hoja.Cells(1, 7).columnwidth = 5
        x1hoja.Cells(1, 8).columnwidth = 5
        x1hoja.Cells(1, 9).columnwidth = 5
        x1hoja.Cells(1, 10).columnwidth = 5
        x1hoja.Cells(1, 11).columnwidth = 6 '8
        x1hoja.Cells(1, 12).columnwidth = 6
        x1hoja.Cells(1, 13).columnwidth = 6 '8
        x1hoja.Cells(1, 14).columnwidth = 5
        lista = csm.listarporsolicitud(idsol)
        fila = 15
        columna = 1
        '*** FIN DEL ENCABEZADO************************************************************************************
        '**********************************************************************************************************
        x1hoja.Cells(fila, columna).Formula = "Ident."
        x1hoja.Cells(fila, columna).Font.Bold = True
        x1hoja.Cells(fila, columna).Font.Size = 8
        x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
        x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        columna = columna + 1
        x1hoja.Cells(fila, columna).Formula = "Rc"
        x1hoja.Cells(fila, columna).Font.Bold = True
        x1hoja.Cells(fila, columna).Font.Size = 8
        x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
        x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        columna = columna + 1
        x1hoja.Cells(fila, columna).Formula = "R Bact."
        x1hoja.Cells(fila, columna).Font.Bold = True
        x1hoja.Cells(fila, columna).Font.Size = 8
        x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
        x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        columna = columna + 1
        x1hoja.Cells(fila, columna).Formula = "Gr"
        x1hoja.Cells(fila, columna).Font.Bold = True
        x1hoja.Cells(fila, columna).Font.Size = 8
        x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
        x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        columna = columna + 1
        x1hoja.Cells(fila, columna).Formula = "Pr"
        x1hoja.Cells(fila, columna).Font.Bold = True
        x1hoja.Cells(fila, columna).Font.Size = 8
        x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
        x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        columna = columna + 1
        x1hoja.Cells(fila, columna).Formula = "Lc*"
        x1hoja.Cells(fila, columna).Font.Bold = True
        x1hoja.Cells(fila, columna).Font.Size = 8
        x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
        x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        columna = columna + 1
        x1hoja.Cells(fila, columna).Formula = "ST"
        x1hoja.Cells(fila, columna).Font.Bold = True
        x1hoja.Cells(fila, columna).Font.Size = 8
        x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
        x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        columna = columna + 1
        x1hoja.Cells(fila, columna).Formula = "Cr*"
        x1hoja.Cells(fila, columna).Font.Bold = True
        x1hoja.Cells(fila, columna).Font.Size = 8
        x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
        x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        columna = columna + 1
        x1hoja.Cells(fila, columna).Formula = "MUN*"
        x1hoja.Cells(fila, columna).Font.Bold = True
        x1hoja.Cells(fila, columna).Font.Size = 8
        x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
        x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        columna = columna + 1
        x1hoja.Cells(fila, columna).Formula = "Inh"
        x1hoja.Cells(fila, columna).Font.Bold = True
        x1hoja.Cells(fila, columna).Font.Size = 8
        x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
        x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        columna = columna + 1
        x1hoja.Cells(fila, columna).Formula = "Esp.Ana.*"
        x1hoja.Cells(fila, columna).Font.Bold = True
        x1hoja.Cells(fila, columna).Font.Size = 8
        x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
        x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        columna = columna + 1
        x1hoja.Cells(fila, columna).Formula = "Psicro.*"
        x1hoja.Cells(fila, columna).Font.Bold = True
        x1hoja.Cells(fila, columna).Font.Size = 8
        x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
        x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        columna = columna + 1
        x1hoja.Cells(fila, columna).Formula = "Caseína*"
        x1hoja.Cells(fila, columna).Font.Bold = True
        x1hoja.Cells(fila, columna).Font.Size = 8
        x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
        x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        columna = columna + 1
        x1hoja.Cells(fila, columna).Formula = "Afla.M1*"
        x1hoja.Cells(fila, columna).Font.Bold = True
        x1hoja.Cells(fila, columna).Font.Size = 8
        x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
        x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        columna = 1
        fila = fila + 1
        x1hoja.Cells(1, 1).RowHeight = 18
        x1hoja.Range("A16", "A17").Merge()
        x1hoja.Range("A16", "A17").Borders.Color = RGB(0, 0, 0)
        x1hoja.Range("A16", "A17").Interior.Color = RGB(192, 192, 192)
        x1hoja.Range("A16", "A17").WrapText = True
        x1hoja.Cells(fila, columna).formula = ""
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
        x1hoja.Cells(fila, columna).Font.Size = 6
        columna = columna + 1
        x1hoja.Cells(1, 1).RowHeight = 18
        x1hoja.Range("B16", "B17").Merge()
        x1hoja.Range("B16", "B17").Borders.Color = RGB(0, 0, 0)
        x1hoja.Range("B16", "B17").Interior.Color = RGB(192, 192, 192)
        x1hoja.Range("B16", "B17").WrapText = True
        x1hoja.Cells(fila, columna).formula = "x 1.000 cel/mL"
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
        x1hoja.Cells(fila, columna).Font.Size = 6
        columna = columna + 1
        x1hoja.Cells(1, 1).RowHeight = 18
        x1hoja.Range("C16", "C17").Merge()
        x1hoja.Range("C16", "C17").Borders.Color = RGB(0, 0, 0)
        x1hoja.Range("C16", "C17").Interior.Color = RGB(192, 192, 192)
        x1hoja.Range("C16", "C17").WrapText = True
        x1hoja.Cells(fila, columna).formula = "x 1.000 eq. UFC/ml"
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
        x1hoja.Cells(fila, columna).Font.Size = 6
        columna = columna + 1
        x1hoja.Cells(1, 1).RowHeight = 18
        x1hoja.Range("D16", "D17").Merge()
        x1hoja.Range("D16", "D17").Borders.Color = RGB(0, 0, 0)
        x1hoja.Range("D16", "D17").Interior.Color = RGB(192, 192, 192)
        x1hoja.Range("D16", "D17").WrapText = True
        x1hoja.Cells(fila, columna).formula = "% peso/vol"
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
        x1hoja.Cells(fila, columna).Font.Size = 6
        columna = columna + 1
        x1hoja.Cells(1, 1).RowHeight = 18
        x1hoja.Range("E16", "E17").Merge()
        x1hoja.Range("E16", "E17").Borders.Color = RGB(0, 0, 0)
        x1hoja.Range("E16", "E17").Interior.Color = RGB(192, 192, 192)
        x1hoja.Range("E16", "E17").WrapText = True
        x1hoja.Cells(fila, columna).formula = "% peso/vol"
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
        x1hoja.Cells(fila, columna).Font.Size = 6
        columna = columna + 1
        x1hoja.Cells(1, 1).RowHeight = 18
        x1hoja.Range("F16", "F17").Merge()
        x1hoja.Range("F16", "F17").Borders.Color = RGB(0, 0, 0)
        x1hoja.Range("F16", "F17").Interior.Color = RGB(192, 192, 192)
        x1hoja.Range("F16", "F17").WrapText = True
        x1hoja.Cells(fila, columna).formula = "% peso/vol"
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
        x1hoja.Cells(fila, columna).Font.Size = 6
        columna = columna + 1
        x1hoja.Cells(1, 1).RowHeight = 18
        x1hoja.Range("G16", "G17").Merge()
        x1hoja.Range("G16", "G17").Borders.Color = RGB(0, 0, 0)
        x1hoja.Range("G16", "G17").Interior.Color = RGB(192, 192, 192)
        x1hoja.Range("G16", "G17").WrapText = True
        x1hoja.Cells(fila, columna).formula = "% peso/vol"
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
        x1hoja.Cells(fila, columna).Font.Size = 6
        columna = columna + 1
        x1hoja.Cells(1, 1).RowHeight = 18
        x1hoja.Range("H16", "H17").Merge()
        x1hoja.Range("H16", "H17").Borders.Color = RGB(0, 0, 0)
        x1hoja.Range("H16", "H17").Interior.Color = RGB(192, 192, 192)
        x1hoja.Range("H16", "H17").WrapText = True
        x1hoja.Cells(fila, columna).formula = "(ºC)"
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
        x1hoja.Cells(fila, columna).Font.Size = 6
        columna = columna + 1
        x1hoja.Cells(1, 1).RowHeight = 18
        x1hoja.Range("I16", "I17").Merge()
        x1hoja.Range("I16", "I17").Borders.Color = RGB(0, 0, 0)
        x1hoja.Range("I16", "I17").Interior.Color = RGB(192, 192, 192)
        x1hoja.Range("I16", "I17").WrapText = True
        x1hoja.Cells(fila, columna).formula = "mg/dl"
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
        x1hoja.Cells(fila, columna).Font.Size = 6
        columna = columna + 1
        x1hoja.Cells(1, 1).RowHeight = 18
        x1hoja.Range("J16", "J17").Merge()
        x1hoja.Range("J16", "J17").Borders.Color = RGB(0, 0, 0)
        x1hoja.Range("J16", "J17").Interior.Color = RGB(192, 192, 192)
        x1hoja.Range("J16", "J17").WrapText = True
        x1hoja.Cells(fila, columna).formula = ""
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
        x1hoja.Cells(fila, columna).Font.Size = 6
        columna = columna + 1
        x1hoja.Cells(1, 1).RowHeight = 18
        x1hoja.Range("K16", "K17").Merge()
        x1hoja.Range("K16", "K17").Borders.Color = RGB(0, 0, 0)
        x1hoja.Range("K16", "K17").Interior.Color = RGB(192, 192, 192)
        x1hoja.Range("K16", "K17").WrapText = True
        x1hoja.Cells(fila, columna).formula = "NMP/L"
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
        x1hoja.Cells(fila, columna).Font.Size = 6
        columna = columna + 1
        x1hoja.Cells(1, 1).RowHeight = 18
        x1hoja.Range("L16", "L17").Merge()
        x1hoja.Range("L16", "L17").Borders.Color = RGB(0, 0, 0)
        x1hoja.Range("L16", "L17").Interior.Color = RGB(192, 192, 192)
        x1hoja.Range("L16", "L17").WrapText = True
        x1hoja.Cells(fila, columna).formula = "x 1000 UFC/ml UFC/mL "
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
        x1hoja.Cells(fila, columna).Font.Size = 6
        columna = columna + 1
        x1hoja.Cells(1, 1).RowHeight = 18
        x1hoja.Range("M16", "M17").Merge()
        x1hoja.Range("M16", "M17").Borders.Color = RGB(0, 0, 0)
        x1hoja.Range("M16", "M17").Interior.Color = RGB(192, 192, 192)
        x1hoja.Range("M16", "M17").WrapText = True
        x1hoja.Cells(fila, columna).formula = ""
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
        x1hoja.Cells(fila, columna).Font.Size = 6
        columna = columna + 1
        x1hoja.Cells(1, 1).RowHeight = 18
        x1hoja.Range("N16", "N17").Merge()
        x1hoja.Range("N16", "N17").Borders.Color = RGB(0, 0, 0)
        x1hoja.Range("N16", "N17").Interior.Color = RGB(192, 192, 192)
        x1hoja.Range("N16", "N17").WrapText = True
        x1hoja.Cells(fila, columna).formula = "ppb"
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
        x1hoja.Cells(fila, columna).Font.Size = 6
        columna = 1
        fila = fila + 2
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each csm In lista
                    Dim c As New dCalidad
                    c.FICHA = idsol
                    c.MUESTRA = Trim(csm.MUESTRA)
                    c = c.buscarxfichaxmuestra
                    If csm.MUESTRA <> "" Then
                        x1hoja.Cells(fila, columna).formula = Trim(csm.MUESTRA)
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    Else
                        x1hoja.Cells(fila, columna).formula = "-"
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    End If
                    If csm.RC = 1 Then
                        If Not c Is Nothing Then
                            x1hoja.Cells(fila, columna).formula = c.RC
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                            If c.RC < 100 Then
                                'contador_rc = contador_rc + 1
                            End If
                        Else
                            x1hoja.Cells(fila, columna).formula = "-"
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        End If
                    Else
                        x1hoja.Cells(fila, columna).formula = "-"
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    End If
                    If csm.RB = 1 Then
                        Dim ibc As New dIbc
                        ibc.FICHA = idsol
                        ibc.MUESTRA = Trim(csm.MUESTRA)
                        ibc = ibc.buscarxfichaxmuestra
                        If Not ibc Is Nothing Then
                            x1hoja.Cells(fila, columna).formula = ibc.RB
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        Else
                            x1hoja.Cells(fila, columna).formula = "-"
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        End If
                    Else
                        x1hoja.Cells(fila, columna).formula = "-"
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    End If
                    If csm.COMPOSICION = 1 Or csm.COMPOSICIONSUERO = 1 Then
                        If Not c Is Nothing Then
                            Dim valgrasa As Double = Val(c.GRASA)
                            If valgrasa < 2.5 Or valgrasa > 5 Then
                                x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
                            End If
                            x1hoja.Cells(fila, columna).formula = c.GRASA.ToString("##,##0.00")
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        Else
                            x1hoja.Cells(fila, columna).formula = "-"
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        End If
                    Else
                        x1hoja.Cells(fila, columna).formula = "-"
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    End If
                    If csm.COMPOSICION = 1 Or csm.COMPOSICIONSUERO = 1 Then
                        If Not c Is Nothing Then
                            Dim valproteina As Double = Val(c.PROTEINA)
                            If valproteina < 2.5 Or valproteina > 4 Then
                                x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
                            End If
                            x1hoja.Cells(fila, columna).formula = c.PROTEINA.ToString("##,##0.00")
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1

                        Else
                            x1hoja.Cells(fila, columna).formula = "-"
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        End If
                    Else
                        x1hoja.Cells(fila, columna).formula = "-"
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    End If
                    If csm.COMPOSICION = 1 Or csm.COMPOSICIONSUERO = 1 Then
                        If Not c Is Nothing Then
                            x1hoja.Cells(fila, columna).formula = c.LACTOSA.ToString("##,##0.00")
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        Else
                            x1hoja.Cells(fila, columna).formula = "-"
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        End If
                    Else
                        x1hoja.Cells(fila, columna).formula = "-"
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    End If
                    If csm.COMPOSICION = 1 Or csm.COMPOSICIONSUERO = 1 Then
                        If Not c Is Nothing Then
                            x1hoja.Cells(fila, columna).formula = c.ST.ToString("##,##0.00")
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        Else
                            x1hoja.Cells(fila, columna).formula = "-"
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        End If
                    Else
                        x1hoja.Cells(fila, columna).formula = "-"
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    End If
                    If csm.CRIOSCOPIA = 1 Or csm.CRIOSCOPIA_CRIOSCOPO = 1 Then
                        If Not c Is Nothing Then
                            If c.CRIOSCOPIA <> -1 Then
                                Dim valcrioscopia As Double = Val(c.CRIOSCOPIA) * -1 / 1000
                                If valcrioscopia > -0.51 Or valcrioscopia < -0.54 Then
                                    x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
                                End If
                                x1hoja.Cells(fila, columna).formula = valcrioscopia.ToString("##,###0.000")
                                x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                x1hoja.Cells(fila, columna).Font.Size = 8
                                columna = columna + 1
                            Else
                                x1hoja.Cells(fila, columna).formula = "-"
                                x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                x1hoja.Cells(fila, columna).Font.Size = 8
                                columna = columna + 1
                            End If
                        Else
                            x1hoja.Cells(fila, columna).formula = "-"
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        End If
                    Else
                        x1hoja.Cells(fila, columna).formula = "-"
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    End If
                    If csm.UREA = 1 Then
                        If Not c Is Nothing Then
                            If c.UREA <> -1 Then
                                Dim valorurea As Integer
                                valorurea = c.UREA * 0.466
                                If valorurea > 20 Or valorurea < 9 Then
                                    x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
                                End If
                                x1hoja.Cells(fila, columna).formula = valorurea
                                x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                x1hoja.Cells(fila, columna).Font.Size = 8
                                columna = columna + 1
                            Else
                                x1hoja.Cells(fila, columna).formula = "-"
                                x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                x1hoja.Cells(fila, columna).Font.Size = 8
                                columna = columna + 1
                            End If
                        Else
                            x1hoja.Cells(fila, columna).formula = "-"
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        End If
                    Else
                        x1hoja.Cells(fila, columna).formula = "-"
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    End If
                    Dim inh As New dInhibidores
                    inh.FICHA = idsol
                    inh.MUESTRA = Trim(csm.MUESTRA)
                    inh = inh.buscarxfichaxmuestra
                    If Not inh Is Nothing Then
                        If inh.RESULTADO = 0 Then
                            x1hoja.Cells(fila, columna).formula = "Negativo"
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 6
                            columna = columna + 1
                        Else
                            x1hoja.Cells(fila, columna).formula = "Positivo"
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 6
                            columna = columna + 1
                        End If
                    Else
                        x1hoja.Cells(fila, columna).formula = "-"
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    End If
                    'ESPORULADOS*******************************************************************************
                    Dim esp As New dEsporulados
                    esp.FICHA = idsol
                    esp.MUESTRA = Trim(csm.MUESTRA)
                    esp = esp.buscarxfichaxmuestra
                    If Not esp Is Nothing Then
                        x1hoja.Cells(fila, columna).formula = esp.RESULTADO
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    Else
                        x1hoja.Cells(fila, columna).formula = "-"
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    End If
                    'PSICROTROFOS*******************************************************************************
                    Dim psi As New dPsicrotrofos
                    psi.FICHA = idsol
                    psi.MUESTRA = Trim(csm.MUESTRA)
                    psi = psi.buscarxfichaxmuestra
                    If Not psi Is Nothing Then
                        x1hoja.Cells(fila, columna).formula = psi.PROMEDIO
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    Else
                        x1hoja.Cells(fila, columna).formula = "-"
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    End If
                    If csm.CASEINA = 1 Then
                        If Not c Is Nothing Then
                            If c.CASEINA <> -1 Then
                                Dim valorcaseina As Double
                                valorcaseina = c.CASEINA
                                x1hoja.Cells(fila, columna).formula = valorcaseina
                                x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                x1hoja.Cells(fila, columna).Font.Size = 8
                                columna = columna + 1
                            Else
                                x1hoja.Cells(fila, columna).formula = "-"
                                x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                x1hoja.Cells(fila, columna).Font.Size = 8
                                columna = columna + 1
                            End If
                        Else
                            x1hoja.Cells(fila, columna).formula = "-"
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        End If
                    Else
                        x1hoja.Cells(fila, columna).formula = "-"
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    End If
                    'AFLATOXINA M1*******************************************************************************
                    Dim m As New dMicotoxinasLeche
                    m.FICHA = idsol
                    m.MUESTRA = Trim(csm.MUESTRA)
                    m = m.buscarxfichaxmuestra
                    If Not m Is Nothing Then
                        x1hoja.Cells(fila, columna).formula = m.RESULTADO
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = 1
                    Else
                        x1hoja.Cells(fila, columna).formula = "-"
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = 1
                    End If
                    columna = 1
                    fila = fila + 1
                Next
                'Referencias
                fila = fila + 1
                columna = 1
            End If
        End If
        'PROTEGE LA HOJA DE EXCEL
        'x1hoja.Protect(Password:="1582782", DrawingObjects:=True, _
        'Contents:=True, Scenarios:=True)
        'GUARDA EL ARCHIVO DE EXCEL
        'x1hoja.SaveAs("\\SRVDATOS\datos (d)\NET\PREINFORMES\CALIDAD\" & idsol & ".xls")
        'x1hoja.PageSetup.CenterFooter = "Página &P de &N"
        x1app.DisplayAlerts = False 'NO PREGUNTA SI EL ARCHIVO EXISTE
        x1hoja.PageSetup.CenterFooter = "Página &P"
        'x1hoja.SaveAs("\\192.168.1.10\e\NET\PREINFORMES\CALIDAD\" & idsol & ".xls")
        x1hoja.SaveAs("C:\PREINFORMES\CALIDAD\" & idsol & ".xls")
        'Marcar como creado
        Dim preinf As New dPreinformes
        preinf.FICHA = idsol
        preinf.marcarcreado()
        preinf = Nothing
        x1app.Visible = False
        x1libro.Close()
        x1app = Nothing
        x1libro = Nothing
        x1hoja = Nothing
        creartxt2(idsol)

        Dim proceso As System.Diagnostics.Process()
        proceso = System.Diagnostics.Process.GetProcessesByName("EXCEL")
        For Each opro As System.Diagnostics.Process In proceso
            'antes de iniciar el proceso obtengo la fecha en que inicie el 
            'proceso para detener todos los procesos que excel que inicio
            'mi código durante el proceso
            opro.Kill()
        Next
    End Sub
    Private Sub pre_informe_calidad()
        Dim pi As New dPreinformes
        Dim lista As New ArrayList
        Dim creapreinformecalidad As Integer = 1
        Dim id_sol As Long = 0
        lista = pi.listarsinmarcarcalidad
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each pi In lista
                    Dim csm As New dCalidadSolicitudMuestra
                    Dim listacsm As New ArrayList
                    Dim ficha As Long = pi.FICHA
                    listacsm = csm.listarporsolicitud(ficha)
                    If Not listacsm Is Nothing Then
                        creapreinformecalidad = 1
                        For Each csm In listacsm
                            If csm.RB = 1 Then
                                Dim ibc As New dIbc
                                ibc.FICHA = csm.ficha
                                ibc.MUESTRA = csm.MUESTRA
                                ibc = ibc.buscarxfichaxmuestra
                                If Not ibc Is Nothing Then
                                Else
                                    creapreinformecalidad = 0
                                    Exit For
                                End If
                                ibc = Nothing
                            End If
                            If csm.PSICROTROFOS = 1 Then
                                Dim psi As New dPsicrotrofos
                                psi.FICHA = csm.ficha
                                psi.MUESTRA = csm.MUESTRA
                                psi = psi.buscarxfichaxmuestra
                                If Not psi Is Nothing Then
                                Else
                                    creapreinformecalidad = 0
                                    Exit For
                                End If
                                psi = Nothing
                            End If
                            If csm.ESPORULADOS = 1 Then
                                Dim esp As New dEsporulados
                                esp.FICHA = csm.ficha
                                esp.MUESTRA = csm.MUESTRA
                                esp = esp.buscarxfichaxmuestra
                                If Not esp Is Nothing Then
                                Else
                                    creapreinformecalidad = 0
                                    Exit For
                                End If
                                esp = Nothing
                            End If
                            If csm.INHIBIDORES = 1 Then
                                Dim inh As New dInhibidores
                                inh.FICHA = csm.ficha
                                inh.MUESTRA = csm.MUESTRA
                                inh = inh.buscarxfichaxmuestra
                                If Not inh Is Nothing Then
                                    If inh.MARCA = 0 Then
                                        creapreinformecalidad = 0
                                        Exit For
                                    End If
                                Else
                                    creapreinformecalidad = 0
                                    Exit For
                                End If
                                inh = Nothing
                            End If
                        Next
                    End If
                    If creapreinformecalidad = 1 Then
                        preinforme_calidad(ficha)
                    End If
                Next
            End If
        End If
        pi = Nothing
        lista = Nothing
    End Sub
    Private Sub pre_informe_control()
        Dim pi As New dPreinformes
        Dim lista As New ArrayList
        Dim id_sol As Long = 0
        lista = pi.listarsinmarcarcontrol
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each pi In lista
                    id_sol = pi.FICHA
                    preinforme_control(id_sol)
                Next
            End If
        End If
        pi = Nothing
        lista = Nothing
    End Sub

#End Region

#Region "SUBIR-INFORMES"

    Private Sub subir_informes()
        Dim picalidad As New dPreinformeCalidad
        Dim picontrol As New dPreinformeControl
        Dim listacalidad As New ArrayList
        Dim listacontrol As New ArrayList
        listacalidad = picalidad.listarparasubir
        listacontrol = picontrol.listarparasubir
        If Not listacalidad Is Nothing Then
            If listacalidad.Count > 0 Then
                For Each picalidad In listacalidad
                    idficha = picalidad.FICHA
                    enviar_copia = picalidad.COPIA
                    _abonado = picalidad.ABONADO
                    _comentarios = picalidad.COMENTARIO
                    Dim sa As New dSolicitudAnalisis
                    sa.ID = idficha
                    sa = sa.buscar
                    If Not sa Is Nothing Then
                        Dim p As New dCliente
                        tipoinforme = sa.IDTIPOINFORME
                        p.ID = sa.IDPRODUCTOR
                        p = p.buscar
                        If Not p Is Nothing Then
                            email = ""
                            email2 = ""
                            If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                email = RTrim(p.NOT_EMAIL_ANALISIS1)
                            End If
                            If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                            End If
                            If email = "" Then
                                If email2 = "" Then
                                    If p.EMAIL <> "" Then
                                        email = RTrim(p.EMAIL)
                                    End If
                                Else
                                    email = email2
                                End If
                            Else
                                If email2 = "" Then
                                    email = email
                                Else
                                    email = email & "," & email2
                                End If
                            End If
                            productorweb_com = p.USUARIO_WEB
                            Dim pw_com As New dProductorWeb_com
                            pw_com.USUARIO = productorweb_com
                            pw_com = pw_com.buscar
                            If Not pw_com Is Nothing Then
                                idproductorweb_com = pw_com.ID
                                'email = RTrim(pw_com.ENVIAR_EMAIL)
                                'email = Replace(pw_com.ENVIAR_EMAIL, " ", "")
                                celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                carpeta = idproductorweb_com
                                crea_carpeta()
                            End If
                            sa = Nothing
                        End If
                    End If
controlexcel:
                    If nombre_pc = "ROBOT" Then
                        subirFicheroXls()
                    End If
                    existeXls()
                    If excel = 1 Then
                        GoTo controlexcel
                    End If
                    moverexcel()
controlpdf:
                    If nombre_pc = "ROBOT" Then
                        subirFicheroPdf()
                    End If
                    existePdf()
                    If pdf = 1 Then
                        GoTo controlpdf
                    End If
                    moverpdf()
                    modificarRegistro()
                    Dim fechaactual As Date = Now()
                    Dim _fecha As String
                    _fecha = Format(fechaactual, "yyyy-MM-dd")
                    picalidad.marcarsubido(_fecha)

                    If tipoinforme = 15 Then
                        enviaremail()
                    End If
                    Dim s As New dSolicitudAnalisis
                    Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                    Dim fecenv As String
                    fecenv = Format(fechaenvio, "yyyy-MM-dd")
                    s.ID = idficha
                    s.actualizarfechaenvio2(fecenv)
                    s.marcar2()
                    s = Nothing
                    ' Grabar estado de la ficha
                    Dim est As New dEstados
                    est.FICHA = idficha
                    est.ESTADO = 8
                    est.FECHA = fecenv
                    est.guardar2()
                    est = Nothing
                    '****************************
                Next
            End If
        End If
        If Not listacontrol Is Nothing Then
            If listacontrol.Count > 0 Then
                For Each picontrol In listacontrol
                    idficha = picontrol.FICHA
                    If idficha = 38 Then
                        Dim tr As Boolean = True
                    End If
                    enviar_copia = picontrol.COPIA
                    _abonado = picontrol.ABONADO
                    _comentarios = picontrol.COMENTARIO
                    Dim sa As New dSolicitudAnalisis
                    sa.ID = idficha
                    sa = sa.buscar
                    If Not sa Is Nothing Then
                        Dim p As New dCliente
                        tipoinforme = sa.IDTIPOINFORME
                        p.ID = sa.IDPRODUCTOR
                        p = p.buscar
                        If Not p Is Nothing Then
                            email = ""
                            email2 = ""
                            If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                email = RTrim(p.NOT_EMAIL_ANALISIS1)
                            End If
                            If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                            End If
                            If email = "" Then
                                If email2 = "" Then
                                    If p.EMAIL <> "" Then
                                        email = RTrim(p.EMAIL)
                                    End If
                                Else
                                    email = email2
                                End If
                            Else
                                If email2 = "" Then
                                    email = email
                                Else
                                    email = email & "," & email2
                                End If
                            End If
                            productorweb_com = p.USUARIO_WEB
                            Dim pw_com As New dProductorWeb_com
                            pw_com.USUARIO = productorweb_com
                            pw_com = pw_com.buscar
                            If Not pw_com Is Nothing Then
                                idproductorweb_com = pw_com.ID
                                'email = RTrim(pw_com.ENVIAR_EMAIL)
                                celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                carpeta = idproductorweb_com
                                crea_carpeta()
                            End If
                            sa = Nothing
                        End If
                    End If
controlexcel2:      If nombre_pc = "ROBOT" Then
                        subirFicheroXls()
                    End If
                    existeXls()
                    If excel = 1 Then
                        GoTo controlexcel2
                    End If
                    moverexcel()
controlpdf2:
                    If nombre_pc = "ROBOT" Then
                        subirFicheroPdf()
                    End If
                    existePdf()
                    If pdf = 1 Then
                        GoTo controlpdf2
                    End If
                    moverpdf()
controlcsv2:
                    If nombre_pc = "ROBOT" Then
                        'subirFicheroCsv()
                    End If
                    'existeCsv()
                    'If csv = 1 Then
                    'GoTo controlcsv2
                    'End If
                    'movertxt()
                    modificarRegistro()
                    Dim fechaactual As Date = Now()
                    Dim _fecha As String
                    _fecha = Format(fechaactual, "yyyy-MM-dd")
                    picontrol.marcarsubido(_fecha)
                    If tipoinforme = 15 Then
                        enviaremail()
                    End If
                    Dim s As New dSolicitudAnalisis
                    Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                    Dim fecenv As String
                    fecenv = Format(fechaenvio, "yyyy-MM-dd")
                    s.ID = idficha
                    s.actualizarfechaenvio2(fecenv)
                    s.marcar2()
                    s = Nothing
                    ' Grabar estado de la ficha
                    Dim est As New dEstados
                    est.FICHA = idficha
                    est.ESTADO = 8
                    est.FECHA = fecenv
                    est.guardar2()
                    est = Nothing
                    '****************************
                Next
            End If
        End If
    End Sub
    Private Sub subir_informes2()
        Dim pi As New dPreinformes
        Dim lista As New ArrayList
        Dim subidoxls As Integer = 0
        Dim subidopdf As Integer = 0
        Dim subidotxt As Integer = 0
        lista = pi.listarparasubir
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each pi In lista
                    Dim cifq As New dControlInformesFQ
                    Dim ficha As Long = 0
                    cifq.FICHA = pi.FICHA
                    cifq = cifq.buscarxficha
                    If Not cifq Is Nothing Then
                        If cifq.CONTROLADO = 1 Then
                            idficha = pi.FICHA
                            _tipoinforme = pi.TIPO
                            enviar_copia = pi.COPIA
                            _abonado = pi.ABONADO
                            _comentarios = pi.COMENTARIO
                            Dim sa As New dSolicitudAnalisis
                            sa.ID = idficha
                            sa = sa.buscar
                            If Not sa Is Nothing Then
                                Dim p As New dCliente
                                tipoinforme = sa.IDTIPOINFORME
                                p.ID = sa.IDPRODUCTOR
                                p = p.buscar
                                If Not p Is Nothing Then
                                    email = ""
                                    email2 = ""
                                    If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                        email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                    End If
                                    If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                        email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                    End If
                                    If email = "" Then
                                        If email2 = "" Then
                                            If p.EMAIL <> "" Then
                                                email = RTrim(p.EMAIL)
                                            End If
                                        Else
                                            email = email2
                                        End If
                                    Else
                                        If email2 = "" Then
                                            email = email
                                        Else
                                            email = email & "," & email2
                                        End If
                                    End If
                                    productorweb_com = p.USUARIO_WEB
                                    Dim pw_com As New dProductorWeb_com
                                    pw_com.USUARIO = productorweb_com
                                    pw_com = pw_com.buscar
                                    If Not pw_com Is Nothing Then
                                        idproductorweb_com = pw_com.ID
                                        'email = RTrim(pw_com.ENVIAR_EMAIL)
                                        celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                        carpeta = idproductorweb_com
                                        crea_carpeta()
                                    End If
                                    sa = Nothing
                                End If
                            End If
controlexcel:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroXls()
                            End If
                            existeXls()
                            If excel = 1 Then
                                GoTo controlexcel
                            End If
                            subidoxls = 1
controlpdf:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroPdf()
                            End If
                            existePdf()
                            If pdf = 1 Then
                                GoTo controlpdf
                            End If
                            subidopdf = 1
                            If pi.TIPO = 1 Then
controltxt:
                                If nombre_pc = "ROBOT" Then
                                    subirFicheroCsv()
                                End If
                                existeCsv()
                                If csv = 1 Then
                                    GoTo controltxt
                                End If
                                movertxt()
                            End If
                            modificarRegistro()
                            Dim fechaactual As Date = Now()
                            Dim _fecha As String
                            _fecha = Format(fechaactual, "yyyy-MM-dd")
                            pi.marcarsubido(_fecha)
                            If tipoinforme = 15 Then
                                enviaremail()
                            End If
                            Dim s As New dSolicitudAnalisis
                            Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                            Dim fecenv As String
                            fecenv = Format(fechaenvio, "yyyy-MM-dd")
                            s.ID = idficha
                            s.actualizarfechaenvio2(fecenv)
                            s.marcar2()
                            s = Nothing
                            If subidoxls = 1 And subidopdf = 1 Then
                                If nombre_pc = "ROBOT" Then
                                    moverexcel()
                                    moverpdf()
                                Else
                                    moverexcel_otrapc()
                                    moverpdf_otrapc()
                                End If
                                ' Grabar estado de la ficha
                                Dim est As New dEstados
                                est.FICHA = idficha
                                est.ESTADO = 8
                                est.FECHA = fecenv
                                est.guardar2()
                                est = Nothing
                                '****************************
                            End If
                        Else
                        End If
                    Else
                    End If
                    cifq = Nothing
                Next
            End If
        End If
    End Sub
    Private Sub subir_informes2b()
        Dim pi As New dPreinformes
        Dim lista As New ArrayList
        Dim subidoxls As Integer = 0
        Dim subidopdf As Integer = 0
        Dim subidotxt As Integer = 0
        lista = pi.listarparasubirmicro
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each pi In lista
                    Dim cimicro As New dControlInformesMicro
                    Dim ficha As Long = 0
                    cimicro.FICHA = pi.FICHA
                    cimicro = cimicro.buscarxficha
                    If Not cimicro Is Nothing Then
                        If cimicro.CONTROLADO = 1 Then
                            idficha = pi.FICHA
                            _tipoinforme = pi.TIPO
                            enviar_copia = pi.COPIA
                            _abonado = pi.ABONADO
                            _comentarios = pi.COMENTARIO
                            Dim sa As New dSolicitudAnalisis
                            sa.ID = idficha
                            sa = sa.buscar
                            If Not sa Is Nothing Then
                                Dim p As New dCliente
                                tipoinforme = sa.IDTIPOINFORME
                                p.ID = sa.IDPRODUCTOR
                                p = p.buscar
                                If Not p Is Nothing Then
                                    email = ""
                                    email2 = ""
                                    If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                        email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                    End If
                                    If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                        email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                    End If
                                    If email = "" Then
                                        If email2 = "" Then
                                            If p.EMAIL <> "" Then
                                                email = RTrim(p.EMAIL)
                                            End If
                                        Else
                                            email = email2
                                        End If
                                    Else
                                        If email2 = "" Then
                                            email = email
                                        Else
                                            email = email & "," & email2
                                        End If
                                    End If
                                    productorweb_com = p.USUARIO_WEB
                                    Dim pw_com As New dProductorWeb_com
                                    pw_com.USUARIO = productorweb_com
                                    pw_com = pw_com.buscar
                                    If Not pw_com Is Nothing Then
                                        idproductorweb_com = pw_com.ID
                                        'email = RTrim(pw_com.ENVIAR_EMAIL)
                                        celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                        carpeta = idproductorweb_com
                                        crea_carpeta()
                                    End If
                                    sa = Nothing
                                End If
                            End If
controlexcel:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroXls()
                            End If
                            existeXls()
                            If excel = 1 Then
                                GoTo controlexcel
                            End If
                            subidoxls = 1
controlpdf:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroPdf()
                            End If
                            existePdf()
                            If pdf = 1 Then
                                GoTo controlpdf
                            End If
                            subidopdf = 1

                            If pi.TIPO = 1 Then
controltxt:
                                If nombre_pc = "ROBOT" Then
                                    subirFicheroCsv()
                                End If
                                existeCsv()
                                If csv = 1 Then
                                    GoTo controltxt
                                End If
                                movertxt()
                            End If
                            modificarRegistro()
                            Dim fechaactual As Date = Now()
                            Dim _fecha As String
                            _fecha = Format(fechaactual, "yyyy-MM-dd")
                            pi.marcarsubido(_fecha)
                            If tipoinforme = 15 Then
                                enviaremail()
                            End If
                            Dim s As New dSolicitudAnalisis
                            Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                            Dim fecenv As String
                            fecenv = Format(fechaenvio, "yyyy-MM-dd")
                            s.ID = idficha
                            s.actualizarfechaenvio2(fecenv)
                            s.marcar2()
                            s = Nothing
                            If subidoxls = 1 And subidopdf = 1 Then
                                If nombre_pc = "ROBOT" Then
                                    moverexcel()
                                    moverpdf()
                                Else
                                    moverexcel_otrapc()
                                    moverpdf_otrapc()
                                End If
                            End If
                        Else
                        End If
                    Else
                    End If
                    cimicro = Nothing
                Next
            End If
        End If
    End Sub
    Private Sub subir_informes_control()
        Dim pi As New dPreinformes
        Dim lista As New ArrayList
        Dim subidoxls As Integer = 0
        Dim subidopdf As Integer = 0
        Dim subidotxt As Integer = 0
        lista = pi.listarparasubircontrol
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each pi In lista
                    Dim cifq As New dControlInformesFQ
                    Dim ficha As Long = 0
                    cifq.FICHA = pi.FICHA
                    cifq = cifq.buscarxficha
                    If Not cifq Is Nothing Then
                        If cifq.CONTROLADO = 1 Then
                            idficha = pi.FICHA
                            _tipoinforme = pi.TIPO
                            enviar_copia = pi.COPIA
                            _abonado = pi.ABONADO
                            _comentarios = pi.COMENTARIO
                            Dim sa As New dSolicitudAnalisis
                            sa.ID = idficha
                            sa = sa.buscar
                            If Not sa Is Nothing Then
                                Dim p As New dCliente
                                tipoinforme = sa.IDTIPOINFORME
                                p.ID = sa.IDPRODUCTOR
                                p = p.buscar
                                If Not p Is Nothing Then
                                    email = ""
                                    email2 = ""
                                    If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                        email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                    End If
                                    If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                        email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                    End If
                                    If email = "" Then
                                        If email2 = "" Then
                                            If p.EMAIL <> "" Then
                                                email = RTrim(p.EMAIL)
                                            End If
                                        Else
                                            email = email2
                                        End If
                                    Else
                                        If email2 = "" Then
                                            email = email
                                        Else
                                            email = email & "," & email2
                                        End If
                                    End If
                                    productorweb_com = p.USUARIO_WEB
                                    Dim pw_com As New dProductorWeb_com
                                    pw_com.USUARIO = productorweb_com
                                    pw_com = pw_com.buscar
                                    If Not pw_com Is Nothing Then
                                        idproductorweb_com = pw_com.ID
                                        'email = RTrim(pw_com.ENVIAR_EMAIL)
                                        celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                        carpeta = idproductorweb_com
                                        crea_carpeta()
                                    End If
                                    sa = Nothing
                                End If
                            End If
controlexcel:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroXls()
                            End If
                            existeXls()
                            If excel = 1 Then
                                GoTo controlexcel
                            End If
                            subidoxls = 1
controlpdf:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroPdf()
                            End If
                            existePdf()
                            If pdf = 1 Then
                                GoTo controlpdf
                            End If
                            subidopdf = 1
                            If pi.TIPO = 1 Then
controltxt:
                                If nombre_pc = "ROBOT" Then
                                    subirFicheroCsv()
                                End If
                                existeCsv()

                                If csv = 1 Then
                                    GoTo controltxt2
                                End If
                                If nombre_pc = "ROBOT" Then
                                    movertxt()
                                Else
                                    movertxt_otrapc()
                                End If
                        End If
                        modificarRegistro()
                        Dim fechaactual As Date = Now()
                        Dim _fecha As String
                        _fecha = Format(fechaactual, "yyyy-MM-dd")
                        pi.marcarsubido(_fecha)
                        If tipoinforme = 15 Then
                            enviaremail()
                        End If
                        Dim s As New dSolicitudAnalisis
                        Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                        Dim fecenv As String
                        fecenv = Format(fechaenvio, "yyyy-MM-dd")
                        s.ID = idficha
                        s.actualizarfechaenvio2(fecenv)
                        s.marcar2()
                        s = Nothing
                        If subidoxls = 1 And subidopdf = 1 Then
                            If nombre_pc = "ROBOT" Then
                                moverexcel()
                                moverpdf()
                            Else
                                moverexcel_otrapc()
                                moverpdf_otrapc()
                            End If
                            ' Grabar estado de la ficha
                            Dim est As New dEstados
                            est.FICHA = idficha
                            est.ESTADO = 8
                            est.FECHA = fecenv
                            est.guardar2()
                            est = Nothing
                            '****************************
                            envio_mail_informes()
                        End If
                    Else
                    End If
                    Else
                    idficha = pi.FICHA
                    _tipoinforme = pi.TIPO
                    enviar_copia = pi.COPIA
                    _abonado = pi.ABONADO
                    _comentarios = pi.COMENTARIO
                    Dim sa As New dSolicitudAnalisis
                    sa.ID = idficha
                    sa = sa.buscar
                    If Not sa Is Nothing Then
                        Dim p As New dCliente
                        tipoinforme = sa.IDTIPOINFORME
                        p.ID = sa.IDPRODUCTOR
                        p = p.buscar
                        If Not p Is Nothing Then
                            email = ""
                            email2 = ""
                            If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                email = RTrim(p.NOT_EMAIL_ANALISIS1)
                            End If
                            If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                            End If
                            If email = "" Then
                                If email2 = "" Then
                                    If p.EMAIL <> "" Then
                                        email = RTrim(p.EMAIL)
                                    End If
                                Else
                                    email = email2
                                End If
                            Else
                                If email2 = "" Then
                                    email = email
                                Else
                                    email = email & "," & email2
                                End If
                            End If
                            productorweb_com = p.USUARIO_WEB
                            Dim pw_com As New dProductorWeb_com
                            pw_com.USUARIO = productorweb_com
                            pw_com = pw_com.buscar
                            If Not pw_com Is Nothing Then
                                idproductorweb_com = pw_com.ID
                                'email = RTrim(pw_com.ENVIAR_EMAIL)
                                celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                carpeta = idproductorweb_com
                                crea_carpeta()
                            End If
                            sa = Nothing
                        End If
                    End If
controlexcel2:
                    If nombre_pc = "ROBOT" Then
                        subirFicheroXls()
                    End If
                    existeXls()
                    If excel = 1 Then
                        GoTo controlexcel2
                    End If
                    subidoxls = 1
controlpdf2:
                    If nombre_pc = "ROBOT" Then
                        subirFicheroPdf()
                    End If
                    existePdf()
                    If pdf = 1 Then
                        GoTo controlpdf2
                    End If
                    subidopdf = 1
                    If pi.TIPO = 1 Then
controltxt2:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroCsv()
                            End If
                            existeCsv()

                            If csv = 1 Then
                                GoTo controltxt2
                            End If
                            If nombre_pc = "ROBOT" Then
                                movertxt()
                            Else
                                movertxt_otrapc()
                            End If
                    End If
                    modificarRegistro()
                    Dim fechaactual As Date = Now()
                    Dim _fecha As String
                    _fecha = Format(fechaactual, "yyyy-MM-dd")
                    pi.marcarsubido(_fecha)
                    If tipoinforme = 15 Then
                        enviaremail()
                    End If
                    Dim s As New dSolicitudAnalisis
                    Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                    Dim fecenv As String
                    fecenv = Format(fechaenvio, "yyyy-MM-dd")
                    s.ID = idficha
                    s.actualizarfechaenvio2(fecenv)
                    s.marcar2()
                    s = Nothing
                    If subidoxls = 1 And subidopdf = 1 Then
                        If nombre_pc = "ROBOT" Then
                            moverexcel()
                            moverpdf()
                        Else
                            moverexcel_otrapc()
                            moverpdf_otrapc()
                        End If
                        ' Grabar estado de la ficha
                        Dim est As New dEstados
                        est.FICHA = idficha
                        est.ESTADO = 8
                        est.FECHA = fecenv
                        est.guardar2()
                        est = Nothing
                        '****************************
                        envio_mail_informes()
                    End If
                    End If
            cifq = Nothing
                Next
            End If
        End If
    End Sub
    Private Sub subir_informes_agua()
        Dim pi As New dPreinformes
        Dim lista As New ArrayList
        Dim subidoxls As Integer = 0
        Dim subidopdf As Integer = 0
        Dim subidotxt As Integer = 0
        lista = pi.listarparasubiragua
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each pi In lista
                    Dim cimicro As New dControlInformesMicro
                    Dim ficha As Long = 0
                    cimicro.FICHA = pi.FICHA
                    cimicro = cimicro.buscarxficha
                    If Not cimicro Is Nothing Then
                        If cimicro.CONTROLADO = 1 Then
                            idficha = pi.FICHA
                            _tipoinforme = pi.TIPO
                            enviar_copia = pi.COPIA
                            _abonado = pi.ABONADO
                            _comentarios = pi.COMENTARIO
                            Dim sa As New dSolicitudAnalisis
                            sa.ID = idficha
                            sa = sa.buscar
                            If Not sa Is Nothing Then
                                Dim p As New dCliente
                                tipoinforme = sa.IDTIPOINFORME
                                p.ID = sa.IDPRODUCTOR
                                p = p.buscar
                                If Not p Is Nothing Then
                                    email = ""
                                    email2 = ""
                                    If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                        email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                    End If
                                    If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                        email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                    End If
                                    If email = "" Then
                                        If email2 = "" Then
                                            If p.EMAIL <> "" Then
                                                email = RTrim(p.EMAIL)
                                            End If
                                        Else
                                            email = email2
                                        End If
                                    Else
                                        If email2 = "" Then
                                            email = email
                                        Else
                                            email = email & "," & email2
                                        End If
                                    End If
                                    productorweb_com = p.USUARIO_WEB
                                    Dim pw_com As New dProductorWeb_com
                                    pw_com.USUARIO = productorweb_com
                                    pw_com = pw_com.buscar
                                    If Not pw_com Is Nothing Then
                                        idproductorweb_com = pw_com.ID
                                        'email = RTrim(pw_com.ENVIAR_EMAIL)
                                        celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                        carpeta = idproductorweb_com
                                        crea_carpeta()
                                    End If
                                    sa = Nothing
                                End If
                            End If
controlexcel:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroXls()
                            End If
                            existeXls()
                            If excel = 1 Then
                                GoTo controlexcel
                            End If
                            subidoxls = 1
controlpdf:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroPdf()
                            End If
                            existePdf()
                            If pdf = 1 Then
                                GoTo controlpdf
                            End If
                            subidopdf = 1
                            If pi.TIPO = 1 Then
controltxt:
                                'If nombre_pc = "ROBOT" Then
                                '    subirFicheroCsv()
                                'End If
                                'existeCsv()
                                'If csv = 1 Then
                                '    GoTo controltxt
                                'End If
                                'movertxt()
                            End If
                            modificarRegistro()
                            Dim fechaactual As Date = Now()
                            Dim _fecha As String
                            _fecha = Format(fechaactual, "yyyy-MM-dd")
                            pi.marcarsubido(_fecha)
                            If tipoinforme = 15 Then
                                enviaremail()
                            End If
                            Dim s As New dSolicitudAnalisis
                            Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                            Dim fecenv As String
                            fecenv = Format(fechaenvio, "yyyy-MM-dd")
                            s.ID = idficha
                            s.actualizarfechaenvio2(fecenv)
                            s.marcar2()
                            s = Nothing
                            If subidoxls = 1 And subidopdf = 1 Then
                                If nombre_pc = "ROBOT" Then
                                    moverexcel()
                                    moverpdf()
                                Else
                                    moverexcel_otrapc()
                                    moverpdf_otrapc()
                                End If
                                ' Grabar estado de la ficha
                                Dim est As New dEstados
                                est.FICHA = idficha
                                est.ESTADO = 8
                                est.FECHA = fecenv
                                est.guardar2()
                                est = Nothing
                                '****************************
                                envio_mail_informes()
                            End If
                        Else
                        End If
                    Else
                        idficha = pi.FICHA
                        _tipoinforme = pi.TIPO
                        enviar_copia = pi.COPIA
                        _abonado = pi.ABONADO
                        _comentarios = pi.COMENTARIO
                        Dim sa As New dSolicitudAnalisis
                        sa.ID = idficha
                        sa = sa.buscar
                        If Not sa Is Nothing Then
                            Dim p As New dCliente
                            tipoinforme = sa.IDTIPOINFORME
                            p.ID = sa.IDPRODUCTOR
                            p = p.buscar
                            If Not p Is Nothing Then
                                email = ""
                                email2 = ""
                                If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                    email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                End If
                                If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                    email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                End If
                                If email = "" Then
                                    If email2 = "" Then
                                        If p.EMAIL <> "" Then
                                            email = RTrim(p.EMAIL)
                                        End If
                                    Else
                                        email = email2
                                    End If
                                Else
                                    If email2 = "" Then
                                        email = email
                                    Else
                                        email = email & "," & email2
                                    End If
                                End If
                                productorweb_com = p.USUARIO_WEB
                                Dim pw_com As New dProductorWeb_com
                                pw_com.USUARIO = productorweb_com
                                pw_com = pw_com.buscar
                                If Not pw_com Is Nothing Then
                                    idproductorweb_com = pw_com.ID
                                    'email = RTrim(pw_com.ENVIAR_EMAIL)
                                    celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                    carpeta = idproductorweb_com
                                    crea_carpeta()
                                End If
                                sa = Nothing
                            End If
                        End If
controlexcel2:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroXls()
                        End If
                        existeXls()
                        If excel = 1 Then
                            GoTo controlexcel2
                        End If
                        subidoxls = 1
controlpdf2:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroPdf()
                        End If
                        existePdf()
                        If pdf = 1 Then
                            GoTo controlpdf2
                        End If
                        subidopdf = 1
                        If pi.TIPO = 1 Then
controltxt2:
                            'If nombre_pc = "ROBOT" Then
                            '    subirFicheroCsv()
                            'End If
                            'existeCsv()
                            'If csv = 1 Then
                            '    GoTo controltxt2
                            'End If
                            'movertxt()
                        End If
                        modificarRegistro()
                        Dim fechaactual As Date = Now()
                        Dim _fecha As String
                        _fecha = Format(fechaactual, "yyyy-MM-dd")
                        pi.marcarsubido(_fecha)
                        If tipoinforme = 15 Then
                            enviaremail()
                        End If
                        Dim s As New dSolicitudAnalisis
                        Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                        Dim fecenv As String
                        fecenv = Format(fechaenvio, "yyyy-MM-dd")
                        s.ID = idficha
                        s.actualizarfechaenvio2(fecenv)
                        s.marcar2()
                        s = Nothing
                        If subidoxls = 1 And subidopdf = 1 Then
                            If nombre_pc = "ROBOT" Then
                                moverexcel()
                                moverpdf()
                            Else
                                moverexcel_otrapc()
                                moverpdf_otrapc()
                            End If
                            ' Grabar estado de la ficha
                            Dim est As New dEstados
                            est.FICHA = idficha
                            est.ESTADO = 8
                            est.FECHA = fecenv
                            est.guardar2()
                            est = Nothing
                            '****************************
                            envio_mail_informes()
                        End If
                    End If
                    cimicro = Nothing
                Next
            End If
        End If
    End Sub
    Private Sub subir_informes_efluentes()
        Dim pi As New dPreinformes
        Dim lista As New ArrayList
        Dim subidoxls As Integer = 0
        Dim subidopdf As Integer = 0
        Dim subidotxt As Integer = 0
        lista = pi.listarparasubirefluentes
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each pi In lista
                    Dim ciefluentes As New dControlInformesEfluentes
                    Dim ficha As Long = 0
                    ciefluentes.FICHA = pi.FICHA
                    ciefluentes = ciefluentes.buscarxficha
                    If Not ciefluentes Is Nothing Then
                        If ciefluentes.CONTROLADO = 1 Then
                            idficha = pi.FICHA
                            _tipoinforme = pi.TIPO
                            enviar_copia = pi.COPIA
                            _abonado = pi.ABONADO
                            _comentarios = pi.COMENTARIO
                            Dim sa As New dSolicitudAnalisis
                            sa.ID = idficha
                            sa = sa.buscar
                            If Not sa Is Nothing Then
                                Dim p As New dCliente
                                tipoinforme = sa.IDTIPOINFORME
                                p.ID = sa.IDPRODUCTOR
                                p = p.buscar
                                If Not p Is Nothing Then
                                    email = ""
                                    email2 = ""
                                    If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                        email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                    End If
                                    If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                        email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                    End If
                                    If email = "" Then
                                        If email2 = "" Then
                                            If p.EMAIL <> "" Then
                                                email = RTrim(p.EMAIL)
                                            End If
                                        Else
                                            email = email2
                                        End If
                                    Else
                                        If email2 = "" Then
                                            email = email
                                        Else
                                            email = email & "," & email2
                                        End If
                                    End If
                                    productorweb_com = p.USUARIO_WEB
                                    Dim pw_com As New dProductorWeb_com
                                    pw_com.USUARIO = productorweb_com
                                    pw_com = pw_com.buscar
                                    If Not pw_com Is Nothing Then
                                        idproductorweb_com = pw_com.ID
                                        'email = RTrim(pw_com.ENVIAR_EMAIL)
                                        celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                        carpeta = idproductorweb_com
                                        crea_carpeta()
                                    End If
                                    sa = Nothing
                                End If
                            End If
controlexcel:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroXls()
                            End If
                            existeXls()
                            If excel = 1 Then
                                GoTo controlexcel
                            End If
                            subidoxls = 1
controlpdf:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroPdf()
                            End If
                            existePdf()
                            If pdf = 1 Then
                                GoTo controlpdf
                            End If
                            subidopdf = 1
                            If pi.TIPO = 1 Then
controltxt:
                                'If nombre_pc = "ROBOT" Then
                                '    subirFicheroCsv()
                                'End If
                                'existeCsv()
                                'If csv = 1 Then
                                '    GoTo controltxt
                                'End If
                                'movertxt()
                            End If
                            modificarRegistro()
                            Dim fechaactual As Date = Now()
                            Dim _fecha As String
                            _fecha = Format(fechaactual, "yyyy-MM-dd")
                            pi.marcarsubido(_fecha)
                            If tipoinforme = 15 Then
                                enviaremail()
                            End If
                            Dim s As New dSolicitudAnalisis
                            Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                            Dim fecenv As String
                            fecenv = Format(fechaenvio, "yyyy-MM-dd")
                            s.ID = idficha
                            s.actualizarfechaenvio2(fecenv)
                            s.marcar2()
                            s = Nothing
                            If subidoxls = 1 And subidopdf = 1 Then
                                If nombre_pc = "ROBOT" Then
                                    moverexcel()
                                    moverpdf()
                                Else
                                    moverexcel_otrapc()
                                    moverpdf_otrapc()
                                End If
                                ' Grabar estado de la ficha
                                Dim est As New dEstados
                                est.FICHA = idficha
                                est.ESTADO = 8
                                est.FECHA = fecenv
                                est.guardar2()
                                est = Nothing
                                '****************************
                                envio_mail_informes()
                            End If
                        Else
                        End If
                    Else
                        idficha = pi.FICHA
                        _tipoinforme = pi.TIPO
                        enviar_copia = pi.COPIA
                        _abonado = pi.ABONADO
                        _comentarios = pi.COMENTARIO
                        Dim sa As New dSolicitudAnalisis
                        sa.ID = idficha
                        sa = sa.buscar
                        If Not sa Is Nothing Then
                            Dim p As New dCliente
                            tipoinforme = sa.IDTIPOINFORME
                            p.ID = sa.IDPRODUCTOR
                            p = p.buscar
                            If Not p Is Nothing Then
                                email = ""
                                email2 = ""
                                If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                    email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                End If
                                If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                    email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                End If
                                If email = "" Then
                                    If email2 = "" Then
                                        If p.EMAIL <> "" Then
                                            email = RTrim(p.EMAIL)
                                        End If
                                    Else
                                        email = email2
                                    End If
                                Else
                                    If email2 = "" Then
                                        email = email
                                    Else
                                        email = email & "," & email2
                                    End If
                                End If
                                productorweb_com = p.USUARIO_WEB
                                Dim pw_com As New dProductorWeb_com
                                pw_com.USUARIO = productorweb_com
                                pw_com = pw_com.buscar
                                If Not pw_com Is Nothing Then
                                    idproductorweb_com = pw_com.ID
                                    'email = RTrim(pw_com.ENVIAR_EMAIL)
                                    celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                    carpeta = idproductorweb_com
                                    crea_carpeta()
                                End If
                                sa = Nothing
                            End If
                        End If
controlexcel2:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroXls()
                        End If
                        existeXls()
                        If excel = 1 Then
                            GoTo controlexcel2
                        End If
                        subidoxls = 1
controlpdf2:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroPdf()
                        End If
                        existePdf()
                        If pdf = 1 Then
                            GoTo controlpdf2
                        End If
                        subidopdf = 1
                        If pi.TIPO = 1 Then
controltxt2:
                            'If nombre_pc = "ROBOT" Then
                            '    subirFicheroCsv()
                            'End If
                            'existeCsv()
                            'If csv = 1 Then
                            '    GoTo controltxt2
                            'End If
                            'movertxt()
                        End If
                        modificarRegistro()
                        Dim fechaactual As Date = Now()
                        Dim _fecha As String
                        _fecha = Format(fechaactual, "yyyy-MM-dd")
                        pi.marcarsubido(_fecha)
                        If tipoinforme = 15 Then
                            enviaremail()
                        End If
                        Dim s As New dSolicitudAnalisis
                        Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                        Dim fecenv As String
                        fecenv = Format(fechaenvio, "yyyy-MM-dd")
                        s.ID = idficha
                        s.actualizarfechaenvio2(fecenv)
                        s.marcar2()
                        s = Nothing
                        If subidoxls = 1 And subidopdf = 1 Then
                            If nombre_pc = "ROBOT" Then
                                moverexcel()
                                moverpdf()
                            Else
                                moverexcel_otrapc()
                                moverpdf_otrapc()
                            End If
                            ' Grabar estado de la ficha
                            Dim est As New dEstados
                            est.FICHA = idficha
                            est.ESTADO = 8
                            est.FECHA = fecenv
                            est.guardar2()
                            est = Nothing
                            '****************************
                            envio_mail_informes()
                        End If
                    End If
                    ciefluentes = Nothing
                Next
            End If
        End If
    End Sub
    Private Sub subir_informes_atb()
        Dim pi As New dPreinformes
        Dim lista As New ArrayList
        Dim subidoxls As Integer = 0
        Dim subidopdf As Integer = 0
        Dim subidotxt As Integer = 0
        lista = pi.listarparasubiratb
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each pi In lista
                    Dim cimicro As New dControlInformesMicro
                    Dim ficha As Long = 0
                    cimicro.FICHA = pi.FICHA
                    cimicro = cimicro.buscarxficha
                    If Not cimicro Is Nothing Then
                        If cimicro.CONTROLADO = 1 Then
                            idficha = pi.FICHA
                            _tipoinforme = pi.TIPO
                            enviar_copia = pi.COPIA
                            _abonado = pi.ABONADO
                            _comentarios = pi.COMENTARIO
                            Dim sa As New dSolicitudAnalisis
                            sa.ID = idficha
                            sa = sa.buscar
                            If Not sa Is Nothing Then
                                Dim p As New dCliente
                                tipoinforme = sa.IDTIPOINFORME
                                p.ID = sa.IDPRODUCTOR
                                p = p.buscar
                                If Not p Is Nothing Then
                                    email = ""
                                    email2 = ""
                                    If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                        email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                    End If
                                    If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                        email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                    End If
                                    If email = "" Then
                                        If email2 = "" Then
                                            If p.EMAIL <> "" Then
                                                email = RTrim(p.EMAIL)
                                            End If
                                        Else
                                            email = email2
                                        End If
                                    Else
                                        If email2 = "" Then
                                            email = email
                                        Else
                                            email = email & "," & email2
                                        End If
                                    End If
                                    productorweb_com = p.USUARIO_WEB
                                    Dim pw_com As New dProductorWeb_com
                                    pw_com.USUARIO = productorweb_com
                                    pw_com = pw_com.buscar
                                    If Not pw_com Is Nothing Then
                                        idproductorweb_com = pw_com.ID
                                        'email = RTrim(pw_com.ENVIAR_EMAIL)
                                        celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                        carpeta = idproductorweb_com
                                        crea_carpeta()
                                    End If
                                    sa = Nothing
                                End If
                            End If
controlexcel:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroXls()
                            End If
                            existeXls()
                            If excel = 1 Then
                                GoTo controlexcel
                            End If
                            subidoxls = 1
controlpdf:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroPdf()
                            End If
                            existePdf()
                            If pdf = 1 Then
                                GoTo controlpdf
                            End If
                            subidopdf = 1
                            If pi.TIPO = 1 Then
controltxt:
                                'If nombre_pc = "ROBOT" Then
                                '    subirFicheroCsv()
                                'End If
                                'existeCsv()
                                'If csv = 1 Then
                                '    GoTo controltxt
                                'End If
                                'movertxt()
                            End If
                            modificarRegistro()
                            Dim fechaactual As Date = Now()
                            Dim _fecha As String
                            _fecha = Format(fechaactual, "yyyy-MM-dd")
                            pi.marcarsubido(_fecha)
                            If tipoinforme = 15 Then
                                enviaremail()
                            End If
                            Dim s As New dSolicitudAnalisis
                            Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                            Dim fecenv As String
                            fecenv = Format(fechaenvio, "yyyy-MM-dd")
                            s.ID = idficha
                            s.actualizarfechaenvio2(fecenv)
                            s.marcar2()
                            s = Nothing
                            If subidoxls = 1 And subidopdf = 1 Then
                                If nombre_pc = "ROBOT" Then
                                    moverexcel()
                                    moverpdf()
                                Else
                                    moverexcel_otrapc()
                                    moverpdf_otrapc()
                                End If
                                ' Grabar estado de la ficha
                                Dim est As New dEstados
                                est.FICHA = idficha
                                est.ESTADO = 8
                                est.FECHA = fecenv
                                est.guardar2()
                                est = Nothing
                                '****************************
                                envio_mail_informes()
                            End If
                        Else
                        End If
                    Else
                        idficha = pi.FICHA
                        _tipoinforme = pi.TIPO
                        enviar_copia = pi.COPIA
                        _abonado = pi.ABONADO
                        _comentarios = pi.COMENTARIO
                        Dim sa As New dSolicitudAnalisis
                        sa.ID = idficha
                        sa = sa.buscar
                        If Not sa Is Nothing Then
                            Dim p As New dCliente
                            tipoinforme = sa.IDTIPOINFORME
                            p.ID = sa.IDPRODUCTOR
                            p = p.buscar
                            If Not p Is Nothing Then
                                email = ""
                                email2 = ""
                                If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                    email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                End If
                                If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                    email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                End If
                                If email = "" Then
                                    If email2 = "" Then
                                        If p.EMAIL <> "" Then
                                            email = RTrim(p.EMAIL)
                                        End If
                                    Else
                                        email = email2
                                    End If
                                Else
                                    If email2 = "" Then
                                        email = email
                                    Else
                                        email = email & "," & email2
                                    End If
                                End If
                                productorweb_com = p.USUARIO_WEB
                                Dim pw_com As New dProductorWeb_com
                                pw_com.USUARIO = productorweb_com
                                pw_com = pw_com.buscar
                                If Not pw_com Is Nothing Then
                                    idproductorweb_com = pw_com.ID
                                    'email = RTrim(pw_com.ENVIAR_EMAIL)
                                    celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                    carpeta = idproductorweb_com
                                    crea_carpeta()
                                End If
                                sa = Nothing
                            End If
                        End If
controlexcel2:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroXls()
                        End If
                        existeXls()
                        If excel = 1 Then
                            GoTo controlexcel2
                        End If
                        subidoxls = 1
controlpdf2:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroPdf()
                        End If
                        existePdf()
                        If pdf = 1 Then
                            GoTo controlpdf2
                        End If
                        subidopdf = 1
                        If pi.TIPO = 1 Then
controltxt2:
                            'If nombre_pc = "ROBOT" Then
                            '    subirFicheroCsv()
                            'End If
                            'existeCsv()
                            'If csv = 1 Then
                            '    GoTo controltxt2
                            'End If
                            'movertxt()
                        End If
                        modificarRegistro()
                        Dim fechaactual As Date = Now()
                        Dim _fecha As String
                        _fecha = Format(fechaactual, "yyyy-MM-dd")
                        pi.marcarsubido(_fecha)

                        If tipoinforme = 15 Then
                            enviaremail()
                        End If
                        Dim s As New dSolicitudAnalisis
                        Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                        Dim fecenv As String
                        fecenv = Format(fechaenvio, "yyyy-MM-dd")
                        s.ID = idficha
                        s.actualizarfechaenvio2(fecenv)
                        s.marcar2()
                        s = Nothing
                        If subidoxls = 1 And subidopdf = 1 Then
                            If nombre_pc = "ROBOT" Then
                                moverexcel()
                                moverpdf()
                            Else
                                moverexcel_otrapc()
                                moverpdf_otrapc()
                            End If
                            ' Grabar estado de la ficha
                            Dim est As New dEstados
                            est.FICHA = idficha
                            est.ESTADO = 8
                            est.FECHA = fecenv
                            est.guardar2()
                            est = Nothing
                            '****************************
                            envio_mail_informes()
                        End If
                    End If
                    cimicro = Nothing
                Next
            End If
        End If
    End Sub
    Private Sub subir_informes_bacteriologia()
        Dim pi As New dPreinformes
        Dim lista As New ArrayList
        Dim subidoxls As Integer = 0
        Dim subidopdf As Integer = 0
        Dim subidotxt As Integer = 0
        lista = pi.listarparasubirbact(tipo)
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each pi In lista
                    Dim cimicro As New dControlInformesMicro
                    Dim ficha As Long = 0
                    cimicro.FICHA = pi.FICHA
                    cimicro = cimicro.buscarxficha
                    If Not cimicro Is Nothing Then
                        If cimicro.CONTROLADO = 1 Then
                            idficha = pi.FICHA
                            _tipoinforme = pi.TIPO
                            enviar_copia = pi.COPIA
                            _abonado = pi.ABONADO
                            _comentarios = pi.COMENTARIO
                            Dim sa As New dSolicitudAnalisis
                            sa.ID = idficha
                            sa = sa.buscar
                            If Not sa Is Nothing Then
                                Dim p As New dCliente
                                tipoinforme = sa.IDTIPOINFORME
                                p.ID = sa.IDPRODUCTOR
                                p = p.buscar
                                If Not p Is Nothing Then
                                    email = ""
                                    email2 = ""
                                    If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                        email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                    End If
                                    If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                        email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                    End If
                                    If email = "" Then
                                        If email2 = "" Then
                                            If p.EMAIL <> "" Then
                                                email = RTrim(p.EMAIL)
                                            End If
                                        Else
                                            email = email2
                                        End If
                                    Else
                                        If email2 = "" Then
                                            email = email
                                        Else
                                            email = email & "," & email2
                                        End If
                                    End If
                                    productorweb_com = p.USUARIO_WEB
                                    Dim pw_com As New dProductorWeb_com
                                    pw_com.USUARIO = productorweb_com
                                    pw_com = pw_com.buscar
                                    If Not pw_com Is Nothing Then
                                        idproductorweb_com = pw_com.ID
                                        'email = RTrim(pw_com.ENVIAR_EMAIL)
                                        celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                        carpeta = idproductorweb_com
                                        crea_carpeta()
                                    End If
                                    sa = Nothing
                                End If
                            End If
controlexcel:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroXls()
                            End If
                            existeXls()
                            If excel = 1 Then
                                GoTo controlexcel
                            End If
                            subidoxls = 1
controlpdf:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroPdf()
                            End If
                            existePdf()
                            If pdf = 1 Then
                                GoTo controlpdf
                            End If
                            subidopdf = 1
                            If pi.TIPO = 1 Then
controltxt:
                                'If nombre_pc = "ROBOT" Then
                                '    subirFicheroCsv()
                                'End If
                                'existeCsv()
                                'If csv = 1 Then
                                '    GoTo controltxt
                                'End If
                                'movertxt()
                            End If
                            modificarRegistro()
                            Dim fechaactual As Date = Now()
                            Dim _fecha As String
                            _fecha = Format(fechaactual, "yyyy-MM-dd")
                            pi.marcarsubido(_fecha)
                            If tipoinforme = 15 Then
                                enviaremail()
                            End If
                            Dim s As New dSolicitudAnalisis
                            Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                            Dim fecenv As String
                            fecenv = Format(fechaenvio, "yyyy-MM-dd")
                            s.ID = idficha
                            s.actualizarfechaenvio2(fecenv)
                            s.marcar2()
                            s = Nothing
                            If subidoxls = 1 And subidopdf = 1 Then
                                If nombre_pc = "ROBOT" Then
                                    moverexcel()
                                    moverpdf()
                                Else
                                    moverexcel_otrapc()
                                    moverpdf_otrapc()
                                End If
                                ' Grabar estado de la ficha
                                Dim est As New dEstados
                                est.FICHA = idficha
                                est.ESTADO = 8
                                est.FECHA = fecenv
                                est.guardar2()
                                est = Nothing
                                '****************************
                                envio_mail_informes()
                            End If
                        Else
                        End If
                    Else
                        idficha = pi.FICHA
                        _tipoinforme = pi.TIPO
                        enviar_copia = pi.COPIA
                        _abonado = pi.ABONADO
                        _comentarios = pi.COMENTARIO
                        Dim sa As New dSolicitudAnalisis
                        sa.ID = idficha
                        sa = sa.buscar
                        If Not sa Is Nothing Then
                            Dim p As New dCliente
                            tipoinforme = sa.IDTIPOINFORME
                            p.ID = sa.IDPRODUCTOR
                            p = p.buscar
                            If Not p Is Nothing Then
                                email = ""
                                email2 = ""
                                If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                    email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                End If
                                If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                    email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                End If
                                If email = "" Then
                                    If email2 = "" Then
                                        If p.EMAIL <> "" Then
                                            email = RTrim(p.EMAIL)
                                        End If
                                    Else
                                        email = email2
                                    End If
                                Else
                                    If email2 = "" Then
                                        email = email
                                    Else
                                        email = email & "," & email2
                                    End If
                                End If
                                productorweb_com = p.USUARIO_WEB
                                Dim pw_com As New dProductorWeb_com
                                pw_com.USUARIO = productorweb_com
                                pw_com = pw_com.buscar
                                If Not pw_com Is Nothing Then
                                    idproductorweb_com = pw_com.ID
                                    'email = RTrim(pw_com.ENVIAR_EMAIL)
                                    celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                    carpeta = idproductorweb_com
                                    crea_carpeta()
                                End If
                                sa = Nothing
                            End If
                        End If
controlexcel2:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroXls()
                        End If
                        existeXls()
                        If excel = 1 Then
                            GoTo controlexcel2
                        End If
                        subidoxls = 1
controlpdf2:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroPdf()
                        End If
                        existePdf()
                        If pdf = 1 Then
                            GoTo controlpdf2
                        End If
                        subidopdf = 1
                        If pi.TIPO = 1 Then
controltxt2:
                            'If nombre_pc = "ROBOT" Then
                            '    subirFicheroCsv()
                            'End If
                            'existeCsv()
                            'If csv = 1 Then
                            '    GoTo controltxt2
                            'End If
                            'movertxt()
                        End If
                        modificarRegistro()
                        Dim fechaactual As Date = Now()
                        Dim _fecha As String
                        _fecha = Format(fechaactual, "yyyy-MM-dd")
                        pi.marcarsubido(_fecha)

                        If tipoinforme = 15 Then
                            enviaremail()
                        End If
                        Dim s As New dSolicitudAnalisis
                        Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                        Dim fecenv As String
                        fecenv = Format(fechaenvio, "yyyy-MM-dd")
                        s.ID = idficha
                        s.actualizarfechaenvio2(fecenv)
                        s.marcar2()
                        s = Nothing
                        If subidoxls = 1 And subidopdf = 1 Then
                            If nombre_pc = "ROBOT" Then
                                moverexcel()
                                moverpdf()
                            Else
                                moverexcel_otrapc()
                                moverpdf_otrapc()
                            End If
                            ' Grabar estado de la ficha
                            Dim est As New dEstados
                            est.FICHA = idficha
                            est.ESTADO = 8
                            est.FECHA = fecenv
                            est.guardar2()
                            est = Nothing
                            '****************************
                            envio_mail_informes()
                        End If
                    End If
                    cimicro = Nothing
                Next
            End If
        End If
    End Sub
    Private Sub subir_informes_parasitologia()
        Dim pi As New dPreinformes
        Dim lista As New ArrayList
        Dim subidoxls As Integer = 0
        Dim subidopdf As Integer = 0
        Dim subidotxt As Integer = 0
        lista = pi.listarparasubirparasitologia
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each pi In lista
                    Dim ficha As Long = 0
                    idficha = pi.FICHA
                    _tipoinforme = pi.TIPO
                    enviar_copia = pi.COPIA
                    _abonado = pi.ABONADO
                    _comentarios = pi.COMENTARIO
                    Dim sa As New dSolicitudAnalisis
                    sa.ID = idficha
                    sa = sa.buscar
                    If Not sa Is Nothing Then
                        Dim p As New dCliente
                        tipoinforme = sa.IDTIPOINFORME
                        p.ID = sa.IDPRODUCTOR
                        p = p.buscar
                        If Not p Is Nothing Then
                            email = ""
                            email2 = ""
                            If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                email = RTrim(p.NOT_EMAIL_ANALISIS1)
                            End If
                            If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                            End If
                            If email = "" Then
                                If email2 = "" Then
                                    If p.EMAIL <> "" Then
                                        email = RTrim(p.EMAIL)
                                    End If
                                Else
                                    email = email2
                                End If
                            Else
                                If email2 = "" Then
                                    email = email
                                Else
                                    email = email & "," & email2
                                End If
                            End If
                            productorweb_com = p.USUARIO_WEB
                            Dim pw_com As New dProductorWeb_com
                            pw_com.USUARIO = productorweb_com
                            pw_com = pw_com.buscar
                            If Not pw_com Is Nothing Then
                                idproductorweb_com = pw_com.ID
                                celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                carpeta = idproductorweb_com
                                crea_carpeta()
                            End If
                            sa = Nothing
                        End If
                    End If
controlexcel2:
                    If nombre_pc = "ROBOT" Then
                        subirFicheroXls()
                    End If
                    existeXls()
                    If excel = 1 Then
                        GoTo controlexcel2
                    End If
                    subidoxls = 1
controlpdf2:
                    If nombre_pc = "ROBOT" Then
                        subirFicheroPdf()
                    End If
                    existePdf()
                    If pdf = 1 Then
                        GoTo controlpdf2
                    End If
                    subidopdf = 1
                    If pi.TIPO = 1 Then
controltxt2:
                        'If nombre_pc = "ROBOT" Then
                        '    subirFicheroCsv()
                        'End If
                        'existeCsv()
                        'If csv = 1 Then
                        '    GoTo controltxt2
                        'End If
                        'movertxt()
                    End If
                    modificarRegistro()
                    Dim fechaactual As Date = Now()
                    Dim _fecha As String
                    _fecha = Format(fechaactual, "yyyy-MM-dd")
                    pi.marcarsubido(_fecha)
                    Dim s As New dSolicitudAnalisis
                    Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                    Dim fecenv As String
                    fecenv = Format(fechaenvio, "yyyy-MM-dd")
                    s.ID = idficha
                    s.actualizarfechaenvio2(fecenv)
                    s.marcar2()
                    s = Nothing
                    If subidoxls = 1 And subidopdf = 1 Then
                        If nombre_pc = "ROBOT" Then
                            moverexcel()
                            moverpdf()
                        Else
                            moverexcel_otrapc()
                            moverpdf_otrapc()
                        End If
                        ' Grabar estado de la ficha
                        Dim est As New dEstados
                        est.FICHA = idficha
                        est.ESTADO = 8
                        est.FECHA = fecenv
                        est.guardar2()
                        est = Nothing
                        '****************************
                        envio_mail_informes()
                    End If
                Next
            End If
        End If
    End Sub
    Private Sub subir_informes_patologia()
        Dim pi As New dPreinformes
        Dim lista As New ArrayList
        Dim subidoxls As Integer = 0
        Dim subidopdf As Integer = 0
        Dim subidotxt As Integer = 0
        lista = pi.listarparasubirpatologia
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each pi In lista
                    Dim ficha As Long = 0
                    idficha = pi.FICHA
                    _tipoinforme = pi.TIPO
                    enviar_copia = pi.COPIA
                    _abonado = pi.ABONADO
                    _comentarios = pi.COMENTARIO
                    Dim sa As New dSolicitudAnalisis
                    sa.ID = idficha
                    sa = sa.buscar
                    If Not sa Is Nothing Then
                        Dim p As New dCliente
                        tipoinforme = sa.IDTIPOINFORME
                        p.ID = sa.IDPRODUCTOR
                        p = p.buscar
                        If Not p Is Nothing Then
                            email = ""
                            email2 = ""
                            If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                email = RTrim(p.NOT_EMAIL_ANALISIS1)
                            End If
                            If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                            End If
                            If email = "" Then
                                If email2 = "" Then
                                    If p.EMAIL <> "" Then
                                        email = RTrim(p.EMAIL)
                                    End If
                                Else
                                    email = email2
                                End If
                            Else
                                If email2 = "" Then
                                    email = email
                                Else
                                    email = email & "," & email2
                                End If
                            End If
                            productorweb_com = p.USUARIO_WEB
                            Dim pw_com As New dProductorWeb_com
                            pw_com.USUARIO = productorweb_com
                            pw_com = pw_com.buscar
                            If Not pw_com Is Nothing Then
                                idproductorweb_com = pw_com.ID
                                celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                carpeta = idproductorweb_com
                                crea_carpeta()
                            End If
                            sa = Nothing
                        End If
                    End If
controlexcel2:
                    If nombre_pc = "ROBOT" Then
                        subirFicheroXls()
                    End If
                    existeXls()
                    If excel = 1 Then
                        GoTo controlexcel2
                    End If
                    subidoxls = 1
controlpdf2:
                    If nombre_pc = "ROBOT" Then
                        subirFicheroPdf()
                    End If
                    existePdf()
                    If pdf = 1 Then
                        GoTo controlpdf2
                    End If
                    subidopdf = 1
                    If pi.TIPO = 1 Then
controltxt2:
                        'If nombre_pc = "ROBOT" Then
                        '    subirFicheroCsv()
                        'End If
                        'existeCsv()
                        'If csv = 1 Then
                        '    GoTo controltxt2
                        'End If
                        'movertxt()
                    End If
                    modificarRegistro()
                    Dim fechaactual As Date = Now()
                    Dim _fecha As String
                    _fecha = Format(fechaactual, "yyyy-MM-dd")
                    pi.marcarsubido(_fecha)
                    Dim s As New dSolicitudAnalisis
                    Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                    Dim fecenv As String
                    fecenv = Format(fechaenvio, "yyyy-MM-dd")
                    s.ID = idficha
                    s.actualizarfechaenvio2(fecenv)
                    s.marcar2()
                    s = Nothing
                    If subidoxls = 1 And subidopdf = 1 Then
                        If nombre_pc = "ROBOT" Then
                            moverexcel()
                            moverpdf()
                        Else
                            moverexcel_otrapc()
                            moverpdf_otrapc()
                        End If
                        ' Grabar estado de la ficha
                        Dim est As New dEstados
                        est.FICHA = idficha
                        est.ESTADO = 8
                        est.FECHA = fecenv
                        est.guardar2()
                        est = Nothing
                        '****************************
                        envio_mail_informes()
                    End If
                Next
            End If
        End If
    End Sub
    Private Sub subir_informes_toxicologia()
        Dim pi As New dPreinformes
        Dim lista As New ArrayList
        Dim subidoxls As Integer = 0
        Dim subidopdf As Integer = 0
        Dim subidotxt As Integer = 0
        lista = pi.listarparasubirtoxicologia
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each pi In lista
                    Dim ficha As Long = 0
                    idficha = pi.FICHA
                    _tipoinforme = pi.TIPO
                    enviar_copia = pi.COPIA
                    _abonado = pi.ABONADO
                    _comentarios = pi.COMENTARIO
                    Dim sa As New dSolicitudAnalisis
                    sa.ID = idficha
                    sa = sa.buscar
                    If Not sa Is Nothing Then
                        Dim p As New dCliente
                        tipoinforme = sa.IDTIPOINFORME
                        p.ID = sa.IDPRODUCTOR
                        p = p.buscar
                        If Not p Is Nothing Then
                            email = ""
                            email2 = ""
                            If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                email = RTrim(p.NOT_EMAIL_ANALISIS1)
                            End If
                            If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                            End If
                            If email = "" Then
                                If email2 = "" Then
                                    If p.EMAIL <> "" Then
                                        email = RTrim(p.EMAIL)
                                    End If
                                Else
                                    email = email2
                                End If
                            Else
                                If email2 = "" Then
                                    email = email
                                Else
                                    email = email & "," & email2
                                End If
                            End If
                            productorweb_com = p.USUARIO_WEB
                            Dim pw_com As New dProductorWeb_com
                            pw_com.USUARIO = productorweb_com
                            pw_com = pw_com.buscar
                            If Not pw_com Is Nothing Then
                                idproductorweb_com = pw_com.ID
                                celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                carpeta = idproductorweb_com
                                crea_carpeta()
                            End If
                            sa = Nothing
                        End If
                    End If
controlexcel2:
                    If nombre_pc = "ROBOT" Then
                        subirFicheroXls()
                    End If
                    existeXls()
                    If excel = 1 Then
                        GoTo controlexcel2
                    End If
                    subidoxls = 1
controlpdf2:
                    If nombre_pc = "ROBOT" Then
                        subirFicheroPdf()
                    End If
                    existePdf()
                    If pdf = 1 Then
                        GoTo controlpdf2
                    End If
                    subidopdf = 1
                    If pi.TIPO = 1 Then
controltxt2:
                        'If nombre_pc = "ROBOT" Then
                        '    subirFicheroCsv()
                        'End If
                        'existeCsv()
                        'If csv = 1 Then
                        '    GoTo controltxt2
                        'End If
                        'movertxt()
                    End If
                    modificarRegistro()
                    Dim fechaactual As Date = Now()
                    Dim _fecha As String
                    _fecha = Format(fechaactual, "yyyy-MM-dd")
                    pi.marcarsubido(_fecha)
                    Dim s As New dSolicitudAnalisis
                    Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                    Dim fecenv As String
                    fecenv = Format(fechaenvio, "yyyy-MM-dd")
                    s.ID = idficha
                    s.actualizarfechaenvio2(fecenv)
                    s.marcar2()
                    s = Nothing
                    If subidoxls = 1 And subidopdf = 1 Then
                        If nombre_pc = "ROBOT" Then
                            moverexcel()
                            moverpdf()
                        Else
                            moverexcel_otrapc()
                            moverpdf_otrapc()
                        End If
                        ' Grabar estado de la ficha
                        Dim est As New dEstados
                        est.FICHA = idficha
                        est.ESTADO = 8
                        est.FECHA = fecenv
                        est.guardar2()
                        est = Nothing
                        '****************************
                        envio_mail_informes()
                    End If
                Next
            End If
        End If
    End Sub
    Private Sub subir_informes_ambiental()
        Dim pi As New dPreinformes
        Dim lista As New ArrayList
        Dim subidoxls As Integer = 0
        Dim subidopdf As Integer = 0
        Dim subidotxt As Integer = 0
        lista = pi.listarparasubirambiental
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each pi In lista
                    Dim cimicro As New dControlInformesMicro
                    Dim ficha As Long = 0
                    cimicro.FICHA = pi.FICHA
                    cimicro = cimicro.buscarxficha
                    If Not cimicro Is Nothing Then
                        If cimicro.CONTROLADO = 1 Then
                            idficha = pi.FICHA
                            _tipoinforme = pi.TIPO
                            enviar_copia = pi.COPIA
                            _abonado = pi.ABONADO
                            _comentarios = pi.COMENTARIO
                            Dim sa As New dSolicitudAnalisis
                            sa.ID = idficha
                            sa = sa.buscar
                            If Not sa Is Nothing Then
                                Dim p As New dCliente
                                tipoinforme = sa.IDTIPOINFORME
                                p.ID = sa.IDPRODUCTOR
                                p = p.buscar
                                If Not p Is Nothing Then
                                    email = ""
                                    email2 = ""
                                    If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                        email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                    End If
                                    If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                        email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                    End If
                                    If email = "" Then
                                        If email2 = "" Then
                                            If p.EMAIL <> "" Then
                                                email = RTrim(p.EMAIL)
                                            End If
                                        Else
                                            email = email2
                                        End If
                                    Else
                                        If email2 = "" Then
                                            email = email
                                        Else
                                            email = email & "," & email2
                                        End If
                                    End If
                                    productorweb_com = p.USUARIO_WEB
                                    Dim pw_com As New dProductorWeb_com
                                    pw_com.USUARIO = productorweb_com
                                    pw_com = pw_com.buscar
                                    If Not pw_com Is Nothing Then
                                        idproductorweb_com = pw_com.ID
                                        'email = RTrim(pw_com.ENVIAR_EMAIL)
                                        celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                        carpeta = idproductorweb_com
                                        crea_carpeta()
                                    End If
                                    sa = Nothing
                                End If
                            End If
controlexcel:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroXls()
                            End If
                            existeXls()
                            If excel = 1 Then
                                GoTo controlexcel
                            End If
                            subidoxls = 1
controlpdf:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroPdf()
                            End If
                            existePdf()
                            If pdf = 1 Then
                                GoTo controlpdf
                            End If
                            subidopdf = 1
                            If pi.TIPO = 1 Then
controltxt:
                                'If nombre_pc = "ROBOT" Then
                                '    subirFicheroCsv()
                                'End If
                                'existeCsv()
                                'If csv = 1 Then
                                '    GoTo controltxt
                                'End If
                                'movertxt()
                            End If
                            modificarRegistro()
                            Dim fechaactual As Date = Now()
                            Dim _fecha As String
                            _fecha = Format(fechaactual, "yyyy-MM-dd")
                            pi.marcarsubido(_fecha)
                            If tipoinforme = 15 Then
                                enviaremail()
                            End If
                            Dim s As New dSolicitudAnalisis
                            Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                            Dim fecenv As String
                            fecenv = Format(fechaenvio, "yyyy-MM-dd")
                            s.ID = idficha
                            s.actualizarfechaenvio2(fecenv)
                            s.marcar2()
                            s = Nothing
                            If subidoxls = 1 And subidopdf = 1 Then
                                If nombre_pc = "ROBOT" Then
                                    moverexcel()
                                    moverpdf()
                                Else
                                    moverexcel_otrapc()
                                    moverpdf_otrapc()
                                End If
                                ' Grabar estado de la ficha
                                Dim est As New dEstados
                                est.FICHA = idficha
                                est.ESTADO = 8
                                est.FECHA = fecenv
                                est.guardar2()
                                est = Nothing
                                '****************************
                                envio_mail_informes()
                            End If
                        Else
                        End If
                    Else
                        idficha = pi.FICHA
                        _tipoinforme = pi.TIPO
                        enviar_copia = pi.COPIA
                        _abonado = pi.ABONADO
                        _comentarios = pi.COMENTARIO
                        Dim sa As New dSolicitudAnalisis
                        sa.ID = idficha
                        sa = sa.buscar
                        If Not sa Is Nothing Then
                            Dim p As New dCliente
                            tipoinforme = sa.IDTIPOINFORME
                            p.ID = sa.IDPRODUCTOR
                            p = p.buscar
                            If Not p Is Nothing Then
                                email = ""
                                email2 = ""
                                If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                    email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                End If
                                If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                    email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                End If
                                If email = "" Then
                                    If email2 = "" Then
                                        If p.EMAIL <> "" Then
                                            email = RTrim(p.EMAIL)
                                        End If
                                    Else
                                        email = email2
                                    End If
                                Else
                                    If email2 = "" Then
                                        email = email
                                    Else
                                        email = email & "," & email2
                                    End If
                                End If
                                productorweb_com = p.USUARIO_WEB
                                Dim pw_com As New dProductorWeb_com
                                pw_com.USUARIO = productorweb_com
                                pw_com = pw_com.buscar
                                If Not pw_com Is Nothing Then
                                    idproductorweb_com = pw_com.ID
                                    'email = RTrim(pw_com.ENVIAR_EMAIL)
                                    celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                    carpeta = idproductorweb_com
                                    crea_carpeta()
                                End If
                                sa = Nothing
                            End If
                        End If
controlexcel2:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroXls()
                        End If
                        existeXls()
                        If excel = 1 Then
                            GoTo controlexcel2
                        End If
                        subidoxls = 1
controlpdf2:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroPdf()
                        End If
                        existePdf()
                        If pdf = 1 Then
                            GoTo controlpdf2
                        End If
                        subidopdf = 1
                        If pi.TIPO = 1 Then
controltxt2:
                            'If nombre_pc = "ROBOT" Then
                            '    subirFicheroCsv()
                            'End If
                            'existeCsv()
                            'If csv = 1 Then
                            '    GoTo controltxt2
                            'End If
                            'movertxt()
                        End If
                        modificarRegistro()
                        Dim fechaactual As Date = Now()
                        Dim _fecha As String
                        _fecha = Format(fechaactual, "yyyy-MM-dd")
                        pi.marcarsubido(_fecha)
                        If tipoinforme = 15 Then
                            enviaremail()
                        End If
                        Dim s As New dSolicitudAnalisis
                        Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                        Dim fecenv As String
                        fecenv = Format(fechaenvio, "yyyy-MM-dd")
                        s.ID = idficha
                        s.actualizarfechaenvio2(fecenv)
                        s.marcar2()
                        s = Nothing
                        If subidoxls = 1 And subidopdf = 1 Then
                            If nombre_pc = "ROBOT" Then
                                moverexcel()
                                moverpdf()
                            Else
                                moverexcel_otrapc()
                                moverpdf_otrapc()
                            End If
                            ' Grabar estado de la ficha
                            Dim est As New dEstados
                            est.FICHA = idficha
                            est.ESTADO = 8
                            est.FECHA = fecenv
                            est.guardar2()
                            est = Nothing
                            '****************************
                            envio_mail_informes()
                        End If
                    End If
                    cimicro = Nothing
                Next
            End If
        End If
    End Sub
    Private Sub subir_informes_serologia()
        Dim pi As New dPreinformes
        Dim lista As New ArrayList
        Dim subidoxls As Integer = 0
        Dim subidopdf As Integer = 0
        Dim subidotxt As Integer = 0
        lista = pi.listarparasubirserologia
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each pi In lista
                    Dim cimicro As New dControlInformesMicro
                    Dim ficha As Long = 0
                    cimicro.FICHA = pi.FICHA
                    cimicro = cimicro.buscarxficha
                    If Not cimicro Is Nothing Then
                        If cimicro.CONTROLADO = 1 Then
                            idficha = pi.FICHA
                            _tipoinforme = pi.TIPO
                            enviar_copia = pi.COPIA
                            _abonado = pi.ABONADO
                            _comentarios = pi.COMENTARIO
                            Dim sa As New dSolicitudAnalisis
                            sa.ID = idficha
                            sa = sa.buscar
                            If Not sa Is Nothing Then
                                Dim p As New dCliente
                                tipoinforme = sa.IDTIPOINFORME
                                p.ID = sa.IDPRODUCTOR
                                p = p.buscar
                                If Not p Is Nothing Then
                                    email = ""
                                    email2 = ""
                                    If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                        email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                    End If
                                    If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                        email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                    End If
                                    If email = "" Then
                                        If email2 = "" Then
                                            If p.EMAIL <> "" Then
                                                email = RTrim(p.EMAIL)
                                            End If
                                        Else
                                            email = email2
                                        End If
                                    Else
                                        If email2 = "" Then
                                            email = email
                                        Else
                                            email = email & "," & email2
                                        End If
                                    End If
                                    productorweb_com = p.USUARIO_WEB
                                    Dim pw_com As New dProductorWeb_com
                                    pw_com.USUARIO = productorweb_com
                                    pw_com = pw_com.buscar
                                    If Not pw_com Is Nothing Then
                                        idproductorweb_com = pw_com.ID
                                        'email = RTrim(pw_com.ENVIAR_EMAIL)
                                        celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                        carpeta = idproductorweb_com
                                        crea_carpeta()
                                    End If
                                    sa = Nothing
                                End If
                            End If
controlexcel:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroXls()
                            End If
                            existeXls()
                            If excel = 1 Then
                                GoTo controlexcel
                            End If
                            subidoxls = 1
controlpdf:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroPdf()
                            End If
                            existePdf()
                            If pdf = 1 Then
                                GoTo controlpdf
                            End If
                            subidopdf = 1
                            If pi.TIPO = 1 Then
controltxt:
                                'If nombre_pc = "ROBOT" Then
                                '    subirFicheroCsv()
                                'End If
                                'existeCsv()
                                'If csv = 1 Then
                                '    GoTo controltxt
                                'End If
                                'movertxt()
                            End If
                            modificarRegistro()
                            Dim fechaactual As Date = Now()
                            Dim _fecha As String
                            _fecha = Format(fechaactual, "yyyy-MM-dd")
                            pi.marcarsubido(_fecha)
                            If tipoinforme = 15 Then
                                enviaremail()
                            End If
                            Dim s As New dSolicitudAnalisis
                            Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                            Dim fecenv As String
                            fecenv = Format(fechaenvio, "yyyy-MM-dd")
                            s.ID = idficha
                            s.actualizarfechaenvio2(fecenv)
                            s.marcar2()
                            s = Nothing
                            If subidoxls = 1 And subidopdf = 1 Then
                                If nombre_pc = "ROBOT" Then
                                    moverexcel()
                                    moverpdf()
                                Else
                                    moverexcel_otrapc()
                                    moverpdf_otrapc()
                                End If
                                ' Grabar estado de la ficha
                                Dim est As New dEstados
                                est.FICHA = idficha
                                est.ESTADO = 8
                                est.FECHA = fecenv
                                est.guardar2()
                                est = Nothing
                                '****************************
                                envio_mail_informes()
                            End If
                        Else
                        End If
                    Else
                        idficha = pi.FICHA
                        _tipoinforme = pi.TIPO
                        enviar_copia = pi.COPIA
                        _abonado = pi.ABONADO
                        _comentarios = pi.COMENTARIO
                        Dim sa As New dSolicitudAnalisis
                        sa.ID = idficha
                        sa = sa.buscar
                        If Not sa Is Nothing Then
                            Dim p As New dCliente
                            tipoinforme = sa.IDTIPOINFORME
                            p.ID = sa.IDPRODUCTOR
                            p = p.buscar
                            If Not p Is Nothing Then
                                email = ""
                                email2 = ""
                                If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                    email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                End If
                                If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                    email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                End If
                                If email = "" Then
                                    If email2 = "" Then
                                        If p.EMAIL <> "" Then
                                            email = RTrim(p.EMAIL)
                                        End If
                                    Else
                                        email = email2
                                    End If
                                Else
                                    If email2 = "" Then
                                        email = email
                                    Else
                                        email = email & "," & email2
                                    End If
                                End If
                                productorweb_com = p.USUARIO_WEB
                                Dim pw_com As New dProductorWeb_com
                                pw_com.USUARIO = productorweb_com
                                pw_com = pw_com.buscar
                                If Not pw_com Is Nothing Then
                                    idproductorweb_com = pw_com.ID
                                    'email = RTrim(pw_com.ENVIAR_EMAIL)
                                    celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                    carpeta = idproductorweb_com
                                    crea_carpeta()
                                End If
                                sa = Nothing
                            End If
                        End If
controlexcel2:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroXls()
                        End If
                        existeXls()
                        If excel = 1 Then
                            GoTo controlexcel2
                        End If
                        subidoxls = 1
controlpdf2:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroPdf()
                        End If
                        existePdf()
                        If pdf = 1 Then
                            GoTo controlpdf2
                        End If
                        subidopdf = 1
                        If pi.TIPO = 1 Then
controltxt2:
                            'If nombre_pc = "ROBOT" Then
                            '    subirFicheroCsv()
                            'End If
                            'existeCsv()
                            'If csv = 1 Then
                            '    GoTo controltxt2
                            'End If
                            'movertxt()
                        End If
                        modificarRegistro()
                        Dim fechaactual As Date = Now()
                        Dim _fecha As String
                        _fecha = Format(fechaactual, "yyyy-MM-dd")
                        pi.marcarsubido(_fecha)
                        If tipoinforme = 15 Then
                            enviaremail()
                        End If
                        Dim s As New dSolicitudAnalisis
                        Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                        Dim fecenv As String
                        fecenv = Format(fechaenvio, "yyyy-MM-dd")
                        s.ID = idficha
                        s.actualizarfechaenvio2(fecenv)
                        s.marcar2()
                        s = Nothing
                        If subidoxls = 1 And subidopdf = 1 Then
                            If nombre_pc = "ROBOT" Then
                                moverexcel()
                                moverpdf()
                            Else
                                moverexcel_otrapc()
                                moverpdf_otrapc()
                            End If
                            ' Grabar estado de la ficha
                            Dim est As New dEstados
                            est.FICHA = idficha
                            est.ESTADO = 8
                            est.FECHA = fecenv
                            est.guardar2()
                            est = Nothing
                            '****************************
                            envio_mail_informes()
                        End If
                    End If
                    cimicro = Nothing
                Next
            End If
        End If
    End Sub
    Private Sub subir_informes_subproductos()
        Dim pi As New dPreinformes
        Dim lista As New ArrayList
        Dim subidoxls As Integer = 0
        Dim subidopdf As Integer = 0
        Dim subidotxt As Integer = 0
        lista = pi.listarparasubirsubproductos
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each pi In lista
                    Dim cimicro As New dControlInformesMicro
                    Dim ficha As Long = 0
                    cimicro.FICHA = pi.FICHA
                    cimicro = cimicro.buscarxficha
                    If Not cimicro Is Nothing Then
                        If cimicro.CONTROLADO = 1 Then
                            idficha = pi.FICHA
                            _tipoinforme = pi.TIPO
                            enviar_copia = pi.COPIA
                            _abonado = pi.ABONADO
                            _comentarios = pi.COMENTARIO
                            Dim sa As New dSolicitudAnalisis
                            sa.ID = idficha
                            sa = sa.buscar
                            If Not sa Is Nothing Then
                                Dim p As New dCliente
                                tipoinforme = sa.IDTIPOINFORME
                                p.ID = sa.IDPRODUCTOR
                                p = p.buscar
                                If Not p Is Nothing Then
                                    email = ""
                                    email2 = ""
                                    If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                        email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                    End If
                                    If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                        email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                    End If
                                    If email = "" Then
                                        If email2 = "" Then
                                            If p.EMAIL <> "" Then
                                                email = RTrim(p.EMAIL)
                                            End If
                                        Else
                                            email = email2
                                        End If
                                    Else
                                        If email2 = "" Then
                                            email = email
                                        Else
                                            email = email & "," & email2
                                        End If
                                    End If
                                    productorweb_com = p.USUARIO_WEB
                                    Dim pw_com As New dProductorWeb_com
                                    pw_com.USUARIO = productorweb_com
                                    pw_com = pw_com.buscar
                                    If Not pw_com Is Nothing Then
                                        idproductorweb_com = pw_com.ID
                                        'email = RTrim(pw_com.ENVIAR_EMAIL)
                                        celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                        carpeta = idproductorweb_com
                                        crea_carpeta()
                                    End If
                                    sa = Nothing
                                End If
                            End If
controlexcel:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroXls()
                            End If
                            existeXls()
                            If excel = 1 Then
                                GoTo controlexcel
                            End If
                            subidoxls = 1
controlpdf:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroPdf()
                            End If
                            existePdf()
                            If pdf = 1 Then
                                GoTo controlpdf
                            End If
                            subidopdf = 1
                            If pi.TIPO = 1 Then
controltxt:
                                'If nombre_pc = "ROBOT" Then
                                '    subirFicheroCsv()
                                'End If
                                'existeCsv()
                                'If csv = 1 Then
                                '    GoTo controltxt
                                'End If
                                'movertxt()
                            End If
                            modificarRegistro()
                            Dim fechaactual As Date = Now()
                            Dim _fecha As String
                            _fecha = Format(fechaactual, "yyyy-MM-dd")
                            pi.marcarsubido(_fecha)
                            If tipoinforme = 15 Then
                                enviaremail()
                            End If
                            Dim s As New dSolicitudAnalisis
                            Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                            Dim fecenv As String
                            fecenv = Format(fechaenvio, "yyyy-MM-dd")
                            s.ID = idficha
                            s.actualizarfechaenvio2(fecenv)
                            s.marcar2()
                            s = Nothing
                            If subidoxls = 1 And subidopdf = 1 Then
                                If nombre_pc = "ROBOT" Then
                                    moverexcel()
                                    moverpdf()
                                Else
                                    moverexcel_otrapc()
                                    moverpdf_otrapc()
                                End If
                                ' Grabar estado de la ficha
                                Dim est As New dEstados
                                est.FICHA = idficha
                                est.ESTADO = 8
                                est.FECHA = fecenv
                                est.guardar2()
                                est = Nothing
                                '****************************
                                envio_mail_informes()
                            End If
                        Else
                        End If
                    Else
                        idficha = pi.FICHA
                        _tipoinforme = pi.TIPO
                        enviar_copia = pi.COPIA
                        _abonado = pi.ABONADO
                        _comentarios = pi.COMENTARIO
                        Dim sa As New dSolicitudAnalisis
                        sa.ID = idficha
                        sa = sa.buscar
                        If Not sa Is Nothing Then
                            Dim p As New dCliente
                            tipoinforme = sa.IDTIPOINFORME
                            p.ID = sa.IDPRODUCTOR
                            p = p.buscar
                            If Not p Is Nothing Then
                                email = ""
                                email2 = ""
                                If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                    email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                End If
                                If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                    email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                End If
                                If email = "" Then
                                    If email2 = "" Then
                                        If p.EMAIL <> "" Then
                                            email = RTrim(p.EMAIL)
                                        End If
                                    Else
                                        email = email2
                                    End If
                                Else
                                    If email2 = "" Then
                                        email = email
                                    Else
                                        email = email & "," & email2
                                    End If
                                End If
                                productorweb_com = p.USUARIO_WEB
                                Dim pw_com As New dProductorWeb_com
                                pw_com.USUARIO = productorweb_com
                                pw_com = pw_com.buscar
                                If Not pw_com Is Nothing Then
                                    idproductorweb_com = pw_com.ID
                                    'email = RTrim(pw_com.ENVIAR_EMAIL)
                                    celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                    carpeta = idproductorweb_com
                                    crea_carpeta()
                                End If
                                sa = Nothing
                            End If
                        End If
controlexcel2:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroXls()
                        End If
                        existeXls()
                        If excel = 1 Then
                            GoTo controlexcel2
                        End If
                        subidoxls = 1
controlpdf2:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroPdf()
                        End If
                        existePdf()
                        If pdf = 1 Then
                            GoTo controlpdf2
                        End If
                        subidopdf = 1
                        If pi.TIPO = 1 Then
controltxt2:
                            'If nombre_pc = "ROBOT" Then
                            '    subirFicheroCsv()
                            'End If
                            'existeCsv()
                            'If csv = 1 Then
                            '    GoTo controltxt2
                            'End If
                            'movertxt()
                        End If
                        modificarRegistro()
                        Dim fechaactual As Date = Now()
                        Dim _fecha As String
                        _fecha = Format(fechaactual, "yyyy-MM-dd")
                        pi.marcarsubido(_fecha)
                        If tipoinforme = 15 Then
                            enviaremail()
                        End If
                        Dim s As New dSolicitudAnalisis
                        Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                        Dim fecenv As String
                        fecenv = Format(fechaenvio, "yyyy-MM-dd")
                        s.ID = idficha
                        s.actualizarfechaenvio2(fecenv)
                        s.marcar2()
                        s = Nothing
                        If subidoxls = 1 And subidopdf = 1 Then
                            If nombre_pc = "ROBOT" Then
                                moverexcel()
                                moverpdf()
                            Else
                                moverexcel_otrapc()
                                moverpdf_otrapc()
                            End If
                            ' Grabar estado de la ficha
                            Dim est As New dEstados
                            est.FICHA = idficha
                            est.ESTADO = 8
                            est.FECHA = fecenv
                            est.guardar2()
                            est = Nothing
                            '****************************
                            envio_mail_informes()
                        End If
                    End If
                    cimicro = Nothing
                Next
            End If
        End If
    End Sub
    Private Sub subir_informes_calidad()
        Dim pi As New dPreinformes
        Dim lista As New ArrayList
        Dim subidoxls As Integer = 0
        Dim subidopdf As Integer = 0
        Dim subidotxt As Integer = 0
        lista = pi.listarparasubircalidad
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each pi In lista
                    Dim cifq As New dControlInformesFQ
                    Dim cimicro As New dControlInformesMicro
                    Dim ficha As Long = 0
                    cifq.FICHA = pi.FICHA
                    cifq = cifq.buscarxficha
                    cimicro.FICHA = pi.FICHA
                    cimicro = cimicro.buscarxficha
                    Dim control As Integer = 0
                    Dim control2 As Integer = 0
                    If Not cifq Is Nothing Then
                        If cifq.CONTROLADO = 1 Then
                            control = 1
                        Else
                            control = 0
                        End If
                    Else
                        control = 1
                    End If
                    If Not cimicro Is Nothing Then
                        If cimicro.CONTROLADO = 1 Then
                            control2 = 1
                        Else
                            control2 = 0
                        End If
                    Else
                        control2 = 1
                    End If
                    If control = 1 And control2 = 1 Then
                        idficha = pi.FICHA
                        _tipoinforme = pi.TIPO
                        enviar_copia = pi.COPIA
                        _abonado = pi.ABONADO
                        _comentarios = pi.COMENTARIO
                        Dim sa As New dSolicitudAnalisis
                        sa.ID = idficha
                        sa = sa.buscar
                        If Not sa Is Nothing Then
                            Dim p As New dCliente
                            tipoinforme = sa.IDTIPOINFORME
                            userid = sa.IDPRODUCTOR
                            p.ID = sa.IDPRODUCTOR
                            p = p.buscar
                            If Not p Is Nothing Then
                                email = ""
                                email2 = ""
                                If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                    email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                End If
                                If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                    email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                End If
                                If email = "" Then
                                    If email2 = "" Then
                                        If p.EMAIL <> "" Then
                                            email = RTrim(p.EMAIL)
                                        End If
                                    Else
                                        email = email2
                                    End If
                                Else
                                    If email2 = "" Then
                                        email = email
                                    Else
                                        email = email & "," & email2
                                    End If
                                End If
                                productorweb_com = p.USUARIO_WEB
                                Dim pw_com As New dProductorWeb_com
                                pw_com.USUARIO = productorweb_com
                                pw_com = pw_com.buscar
                                If Not pw_com Is Nothing Then
                                    idproductorweb_com = pw_com.ID
                                    'email = RTrim(pw_com.ENVIAR_EMAIL)
                                    celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                    carpeta = idproductorweb_com
                                    crea_carpeta()
                                End If
                                sa = Nothing
                            End If
                        End If
controlexcel:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroXls()
                        End If
                        existeXls()
                        If excel = 1 Then
                            GoTo controlexcel
                        End If
                        subidoxls = 1
controlpdf:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroPdf()
                        End If
                        existePdf()
                        If pdf = 1 Then
                            GoTo controlpdf
                        End If
                        subidopdf = 1
                        If pi.TIPO = 1 Then
controltxt:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroCsv()
                            End If
                            existeCsv()
                            If csv = 1 Then
                                GoTo controltxt
                            End If
                            movertxt()
                        End If
                        If userid = 6299 Then
                            If pi.TIPO = 10 Then
controltxt2:
                                If nombre_pc = "ROBOT" Then
                                    subirFicheroCsv()
                                End If
                                existeCsv()
                                If csv = 1 Then
                                    GoTo controltxt2
                                End If
                                movertxt()
                            End If
                        End If
                        
                        modificarRegistro()
                        Dim fechaactual As Date = Now()
                        Dim _fecha As String
                        _fecha = Format(fechaactual, "yyyy-MM-dd")
                        pi.marcarsubido(_fecha)
                        If tipoinforme = 15 Then
                            enviaremail()
                        End If
                        Dim s As New dSolicitudAnalisis
                        Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                        Dim fecenv As String
                        fecenv = Format(fechaenvio, "yyyy-MM-dd")
                        s.ID = idficha
                        s.actualizarfechaenvio2(fecenv)
                        s.marcar2()
                        s = Nothing
                        If subidoxls = 1 And subidopdf = 1 Then
                            If nombre_pc = "ROBOT" Then
                                moverexcel()
                                moverpdf()
                            Else
                                moverexcel_otrapc()
                                moverpdf_otrapc()
                            End If
                            ' Grabar estado de la ficha
                            Dim est As New dEstados
                            est.FICHA = idficha
                            est.ESTADO = 8
                            est.FECHA = fecenv
                            est.guardar2()
                            est = Nothing
                            '****************************
                            envio_mail_informes()
                        End If
                    End If
                    cifq = Nothing
                Next
            End If
        End If
    End Sub
    Private Sub subir_informes_nutricion()
        Dim pi As New dPreinformes
        Dim lista As New ArrayList
        Dim subidoxls As Integer = 0
        Dim subidopdf As Integer = 0
        lista = pi.listarparasubirnutricion
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each pi In lista
                    Dim cinutricion As New dControlInformesNutricion
                    Dim ficha As Long = 0
                    cinutricion.FICHA = pi.FICHA
                    cinutricion = cinutricion.buscarxficha
                    If Not cinutricion Is Nothing Then
                        If cinutricion.CONTROLADO = 1 Then
                            idficha = pi.FICHA
                            _tipoinforme = pi.TIPO
                            enviar_copia = pi.COPIA
                            _abonado = pi.ABONADO
                            _comentarios = pi.COMENTARIO
                            Dim sa As New dSolicitudAnalisis
                            sa.ID = idficha
                            sa = sa.buscar
                            If Not sa Is Nothing Then
                                Dim p As New dCliente
                                tipoinforme = sa.IDTIPOINFORME
                                p.ID = sa.IDPRODUCTOR
                                p = p.buscar
                                If Not p Is Nothing Then
                                    email = ""
                                    email2 = ""
                                    If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                        email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                    End If
                                    If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                        email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                    End If
                                    If email = "" Then
                                        If email2 = "" Then
                                            If p.EMAIL <> "" Then
                                                email = RTrim(p.EMAIL)
                                            End If
                                        Else
                                            email = email2
                                        End If
                                    Else
                                        If email2 = "" Then
                                            email = email
                                        Else
                                            email = email & "," & email2
                                        End If
                                    End If
                                    productorweb_com = p.USUARIO_WEB
                                    Dim pw_com As New dProductorWeb_com
                                    pw_com.USUARIO = productorweb_com
                                    pw_com = pw_com.buscar
                                    If Not pw_com Is Nothing Then
                                        idproductorweb_com = pw_com.ID
                                        'email = RTrim(pw_com.ENVIAR_EMAIL)
                                        celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                        carpeta = idproductorweb_com
                                        crea_carpeta()
                                    End If
                                    sa = Nothing
                                End If
                            End If
controlexcel:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroXls()
                            End If
                            existeXls()
                            If excel = 1 Then
                                GoTo controlexcel
                            End If
                            subidoxls = 1
controlpdf:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroPdf()
                            End If
                            existePdf()
                            If pdf = 1 Then
                                GoTo controlpdf
                            End If
                            subidopdf = 1
                            If pi.TIPO = 1 Then
controltxt:
                                'If nombre_pc = "ROBOT" Then
                                '    subirFicheroCsv()
                                'End If
                                'existeCsv()
                                'If csv = 1 Then
                                '    GoTo controltxt
                                'End If
                                'movertxt()
                            End If
                            modificarRegistro()
                            Dim fechaactual As Date = Now()
                            Dim _fecha As String
                            _fecha = Format(fechaactual, "yyyy-MM-dd")
                            pi.marcarsubido(_fecha)
                            If tipoinforme = 15 Then
                                enviaremail()
                            End If
                            Dim s As New dSolicitudAnalisis
                            Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                            Dim fecenv As String
                            fecenv = Format(fechaenvio, "yyyy-MM-dd")
                            s.ID = idficha
                            s.actualizarfechaenvio2(fecenv)
                            s.marcar2()
                            s = Nothing
                            If subidoxls = 1 And subidopdf = 1 Then
                                If nombre_pc = "ROBOT" Then
                                    moverexcel()
                                    moverpdf()
                                Else
                                    moverexcel_otrapc()
                                    moverpdf_otrapc()
                                End If
                                ' Grabar estado de la ficha
                                Dim est As New dEstados
                                est.FICHA = idficha
                                est.ESTADO = 8
                                est.FECHA = fecenv
                                est.guardar2()
                                est = Nothing
                                '****************************
                                envio_mail_informes()
                            End If
                        Else
                        End If
                    Else
                        idficha = pi.FICHA
                        _tipoinforme = pi.TIPO
                        enviar_copia = pi.COPIA
                        _abonado = pi.ABONADO
                        _comentarios = pi.COMENTARIO
                        Dim sa As New dSolicitudAnalisis
                        sa.ID = idficha
                        sa = sa.buscar
                        If Not sa Is Nothing Then
                            Dim p As New dCliente
                            tipoinforme = sa.IDTIPOINFORME
                            p.ID = sa.IDPRODUCTOR
                            p = p.buscar
                            If Not p Is Nothing Then
                                email = ""
                                email2 = ""
                                If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                    email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                End If
                                If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                    email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                End If
                                If email = "" Then
                                    If email2 = "" Then
                                        If p.EMAIL <> "" Then
                                            email = RTrim(p.EMAIL)
                                        End If
                                    Else
                                        email = email2
                                    End If
                                Else
                                    If email2 = "" Then
                                        email = email
                                    Else
                                        email = email & "," & email2
                                    End If
                                End If
                                productorweb_com = p.USUARIO_WEB
                                Dim pw_com As New dProductorWeb_com
                                pw_com.USUARIO = productorweb_com
                                pw_com = pw_com.buscar
                                If Not pw_com Is Nothing Then
                                    idproductorweb_com = pw_com.ID
                                    'email = RTrim(pw_com.ENVIAR_EMAIL)
                                    celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                    carpeta = idproductorweb_com
                                    crea_carpeta()
                                End If
                                sa = Nothing
                            End If
                        End If
controlexcel2:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroXls()
                        End If
                        existeXls()
                        If excel = 1 Then
                            GoTo controlexcel2
                        End If
                        subidoxls = 1
controlpdf2:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroPdf()
                        End If
                        existePdf()
                        If pdf = 1 Then
                            GoTo controlpdf2
                        End If
                        subidopdf = 1
                        If pi.TIPO = 1 Then
controltxt2:
                            'If nombre_pc = "ROBOT" Then
                            '    subirFicheroCsv()
                            'End If
                            'existeCsv()
                            'If csv = 1 Then
                            '    GoTo controltxt2
                            'End If
                            'movertxt()
                        End If
                        modificarRegistro()
                        Dim fechaactual As Date = Now()
                        Dim _fecha As String
                        _fecha = Format(fechaactual, "yyyy-MM-dd")
                        pi.marcarsubido(_fecha)
                        If tipoinforme = 15 Then
                            enviaremail()
                        End If
                        Dim s As New dSolicitudAnalisis
                        Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                        Dim fecenv As String
                        fecenv = Format(fechaenvio, "yyyy-MM-dd")
                        s.ID = idficha
                        s.actualizarfechaenvio2(fecenv)
                        s.marcar2()
                        s = Nothing
                        If subidoxls = 1 And subidopdf = 1 Then
                            If nombre_pc = "ROBOT" Then
                                moverexcel()
                                moverpdf()
                            Else
                                moverexcel_otrapc()
                                moverpdf_otrapc()
                            End If
                            ' Grabar estado de la ficha
                            Dim est As New dEstados
                            est.FICHA = idficha
                            est.ESTADO = 8
                            est.FECHA = fecenv
                            est.guardar2()
                            est = Nothing
                            '****************************
                            envio_mail_informes()
                        End If
                    End If
                    cinutricion = Nothing
                Next
            End If
        End If
    End Sub
    Private Sub subir_informes_suelos()
        Dim pi As New dPreinformes
        Dim lista As New ArrayList
        Dim subidoxls As Integer = 0
        Dim subidopdf As Integer = 0
        Dim subidotxt As Integer = 0
        lista = pi.listarparasubirsuelos
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each pi In lista
                    Dim cisuelos As New dControlInformesSuelos
                    Dim ficha As Long = 0
                    cisuelos.FICHA = pi.FICHA
                    cisuelos = cisuelos.buscarxficha
                    If Not cisuelos Is Nothing Then
                        If cisuelos.CONTROLADO = 1 Then
                            idficha = pi.FICHA
                            _tipoinforme = pi.TIPO
                            enviar_copia = pi.COPIA
                            _abonado = pi.ABONADO
                            _comentarios = pi.COMENTARIO
                            Dim sa As New dSolicitudAnalisis
                            sa.ID = idficha
                            sa = sa.buscar
                            If Not sa Is Nothing Then
                                Dim p As New dCliente
                                tipoinforme = sa.IDTIPOINFORME
                                p.ID = sa.IDPRODUCTOR
                                p = p.buscar
                                If Not p Is Nothing Then
                                    email = ""
                                    email2 = ""
                                    If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                        email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                    End If
                                    If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                        email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                    End If
                                    If email = "" Then
                                        If email2 = "" Then
                                            If p.EMAIL <> "" Then
                                                email = RTrim(p.EMAIL)
                                            End If
                                        Else
                                            email = email2
                                        End If
                                    Else
                                        If email2 = "" Then
                                            email = email
                                        Else
                                            email = email & "," & email2
                                        End If
                                    End If
                                    productorweb_com = p.USUARIO_WEB
                                    Dim pw_com As New dProductorWeb_com
                                    pw_com.USUARIO = productorweb_com
                                    pw_com = pw_com.buscar
                                    If Not pw_com Is Nothing Then
                                        idproductorweb_com = pw_com.ID
                                        'email = RTrim(pw_com.ENVIAR_EMAIL)
                                        celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                        carpeta = idproductorweb_com
                                        crea_carpeta()
                                    End If
                                    sa = Nothing
                                End If
                            End If
controlexcel:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroXls()
                            End If
                            existeXls()
                            If excel = 1 Then
                                GoTo controlexcel
                            End If
                            subidoxls = 1
controlpdf:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroPdf()
                            End If
                            existePdf()
                            If pdf = 1 Then
                                GoTo controlpdf
                            End If
                            subidopdf = 1
                            If pi.TIPO = 1 Then
controltxt:
                                'If nombre_pc = "ROBOT" Then
                                '    subirFicheroCsv()
                                'End If
                                'existeCsv()
                                'If csv = 1 Then
                                '    GoTo controltxt
                                'End If
                                'movertxt()
                            End If
                            modificarRegistro()
                            Dim fechaactual As Date = Now()
                            Dim _fecha As String
                            _fecha = Format(fechaactual, "yyyy-MM-dd")
                            pi.marcarsubido(_fecha)
                            If tipoinforme = 15 Then
                                enviaremail()
                            End If
                            Dim s As New dSolicitudAnalisis
                            Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                            Dim fecenv As String
                            fecenv = Format(fechaenvio, "yyyy-MM-dd")
                            s.ID = idficha
                            s.actualizarfechaenvio2(fecenv)
                            s.marcar2()
                            s = Nothing
                            If subidoxls = 1 And subidopdf = 1 Then
                                If nombre_pc = "ROBOT" Then
                                    moverexcel()
                                    moverpdf()
                                Else
                                    moverexcel_otrapc()
                                    moverpdf_otrapc()
                                End If
                                ' Grabar estado de la ficha
                                Dim est As New dEstados
                                est.FICHA = idficha
                                est.ESTADO = 8
                                est.FECHA = fecenv
                                est.guardar2()
                                est = Nothing
                                '****************************
                                envio_mail_informes()
                            End If
                        Else
                        End If
                    Else
                        idficha = pi.FICHA
                        _tipoinforme = pi.TIPO
                        enviar_copia = pi.COPIA
                        _abonado = pi.ABONADO
                        _comentarios = pi.COMENTARIO
                        Dim sa As New dSolicitudAnalisis
                        sa.ID = idficha
                        sa = sa.buscar
                        If Not sa Is Nothing Then
                            Dim p As New dCliente
                            tipoinforme = sa.IDTIPOINFORME
                            p.ID = sa.IDPRODUCTOR
                            p = p.buscar
                            If Not p Is Nothing Then
                                email = ""
                                email2 = ""
                                If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                    email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                End If
                                If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                    email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                End If
                                If email = "" Then
                                    If email2 = "" Then
                                        If p.EMAIL <> "" Then
                                            email = RTrim(p.EMAIL)
                                        End If
                                    Else
                                        email = email2
                                    End If
                                Else
                                    If email2 = "" Then
                                        email = email
                                    Else
                                        email = email & "," & email2
                                    End If
                                End If
                                productorweb_com = p.USUARIO_WEB
                                Dim pw_com As New dProductorWeb_com
                                pw_com.USUARIO = productorweb_com
                                pw_com = pw_com.buscar
                                If Not pw_com Is Nothing Then
                                    idproductorweb_com = pw_com.ID
                                    'email = RTrim(pw_com.ENVIAR_EMAIL)
                                    celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                    carpeta = idproductorweb_com
                                    crea_carpeta()
                                End If
                                sa = Nothing
                            End If
                        End If
controlexcel2:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroXls()
                        End If
                        existeXls()
                        If excel = 1 Then
                            GoTo controlexcel2
                        End If
                        subidoxls = 1
controlpdf2:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroPdf()
                        End If
                        existePdf()
                        If pdf = 1 Then
                            GoTo controlpdf2
                        End If
                        subidopdf = 1
                        If pi.TIPO = 1 Then
controltxt2:
                            'If nombre_pc = "ROBOT" Then
                            '    subirFicheroCsv()
                            'End If
                            'existeCsv()
                            'If csv = 1 Then
                            '    GoTo controltxt2
                            'End If
                            'movertxt()
                        End If
                        modificarRegistro()
                        Dim fechaactual As Date = Now()
                        Dim _fecha As String
                        _fecha = Format(fechaactual, "yyyy-MM-dd")
                        pi.marcarsubido(_fecha)
                        If tipoinforme = 15 Then
                            enviaremail()
                        End If
                        Dim s As New dSolicitudAnalisis
                        Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                        Dim fecenv As String
                        fecenv = Format(fechaenvio, "yyyy-MM-dd")
                        s.ID = idficha
                        s.actualizarfechaenvio2(fecenv)
                        s.marcar2()
                        s = Nothing
                        If subidoxls = 1 And subidopdf = 1 Then
                            If nombre_pc = "ROBOT" Then
                                moverexcel()
                                moverpdf()
                            Else
                                moverexcel_otrapc()
                                moverpdf_otrapc()
                            End If
                            ' Grabar estado de la ficha
                            Dim est As New dEstados
                            est.FICHA = idficha
                            est.ESTADO = 8
                            est.FECHA = fecenv
                            est.guardar2()
                            est = Nothing
                            '****************************
                            envio_mail_informes()
                        End If
                    End If
                    cisuelos = Nothing
                Next
            End If
        End If
    End Sub
    Private Sub subir_informes_foliar()
        Dim pi As New dPreinformes
        Dim lista As New ArrayList
        Dim subidoxls As Integer = 0
        Dim subidopdf As Integer = 0
        Dim subidotxt As Integer = 0
        lista = pi.listarparasubirfoliar
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each pi In lista
                    Dim cisuelos As New dControlInformesSuelos
                    Dim ficha As Long = 0
                    cisuelos.FICHA = pi.FICHA
                    cisuelos = cisuelos.buscarxficha
                    If Not cisuelos Is Nothing Then
                        If cisuelos.CONTROLADO = 1 Then
                            idficha = pi.FICHA
                            _tipoinforme = pi.TIPO
                            enviar_copia = pi.COPIA
                            _abonado = pi.ABONADO
                            _comentarios = pi.COMENTARIO
                            Dim sa As New dSolicitudAnalisis
                            sa.ID = idficha
                            sa = sa.buscar
                            If Not sa Is Nothing Then
                                Dim p As New dCliente
                                tipoinforme = sa.IDTIPOINFORME
                                p.ID = sa.IDPRODUCTOR
                                p = p.buscar
                                If Not p Is Nothing Then
                                    email = ""
                                    email2 = ""
                                    If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                        email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                    End If
                                    If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                        email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                    End If
                                    If email = "" Then
                                        If email2 = "" Then
                                            If p.EMAIL <> "" Then
                                                email = RTrim(p.EMAIL)
                                            End If
                                        Else
                                            email = email2
                                        End If
                                    Else
                                        If email2 = "" Then
                                            email = email
                                        Else
                                            email = email & "," & email2
                                        End If
                                    End If
                                    productorweb_com = p.USUARIO_WEB
                                    Dim pw_com As New dProductorWeb_com
                                    pw_com.USUARIO = productorweb_com
                                    pw_com = pw_com.buscar
                                    If Not pw_com Is Nothing Then
                                        idproductorweb_com = pw_com.ID
                                        'email = RTrim(pw_com.ENVIAR_EMAIL)
                                        celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                        carpeta = idproductorweb_com
                                        crea_carpeta()
                                    End If
                                    sa = Nothing
                                End If
                            End If
controlexcel:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroXls()
                            End If
                            existeXls()
                            If excel = 1 Then
                                GoTo controlexcel
                            End If
                            subidoxls = 1
controlpdf:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroPdf()
                            End If
                            existePdf()
                            If pdf = 1 Then
                                GoTo controlpdf
                            End If
                            subidopdf = 1
                            If pi.TIPO = 1 Then
controltxt:
                                'If nombre_pc = "ROBOT" Then
                                '    subirFicheroCsv()
                                'End If
                                'existeCsv()
                                'If csv = 1 Then
                                '    GoTo controltxt
                                'End If
                                'movertxt()
                            End If
                            modificarRegistro()
                            Dim fechaactual As Date = Now()
                            Dim _fecha As String
                            _fecha = Format(fechaactual, "yyyy-MM-dd")
                            pi.marcarsubido(_fecha)
                            If tipoinforme = 15 Then
                                enviaremail()
                            End If
                            Dim s As New dSolicitudAnalisis
                            Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                            Dim fecenv As String
                            fecenv = Format(fechaenvio, "yyyy-MM-dd")
                            s.ID = idficha
                            s.actualizarfechaenvio2(fecenv)
                            s.marcar2()
                            s = Nothing
                            If subidoxls = 1 And subidopdf = 1 Then
                                If nombre_pc = "ROBOT" Then
                                    moverexcel()
                                    moverpdf()
                                Else
                                    moverexcel_otrapc()
                                    moverpdf_otrapc()
                                End If
                                ' Grabar estado de la ficha
                                Dim est As New dEstados
                                est.FICHA = idficha
                                est.ESTADO = 8
                                est.FECHA = fecenv
                                est.guardar2()
                                est = Nothing
                                '****************************
                                envio_mail_informes()
                            End If
                        Else
                        End If
                    Else
                        idficha = pi.FICHA
                        _tipoinforme = pi.TIPO
                        enviar_copia = pi.COPIA
                        _abonado = pi.ABONADO
                        _comentarios = pi.COMENTARIO
                        Dim sa As New dSolicitudAnalisis
                        sa.ID = idficha
                        sa = sa.buscar
                        If Not sa Is Nothing Then
                            Dim p As New dCliente
                            tipoinforme = sa.IDTIPOINFORME
                            p.ID = sa.IDPRODUCTOR
                            p = p.buscar
                            If Not p Is Nothing Then
                                email = ""
                                email2 = ""
                                If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                    email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                End If
                                If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                    email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                End If
                                If email = "" Then
                                    If email2 = "" Then
                                        If p.EMAIL <> "" Then
                                            email = RTrim(p.EMAIL)
                                        End If
                                    Else
                                        email = email2
                                    End If
                                Else
                                    If email2 = "" Then
                                        email = email
                                    Else
                                        email = email & "," & email2
                                    End If
                                End If
                                productorweb_com = p.USUARIO_WEB
                                Dim pw_com As New dProductorWeb_com
                                pw_com.USUARIO = productorweb_com
                                pw_com = pw_com.buscar
                                If Not pw_com Is Nothing Then
                                    idproductorweb_com = pw_com.ID
                                    'email = RTrim(pw_com.ENVIAR_EMAIL)
                                    celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                    carpeta = idproductorweb_com
                                    crea_carpeta()
                                End If
                                sa = Nothing
                            End If
                        End If
controlexcel2:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroXls()
                        End If
                        existeXls()
                        If excel = 1 Then
                            GoTo controlexcel2
                        End If
                        subidoxls = 1
controlpdf2:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroPdf()
                        End If
                        existePdf()
                        If pdf = 1 Then
                            GoTo controlpdf2
                        End If
                        subidopdf = 1
                        If pi.TIPO = 1 Then
controltxt2:
                            'If nombre_pc = "ROBOT" Then
                            '    subirFicheroCsv()
                            'End If
                            'existeCsv()
                            'If csv = 1 Then
                            '    GoTo controltxt2
                            'End If
                            'movertxt()
                        End If
                        modificarRegistro()
                        Dim fechaactual As Date = Now()
                        Dim _fecha As String
                        _fecha = Format(fechaactual, "yyyy-MM-dd")
                        pi.marcarsubido(_fecha)
                        If tipoinforme = 15 Then
                            enviaremail()
                        End If
                        Dim s As New dSolicitudAnalisis
                        Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                        Dim fecenv As String
                        fecenv = Format(fechaenvio, "yyyy-MM-dd")
                        s.ID = idficha
                        s.actualizarfechaenvio2(fecenv)
                        s.marcar2()
                        s = Nothing
                        If subidoxls = 1 And subidopdf = 1 Then
                            If nombre_pc = "ROBOT" Then
                                moverexcel()
                                moverpdf()
                            Else
                                moverexcel_otrapc()
                                moverpdf_otrapc()
                            End If
                            ' Grabar estado de la ficha
                            Dim est As New dEstados
                            est.FICHA = idficha
                            est.ESTADO = 8
                            est.FECHA = fecenv
                            est.guardar2()
                            est = Nothing
                            '****************************
                            envio_mail_informes()
                        End If
                    End If
                    cisuelos = Nothing
                Next
            End If
        End If
    End Sub
    Private Sub subir_informes_brucelosis()
        Dim pi As New dPreinformes
        Dim lista As New ArrayList
        Dim subidoxls As Integer = 0
        Dim subidopdf As Integer = 0
        Dim subidotxt As Integer = 0
        lista = pi.listarparasubirbrucelosis
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each pi In lista
                    idficha = pi.FICHA
                    _tipoinforme = pi.TIPO
                    enviar_copia = pi.COPIA
                    _abonado = pi.ABONADO
                    _comentarios = pi.COMENTARIO
                    Dim sa As New dSolicitudAnalisis
                    sa.ID = idficha
                    sa = sa.buscar
                    If Not sa Is Nothing Then
                        Dim p As New dCliente
                        tipoinforme = sa.IDTIPOINFORME
                        p.ID = sa.IDPRODUCTOR
                        p = p.buscar
                        If Not p Is Nothing Then
                            email = ""
                            email2 = ""
                            If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                email = RTrim(p.NOT_EMAIL_ANALISIS1)
                            End If
                            If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                            End If
                            If email = "" Then
                                If email2 = "" Then
                                    If p.EMAIL <> "" Then
                                        email = RTrim(p.EMAIL)
                                    End If
                                Else
                                    email = email2
                                End If
                            Else
                                If email2 = "" Then
                                    email = email
                                Else
                                    email = email & "," & email2
                                End If
                            End If
                            productorweb_com = p.USUARIO_WEB
                            Dim pw_com As New dProductorWeb_com
                            pw_com.USUARIO = productorweb_com
                            pw_com = pw_com.buscar
                            If Not pw_com Is Nothing Then
                                idproductorweb_com = pw_com.ID
                                'email = RTrim(pw_com.ENVIAR_EMAIL)
                                celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                carpeta = idproductorweb_com
                                crea_carpeta()
                            End If
                            sa = Nothing
                        End If
                    End If
controlexcel:
                    If nombre_pc = "ROBOT" Then
                        subirFicheroXls()
                    End If
                    existeXls()
                    If excel = 1 Then
                        GoTo controlexcel
                    End If
                    subidoxls = 1
controlpdf:
                    If nombre_pc = "ROBOT" Then
                        subirFicheroPdf()
                    End If
                    existePdf()
                    If pdf = 1 Then
                        GoTo controlpdf
                    End If
                    subidopdf = 1
                    If pi.TIPO = 1 Then
controltxt:
                        'If nombre_pc = "ROBOT" Then
                        '    subirFicheroCsv()
                        'End If
                        'existeCsv()
                        'If csv = 1 Then
                        '    GoTo controltxt
                        'End If
                        'movertxt()
                    End If
                    modificarRegistro()
                    Dim fechaactual As Date = Now()
                    Dim _fecha As String
                    _fecha = Format(fechaactual, "yyyy-MM-dd")
                    pi.marcarsubido(_fecha)
                    If tipoinforme = 15 Then
                        enviaremail()
                    End If
                    Dim s As New dSolicitudAnalisis
                    Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                    Dim fecenv As String
                    fecenv = Format(fechaenvio, "yyyy-MM-dd")
                    s.ID = idficha
                    s.actualizarfechaenvio2(fecenv)
                    s.marcar2()
                    s = Nothing
                    If subidoxls = 1 And subidopdf = 1 Then
                        If nombre_pc = "ROBOT" Then
                            moverexcel()
                            moverpdf()
                        Else
                            moverexcel_otrapc()
                            moverpdf_otrapc()
                        End If
                        ' Grabar estado de la ficha
                        Dim est As New dEstados
                        est.FICHA = idficha
                        est.ESTADO = 8
                        est.FECHA = fecenv
                        est.guardar2()
                        est = Nothing
                        '****************************
                        envio_mail_informes()
                    End If
                Next
            End If
        End If
    End Sub
    Private Sub subir_informes_gestor()
        Dim subidoxls As Integer = 0
        Dim subidopdf As Integer = 0
        Dim subidotxt As Integer = 0
        Dim sa As New dSolicitudAnalisis
        sa.ID = idficha
        sa = sa.buscar
        If Not sa Is Nothing Then
            Dim p As New dCliente
            tipoinforme = sa.IDTIPOINFORME
            p.ID = sa.IDPRODUCTOR
            p = p.buscar
            If Not p Is Nothing Then
                productorweb_com = p.USUARIO_WEB
                Dim pw_com As New dProductorWeb_com
                pw_com.USUARIO = productorweb_com
                pw_com = pw_com.buscar
                If Not pw_com Is Nothing Then
                    idproductorweb_com = pw_com.ID
                    carpeta = idproductorweb_com
                    crea_carpeta()
                End If
                sa = Nothing
            End If
        End If
controlexcel:
        subirFicheroXls_G()
        existeXls_G()
        If excel = 1 Then
            GoTo controlexcel
        End If
        subidoxls = 1
controlpdf:
        subirFicheroPdf_G()
        existePdf_G()
        If pdf = 1 Then
            GoTo controlpdf
        End If
        subidopdf = 1
        If tipoinforme = 1 Then
controltxt:
            subirFicheroCsv_G()
            existeCsv_G()
            If csv = 1 Then
                GoTo controltxt
            End If
        End If
        Dim s As New dSolicitudAnalisis
        Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
        Dim fecenv As String
        fecenv = Format(fechaenvio, "yyyy-MM-dd")
        s.ID = idficha
        s.marcar2()
        s = Nothing
        If subidoxls = 1 And subidopdf = 1 Then
            ' Grabar estado de la ficha
            Dim est As New dEstados
            est.FICHA = idficha
            est.ESTADO = 8
            est.FECHA = fecenv
            est.guardar2()
            est = Nothing
            '****************************
        End If
    End Sub

    Public Function subirFicheroXls() As String
        Dim fichero As String = ""
        Dim destino As String = ""
        Dim user As String = "colaveco"
        Dim pass As String = "NUEVA**!!COL22"
        If tipoinforme = 1 Then
            crea_control_lechero_com()
            fichero = "C:\INFORMES PARA SUBIR\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/control_lechero/" & idficha & ".xls"
        ElseIf tipoinforme = 3 Then
            crea_agua_com()
            'fichero = "\\192.168.1.10\e\NET\INFORMES PARA SUBIR\" & idficha & ".xls"
            fichero = "C:\INFORMES PARA SUBIR\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/agua/" & idficha & ".xls"
        ElseIf tipoinforme = 4 Then
            crea_antibiograma_com()
            'fichero = "\\192.168.1.10\e\NET\INFORMES PARA SUBIR\" & idficha & ".xls"
            fichero = "C:\INFORMES PARA SUBIR\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/antibiograma/" & idficha & ".xls"
        ElseIf tipoinforme = 6 Then
            crea_parasitologia_com()
            'fichero = "\\192.168.1.10\e\NET\PARASITOLOGIA\" & idficha & ".xls"
            fichero = "C:\INFORMES PARA SUBIR\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/parasitologia/" & idficha & ".xls"
        ElseIf tipoinforme = 7 Then
            crea_productos_subproductos_com()
            'fichero = "\\192.168.1.10\e\NET\INFORMES PARA SUBIR\" & idficha & ".xls"
            fichero = "C:\INFORMES PARA SUBIR\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/productos_subproductos/" & idficha & ".xls"
        ElseIf tipoinforme = 8 Then
            crea_serologia_com()
            fichero = "C:\INFORMES PARA SUBIR\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/serologia/" & idficha & ".xls"
        ElseIf tipoinforme = 9 Then
            crea_patologia_com()
            fichero = "C:\INFORMES PARA SUBIR\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/patologia/" & idficha & ".xls"
        ElseIf tipoinforme = 10 Then
            crea_calidad_de_leche_com()
            fichero = "C:\INFORMES PARA SUBIR\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/calidad_de_leche/" & idficha & ".xls"
        ElseIf tipoinforme = 11 Then
            crea_ambiental_com()
            fichero = "C:\INFORMES PARA SUBIR\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/ambiental/" & idficha & ".xls"
        ElseIf tipoinforme = 13 Then
            crea_agro_nutricion_com()
            fichero = "C:\INFORMES PARA SUBIR\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/agro_nutricion/" & idficha & ".xls"
        ElseIf tipoinforme = 14 Then
            crea_agro_suelos_com()
            fichero = "C:\INFORMES PARA SUBIR\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/agro_suelos/" & idficha & ".xls"
        ElseIf tipoinforme = 15 Then
            crea_brucelosis_leche_com()
            fichero = "C:\INFORMES PARA SUBIR\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/brucelosis_leche/" & idficha & ".xls"
        ElseIf tipoinforme = 16 Then
            crea_efluentes_com()
            fichero = "C:\INFORMES PARA SUBIR\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/efluentes/" & idficha & ".xls"
        ElseIf tipoinforme = 17 Then
            crea_antibiograma_com()
            fichero = "C:\INFORMES PARA SUBIR\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/antibiograma/" & idficha & ".xls"
        ElseIf tipoinforme = 18 Then
            crea_antibiograma_com()
            fichero = "C:\INFORMES PARA SUBIR\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/antibiograma/" & idficha & ".xls"
        ElseIf tipoinforme = 19 Then
            crea_agro_suelos_com()
            fichero = "C:\INFORMES PARA SUBIR\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/agro_suelos/" & idficha & ".xls"
        ElseIf tipoinforme = 20 Then
            crea_patologia_com()
            fichero = "C:\INFORMES PARA SUBIR\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/patologia/" & idficha & ".xls"
        ElseIf tipoinforme = 99 Then
            crea_otros_servicios_com()
            fichero = "\\192.168.1.10\e\NET\OTROS\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/otros_servicios/" & idficha & ".xls"
        End If
        Dim infoFichero As New FileInfo(fichero)
        Dim uri As String
        uri = destino
        Dim peticionFTP As FtpWebRequest
        ' Creamos una peticion FTP con la dirección del fichero que vamos a subir
        peticionFTP = CType(FtpWebRequest.Create(New Uri(destino)), FtpWebRequest)
        ' Fijamos el usuario y la contraseña de la petición
        peticionFTP.Credentials = New NetworkCredential(user, pass)
        peticionFTP.KeepAlive = False
        peticionFTP.UsePassive = False
        ' Seleccionamos el comando que vamos a utilizar: Subir un fichero
        peticionFTP.Method = WebRequestMethods.Ftp.UploadFile
        ' Especificamos el tipo de transferencia de datos
        peticionFTP.UseBinary = True
        ' Informamos al servidor sobre el tamaño del fichero que vamos a subir
        peticionFTP.ContentLength = infoFichero.Length
        ' Fijamos un buffer de 150KB
        Dim longitudBuffer As Integer
        longitudBuffer = 153600
        Dim lector As Byte() = New Byte(153600) {}
        Dim num As Integer
        ' Abrimos el fichero para subirlo
        Dim fs As FileStream
        fs = infoFichero.OpenRead()
        Try
            Dim escritor As Stream
            If infoFichero.Length > 0 Then
                escritor = peticionFTP.GetRequestStream()
                ' Leemos 150 KB del fichero en cada iteración
                num = fs.Read(lector, 0, longitudBuffer)
                While (num <> 0)
                    ' Escribimos el contenido del flujo de lectura en el
                    ' flujo de escritura del comando FTP
                    escritor.Write(lector, 0, num)
                    num = fs.Read(lector, 0, longitudBuffer)
                End While
                escritor.Close()
                fs.Close()
                ' Si todo ha ido bien, se devolverá String.Empty
                Return String.Empty
            End If
        Catch ex As Exception
            ' Si se produce algún fallo, se devolverá el mensaje del error
            Return ex.Message
        End Try
    End Function

    Public Function subirFicheroPdf() As String
        Dim fichero As String = ""
        Dim destino As String = ""
        Dim user As String = "colaveco"
        Dim pass As String = "NUEVA**!!COL22"
        If tipoinforme = 1 Then
            crea_control_lechero_com()
            fichero = "C:\INFORMES PARA SUBIR\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/control_lechero/" & idficha & ".pdf"
        ElseIf tipoinforme = 3 Then
            crea_agua_com()
            fichero = "C:\INFORMES PARA SUBIR\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/agua/" & idficha & ".pdf"
        ElseIf tipoinforme = 4 Then
            crea_antibiograma_com()
            fichero = "C:\INFORMES PARA SUBIR\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/antibiograma/" & idficha & ".pdf"
        ElseIf tipoinforme = 6 Then
            crea_parasitologia_com()
            fichero = "C:\INFORMES PARA SUBIR\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/parasitologia/" & idficha & ".pdf"
        ElseIf tipoinforme = 7 Then
            crea_productos_subproductos_com()
            fichero = "C:\INFORMES PARA SUBIR\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/productos_subproductos/" & idficha & ".pdf"
        ElseIf tipoinforme = 8 Then
            crea_serologia_com()
            fichero = "C:\INFORMES PARA SUBIR\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/serologia/" & idficha & ".pdf"
        ElseIf tipoinforme = 9 Then
            crea_patologia_com()
            fichero = "C:\INFORMES PARA SUBIR\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/patologia/" & idficha & ".pdf"
        ElseIf tipoinforme = 10 Then
            crea_calidad_de_leche_com()
            fichero = "C:\INFORMES PARA SUBIR\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/calidad_de_leche/" & idficha & ".pdf"
        ElseIf tipoinforme = 11 Then
            crea_ambiental_com()
            fichero = "C:\INFORMES PARA SUBIR\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/ambiental/" & idficha & ".pdf"
        ElseIf tipoinforme = 13 Then
            crea_agro_nutricion_com()
            fichero = "C:\INFORMES PARA SUBIR\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/agro_nutricion/" & idficha & ".pdf"
        ElseIf tipoinforme = 14 Then
            crea_agro_suelos_com()
            fichero = "C:\INFORMES PARA SUBIR\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/agro_suelos/" & idficha & ".pdf"
        ElseIf tipoinforme = 15 Then
            crea_brucelosis_leche_com()
            fichero = "C:\INFORMES PARA SUBIR\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/brucelosis_leche/" & idficha & ".pdf"
        ElseIf tipoinforme = 16 Then
            crea_efluentes_com()
            fichero = "C:\INFORMES PARA SUBIR\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/efluentes/" & idficha & ".pdf"
        ElseIf tipoinforme = 17 Then
            crea_antibiograma_com()
            fichero = "C:\INFORMES PARA SUBIR\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/antibiograma/" & idficha & ".pdf"
        ElseIf tipoinforme = 18 Then
            crea_antibiograma_com()
            fichero = "C:\INFORMES PARA SUBIR\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/antibiograma/" & idficha & ".pdf"
        ElseIf tipoinforme = 19 Then
            crea_agro_suelos_com()
            fichero = "C:\INFORMES PARA SUBIR\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/agro_suelos/" & idficha & ".pdf"
        ElseIf tipoinforme = 20 Then
            crea_patologia_com()
            fichero = "C:\INFORMES PARA SUBIR\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/patologia/" & idficha & ".pdf"
        ElseIf tipoinforme = 99 Then
            crea_otros_servicios_com()
            fichero = "\\192.168.1.10\e\NET\OTROS\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/otros_servicios/" & idficha & ".pdf"
        End If
        Dim infoFichero As New FileInfo(fichero)
        Dim uri As String
        uri = destino
        Dim peticionFTP As FtpWebRequest
        ' Creamos una peticion FTP con la dirección del fichero que vamos a subir
        peticionFTP = CType(FtpWebRequest.Create(New Uri(destino)), FtpWebRequest)
        ' Fijamos el usuario y la contraseña de la petición
        peticionFTP.Credentials = New NetworkCredential(user, pass)
        peticionFTP.KeepAlive = False
        peticionFTP.UsePassive = False
        ' Seleccionamos el comando que vamos a utilizar: Subir un fichero
        peticionFTP.Method = WebRequestMethods.Ftp.UploadFile
        ' Especificamos el tipo de transferencia de datos
        peticionFTP.UseBinary = True
        ' Informamos al servidor sobre el tamaño del fichero que vamos a subir
        'peticionFTP.ContentLength = infoFichero.Length
        '**********************************************************************
        Try
            peticionFTP.ContentLength = infoFichero.Length
        Catch ex As Exception
            ' Si se produce algún fallo, se devolverá el mensaje del error
            Return ex.Message
        End Try
        '**********************************************************************
        ' Fijamos un buffer de 150KB
        Dim longitudBuffer As Integer
        longitudBuffer = 153600
        Dim lector As Byte() = New Byte(153600) {}
        Dim num As Integer
        ' Abrimos el fichero para subirlo
        Dim fs As FileStream
        fs = infoFichero.OpenRead()
        Try
            Dim escritor As Stream
            escritor = peticionFTP.GetRequestStream()
            ' Leemos 150 KB del fichero en cada iteración
            num = fs.Read(lector, 0, longitudBuffer)
            While (num <> 0)
                ' Escribimos el contenido del flujo de lectura en el
                ' flujo de escritura del comando FTP
                escritor.Write(lector, 0, num)
                num = fs.Read(lector, 0, longitudBuffer)
            End While
            escritor.Close()
            fs.Close()
            ' Si todo ha ido bien, se devolverá String.Empty
            Return String.Empty
        Catch ex As Exception
            ' Si se produce algún fallo, se devolverá el mensaje del error
            'MsgBox("No encuentro el PDF " & idficha & " en Informes para subir")
            Return ex.Message
        End Try
    End Function

    Public Function subirFicheroCsv() As String
        Dim fichero As String = ""
        Dim destino As String = ""
        Dim user As String = "colaveco"
        Dim pass As String = "NUEVA**!!COL22"
        If tipoinforme = 1 Then
            fichero = "C:\INFORMES PARA SUBIR\" & idficha & ".txt"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/control_lechero/" & idficha & ".txt"
        End If
        If tipoinforme = 10 Then
            fichero = "C:\INFORMES PARA SUBIR\" & idficha & ".txt"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/calidad_de_leche/" & idficha & ".txt"
        End If
        If fichero <> "" Then
            Dim infoFichero As New FileInfo(fichero)
            Dim uri As String
            uri = destino
            Dim peticionFTP As FtpWebRequest
            ' Creamos una peticion FTP con la dirección del fichero que vamos a subir
            peticionFTP = CType(FtpWebRequest.Create(New Uri(destino)), FtpWebRequest)
            ' Fijamos el usuario y la contraseña de la petición
            peticionFTP.Credentials = New NetworkCredential(user, pass)
            peticionFTP.KeepAlive = False
            peticionFTP.UsePassive = False
            ' Seleccionamos el comando que vamos a utilizar: Subir un fichero
            peticionFTP.Method = WebRequestMethods.Ftp.UploadFile
            ' Especificamos el tipo de transferencia de datos
            peticionFTP.UseBinary = True
            ' Informamos al servidor sobre el tamaño del fichero que vamos a subir
            peticionFTP.ContentLength = infoFichero.Length
            ' Fijamos un buffer de 150KB
            Dim longitudBuffer As Integer
            longitudBuffer = 153600
            Dim lector As Byte() = New Byte(153600) {}
            Dim num As Integer
            ' Abrimos el fichero para subirlo
            Dim fs As FileStream
            fs = infoFichero.OpenRead()
            Try
                Dim escritor As Stream
                escritor = peticionFTP.GetRequestStream()
                ' Leemos 150 KB del fichero en cada iteración
                num = fs.Read(lector, 0, longitudBuffer)
                While (num <> 0)
                    ' Escribimos el contenido del flujo de lectura en el
                    ' flujo de escritura del comando FTP
                    escritor.Write(lector, 0, num)
                    num = fs.Read(lector, 0, longitudBuffer)
                End While
                escritor.Close()
                fs.Close()
                ' Si todo ha ido bien, se devolverá String.Empty
                Return String.Empty
            Catch ex As Exception
                MsgBox("No esta creado el .TXT de la ficha + " & idficha & " + en la ruta C:\INFORMES PARA SUBIR\", MsgBoxStyle.Critical, "Atención")
                ' Si se produce algún fallo, se devolverá el mensaje del error
                Return ex.Message
            End Try
        End If
    End Function

#End Region

#Region "CREAR-FTP-COLAVECO.COM"

    Public Sub crea_brucelosis_leche_com()
        Dim carpeta As Long = idproductorweb_com
        Dim pweb_com As New dProductorWeb_com
        Dim user As String = "colaveco"
        Dim pass As String = "NUEVA**!!COL22"
        Dim peticionFTP As FtpWebRequest
        Dim dir As String = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & carpeta & "/brucelosis_leche/"
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
        Dim carpeta As Long = idproductorweb_com
        Dim pweb_com As New dProductorWeb_com
        Dim user As String = "colaveco"
        Dim pass As String = "NUEVA**!!COL22"
        Dim peticionFTP As FtpWebRequest
        Dim dir As String = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & carpeta & "/agro_suelos/"
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
        Dim carpeta As Long = idproductorweb_com
        Dim pweb_com As New dProductorWeb_com
        Dim user As String = "colaveco"
        Dim pass As String = "NUEVA**!!COL22"
        Dim peticionFTP As FtpWebRequest
        Dim dir As String = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & carpeta & "/control_lechero/"
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
        Dim carpeta As Long = idproductorweb_com
        Dim pweb_com As New dProductorWeb_com
        Dim user As String = "colaveco"
        Dim pass As String = "NUEVA**!!COL22"
        Dim peticionFTP As FtpWebRequest
        Dim dir As String = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & carpeta & "/agua/"
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
        Dim carpeta As Long = idproductorweb_com
        Dim pweb_com As New dProductorWeb_com
        Dim user As String = "colaveco"
        Dim pass As String = "NUEVA**!!COL22"
        Dim peticionFTP As FtpWebRequest
        Dim dir As String = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & carpeta & "/antibiograma/"
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
        Dim carpeta As Long = idproductorweb_com
        Dim pweb_com As New dProductorWeb_com
        Dim user As String = "colaveco"
        Dim pass As String = "NUEVA**!!COL22"
        Dim peticionFTP As FtpWebRequest
        Dim dir As String = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & carpeta & "/parasitologia/"
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
        Dim carpeta As Long = idproductorweb_com
        Dim pweb_com As New dProductorWeb_com
        Dim user As String = "colaveco"
        Dim pass As String = "NUEVA**!!COL22"
        Dim peticionFTP As FtpWebRequest
        Dim dir As String = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & carpeta & "/productos_subproductos/"
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
        Dim carpeta As Long = idproductorweb_com
        Dim pweb_com As New dProductorWeb_com
        Dim user As String = "colaveco"
        Dim pass As String = "NUEVA**!!COL22"
        Dim peticionFTP As FtpWebRequest
        Dim dir As String = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & carpeta & "/serologia/"
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
        Dim carpeta As Long = idproductorweb_com
        Dim pweb_com As New dProductorWeb_com
        Dim user As String = "colaveco"
        Dim pass As String = "NUEVA**!!COL22"
        Dim peticionFTP As FtpWebRequest
        Dim dir As String = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & carpeta & "/patologia/"
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
        Dim carpeta As Long = idproductorweb_com
        Dim pweb_com As New dProductorWeb_com
        Dim user As String = "colaveco"
        Dim pass As String = "NUEVA**!!COL22"
        Dim peticionFTP As FtpWebRequest
        Dim dir As String = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & carpeta & "/calidad_de_leche/"
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
        Dim carpeta As Long = idproductorweb_com
        Dim pweb_com As New dProductorWeb_com
        Dim user As String = "colaveco"
        Dim pass As String = "NUEVA**!!COL22"
        Dim peticionFTP As FtpWebRequest
        Dim dir As String = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & carpeta & "/ambiental/"
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
        Dim carpeta As Long = idproductorweb_com
        Dim pweb_com As New dProductorWeb_com
        Dim user As String = "colaveco"
        Dim pass As String = "NUEVA**!!COL22"
        Dim peticionFTP As FtpWebRequest
        Dim dir As String = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & carpeta & "/agro_nutricion/"
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
        Dim carpeta As Long = idproductorweb_com
        Dim pweb_com As New dProductorWeb_com
        Dim user As String = "colaveco"
        Dim pass As String = "NUEVA**!!COL22"
        Dim peticionFTP As FtpWebRequest
        Dim dir As String = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & carpeta & "/otros_servicios/"
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
    Public Sub crea_efluentes_com()
        Dim carpeta As Long = idproductorweb_com
        Dim pweb_com As New dProductorWeb_com
        Dim user As String = "colaveco"
        Dim pass As String = "NUEVA**!!COL22"
        Dim peticionFTP As FtpWebRequest
        Dim dir As String = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & carpeta & "/efluentes/"
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

#End Region

#Region "Otros"

    Private Sub creartxt2(ByVal idsol As Long)
        Dim sa As New dSolicitudAnalisis
        sa.ID = idsol
        sa = sa.buscar
        If sa.IDPRODUCTOR = 6299 Or sa.IDPRODUCTOR = 2705 Then
            Dim idficha As Long = idsol
            Dim oSW As New StreamWriter("\\192.168.1.10\e\NET\CALIDAD\" & idficha & ".txt")
            Dim oSW2 As New StreamWriter("C:\INFORMES PARA SUBIR\" & idficha & ".txt")
            Dim csm As New dCalidadSolicitudMuestra
            Dim lista As New ArrayList
            lista = csm.listarporsolicitud(idficha)
            Dim s As New dSolicitudAnalisis
            s.ID = idficha
            s = s.buscar
            Dim fecha As String = ""
            If Not s Is Nothing Then
                fecha = s.FECHAINGRESO
            End If
            If Not lista Is Nothing Then
                If lista.Count > 0 Then
                    Dim Linea As String = ""
                    Linea = ""
                    For Each csm In lista
                        Dim c As New dCalidad
                        Dim finmatricula As String = ""
                        Dim matricula As String = ""
                        Dim largocadena As String = ""
                        c.FICHA = idficha
                        c.MUESTRA = Trim(csm.MUESTRA)
                        c = c.buscarxfichaxmuestra
                        If Not c Is Nothing Then
                            largocadena = c.MUESTRA
                            If largocadena.Length > 1 Then
                                finmatricula = Mid(c.MUESTRA, Len(c.MUESTRA) - 1, 2)
                            End If
                        Else
                            largocadena = Trim(csm.MUESTRA)
                            If largocadena.Length > 1 Then
                                finmatricula = Mid(csm.MUESTRA, Len(csm.MUESTRA) - 1, 2)
                            End If
                        End If
                        'If finmatricula = "T1" Or finmatricula = "T2" Or finmatricula = "T3" Or finmatricula = "T4" Or finmatricula = "T5" Or finmatricula = "T6" Or finmatricula = "T7" Or finmatricula = "T8" Or finmatricula = "T9" Or finmatricula = "t1" Or finmatricula = "t2" Or finmatricula = "t3" Or finmatricula = "t4" Or finmatricula = "t5" Or finmatricula = "t6" Or finmatricula = "t7" Or finmatricula = "t8" Or finmatricula = "t9" Then
                        '    matricula = Mid(csm.MUESTRA, 1, Len(csm.MUESTRA) - 2)
                        'Else
                        '    If Not c Is Nothing Then
                        '        matricula = c.MUESTRA
                        '    Else
                        matricula = csm.MUESTRA
                        '    End If
                        'End If
                        If matricula <> "" Then
                            Linea = Linea & matricula & ";"
                        Else
                            Linea = Linea & "-" & ";"
                        End If
                        Dim ibc As New dIbc
                        ibc.FICHA = idficha
                        If Not c Is Nothing Then
                            ibc.MUESTRA = Trim(c.MUESTRA)
                        Else
                            ibc.MUESTRA = Trim(csm.MUESTRA)
                        End If
                        ibc = ibc.buscarxfichaxmuestra
                        If csm.RB = 1 Then
                            If Not ibc Is Nothing Then
                                If ibc.RB <> -1 Then
                                    Linea = Linea & ibc.RB & ";"
                                Else
                                    Linea = Linea & "-" & ";"
                                End If
                            Else
                                Linea = Linea & "-" & ";"
                            End If
                        Else
                            Linea = Linea & "-" & ";"
                        End If
                        ibc = Nothing
                        If csm.RC = 1 Then
                            If Not c Is Nothing Then
                                If c.RC <> -1 Then
                                    Linea = Linea & c.RC & ";"
                                Else
                                    Linea = Linea & "-" & ";"
                                End If
                            Else
                                Linea = Linea & "-" & ";"
                            End If
                        Else
                            Linea = Linea & "-" & ";"
                        End If
                        If csm.COMPOSICION = 1 Then
                            If Not c Is Nothing Then
                                If c.GRASA <> -1 Then
                                    Linea = Linea & c.GRASA & "; " & c.PROTEINA & "; " & c.LACTOSA & "; " & c.ST & "; " & "-0." & c.CRIOSCOPIA
                                Else
                                    Linea = Linea & "-" & ";" & "-" & ";" & "-" & ";" & "-" & ";" & "-"
                                End If
                            Else
                                Linea = Linea & "-" & ";" & "-" & ";" & "-" & ";" & "-" & ";" & "-"
                            End If
                        Else
                            Linea = Linea & "-" & ";" & "-" & ";" & "-" & ";" & "-" & ";" & "-"
                        End If
                        oSW.WriteLine(Linea)
                        oSW2.WriteLine(Linea)
                        Linea = ""
                    Next
                End If
            End If
            oSW.Flush()
            oSW2.Flush()
        End If
    End Sub
    Private Sub creartxt(ByVal idsol As Long)
        Dim idficha As Long = idsol
        Dim oSW As New StreamWriter("\\192.168.1.10\e\NET\CONTROL_LECHERO\" & idficha & ".txt")
        Dim c As New dControl
        Dim lista4 As New ArrayList
        lista4 = c.listarporsolicitud(idficha)
        Dim secuencial As Integer = 1
        If Not lista4 Is Nothing Then
            If lista4.Count > 0 Then
                Dim cs As New dControlSolicitud
                cs.FICHA = idficha
                cs = cs.buscar
                Dim Linea As String = ""
                For Each c In lista4
                    Linea = Linea & secuencial & Chr(9)
                    If c.MUESTRA <> "" Then
                        Linea = Linea & c.MUESTRA + Chr(9)
                    Else
                        Linea = Linea & "-" & Chr(9)
                    End If
                    If c.GRASA = -1 Or c.GRASA = 0 Then
                        Linea = Linea & "-" & Chr(9)
                    Else
                        Dim valgrasa = Val(c.GRASA)
                        Linea = Linea & valgrasa & Chr(9)
                    End If
                    If c.PROTEINA = -1 Or c.PROTEINA = 0 Then
                        Linea = Linea & "-" & Chr(9)
                    Else
                        Dim valproteina = Val(c.PROTEINA)
                        Linea = Linea & valproteina & Chr(9)
                    End If
                    If c.LACTOSA = -1 Or c.LACTOSA = 0 Then
                        Linea = Linea & "-" & Chr(9)
                    Else
                        Linea = Linea & c.LACTOSA & Chr(9)
                    End If
                    Linea = Linea & "0" & Chr(9)
                    If c.RC = -1 Then
                        Linea = Linea & "-" '& vbNewLine
                    Else
                        If c.GRASA = -1 Or c.GRASA = 0 Then
                            Linea = Linea & "-" & Chr(9)
                        Else
                            If c.RC < 4 Then
                                Linea = Linea & "4" ' & vbNewLine
                            Else
                                Linea = Linea & c.RC ' & vbNewLine
                            End If
                        End If
                    End If
                    oSW.WriteLine(Linea)
                    Linea = ""
                    secuencial = secuencial + 1
                Next
            End If
        End If
        Dim sa2 As New dSolicitudAnalisis
        sa2.ID = idficha
        sa2.NMUESTRAS = secuencial - 1
        oSW.Flush()
    End Sub
    Private Sub enviocopia()
        Dim _Message As New System.Net.Mail.MailMessage()
        Dim _SMTP As New System.Net.Mail.SmtpClient
        Dim enviarcopia As String = ""
        Dim fichero As String = ""
        Dim tipo As String = ""
        enviarcopia = enviar_copia
        If tipoinforme = 1 Then
            fichero = "\\192.168.1.10\e\NET\CONTROL_LECHERO\" & idficha & ".xls"
            tipo = "Control lechero"
        ElseIf tipoinforme = 3 Then
            fichero = "\\192.168.1.10\e\NET\AGUA\" & idficha & ".xls"
            tipo = "Agua"
        ElseIf tipoinforme = 4 Then
            fichero = "\\192.168.1.10\e\NET\ANTIBIOGRAMA\" & idficha & ".xls"
            tipo = "Antibiograma"
        ElseIf tipoinforme = 5 Then
            fichero = "\\192.168.1.10\e\NET\PAL\" & idficha & ".xls"
            tipo = "PAL"
        ElseIf tipoinforme = 6 Then
            fichero = "\\192.168.1.10\e\NET\PARASITOLOGIA\" & idficha & ".xls"
            tipo = "Parasitología"
        ElseIf tipoinforme = 7 Then
            fichero = "\\192.168.1.10\e\NET\ALIMENTOS\" & idficha & ".xls"
            tipo = "Alimentos"
        ElseIf tipoinforme = 8 Then
            fichero = "\\192.168.1.10\e\NET\SEROLOGIA\" & idficha & ".xls"
            tipo = "Serología"
        ElseIf tipoinforme = 9 Then
            fichero = "\\192.168.1.10\e\NET\PATOLOGIA\" & idficha & ".xls"
            tipo = "Patología"
        ElseIf tipoinforme = 10 Then
            fichero = "\\192.168.1.10\e\NET\CALIDAD\" & idficha & ".xls"
            tipo = "Calidad de leche"
        ElseIf tipoinforme = 11 Then
            fichero = "\\192.168.1.10\e\NET\AMBIENTAL\" & idficha & ".xls"
            tipo = "Prueba ambiental"
        ElseIf tipoinforme = 12 Then
            fichero = "\\192.168.1.10\e\NET\LACTOMETROS\" & idficha & ".xls"
            tipo = "Lactómetros"
        ElseIf tipoinforme = 13 Then
            fichero = "\\192.168.1.10\e\NET\NUTRICION\" & idficha & ".xls"
            tipo = "Nutrición"
        ElseIf tipoinforme = 14 Then
            fichero = "\\192.168.1.10\e\NET\Suelos\" & idficha & ".xls"
            tipo = "Suelos"
        ElseIf tipoinforme = 15 Then
            fichero = "\\192.168.1.10\e\NET\Brucelosis en leche\" & idficha & ".xls"
            tipo = "Brucelosis en leche"
        ElseIf tipoinforme = 99 Then
            fichero = "\\192.168.1.10\e\NET\OTROS\" & idficha & ".xls"
            tipo = "Otros"
        End If
        If enviarcopia <> "" Then
            'CONFIGURACIÓN DEL STMP 
            _SMTP.Credentials = New System.Net.NetworkCredential("notificaciones@colaveco.com.uy", "19912021Notificaciones")
            _SMTP.Host = "170.249.199.66"
            _SMTP.Port = 25
            _SMTP.EnableSsl = False


            ' CONFIGURACION DEL MENSAJE 
            _Message.[To].Add(enviarcopia)
            _Message.[To].Add("envios@colaveco.com.uy")
            'Cuenta de Correo al que se le quiere enviar el e-mail 
            _Message.From = New System.Net.Mail.MailAddress("notificaciones@colaveco.com.uy", "COLAVECO", System.Text.Encoding.UTF8)
            'Quien lo envía 
            _Message.Subject = "Informe" & " " & idficha & " - " & tipo
            'Sujeto del e-mail 
            _Message.SubjectEncoding = System.Text.Encoding.UTF8
            'Codificacion 
            _Message.Body = "Adjunto informe:" & " " & tipo
            'contenido del mail 
            _Message.BodyEncoding = System.Text.Encoding.UTF8 '
            _Message.Priority = System.Net.Mail.MailPriority.Normal
            _Message.IsBodyHtml = False
            ' ADICION DE DATOS ADJUNTOS ‘
            Dim _File As String = fichero 'My.Application.Info.DirectoryPath & fichero 'archivo que se quiere adjuntar ‘
            Dim _Attachment As New System.Net.Mail.Attachment(_File, System.Net.Mime.MediaTypeNames.Application.Octet) '
            _Message.Attachments.Add(_Attachment) 'ENVIO 
            Try
                '_SMTP.Send(_Message)
                'MessageBox.Show("Correo enviado!", "Correo", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As System.Net.Mail.SmtpException ' MessageBox.Show(ex.ToString) 
            End Try
            _SMTP.Send(_Message)
            'MessageBox.Show("Pedidos enviados!", "Correo", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub
    Public Function existeXls() As Boolean
        Dim fichero As String = ""
        Dim destino As String = ""
        Dim user As String = "colaveco"
        Dim pass As String = "NUEVA**!!COL22"
        If tipoinforme = 1 Then
            fichero = "\\192.168.1.10\e\NET\CONTROL_LECHERO\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/control_lechero/" & idficha & ".xls"
        ElseIf tipoinforme = 3 Then
            fichero = "\\192.168.1.10\e\NET\AGUA\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/agua/" & idficha & ".xls"
        ElseIf tipoinforme = 4 Then
            fichero = "\\192.168.1.10\e\NET\ANTIBIOGRAMA\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/antibiograma/" & idficha & ".xls"
        ElseIf tipoinforme = 5 Then
            fichero = "\\192.168.1.10\e\NET\PAL\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy.uy/www/gestor/data_file/" & idproductorweb_com & "/pal/" & idficha & ".xls"
        ElseIf tipoinforme = 6 Then
            fichero = "\\192.168.1.10\e\NET\PARASITOLOGIA\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/parasitologia/" & idficha & ".xls"
        ElseIf tipoinforme = 7 Then
            fichero = "\\192.168.1.10\e\NET\ALIMENTOS\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/productos_subproductos/" & idficha & ".xls"
        ElseIf tipoinforme = 8 Then
            fichero = "\\192.168.1.10\e\NET\SEROLOGIA\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/serologia/" & idficha & ".xls"
        ElseIf tipoinforme = 9 Then
            fichero = "\\192.168.1.10\e\NET\PATOLOGIA\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/patologia/" & idficha & ".xls"
        ElseIf tipoinforme = 10 Then
            fichero = "\\192.168.1.10\e\NET\CALIDAD\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/calidad_de_leche/" & idficha & ".xls"
        ElseIf tipoinforme = 11 Then
            fichero = "\\192.168.1.10\e\NET\AMBIENTAL\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/ambiental/" & idficha & ".xls"
        ElseIf tipoinforme = 12 Then
            fichero = "\\192.168.1.10\e\NET\LACTOMETROS\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/lactometros_chequeos_maquina/" & idficha & ".xls"
        ElseIf tipoinforme = 13 Then
            fichero = "\\192.168.1.10\e\NET\NUTRICION\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/agro_nutricion/" & idficha & ".xls"
        ElseIf tipoinforme = 14 Then
            fichero = "\\192.168.1.10\e\NET\Suelos\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/agro_suelos/" & idficha & ".xls"
        ElseIf tipoinforme = 15 Then
            fichero = "\\192.168.1.10\e\NET\Brucelosis en leche\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/brucelosis_leche/" & idficha & ".xls"
        ElseIf tipoinforme = 16 Then
            fichero = "\\192.168.1.10\e\NET\Efluentes\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/efluentes/" & idficha & ".xls"
        ElseIf tipoinforme = 17 Then
            fichero = "\\192.168.1.10\e\NET\ANTIBIOGRAMA\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/antibiograma/" & idficha & ".xls"
        ElseIf tipoinforme = 18 Then
            fichero = "\\192.168.1.10\e\NET\ANTIBIOGRAMA\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/antibiograma/" & idficha & ".xls"
        ElseIf tipoinforme = 19 Then
            fichero = "\\192.168.1.10\e\NET\Suelos\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/agro_suelos/" & idficha & ".xls"
        ElseIf tipoinforme = 20 Then
            fichero = "\\192.168.1.10\e\NET\Toxicologia\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/patologia/" & idficha & ".xls"

        ElseIf tipoinforme = 99 Then
            fichero = "\\192.168.1.10\e\NET\OTROS\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/otros_servicios/" & idficha & ".xls"
        End If
        Dim peticionFTP As FtpWebRequest
        ' Creamos una peticion FTP con la dirección del objeto que queremos saber si existe
        peticionFTP = CType(WebRequest.Create(New Uri(destino)), FtpWebRequest)
        ' Fijamos el usuario y la contraseña de la petición
        peticionFTP.Credentials = New NetworkCredential(user, pass)
        ' Para saber si el objeto existe, solicitamos la fecha de creación del mismo
        peticionFTP.Method = WebRequestMethods.Ftp.GetDateTimestamp
        peticionFTP.UsePassive = False
        Try
            ' Si el objeto existe, se devolverá True
            Dim respuestaFTP As FtpWebResponse
            respuestaFTP = CType(peticionFTP.GetResponse(), FtpWebResponse)
            excel = 0
            Return True
        Catch ex As Exception
            mensaje = mensaje & " excel(com) - "
            excel = 1
            MsgBox("No esta creado el .XLS en \\192.168.1.10\e\NET\... de la ficha + " & idficha & "", MsgBoxStyle.Critical, "Atención")
            ' Si el objeto no existe, se producirá un error y al entrar por el Catch
            ' se devolverá falso
            Return False
        End Try
    End Function
    Public Function existePdf() As Boolean
        Dim fichero As String = ""
        Dim destino As String = ""
        Dim user As String = "colaveco"
        Dim pass As String = "NUEVA**!!COL22"
        If tipoinforme = 1 Then
            fichero = "\\192.168.1.10\e\NET\CONTROL_LECHERO\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/control_lechero/" & idficha & ".pdf"
        ElseIf tipoinforme = 3 Then
            fichero = "\\192.168.1.10\e\NET\AGUA\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/agua/" & idficha & ".pdf"
        ElseIf tipoinforme = 4 Then
            fichero = "\\192.168.1.10\e\NET\ANTIBIOGRAMA\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/antibiograma/" & idficha & ".pdf"
        ElseIf tipoinforme = 5 Then
            fichero = "\\192.168.1.10\e\NET\PAL\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/pal/" & idficha & ".pdf"
        ElseIf tipoinforme = 6 Then
            fichero = "\\192.168.1.10\e\NET\PARASITOLOGIA\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/parasitologia/" & idficha & ".pdf"
        ElseIf tipoinforme = 7 Then
            fichero = "\\192.168.1.10\e\NET\ALIMENTOS\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/productos_subproductos/" & idficha & ".pdf"
        ElseIf tipoinforme = 8 Then
            fichero = "\\192.168.1.10\e\NET\SEROLOGIA\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/serologia/" & idficha & ".pdf"
        ElseIf tipoinforme = 9 Then
            fichero = "\\192.168.1.10\e\NET\PATOLOGIA\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/patologia/" & idficha & ".pdf"
        ElseIf tipoinforme = 10 Then
            fichero = "\\192.168.1.10\e\NET\CALIDAD\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/calidad_de_leche/" & idficha & ".pdf"
        ElseIf tipoinforme = 11 Then
            fichero = "\\192.168.1.10\e\NET\AMBIENTAL\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/ambiental/" & idficha & ".pdf"
        ElseIf tipoinforme = 12 Then
            fichero = "\\192.168.1.10\e\NET\LACTOMETROS\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/lactometros_chequeos_maquina/" & idficha & ".pdf"
        ElseIf tipoinforme = 13 Then
            fichero = "\\192.168.1.10\e\NET\NUTRICION\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/agro_nutricion/" & idficha & ".pdf"
        ElseIf tipoinforme = 14 Then
            fichero = "\\192.168.1.10\e\NET\Suelos\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/agro_suelos/" & idficha & ".pdf"
        ElseIf tipoinforme = 15 Then
            fichero = "\\192.168.1.10\e\NET\Brucelosis en leche\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/brucelosis_leche/" & idficha & ".pdf"
        ElseIf tipoinforme = 16 Then
            fichero = "\\192.168.1.10\e\NET\Efluentes\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/efluentes/" & idficha & ".pdf"
        ElseIf tipoinforme = 17 Then
            fichero = "\\192.168.1.10\e\NET\ANTIBIOGRAMA\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/antibiograma/" & idficha & ".pdf"
        ElseIf tipoinforme = 18 Then
            fichero = "\\192.168.1.10\e\NET\ANTIBIOGRAMA\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/antibiograma/" & idficha & ".pdf"
        ElseIf tipoinforme = 19 Then
            fichero = "\\192.168.1.10\e\NET\Suelos\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/agro_suelos/" & idficha & ".pdf"
        ElseIf tipoinforme = 20 Then
            fichero = "\\192.168.1.10\e\NET\Toxicologia\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/patologia/" & idficha & ".pdf"
        ElseIf tipoinforme = 99 Then
            fichero = "\\192.168.1.10\e\NET\OTROS\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/otros_servicios/" & idficha & ".pdf"
        End If
        Dim peticionFTP As FtpWebRequest
        ' Creamos una peticion FTP con la dirección del objeto que queremos saber si existe
        peticionFTP = CType(WebRequest.Create(New Uri(destino)), FtpWebRequest)
        ' Fijamos el usuario y la contraseña de la petición
        peticionFTP.Credentials = New NetworkCredential(user, pass)
        ' Para saber si el objeto existe, solicitamos la fecha de creación del mismo
        peticionFTP.Method = WebRequestMethods.Ftp.GetDateTimestamp
        peticionFTP.UsePassive = False
        Try
            ' Si el objeto existe, se devolverá True
            Dim respuestaFTP As FtpWebResponse
            respuestaFTP = CType(peticionFTP.GetResponse(), FtpWebResponse)
            pdf = 0
            Return True
        Catch ex As Exception
            mensaje = mensaje & " pdf(com) - "
            pdf = 1
            MsgBox("No esta creado el PDF de la ficha + " & idficha & "", MsgBoxStyle.Critical, "Atención")
            ' Si el objeto no existe, se producirá un error y al entrar por el Catch
            ' se devolverá falso
            Return False
        End Try
    End Function
    Public Function existeCsv() As Boolean
        Dim fichero As String = ""
        Dim destino As String = ""
        Dim user As String = "colaveco"
        Dim pass As String = "NUEVA**!!COL22"
        If tipoinforme = 1 Then
            fichero = "\\192.168.1.10\e\NET\CONTROL_LECHERO\" & idficha & ".txt"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/control_lechero/" & idficha & ".txt"



            Dim peticionFTP As FtpWebRequest
            ' Creamos una peticion FTP con la dirección del objeto que queremos saber si existe
            peticionFTP = CType(WebRequest.Create(New Uri(destino)), FtpWebRequest)
            ' Fijamos el usuario y la contraseña de la petición
            peticionFTP.Credentials = New NetworkCredential(user, pass)
            ' Para saber si el objeto existe, solicitamos la fecha de creación del mismo
            peticionFTP.Method = WebRequestMethods.Ftp.GetDateTimestamp
            peticionFTP.UsePassive = False
            Try
                ' Si el objeto existe, se devolverá True
                Dim respuestaFTP As FtpWebResponse
                respuestaFTP = CType(peticionFTP.GetResponse(), FtpWebResponse)
                csv = 0
                Return True
            Catch ex As Exception
                mensaje = mensaje & " csv(com) - "
                csv = 1
                MsgBox("Si aun no esta la ficha en el Z:\, moverla a su carpeta corerspondiente, Ficha: " & idficha & ".txt", MsgBoxStyle.Critical, "Atención")
                ' Si el objeto no existe, se producirá un error y al entrar por el Catch
                ' se devolverá falso
                Return False
            End Try
        End If
        If tipoinforme = 10 Then
            fichero = "\\192.168.1.10\e\NET\CALIDAD\" & idficha & ".txt"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/calidad_de_leche/" & idficha & ".txt"



            Dim peticionFTP As FtpWebRequest
            ' Creamos una peticion FTP con la dirección del objeto que queremos saber si existe
            peticionFTP = CType(WebRequest.Create(New Uri(destino)), FtpWebRequest)
            ' Fijamos el usuario y la contraseña de la petición
            peticionFTP.Credentials = New NetworkCredential(user, pass)
            ' Para saber si el objeto existe, solicitamos la fecha de creación del mismo
            peticionFTP.Method = WebRequestMethods.Ftp.GetDateTimestamp
            peticionFTP.UsePassive = False
            Try
                ' Si el objeto existe, se devolverá True
                Dim respuestaFTP As FtpWebResponse
                respuestaFTP = CType(peticionFTP.GetResponse(), FtpWebResponse)
                csv = 0
                Return True
            Catch ex As Exception
                mensaje = mensaje & " csv(com) - "
                csv = 1
                MsgBox("Si aun no esta la ficha en el Z:\, moverla a su carpeta corerspondiente, Ficha: " & idficha & "", MsgBoxStyle.Critical, "Atención")

                ' Si el objeto no existe, se producirá un error y al entrar por el Catch
                ' se devolverá falso
                Return False
            End Try
        End If
    End Function
    Private Sub actualizar_cajas()
        Dim c As New dCajas
        Dim fecha As Date
        Dim fec As String = ""
        Dim lista_cajas As New ArrayList
        lista_cajas = c.listar
        For Each c In lista_cajas
            Dim ec As New dEnvioCajas
            ec.IDCAJA = c.CODIGO
            ec = ec.buscarultimoenvioxcaja
            If Not ec Is Nothing Then
                Dim c2 As New dCajas
                c2.CODIGO = ec.IDCAJA
                If ec.RECIBIDO = 1 Then
                    c2.ESTADO = 1
                    c2.IDCLIENTE = -1
                    fecha = ec.FECHARECIBO
                    fec = Format(fecha, "yyyy-MM-dd")
                    c2.FECHA = fec
                Else
                    If ec.IDCAJA = "Cons-Devolución" Or ec.IDCAJA = "Caja-Devolución" Or ec.IDCAJA = "Bolsa" Or ec.IDCAJA = "Gradilla" Then
                        c2.ESTADO = 1
                        c2.IDCLIENTE = -1
                        fecha = ec.FECHAENVIO
                        fec = Format(fecha, "yyyy-MM-dd")
                        c2.FECHA = fec
                    Else
                        c2.ESTADO = 2
                        c2.IDCLIENTE = ec.IDPRODUCTOR
                        fecha = ec.FECHAENVIO
                        fec = Format(fecha, "yyyy-MM-dd")
                        c2.FECHA = fec
                    End If

                End If
                c2.modificar2()
                c2 = Nothing
            End If
            ec = Nothing
        Next
    End Sub
    Public Sub crea_carpeta()
        Dim pweb_com As New dProductorWeb_com
        Dim user As String = "colaveco"
        Dim pass As String = "NUEVA**!!COL22"
        Dim peticionFTP As FtpWebRequest
        Dim dir As String = "ftp://170.249.199.66/public_html/gestor/data_file/" & carpeta
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
    Public Sub modificarRegistro()
        Dim idnet As Long = 0
        Dim sa_ As New dSolicitudAnalisis
        sa_.ID = idficha
        sa_ = sa_.buscar
        If Not sa_ Is Nothing Then
            idnet = sa_.IDPRODUCTOR
        End If
        If tipoinforme = 1 Then 'SI EL TIPO DE INFORME ES DE CONTROL LECHERO
            Dim cw_com As New dControlLecheroWeb_com
            Dim pw_com As New dProductorWeb_com
            pw_com.USUARIO = productorweb_com
            pw_com = pw_com.buscar
            Dim idproductorweb_com As Long = pw_com.ID
            Dim comentarios As String = ""
            If _comentarios.Length > 0 Then
                comentarios = _comentarios
            End If
            Dim abonado As Integer = _abonado
            Dim fecha_emision As Date = DateFecha.Value.ToString("yyyy-MM-dd")
            Dim path_excel As String = ""
            path_excel = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/control_lechero/" & idficha & ".xls"
            Dim path_pdf As String = ""
            path_pdf = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/control_lechero/" & idficha & ".pdf"
            Dim path_csv As String = ""
            path_csv = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/control_lechero/" & idficha & ".txt"
            Dim id_estado As Integer = 0
            If _abonado = 0 Then
                id_estado = 6
            ElseIf _abonado = 1 Then
                id_estado = 5
            Else
                id_estado = 5
            End If
            cw_com.FICHA = idficha
            cw_com = cw_com.buscar
            If Not cw_com Is Nothing Then
                If comentarios <> "" Then
                    cw_com.COMENTARIO = comentarios
                End If
                cw_com.ABONADO = abonado
                Dim fechaemi As String
                fechaemi = Format(fecha_emision, "yyyy-MM-dd")
                cw_com.FECHA_EMISION = fechaemi
                cw_com.PATH_EXCEL = path_excel
                cw_com.PATH_PDF = path_pdf
                cw_com.PATH_CSV = path_csv
                cw_com.ID_ESTADO = id_estado
                If (cw_com.modificar2()) Then
                    'MsgBox("Solicitud guardada", MsgBoxStyle.Information, "Atención")
                    'limpiar()
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            Else
                Dim cweb_com As New dControlLecheroWeb_com
                cweb_com.ID_USUARIO = idproductorweb_com

                If comentarios <> "" Then
                    cweb_com.COMENTARIO = comentarios
                End If
                cweb_com.ABONADO = abonado
                Dim fechaemi As String
                fechaemi = Format(fecha_emision, "yyyy-MM-dd")
                cweb_com.FECHA_CREADO = fechaemi
                cweb_com.FECHA_EMISION = fechaemi
                cweb_com.PATH_EXCEL = path_excel
                cweb_com.PATH_PDF = path_pdf
                cweb_com.PATH_CSV = path_csv
                cweb_com.FICHA = idficha
                cweb_com.ID_ESTADO = id_estado
                cweb_com.ID_LIBRO = idficha
                If (cweb_com.guardar()) Then
                    'MsgBox("Solicitud guardada", MsgBoxStyle.Information, "Atención")
                    'limpiar()
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            End If
        ElseIf tipoinforme = 3 Then 'SI EL TIPO DE INFORME ES DE AGUA
            Dim aw_com As New dAguaWeb_com
            Dim pw_com As New dProductorWeb_com
            pw_com.USUARIO = productorweb_com
            pw_com = pw_com.buscar
            Dim idproductorweb_com As Long = pw_com.ID
            Dim comentarios As String = ""
            If _comentarios.Length > 0 Then
                comentarios = _comentarios
            End If
            Dim abonado As Integer = _abonado
            Dim fecha_emision As Date = DateFecha.Value.ToString("yyyy-MM-dd")
            Dim path_excel As String = ""
            path_excel = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/agua/" & idficha & ".xls"
            Dim path_pdf As String = ""
            path_pdf = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/agua/" & idficha & ".pdf"
            Dim path_csv As String = ""
            path_csv = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/agua/" & idficha & ".txt"
            Dim id_estado As Integer = 3
            aw_com.FICHA = idficha
            aw_com = aw_com.buscar
            If Not aw_com Is Nothing Then
                If comentarios <> "" Then
                    aw_com.COMENTARIO = comentarios
                End If
                aw_com.ABONADO = abonado
                Dim fechaemi As String
                fechaemi = Format(fecha_emision, "yyyy-MM-dd")
                aw_com.FECHA_EMISION = fechaemi
                aw_com.PATH_EXCEL = path_excel
                aw_com.PATH_PDF = path_pdf
                aw_com.PATH_CSV = path_csv
                aw_com.ID_ESTADO = id_estado
                If (aw_com.modificar2()) Then
                    'MsgBox("Solicitud guardada", MsgBoxStyle.Information, "Atención")
                    'limpiar()
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            Else
                Dim aweb_com As New dAguaWeb_com
                aweb_com.ID_USUARIO = idproductorweb_com
                If comentarios <> "" Then
                    aweb_com.COMENTARIO = comentarios
                End If
                aweb_com.ABONADO = abonado
                Dim fechaemi As String
                fechaemi = Format(fecha_emision, "yyyy-MM-dd")
                aweb_com.FECHA_CREADO = fechaemi
                aweb_com.FECHA_EMISION = fechaemi
                aweb_com.PATH_EXCEL = path_excel
                aweb_com.PATH_PDF = path_pdf
                aweb_com.PATH_CSV = path_csv
                aweb_com.FICHA = idficha
                aweb_com.ID_ESTADO = id_estado
                aweb_com.ID_LIBRO = idficha
                If (aweb_com.guardar()) Then
                    'MsgBox("Solicitud guardada", MsgBoxStyle.Information, "Atención")
                    'limpiar()
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            End If
        ElseIf tipoinforme = 4 Then 'SI EL TIPO DE INFORME ES DE BACTERIOLOGÍA Y ANTIBIOGRAMA
            Dim aw_com As New dAntibiogramaWeb_com
            Dim pw_com As New dProductorWeb_com
            pw_com.USUARIO = productorweb_com
            pw_com = pw_com.buscar
            Dim idproductorweb_com As Long = pw_com.ID
            Dim comentarios As String = ""
            If _comentarios.Length > 0 Then
                comentarios = _comentarios
            End If
            Dim abonado As Integer = _abonado
            Dim fecha_emision As Date = DateFecha.Value.ToString("yyyy-MM-dd")
            Dim path_excel As String = ""
            path_excel = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/antibiograma/" & idficha & ".xls"
            Dim path_pdf As String = ""
            path_pdf = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/antibiograma/" & idficha & ".pdf"
            Dim path_csv As String = ""
            path_csv = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/antibiograma/" & idficha & ".txt"
            Dim id_estado As Integer = 3
            aw_com.FICHA = idficha
            aw_com = aw_com.buscar
            If Not aw_com Is Nothing Then
                If comentarios <> "" Then
                    aw_com.COMENTARIO = comentarios
                End If
                aw_com.ABONADO = abonado
                Dim fechaemi As String
                fechaemi = Format(fecha_emision, "yyyy-MM-dd")
                aw_com.FECHA_EMISION = fechaemi
                aw_com.PATH_EXCEL = path_excel
                aw_com.PATH_PDF = path_pdf
                aw_com.PATH_CSV = path_csv
                aw_com.ID_ESTADO = id_estado
                If (aw_com.modificar2()) Then
                    'MsgBox("Solicitud guardada", MsgBoxStyle.Information, "Atención")
                    'limpiar()
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            Else
                Dim aweb_com As New dAntibiogramaWeb_com
                aweb_com.ID_USUARIO = idproductorweb_com
                If comentarios <> "" Then
                    aweb_com.COMENTARIO = comentarios
                End If
                aweb_com.ABONADO = abonado
                Dim fechaemi As String
                fechaemi = Format(fecha_emision, "yyyy-MM-dd")
                aweb_com.FECHA_CREADO = fechaemi
                aweb_com.FECHA_EMISION = fechaemi
                aweb_com.PATH_EXCEL = path_excel
                aweb_com.PATH_PDF = path_pdf
                aweb_com.PATH_CSV = path_csv
                aweb_com.FICHA = idficha
                aweb_com.ID_ESTADO = id_estado
                aweb_com.ID_LIBRO = idficha
                If (aweb_com.guardar()) Then
                    'MsgBox("Solicitud guardada", MsgBoxStyle.Information, "Atención")
                    'limpiar()
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            End If
        ElseIf tipoinforme = 5 Then 'SI EL TIPO DE INFORME ES DE PAL
            Dim palw_com As New dPalWeb_com
            Dim pw_com As New dProductorWeb_com
            pw_com.USUARIO = productorweb_com
            pw_com = pw_com.buscar
            Dim idproductorweb_com As Long = pw_com.ID
            Dim comentarios As String = ""
            If _comentarios.Length > 0 Then
                comentarios = _comentarios
            End If
            Dim abonado As Integer = _abonado
            Dim fecha_emision As Date = DateFecha.Value.ToString("yyyy-MM-dd")
            Dim path_excel As String = ""
            path_excel = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/pal/" & idficha & ".xls"
            Dim path_pdf As String = ""
            path_pdf = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/pal/" & idficha & ".pdf"
            Dim path_csv As String = ""
            path_csv = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/pal/" & idficha & ".txt"
            Dim id_estado As Integer = 3
            palw_com.FICHA = idficha
            palw_com = palw_com.buscar
            If Not palw_com Is Nothing Then
                If comentarios <> "" Then
                    palw_com.COMENTARIO = comentarios
                End If
                palw_com.ABONADO = abonado
                Dim fechaemi As String
                fechaemi = Format(fecha_emision, "yyyy-MM-dd")
                palw_com.FECHA_EMISION = fechaemi
                palw_com.PATH_EXCEL = path_excel
                palw_com.PATH_PDF = path_pdf
                palw_com.PATH_CSV = path_csv
                palw_com.ID_ESTADO = id_estado
                If (palw_com.modificar2()) Then
                    'MsgBox("Solicitud guardada", MsgBoxStyle.Information, "Atención")
                    'limpiar()
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            Else
                Dim palweb_com As New dPalWeb_com
                palweb_com.ID_USUARIO = idproductorweb_com
                If comentarios <> "" Then
                    palweb_com.COMENTARIO = comentarios
                End If
                palweb_com.ABONADO = abonado
                Dim fechaemi As String
                fechaemi = Format(fecha_emision, "yyyy-MM-dd")
                palweb_com.FECHA_CREADO = fechaemi
                palweb_com.FECHA_EMISION = fechaemi
                palweb_com.PATH_EXCEL = path_excel
                palweb_com.PATH_PDF = path_pdf
                palweb_com.PATH_CSV = path_csv
                palweb_com.FICHA = idficha
                palweb_com.ID_ESTADO = id_estado
                palweb_com.ID_LIBRO = idficha
                If (palweb_com.guardar()) Then
                    'MsgBox("Solicitud guardada", MsgBoxStyle.Information, "Atención")
                    'limpiar()
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            End If
        ElseIf tipoinforme = 6 Then 'SI EL TIPO DE INFORME ES DE PARASITOLOGÍA
            Dim paw_com As New dParasitologiaWeb_com
            Dim pw_com As New dProductorWeb_com
            pw_com.USUARIO = productorweb_com
            pw_com = pw_com.buscar
            Dim idproductorweb_com As Long = pw_com.ID
            Dim comentarios As String = ""
            If _comentarios.Length > 0 Then
                comentarios = _comentarios
            End If
            Dim abonado As Integer = _abonado
            Dim fecha_emision As Date = DateFecha.Value.ToString("yyyy-MM-dd")
            Dim path_excel As String = ""
            path_excel = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/parasitologia/" & idficha & ".xls"
            Dim path_pdf As String = ""
            path_pdf = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/parasitologia/" & idficha & ".pdf"
            Dim path_csv As String = ""
            path_csv = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/parasitologia/" & idficha & ".txt"
            Dim id_estado As Integer = 3
            paw_com.FICHA = idficha
            paw_com = paw_com.buscar
            If Not paw_com Is Nothing Then
                If comentarios <> "" Then
                    paw_com.COMENTARIO = comentarios
                End If
                paw_com.ABONADO = abonado
                Dim fechaemi As String
                fechaemi = Format(fecha_emision, "yyyy-MM-dd")
                paw_com.FECHA_EMISION = fechaemi
                paw_com.PATH_EXCEL = path_excel
                paw_com.PATH_PDF = path_pdf
                paw_com.PATH_CSV = path_csv
                paw_com.ID_ESTADO = id_estado
                If (paw_com.modificar2()) Then
                    'MsgBox("Solicitud guardada", MsgBoxStyle.Information, "Atención")
                    'limpiar()
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            Else
                Dim pweb_com As New dParasitologiaWeb_com
                pweb_com.ID_USUARIO = idproductorweb_com
                If comentarios <> "" Then
                    pweb_com.COMENTARIO = comentarios
                End If
                pweb_com.ABONADO = abonado
                Dim fechaemi As String
                fechaemi = Format(fecha_emision, "yyyy-MM-dd")
                pweb_com.FECHA_CREADO = fechaemi
                pweb_com.FECHA_EMISION = fechaemi
                pweb_com.PATH_EXCEL = path_excel
                pweb_com.PATH_PDF = path_pdf
                pweb_com.PATH_CSV = path_csv
                pweb_com.FICHA = idficha
                pweb_com.ID_ESTADO = id_estado
                pweb_com.ID_LIBRO = idficha
                If (pweb_com.guardar()) Then
                    'MsgBox("Solicitud guardada", MsgBoxStyle.Information, "Atención")
                    'limpiar()
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            End If
        ElseIf tipoinforme = 7 Then 'SI EL TIPO DE INFORME ES DE ALIMENTOS
            Dim spw_com As New dSubproductosWeb_com
            Dim pw_com As New dProductorWeb_com
            pw_com.USUARIO = productorweb_com
            pw_com = pw_com.buscar
            Dim idproductorweb_com As Long = pw_com.ID
            Dim comentarios As String = ""
            If _comentarios.Length > 0 Then
                comentarios = _comentarios
            End If
            Dim abonado As Integer = _abonado
            Dim fecha_emision As Date = DateFecha.Value.ToString("yyyy-MM-dd")
            Dim path_excel As String = ""
            path_excel = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/productos_subproductos/" & idficha & ".xls"
            Dim path_pdf As String = ""
            path_pdf = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/productos_subproductos/" & idficha & ".pdf"
            Dim path_csv As String = ""
            path_csv = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/productos_subproductos/" & idficha & ".txt"
            Dim id_estado As Integer = 3
            spw_com.FICHA = idficha
            spw_com = spw_com.buscar
            If Not spw_com Is Nothing Then
                If comentarios <> "" Then
                    spw_com.COMENTARIO = comentarios
                End If
                spw_com.ABONADO = abonado
                Dim fechaemi As String
                fechaemi = Format(fecha_emision, "yyyy-MM-dd")
                spw_com.FECHA_EMISION = fechaemi
                spw_com.PATH_EXCEL = path_excel
                spw_com.PATH_PDF = path_pdf
                spw_com.PATH_CSV = path_csv
                spw_com.ID_ESTADO = id_estado
                If (spw_com.modificar2()) Then
                    'MsgBox("Solicitud guardada", MsgBoxStyle.Information, "Atención")
                    'limpiar()
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            Else
                Dim spweb_com As New dSubproductosWeb_com
                spweb_com.ID_USUARIO = idproductorweb_com
                If comentarios <> "" Then
                    spweb_com.COMENTARIO = comentarios
                End If
                spweb_com.ABONADO = abonado
                Dim fechaemi As String
                fechaemi = Format(fecha_emision, "yyyy-MM-dd")
                spweb_com.FECHA_CREADO = fechaemi
                spweb_com.FECHA_EMISION = fechaemi
                spweb_com.PATH_EXCEL = path_excel
                spweb_com.PATH_PDF = path_pdf
                spweb_com.PATH_CSV = path_csv
                spweb_com.FICHA = idficha
                spweb_com.ID_ESTADO = id_estado
                spweb_com.ID_LIBRO = idficha
                If (spweb_com.guardar()) Then
                    'MsgBox("Solicitud guardada", MsgBoxStyle.Information, "Atención")
                    'limpiar()
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            End If
        ElseIf tipoinforme = 8 Then 'SI EL TIPO DE INFORME ES DE SEROLOGÍA
            Dim sw_com As New dSerologiaWeb_com
            Dim pw_com As New dProductorWeb_com
            pw_com.USUARIO = productorweb_com
            pw_com = pw_com.buscar
            Dim idproductorweb_com As Long = pw_com.ID
            Dim comentarios As String = ""
            If _comentarios.Length > 0 Then
                comentarios = _comentarios
            End If
            Dim abonado As Integer = _abonado
            Dim fecha_emision As Date = DateFecha.Value.ToString("yyyy-MM-dd")
            Dim path_excel As String = ""
            path_excel = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/serologia/" & idficha & ".xls"
            Dim path_pdf As String = ""
            path_pdf = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/serologia/" & idficha & ".pdf"
            Dim path_csv As String = ""
            path_csv = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/serologia/" & idficha & ".txt"
            Dim id_estado As Integer = 3
            sw_com.FICHA = idficha
            sw_com = sw_com.buscar
            If Not sw_com Is Nothing Then
                If comentarios <> "" Then
                    sw_com.COMENTARIO = comentarios
                End If
                sw_com.ABONADO = abonado
                Dim fechaemi As String
                fechaemi = Format(fecha_emision, "yyyy-MM-dd")
                sw_com.FECHA_EMISION = fechaemi
                sw_com.PATH_EXCEL = path_excel
                sw_com.PATH_PDF = path_pdf
                sw_com.PATH_CSV = path_csv
                sw_com.ID_ESTADO = id_estado
                If (sw_com.modificar2()) Then
                    'MsgBox("Solicitud guardada", MsgBoxStyle.Information, "Atención")
                    'limpiar()
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            Else
                Dim sweb_com As New dSerologiaWeb_com
                sweb_com.ID_USUARIO = idproductorweb_com
                If comentarios <> "" Then
                    sweb_com.COMENTARIO = comentarios
                End If
                sweb_com.ABONADO = abonado
                Dim fechaemi As String
                fechaemi = Format(fecha_emision, "yyyy-MM-dd")
                sweb_com.FECHA_CREADO = fechaemi
                sweb_com.FECHA_EMISION = fechaemi
                sweb_com.PATH_EXCEL = path_excel
                sweb_com.PATH_PDF = path_pdf
                sweb_com.PATH_CSV = path_csv
                sweb_com.FICHA = idficha
                sweb_com.ID_ESTADO = id_estado
                sweb_com.ID_LIBRO = idficha
                If (sweb_com.guardar()) Then
                    'MsgBox("Solicitud guardada", MsgBoxStyle.Information, "Atención")
                    'limpiar()
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            End If
        ElseIf tipoinforme = 9 Then 'SI EL TIPO DE INFORME ES DE PATOLOGÍA - TOXICOLOGÍA
            Dim patw_com As New dPatologiaWeb_com
            Dim pw_com As New dProductorWeb_com
            pw_com.USUARIO = productorweb_com
            pw_com = pw_com.buscar
            Dim idproductorweb_com As Long = pw_com.ID
            Dim comentarios As String = ""
            If _comentarios.Length > 0 Then
                comentarios = _comentarios
            End If
            Dim abonado As Integer = _abonado
            Dim fecha_emision As Date = DateFecha.Value.ToString("yyyy-MM-dd")
            Dim path_excel As String = ""
            path_excel = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/patologia/" & idficha & ".xls"
            Dim path_pdf As String = ""
            path_pdf = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/patologia/" & idficha & ".pdf"
            Dim path_csv As String = ""
            path_csv = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/patologia/" & idficha & ".txt"
            Dim id_estado As Integer = 3
            patw_com.FICHA = idficha
            patw_com = patw_com.buscar
            If Not patw_com Is Nothing Then
                If comentarios <> "" Then
                    patw_com.COMENTARIO = comentarios
                End If
                patw_com.ABONADO = abonado
                Dim fechaemi As String
                fechaemi = Format(fecha_emision, "yyyy-MM-dd")
                patw_com.FECHA_EMISION = fechaemi
                patw_com.PATH_EXCEL = path_excel
                patw_com.PATH_PDF = path_pdf
                patw_com.PATH_CSV = path_csv
                patw_com.ID_ESTADO = id_estado
                If (patw_com.modificar2()) Then
                    'MsgBox("Solicitud guardada", MsgBoxStyle.Information, "Atención")
                    'limpiar()
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            Else
                Dim patoweb_com As New dPatologiaWeb_com
                patoweb_com.ID_USUARIO = idproductorweb_com
                If comentarios <> "" Then
                    patoweb_com.COMENTARIO = comentarios
                End If
                patoweb_com.ABONADO = abonado
                Dim fechaemi As String
                fechaemi = Format(fecha_emision, "yyyy-MM-dd")
                patoweb_com.FECHA_CREADO = fechaemi
                patoweb_com.FECHA_EMISION = fechaemi
                patoweb_com.PATH_EXCEL = path_excel
                patoweb_com.PATH_PDF = path_pdf
                patoweb_com.PATH_CSV = path_csv
                patoweb_com.FICHA = idficha
                patoweb_com.ID_ESTADO = id_estado
                patoweb_com.ID_LIBRO = idficha
                If (patoweb_com.guardar()) Then
                    'MsgBox("Solicitud guardada", MsgBoxStyle.Information, "Atención")
                    'limpiar()
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            End If
        ElseIf tipoinforme = 10 Then 'SI EL TIPO DE INFORME ES DE CALIDAD
            Dim cw_com As New dCalidadWeb_com
            Dim pw_com As New dProductorWeb_com
            pw_com.USUARIO = productorweb_com
            pw_com = pw_com.buscar
            Dim idproductorweb_com As Long = pw_com.ID
            Dim comentarios As String = ""
            If _comentarios.Length > 0 Then
                comentarios = _comentarios
            End If
            Dim abonado As Integer = _abonado
            Dim fecha_emision As Date = DateFecha.Value.ToString("yyyy-MM-dd")
            Dim path_excel As String = ""
            path_excel = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/calidad_de_leche/" & idficha & ".xls"
            Dim path_pdf As String = ""
            path_pdf = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/calidad_de_leche/" & idficha & ".pdf"
            'Dim path_csv As String = ""
            'path_csv = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/calidad_de_leche/" & idficha & ".txt"
            Dim id_estado As Integer = 3
            cw_com.FICHA = idficha
            cw_com = cw_com.buscar
            If Not cw_com Is Nothing Then
                If comentarios <> "" Then
                    cw_com.COMENTARIO = comentarios
                End If
                cw_com.ABONADO = abonado
                Dim fechaemi As String
                fechaemi = Format(fecha_emision, "yyyy-MM-dd")
                cw_com.FECHA_EMISION = fechaemi
                cw_com.PATH_EXCEL = path_excel
                cw_com.PATH_PDF = path_pdf
                'cw_com.PATH_CSV = path_csv
                cw_com.ID_ESTADO = id_estado
                If (cw_com.modificar2()) Then
                    'MsgBox("Solicitud guardada", MsgBoxStyle.Information, "Atención")
                    'limpiar()
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            Else
                Dim calweb_com As New dCalidadWeb_com
                calweb_com.ID_USUARIO = idproductorweb_com
                If comentarios <> "" Then
                    calweb_com.COMENTARIO = comentarios
                End If
                calweb_com.ABONADO = abonado
                Dim fechaemi As String
                fechaemi = Format(fecha_emision, "yyyy-MM-dd")
                calweb_com.FECHA_CREADO = fechaemi
                calweb_com.FECHA_EMISION = fechaemi
                calweb_com.PATH_EXCEL = path_excel
                calweb_com.PATH_PDF = path_pdf
                'calweb_com.PATH_CSV = path_csv
                calweb_com.FICHA = idficha
                calweb_com.ID_ESTADO = id_estado
                calweb_com.ID_LIBRO = idficha
                If (calweb_com.guardar()) Then
                    'MsgBox("Solicitud guardada", MsgBoxStyle.Information, "Atención")
                    'limpiar()
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            End If
        ElseIf tipoinforme = 11 Then 'SI EL TIPO DE INFORME ES AMBIENTAL
            Dim aw_com As New dAmbientalWeb_com
            Dim pw_com As New dProductorWeb_com
            pw_com.USUARIO = productorweb_com
            pw_com = pw_com.buscar
            Dim idproductorweb_com As Long = pw_com.ID
            Dim comentarios As String = ""
            If _comentarios.Length > 0 Then
                comentarios = _comentarios
            End If
            Dim abonado As Integer = _abonado
            Dim fecha_emision As Date = DateFecha.Value.ToString("yyyy-MM-dd")
            Dim path_excel As String = ""
            path_excel = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/ambiental/" & idficha & ".xls"
            Dim path_pdf As String = ""
            path_pdf = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/ambiental/" & idficha & ".pdf"
            Dim path_csv As String = ""
            path_csv = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/ambiental/" & idficha & ".txt"
            Dim id_estado As Integer = 3
            aw_com.FICHA = idficha
            aw_com = aw_com.buscar
            If Not aw_com Is Nothing Then
                If comentarios <> "" Then
                    aw_com.COMENTARIO = comentarios
                End If
                aw_com.ABONADO = abonado
                Dim fechaemi As String
                fechaemi = Format(fecha_emision, "yyyy-MM-dd")
                aw_com.FECHA_EMISION = fechaemi
                aw_com.PATH_EXCEL = path_excel
                aw_com.PATH_PDF = path_pdf
                aw_com.PATH_CSV = path_csv
                aw_com.ID_ESTADO = id_estado
                If (aw_com.modificar2()) Then
                    'MsgBox("Solicitud guardada", MsgBoxStyle.Information, "Atención")
                    'limpiar()
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            Else
                Dim aweb_com As New dAmbientalWeb_com
                aweb_com.ID_USUARIO = idproductorweb_com
                If comentarios <> "" Then
                    aweb_com.COMENTARIO = comentarios
                End If
                aweb_com.ABONADO = abonado
                Dim fechaemi As String
                fechaemi = Format(fecha_emision, "yyyy-MM-dd")
                aweb_com.FECHA_CREADO = fechaemi
                aweb_com.FECHA_EMISION = fechaemi
                aweb_com.PATH_EXCEL = path_excel
                aweb_com.PATH_PDF = path_pdf
                aweb_com.PATH_CSV = path_csv
                aweb_com.FICHA = idficha
                aweb_com.ID_ESTADO = id_estado
                aweb_com.ID_LIBRO = idficha
                If (aweb_com.guardar()) Then
                    'MsgBox("Solicitud guardada", MsgBoxStyle.Information, "Atención")
                    'limpiar()
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            End If
        ElseIf tipoinforme = 12 Then 'SI EL TIPO DE INFORME ES DE LACTÓMETROS
            Dim lw_com As New dLactometrosWeb_com
            Dim pw_com As New dProductorWeb_com
            pw_com.USUARIO = productorweb_com
            pw_com = pw_com.buscar
            Dim idproductorweb_com As Long = pw_com.ID
            Dim comentarios As String = ""
            If _comentarios.Length > 0 Then
                comentarios = _comentarios
            End If
            Dim abonado As Integer = _abonado
            Dim fecha_emision As Date = DateFecha.Value.ToString("yyyy-MM-dd")
            Dim path_excel As String = ""
            path_excel = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/lactometros_chequeos_maquina/" & idficha & ".xls"
            Dim path_pdf As String = ""
            path_pdf = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/lactometros_chequeos_maquina/" & idficha & ".pdf"
            Dim path_csv As String = ""
            path_csv = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/lactometros_chequeos_maquina/" & idficha & ".txt"
            Dim id_estado As Integer = 3
            lw_com.FICHA = idficha
            lw_com = lw_com.buscar
            If Not lw_com Is Nothing Then
                If comentarios <> "" Then
                    lw_com.COMENTARIO = comentarios
                End If
                lw_com.ABONADO = abonado
                Dim fechaemi As String
                fechaemi = Format(fecha_emision, "yyyy-MM-dd")
                lw_com.FECHA_EMISION = fechaemi
                lw_com.PATH_EXCEL = path_excel
                lw_com.PATH_PDF = path_pdf
                lw_com.PATH_CSV = path_csv
                lw_com.ID_ESTADO = id_estado
                If (lw_com.modificar2()) Then
                    'MsgBox("Solicitud guardada", MsgBoxStyle.Information, "Atención")
                    'limpiar()
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            Else
                Dim lactweb_com As New dLactometrosWeb_com
                lactweb_com.ID_USUARIO = idproductorweb_com
                If comentarios <> "" Then
                    lactweb_com.COMENTARIO = comentarios
                End If
                lactweb_com.ABONADO = abonado
                Dim fechaemi As String
                fechaemi = Format(fecha_emision, "yyyy-MM-dd")
                lactweb_com.FECHA_CREADO = fechaemi
                lactweb_com.FECHA_EMISION = fechaemi
                lactweb_com.PATH_EXCEL = path_excel
                lactweb_com.PATH_PDF = path_pdf
                lactweb_com.PATH_CSV = path_csv
                lactweb_com.FICHA = idficha
                lactweb_com.ID_ESTADO = id_estado
                lactweb_com.ID_LIBRO = idficha
                If (lactweb_com.guardar()) Then
                    'MsgBox("Solicitud guardada", MsgBoxStyle.Information, "Atención")
                    'limpiar()
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            End If
        ElseIf tipoinforme = 13 Then 'SI EL TIPO DE INFORME ES DE NUTRICIÓN
            Dim aw_com As New dAgroNutricionWeb_com
            Dim pw_com As New dProductorWeb_com
            pw_com.USUARIO = productorweb_com
            pw_com = pw_com.buscar
            Dim idproductorweb_com As Long = pw_com.ID
            Dim comentarios As String = ""
            If _comentarios.Length > 0 Then
                comentarios = _comentarios
            End If
            Dim abonado As Integer = _abonado
            Dim fecha_emision As Date = DateFecha.Value.ToString("yyyy-MM-dd")
            Dim path_excel As String = ""
            path_excel = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/agro_nutricion/" & idficha & ".xls"
            Dim path_pdf As String = ""
            path_pdf = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/agro_nutricion/" & idficha & ".pdf"
            Dim path_csv As String = ""
            path_csv = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/agro_nutricion/" & idficha & ".txt"
            Dim id_estado As Integer = 3
            aw_com.FICHA = idficha
            aw_com = aw_com.buscar
            If Not aw_com Is Nothing Then
                If comentarios <> "" Then
                    aw_com.COMENTARIO = comentarios
                End If
                aw_com.ABONADO = abonado
                Dim fechaemi As String
                fechaemi = Format(fecha_emision, "yyyy-MM-dd")
                aw_com.FECHA_EMISION = fechaemi
                aw_com.PATH_EXCEL = path_excel
                aw_com.PATH_PDF = path_pdf
                aw_com.PATH_CSV = path_csv
                aw_com.ID_ESTADO = id_estado
                If (aw_com.modificar2()) Then
                    'MsgBox("Solicitud guardada", MsgBoxStyle.Information, "Atención")
                    'limpiar()
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            Else
                Dim aweb_com As New dAgroNutricionWeb_com
                aweb_com.ID_USUARIO = idproductorweb_com
                If comentarios <> "" Then
                    aweb_com.COMENTARIO = comentarios
                End If
                aweb_com.ABONADO = abonado
                Dim fechaemi As String
                fechaemi = Format(fecha_emision, "yyyy-MM-dd")
                aweb_com.FECHA_CREADO = fechaemi
                aweb_com.FECHA_EMISION = fechaemi
                aweb_com.PATH_EXCEL = path_excel
                aweb_com.PATH_PDF = path_pdf
                aweb_com.PATH_CSV = path_csv
                aweb_com.FICHA = idficha
                aweb_com.ID_ESTADO = id_estado
                aweb_com.ID_LIBRO = idficha
                If (aweb_com.guardar()) Then
                    'MsgBox("Solicitud guardada", MsgBoxStyle.Information, "Atención")
                    'limpiar()
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            End If
        ElseIf tipoinforme = 14 Then 'SI EL TIPO DE INFORME ES DE SUELOS
            Dim aw_com As New dAgroSuelosWeb_com
            Dim pw_com As New dProductorWeb_com
            pw_com.USUARIO = productorweb_com
            pw_com = pw_com.buscar
            Dim idproductorweb_com As Long = pw_com.ID
            Dim comentarios As String = ""
            If _comentarios.Length > 0 Then
                comentarios = _comentarios
            End If
            Dim abonado As Integer = _abonado
            Dim fecha_emision As Date = DateFecha.Value.ToString("yyyy-MM-dd")
            Dim path_excel As String = ""
            path_excel = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/agro_suelos/" & idficha & ".xls"
            Dim path_pdf As String = ""
            path_pdf = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/agro_suelos/" & idficha & ".pdf"
            Dim path_csv As String = ""
            path_csv = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/agro_suelos/" & idficha & ".txt"
            Dim id_estado As Integer = 3
            aw_com.FICHA = idficha
            aw_com = aw_com.buscar
            If Not aw_com Is Nothing Then
                If comentarios <> "" Then
                    aw_com.COMENTARIO = comentarios
                End If
                aw_com.ABONADO = abonado
                Dim fechaemi As String
                fechaemi = Format(fecha_emision, "yyyy-MM-dd")
                aw_com.FECHA_EMISION = fechaemi
                aw_com.PATH_EXCEL = path_excel
                aw_com.PATH_PDF = path_pdf
                aw_com.PATH_CSV = path_csv
                aw_com.ID_ESTADO = id_estado
                If (aw_com.modificar2()) Then
                    'MsgBox("Solicitud guardada", MsgBoxStyle.Information, "Atención")
                    'limpiar()
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            Else
                Dim aweb_com As New dAgroSuelosWeb_com
                aweb_com.ID_USUARIO = idproductorweb_com
                If comentarios <> "" Then
                    aweb_com.COMENTARIO = comentarios
                End If
                aweb_com.ABONADO = abonado
                Dim fechaemi As String
                fechaemi = Format(fecha_emision, "yyyy-MM-dd")
                aweb_com.FECHA_CREADO = fechaemi
                aweb_com.FECHA_EMISION = fechaemi
                aweb_com.PATH_EXCEL = path_excel
                aweb_com.PATH_PDF = path_pdf
                aweb_com.PATH_CSV = path_csv
                aweb_com.FICHA = idficha
                aweb_com.ID_ESTADO = id_estado
                aweb_com.ID_LIBRO = idficha
                If (aweb_com.guardar()) Then
                    'MsgBox("Solicitud guardada", MsgBoxStyle.Information, "Atención")
                    'limpiar()
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            End If
        ElseIf tipoinforme = 15 Then 'SI EL TIPO DE INFORME ES DE BRUCELOSIS EN LECHE
            Dim bw_com As New dBrucelosisLecheWeb_com
            Dim pw_com As New dProductorWeb_com
            pw_com.USUARIO = productorweb_com
            pw_com = pw_com.buscar
            Dim idproductorweb_com As Long = pw_com.ID
            Dim comentarios As String = ""
            If _comentarios.Length > 0 Then
                comentarios = _comentarios
            End If
            Dim abonado As Integer = _abonado
            Dim fecha_emision As Date = DateFecha.Value.ToString("yyyy-MM-dd")
            Dim path_excel As String = ""
            path_excel = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/brucelosis_leche/" & idficha & ".xls"
            Dim path_pdf As String = ""
            path_pdf = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/brucelosis_leche/" & idficha & ".pdf"
            Dim path_csv As String = ""
            path_csv = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/brucelosis_leche/" & idficha & ".txt"
            Dim id_estado As Integer = 3
            bw_com.FICHA = idficha
            bw_com = bw_com.buscar
            If Not bw_com Is Nothing Then
                If comentarios <> "" Then
                    bw_com.COMENTARIO = comentarios
                End If
                bw_com.ABONADO = abonado
                Dim fechaemi As String
                fechaemi = Format(fecha_emision, "yyyy-MM-dd")
                bw_com.FECHA_EMISION = fechaemi
                bw_com.PATH_EXCEL = path_excel
                bw_com.PATH_PDF = path_pdf
                bw_com.PATH_CSV = path_csv
                bw_com.ID_ESTADO = id_estado
                If (bw_com.modificar2()) Then
                    'MsgBox("Solicitud guardada", MsgBoxStyle.Information, "Atención")
                    'limpiar()
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            Else
                Dim bweb_com As New dBrucelosisLecheWeb_com
                bweb_com.ID_USUARIO = idproductorweb_com
                If comentarios <> "" Then
                    bweb_com.COMENTARIO = comentarios
                End If
                bweb_com.ABONADO = abonado
                Dim fechaemi As String
                fechaemi = Format(fecha_emision, "yyyy-MM-dd")
                bweb_com.FECHA_CREADO = fechaemi
                bweb_com.FECHA_EMISION = fechaemi
                bweb_com.PATH_EXCEL = path_excel
                bweb_com.PATH_PDF = path_pdf
                bweb_com.PATH_CSV = path_csv
                bweb_com.FICHA = idficha
                bweb_com.ID_ESTADO = id_estado
                bweb_com.ID_LIBRO = idficha
                If (bweb_com.guardar()) Then
                    'MsgBox("Solicitud guardada", MsgBoxStyle.Information, "Atención")
                    'limpiar()
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            End If
        ElseIf tipoinforme = 16 Then 'SI EL TIPO DE INFORME ES DE EFLUENTES
            Dim bw_com As New dBrucelosisLecheWeb_com
            Dim pw_com As New dProductorWeb_com
            pw_com.USUARIO = productorweb_com
            pw_com = pw_com.buscar
            Dim idproductorweb_com As Long = pw_com.ID
            Dim comentarios As String = ""
            If _comentarios.Length > 0 Then
                comentarios = _comentarios
            End If
            Dim abonado As Integer = _abonado
            Dim fecha_emision As Date = DateFecha.Value.ToString("yyyy-MM-dd")
            Dim path_excel As String = ""
            path_excel = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/efluentes/" & idficha & ".xls"
            Dim path_pdf As String = ""
            path_pdf = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/efluentes/" & idficha & ".pdf"
            Dim path_csv As String = ""
            path_csv = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/efluentes/" & idficha & ".txt"
            Dim id_estado As Integer = 3
            bw_com.FICHA = idficha
            bw_com = bw_com.buscar
            If Not bw_com Is Nothing Then
                If comentarios <> "" Then
                    bw_com.COMENTARIO = comentarios
                End If
                bw_com.ABONADO = abonado
                Dim fechaemi As String
                fechaemi = Format(fecha_emision, "yyyy-MM-dd")
                bw_com.FECHA_EMISION = fechaemi
                bw_com.PATH_EXCEL = path_excel
                bw_com.PATH_PDF = path_pdf
                bw_com.PATH_CSV = path_csv
                bw_com.ID_ESTADO = id_estado
                If (bw_com.modificar2()) Then
                    'MsgBox("Solicitud guardada", MsgBoxStyle.Information, "Atención")
                    'limpiar()
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            Else
                Dim bweb_com As New dBrucelosisLecheWeb_com
                bweb_com.ID_USUARIO = idproductorweb_com
                If comentarios <> "" Then
                    bweb_com.COMENTARIO = comentarios
                End If
                bweb_com.ABONADO = abonado
                Dim fechaemi As String
                fechaemi = Format(fecha_emision, "yyyy-MM-dd")
                bweb_com.FECHA_CREADO = fechaemi
                bweb_com.FECHA_EMISION = fechaemi
                bweb_com.PATH_EXCEL = path_excel
                bweb_com.PATH_PDF = path_pdf
                bweb_com.PATH_CSV = path_csv
                bweb_com.FICHA = idficha
                bweb_com.ID_ESTADO = id_estado
                bweb_com.ID_LIBRO = idficha
                If (bweb_com.guardar()) Then
                    'MsgBox("Solicitud guardada", MsgBoxStyle.Information, "Atención")
                    'limpiar()
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            End If
        ElseIf tipoinforme = 17 Then 'SI EL TIPO DE INFORME ES DE BACTERIOLOGÍA
            Dim aw_com As New dAntibiogramaWeb_com
            Dim pw_com As New dProductorWeb_com
            pw_com.USUARIO = productorweb_com
            pw_com = pw_com.buscar
            Dim idproductorweb_com As Long = pw_com.ID
            Dim comentarios As String = ""
            If _comentarios.Length > 0 Then
                comentarios = _comentarios
            End If
            Dim abonado As Integer = _abonado
            Dim fecha_emision As Date = DateFecha.Value.ToString("yyyy-MM-dd")
            Dim path_excel As String = ""
            path_excel = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/antibiograma/" & idficha & ".xls"
            Dim path_pdf As String = ""
            path_pdf = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/antibiograma/" & idficha & ".pdf"
            Dim path_csv As String = ""
            path_csv = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/antibiograma/" & idficha & ".txt"
            Dim id_estado As Integer = 3
            aw_com.FICHA = idficha
            aw_com = aw_com.buscar
            If Not aw_com Is Nothing Then
                If comentarios <> "" Then
                    aw_com.COMENTARIO = comentarios
                End If
                aw_com.ABONADO = abonado
                Dim fechaemi As String
                fechaemi = Format(fecha_emision, "yyyy-MM-dd")
                aw_com.FECHA_EMISION = fechaemi
                aw_com.PATH_EXCEL = path_excel
                aw_com.PATH_PDF = path_pdf
                aw_com.PATH_CSV = path_csv
                aw_com.ID_ESTADO = id_estado
                If (aw_com.modificar2()) Then
                    'MsgBox("Solicitud guardada", MsgBoxStyle.Information, "Atención")
                    'limpiar()
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            Else
                Dim aweb_com As New dAntibiogramaWeb_com
                aweb_com.ID_USUARIO = idproductorweb_com
                If comentarios <> "" Then
                    aweb_com.COMENTARIO = comentarios
                End If
                aweb_com.ABONADO = abonado
                Dim fechaemi As String
                fechaemi = Format(fecha_emision, "yyyy-MM-dd")
                aweb_com.FECHA_CREADO = fechaemi
                aweb_com.FECHA_EMISION = fechaemi
                aweb_com.PATH_EXCEL = path_excel
                aweb_com.PATH_PDF = path_pdf
                aweb_com.PATH_CSV = path_csv
                aweb_com.FICHA = idficha
                aweb_com.ID_ESTADO = id_estado
                aweb_com.ID_LIBRO = idficha
                If (aweb_com.guardar()) Then
                    'MsgBox("Solicitud guardada", MsgBoxStyle.Information, "Atención")
                    'limpiar()
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            End If
        ElseIf tipoinforme = 18 Then 'SI EL TIPO DE INFORME ES DE BACTERIOLOGÍA CLÍNICA AERÓBICA
            Dim aw_com As New dAntibiogramaWeb_com
            Dim pw_com As New dProductorWeb_com
            pw_com.USUARIO = productorweb_com
            pw_com = pw_com.buscar
            Dim idproductorweb_com As Long = pw_com.ID
            Dim comentarios As String = ""
            If _comentarios.Length > 0 Then
                comentarios = _comentarios
            End If
            Dim abonado As Integer = _abonado
            Dim fecha_emision As Date = DateFecha.Value.ToString("yyyy-MM-dd")
            Dim path_excel As String = ""
            path_excel = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/antibiograma/" & idficha & ".xls"
            Dim path_pdf As String = ""
            path_pdf = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/antibiograma/" & idficha & ".pdf"
            Dim path_csv As String = ""
            path_csv = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/antibiograma/" & idficha & ".txt"
            Dim id_estado As Integer = 3
            aw_com.FICHA = idficha
            aw_com = aw_com.buscar
            If Not aw_com Is Nothing Then
                If comentarios <> "" Then
                    aw_com.COMENTARIO = comentarios
                End If
                aw_com.ABONADO = abonado
                Dim fechaemi As String
                fechaemi = Format(fecha_emision, "yyyy-MM-dd")
                aw_com.FECHA_EMISION = fechaemi
                aw_com.PATH_EXCEL = path_excel
                aw_com.PATH_PDF = path_pdf
                aw_com.PATH_CSV = path_csv
                aw_com.ID_ESTADO = id_estado
                If (aw_com.modificar2()) Then
                    'MsgBox("Solicitud guardada", MsgBoxStyle.Information, "Atención")
                    'limpiar()
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            Else
                Dim aweb_com As New dAntibiogramaWeb_com
                aweb_com.ID_USUARIO = idproductorweb_com
                If comentarios <> "" Then
                    aweb_com.COMENTARIO = comentarios
                End If
                aweb_com.ABONADO = abonado
                Dim fechaemi As String
                fechaemi = Format(fecha_emision, "yyyy-MM-dd")
                aweb_com.FECHA_CREADO = fechaemi
                aweb_com.FECHA_EMISION = fechaemi
                aweb_com.PATH_EXCEL = path_excel
                aweb_com.PATH_PDF = path_pdf
                aweb_com.PATH_CSV = path_csv
                aweb_com.FICHA = idficha
                aweb_com.ID_ESTADO = id_estado
                aweb_com.ID_LIBRO = idficha
                If (aweb_com.guardar()) Then
                    'MsgBox("Solicitud guardada", MsgBoxStyle.Information, "Atención")
                    'limpiar()
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            End If
        ElseIf tipoinforme = 19 Then 'SI EL TIPO DE INFORME ES DE FOLIAR
            Dim aw_com As New dAgroSuelosWeb_com
            Dim pw_com As New dProductorWeb_com
            pw_com.USUARIO = productorweb_com
            pw_com = pw_com.buscar
            Dim idproductorweb_com As Long = pw_com.ID
            Dim comentarios As String = ""
            If _comentarios.Length > 0 Then
                comentarios = _comentarios
            End If
            Dim abonado As Integer = _abonado
            Dim fecha_emision As Date = DateFecha.Value.ToString("yyyy-MM-dd")
            Dim path_excel As String = ""
            path_excel = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/agro_suelos/" & idficha & ".xls"
            Dim path_pdf As String = ""
            path_pdf = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/agro_suelos/" & idficha & ".pdf"
            Dim path_csv As String = ""
            path_csv = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/agro_suelos/" & idficha & ".txt"
            Dim id_estado As Integer = 3
            aw_com.FICHA = idficha
            aw_com = aw_com.buscar
            If Not aw_com Is Nothing Then
                If comentarios <> "" Then
                    aw_com.COMENTARIO = comentarios
                End If
                aw_com.ABONADO = abonado
                Dim fechaemi As String
                fechaemi = Format(fecha_emision, "yyyy-MM-dd")
                aw_com.FECHA_EMISION = fechaemi
                aw_com.PATH_EXCEL = path_excel
                aw_com.PATH_PDF = path_pdf
                aw_com.PATH_CSV = path_csv
                aw_com.ID_ESTADO = id_estado
                If (aw_com.modificar2()) Then
                    'MsgBox("Solicitud guardada", MsgBoxStyle.Information, "Atención")
                    'limpiar()
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            Else
                Dim aweb_com As New dAgroSuelosWeb_com
                aweb_com.ID_USUARIO = idproductorweb_com
                If comentarios <> "" Then
                    aweb_com.COMENTARIO = comentarios
                End If
                aweb_com.ABONADO = abonado
                Dim fechaemi As String
                fechaemi = Format(fecha_emision, "yyyy-MM-dd")
                aweb_com.FECHA_CREADO = fechaemi
                aweb_com.FECHA_EMISION = fechaemi
                aweb_com.PATH_EXCEL = path_excel
                aweb_com.PATH_PDF = path_pdf
                aweb_com.PATH_CSV = path_csv
                aweb_com.FICHA = idficha
                aweb_com.ID_ESTADO = id_estado
                aweb_com.ID_LIBRO = idficha
                If (aweb_com.guardar()) Then
                    'MsgBox("Solicitud guardada", MsgBoxStyle.Information, "Atención")
                    'limpiar()
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            End If
        ElseIf tipoinforme = 99 Then 'SI EL TIPO DE INFORME ES DE OTROS SERVICIOS
            Dim ow_com As New dOtrosServiciosWeb_com
            Dim pw_com As New dProductorWeb_com
            pw_com.USUARIO = productorweb_com
            pw_com = pw_com.buscar
            Dim idproductorweb_com As Long = pw_com.ID
            Dim comentarios As String = ""
            If _comentarios.Length > 0 Then
                comentarios = _comentarios
            End If
            Dim abonado As Integer = _abonado
            Dim fecha_emision As Date = DateFecha.Value.ToString("yyyy-MM-dd")
            Dim path_excel As String = ""
            path_excel = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/otros_servicios/" & idficha & ".xls"
            Dim path_pdf As String = ""
            path_pdf = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/otros_servicios/" & idficha & ".pdf"
            Dim path_csv As String = ""
            path_csv = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/otros_servicios/" & idficha & ".txt"
            Dim id_estado As Integer = 3
            ow_com.FICHA = idficha
            ow_com = ow_com.buscar
            If Not ow_com Is Nothing Then
                If comentarios <> "" Then
                    ow_com.COMENTARIO = comentarios
                End If
                ow_com.ABONADO = abonado
                Dim fechaemi As String
                fechaemi = Format(fecha_emision, "yyyy-MM-dd")
                ow_com.FECHA_EMISION = fechaemi
                ow_com.PATH_EXCEL = path_excel
                ow_com.PATH_PDF = path_pdf
                ow_com.PATH_CSV = path_csv
                ow_com.ID_ESTADO = id_estado
                If (ow_com.modificar2()) Then
                    'MsgBox("Solicitud guardada", MsgBoxStyle.Information, "Atención")
                    'limpiar()
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            Else
                Dim oweb_com As New dOtrosServiciosWeb_com
                oweb_com.ID_USUARIO = idproductorweb_com
                If comentarios <> "" Then
                    oweb_com.COMENTARIO = comentarios
                End If
                oweb_com.ABONADO = abonado
                Dim fechaemi As String
                fechaemi = Format(fecha_emision, "yyyy-MM-dd")
                oweb_com.FECHA_CREADO = fechaemi
                oweb_com.FECHA_EMISION = fechaemi
                oweb_com.PATH_EXCEL = path_excel
                oweb_com.PATH_PDF = path_pdf
                oweb_com.PATH_CSV = path_csv
                oweb_com.FICHA = idficha
                oweb_com.ID_ESTADO = id_estado
                oweb_com.ID_LIBRO = idficha
                If (oweb_com.guardar()) Then
                    'MsgBox("Solicitud guardada", MsgBoxStyle.Information, "Atención")
                    'limpiar()
                Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
                End If
            End If
        End If
        enviar_notificacion_resultado()
        '*** CREA RESULTADO EN GESTOR NUEVO *******************************************************************************************
        Dim resultado As New Dictionary(Of String, dResultado)
        Dim carpeta As String = ""
        If tipoinforme = 1 Then
            carpeta = "control_lechero"
        ElseIf tipoinforme = 3 Then
            carpeta = "agua"
        ElseIf tipoinforme = 4 Then
            carpeta = "antibiograma"
        ElseIf tipoinforme = 6 Then
            carpeta = "parasitologia"
        ElseIf tipoinforme = 7 Then
            carpeta = "productos_subproductos"
        ElseIf tipoinforme = 8 Then
            carpeta = "serologia"
        ElseIf tipoinforme = 9 Then
            carpeta = "patologia"
        ElseIf tipoinforme = 10 Then
            carpeta = "calidad_de_leche"
        ElseIf tipoinforme = 11 Then
            carpeta = "ambiental"
        ElseIf tipoinforme = 13 Then
            carpeta = "agro_nutricion"
        ElseIf tipoinforme = 14 Then
            carpeta = "agro_suelos"
        ElseIf tipoinforme = 15 Then
            carpeta = "brucelosis_leche"
        ElseIf tipoinforme = 16 Then
            carpeta = "efluentes"
        ElseIf tipoinforme = 17 Then
            carpeta = "antibiograma"
        ElseIf tipoinforme = 18 Then
            carpeta = "antibiograma"
        ElseIf tipoinforme = 19 Then
            carpeta = "agro_suelos"
        End If
        'If tipoinforme = 19 Then
        '    tipoinforme = 14
        'End If
        Dim rg As New dResultado
        Dim fechaemi2 As String
        Dim fecha_emision2 As Date = DateFecha.Value.ToString("yyyy-MM-dd")
        fechaemi2 = Format(fecha_emision2, "yyyy-MM-dd")
        rg.ficha = idficha
        rg.comentarios = _comentarios
        rg.idnet_usuario = idnet
        rg.abonado = _abonado
        rg.fecha_creado = fechaemi2
        rg.fecha_emision = fechaemi2
        rg.path_excel = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/" & carpeta & "/" & idficha & ".xls"
        rg.path_pdf = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/" & carpeta & "/" & idficha & ".pdf"
        rg.path_csv = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/" & carpeta & "/" & idficha & ".txt"
        rg.id_estado = 3
        rg.id_libro = idficha
        rg.idnet_tipo_informe = tipoinforme
        resultado.Add("resultado", rg)
        Dim parameters As String = JsonConvert.SerializeObject(resultado, Formatting.None)
        Dim status As HttpStatusCode = HttpStatusCode.ExpectationFailed
        Dim response As Byte() = PostResponse("http://colaveco-gestor.herokuapp.com/resultados", "POST", parameters, status)
    End Sub
    Private Sub moverexcel()
        Dim sRutaDestino As String
        If tipoinforme = 1 Then
            sRutaDestino = "\\192.168.1.10\e\NET\CONTROL_LECHERO\" & idficha & ".xls"
        ElseIf tipoinforme = 3 Then
            sRutaDestino = "\\192.168.1.10\e\NET\AGUA\" & idficha & ".xls"
        ElseIf tipoinforme = 4 Then
            sRutaDestino = "\\192.168.1.10\e\NET\ANTIBIOGRAMA\" & idficha & ".xls"
        ElseIf tipoinforme = 5 Then
            sRutaDestino = "\\192.168.1.10\e\NET\PAL\" & idficha & ".xls"
        ElseIf tipoinforme = 6 Then
            sRutaDestino = "\\192.168.1.10\e\NET\PARASITOLOGIA\" & idficha & ".xls"
        ElseIf tipoinforme = 7 Then
            sRutaDestino = "\\192.168.1.10\e\NET\ALIMENTOS\" & idficha & ".xls"
        ElseIf tipoinforme = 8 Then
            sRutaDestino = "\\192.168.1.10\e\NET\SEROLOGIA\" & idficha & ".xls"
        ElseIf tipoinforme = 9 Then
            sRutaDestino = "\\192.168.1.10\e\NET\PATOLOGIA\" & idficha & ".xls"
        ElseIf tipoinforme = 10 Then
            sRutaDestino = "\\192.168.1.10\e\NET\CALIDAD\" & idficha & ".xls"
        ElseIf tipoinforme = 11 Then
            sRutaDestino = "\\192.168.1.10\e\NET\AMBIENTAL\" & idficha & ".xls"
        ElseIf tipoinforme = 12 Then
            sRutaDestino = "\\192.168.1.10\e\NET\LACTOMETROS\" & idficha & ".xls"
        ElseIf tipoinforme = 13 Then
            sRutaDestino = "\\192.168.1.10\e\NET\NUTRICION\" & idficha & ".xls"
        ElseIf tipoinforme = 14 Then
            sRutaDestino = "\\192.168.1.10\e\NET\SUELOS\" & idficha & ".xls"
        ElseIf tipoinforme = 15 Then
            sRutaDestino = "\\192.168.1.10\e\NET\Bruselocis en leche\" & idficha & ".xls"
        ElseIf tipoinforme = 16 Then
            sRutaDestino = "\\192.168.1.10\e\NET\Efluentes\" & idficha & ".xls"
        ElseIf tipoinforme = 17 Then
            sRutaDestino = "\\192.168.1.10\e\NET\BACTEREOLOGIATANQUE\" & idficha & ".xls"
        ElseIf tipoinforme = 18 Then
            sRutaDestino = "\\192.168.1.10\e\NET\BACTEREOLOGIACLINICAAEROBICA\" & idficha & ".xls"
        ElseIf tipoinforme = 19 Then
            sRutaDestino = "\\192.168.1.10\e\NET\FOLIARES\" & idficha & ".xls"
        ElseIf tipoinforme = 20 Then
            sRutaDestino = "\\192.168.1.10\e\NET\TOXICOLOGIA\" & idficha & ".xls"
        ElseIf tipoinforme = 99 Then
            sRutaDestino = "\\192.168.1.10\e\NET\OTROS\" & idficha & ".xls"
        End If
        'Dim sArchivoOrigen As String = "\\192.168.1.10\e\NET\INFORMES PARA SUBIR\" & idficha & ".txt"
        Dim sArchivoOrigen As String = "C:\INFORMES PARA SUBIR\" & idficha & ".xls"

        Try
            ' Mover el fichero.si existe lo sobreescribe  
            My.Computer.FileSystem.MoveFile(sArchivoOrigen, _
                                            sRutaDestino, _
                                            True)
            'MsgBox("Ok.", MsgBoxStyle.Information, "Mover archivo")
            ' errores  
        Catch ex As Exception
            'MsgBox(ex.Message.ToString, MsgBoxStyle.Critical)
        End Try
    End Sub
    Private Sub moverexcel_otrapc()
        If tipoinforme = 10 Then
            'Dim sArchivoOrigen As String = "\\192.168.1.10\e\NET\INFORMES PARA SUBIR\" & idficha & ".xls"
            Dim sArchivoOrigen As String = "\\ROBOT\INFORMES PARA SUBIR\" & idficha & ".xls"
            Dim sRutaDestino As String = "\\192.168.1.10\e\NET\CALIDAD\" & idficha & ".xls"
            Try
                ' Mover el fichero.si existe lo sobreescribe  
                My.Computer.FileSystem.MoveFile(sArchivoOrigen, _
                                                sRutaDestino, _
                                                True)
                'MsgBox("Ok.", MsgBoxStyle.Information, "Mover archivo")
                ' errores  
            Catch ex As Exception
                'MsgBox(ex.Message.ToString, MsgBoxStyle.Critical)
            End Try
        ElseIf tipoinforme = 1 Then
            'Dim sArchivoOrigen As String = "\\192.168.1.10\e\NET\INFORMES PARA SUBIR\" & idficha & ".xls"
            Dim sArchivoOrigen As String = "\\ROBOT\INFORMES PARA SUBIR\" & idficha & ".xls"
            Dim sRutaDestino As String = "\\192.168.1.10\e\NET\CONTROL_LECHERO\" & idficha & ".xls"
            Try
                ' Mover el fichero.si existe lo sobreescribe  
                My.Computer.FileSystem.MoveFile(sArchivoOrigen, _
                                                sRutaDestino, _
                                                True)
                'MsgBox("Ok.", MsgBoxStyle.Information, "Mover archivo")
                ' errores  
            Catch ex As Exception
                'MsgBox(ex.Message.ToString, MsgBoxStyle.Critical)
            End Try
        ElseIf tipoinforme = 3 Then
            'Dim sArchivoOrigen As String = "\\192.168.1.10\e\NET\INFORMES PARA SUBIR\" & idficha & ".xls"
            Dim sArchivoOrigen As String = "\\ROBOT\INFORMES PARA SUBIR\" & idficha & ".xls"
            Dim sRutaDestino As String = "\\192.168.1.10\e\NET\AGUA\" & idficha & ".xls"
            Try
                ' Mover el fichero.si existe lo sobreescribe  
                My.Computer.FileSystem.MoveFile(sArchivoOrigen, _
                                                sRutaDestino, _
                                                True)
                'MsgBox("Ok.", MsgBoxStyle.Information, "Mover archivo")
                ' errores  
            Catch ex As Exception
                'MsgBox(ex.Message.ToString, MsgBoxStyle.Critical)
            End Try
        ElseIf tipoinforme = 4 Then
            'Dim sArchivoOrigen As String = "\\192.168.1.10\e\NET\INFORMES PARA SUBIR\" & idficha & ".xls"
            Dim sArchivoOrigen As String = "\\ROBOT\INFORMES PARA SUBIR\" & idficha & ".xls"
            Dim sRutaDestino As String = "\\192.168.1.10\e\NET\ANTIBIOGRAMA\" & idficha & ".xls"
            Try
                ' Mover el fichero.si existe lo sobreescribe  
                My.Computer.FileSystem.MoveFile(sArchivoOrigen, _
                                                sRutaDestino, _
                                                True)
                'MsgBox("Ok.", MsgBoxStyle.Information, "Mover archivo")
                ' errores  
            Catch ex As Exception
                'MsgBox(ex.Message.ToString, MsgBoxStyle.Critical)
            End Try
        ElseIf tipoinforme = 7 Then
            'Dim sArchivoOrigen As String = "\\192.168.1.10\e\NET\INFORMES PARA SUBIR\" & idficha & ".xls"
            Dim sArchivoOrigen As String = "\\ROBOT\INFORMES PARA SUBIR\" & idficha & ".xls"
            Dim sRutaDestino As String = "\\192.168.1.10\e\NET\ALIMENTOS\" & idficha & ".xls"
            Try
                ' Mover el fichero.si existe lo sobreescribe  
                My.Computer.FileSystem.MoveFile(sArchivoOrigen, _
                                                sRutaDestino, _
                                                True)
                'MsgBox("Ok.", MsgBoxStyle.Information, "Mover archivo")
                ' errores  
            Catch ex As Exception
                'MsgBox(ex.Message.ToString, MsgBoxStyle.Critical)
            End Try
        ElseIf tipoinforme = 13 Then
            'Dim sArchivoOrigen As String = "\\192.168.1.10\e\NET\INFORMES PARA SUBIR\" & idficha & ".xls"
            Dim sArchivoOrigen As String = "\\ROBOT\INFORMES PARA SUBIR\" & idficha & ".xls"
            Dim sRutaDestino As String = "\\192.168.1.10\e\NET\NUTRICION\" & idficha & ".xls"
            Try
                ' Mover el fichero.si existe lo sobreescribe  
                My.Computer.FileSystem.MoveFile(sArchivoOrigen, _
                                                sRutaDestino, _
                                                True)
                'MsgBox("Ok.", MsgBoxStyle.Information, "Mover archivo")
                ' errores  
            Catch ex As Exception
                'MsgBox(ex.Message.ToString, MsgBoxStyle.Critical)
            End Try
        ElseIf tipoinforme = 14 Then
            'Dim sArchivoOrigen As String = "\\192.168.1.10\e\NET\INFORMES PARA SUBIR\" & idficha & ".xls"
            Dim sArchivoOrigen As String = "\\ROBOT\INFORMES PARA SUBIR\" & idficha & ".xls"
            Dim sRutaDestino As String = "\\192.168.1.10\e\NET\Suelos\" & idficha & ".xls"
            Try
                ' Mover el fichero.si existe lo sobreescribe  
                My.Computer.FileSystem.MoveFile(sArchivoOrigen, _
                                                sRutaDestino, _
                                                True)
                'MsgBox("Ok.", MsgBoxStyle.Information, "Mover archivo")
                ' errores  
            Catch ex As Exception
                'MsgBox(ex.Message.ToString, MsgBoxStyle.Critical)
            End Try
        ElseIf tipoinforme = 15 Then
            'Dim sArchivoOrigen As String = "\\192.168.1.10\e\NET\INFORMES PARA SUBIR\" & idficha & ".xls"
            Dim sArchivoOrigen As String = "\\ROBOT\INFORMES PARA SUBIR\" & idficha & ".xls"
            Dim sRutaDestino As String = "\\192.168.1.10\e\NET\Brucelosis en leche\" & idficha & ".xls"
            Try
                ' Mover el fichero.si existe lo sobreescribe  
                My.Computer.FileSystem.MoveFile(sArchivoOrigen, _
                                                sRutaDestino, _
                                                True)
                'MsgBox("Ok.", MsgBoxStyle.Information, "Mover archivo")
                ' errores  
            Catch ex As Exception
                'MsgBox(ex.Message.ToString, MsgBoxStyle.Critical)
            End Try
        ElseIf tipoinforme = 11 Then
            Dim sArchivoOrigen As String = "\\ROBOT\INFORMES PARA SUBIR\" & idficha & ".xls"
            Dim sRutaDestino As String = "\\192.168.1.10\e\NET\AMBIENTAL\" & idficha & ".xls"
            Try
                ' Mover el fichero.si existe lo sobreescribe  
                My.Computer.FileSystem.MoveFile(sArchivoOrigen, _
                                                sRutaDestino, _
                                                True)
                'MsgBox("Ok.", MsgBoxStyle.Information, "Mover archivo")
                ' errores  
            Catch ex As Exception
                'MsgBox(ex.Message.ToString, MsgBoxStyle.Critical)
            End Try
        End If
    End Sub
    Private Sub moverpdf()
        Dim sRutaDestino As String
        If tipoinforme = 1 Then
            sRutaDestino = "\\192.168.1.10\e\NET\CONTROL_LECHERO\" & idficha & ".pdf"
        ElseIf tipoinforme = 3 Then
            sRutaDestino = "\\192.168.1.10\e\NET\AGUA\" & idficha & ".pdf"
        ElseIf tipoinforme = 4 Then
            sRutaDestino = "\\192.168.1.10\e\NET\ANTIBIOGRAMA\" & idficha & ".pdf"
        ElseIf tipoinforme = 5 Then
            sRutaDestino = "\\192.168.1.10\e\NET\PAL\" & idficha & ".pdf"
        ElseIf tipoinforme = 6 Then
            sRutaDestino = "\\192.168.1.10\e\NET\PARASITOLOGIA\" & idficha & ".pdf"
        ElseIf tipoinforme = 7 Then
            sRutaDestino = "\\192.168.1.10\e\NET\ALIMENTOS\" & idficha & ".pdf"
        ElseIf tipoinforme = 8 Then
            sRutaDestino = "\\192.168.1.10\e\NET\SEROLOGIA\" & idficha & ".pdf"
        ElseIf tipoinforme = 9 Then
            sRutaDestino = "\\192.168.1.10\e\NET\PATOLOGIA\" & idficha & ".pdf"
        ElseIf tipoinforme = 10 Then
            sRutaDestino = "\\192.168.1.10\e\NET\CALIDAD\" & idficha & ".pdf"
        ElseIf tipoinforme = 11 Then
            sRutaDestino = "\\192.168.1.10\e\NET\AMBIENTAL\" & idficha & ".pdf"
        ElseIf tipoinforme = 12 Then
            sRutaDestino = "\\192.168.1.10\e\NET\LACTOMETROS\" & idficha & ".pdf"
        ElseIf tipoinforme = 13 Then
            sRutaDestino = "\\192.168.1.10\e\NET\NUTRICION\" & idficha & ".pdf"
        ElseIf tipoinforme = 14 Then
            sRutaDestino = "\\192.168.1.10\e\NET\SUELOS\" & idficha & ".pdf"
        ElseIf tipoinforme = 15 Then
            sRutaDestino = "\\192.168.1.10\e\NET\Bruselocis en leche\" & idficha & ".pdf"
        ElseIf tipoinforme = 16 Then
            sRutaDestino = "\\192.168.1.10\e\NET\Efluentes\" & idficha & ".pdf"
        ElseIf tipoinforme = 17 Then
            sRutaDestino = "\\192.168.1.10\e\NET\BACTEREOLOGIATANQUE\" & idficha & ".pdf"
        ElseIf tipoinforme = 18 Then
            sRutaDestino = "\\192.168.1.10\e\NET\BACTEREOLOGIACLINICAAEROBICA\" & idficha & ".pdf"
        ElseIf tipoinforme = 19 Then
            sRutaDestino = "\\192.168.1.10\e\NET\FOLIARES\" & idficha & ".pdf"
        ElseIf tipoinforme = 20 Then
            sRutaDestino = "\\192.168.1.10\e\NET\TOXICOLOGIA\" & idficha & ".pdf"
        ElseIf tipoinforme = 99 Then
            sRutaDestino = "\\192.168.1.10\e\NET\OTROS\" & idficha & ".pdf"
        End If
        'Dim sArchivoOrigen As String = "\\192.168.1.10\e\NET\INFORMES PARA SUBIR\" & idficha & ".txt"
        Dim sArchivoOrigen As String = "C:\INFORMES PARA SUBIR\" & idficha & ".pdf"

        Try
            ' Mover el fichero.si existe lo sobreescribe  
            My.Computer.FileSystem.MoveFile(sArchivoOrigen, _
                                            sRutaDestino, _
                                            True)
            'MsgBox("Ok.", MsgBoxStyle.Information, "Mover archivo")
            ' errores  
        Catch ex As Exception
            'MsgBox(ex.Message.ToString, MsgBoxStyle.Critical)
        End Try
    End Sub
    Private Sub moverpdf_otrapc()
        If tipoinforme = 10 Then
            'Dim sArchivoOrigen As String = "\\192.168.1.10\e\NET\INFORMES PARA SUBIR\" & idficha & ".pdf"
            Dim sArchivoOrigen As String = "\\ROBOT\INFORMES PARA SUBIR\" & idficha & ".pdf"
            Dim sRutaDestino As String = "\\192.168.1.10\e\NET\CALIDAD\" & idficha & ".pdf"
            Try
                ' Mover el fichero.si existe lo sobreescribe  
                My.Computer.FileSystem.MoveFile(sArchivoOrigen, _
                                                sRutaDestino, _
                                                True)
                'MsgBox("Ok.", MsgBoxStyle.Information, "Mover archivo")
                ' errores  
            Catch ex As Exception
                'MsgBox(ex.Message.ToString, MsgBoxStyle.Critical)
            End Try
        ElseIf tipoinforme = 1 Then
            'Dim sArchivoOrigen As String = "\\192.168.1.10\e\NET\INFORMES PARA SUBIR\" & idficha & ".pdf"
            Dim sArchivoOrigen As String = "\\ROBOT\INFORMES PARA SUBIR\" & idficha & ".pdf"
            Dim sRutaDestino As String = "\\192.168.1.10\e\NET\CONTROL_LECHERO\" & idficha & ".pdf"
            Try
                ' Mover el fichero.si existe lo sobreescribe  
                My.Computer.FileSystem.MoveFile(sArchivoOrigen, _
                                                sRutaDestino, _
                                                True)
                'MsgBox("Ok.", MsgBoxStyle.Information, "Mover archivo")
                ' errores  
            Catch ex As Exception
                'MsgBox(ex.Message.ToString, MsgBoxStyle.Critical)
            End Try
        ElseIf tipoinforme = 3 Then
            'Dim sArchivoOrigen As String = "\\192.168.1.10\e\NET\INFORMES PARA SUBIR\" & idficha & ".pdf"
            Dim sArchivoOrigen As String = "\\ROBOT\INFORMES PARA SUBIR\" & idficha & ".pdf"
            Dim sRutaDestino As String = "\\192.168.1.10\e\NET\AGUA\" & idficha & ".pdf"
            Try
                ' Mover el fichero.si existe lo sobreescribe  
                My.Computer.FileSystem.MoveFile(sArchivoOrigen, _
                                                sRutaDestino, _
                                                True)
                'MsgBox("Ok.", MsgBoxStyle.Information, "Mover archivo")
                ' errores  
            Catch ex As Exception
                'MsgBox(ex.Message.ToString, MsgBoxStyle.Critical)
            End Try
        ElseIf tipoinforme = 4 Then
            'Dim sArchivoOrigen As String = "\\192.168.1.10\e\NET\INFORMES PARA SUBIR\" & idficha & ".pdf"
            Dim sArchivoOrigen As String = "\\ROBOT\INFORMES PARA SUBIR\" & idficha & ".pdf"
            Dim sRutaDestino As String = "\\192.168.1.10\e\NET\ANTIBIOGRAMA\" & idficha & ".pdf"
            Try
                ' Mover el fichero.si existe lo sobreescribe  
                My.Computer.FileSystem.MoveFile(sArchivoOrigen, _
                                                sRutaDestino, _
                                                True)
                'MsgBox("Ok.", MsgBoxStyle.Information, "Mover archivo")
                ' errores  
            Catch ex As Exception
                'MsgBox(ex.Message.ToString, MsgBoxStyle.Critical)
            End Try
        ElseIf tipoinforme = 7 Then
            'Dim sArchivoOrigen As String = "\\192.168.1.10\e\NET\INFORMES PARA SUBIR\" & idficha & ".pdf"
            Dim sArchivoOrigen As String = "\\ROBOT\INFORMES PARA SUBIR\" & idficha & ".pdf"
            Dim sRutaDestino As String = "\\192.168.1.10\e\NET\ALIMENTOS\" & idficha & ".pdf"
            Try
                ' Mover el fichero.si existe lo sobreescribe  
                My.Computer.FileSystem.MoveFile(sArchivoOrigen, _
                                                sRutaDestino, _
                                                True)
                'MsgBox("Ok.", MsgBoxStyle.Information, "Mover archivo")
                ' errores  
            Catch ex As Exception
                'MsgBox(ex.Message.ToString, MsgBoxStyle.Critical)
            End Try
        ElseIf tipoinforme = 13 Then
            'Dim sArchivoOrigen As String = "\\192.168.1.10\e\NET\INFORMES PARA SUBIR\" & idficha & ".pdf"
            Dim sArchivoOrigen As String = "\\ROBOT\INFORMES PARA SUBIR\" & idficha & ".pdf"
            Dim sRutaDestino As String = "\\192.168.1.10\e\NET\NUTRICION\" & idficha & ".pdf"
            Try
                ' Mover el fichero.si existe lo sobreescribe  
                My.Computer.FileSystem.MoveFile(sArchivoOrigen, _
                                                sRutaDestino, _
                                                True)
                'MsgBox("Ok.", MsgBoxStyle.Information, "Mover archivo")
                ' errores  
            Catch ex As Exception
                'MsgBox(ex.Message.ToString, MsgBoxStyle.Critical)
            End Try
        ElseIf tipoinforme = 14 Then
            'Dim sArchivoOrigen As String = "\\192.168.1.10\e\NET\INFORMES PARA SUBIR\" & idficha & ".pdf"
            Dim sArchivoOrigen As String = "\\ROBOT\INFORMES PARA SUBIR\" & idficha & ".pdf"
            Dim sRutaDestino As String = "\\192.168.1.10\e\NET\Suelos\" & idficha & ".pdf"
            Try
                ' Mover el fichero.si existe lo sobreescribe  
                My.Computer.FileSystem.MoveFile(sArchivoOrigen, _
                                                sRutaDestino, _
                                                True)
                'MsgBox("Ok.", MsgBoxStyle.Information, "Mover archivo")
                ' errores  
            Catch ex As Exception
                'MsgBox(ex.Message.ToString, MsgBoxStyle.Critical)
            End Try
        ElseIf tipoinforme = 15 Then
            'Dim sArchivoOrigen As String = "\\192.168.1.10\e\NET\INFORMES PARA SUBIR\" & idficha & ".pdf"
            Dim sArchivoOrigen As String = "\\ROBOT\INFORMES PARA SUBIR\" & idficha & ".pdf"
            Dim sRutaDestino As String = "\\192.168.1.10\e\NET\Brucelosis en leche\" & idficha & ".pdf"
            Try
                ' Mover el fichero.si existe lo sobreescribe  
                My.Computer.FileSystem.MoveFile(sArchivoOrigen, _
                                                sRutaDestino, _
                                                True)
                'MsgBox("Ok.", MsgBoxStyle.Information, "Mover archivo")
                ' errores  
            Catch ex As Exception
                'MsgBox(ex.Message.ToString, MsgBoxStyle.Critical)
            End Try
        ElseIf tipoinforme = 11 Then
            'Dim sArchivoOrigen As String = "\\192.168.1.10\e\NET\INFORMES PARA SUBIR\" & idficha & ".pdf"
            Dim sArchivoOrigen As String = "\\ROBOT\INFORMES PARA SUBIR\" & idficha & ".pdf"
            Dim sRutaDestino As String = "\\192.168.1.10\e\NET\AMBIENTAL\" & idficha & ".pdf"
            Try
                ' Mover el fichero.si existe lo sobreescribe  
                My.Computer.FileSystem.MoveFile(sArchivoOrigen, _
                                                sRutaDestino, _
                                                True)
                'MsgBox("Ok.", MsgBoxStyle.Information, "Mover archivo")
                ' errores  
            Catch ex As Exception
                'MsgBox(ex.Message.ToString, MsgBoxStyle.Critical)
            End Try
        End If
    End Sub
    Private Sub movertxt()
        Dim sRutaDestino As String
        If tipoinforme = 1 Then
            sRutaDestino = "\\192.168.1.10\e\NET\CONTROL_LECHERO\" & idficha & ".txt"
        ElseIf tipoinforme = 3 Then
            sRutaDestino = "\\192.168.1.10\e\NET\AGUA\" & idficha & ".txt"
        ElseIf tipoinforme = 4 Then
            sRutaDestino = "\\192.168.1.10\e\NET\ANTIBIOGRAMA\" & idficha & ".txt"
        ElseIf tipoinforme = 5 Then
            sRutaDestino = "\\192.168.1.10\e\NET\PAL\" & idficha & ".txt"
        ElseIf tipoinforme = 6 Then
            sRutaDestino = "\\192.168.1.10\e\NET\PARASITOLOGIA\" & idficha & ".txt"
        ElseIf tipoinforme = 7 Then
            sRutaDestino = "\\192.168.1.10\e\NET\ALIMENTOS\" & idficha & ".txt"
        ElseIf tipoinforme = 8 Then
            sRutaDestino = "\\192.168.1.10\e\NET\SEROLOGIA\" & idficha & ".txt"
        ElseIf tipoinforme = 9 Then
            sRutaDestino = "\\192.168.1.10\e\NET\PATOLOGIA\" & idficha & ".txt"
        ElseIf tipoinforme = 10 Then
            sRutaDestino = "\\192.168.1.10\e\NET\CALIDAD\" & idficha & ".txt"
        ElseIf tipoinforme = 11 Then
            sRutaDestino = "\\192.168.1.10\e\NET\AMBIENTAL\" & idficha & ".txt"
        ElseIf tipoinforme = 12 Then
            sRutaDestino = "\\192.168.1.10\e\NET\LACTOMETROS\" & idficha & ".txt"
        ElseIf tipoinforme = 13 Then
            sRutaDestino = "\\192.168.1.10\e\NET\NUTRICION\" & idficha & ".txt"
        ElseIf tipoinforme = 14 Then
            sRutaDestino = "\\192.168.1.10\e\NET\SUELOS\" & idficha & ".txt"
        ElseIf tipoinforme = 15 Then
            sRutaDestino = "\\192.168.1.10\e\NET\Bruselocis en leche\" & idficha & ".txt"
        ElseIf tipoinforme = 16 Then
            sRutaDestino = "\\192.168.1.10\e\NET\Efluentes\" & idficha & ".txt"
        ElseIf tipoinforme = 17 Then
            sRutaDestino = "\\192.168.1.10\e\NET\BACTEREOLOGIATANQUE\" & idficha & ".txt"
        ElseIf tipoinforme = 18 Then
            sRutaDestino = "\\192.168.1.10\e\NET\BACTEREOLOGIACLINICAAEROBICA\" & idficha & ".txt"
        ElseIf tipoinforme = 19 Then
            sRutaDestino = "\\192.168.1.10\e\NET\FOLIARES\" & idficha & ".txt"
        ElseIf tipoinforme = 20 Then
            sRutaDestino = "\\192.168.1.10\e\NET\TOXICOLOGIA\" & idficha & ".txt"
        ElseIf tipoinforme = 99 Then
            sRutaDestino = "\\192.168.1.10\e\NET\OTROS\" & idficha & ".txt"
        End If
        'Dim sArchivoOrigen As String = "\\192.168.1.10\e\NET\INFORMES PARA SUBIR\" & idficha & ".txt"
        Dim sArchivoOrigen As String = "C:\INFORMES PARA SUBIR\" & idficha & ".txt"

        Try
            ' Mover el fichero.si existe lo sobreescribe  
            My.Computer.FileSystem.MoveFile(sArchivoOrigen, _
                                            sRutaDestino, _
                                            True)
            'MsgBox("Ok.", MsgBoxStyle.Information, "Mover archivo")
            ' errores  
        Catch ex As Exception
            'MsgBox(ex.Message.ToString, MsgBoxStyle.Critical)
        End Try

    End Sub
    Private Sub movertxt_otrapc()
        If tipoinforme = 1 Then
            'Dim sArchivoOrigen As String = "\\192.168.1.10\e\NET\INFORMES PARA SUBIR\" & idficha & ".txt"
            Dim sArchivoOrigen As String = "\\ROBOT\INFORMES PARA SUBIR\" & idficha & ".txt"
            Dim sRutaDestino As String = "\\192.168.1.10\e\NET\CONTROL_LECHERO\" & idficha & ".txt"
            Try
                ' Mover el fichero.si existe lo sobreescribe  
                My.Computer.FileSystem.MoveFile(sArchivoOrigen, _
                                                sRutaDestino, _
                                                True)
                'MsgBox("Ok.", MsgBoxStyle.Information, "Mover archivo")
                ' errores  
            Catch ex As Exception
                'MsgBox(ex.Message.ToString, MsgBoxStyle.Critical)
            End Try
        End If
    End Sub
    Private Sub creapreinformesmanual()
        Dim pi As New dPreinformes
        Dim lista As New ArrayList
        Dim creapreinformecalidad As Integer = 1
        lista = pi.listarsinmarcarcalidad
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each pi In lista
                    Dim csm As New dCalidadSolicitudMuestra
                    Dim listacsm As New ArrayList
                    Dim ficha As Long = pi.FICHA
                    listacsm = csm.listarporsolicitud(ficha)
                    If Not listacsm Is Nothing Then
                        creapreinformecalidad = 1
                        For Each csm In listacsm
                            If csm.RB = 1 Then
                                Dim ibc As New dIbc
                                ibc.FICHA = csm.FICHA
                                ibc.MUESTRA = csm.MUESTRA
                                ibc = ibc.buscarxfichaxmuestra
                                If Not ibc Is Nothing Then
                                Else
                                    creapreinformecalidad = 0
                                    Exit For
                                End If
                                ibc = Nothing
                            End If
                            If csm.PSICROTROFOS = 1 Then
                                Dim psi As New dPsicrotrofos
                                psi.FICHA = csm.FICHA
                                psi.MUESTRA = csm.MUESTRA
                                psi = psi.buscarxfichaxmuestra
                                If Not psi Is Nothing Then
                                Else
                                    creapreinformecalidad = 0
                                    Exit For
                                End If
                                psi = Nothing
                            End If
                            If csm.ESPORULADOS = 1 Then
                                Dim esp As New dEsporulados
                                esp.FICHA = csm.FICHA
                                esp.MUESTRA = csm.MUESTRA
                                esp = esp.buscarxfichaxmuestra
                                If Not esp Is Nothing Then
                                Else
                                    creapreinformecalidad = 0
                                    Exit For
                                End If
                                esp = Nothing
                            End If
                            If csm.INHIBIDORES = 1 Then
                                Dim inh As New dInhibidores
                                inh.FICHA = csm.FICHA
                                inh.MUESTRA = csm.MUESTRA
                                inh = inh.buscarxfichaxmuestra
                                If Not inh Is Nothing Then
                                    If inh.MARCA = 0 Then
                                        creapreinformecalidad = 0
                                        Exit For
                                    End If
                                Else
                                    creapreinformecalidad = 0
                                    Exit For
                                End If
                                inh = Nothing
                            End If
                            If csm.AFLATOXINA = 1 Then
                                Dim ml As New dMicotoxinasLeche
                                ml.FICHA = csm.FICHA
                                ml.MUESTRA = csm.MUESTRA
                                ml = ml.buscarxfichaxmuestra
                                If Not ml Is Nothing Then
                                Else
                                    creapreinformecalidad = 0
                                    Exit For
                                End If
                                ml = Nothing
                            End If
                        Next
                    End If
                    If creapreinformecalidad = 1 Then
                        preinforme_calidad(ficha)
                    End If
                Next
            End If
        End If
        pre_informe_control()
    End Sub
    Private Sub soloibc()
        Dim extension As String
        Dim nombrearchivo As String = ""
        Dim linea As Integer
        '**********************************************************************************
        Dim El_Ping As Boolean
        Dim eco As New System.Net.NetworkInformation.Ping
        Dim res As System.Net.NetworkInformation.PingReply
        Dim ip As Net.IPAddress
        ip = Net.IPAddress.Parse("192.168.1.50")
        res = eco.Send(ip)
        If res.Status = System.Net.NetworkInformation.IPStatus.Success Then
            El_Ping = (My.Computer.Network.Ping("ibc1123"))
        End If
        'Acá mandamos los mensajes para las 2 posibilidades
        If El_Ping = False Then
            'si no se pudo acceder ,avisamos
            'MsgBox("El servidor no está disponible.", MsgBoxStyle.Critical, "Error")
        Else
            'MsgBox("Servidor disponible.", MsgBoxStyle.Information, "Aviso")
            Dim folder As New DirectoryInfo("\\Ibc1123\Carol")
            For Each file As FileInfo In folder.GetFiles("*.csv")
                'ListBox1.Items.Add(file.Name)
                nombrearchivo = file.Name
                linea = 1
                extension = Microsoft.VisualBasic.Right(file.Name, 3)
                Dim objReader2 As New StreamReader("\\Ibc1123\Carol\" & file.Name)
                Dim sLine As String = ""
                Dim arraytext() As String
                Dim ficha As String = ""
                Dim ficha2 As String = ""
                Dim ficha3 As String
                Dim muestra As String = ""
                Dim idibc As Integer = 0
                Dim ibc As Long = 0
                Dim rb As Integer = 0
                ' *** SI EL ARCHIVO ES CSV **************************************************************************************
                If extension = "csv" Or extension = "CSV" Then
                    Dim c As New dImpIbc()
                    Do
                        sLine = objReader2.ReadLine()
                        If Not sLine Is Nothing Then
                            arraytext = Split(sLine, ",")
                            Dim muestra2 As String
                            Dim muestrax As String
                            If Trim(arraytext(1)) <> "" Then
                                muestra = arraytext(1)
                                muestrax = Replace(muestra, Chr(34), "")
                                If muestrax <> "" Then
                                    muestra2 = muestrax
                                Else
                                    muestra = arraytext(7)
                                    muestrax = Replace(muestra, Chr(34), "")
                                    If muestrax <> "" Then
                                        muestra2 = muestrax
                                    Else
                                        muestra2 = "error"
                                    End If
                                End If
                            Else
                                If arraytext.Length > 7 Then
                                    muestra = arraytext(7)
                                    muestrax = Replace(muestra, Chr(34), "")
                                    If muestrax <> "" Then
                                        muestra2 = muestrax
                                    Else
                                        muestra2 = "error"
                                    End If
                                Else
                                    muestra2 = "error"
                                End If
                            End If
                            If Trim(arraytext(2)) <> "" Then
                                idibc = arraytext(2)
                            Else
                                idibc = -1
                            End If
                            If Trim(arraytext(4)) <> "" Then
                                ibc = arraytext(4)
                            Else
                                ibc = -1
                            End If
                            If Trim(arraytext(5)) <> "" Then
                                rb = arraytext(5)
                            Else
                                rb = -1
                            End If
                            ficha2 = Mid(file.Name, Len(file.Name) - 4, 1)
                            ficha3 = Mid(file.Name, 1, 1)
                            If ficha2 = "a" Or ficha2 = "A" Or ficha2 = "b" Or ficha2 = "B" Or ficha2 = "c" Or ficha2 = "C" Or ficha2 = "d" Or ficha2 = "D" Then
                                ficha = Mid(file.Name, 1, Len(file.Name) - 5)
                            Else
                                ficha = Mid(file.Name, 1, Len(file.Name) - 4)
                            End If
                            If Mid(ficha, 1, 1) = "l" Or Mid(ficha, 1, 1) = "L" Then
                                Dim MyString As String = ficha
                                Dim MyChar As Char() = {"l"c, "L"c}
                                Dim NewString As String = MyString.TrimStart(MyChar)
                                ficha3 = NewString
                            Else
                                ficha3 = ficha
                            End If
                            Dim fechaoriginal As Date = Now()
                            Dim fecha As String
                            fecha = Format(fechaoriginal, "yyyy-MM-dd")
                            c.FICHA = ficha3
                            c.MUESTRA = muestra2
                            c.IDIBC = idibc
                            c.IBC = ibc
                            c.RB = rb
                            c.FECHA = fecha
                            c.guardar()
                        End If
                        linea = linea + 1
                    Loop Until sLine Is Nothing
                    objReader2.Close()
                End If
                '*** MOVER ARCHIVO ***********************************************************************
                Dim sArchivoOrigen As String = "\\Ibc1123\Carol\" & nombrearchivo
                'Dim sRutaDestino1 As String = "d:\documentos\secretaria\analisis\leche\ibc\" & nombrearchivo
                Dim sRutaDestino1 As String = "Y:\documentos\secretaria\analisis\leche\ibc\" & nombrearchivo
                Dim sRutaDestino As String = "\\Ibc1123\Carol\pasados\" & nombrearchivo
                Try
                    ' Mover el fichero.si existe lo sobreescribe  
                    My.Computer.FileSystem.CopyFile(sArchivoOrigen, _
                                                    sRutaDestino1, _
                                                    True)

                    My.Computer.FileSystem.MoveFile(sArchivoOrigen, _
                                                    sRutaDestino, _
                                                    True)
                    'MsgBox("Ok.", MsgBoxStyle.Information, "Mover archivo")
                    ' errores  
                Catch ex As Exception
                    MsgBox(ex.Message.ToString, MsgBoxStyle.Critical)
                End Try
            Next
        End If
        '***********************************************************************************
    End Sub
    Private Sub buscarinformes()
        Dim _hora As String = ""
        Dim hora As Integer = 0
        Dim minuto As Integer = 0
        _hora = Now.ToString("HH:mm")
        hora = Mid(_hora, 1, 2)
        minuto = Mid(_hora, 4, 2)
        If hora = 8 And minuto < 30 Then
            buscarinformesfq()
        End If
    End Sub
    Private Sub buscarinformesfq()
        Dim pi As New dPreinformes
        Dim fechadesde As Date = Now
        Dim fechahasta As Date = Now
        Dim fechad As String
        Dim fechah As String
        fechad = Format(fechadesde, "yyyy-MM-dd")
        fechah = Format(fechahasta, "yyyy-MM-dd")
        Dim lcontrol As New ArrayList
        Dim lcalidad As New ArrayList
        Dim ficha As Long = 0
        Dim fecha As Date = Now
        Dim tipo As Integer = 0
        Dim resultado As Integer = 0
        Dim coincide As Integer = 0
        Dim observaciones As String = ""
        Dim controlador As Integer = 0
        Dim controlado As Integer = 0
        lcontrol = pi.listarsincontrolclechero(fechad, fechah)
        lcalidad = pi.listarsincontrolcalidad(fechad, fechah)
        '*** CONTROL ***********************************************************************
        If Not lcontrol Is Nothing Then
            If lcontrol.Count > 0 Then
                Dim ci As New dControlInformesFQ
                For Each pi In lcontrol
                    ci.FECHACONTROL = fechad
                    ci.FICHA = pi.FICHA
                    ci.FECHA = fechad
                    ci.TIPO = pi.TIPO
                    ci.RESULTADO = 0
                    ci.COINCIDE = 0
                    ci.OBSERVACIONES = ""
                    ci.CONTROLADOR = 100
                    ci.CONTROLADO = 0
                    ci.guardar()
                Next
            End If
        End If
        '*** CALIDAD ***********************************************************************
        If Not lcalidad Is Nothing Then
            If lcalidad.Count > 0 Then
                Dim ci As New dControlInformesFQ
                For Each pi In lcalidad
                    ci.FECHACONTROL = fechad
                    ci.FICHA = pi.FICHA
                    ci.FECHA = fechad
                    ci.TIPO = pi.TIPO
                    ci.RESULTADO = 0
                    ci.COINCIDE = 0
                    ci.OBSERVACIONES = ""
                    ci.CONTROLADOR = 100
                    ci.CONTROLADO = 0
                    ci.guardar()
                Next
            End If
        End If
    End Sub
    Private Sub buscarinformesmicro()
        Dim pi As New dPreinformes
        Dim fechadesde As Date = Now
        Dim fechahasta As Date = Now
        Dim fechad As String
        Dim fechah As String
        fechad = Format(fechadesde, "yyyy-MM-dd")
        fechah = Format(fechahasta, "yyyy-MM-dd")
        Dim lcalidad As New ArrayList
        Dim lagua As New ArrayList
        Dim lsubproductos As New ArrayList
        Dim ficha As Long = 0
        Dim fecha As Date = Now
        Dim tipo As Integer = 0
        Dim resultado As Integer = 0
        Dim coincide As Integer = 0
        Dim observaciones As String = ""
        Dim controlador As Integer = 0
        Dim controlado As Integer = 0
        lcalidad = pi.listarsincontrolcalidad(fechad, fechah)
        lagua = pi.listarsincontrolagua(fechad, fechah)
        lsubproductos = pi.listarsincontrolsubproductos(fechad, fechah)
        '*** CALIDAD ***********************************************************************
        If Not lcalidad Is Nothing Then
            If lcalidad.Count > 0 Then
                Dim ci As New dControlInformesFQ
                For Each pi In lcalidad
                    ci.FECHACONTROL = fechad
                    ci.FICHA = pi.FICHA
                    ci.FECHA = fechad
                    ci.TIPO = pi.TIPO
                    ci.RESULTADO = 0
                    ci.COINCIDE = 0
                    ci.OBSERVACIONES = ""
                    ci.CONTROLADOR = 100
                    ci.CONTROLADO = 0
                    ci.guardar()
                Next
            End If
        End If
        '*** AGUA ***********************************************************************
        If Not lagua Is Nothing Then
            If lagua.Count > 0 Then
                Dim ci As New dControlInformesMicro
                For Each pi In lagua
                    ci.FECHACONTROL = fechad
                    ci.FICHA = pi.FICHA
                    ci.FECHA = fechad
                    ci.TIPO = pi.TIPO
                    ci.RESULTADO = 0
                    ci.COINCIDE = 0
                    ci.OBSERVACIONES = ""
                    ci.CONTROLADOR = 100
                    ci.CONTROLADO = 0
                    ci.guardar()
                Next
            End If
        End If
        '*** ALIMENTOS ***********************************************************************
        If Not lsubproductos Is Nothing Then
            If lsubproductos.Count > 0 Then
                Dim ci As New dControlInformesMicro
                For Each pi In lsubproductos
                    ci.FECHACONTROL = fechad
                    ci.FICHA = pi.FICHA
                    ci.FECHA = fechad
                    ci.TIPO = pi.TIPO
                    ci.RESULTADO = 0
                    ci.COINCIDE = 0
                    ci.OBSERVACIONES = ""
                    ci.CONTROLADOR = 100
                    ci.CONTROLADO = 0
                    ci.guardar()
                Next
            End If
        End If
    End Sub
    Private Sub marcarenvio()
        Dim c As New dCompras
        c.ID = compraid
        c.marcarenviado()
    End Sub
    Private Sub enviaravisoenviocajas()
        Dim p As New dPedidos
        Dim listap As New ArrayList
        Dim flag As Integer = 1
        listap = p.listarsinavisar
        If Not listap Is Nothing Then
            For Each p In listap
                flag = 1
                Dim env As New dEnvioCajas
                Dim listae As New ArrayList
                listae = env.listarcargada(p.ID)
                If Not listae Is Nothing Then
                    For Each env In listae
                        If env.CARGADA = 0 Then
                            flag = 0
                        End If
                    Next
                    If flag = 1 Then
                        enviomailcajas(p.ID)
                        enviar_notificacion_envio(p.ID)
                        p.marcar(p.ID)
                    End If
                End If
            Next
        End If

    End Sub
    Private Sub enviar_notificacion_envio(ByVal _id As Long)
        Dim p As New dPedidos
        p.ID = _id
        p = p.buscar
        Dim eagencia As String = ""
        Dim eremito As String = ""
        Dim fechaactual As Date = Now()
        Dim _fecha As String
        _fecha = Format(fechaactual, "yyyy-MM-dd")
        Dim notificacion As New Dictionary(Of String, dNotificaciones)
        Dim nt As New dNotificaciones
        Dim _tipo As String = ""
        Dim _mensaje As String = ""
        Dim nuevoid As Long = p.IDPRODUCTOR 'CType(TextIdProductor.Text, Long)
        Dim _detalle As String = ""
        Dim _detalle_envio As String = ""
        Dim ag As New dEmpresaT
        ag.ID = p.IDAGENCIA
        ag = ag.buscar
        eagencia = ag.NOMBRE
        'Dim eagencia As String = ComboAgencia.Text
        If eagencia = "RETIRA EN COLAVECO" Then
            _mensaje = "Su pedido de frascos está pronto para retirar en Colaveco"
        ElseIf eagencia = "Retira ahora" Then
            _mensaje = "Su pedido de frascos está pronto para retirar en Colaveco"
        Else
            _mensaje = "Su pedido de frascos fué despachado por " & eagencia
        End If
        _tipo = "envio_frasco"
        Dim env As New dEnvioCajas
        Dim idped As Long = _id 'CType(TextId.Text, Long)
        Dim lista As New ArrayList
        lista = env.listarporid(idped)
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each env In lista
                    _detalle_envio = _detalle_envio & env.IDCAJA & ", "
                Next
            End If
        End If
        '_detalle = "<p><b>Fecha de despacho:</b> " + _fecha + " </p><p><b>Agencia:</b> " + eagencia + " </p><p><b>Destino:</b> " + TextDireccion.Text + " </p><p><b>Detalle de Envio:</b> " + _detalle_envio + " </p>"
        _detalle = "<p><b>Fecha de despacho:</b> " + _fecha + " </p><p><b>Agencia:</b> " + eagencia + " </p><p><b>Destino:</b> " + p.DIRECCION + " </p><p><b>Detalle de Envio:</b> " + _detalle_envio + " </p>"
        nt.fecha = _fecha
        nt.tipo = _tipo
        nt.mensaje = _mensaje
        nt.idnet_usuario = nuevoid
        nt.detalle = _detalle
        notificacion.Add("notification", nt)
        Dim parameters As String = JsonConvert.SerializeObject(notificacion, Formatting.None)
        Dim status As HttpStatusCode = HttpStatusCode.ExpectationFailed
        Dim response As Byte() = PostResponse("http://colaveco-gestor.herokuapp.com/notifications", "POST", parameters, status)
    End Sub
    Private Sub solicitud_frascos()
        Dim status As HttpStatusCode = HttpStatusCode.ExpectationFailed
        Dim response As Byte() = GetResponse("http://colaveco-gestor.herokuapp.com/solicitudfrascos/nuevos", "GET", status)
        Dim responseString As String
        If response IsNot Nothing Then
            responseString = System.Text.Encoding.UTF8.GetString(response)
        Else
            responseString = "NULL"
        End If
        Console.WriteLine("Response Code: " & status)
        Console.WriteLine("Response String: " & responseString)
        Dim frascos As New List(Of dFrascosGestor)
        frascos = JsonConvert.DeserializeObject(Of List(Of dFrascosGestor))(responseString)
        For Each a In frascos
            Dim pw As New dPedidosWeb
            Dim fecha As Date = Now
            Dim fec As String
            fec = Format(fecha, "yyyy-MM-dd")
            pw.FECHA = fec
            pw.CODIGO = a.idnet
            pw.NOMBRE = a.nombre
            pw.DIRECCION = a.direccion
            If a.agencia = "Agencia Central" Then
                pw.AGENCIA = 1
            ElseIf a.agencia = "Tiempost" Then
                pw.AGENCIA = 2
            ElseIf a.agencia = "Cia. Colonia" Then
                pw.AGENCIA = 3
            ElseIf a.agencia = "Cot" Then
                pw.AGENCIA = 4
            ElseIf a.agencia = "Comsa" Then
                pw.AGENCIA = 5
            ElseIf a.agencia = "Turil" Then
                pw.AGENCIA = 6
            ElseIf a.agencia = "Retiro en Colaveco" Then
                pw.AGENCIA = 7
            ElseIf a.agencia = "No proporcionado" Then
                pw.AGENCIA = 8
            ElseIf a.agencia = "Correo" Then
                pw.AGENCIA = 9
            ElseIf a.agencia = "Retiro en Florida" Then
                pw.AGENCIA = 10
            ElseIf a.agencia = "Retiro en Cardal" Then
                pw.AGENCIA = 11
            ElseIf a.agencia = "Retiro en Canelones" Then
                pw.AGENCIA = 12
            ElseIf a.agencia = "Retiro ahora" Then
                pw.AGENCIA = 13
            End If
            pw.TELEFONO = a.telefono
            pw.EMAIL = a.email
            pw.CCONSERVANTE = a.frascos_con_c
            pw.SCONSERVANTE = a.frascos_sin_c
            pw.AGUA = a.frascos_agua
            pw.SANGRE = a.frascos_sangre
            pw.OBSERVACIONES = a.observaciones
            pw.REALIZADO = 0
            pw.CANCELADO = 0
            pw.guardar()
            pw = Nothing
        Next
    End Sub
    Private Sub actualizacion_clientes()
        Dim status As HttpStatusCode = HttpStatusCode.ExpectationFailed
        Dim response As Byte() = GetResponse("http://colaveco-gestor.herokuapp.com/users/actualizados", "GET", status)
        Dim responseString As String
        If response IsNot Nothing Then
            responseString = System.Text.Encoding.UTF8.GetString(response)
        Else
            responseString = "NULL"
        End If
        Console.WriteLine("Response Code: " & status)
        Console.WriteLine("Response String: " & responseString)
        Dim usuario As New List(Of dUsuarioGestor)
        usuario = JsonConvert.DeserializeObject(Of List(Of dUsuarioGestor))(responseString)
        For Each a In usuario
            Dim c As New dCliente
            c.ID = a.idnet
            c.NOMBRE = a.nombre
            c.EMAIL = a.email
            c.NOMBRE_EMAIL1 = a.tecnico_email_nombre_1
            c.EMAIL1 = a.tecnico_email_1
            c.NOMBRE_EMAIL2 = a.tecnico_email_nombre_2
            c.EMAIL2 = a.tecnico_email_2
            c.ENVIO = a.direccion_frasco
            c.USUARIO_WEB = a.usuario_web
            c.NOMBRE_CELULAR1 = a.tecnico_celular_nombre_1
            c.CELULAR = a.tecnico_celular_1
            c.NOMBRE_CELULAR2 = a.tecnico_celular_nombre_2
            c.CELULAR2 = a.tecnico_celular_2
            c.DIRECCION = a.direccion
            c.NOMBRE_TELEFONO1 = a.tecnico_telefono_nombre_1
            c.TELEFONO1 = a.tecnico_telefono_1
            c.NOMBRE_TELEFONO2 = a.tecnico_telefono_nombre_2
            c.TELEFONO2 = a.tecnico_telefono_2
            c.DICOSE = a.dicose
            c.IDAGENCIA = a.agencia_frasco
            c.FAC_RSOCIAL = a.razon_social
            c.FAC_CEDULA = a.documento
            c.FAC_RUT = a.rut
            c.FAC_DIRECCION = a.fac_direccion
            c.FAC_LOCALIDAD = a.fac_localidad
            c.FAC_DEPARTAMENTO = a.fac_departamento
            c.COB_NOMBRE_TELEFONO1 = a.cobranza_telefono_nombre_1
            c.FAC_TELEFONOS = a.cobranza_telefono_1
            c.COB_NOMBRE_TELEFONO2 = a.cobranza_telefono_nombre_2
            c.COB_TELEFONO2 = a.cobranza_telefono_2
            c.COB_NOMBRE_CELULAR1 = a.cobranza_celular_nombre_1
            c.COB_CELULAR1 = a.cobranza_celular_1
            c.COB_NOMBRE_CELULAR2 = a.cobranza_celular_nombre_2
            c.COB_CELULAR2 = a.cobranza_celular_2
            c.COB_NOMBRE_EMAIL1 = a.cobranza_email_nombre_1
            c.COB_EMAIL1 = a.cobranza_email_1
            c.COB_NOMBRE_EMAIL2 = a.cobranza_email_nombre_2
            c.COB_EMAIL2 = a.cobranza_email_2
            c.FAC_EMAIL = a.fac_email_envio
            c.NOT_EMAIL_FRASCOS1 = a.notificacion_frasco_1
            c.NOT_EMAIL_FRASCOS2 = a.notificacion_frasco_2
            c.NOT_EMAIL_MUESTRAS1 = a.notificacion_solicitud_1
            c.NOT_EMAIL_MUESTRAS2 = a.notificacion_solicitud_2
            c.NOT_EMAIL_ANALISIS1 = a.notificacion_resultado_1
            c.NOT_EMAIL_ANALISIS2 = a.notificacion_resultado_2
            c.NOT_EMAIL_GENERAL1 = a.notificacion_avisos_1
            c.NOT_EMAIL_GENERAL2 = a.notificacion_avisos_2
            c.modificardesdeweb()
            c = Nothing
        Next
    End Sub
    Private Sub cargarpedidosweb()
        Dim pw_com As New dPedidosWeb_com
        Dim fecha As Date = Now
        Dim fec As String
        fec = Format(fecha, "yyyy-MM-dd")
        Dim lista As New ArrayList
        lista = pw_com.listar
        If Not lista Is Nothing Then
            For Each pw_com In lista
                Dim pw As New dPedidosWeb
                pw.FECHA = fec
                pw.CODIGO = pw_com.CODIGO
                pw.NOMBRE = pw_com.NOMBRE
                pw.DIRECCION = pw_com.DIRECCION
                pw.AGENCIA = pw_com.AGENCIA
                pw.TELEFONO = pw_com.TELEFONO
                pw.EMAIL = pw_com.EMAIL
                pw.CCONSERVANTE = pw_com.CCONSERVANTE
                pw.SCONSERVANTE = pw_com.SCONSERVANTE
                pw.AGUA = pw_com.AGUA
                pw.SANGRE = pw_com.SANGRE
                pw.OBSERVACIONES = pw_com.OBSERVACIONES
                pw.REALIZADO = 0
                pw.CANCELADO = 0
                pw.guardar()
                pw_com.eliminar()
                pw = Nothing
            Next
        End If
    End Sub
    Private Sub moverarchivossubidos()
        Dim extension As String
        Dim nombrearchivo As String = ""
        Dim numficha As Long = 0
        'Dim folder As New DirectoryInfo("\\192.168.1.10\e\NET\Informes para subir")
        Dim folder As New DirectoryInfo("C:\INFORMES PARA SUBIR")
        Dim _ficheros() As String
        '_ficheros = Directory.GetFiles("\\192.168.1.10\e\NET\Informes para subir")
        _ficheros = Directory.GetFiles("C:\INFORMES PARA SUBIR")
        If Not (_ficheros.Length > 0) Then
        Else
            For Each file As FileInfo In folder.GetFiles("*.*")
                nombrearchivo = file.Name
                extension = Microsoft.VisualBasic.Right(file.Name, 3)
                If extension = "xls" Or extension = "pdf" Or extension = "txt" Or extension = "XLS" Or extension = "PDF" Or extension = "TXT" Then
                    numficha = Mid(file.Name, 1, Len(file.Name) - 4)
                    idficha = numficha
                    Dim sa As New dSolicitudAnalisis
                    sa.ID = numficha
                    sa = sa.buscar
                    If Not sa Is Nothing Then
                        tipoinforme = sa.IDTIPOINFORME
                        If sa.MARCA = 1 Then
                            If extension = "xls" Or extension = "XLS" Then
                                moverexcel()
                            End If
                            If extension = "pdf" Or extension = "PDF" Then
                                moverpdf()
                            End If
                            If extension = "txt" Or extension = "TXT" Then
                                movertxt()
                            End If
                        End If
                    End If
                    sa = Nothing
                End If
            Next
        End If
    End Sub
    Private Sub controles_ibc()
        Dim hoy As Date = Now
        Dim ano As Integer = 0
        Dim mes As Integer = 0
        Dim mes_ As String = ""
        ano = hoy.Year
        mes = hoy.Month
        Dim _mes As New dMeses
        _mes.ID = mes
        _mes = _mes.buscar
        If Not _mes Is Nothing Then
            mes_ = _mes.NOMBRE
        End If
    End Sub
    Private Sub agua(ByVal ano As Integer, ByVal mes As String)
        Dim extension As String
        Dim nombrearchivo As String = ""
        Dim linea As Integer
        Dim ficha As String = ""
        Dim ficha2 As String = ""
        Dim ficha3 As String = "0"
        Dim fecha2 As String = ""
        '**********************************************************************************
        Dim El_Ping As Boolean
        Dim eco As New System.Net.NetworkInformation.Ping
        Dim res As System.Net.NetworkInformation.PingReply
        Dim ip As Net.IPAddress
        ip = Net.IPAddress.Parse("192.168.1.50")
        res = eco.Send(ip)
        If res.Status = System.Net.NetworkInformation.IPStatus.Success Then
            El_Ping = (My.Computer.Network.Ping("ibc1123"))
        End If
        If El_Ping = False Then
        Else
            Dim folder As New DirectoryInfo("\\Ibc1123\Agua\" & ano & "\" & mes)
            For Each file As FileInfo In folder.GetFiles("*.csv")
                nombrearchivo = file.Name
                linea = 1
                extension = Microsoft.VisualBasic.Right(file.Name, 3)
                Dim objReader2 As New StreamReader("\\Ibc1123\Agua\" & ano & "\" & mes & "\" & file.Name)
                Dim sLine As String = ""
                Dim arraytext() As String
                Dim muestra As String = ""
                Dim idibc As Integer = 0
                Dim ibc As Long = 0
                Dim rb As Integer = 0
                If extension = "csv" Or extension = "CSV" Then
                    Do
                        sLine = objReader2.ReadLine()
                        If Not sLine Is Nothing Then
                            arraytext = Split(sLine, ",")
                            Dim muestra2 As String
                            Dim muestrax As String
                            If Trim(arraytext(1)) <> "" Then
                                muestra = arraytext(1)
                                muestrax = Replace(muestra, Chr(34), "")
                                If muestrax <> "" Then
                                    muestra2 = muestrax
                                Else
                                    muestra = arraytext(7)
                                    muestrax = Replace(muestra, Chr(34), "")
                                    If muestrax <> "" Then
                                        muestra2 = muestrax
                                    Else
                                        muestra2 = "error"
                                    End If
                                End If
                            Else
                                If arraytext.Length > 7 Then
                                    muestra = arraytext(7)
                                    muestrax = Replace(muestra, Chr(34), "")
                                    If muestrax <> "" Then
                                        muestra2 = muestrax
                                    Else
                                        muestra2 = "error"
                                    End If
                                Else
                                    muestra2 = "error"
                                End If
                            End If
                            If Trim(arraytext(2)) <> "" Then
                                idibc = arraytext(2)
                            Else
                                idibc = -1
                            End If
                            If Trim(arraytext(4)) <> "" Then
                                ibc = arraytext(4)
                            Else
                                ibc = -1
                            End If
                            If Trim(arraytext(5)) <> "" Then
                                rb = arraytext(5)
                            Else
                                rb = -1
                            End If
                            ficha2 = Mid(file.Name, Len(file.Name) - 4, 1)
                            ficha3 = Mid(file.Name, 1, 1)
                            If ficha2 = "a" Or ficha2 = "A" Or ficha2 = "b" Or ficha2 = "B" Or ficha2 = "c" Or ficha2 = "C" Or ficha2 = "d" Or ficha2 = "D" Then
                                ficha = Mid(file.Name, 1, Len(file.Name) - 5)
                            Else
                                ficha = Mid(file.Name, 1, Len(file.Name) - 4)
                            End If
                            If Mid(ficha, 1, 1) = "l" Or Mid(ficha, 1, 1) = "L" Then
                                Dim MyString As String = ficha
                                Dim MyChar As Char() = {"l"c, "L"c}
                                Dim NewString As String = MyString.TrimStart(MyChar)
                                ficha3 = NewString
                            Else
                                ficha3 = ficha
                            End If
                            Dim fechaoriginal As Date = Now()
                            Dim fecha As String
                            fecha = Format(fechaoriginal, "yyyy-MM-dd")
                            fecha2 = Format(fechaoriginal, "yyyy-MM-dd")
                        End If
                        linea = linea + 1
                    Loop Until sLine Is Nothing
                    objReader2.Close()
                End If
            Next
        End If
    End Sub
    Private Sub cargarsolicitudesweb()

        Dim sw As New dSolicitudWeb
        Dim lista As New ArrayList
        lista = sw.listarsincargar
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each sw In lista
                    barra = barra + 1
                    ProgressBar1.Value = barra
                    If barra = 100 Then
                        barra = 0
                    End If
                    Dim s As New dSolicitudAnalisis
                    s.ID = sw.FICHA
                    id_ficha_ = s.ID
                    s = s.buscar
                    nficha = sw.FICHA
                    If Not s Is Nothing Then
                        nmuestrasweb = s.NMUESTRAS
                        Dim ti As New dTipoInforme
                        ti.ID = s.IDTIPOINFORME
                        ti = ti.buscar
                        If Not ti Is Nothing Then
                            tipoinformeweb = ti.NOMBRE
                        End If
                        Dim si As New dSubInforme
                        si.ID = s.IDSUBINFORME
                        si = si.buscar
                        If Not si Is Nothing Then
                            subtipoinformeweb = si.NOMBRE
                        End If
                        observacionesweb = s.OBSERVACIONES
                        Dim m As New dMuestras
                        m.ID = s.IDMUESTRA
                        m = m.buscar
                        If Not m Is Nothing Then
                            muestraweb = m.NOMBRE
                        End If
                        Dim c As New dCliente
                        c.ID = s.IDPRODUCTOR
                        c = c.buscar
                        If Not c Is Nothing Then
                            email = ""
                            email2 = ""
                            If c.NOT_EMAIL_MUESTRAS1 <> "" Then
                                email = RTrim(c.NOT_EMAIL_MUESTRAS1)
                            End If
                            If c.NOT_EMAIL_MUESTRAS2 <> "" Then
                                email2 = RTrim(c.NOT_EMAIL_MUESTRAS2)
                            End If
                            If email = "" Then
                                If email2 = "" Then
                                    If c.EMAIL <> "" Then
                                        email = RTrim(c.EMAIL)
                                    End If
                                Else
                                    email = email2
                                End If
                            Else
                                If email2 = "" Then
                                    email = email
                                Else
                                    email = email & "," & email2
                                End If
                            End If
                            nombreproductorweb = c.NOMBRE
                            productorweb_com = c.USUARIO_WEB
                            Dim pw_com As New dProductorWeb_com
                            pw_com.USUARIO = productorweb_com
                            pw_com = pw_com.buscar
                            If Not pw_com Is Nothing Then
                                idproductorweb_com = pw_com.ID
                                'email = RTrim(pw_com.ENVIAR_EMAIL)
                                celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                InsertarRegistro_com()
                                sw.marcarcargado()
                                nficha = 0
                                email = ""
                                email2 = ""
                            End If
                        End If
                    End If
                Next
            End If
        End If
    End Sub
    Private Sub InsertarRegistro_com()
        Dim idnet As Long = 0
        idficha = nficha
        Dim sa_ As New dSolicitudAnalisis
        sa_.ID = idficha
        sa_ = sa_.buscar
        If Not sa_ Is Nothing Then
            idnet = sa_.IDPRODUCTOR
        End If
        Dim tipoinformesw As String
        tipoinformesw = tipoinformeweb
        If tipoinformesw = "Control Lechero" Then 'SI EL TIPO DE INFORME ES DE CONTROL LECHERO
            tipoinforme = 1
            Dim cw_com As New dControlLecheroWeb_com
            Dim pw_com As New dProductorWeb_com
            pw_com.USUARIO = productorweb_com
            pw_com = pw_com.buscar
            Dim idproductorweb_com As Long = pw_com.ID
            Dim fecha_emision As Date = DateFecha.Value.ToString("yyyy-MM-dd")
            Dim fechaemi As String
            fechaemi = Format(fecha_emision, "yyyy-MM-dd")
            cw_com.ID_USUARIO = idproductorweb_com
            cw_com.ABONADO = 0
            cw_com.FECHA_CREADO = fechaemi
            cw_com.FECHA_EMISION = fechaemi
            cw_com.FICHA = idficha
            cw_com.ID_ESTADO = 1
            cw_com.ID_LIBRO = idficha
            If (cw_com.guardar()) Then
            Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
            End If
        ElseIf tipoinformesw = "Calidad de leche" Then 'SI EL TIPO DE INFORME ES DE CALIDAD DE LECHE
            tipoinforme = 10
            Dim cw_com As New dCalidadWeb_com
            Dim pw_com As New dProductorWeb_com
            pw_com.USUARIO = productorweb_com
            pw_com = pw_com.buscar
            Dim idproductorweb_com As Long = pw_com.ID
            Dim fecha_emision As Date = DateFecha.Value.ToString("yyyy-MM-dd")
            Dim fechaemi As String
            fechaemi = Format(fecha_emision, "yyyy-MM-dd")
            cw_com.ID_USUARIO = idproductorweb_com
            cw_com.ABONADO = 0
            cw_com.FECHA_CREADO = fechaemi
            cw_com.FECHA_EMISION = fechaemi
            cw_com.FICHA = idficha
            cw_com.ID_ESTADO = 1
            cw_com.ID_LIBRO = idficha
            If (cw_com.guardar()) Then
            Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
            End If
        ElseIf tipoinformesw = "Agua" Then 'SI EL TIPO DE INFORME ES DE AGUA
            tipoinforme = 3
            Dim aw_com As New dAguaWeb_com
            Dim pw_com As New dProductorWeb_com
            pw_com.USUARIO = productorweb_com
            pw_com = pw_com.buscar
            Dim idproductorweb_com As Long = pw_com.ID
            Dim fecha_emision As Date = DateFecha.Value.ToString("yyyy-MM-dd")
            Dim fechaemi As String
            fechaemi = Format(fecha_emision, "yyyy-MM-dd")
            aw_com.ID_USUARIO = idproductorweb_com
            aw_com.ABONADO = 0
            aw_com.FECHA_CREADO = fechaemi
            aw_com.FECHA_EMISION = fechaemi
            aw_com.FICHA = idficha
            aw_com.ID_ESTADO = 1
            aw_com.ID_LIBRO = idficha
            If (aw_com.guardar()) Then
            Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
            End If
        ElseIf tipoinformesw = "Parasitología" Then 'SI EL TIPO DE INFORME ES DE PARASITOLOGÍA
            tipoinforme = 6
            Dim parw_com As New dParasitologiaWeb_com
            Dim pw_com As New dProductorWeb_com
            pw_com.USUARIO = productorweb_com
            pw_com = pw_com.buscar
            Dim idproductorweb_com As Long = pw_com.ID
            Dim fecha_emision As Date = DateFecha.Value.ToString("yyyy-MM-dd")
            Dim fechaemi As String
            fechaemi = Format(fecha_emision, "yyyy-MM-dd")
            parw_com.ID_USUARIO = idproductorweb_com
            parw_com.ABONADO = 0
            parw_com.FECHA_CREADO = fechaemi
            parw_com.FECHA_EMISION = fechaemi
            parw_com.FICHA = idficha
            parw_com.ID_ESTADO = 1
            parw_com.ID_LIBRO = idficha
            If (parw_com.guardar()) Then
            Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
            End If
        ElseIf tipoinformesw = "Alimentos" Then 'SI EL TIPO DE INFORME ES DE Alimentos
            tipoinforme = 7
            Dim spw_com As New dSubproductosWeb_com
            Dim pw_com As New dProductorWeb_com
            pw_com.USUARIO = productorweb_com
            pw_com = pw_com.buscar
            Dim idproductorweb_com As Long = pw_com.ID
            Dim fecha_emision As Date = DateFecha.Value.ToString("yyyy-MM-dd")
            Dim fechaemi As String
            fechaemi = Format(fecha_emision, "yyyy-MM-dd")
            spw_com.ID_USUARIO = idproductorweb_com
            spw_com.ABONADO = 0
            spw_com.FECHA_CREADO = fechaemi
            spw_com.FECHA_EMISION = fechaemi
            spw_com.FICHA = idficha
            spw_com.ID_ESTADO = 1
            spw_com.ID_LIBRO = idficha
            If (spw_com.guardar()) Then
            Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
            End If
        ElseIf tipoinformesw = "Serología" Then 'SI EL TIPO DE INFORME ES DE SEROLOGÍA
            tipoinforme = 8
            Dim sw_com As New dSerologiaWeb_com
            Dim pw_com As New dProductorWeb_com
            pw_com.USUARIO = productorweb_com
            pw_com = pw_com.buscar
            Dim idproductorweb_com As Long = pw_com.ID
            Dim fecha_emision As Date = DateFecha.Value.ToString("yyyy-MM-dd")
            Dim fechaemi As String
            fechaemi = Format(fecha_emision, "yyyy-MM-dd")
            sw_com.ID_USUARIO = idproductorweb_com
            sw_com.ABONADO = 0
            sw_com.FECHA_CREADO = fechaemi
            sw_com.FECHA_EMISION = fechaemi
            sw_com.FICHA = idficha
            sw_com.ID_ESTADO = 1
            sw_com.ID_LIBRO = idficha
            If (sw_com.guardar()) Then
            Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
            End If
        ElseIf tipoinformesw = "Patología - Toxicología" Then 'SI EL TIPO DE INFORME ES DE PATOLOGÍA - TOXICOLOGÍA
            tipoinforme = 9
            Dim paw_com As New dPatologiaWeb_com
            Dim pw_com As New dProductorWeb_com
            pw_com.USUARIO = productorweb_com
            pw_com = pw_com.buscar
            Dim idproductorweb_com As Long = pw_com.ID
            Dim fecha_emision As Date = DateFecha.Value.ToString("yyyy-MM-dd")
            Dim fechaemi As String
            fechaemi = Format(fecha_emision, "yyyy-MM-dd")
            paw_com.ID_USUARIO = idproductorweb_com
            paw_com.ABONADO = 0
            paw_com.FECHA_CREADO = fechaemi
            paw_com.FECHA_EMISION = fechaemi
            paw_com.FICHA = idficha
            paw_com.ID_ESTADO = 1
            paw_com.ID_LIBRO = idficha
            If (paw_com.guardar()) Then
            Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
            End If
        ElseIf tipoinformesw = "Ambiental" Then 'SI EL TIPO DE INFORME ES AMBIENTAL
            tipoinforme = 11
            Dim aw_com As New dAmbientalWeb_com
            Dim pw_com As New dProductorWeb_com
            pw_com.USUARIO = productorweb_com
            pw_com = pw_com.buscar
            Dim idproductorweb_com As Long = pw_com.ID
            Dim fecha_emision As Date = DateFecha.Value.ToString("yyyy-MM-dd")
            Dim fechaemi As String
            fechaemi = Format(fecha_emision, "yyyy-MM-dd")
            aw_com.ID_USUARIO = idproductorweb_com
            aw_com.ABONADO = 0
            aw_com.FECHA_CREADO = fechaemi
            aw_com.FECHA_EMISION = fechaemi
            aw_com.FICHA = idficha
            aw_com.ID_ESTADO = 1
            aw_com.ID_LIBRO = idficha
            If (aw_com.guardar()) Then
            Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
            End If
        ElseIf tipoinformesw = "Lactómetros - Chequeos" Then 'SI EL TIPO DE INFORME ES DE LACTÓMETROS
            tipoinforme = 12
            Dim lw_com As New dLactometrosWeb_com
            Dim pw_com As New dProductorWeb_com
            pw_com.USUARIO = productorweb_com
            pw_com = pw_com.buscar
            Dim idproductorweb_com As Long = pw_com.ID
            Dim fecha_emision As Date = DateFecha.Value.ToString("yyyy-MM-dd")
            Dim fechaemi As String
            fechaemi = Format(fecha_emision, "yyyy-MM-dd")
            lw_com.ID_USUARIO = idproductorweb_com
            lw_com.ABONADO = 0
            lw_com.FECHA_CREADO = fechaemi
            lw_com.FECHA_EMISION = fechaemi
            lw_com.FICHA = idficha
            lw_com.ID_ESTADO = 1
            lw_com.ID_LIBRO = idficha
            If (lw_com.guardar()) Then
            Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
            End If
        ElseIf tipoinformesw = "Nutrición" Then 'SI EL TIPO DE INFORME ES DE NUTRICIÓN
            tipoinforme = 13
            Dim aw_com As New dAgroNutricionWeb_com
            Dim pw_com As New dProductorWeb_com
            pw_com.USUARIO = productorweb_com
            pw_com = pw_com.buscar
            Dim idproductorweb_com As Long = pw_com.ID
            Dim fecha_emision As Date = DateFecha.Value.ToString("yyyy-MM-dd")
            Dim fechaemi As String
            fechaemi = Format(fecha_emision, "yyyy-MM-dd")
            aw_com.ID_USUARIO = idproductorweb_com
            aw_com.ABONADO = 0
            aw_com.FECHA_CREADO = fechaemi
            aw_com.FECHA_EMISION = fechaemi
            aw_com.FICHA = idficha
            aw_com.ID_ESTADO = 1
            aw_com.ID_LIBRO = idficha
            If (aw_com.guardar()) Then
            Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
            End If
        ElseIf tipoinformesw = "Otros Servicios" Then 'SI EL TIPO DE INFORME ES DE OTROS SERVICIOS
            tipoinforme = 99
            Dim osw_com As New dOtrosServiciosWeb_com
            Dim pw_com As New dProductorWeb_com
            pw_com.USUARIO = productorweb_com
            pw_com = pw_com.buscar
            Dim idproductorweb_com As Long = pw_com.ID
            Dim fecha_emision As Date = DateFecha.Value.ToString("yyyy-MM-dd")
            Dim fechaemi As String
            fechaemi = Format(fecha_emision, "yyyy-MM-dd")
            osw_com.ID_USUARIO = idproductorweb_com
            osw_com.ABONADO = 0
            osw_com.FECHA_CREADO = fechaemi
            osw_com.FECHA_EMISION = fechaemi
            osw_com.FICHA = idficha
            osw_com.ID_ESTADO = 1
            osw_com.ID_LIBRO = idficha
            If (osw_com.guardar()) Then
            Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
            End If
        ElseIf tipoinformesw = "Suelos" Then 'SI EL TIPO DE INFORME ES DE SUELOS
            tipoinforme = 14
            Dim aw_com As New dAgroSuelosWeb_com
            Dim pw_com As New dProductorWeb_com
            pw_com.USUARIO = productorweb_com
            pw_com = pw_com.buscar
            Dim idproductorweb_com As Long = pw_com.ID
            Dim fecha_emision As Date = DateFecha.Value.ToString("yyyy-MM-dd")
            Dim fechaemi As String
            fechaemi = Format(fecha_emision, "yyyy-MM-dd")
            aw_com.ID_USUARIO = idproductorweb_com
            aw_com.ABONADO = 0
            aw_com.FECHA_CREADO = fechaemi
            aw_com.FECHA_EMISION = fechaemi
            aw_com.FICHA = idficha
            aw_com.ID_ESTADO = 1
            aw_com.ID_LIBRO = idficha
            If (aw_com.guardar()) Then
            Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
            End If
        ElseIf tipoinformesw = "Brucelosis en leche" Then 'SI EL TIPO DE INFORME ES DE BRUCELOSIS EN LECHE
            tipoinforme = 15
            Dim bw_com As New dBrucelosisLecheWeb_com
            Dim pw_com As New dProductorWeb_com
            pw_com.USUARIO = productorweb_com
            pw_com = pw_com.buscar
            Dim idproductorweb_com As Long = pw_com.ID
            Dim fecha_emision As Date = DateFecha.Value.ToString("yyyy-MM-dd")
            Dim fechaemi As String
            fechaemi = Format(fecha_emision, "yyyy-MM-dd")
            bw_com.ID_USUARIO = idproductorweb_com
            bw_com.ABONADO = 0
            bw_com.FECHA_CREADO = fechaemi
            bw_com.FECHA_EMISION = fechaemi
            bw_com.FICHA = idficha
            bw_com.ID_ESTADO = 1
            bw_com.ID_LIBRO = idficha
            If (bw_com.guardar()) Then
            Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
            End If
        ElseIf tipoinformesw = "Bacteriología y Antibiograma" Then 'SI EL TIPO DE INFORME ES DE BACTERIOLOGIA Y ANTIBIOGRAMA
            tipoinforme = 4
            Dim aw_com As New dAntibiogramaWeb_com
            Dim pw_com As New dProductorWeb_com
            pw_com.USUARIO = productorweb_com
            pw_com = pw_com.buscar
            Dim idproductorweb_com As Long = pw_com.ID
            Dim fecha_emision As Date = DateFecha.Value.ToString("yyyy-MM-dd")
            Dim fechaemi As String
            fechaemi = Format(fecha_emision, "yyyy-MM-dd")
            aw_com.ID_USUARIO = idproductorweb_com
            aw_com.ABONADO = 0
            aw_com.FECHA_CREADO = fechaemi
            aw_com.FECHA_EMISION = fechaemi
            aw_com.FICHA = idficha
            aw_com.ID_ESTADO = 1
            aw_com.ID_LIBRO = idficha
            If (aw_com.guardar()) Then
            Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
            End If
        ElseIf tipoinformesw = "Efluentes" Then 'SI EL TIPO DE INFORME ES DE BACTERIOLOGIA Y ANTIBIOGRAMA
            tipoinforme = 16
            Dim ew_com As New defluentesWeb_com
            Dim pw_com As New dProductorWeb_com
            pw_com.USUARIO = productorweb_com
            pw_com = pw_com.buscar
            Dim idproductorweb_com As Long = pw_com.ID
            Dim fecha_emision As Date = DateFecha.Value.ToString("yyyy-MM-dd")
            Dim fechaemi As String
            fechaemi = Format(fecha_emision, "yyyy-MM-dd")
            ew_com.ID_USUARIO = idproductorweb_com
            ew_com.ABONADO = 0
            ew_com.FECHA_CREADO = fechaemi
            ew_com.FECHA_EMISION = fechaemi
            ew_com.FICHA = idficha
            ew_com.ID_ESTADO = 1
            ew_com.ID_LIBRO = idficha
            If (ew_com.guardar()) Then
            Else : MsgBox("Error", MsgBoxStyle.Critical, "Atención")
            End If
        End If
        enviomail_sw()
        enviar_notificacion_solicitud()
        'Insert tabla preinforme **********************************************
        Dim pi As New dPreinformes
        Dim fechaactual As Date = Now()
        Dim _fecha As String
        _fecha = Format(fechaactual, "yyyy-MM-dd")
        pi.FICHA = idficha
        pi = pi.buscar
        If Not pi Is Nothing Then
        Else
            Dim pi2 As New dPreinformes
            pi2.FICHA = idficha
            pi2.TIPO = tipoinforme
            pi2.CREADO = 0
            pi2.FECHA = _fecha
            pi2.guardar()
            pi2 = Nothing
        End If
        pi = Nothing
        '***********************************************************************************
        '*** CREA RESULTADO EN GESTOR NUEVO *******************************************************************************************
        Dim resultado As New Dictionary(Of String, dResultado)
        Dim carpeta As String = ""
        If tipoinforme = 1 Then
            carpeta = "control_lechero"
        ElseIf tipoinforme = 3 Then
            carpeta = "agua"
        ElseIf tipoinforme = 4 Then
            carpeta = "antibiograma"
        ElseIf tipoinforme = 6 Then
            carpeta = "parasitologia"
        ElseIf tipoinforme = 7 Then
            carpeta = "productos_subproductos"
        ElseIf tipoinforme = 8 Then
            carpeta = "serologia"
        ElseIf tipoinforme = 9 Then
            carpeta = "patologia"
        ElseIf tipoinforme = 10 Then
            carpeta = "calidad_de_leche"
        ElseIf tipoinforme = 11 Then
            carpeta = "ambiental"
        ElseIf tipoinforme = 13 Then
            carpeta = "agro_nutricion"
        ElseIf tipoinforme = 14 Then
            carpeta = "agro_suelos"
        ElseIf tipoinforme = 15 Then
            carpeta = "brucelosis_leche"
        ElseIf tipoinforme = 16 Then
            carpeta = "agua"
        ElseIf tipoinforme = 17 Then
            carpeta = "antibiograma"
        ElseIf tipoinforme = 18 Then
            carpeta = "antibiograma"
        ElseIf tipoinforme = 19 Then
            carpeta = "agro_suelos"
        End If
        Dim rg As New dResultado
        Dim fechaemi2 As String
        Dim fecha_emision2 As Date = DateFecha.Value.ToString("yyyy-MM-dd")
        fechaemi2 = Format(fecha_emision2, "yyyy-MM-dd")
        rg.ficha = id_ficha_
        rg.comentarios = ""
        rg.idnet_usuario = idnet
        rg.abonado = True
        rg.fecha_creado = fechaemi2
        rg.fecha_emision = fechaemi2
        rg.path_excel = ""
        rg.path_pdf = ""
        rg.path_csv = ""
        rg.id_estado = 1
        rg.id_libro = id_ficha_
        rg.idnet_tipo_informe = tipoinforme
        resultado.Add("resultado", rg)
        Dim parameters As String = JsonConvert.SerializeObject(resultado, Formatting.None)
        Dim status As HttpStatusCode = HttpStatusCode.ExpectationFailed
        Dim response As Byte() = PostResponse("http://colaveco-gestor.herokuapp.com/resultados", "POST", parameters, status)
    End Sub
    Private Sub subir_ctacte()
        Dim fechaactual As Date = Now()
        Dim _fecha As String
        _fecha = Format(fechaactual, "yyyy-MM-dd")
        Dim movimientosDict As New Dictionary(Of String, List(Of dMovimientos))
        Dim movimientos As New List(Of dMovimientos)
        Dim m As New dMovCte2
        Dim listamovimientos As New ArrayList
        listamovimientos = m.listarxfecha(_fecha)
        If Not listamovimientos Is Nothing Then
            For Each m In listamovimientos
                barra = barra + 1
                ProgressBar1.Value = barra
                If barra = 100 Then
                    barra = 0
                End If
                Dim movi As New dMovimientos
                Dim _path As String = ""
                Dim path As String = "/home/colaveco/public_html/gestor/facturas/"
                movi.idnet_movimiento = m.MCCNRO
                movi.fecha_emision = m.MCCFCH
                movi.fecha_vencimiento = m.MCCVTO
                Dim tipo As String = ""
                If m.DOCCOD = "NF" Then
                    tipo = "NC"
                ElseIf m.DOCCOD = "NI" Then
                    tipo = "NC"
                ElseIf m.DOCCOD = "01" Then
                    tipo = "R"
                ElseIf m.DOCCOD = "02" Then
                    tipo = "ND"
                ElseIf m.DOCCOD = "AA" Then
                    tipo = "AA"
                ElseIf m.DOCCOD = "AD" Then
                    tipo = "AD"
                ElseIf m.DOCCOD = "CI" Then
                    tipo = "F"
                ElseIf m.DOCCOD = "FA" Then
                    tipo = "F"
                ElseIf m.DOCCOD = "FF" Then
                    tipo = "F"
                ElseIf m.DOCCOD = "NC" Then
                    tipo = "NC"
                End If
                movi.tipo = tipo
                movi.numero = m.MCCDOC
                movi.detalle = m.MCCDES
                movi.importe = m.MCCIMP
                movi.tipo_movimiento = m.MCCCOD
                movi.importe_pago = m.MCCPAG
                movi.idnet_usuario = m.CLICOD
                If m.MCCTIP = "V" Then
                    Dim f As New dFactur
                    f.FACNRO = m.MCCCMP
                    f = f.buscar
                    If Not f Is Nothing Then
                        Dim c As String = "\\192.168.1.106\apls\soft\"
                        Dim tx As String = f.FACPDF
                        If tx.Contains(c) Then
                            _path = tx.Replace(c, "")
                            factura_origen = _path
                            path = path & _path
                        End If
                    End If
                End If
                movi.path_pdf = path
                movimientos.Add(movi)
                movi = Nothing
                tipo = Nothing
                subirFacturaPdf()
            Next
            movimientosDict.Add("movimientos", movimientos)
            Dim parameters As String = JsonConvert.SerializeObject(movimientosDict, Formatting.None)
            Dim status As HttpStatusCode = HttpStatusCode.ExpectationFailed
            Dim response As Byte() = PostResponse("http://colaveco-gestor.herokuapp.com/factura_movimientos/migrar", "POST", parameters, status)
        End If
    End Sub
    Private Sub subir_ctacte2()
        Dim un As New dUltimoNumero
        Dim ultnum As Long = 0
        un = un.buscar
        ultnum = un.CTACTE
        Dim fechaactual As Date = Now()
        Dim _fecha As String
        _fecha = Format(fechaactual, "yyyy-MM-dd")
        Dim movimientosDict As New Dictionary(Of String, List(Of dMovimientos))
        Dim movimientos As New List(Of dMovimientos)
        Dim m As New dMovCte2
        Dim listamovimientos As New ArrayList
        listamovimientos = m.listarxid(ultnum)
        If Not listamovimientos Is Nothing Then
            For Each m In listamovimientos
                barra = barra + 1
                ProgressBar1.Value = barra
                If barra = 100 Then
                    barra = 0
                End If
                Dim movi As New dMovimientos
                Dim _path As String = ""
                Dim path As String = "/home/colaveco/public_html/gestor/facturas/"
                movi.idnet_movimiento = m.MCCNRO
                movi.fecha_emision = m.MCCFCH
                movi.fecha_vencimiento = m.MCCVTO
                Dim tipo As String = ""
                If m.DOCCOD = "NF" Then
                    tipo = "NC"
                ElseIf m.DOCCOD = "NI" Then
                    tipo = "NC"
                ElseIf m.DOCCOD = "01" Then
                    tipo = "R"
                ElseIf m.DOCCOD = "02" Then
                    tipo = "ND"
                ElseIf m.DOCCOD = "AA" Then
                    tipo = "AA"
                ElseIf m.DOCCOD = "AD" Then
                    tipo = "AD"
                ElseIf m.DOCCOD = "CI" Then
                    tipo = "F"
                ElseIf m.DOCCOD = "FA" Then
                    tipo = "F"
                ElseIf m.DOCCOD = "FF" Then
                    tipo = "F"
                ElseIf m.DOCCOD = "NC" Then
                    tipo = "NC"
                End If
                movi.tipo = tipo
                movi.numero = m.MCCDOC
                movi.detalle = m.MCCDES
                movi.importe = m.MCCIMP
                movi.tipo_movimiento = m.MCCCOD
                movi.importe_pago = m.MCCPAG
                movi.idnet_usuario = m.CLICOD
                If m.MCCTIP = "V" Then
                    Dim f As New dFactur
                    f.FACNRO = m.MCCCMP
                    f = f.buscar
                    If Not f Is Nothing Then
                        Dim c As String = "\\192.168.1.106\apls\soft\"
                        Dim tx As String = f.FACPDF
                        If tx.Contains(c) Then
                            _path = tx.Replace(c, "")
                            factura_origen = _path
                            path = path & _path
                        End If
                    End If
                End If
                movi.path_pdf = path
                movimientos.Add(movi)
                movi = Nothing
                tipo = Nothing
                subirFacturaPdf()
                un.CTACTE = m.MCCNRO
                un.modificarctacte()
            Next
            movimientosDict.Add("movimientos", movimientos)
            Dim parameters As String = JsonConvert.SerializeObject(movimientosDict, Formatting.None)
            Dim status As HttpStatusCode = HttpStatusCode.ExpectationFailed
            Dim response As Byte() = PostResponse("http://colaveco-gestor.herokuapp.com/factura_movimientos/migrar", "POST", parameters, status)
        End If
    End Sub
    Private Sub enviar_notificacion_solicitud()
        Dim fechaactual As Date = Now()
        Dim _fecha As String
        _fecha = Format(fechaactual, "yyyy-MM-dd")
        Dim notificacion As New Dictionary(Of String, dNotificaciones)
        Dim nt As New dNotificaciones
        Dim _tipo As String = ""
        Dim _mensaje As String = ""
        Dim nuevoid As Long = 0
        Dim _detalle As String = ""
        Dim _tipoinforme As String = ""
        Dim _subtipo As String = ""
        Dim _nombrecliente As String = ""
        Dim _muestrasingresadas As String = ""
        Dim _tipomuestra As String = ""
        Dim _tiempo As String = 0
        Dim sa As New dSolicitudAnalisis
        sa.ID = idficha
        sa = sa.buscar
        If Not sa Is Nothing Then
            Dim ti As New dTipoInforme
            ti.ID = sa.IDTIPOINFORME
            ti = ti.buscar
            If Not ti Is Nothing Then
                _tipoinforme = ti.NOMBRE
            End If
            Dim sti As New dSubInforme
            sti.ID = sa.IDSUBINFORME
            sti = sti.buscar
            If Not sti Is Nothing Then
                _subtipo = sti.NOMBRE
            End If
            nuevoid = sa.IDPRODUCTOR
            Dim pro As New dCliente
            pro.ID = sa.IDPRODUCTOR
            pro = pro.buscar
            If Not pro Is Nothing Then
                _nombrecliente = pro.NOMBRE
            Else
                _nombrecliente = ""
            End If
            If sa.NMUESTRAS = 0 Then
                _muestrasingresadas = "n/a"
            Else
                _muestrasingresadas = sa.NMUESTRAS
            End If
            Dim m As New dMuestras
            m.ID = sa.IDMUESTRA
            m = m.buscar
            If Not m Is Nothing Then
                _tipomuestra = m.NOMBRE
            End If
        End If
        _tipo = "solicitud_creada"
        _mensaje = "Ha ingresado una solicitud de análisis de " & _tipoinforme & ", con el número " & idficha
        '******************************************************************************************************************************************
        Dim texto As String = ""
        Dim texto2 As String = ""
        Dim texto3 As String = ""
        Dim sm As New dRelSolicitudMuestras
        Dim spal As New dSolicitudPAL
        Dim csm As New dCalidadSolicitudMuestra
        Dim cs As New dControlSolicitud
        Dim a2 As New dAntibiograma2
        Dim sn As New dSolicitudNutricion
        Dim ss As New dSolicitudSuelos
        Dim bl As New dBrucelosis
        Dim lista2 As New ArrayList
        Dim lista3 As New ArrayList
        Dim lista4 As New ArrayList
        Dim lista5 As New ArrayList
        Dim lista6 As New ArrayList
        Dim lista7 As New ArrayList
        Dim lista10 As New ArrayList
        Dim listabl As New ArrayList
        Dim listanutricion As New ArrayList
        Dim listasuelos As New ArrayList
        lista4 = sm.listarporficha(idficha)
        lista5 = csm.listarporsolicitud3(idficha)
        lista6 = cs.listarporsolicitud(idficha)
        lista7 = a2.listarporsolicitud(idficha)
        lista10 = spal.listarporsolicitud(idficha)
        listanutricion = sn.listarporsolicitud(idficha)
        listasuelos = ss.listarporsolicitud(idficha)
        listabl = sm.listarporficha(idficha)
        ' SI ES PRODUCTOS LÁCTEOS ********************************************************************************
        If _tipoinforme = "Alimentos" Then
            _tiempo = "3"
            Dim sp As New dSubproducto
            Dim lista As New ArrayList
            texto = ""
            lista = sp.listarporsolicitud(idficha)
            If Not lista Is Nothing Then
                If lista.Count > 0 Then
                    For Each sp In lista
                        If sp.ESTAFCOAGPOSITIVO = 1 Then
                            texto = texto + " - Estaf. Coag. Positivo"
                        End If
                        If sp.CF = 1 Then
                            texto = texto + " - CF"
                        End If
                        If sp.MOHOSYLEVADURAS = 1 Then
                            texto = texto + " - Mohos y levaduras"
                            _tiempo = "5"
                        End If
                        If sp.CT = 1 Then
                            texto = texto + " - Coliformes Totales"
                        End If
                        If sp.ECOLI = 1 Then
                            texto = texto + " - E. Coli"
                        End If
                        If sp.SALMONELLA = 1 Then
                            texto = texto + " - Salmonella"
                            _tiempo = "7"
                        End If
                        If sp.LISTERIASPP = 1 Then
                            texto = texto + " - Listeria spp"
                            _tiempo = "7"
                        End If
                        If sp.HUMEDAD = 1 Then
                            texto = texto + " - Humedad"
                        End If
                        If sp.MGRASA = 1 Then
                            texto = texto + " - M. Grasa"
                        End If
                        If sp.PH = 1 Then
                            texto = texto + " - pH"
                        End If
                        If sp.CLORUROS = 1 Then
                            texto = texto + " - Cloruros"
                        End If
                        If sp.PROTEINAS = 1 Then
                            texto = texto + " - Proteínas"
                        End If
                        If sp.ENTEROBACTERIAS = 1 Then
                            texto = texto + " - Enterobacterias"
                        End If
                        If sp.LISTERIAAMBIENTAL = 1 Then
                            texto = texto + " - Listeria Ambiental"
                            _tiempo = "7"
                        End If
                        If sp.ESPORANAERMESOFILO = 1 Then
                            texto = texto + " - Espor. Anaer. Mesófilos"
                        End If
                        If sp.TERMOFILOS = 1 Then
                            texto = texto + " - Termodúricos"
                        End If
                        If sp.PSICROTROFOS = 1 Then
                            texto = texto + " - Psicrótrofos"
                        End If
                        If sp.RB = 1 Then
                            texto = texto + " - RB"
                        End If
                        If sp.TABLANUTRICIONAL = 1 Then
                            texto = texto + " - Tabla nutricional"
                        End If
                        If sp.LISTERIAMONOCITOGENES = 1 Then
                            texto = texto + " - Listeria monocitógenes"
                            _tiempo = "7"
                        End If
                        If sp.CENIZAS = 1 Then
                            texto = texto + " - Cenizas"
                        End If
                    Next
                End If
            End If
            ' SI ES AGUA ********************************************************************************
        ElseIf _tipoinforme = "Agua" Then
            _tiempo = "2"
            Dim a1 As New dAgua
            texto = ""
            a1.ID = idficha
            a1 = a1.buscar()
            texto = _subtipo
            If Not a1 Is Nothing Then
                If a1.HET22 = 1 Then
                    texto = texto & " " & " - Heterotróficos 22"
                End If
                If a1.HET35 = 1 Then
                    texto = texto & " " & " - Heterotróficos 35"
                End If
                If a1.HET37 = 1 Then
                    texto = texto & " " & " - Heterotróficos 37"
                End If
                If a1.CLORO = 1 Then
                    texto = texto & " " & " - Cloro"
                End If
                If a1.CONDUCTIVIDAD = 1 Then
                    texto = texto & " " & " - Conductividad"
                End If
                If a1.PH = 1 Then
                    texto = texto & " " & " - pH"
                End If
                If a1.ECOLI = 1 Then
                    texto = texto & " " & " - Ecoli"
                End If
            End If
            ' SI ES CALIDAD DE LECHE ********************************************************************************
        ElseIf _tipoinforme = "Calidad de leche" Then
            _tiempo = "1"
            Dim rb As Integer = 0
            Dim rc As Integer = 0
            Dim comp As Integer = 0
            Dim criosc As Integer = 0
            Dim inh As Integer = 0
            Dim espor As Integer = 0
            Dim urea As Integer = 0
            Dim term As Integer = 0
            Dim psicr As Integer = 0
            Dim crioscopo As Integer = 0
            texto = ""
            If Not lista5 Is Nothing Then
                If lista5.Count > 0 Then
                    For Each csm In lista5
                        If csm.RB = 1 Then
                            rb = 1
                        End If
                        If csm.RC = 1 Then
                            rc = 1
                        End If
                        If csm.COMPOSICION = 1 Then
                            comp = 1
                        End If
                        If csm.CRIOSCOPIA = 1 Then
                            criosc = 1
                        End If
                        If csm.INHIBIDORES = 1 Then
                            inh = 1
                        End If
                        If csm.ESPORULADOS = 1 Then
                            espor = 1
                        End If
                        If csm.UREA = 1 Then
                            urea = 1
                        End If
                        If csm.TERMOFILOS = 1 Then
                            term = 1
                        End If
                        If csm.PSICROTROFOS = 1 Then
                            psicr = 1
                        End If
                        If csm.CRIOSCOPIA_CRIOSCOPO = 1 Then
                            crioscopo = 1
                        End If
                    Next
                End If
            End If
            If rb = 1 Then
                texto = texto + " - RB"
            End If
            If rc = 1 Then
                texto = texto + " - RC"
            End If
            If comp = 1 Then
                texto = texto + " - Composición"
            End If
            If criosc = 1 Then
                texto = texto + " - Crioscopía"
            End If
            If inh = 1 Then
                texto = texto + " - Inhibidores"
            End If
            If espor = 1 Then
                texto = texto + " - Esporulados"
                _tiempo = "7"
            End If
            If urea = 1 Then
                texto = texto + " - Urea"
            End If
            If term = 1 Then
                texto = texto + " - Termófilos"
            End If
            If psicr = 1 Then
                texto = texto + " - Psicrótrofos"
            End If
            If crioscopo = 1 Then
                texto = texto + " - Crioscopía (crióscopo)"
            End If
            ' SI ES CONTROL LECHERO ********************************************************************************
        ElseIf _tipoinforme = "Control Lechero" Then
            _tiempo = "1"
            Dim rc As Integer = 0
            Dim comp As Integer = 0
            Dim urea As Integer = 0
            texto = ""
            If Not lista6 Is Nothing Then
                If lista6.Count > 0 Then
                    For Each cs In lista6
                        If cs.RC = 1 Then
                            rc = 1
                        End If
                        If cs.COMPOSICION = 1 Then
                            comp = 1
                        End If
                        If cs.UREA = 1 Then
                            urea = 1
                        End If
                    Next
                End If
            End If
            If rc = 1 Then
                texto = texto + " - RC"
            End If
            If comp = 1 Then
                texto = texto + " - Composición"
            End If
            If urea = 1 Then
                texto = texto + " - Urea"
            End If
            ' SI ANTIBIOGRAMA ********************************************************************************
        ElseIf _tipoinforme = "Bacteriología y Antibiograma" Then
            _tiempo = "3"
            Dim aislamiento As Integer = 0
            Dim antibiograma As Integer = 0
            texto = ""
            If Not lista7 Is Nothing Then
                If lista7.Count > 0 Then
                    For Each a2 In lista7
                        If a2.AISLAMIENTO = 1 Then
                            aislamiento = 1
                        End If
                        If a2.ANTIBIOGRAMA = 1 Then
                            antibiograma = 1
                        End If
                    Next
                End If
            End If
            If aislamiento = 1 Then
                texto = texto + " - Aislamiento"
            End If
            If antibiograma = 1 Then
                texto = texto + " - Antibiograma"
            End If
            ' SI ES AMBIENTAL ********************************************************************************
        ElseIf _tipoinforme = "Ambiental" Then
            _tiempo = "2"
            Dim ambs As New dAmbientalSolicitud
            Dim lista8 As ArrayList
            lista8 = ambs.listarporsolicitud(idficha)
            Dim enterobacterias As Integer = 0
            Dim listambiental As Integer = 0
            Dim listmono As Integer = 0
            Dim salmonella As Integer = 0
            Dim ecoli As Integer = 0
            Dim mohosylevaduras As Integer = 0
            Dim rb As Integer = 0
            Dim ct As Integer = 0
            Dim cf As Integer = 0
            Dim pseudomonaspp As Integer = 0
            texto = ""
            If Not lista8 Is Nothing Then
                If lista8.Count > 0 Then
                    For Each ambs In lista8
                        If ambs.ENTEROBACTERIAS = 1 Then
                            enterobacterias = 1
                        End If
                        If ambs.LISTAMBIENTAL = 1 Then
                            listambiental = 1
                        End If
                        If ambs.LISTMONO = 1 Then
                            listmono = 1
                        End If
                        If ambs.SALMONELLA = 1 Then
                            salmonella = 1
                        End If
                        If ambs.ECOLI = 1 Then
                            ecoli = 1
                        End If
                        If ambs.MOHOSYLEVADURAS = 1 Then
                            mohosylevaduras = 1
                        End If
                        If ambs.RB = 1 Then
                            rb = 1
                        End If
                        If ambs.CT = 1 Then
                            ct = 1
                        End If
                        If ambs.CF = 1 Then
                            cf = 1
                        End If
                        If ambs.PSEUDOMONASPP = 1 Then
                            pseudomonaspp = 1
                        End If
                    Next
                End If
            End If
            If enterobacterias = 1 Then
                texto = texto + " - Enterobacterias"
            End If
            If listambiental = 1 Then
                texto = texto + " - Listeria ambiental"
            End If
            If listmono = 1 Then
                texto = texto + " - Listeria monocitógenes"
            End If
            If salmonella = 1 Then
                texto = texto + " - Salmonella"
            End If
            If ecoli = 1 Then
                texto = texto + " - E. Coli"
            End If
            If mohosylevaduras = 1 Then
                texto = texto + " - Mohos y levaduras"
            End If
            If rb = 1 Then
                texto = texto + " - RB"
            End If
            If ct = 1 Then
                texto = texto + " - Coliformes totales"
            End If
            If cf = 1 Then
                texto = texto + " - CF"
            End If
            If pseudomonaspp = 1 Then
                texto = texto + " - Pseudomona spp"
            End If
            ' SI ES PARASITOLOGÍA ********************************************************************************
        ElseIf _tipoinforme = "Parasitología" Then
            _tiempo = "2"
            Dim p As New dParasitologiaSolicitud
            Dim lista9 As ArrayList
            lista9 = p.listarporsolicitud(idficha)
            Dim gastrointestinales As Integer = 0
            Dim fasciola As Integer = 0
            Dim coccidias As Integer = 0
            texto = ""
            If Not lista9 Is Nothing Then
                If lista9.Count > 0 Then
                    For Each p In lista9
                        If p.GASTROINTESTINALES = 1 Then
                            gastrointestinales = 1
                        End If
                        If p.FASCIOLA = 1 Then
                            fasciola = 1
                        End If
                        If p.COCCIDIAS = 1 Then
                            coccidias = 1
                        End If
                    Next
                End If
            End If
            If gastrointestinales = 1 Then
                texto = texto + " - Gastrointestinales"
            End If
            If fasciola = 1 Then
                texto = texto + " - Fasciola"
            End If
            If coccidias = 1 Then
                texto = texto + " - Coccidias"
            End If
            ' SI ES NUTRICIÓN ********************************************************************************
        ElseIf _tipoinforme = "Nutrición" Then
            _tiempo = "4"
            Dim mga As Integer = 0
            Dim mgb As Integer = 0
            Dim ensilados As Integer = 0
            Dim pasturas As Integer = 0
            Dim extetereo As Integer = 0
            Dim nida As Integer = 0
            texto = ""
            If Not listanutricion Is Nothing Then
                If listanutricion.Count > 0 Then
                    For Each sn In listanutricion
                        texto = texto & " // " & sn.MUESTRA & " - "
                        If sn.MGA = 1 Then
                            texto = texto & "MGA - "
                        End If
                        If sn.MGB = 1 Then
                            texto = texto & "MGB - "
                        End If
                        If sn.ENSILADOS = 1 Then
                            texto = texto & "Ensilados - "
                        End If
                        If sn.PASTURAS = 1 Then
                            texto = texto & "Pasturas - "
                        End If
                        If sn.EXTETEREO = 1 Then
                            texto = texto & "Extracto etéreo - "
                        End If
                        If sn.NIDA = 1 Then
                            texto = texto & "NIDA - "
                        End If
                    Next
                End If
            End If
            ' SI ES SUELOS ********************************************************************************
        ElseIf _tipoinforme = "Suelos" Then
            _tiempo = "2"
            Dim nitratos As Integer = 0
            Dim mineralizacion As Integer = 0
            Dim fosforobray As Integer = 0
            Dim fosforocitrico As Integer = 0
            Dim phagua As Integer = 0
            Dim phkci As Integer = 0
            Dim materiaorg As Integer = 0
            Dim potasioint As Integer = 0
            Dim sulfatos As Integer = 0
            Dim nitrogenovegetal As Integer = 0
            texto = ""
            If Not listasuelos Is Nothing Then
                If listasuelos.Count > 0 Then
                    For Each ss In listasuelos
                        texto = texto & " // " & ss.MUESTRA & " - "
                        If ss.NITRATOS = 1 Then
                            texto = texto & "Nitratos - "
                        End If
                        If ss.MINERALIZACION = 1 Then
                            texto = texto & "Mineralización - "
                        End If
                        If ss.FOSFOROBRAY = 1 Then
                            texto = texto & "Fósforo Bray I - "
                        End If
                        If ss.FOSFOROCITRICO = 1 Then
                            texto = texto & "Fósforo Ac.Cítrico - "
                        End If
                        If ss.PHAGUA = 1 Then
                            texto = texto & "pH Agua - "
                        End If
                        If ss.PHKCI = 1 Then
                            texto = texto & "pH KCI - "
                        End If
                        If ss.MATERIAORG = 1 Then
                            texto = texto & "Materia orgánica - "
                        End If
                        If ss.POTASIOINT = 1 Then
                            texto = texto & "Potasio intercambiable - "
                        End If
                        If ss.SULFATOS = 1 Then
                            texto = texto & "Sulfatos - "
                        End If
                        If ss.NITROGENOVEGETAL = 1 Then
                            texto = texto & "Nitrógeno vegetal - "
                        End If
                    Next
                End If
            End If
        ElseIf _tipoinforme = "Otros" Then
            _tiempo = "7"
        ElseIf _tipoinforme = "Serología" Then
            If _subtipo = "Brucelosis" Then
                _tiempo = "2"
            ElseIf _subtipo = "Serología otros" Then
                _tiempo = "10"
            End If
        ElseIf _tipoinforme = "Patología - Toxicología" Then
            _tiempo = "5"
        End If
        '*** LISTADO DE MUESTRAS *********************************************************************************
        ' SI ES PRODUCTOS LÁCTEOS ********************************************************************************
        If _tipoinforme = "Alimentos" Then
            texto2 = ""
            If Not lista4 Is Nothing Then
                If lista4.Count > 0 Then
                    For Each sm In lista4
                        texto2 = texto2 + sm.IDMUESTRA & " - "
                    Next
                End If
            End If
            ' SI ES AGUA ********************************************************************************
        ElseIf _tipoinforme = "Agua" Then
            texto2 = ""
            If Not lista4 Is Nothing Then
                If lista4.Count > 0 Then
                    For Each sm In lista4
                        texto2 = texto2 + sm.IDMUESTRA & " - "
                    Next
                End If
            End If
            ' SI ES CALIDAD ********************************************************************************
        ElseIf _tipoinforme = "Calidad de leche" Then
            texto2 = ""
            Dim cuenta_rb As Integer = 0
            Dim cuenta_rc As Integer = 0
            Dim cuenta_comp As Integer = 0
            Dim cuenta_criosc As Integer = 0
            Dim cuenta_inhib As Integer = 0
            Dim cuenta_espor As Integer = 0
            Dim cuenta_urea As Integer = 0
            Dim cuenta_termo As Integer = 0
            Dim cuenta_psicro As Integer = 0
            Dim cuenta_criosc_criosc As Integer = 0
            Dim cuenta_caseina As Integer = 0
            If Not lista5 Is Nothing Then
                If lista5.Count > 0 Then
                    For Each csm In lista5
                        texto2 = texto2 + csm.MUESTRA
                        If csm.RB = 1 Then
                            cuenta_rb = cuenta_rb + 1
                        End If
                        If csm.RC = 1 Then
                            cuenta_rc = cuenta_rc + 1
                        End If
                        If csm.COMPOSICION = 1 Then
                            cuenta_comp = cuenta_comp + 1
                        End If
                        If csm.CRIOSCOPIA = 1 Then
                            cuenta_criosc = cuenta_criosc + 1
                        End If
                        If csm.INHIBIDORES = 1 Then
                            cuenta_inhib = cuenta_inhib + 1
                        End If
                        If csm.ESPORULADOS = 1 Then
                            cuenta_espor = cuenta_espor + 1
                        End If
                        If csm.UREA = 1 Then
                            cuenta_urea = cuenta_urea + 1
                        End If
                        If csm.TERMOFILOS = 1 Then
                            cuenta_termo = cuenta_termo + 1
                        End If
                        If csm.PSICROTROFOS = 1 Then
                            cuenta_psicro = cuenta_psicro + 1
                        End If
                        If csm.CRIOSCOPIA_CRIOSCOPO = 1 Then
                            cuenta_criosc_criosc = cuenta_criosc_criosc + 1
                        End If
                        If csm.CASEINA = 1 Then
                            cuenta_caseina = cuenta_caseina + 1
                        End If
                        texto2 = texto2 + " - "
                    Next
                End If
            End If
            If cuenta_rb > 0 Then
                texto3 = texto3 & cuenta_rb & " RB - "
            End If
            If cuenta_rc > 0 Then
                texto3 = texto3 & cuenta_rc & " RC - "
            End If
            If cuenta_comp > 0 Then
                texto3 = texto3 & cuenta_comp & " Comp. - "
            End If
            If cuenta_criosc > 0 Then
                texto3 = texto3 & cuenta_criosc & " Criosc. - "
            End If
            If cuenta_inhib > 0 Then
                texto3 = texto3 & cuenta_inhib & " Inhib. - "
            End If
            If cuenta_espor > 0 Then
                texto3 = texto3 & cuenta_espor & " Espor. - "
            End If
            If cuenta_urea > 0 Then
                texto3 = texto3 & cuenta_urea & " Urea - "
            End If
            If cuenta_termo > 0 Then
                texto3 = texto3 & cuenta_termo & " Termof. - "
            End If
            If cuenta_psicro > 0 Then
                texto3 = texto3 & cuenta_psicro & " Psicro. - "
            End If
            If cuenta_criosc_criosc > 0 Then
                texto3 = texto3 & cuenta_criosc_criosc & " Criosc.(Crióscopo) - "
            End If
            If cuenta_caseina > 0 Then
                texto3 = texto3 & cuenta_caseina & " Caseina - "
            End If
            ' SI ES CONTROL LECHERO ********************************************************************************
        ElseIf _tipoinforme = "Control Lechero" Then
            texto2 = ""
            ' SI ES ANTIBIOGRAMA ********************************************************************************
        ElseIf _tipoinforme = "Bacteriología y Antibiograma" Then
            texto2 = ""
            If Not lista4 Is Nothing Then
                If lista4.Count > 0 Then
                    For Each sm In lista4
                        texto2 = texto2 + sm.IDMUESTRA & " - "
                    Next
                End If
            End If
            ' SI ES AMBIENTAL ********************************************************************************
        ElseIf _tipoinforme = "Ambiental" Then
            texto2 = ""
            If Not lista4 Is Nothing Then
                If lista4.Count > 0 Then
                    For Each sm In lista4
                        texto2 = texto2 + sm.IDMUESTRA & " - "
                    Next
                End If
            End If
            ' SI ES PARASITOLOGÍA ********************************************************************************
        ElseIf _tipoinforme = "Parasitología" Then
            texto2 = ""
            If Not lista4 Is Nothing Then
                If lista4.Count > 0 Then
                    For Each sm In lista4
                        texto2 = texto2 + sm.IDMUESTRA & " - "
                    Next
                End If
            End If
            ' SI ES BRUCELOSIS LECHE ********************************************************************************
        ElseIf _tipoinforme = "Brucelosis en leche" Then
            _tiempo = "3"
            texto2 = ""
            If Not listabl Is Nothing Then
                If listabl.Count > 0 Then
                    For Each sm In listabl
                        texto2 = texto2 + sm.IDMUESTRA & " - "
                    Next
                End If
            End If
        End If
        '********************************************************************************************************************
        _detalle = "<p><b>Fecha de recepción:</b> " + _fecha + " </p><p><b>A nombre de:</b> " + _nombrecliente + " </p><p><b>Muestras ingresadas:</b> " + _muestrasingresadas + " </p><p><b>Tipo de muestra:</b> " + _tipomuestra + " </p><p><b>Tipo de informe:</b> " + _tipoinforme + " </p><p><b>Subtipo:</b> " + _subtipo + " </p><p><b>Análisis Solicitado:</b> " + texto + " </p><p><b>Identificación de las muestras:</b> " + texto2 + " </p><p><b>Tiempo estimado de entrega:</b> " + _tiempo + " </p>"
        nt.fecha = _fecha
        nt.tipo = _tipo
        nt.mensaje = _mensaje
        nt.idnet_usuario = nuevoid
        notificacion.Add("notification", nt)
        Dim parameters As String = JsonConvert.SerializeObject(notificacion, Formatting.None)
        Dim status As HttpStatusCode = HttpStatusCode.ExpectationFailed
        Dim response As Byte() = PostResponse("http://colaveco-gestor.herokuapp.com/notifications", "POST", parameters, status)
    End Sub
    Private Sub enviar_notificacion_resultado()
        Dim fechaactual As Date = Now()
        Dim _fecha As String
        _fecha = Format(fechaactual, "yyyy-MM-dd")
        Dim notificacion As New Dictionary(Of String, dNotificaciones)
        Dim nt As New dNotificaciones
        Dim _tipo As String = ""
        Dim _mensaje As String = ""
        Dim nuevoid As Long = 0
        Dim tipoinforme As String = ""
        Dim sa As New dSolicitudAnalisis
        sa.ID = idficha
        sa = sa.buscar
        If Not sa Is Nothing Then
            Dim ti As New dTipoInforme
            ti.ID = sa.IDTIPOINFORME
            ti = ti.buscar
            If Not ti Is Nothing Then
                tipoinforme = ti.NOMBRE
            End If
            nuevoid = sa.IDPRODUCTOR
        End If
        _tipo = "resultado"
        _mensaje = "El resultado de su análisis de " & tipoinforme & ", número " & idficha & " está finalizado."
        nt.fecha = _fecha
        nt.tipo = _tipo
        nt.mensaje = _mensaje
        nt.idnet_usuario = nuevoid
        notificacion.Add("notification", nt)
        Dim parameters As String = JsonConvert.SerializeObject(notificacion, Formatting.None)
        Dim status As HttpStatusCode = HttpStatusCode.ExpectationFailed
        Dim response As Byte() = PostResponse("http://colaveco-gestor.herokuapp.com/notifications", "POST", parameters, status)
    End Sub
    Public Sub modificarRegistro2(ByVal _abonado As Integer)
        Dim idnet As Long = 0
        Dim sa_ As New dSolicitudAnalisis
        Dim fechacreado As String = ""
        Dim fechaenviado As String = ""
        sa_.ID = idficha
        sa_ = sa_.buscar
        If Not sa_ Is Nothing Then
            idnet = sa_.IDPRODUCTOR
            tipoinforme = sa_.IDTIPOINFORME
            Dim c As New dCliente
            c.ID = sa_.IDPRODUCTOR
            c = c.buscar
            If Not c Is Nothing Then
                productorweb_com = c.USUARIO_WEB
                Dim pw_com As New dProductorWeb_com
                pw_com.USUARIO = productorweb_com
                pw_com = pw_com.buscar
                If Not pw_com Is Nothing Then
                    idproductorweb_com = pw_com.ID
                End If
                fechacreado = sa_.FECHAINGRESO
                fechaenviado = sa_.FECHAENVIO
            End If
        End If
        '*** CREA RESULTADO EN GESTOR NUEVO *******************************************************************************************
        Dim resultado As New Dictionary(Of String, dResultado)
        Dim carpeta As String = ""
        If tipoinforme = 1 Then
            carpeta = "control_lechero"
        ElseIf tipoinforme = 3 Then
            carpeta = "agua"
        ElseIf tipoinforme = 4 Then
            carpeta = "antibiograma"
        ElseIf tipoinforme = 6 Then
            carpeta = "parasitologia"
        ElseIf tipoinforme = 7 Then
            carpeta = "productos_subproductos"
        ElseIf tipoinforme = 8 Then
            carpeta = "serologia"
        ElseIf tipoinforme = 9 Then
            carpeta = "patologia"
        ElseIf tipoinforme = 10 Then
            carpeta = "calidad_de_leche"
        ElseIf tipoinforme = 11 Then
            carpeta = "ambiental"
        ElseIf tipoinforme = 13 Then
            carpeta = "agro_nutricion"
        ElseIf tipoinforme = 14 Then
            carpeta = "agro_suelos"
        ElseIf tipoinforme = 15 Then
            carpeta = "brucelosis_leche"
        ElseIf tipoinforme = 16 Then
            carpeta = "efluentes"
        ElseIf tipoinforme = 17 Then
            carpeta = "antibiograma"
        ElseIf tipoinforme = 18 Then
            carpeta = "antibiograma"
        ElseIf tipoinforme = 19 Then
            carpeta = "agro_suelos"
        ElseIf tipoinforme = 20 Then
            carpeta = "patologia"
        End If
        Dim rg As New dResultado
        Dim fechaemi2 As String
        Dim fecha_emision2 As Date = DateFecha.Value.ToString("yyyy-MM-dd")
        fechaemi2 = Format(fecha_emision2, "yyyy-MM-dd")
        rg.ficha = idficha
        rg.comentarios = _comentarios
        rg.idnet_usuario = idnet
        rg.abonado = _abonado
        rg.fecha_creado = fechacreado 'fechaemi2
        rg.fecha_emision = fechaenviado 'fechaemi2
        rg.path_excel = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/" & carpeta & "/" & idficha & ".xls"
        rg.path_pdf = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/" & carpeta & "/" & idficha & ".pdf"
        rg.path_csv = "/home/colaveco/public_html/gestor/data_file/" & idproductorweb_com & "/" & carpeta & "/" & idficha & ".txt"
        rg.id_estado = 3
        rg.id_libro = idficha
        rg.idnet_tipo_informe = tipoinforme
        resultado.Add("resultado", rg)
        Dim parameters As String = JsonConvert.SerializeObject(resultado, Formatting.None)
        Dim status As HttpStatusCode = HttpStatusCode.ExpectationFailed
        Dim response As Byte() = PostResponse("http://colaveco-gestor.herokuapp.com/resultados", "POST", parameters, status)
        subir_informes_gestor()
    End Sub
    Public Function subirFacturaPdf() As String
        Dim fichero As String = ""
        Dim destino As String = ""
        Dim user As String = "colaveco"
        Dim pass As String = "NUEVA**!!COL22"
        fichero = "\\192.168.1.106\apls\soft\" & factura_origen
        destino = "ftp://colaveco.com.uy/www/gestor/facturas/" & factura_origen
        Dim infoFichero As New FileInfo(fichero)
        Dim uri As String
        uri = destino
        Dim peticionFTP As FtpWebRequest
        ' Creamos una peticion FTP con la dirección del fichero que vamos a subir
        peticionFTP = CType(FtpWebRequest.Create(New Uri(destino)), FtpWebRequest)
        ' Fijamos el usuario y la contraseña de la petición
        peticionFTP.Credentials = New NetworkCredential(user, pass)
        peticionFTP.KeepAlive = False
        peticionFTP.UsePassive = False
        ' Seleccionamos el comando que vamos a utilizar: Subir un fichero
        peticionFTP.Method = WebRequestMethods.Ftp.UploadFile
        ' Especificamos el tipo de transferencia de datos
        peticionFTP.UseBinary = True
        ' Informamos al servidor sobre el tamaño del fichero que vamos a subir
        'peticionFTP.ContentLength = infoFichero.Length
        '**********************************************************************
        Try
            peticionFTP.ContentLength = infoFichero.Length
        Catch ex As Exception
            ' Si se produce algún fallo, se devolverá el mensaje del error
            Return ex.Message
        End Try
        '**********************************************************************
        ' Fijamos un buffer de 150KB
        Dim longitudBuffer As Integer
        longitudBuffer = 153600
        Dim lector As Byte() = New Byte(153600) {}
        Dim num As Integer
        ' Abrimos el fichero para subirlo
        Dim fs As FileStream
        fs = infoFichero.OpenRead()
        Try
            Dim escritor As Stream
            escritor = peticionFTP.GetRequestStream()
            ' Leemos 150 KB del fichero en cada iteración
            num = fs.Read(lector, 0, longitudBuffer)
            While (num <> 0)
                ' Escribimos el contenido del flujo de lectura en el
                ' flujo de escritura del comando FTP
                escritor.Write(lector, 0, num)
                num = fs.Read(lector, 0, longitudBuffer)
            End While
            escritor.Close()
            fs.Close()
            ' Si todo ha ido bien, se devolverá String.Empty
            Return String.Empty
        Catch ex As Exception
            ' Si se produce algún fallo, se devolverá el mensaje del error
            Return ex.Message
        End Try
    End Function
    Public Function subirFicheroXls_G() As String
        Dim fichero As String = ""
        Dim destino As String = ""
        Dim user As String = "colaveco"
        Dim pass As String = "NUEVA**!!COL22"
        If tipoinforme = 1 Then
            crea_control_lechero_com()
            fichero = "\\192.168.1.10\e\NET\CONTROL_LECHERO\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/control_lechero/" & idficha & ".xls"
        ElseIf tipoinforme = 3 Then
            crea_agua_com()
            fichero = "\\192.168.1.10\e\NET\AGUA\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/agua/" & idficha & ".xls"
        ElseIf tipoinforme = 4 Then
            crea_antibiograma_com()
            fichero = "\\192.168.1.10\e\NET\ANTIBIOGRAMA\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/antibiograma/" & idficha & ".xls"
        ElseIf tipoinforme = 6 Then
            crea_parasitologia_com()
            fichero = "\\192.168.1.10\e\NET\PARASITOLOGIA\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/parasitologia/" & idficha & ".xls"
        ElseIf tipoinforme = 7 Then
            crea_productos_subproductos_com()
            fichero = "\\192.168.1.10\e\NET\ALIMENTOS\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/productos_subproductos/" & idficha & ".xls"
        ElseIf tipoinforme = 8 Then
            crea_serologia_com()
            fichero = "\\192.168.1.10\e\NET\SEROLOGIA\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/serologia/" & idficha & ".xls"
        ElseIf tipoinforme = 9 Then
            crea_patologia_com()
            fichero = "\\192.168.1.10\e\NET\PATOLOGIA\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/patologia/" & idficha & ".xls"
        ElseIf tipoinforme = 10 Then
            crea_calidad_de_leche_com()
            fichero = "\\192.168.1.10\e\NET\CALIDAD\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/calidad_de_leche/" & idficha & ".xls"
        ElseIf tipoinforme = 11 Then
            crea_ambiental_com()
            fichero = "\\192.168.1.10\e\NET\AMBIENTAL\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/ambiental/" & idficha & ".xls"
        ElseIf tipoinforme = 13 Then
            crea_agro_nutricion_com()
            fichero = "\\192.168.1.10\e\NET\NUTRICION\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/agro_nutricion/" & idficha & ".xls"
        ElseIf tipoinforme = 14 Then
            crea_agro_suelos_com()
            fichero = "\\192.168.1.10\e\NET\Suelos\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/agro_suelos/" & idficha & ".xls"
        ElseIf tipoinforme = 15 Then
            crea_brucelosis_leche_com()
            fichero = "\\192.168.1.10\e\NET\Brucelosis en leche\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/brucelosis_leche/" & idficha & ".xls"
        ElseIf tipoinforme = 16 Then
            crea_efluentes_com()
            fichero = "\\192.168.1.10\e\NET\Efluentes\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/efluentes/" & idficha & ".xls"
        ElseIf tipoinforme = 17 Then
            crea_antibiograma_com()
            fichero = "\\192.168.1.10\e\NET\ANTIBIOGRAMA\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/antibiograma/" & idficha & ".xls"
        ElseIf tipoinforme = 18 Then
            crea_antibiograma_com()
            fichero = "\\192.168.1.10\e\NET\ANTIBIOGRAMA\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/antibiograma/" & idficha & ".xls"
        ElseIf tipoinforme = 19 Then
            crea_agro_suelos_com()
            fichero = "\\192.168.1.10\e\NET\Suelos\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/agro_suelos/" & idficha & ".xls"
        ElseIf tipoinforme = 20 Then
            crea_patologia_com()
            fichero = "\\192.168.1.10\e\NET\PATOLOGIA\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/patologia/" & idficha & ".xls"
        ElseIf tipoinforme = 99 Then
            crea_otros_servicios_com()
            fichero = "\\192.168.1.10\e\NET\OTROS\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/otros_servicios/" & idficha & ".xls"
        End If
        Dim infoFichero As New FileInfo(fichero)
        Dim uri As String
        uri = destino
        Dim peticionFTP As FtpWebRequest
        ' Creamos una peticion FTP con la dirección del fichero que vamos a subir
        peticionFTP = CType(FtpWebRequest.Create(New Uri(destino)), FtpWebRequest)
        ' Fijamos el usuario y la contraseña de la petición
        peticionFTP.Credentials = New NetworkCredential(user, pass)
        peticionFTP.KeepAlive = False
        peticionFTP.UsePassive = False
        ' Seleccionamos el comando que vamos a utilizar: Subir un fichero
        peticionFTP.Method = WebRequestMethods.Ftp.UploadFile
        ' Especificamos el tipo de transferencia de datos
        peticionFTP.UseBinary = True
        ' Informamos al servidor sobre el tamaño del fichero que vamos a subir
        peticionFTP.ContentLength = infoFichero.Length
        ' Fijamos un buffer de 150KB
        Dim longitudBuffer As Integer
        longitudBuffer = 153600
        Dim lector As Byte() = New Byte(153600) {}
        Dim num As Integer
        ' Abrimos el fichero para subirlo
        Dim fs As FileStream
        fs = infoFichero.OpenRead()
        Try
            Dim escritor As Stream
            escritor = peticionFTP.GetRequestStream()
            ' Leemos 150 KB del fichero en cada iteración
            num = fs.Read(lector, 0, longitudBuffer)
            While (num <> 0)
                ' Escribimos el contenido del flujo de lectura en el
                ' flujo de escritura del comando FTP
                escritor.Write(lector, 0, num)
                num = fs.Read(lector, 0, longitudBuffer)
            End While
            escritor.Close()
            fs.Close()
            ' Si todo ha ido bien, se devolverá String.Empty
            Return String.Empty
        Catch ex As Exception
            ' Si se produce algún fallo, se devolverá el mensaje del error
            Return ex.Message
        End Try
    End Function
    Public Function subirFicheroPdf_G() As String
        Dim fichero As String = ""
        Dim destino As String = ""
        Dim user As String = "colaveco"
        Dim pass As String = "NUEVA**!!COL22"
        If tipoinforme = 1 Then
            crea_control_lechero_com()
            fichero = "\\192.168.1.10\e\NET\CONTROL_LECHERO\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/control_lechero/" & idficha & ".pdf"
        ElseIf tipoinforme = 3 Then
            crea_agua_com()
            fichero = "\\192.168.1.10\e\NET\AGUA\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/agua/" & idficha & ".pdf"
        ElseIf tipoinforme = 4 Then
            crea_antibiograma_com()
            fichero = "\\192.168.1.10\e\NET\ANTIBIOGRAMA\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/antibiograma/" & idficha & ".pdf"
        ElseIf tipoinforme = 6 Then
            crea_parasitologia_com()
            fichero = "\\192.168.1.10\e\NET\PARASITOLOGIA\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/parasitologia/" & idficha & ".pdf"
        ElseIf tipoinforme = 7 Then
            crea_productos_subproductos_com()
            fichero = "\\192.168.1.10\e\NET\ALIMENTOS\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/productos_subproductos/" & idficha & ".pdf"
        ElseIf tipoinforme = 8 Then
            crea_serologia_com()
            fichero = "\\192.168.1.10\e\NET\SEROLOGIA\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/serologia/" & idficha & ".pdf"
        ElseIf tipoinforme = 9 Then
            crea_patologia_com()
            fichero = "\\192.168.1.10\e\NET\PATOLOGIA\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/patologia/" & idficha & ".pdf"
        ElseIf tipoinforme = 10 Then
            crea_calidad_de_leche_com()
            fichero = "\\192.168.1.10\e\NET\CALIDAD\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/calidad_de_leche/" & idficha & ".pdf"
        ElseIf tipoinforme = 11 Then
            crea_ambiental_com()
            fichero = "\\192.168.1.10\e\NET\AMBIENTAL\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/ambiental/" & idficha & ".pdf"
        ElseIf tipoinforme = 13 Then
            crea_agro_nutricion_com()
            fichero = "\\192.168.1.10\e\NET\NUTRICION\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/agro_nutricion/" & idficha & ".pdf"
        ElseIf tipoinforme = 14 Then
            crea_agro_suelos_com()
            fichero = "\\192.168.1.10\e\NET\Suelos\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/agro_suelos/" & idficha & ".pdf"
        ElseIf tipoinforme = 15 Then
            crea_brucelosis_leche_com()
            fichero = "\\192.168.1.10\e\NET\Brucelosis en leche\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/brucelosis_leche/" & idficha & ".pdf"
        ElseIf tipoinforme = 16 Then
            crea_efluentes_com()
            fichero = "\\192.168.1.10\e\NET\Efluentes\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/efluentes/" & idficha & ".pdf"
        ElseIf tipoinforme = 17 Then
            crea_antibiograma_com()
            fichero = "\\192.168.1.10\e\NET\ANTIBIOGRAMA\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/antibiograma/" & idficha & ".pdf"
        ElseIf tipoinforme = 18 Then
            crea_antibiograma_com()
            fichero = "\\192.168.1.10\e\NET\ANTIBIOGRAMA\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/antibiograma/" & idficha & ".pdf"
        ElseIf tipoinforme = 19 Then
            crea_agro_suelos_com()
            fichero = "\\192.168.1.10\e\NET\Suelos\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/agro_suelos/" & idficha & ".pdf"
        ElseIf tipoinforme = 20 Then
            crea_patologia_com()
            fichero = "\\192.168.1.10\e\NET\PATOLOGIA\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/patologia/" & idficha & ".pdf"
        ElseIf tipoinforme = 99 Then
            crea_otros_servicios_com()
            fichero = "\\192.168.1.10\e\NET\OTROS\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/otros_servicios/" & idficha & ".pdf"
        End If
        Dim infoFichero As New FileInfo(fichero)
        Dim uri As String
        uri = destino
        Dim peticionFTP As FtpWebRequest
        ' Creamos una peticion FTP con la dirección del fichero que vamos a subir
        peticionFTP = CType(FtpWebRequest.Create(New Uri(destino)), FtpWebRequest)
        ' Fijamos el usuario y la contraseña de la petición
        peticionFTP.Credentials = New NetworkCredential(user, pass)
        peticionFTP.KeepAlive = False
        peticionFTP.UsePassive = False
        ' Seleccionamos el comando que vamos a utilizar: Subir un fichero
        peticionFTP.Method = WebRequestMethods.Ftp.UploadFile
        ' Especificamos el tipo de transferencia de datos
        peticionFTP.UseBinary = True
        ' Informamos al servidor sobre el tamaño del fichero que vamos a subir
        'peticionFTP.ContentLength = infoFichero.Length
        '**********************************************************************
        Try
            peticionFTP.ContentLength = infoFichero.Length
        Catch ex As Exception
            ' Si se produce algún fallo, se devolverá el mensaje del error
            Return ex.Message
        End Try
        '*********************************************************************
        ' Fijamos un buffer de 150KB
        Dim longitudBuffer As Integer
        longitudBuffer = 153600
        Dim lector As Byte() = New Byte(153600) {}
        Dim num As Integer
        ' Abrimos el fichero para subirlo
        Dim fs As FileStream
        fs = infoFichero.OpenRead()
        Try
            Dim escritor As Stream
            escritor = peticionFTP.GetRequestStream()
            ' Leemos 150 KB del fichero en cada iteración
            num = fs.Read(lector, 0, longitudBuffer)
            While (num <> 0)
                ' Escribimos el contenido del flujo de lectura en el
                ' flujo de escritura del comando FTP
                escritor.Write(lector, 0, num)
                num = fs.Read(lector, 0, longitudBuffer)
            End While
            escritor.Close()
            fs.Close()
            ' Si todo ha ido bien, se devolverá String.Empty
            Return String.Empty
        Catch ex As Exception
            ' Si se produce algún fallo, se devolverá el mensaje del error
            Return ex.Message
        End Try
    End Function
    Public Function subirFicheroCsv_G() As String
        Dim fichero As String = ""
        Dim destino As String = ""
        Dim user As String = "colaveco"
        Dim pass As String = "NUEVA**!!COL22"
        If tipoinforme = 1 Then
            fichero = "\\192.168.1.10\e\NET\CONTROL_LECHERO\" & idficha & ".txt"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/control_lechero/" & idficha & ".txt"
        End If
        Dim infoFichero As New FileInfo(fichero)
        Dim uri As String
        uri = destino
        Dim peticionFTP As FtpWebRequest
        ' Creamos una peticion FTP con la dirección del fichero que vamos a subir
        peticionFTP = CType(FtpWebRequest.Create(New Uri(destino)), FtpWebRequest)
        ' Fijamos el usuario y la contraseña de la petición
        peticionFTP.Credentials = New NetworkCredential(user, pass)
        peticionFTP.KeepAlive = False
        peticionFTP.UsePassive = False
        ' Seleccionamos el comando que vamos a utilizar: Subir un fichero
        peticionFTP.Method = WebRequestMethods.Ftp.UploadFile
        ' Especificamos el tipo de transferencia de datos
        peticionFTP.UseBinary = True
        ' Informamos al servidor sobre el tamaño del fichero que vamos a subir
        peticionFTP.ContentLength = infoFichero.Length
        ' Fijamos un buffer de 150KB
        Dim longitudBuffer As Integer
        longitudBuffer = 153600
        Dim lector As Byte() = New Byte(153600) {}
        Dim num As Integer
        ' Abrimos el fichero para subirlo
        Dim fs As FileStream
        fs = infoFichero.OpenRead()
        Try
            Dim escritor As Stream
            escritor = peticionFTP.GetRequestStream()
            ' Leemos 150 KB del fichero en cada iteración
            num = fs.Read(lector, 0, longitudBuffer)
            While (num <> 0)
                ' Escribimos el contenido del flujo de lectura en el
                ' flujo de escritura del comando FTP
                escritor.Write(lector, 0, num)
                num = fs.Read(lector, 0, longitudBuffer)
            End While
            escritor.Close()
            fs.Close()
            ' Si todo ha ido bien, se devolverá String.Empty
            Return String.Empty
        Catch ex As Exception
            ' Si se produce algún fallo, se devolverá el mensaje del error
            Return ex.Message
        End Try
    End Function
    Private Sub revisar_cuentas_corrientes()
        Dim sv As New dSinVisualizacion
        Dim cliente As Long = 0
        Dim lista As New ArrayList
        'Dim idficha As Long = 0
        lista = sv.listarsv
        If Not lista Is Nothing Then
            For Each sv In lista
                barra = barra + 1
                ProgressBar1.Value = barra
                If barra = 100 Then
                    barra = 0
                End If
                Dim sa As New dSolicitudAnalisis
                sa.ID = sv.FICHA
                sa = sa.buscar
                If Not sa Is Nothing Then
                    cliente = sa.IDPRODUCTOR
                    idficha = sa.ID

                End If
                '*** VERIFICA SI EL CLIENTE TIENE DEUDA ***************************************
                '****SI EL CLIENTE ES CONTADO*******************************************
                Dim cli As New dCliente
                cli.ID = cliente
                cli = cli.buscar
                If Not cli Is Nothing Then
                    If cli.PROLESA = 0 Then
                        If cli.FAC_CONTADO = 1 Then
                            Dim f As New dFacturacion
                            Dim listaF As New ArrayList
                            listaF = f.listarxficha(idficha)
                            If Not listaF Is Nothing Then
                                For Each f In listaF
                                    If f.FACTURA <> 0 And f.FACTURA <> 999999 Then
                                        Dim mc As New dMovCte
                                        mc.MCCCMP = f.FACTURA
                                        mc = mc.buscarxcomprobante
                                        If Not mc Is Nothing Then
                                            If mc.MCCPAG >= mc.MCCIMP Then
                                                Dim fecha As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                                                Dim fec As String
                                                fec = Format(fecha, "yyyy-MM-dd")
                                                Dim c As New dSinVisualizacion
                                                c.FICHA = idficha
                                                c = c.buscarxficha
                                                Dim abonado As Integer = 0
                                                c.marcarvisualizacion(fec)
                                                abonado = 2
                                                _abonado = 2
                                                marcarweb(idficha, abonado)
                                                modificarRegistro2(_abonado)
                                            Else
                                                Dim diferencia As Double = 0
                                                diferencia = mc.MCCIMP - mc.MCCPAG
                                                If diferencia < 50 Then
                                                    Dim fecha As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                                                    Dim fec As String
                                                    fec = Format(fecha, "yyyy-MM-dd")
                                                    Dim c As New dSinVisualizacion
                                                    c.FICHA = idficha
                                                    c = c.buscarxficha
                                                    Dim abonado As Integer = 0
                                                    c.marcarvisualizacion(fec)
                                                    abonado = 2
                                                    _abonado = 2
                                                    marcarweb(idficha, abonado)
                                                    modificarRegistro2(_abonado)
                                                End If
                                            End If
                                        Else
                                            Dim fecha As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                                            Dim fec As String
                                            fec = Format(fecha, "yyyy-MM-dd")
                                            Dim c As New dSinVisualizacion
                                            c.FICHA = idficha
                                            c = c.buscarxficha
                                            Dim abonado As Integer = 0
                                            c.marcarvisualizacion(fec)
                                            abonado = 2
                                            _abonado = 2
                                            marcarweb(idficha, abonado)
                                            modificarRegistro2(_abonado)
                                        End If
                                    End If
                                Next
                            End If
                        Else
                            'VERIFICA SI EL CLIENTE FACTURA A OTRO CLIENTE
                            Dim cl As New dClient
                            cl.CLICOD = cliente
                            cl = cl.buscarxcli
                            If Not cl Is Nothing Then
                                If cl.CLISCT <> 0 Then
                                    cliente = cl.CLISCT
                                End If
                            End If
                            '*************************************************
                            Dim mc As New dMovCte
                            Dim listamc As New ArrayList
                            Dim fechaactual As Date = Now.ToString("yyyy-MM-dd")
                            Dim fechaact As String = Format(fechaactual, "yyyy-MM-dd")
                            Dim vencido As Integer = 0
                            listamc = mc.listarxcli(cliente)
                            If Not listamc Is Nothing Then
                                For Each mc In listamc
                                    Dim fechavto As Date = mc.MCCVTO
                                    Dim fecvto As String = Format(fechavto, "yyyy-MM-dd")
                                    If fecvto < fechaact Then
                                        If mc.MCCPAG < mc.MCCIMP Then
                                            Dim diferencia As Double = 0
                                            diferencia = mc.MCCIMP - mc.MCCPAG
                                            If diferencia > 100 Then
                                                vencido = 1
                                            End If
                                        End If
                                    End If
                                Next
                            End If
                            If vencido = 0 Then
                                Dim fecha As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                                Dim fec As String
                                fec = Format(fecha, "yyyy-MM-dd")
                                Dim c As New dSinVisualizacion
                                c.FICHA = idficha
                                c = c.buscarxficha
                                Dim abonado As Integer = 0
                                c.marcarvisualizacion(fec)
                                abonado = 1
                                _abonado = 1
                                marcarweb(idficha, abonado)
                                modificarRegistro2(_abonado)
                            End If
                        End If
                    Else
                        Dim f As New dFacturacion
                        Dim listaF As New ArrayList
                        Dim _abonado As Integer = 0
                        listaF = f.listarxficha(idficha)
                        If Not listaF Is Nothing Then
                            For Each f In listaF
                                If f.FACTURA <> 0 And f.FACTURA <> 999999 Then
                                    _abonado = 1
                                End If
                            Next
                            If _abonado = 1 Then
                                Dim fecha As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                                Dim fec As String
                                fec = Format(fecha, "yyyy-MM-dd")
                                Dim c As New dSinVisualizacion
                                c.FICHA = idficha
                                c = c.buscarxficha
                                Dim abonado As Integer = 0
                                c.marcarvisualizacion(fec)
                                abonado = 2
                                _abonado = 2
                                marcarweb(idficha, abonado)
                                modificarRegistro2(_abonado)
                            End If
                        End If
                    End If
                    '*******************************************************************************
                End If
            Next
        End If
    End Sub
    Private Sub marcarweb(ByVal ficha As Long, ByVal abonado As Integer)
        Dim _ficha As Long = ficha
        Dim _abonado As Integer = abonado
        Dim sa As New dSolicitudAnalisis
        sa.ID = _ficha
        sa = sa.buscar
        Dim tipoinforme As Integer = 0
        If Not sa Is Nothing Then
            tipoinforme = sa.IDTIPOINFORME
        End If
        If tipoinforme = 1 Then
            Dim clw As New dControlLecheroWeb_com
            clw.modificarabonado(ficha, abonado)
            clw = Nothing
        ElseIf tipoinforme = 3 Then
            Dim aw As New dAguaWeb_com
            aw.modificarabonado(ficha, abonado)
            aw = Nothing
        ElseIf tipoinforme = 4 Then
            Dim atbw As New dAntibiogramaWeb_com
            atbw.modificarabonado(ficha, abonado)
            atbw = Nothing
        ElseIf tipoinforme = 5 Then
            Dim palw As New dPalWeb_com
            palw.modificarabonado(ficha, abonado)
            palw = Nothing
        ElseIf tipoinforme = 6 Then
            Dim pstlgw As New dParasitologiaWeb_com
            pstlgw.modificarabonado(ficha, abonado)
            pstlgw = Nothing
        ElseIf tipoinforme = 7 Then
            Dim spw As New dSubproductosWeb_com
            spw.modificarabonado(ficha, abonado)
            spw = Nothing
        ElseIf tipoinforme = 8 Then
            Dim sw As New dSerologiaWeb_com
            sw.modificarabonado(ficha, abonado)
            sw = Nothing
        ElseIf tipoinforme = 9 Then
            Dim ptlgw As New dPatologiaWeb_com
            ptlgw.modificarabonado(ficha, abonado)
            ptlgw = Nothing
        ElseIf tipoinforme = 10 Then
            Dim calw As New dCalidadWeb_com
            calw.modificarabonado(ficha, abonado)
            calw = Nothing
        ElseIf tipoinforme = 11 Then
            Dim ambw As New dAmbientalWeb_com
            ambw.modificarabonado(ficha, abonado)
            ambw = Nothing
        ElseIf tipoinforme = 12 Then
            Dim lw As New dLactometrosWeb_com
            lw.modificarabonado(ficha, abonado)
            lw = Nothing
        ElseIf tipoinforme = 13 Then
            Dim anw As New dAgroNutricionWeb_com
            anw.modificarabonado(ficha, abonado)
            anw = Nothing
        ElseIf tipoinforme = 14 Then
            Dim asw As New dAgroSuelosWeb_com
            asw.modificarabonado(ficha, abonado)
            asw = Nothing
        ElseIf tipoinforme = 15 Then
            Dim blw As New dBrucelosisLecheWeb_com
            blw.modificarabonado(ficha, abonado)
            blw = Nothing
        ElseIf tipoinforme = 99 Then
            Dim ow As New dOtrosServiciosWeb_com
            ow.modificarabonado(ficha, abonado)
            ow = Nothing
        End If
    End Sub
    Public Function existeXls_G() As Boolean
        Dim fichero As String = ""
        Dim destino As String = ""
        Dim user As String = "colaveco"
        Dim pass As String = "NUEVA**!!COL22"
        If tipoinforme = 1 Then
            fichero = "\\192.168.1.10\e\NET\CONTROL_LECHERO\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/control_lechero/" & idficha & ".xls"
        ElseIf tipoinforme = 3 Then
            fichero = "\\192.168.1.10\e\NET\AGUA\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/agua/" & idficha & ".xls"
        ElseIf tipoinforme = 4 Then
            fichero = "\\192.168.1.10\e\NET\ANTIBIOGRAMA\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/antibiograma/" & idficha & ".xls"
        ElseIf tipoinforme = 5 Then
            fichero = "\\192.168.1.10\e\NET\PAL\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy.uy/www/gestor/data_file/" & idproductorweb_com & "/pal/" & idficha & ".xls"
        ElseIf tipoinforme = 6 Then
            fichero = "\\192.168.1.10\e\NET\PARASITOLOGIA\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/parasitologia/" & idficha & ".xls"
        ElseIf tipoinforme = 7 Then
            fichero = "\\192.168.1.10\e\NET\ALIMENTOS\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/productos_subproductos/" & idficha & ".xls"
        ElseIf tipoinforme = 8 Then
            fichero = "\\192.168.1.10\e\NET\SEROLOGIA\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/serologia/" & idficha & ".xls"
        ElseIf tipoinforme = 9 Then
            fichero = "\\192.168.1.10\e\NET\PATOLOGIA\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/patologia/" & idficha & ".xls"
        ElseIf tipoinforme = 10 Then
            fichero = "\\192.168.1.10\e\NET\CALIDAD\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/calidad_de_leche/" & idficha & ".xls"
        ElseIf tipoinforme = 11 Then
            fichero = "\\192.168.1.10\e\NET\AMBIENTAL\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/ambiental/" & idficha & ".xls"
        ElseIf tipoinforme = 12 Then
            fichero = "\\192.168.1.10\e\NET\LACTOMETROS\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/lactometros_chequeos_maquina/" & idficha & ".xls"
        ElseIf tipoinforme = 13 Then
            fichero = "\\192.168.1.10\e\NET\NUTRICION\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/agro_nutricion/" & idficha & ".xls"
        ElseIf tipoinforme = 14 Then
            fichero = "\\192.168.1.10\e\NET\Suelos\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/agro_suelos/" & idficha & ".xls"
        ElseIf tipoinforme = 15 Then
            fichero = "\\192.168.1.10\e\NET\Brucelosis en leche\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/brucelosis_leche/" & idficha & ".xls"
        ElseIf tipoinforme = 16 Then
            fichero = "\\192.168.1.10\e\NET\Efluentes\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/efluentes/" & idficha & ".xls"
        ElseIf tipoinforme = 99 Then
            fichero = "\\192.168.1.10\e\NET\OTROS\" & idficha & ".xls"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/otros_servicios/" & idficha & ".xls"
        End If
        Dim peticionFTP As FtpWebRequest
        ' Creamos una peticion FTP con la dirección del objeto que queremos saber si existe
        peticionFTP = CType(WebRequest.Create(New Uri(destino)), FtpWebRequest)
        ' Fijamos el usuario y la contraseña de la petición
        peticionFTP.Credentials = New NetworkCredential(user, pass)
        ' Para saber si el objeto existe, solicitamos la fecha de creación del mismo
        peticionFTP.Method = WebRequestMethods.Ftp.GetDateTimestamp
        peticionFTP.UsePassive = False
        Try
            ' Si el objeto existe, se devolverá True
            Dim respuestaFTP As FtpWebResponse
            respuestaFTP = CType(peticionFTP.GetResponse(), FtpWebResponse)
            excel = 0
            Return True
        Catch ex As Exception
            mensaje = mensaje & " excel(com) - "
            excel = 1
            ' Si el objeto no existe, se producirá un error y al entrar por el Catch
            ' se devolverá falso
            Return False
        End Try
    End Function
    Public Function existePdf_G() As Boolean
        Dim fichero As String = ""
        Dim destino As String = ""
        Dim user As String = "colaveco"
        Dim pass As String = "NUEVA**!!COL22"
        If tipoinforme = 1 Then
            fichero = "\\192.168.1.10\e\NET\CONTROL_LECHERO\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/control_lechero/" & idficha & ".pdf"
        ElseIf tipoinforme = 3 Then
            fichero = "\\192.168.1.10\e\NET\AGUA\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/agua/" & idficha & ".pdf"
        ElseIf tipoinforme = 4 Then
            fichero = "\\192.168.1.10\e\NET\ANTIBIOGRAMA\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/antibiograma/" & idficha & ".pdf"
        ElseIf tipoinforme = 5 Then
            fichero = "\\192.168.1.10\e\NET\PAL\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/pal/" & idficha & ".pdf"
        ElseIf tipoinforme = 6 Then
            fichero = "\\192.168.1.10\e\NET\PARASITOLOGIA\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/parasitologia/" & idficha & ".pdf"
        ElseIf tipoinforme = 7 Then
            fichero = "\\192.168.1.10\e\NET\ALIMENTOS\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/productos_subproductos/" & idficha & ".pdf"
        ElseIf tipoinforme = 8 Then
            fichero = "\\192.168.1.10\e\NET\SEROLOGIA\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/serologia/" & idficha & ".pdf"
        ElseIf tipoinforme = 9 Then
            fichero = "\\192.168.1.10\e\NET\PATOLOGIA\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/patologia/" & idficha & ".pdf"
        ElseIf tipoinforme = 10 Then
            fichero = "\\192.168.1.10\e\NET\CALIDAD\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/calidad_de_leche/" & idficha & ".pdf"
        ElseIf tipoinforme = 11 Then
            fichero = "\\192.168.1.10\e\NET\AMBIENTAL\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/ambiental/" & idficha & ".pdf"
        ElseIf tipoinforme = 12 Then
            fichero = "\\192.168.1.10\e\NET\LACTOMETROS\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/lactometros_chequeos_maquina/" & idficha & ".pdf"
        ElseIf tipoinforme = 13 Then
            fichero = "\\192.168.1.10\e\NET\NUTRICION\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/agro_nutricion/" & idficha & ".pdf"
        ElseIf tipoinforme = 14 Then
            fichero = "\\192.168.1.10\e\NET\Suelos\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/agro_suelos/" & idficha & ".pdf"
        ElseIf tipoinforme = 15 Then
            fichero = "\\192.168.1.10\e\NET\Brucelosis en leche\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/brucelosis_leche/" & idficha & ".pdf"
        ElseIf tipoinforme = 16 Then
            fichero = "\\192.168.1.10\e\NET\Efluentes\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/efluentes/" & idficha & ".pdf"
        ElseIf tipoinforme = 99 Then
            fichero = "\\192.168.1.10\e\NET\OTROS\" & idficha & ".pdf"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/otros_servicios/" & idficha & ".pdf"
        End If
        Dim peticionFTP As FtpWebRequest
        ' Creamos una peticion FTP con la dirección del objeto que queremos saber si existe
        peticionFTP = CType(WebRequest.Create(New Uri(destino)), FtpWebRequest)
        ' Fijamos el usuario y la contraseña de la petición
        peticionFTP.Credentials = New NetworkCredential(user, pass)
        ' Para saber si el objeto existe, solicitamos la fecha de creación del mismo
        peticionFTP.Method = WebRequestMethods.Ftp.GetDateTimestamp
        peticionFTP.UsePassive = False
        Try
            ' Si el objeto existe, se devolverá True
            Dim respuestaFTP As FtpWebResponse
            respuestaFTP = CType(peticionFTP.GetResponse(), FtpWebResponse)
            pdf = 0
            Return True
        Catch ex As Exception
            mensaje = mensaje & " pdf(com) - "
            pdf = 1
            ' Si el objeto no existe, se producirá un error y al entrar por el Catch
            ' se devolverá falso
            Return False
        End Try
    End Function
    Public Function existeCsv_G() As Boolean
        Dim fichero As String = ""
        Dim destino As String = ""
        Dim user As String = "colaveco"
        Dim pass As String = "NUEVA**!!COL22"
        If tipoinforme = 1 Then
            fichero = "\\192.168.1.10\e\NET\CONTROL_LECHERO\" & idficha & ".txt"
            destino = "ftp://colaveco.com.uy/public_html/gestor/data_file/" & idproductorweb_com & "/control_lechero/" & idficha & ".txt"
        End If
        Dim peticionFTP As FtpWebRequest
        ' Creamos una peticion FTP con la dirección del objeto que queremos saber si existe
        peticionFTP = CType(WebRequest.Create(New Uri(destino)), FtpWebRequest)
        ' Fijamos el usuario y la contraseña de la petición
        peticionFTP.Credentials = New NetworkCredential(user, pass)
        ' Para saber si el objeto existe, solicitamos la fecha de creación del mismo
        peticionFTP.Method = WebRequestMethods.Ftp.GetDateTimestamp
        peticionFTP.UsePassive = False
        Try
            ' Si el objeto existe, se devolverá True
            Dim respuestaFTP As FtpWebResponse
            respuestaFTP = CType(peticionFTP.GetResponse(), FtpWebResponse)
            csv = 0
            Return True
        Catch ex As Exception
            mensaje = mensaje & " csv(com) - "
            csv = 1
            ' Si el objeto no existe, se producirá un error y al entrar por el Catch
            ' se devolverá falso
            Return False
        End Try
    End Function

#End Region

#Region "SALDOS"

    Private Sub calcularsaldos()
        Dim s As New dSaldos
        s.eliminartodos()
        Dim m As New dMovCte
        Dim c As New dCliente
        Dim idcliente As Long = 0
        Dim debe As Double = 0
        Dim haber As Double = 0
        Dim saldo As Double = 0
        Dim listaclientes As New ArrayList
        Dim listamovimientos As New ArrayList
        listaclientes = c.listar
        If Not listaclientes Is Nothing Then
            For Each c In listaclientes
                idcliente = c.ID
                listamovimientos = m.saldosxcli(idcliente)
                If Not listamovimientos Is Nothing Then
                    For Each m In listamovimientos
                        If m.MCCCOD = 1 Then
                            debe = debe + m.MCCIMP
                        ElseIf m.MCCCOD = 2 Then
                            haber = haber + m.MCCIMP
                        End If
                    Next
                    saldo = debe - haber
                End If
                s.IDCLIENTE = idcliente
                s.SALDO = saldo
                saldo = 0
                debe = 0
                haber = 0
                s.guardar()
            Next
        End If
        subirsaldos()
    End Sub
    Private Sub subirsaldos()
        Dim s As New dSaldos
        Dim c As New dCliente
        Dim cw As New dProductorWeb_com
        Dim usuario As String = ""
        Dim saldo As Double = 0
        Dim listasaldos As New ArrayList
        listasaldos = s.listar
        If Not listasaldos Is Nothing Then
            For Each s In listasaldos
                c.ID = s.IDCLIENTE
                c = c.buscar
                If Not c Is Nothing Then
                    usuario = c.USUARIO_WEB
                End If
                saldo = s.SALDO
                cw.USUARIO = usuario
                cw.SALDO_PESOS = saldo
                cw.actualizarsaldo()
            Next
        End If
    End Sub

#End Region

#Region "ENVIO-EMAILS"

    Private Sub enviomail()
        Dim _Message As New System.Net.Mail.MailMessage()
        Dim _SMTP As New System.Net.Mail.SmtpClient
        Dim sa As New dSolicitudAnalisis
        Dim p As New dCliente
        Dim ti As New dTipoInforme
        Dim nombre_productor As String = ""
        Dim tipo_analisis As String = ""
        nficha = idficha
        sa.ID = nficha
        sa = sa.buscar
        If Not sa Is Nothing Then
            p.ID = sa.IDPRODUCTOR
            p = p.buscar
            If Not p Is Nothing Then
                nombre_productor = p.NOMBRE
            End If
            ti.ID = sa.IDTIPOINFORME
            ti = ti.buscar
            If Not ti Is Nothing Then
                tipo_analisis = ti.NOMBRE
            End If
        End If
        Dim texto As String = ""
        texto = "Nos es grato comunicarle que el informe Nº " & " " & nficha & " - " & tipo_analisis & " (" & nombre_productor & ")," & "se encuentra disponible en la web de Colaveco." & vbCrLf _
            & "Para poder acceder a los resultados debe ir a www.colaveco.com.uy/gestor y digitar su usuario y contraseña." & vbCrLf _
            & "Sino cuenta con usuario y contraseña, favor solicitarla en administración al correo electrónico colaveco@gmail.com o al teléfono 4554 5311." & vbCrLf _
            & "Agradecemos su confianza y quedamos a sus órdenes." & vbCrLf & vbCrLf _
            & "Sin mas, saluda muy atte." & vbCrLf & vbCrLf _
            & "Administración - COLAVECO"
        If email <> "" And email <> "no aportado" Then
            'CONFIGURACIÓN DEL STMP 
            _SMTP.Credentials = New System.Net.NetworkCredential("notificaciones@colaveco.com.uy", "19912021Notificaciones")
            _SMTP.Host = "170.249.199.66"
            _SMTP.Port = 25
            _SMTP.EnableSsl = False
            _Message.From = New System.Net.Mail.MailAddress("notificaciones@colaveco.com.uy", "COLAVECO", System.Text.Encoding.UTF8)

            ' CONFIGURACION DEL MENSAJE 
            Try
                _Message.[To].Add(email)
                _Message.[To].Add("envios@colaveco.com.uy")
            Catch ex As System.Net.Mail.SmtpException ' MessageBox.Show(ex.ToString) 
            End Try
            'Cuenta de Correo al que se le quiere enviar el e-mail 
            _Message.From = New System.Net.Mail.MailAddress("colaveco@gmail.com", "COLAVECO", System.Text.Encoding.UTF8)
            'Quien lo envía 
            _Message.Subject = "Informe" & " Nº " & nficha & " - Colaveco"
            'Sujeto del e-mail 
            _Message.SubjectEncoding = System.Text.Encoding.UTF8
            'Codificacion 
            _Message.Body = texto
            'contenido del mail 
            _Message.BodyEncoding = System.Text.Encoding.UTF8 '
            _Message.Priority = System.Net.Mail.MailPriority.Normal
            _Message.IsBodyHtml = False
            ' ADICION DE DATOS ADJUNTOS ‘
            'Dim _File As String = My.Application.Info.DirectoryPath & "archivo" 'archivo que se quiere adjuntar ‘
            'Dim _Attachment As New System.Net.Mail.Attachment(_File, System.Net.Mime.MediaTypeNames.Application.Octet) '
            '_Message.Attachments.Add(_Attachment) 'ENVIO 
            Try
                _SMTP.Send(_Message)
                'MessageBox.Show("Correo enviado!", "Correo", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As System.Net.Mail.SmtpException ' MessageBox.Show(ex.ToString) 
            End Try
        End If
        email = ""
        nficha = 0
    End Sub
    Private Sub enviosms()
        Dim num1 As String = ""
        Dim num2 As String = ""
        Dim email1 As String = ""
        Dim email2 As String = ""
        Dim sms As String = ""
        Dim sms1 As String = ""
        Dim sms2 As String = ""
        Dim cel1 As String = ""
        Dim cel2 As String = ""
        Dim largotexto As Integer = 0
        Dim celular1 As String = ""
        Dim celular2 As String = ""
        Dim _Message As New System.Net.Mail.MailMessage()
        Dim _SMTP As New System.Net.Mail.SmtpClient
        Dim texto As String = celular
        Dim cantcaracteres As Integer = Len(texto)
        If celular <> "" Then
            largotexto = celular.Length
        End If
        nficha = idficha
        Dim posicion As Integer
        Dim posicion1 As Integer
        Dim posicion2 As Integer
        posicion = InStr(celular, ",")
        If posicion > 0 Then
            posicion1 = posicion - 1
            posicion2 = posicion + 1
            cel1 = Mid(celular, 1, posicion1)
            cel2 = Mid(celular, posicion2, largotexto)
            If Mid(cel1, 1, 2) = "09" Then
                celular1 = cel1.Remove(0, 2)
            Else
                celular1 = cel1
            End If
            email = celular1
            num1 = Mid(celular1, 1, 1)
            If num1 = "9" Or num1 = "8" Or num1 = "1" Then
                'ancel es numero (sin 09 inicial + pin)
                sms1 = email & "@antelinfo.com.uy"
            ElseIf num1 = "3" Or num1 = "4" Or num1 = "5" Then
                'movistar es numero (sin 0 inicial + pin)
                If Mid(celular, 1, 1) = "0" Then
                    celular1 = celular.Remove(0, 1)
                End If
                email = celular1
                sms1 = email & "@sms.movistar.com.uy"
            ElseIf num1 = "6" Or num1 = "7" Then
                'claro es numero (sin 0 inicial sin pin)
                If Mid(celular, 1, 1) = "0" Then
                    celular2 = celular.Remove(0, 1)
                End If
                email = celular1
                sms1 = email & "@sms.ctimovil.com.uy"
            End If
            '*****************************************
            If Mid(cel2, 1, 2) = "09" Then
                celular2 = cel2.Remove(0, 2)
            Else
                celular2 = cel2
            End If
            email2 = celular2
            num2 = Mid(celular2, 1, 1)

            If num2 = "9" Or num2 = "8" Or num2 = "1" Then
                'ancel es numero (sin 09 inicial + pin)
                sms2 = email2 & "@antelinfo.com.uy"
            ElseIf num2 = "3" Or num2 = "4" Or num2 = "5" Then
                'movistar es numero (sin 0 inicial + pin)
                If Mid(celular2, 1, 1) = "0" Then
                    celular2 = celular2.Remove(0, 1)
                End If
                email2 = celular2
                sms2 = email2 & "@sms.movistar.com.uy"
            ElseIf num2 = "6" Or num2 = "7" Then
                'claro es numero (sin 0 inicial sin pin)
                If Mid(celular2, 1, 1) = "0" Then
                    celular2 = celular2.Remove(0, 1)
                End If
                email2 = celular2
                sms2 = email2 & "@sms.ctimovil.com.uy"
            End If
            sms = sms1 & "," & sms2
        Else
            'Dim celular As String = ""
            'celular = TextCelular1.Text.Trim
            nficha = idficha
            If Mid(celular, 1, 2) = "09" Then
                celular2 = celular.Remove(0, 2)
            Else
                celular2 = celular
            End If
            email = celular2
            num1 = Mid(celular2, 1, 1)
            If num1 = "9" Or num1 = "8" Or num1 = "1" Then
                'ancel es numero (sin 09 inicial + pin)
                sms = email & "@antelinfo.com.uy"
            ElseIf num1 = "3" Or num1 = "4" Or num1 = "5" Then
                'movistar es numero (sin 0 inicial + pin)
                If Mid(celular, 1, 1) = "0" Then
                    celular2 = celular.Remove(0, 1)
                End If
                email = celular2
                sms = email & "@sms.movistar.com.uy"
            ElseIf num1 = "6" Or num1 = "7" Then
                'claro es numero (sin 0 inicial sin pin)
                If Mid(celular, 1, 1) = "0" Then
                    celular2 = celular.Remove(0, 1)
                End If
                email = celular2
                sms = email & "@sms.ctimovil.com.uy"
            End If
        End If
        Dim sa As New dSolicitudAnalisis
        Dim p As New dCliente
        Dim ti As New dTipoInforme
        Dim nombre_productor As String = ""
        Dim tipo_analisis As String = ""
        nficha = idficha
        sa.ID = nficha
        sa = sa.buscar
        If Not sa Is Nothing Then
            p.ID = sa.IDPRODUCTOR
            p = p.buscar
            If Not p Is Nothing Then
                nombre_productor = p.NOMBRE
            End If
            ti.ID = sa.IDTIPOINFORME
            ti = ti.buscar
            If Not ti Is Nothing Then
                tipo_analisis = ti.NOMBRE
            End If
        End If
        If sms <> "" Then
            'CONFIGURACIÓN DEL STMP 
            _SMTP.Credentials = New System.Net.NetworkCredential("notificaciones@colaveco.com.uy", "19912021Notificaciones")
            _SMTP.Host = "170.249.199.66"
            _SMTP.Port = 25
            _SMTP.EnableSsl = False


            ' CONFIGURACION DEL MENSAJE 
            _Message.[To].Add(sms)
            _Message.[To].Add("envios@colaveco.com.uy")
            'Cuenta de Correo al que se le quiere enviar el e-mail 
            _Message.From = New System.Net.Mail.MailAddress("notificaciones@colaveco.com.uy", "COLAVECO", System.Text.Encoding.UTF8)
            'Quien lo envía 
            _Message.Subject = "El informe Nº " & " " & nficha & " - " & tipo_analisis & " (" & nombre_productor & ")," & "se ha subido a la web. Gracias."
            'Sujeto del e-mail 
            _Message.SubjectEncoding = System.Text.Encoding.UTF8
            'Codificacion 
            '_Message.Body = "Se han enviado las siguientes cajas:" & " " & ecaja1 & ", " & "por" & " " & eagencia & " " & "envío nº" & " " & eremito & ""
            '_Message.Body = "Colaveco ha publicado un informe. Ingrese al sitio http://www.colaveco.com.uy/gestor"
            'contenido del mail 
            _Message.BodyEncoding = System.Text.Encoding.UTF8 '
            _Message.Priority = System.Net.Mail.MailPriority.Normal
            _Message.IsBodyHtml = False
            ' ADICION DE DATOS ADJUNTOS ‘
            'Dim _File As String = My.Application.Info.DirectoryPath & "archivo" 'archivo que se quiere adjuntar ‘
            'Dim _Attachment As New System.Net.Mail.Attachment(_File, System.Net.Mime.MediaTypeNames.Application.Octet) '
            '_Message.Attachments.Add(_Attachment) 'ENVIO 
            Try
                _SMTP.Send(_Message)
                'MessageBox.Show("Mensaje enviado!", "Correo", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As System.Net.Mail.SmtpException ' MessageBox.Show(ex.ToString) 
            End Try
        End If
        email = ""
        texto = ""
    End Sub
    Private Sub enviaremail() 'ENVIA EMAIL A MGAP (BRUCELOSIS)
        Dim _Message As New System.Net.Mail.MailMessage()
        Dim _SMTP As New System.Net.Mail.SmtpClient
        Dim email As String = ""
        Dim destinatario As String = ""
        Dim archivo As String = ""
        archivo = idficha
        email = "unepi@mgap.gub.uy"
        If email <> "" And email <> "no aportado" Then
            'CONFIGURACIÓN DEL STMP 
            _SMTP.Credentials = New System.Net.NetworkCredential("notificaciones@colaveco.com.uy", "19912021Notificaciones")
            _SMTP.Host = "170.249.199.66"
            _SMTP.Port = 25
            _SMTP.EnableSsl = False
            _Message.[To].Add(email)
            _Message.[To].Add("envios@colaveco.com.uy")
            'Cuenta de Correo al que se le quiere enviar el e-mail 
            _Message.From = New System.Net.Mail.MailAddress("notificaciones@colaveco.com.uy", "COLAVECO", System.Text.Encoding.UTF8)
            '_Message.From = New System.Net.Mail.MailAddress("colaveco@gmail.com", "COLAVECO", System.Text.Encoding.UTF8)
            'Quien lo envía 
            _Message.Subject = "Colaveco - Brucelosis en leche"
            'Sujeto del e-mail 
            _Message.SubjectEncoding = System.Text.Encoding.UTF8
            'Codificacion 
            _Message.Body = "Adjuntamos informe de Brucelsois en leche."
            'contenido del mail 
            _Message.BodyEncoding = System.Text.Encoding.UTF8 '
            _Message.Priority = System.Net.Mail.MailPriority.Normal
            _Message.IsBodyHtml = False
            ' ADICION DE DATOS ADJUNTOS ‘
            'Dim _File As String = "\\192.168.1.10\e\NET\INFORMES PARA SUBIR\" & archivo & ".xls" 'archivo que se quiere adjuntar ‘
            Dim _File As String = ""
            If nombre_pc = "ROBOT" Then
                _File = "C:\INFORMES PARA SUBIR\" & archivo & ".pdf" 'archivo que se quiere adjuntar ‘
            Else
                _File = "\\ROBOT\\INFORMES PARA SUBIR\" & archivo & ".pdf" 'archivo que se quiere adjuntar ‘
            End If
            Dim _Attachment As New System.Net.Mail.Attachment(_File, System.Net.Mime.MediaTypeNames.Application.Octet) '
            _Message.Attachments.Add(_Attachment) 'ENVIO 
            Try
                _SMTP.Send(_Message)
                'MessageBox.Show("Correo enviado!", "Correo", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As System.Net.Mail.SmtpException ' MessageBox.Show(ex.ToString) 
            End Try
        End If
        email = ""
    End Sub
    Private Sub enviaremail2() 'ENVIA EMAIL A MGAP (BRUCELOSIS)
        Dim _Message As New System.Net.Mail.MailMessage()
        Dim _SMTP As New System.Net.Mail.SmtpClient
        Dim email As String = ""
        Dim destinatario As String = ""
        Dim archivo As String = ""
        archivo = idficha
        email = "decano@fvet.edu.uy"
        If email <> "" And email <> "no aportado" Then
            'CONFIGURACIÓN DEL STMP 
            _SMTP.Credentials = New System.Net.NetworkCredential("notificaciones@colaveco.com.uy", "19912021Notificaciones")
            _SMTP.Host = "170.249.199.66"
            _SMTP.Port = 25
            _SMTP.EnableSsl = False
            ' CONFIGURACION DEL MENSAJE 
            _Message.[To].Add(email)
            _Message.[To].Add("envios@colaveco.com.uy")
            'Cuenta de Correo al que se le quiere enviar el e-mail 
            _Message.From = New System.Net.Mail.MailAddress("notificaciones@colaveco.com.uy", "COLAVECO", System.Text.Encoding.UTF8)
            'Quien lo envía 
            _Message.Subject = "Colaveco - Brucelosis en leche"
            'Sujeto del e-mail 
            _Message.SubjectEncoding = System.Text.Encoding.UTF8
            'Codificacion 
            _Message.Body = "Adjuntamos informe de Brucelsois en leche."
            'contenido del mail 
            _Message.BodyEncoding = System.Text.Encoding.UTF8 '
            _Message.Priority = System.Net.Mail.MailPriority.Normal
            _Message.IsBodyHtml = False
            ' ADICION DE DATOS ADJUNTOS ‘
            'Dim _File As String = "\\192.168.1.10\e\NET\INFORMES PARA SUBIR\" & archivo & ".xls" 'archivo que se quiere adjuntar ‘
            Dim _File As String = ""
            If nombre_pc = "ROBOT" Then
                _File = "C:\INFORMES PARA SUBIR\" & archivo & ".pdf" 'archivo que se quiere adjuntar ‘
            Else
                _File = "\\ROBOT\\INFORMES PARA SUBIR\" & archivo & ".pdf" 'archivo que se quiere adjuntar ‘
            End If
            Dim _Attachment As New System.Net.Mail.Attachment(_File, System.Net.Mime.MediaTypeNames.Application.Octet) '
            _Message.Attachments.Add(_Attachment) 'ENVIO 
            Try
                _SMTP.Send(_Message)
                'MessageBox.Show("Correo enviado!", "Correo", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As System.Net.Mail.SmtpException ' MessageBox.Show(ex.ToString) 
            End Try
        End If
        email = ""
    End Sub
    Private Sub enviomail_sw()
        Dim _Message As New System.Net.Mail.MailMessage()
        Dim _SMTP As New System.Net.Mail.SmtpClient
        '******************************************************************************************************************************************
        Dim ficha As String = nficha
        Dim fecha As Date = DateFecha.Value
        Dim nmuestras As Integer
        nmuestras = nmuestrasweb
        Dim muestra As String = muestraweb
        Dim solicitud As String = ""
        Dim texto As String = ""
        Dim texto2 As String = ""
        Dim texto3 As String = ""
        Dim tipoinforme As String = tipoinformeweb
        Dim subtipoinforme As String = subtipoinformeweb
        Dim observaciones As String = observacionesweb
        Dim titulo As String = ""
        Dim enc_ficha As String = ""
        Dim enc_fecha As String = ""
        Dim enc_cliente As String = ""
        Dim enc_muestras As String = ""
        Dim enc_muestrade As String = ""
        Dim cuerpo_analisis As String = ""
        Dim cuerpo_muestras As String = ""
        Dim pie_observaciones As String = ""
        Dim pie_estadosolicitud As String = "En nuestro sitio web http://www.colaveco.com.uy/gestor, puede ver el estado de su solicitud."
        Dim nombre_productor As String = nombreproductorweb
        Dim sm As New dRelSolicitudMuestras
        Dim spal As New dSolicitudPAL
        Dim csm As New dCalidadSolicitudMuestra
        Dim cs As New dControlSolicitud
        Dim a2 As New dAntibiograma2
        Dim sn As New dSolicitudNutricion
        Dim ss As New dSolicitudSuelos
        Dim bl As New dBrucelosis
        Dim na As New dNuevoAnalisis
        Dim lista2 As New ArrayList
        Dim lista3 As New ArrayList
        Dim lista4 As New ArrayList
        Dim lista5 As New ArrayList
        Dim lista6 As New ArrayList
        Dim lista7 As New ArrayList
        Dim lista10 As New ArrayList
        Dim listabl As New ArrayList
        Dim listanutricion As New ArrayList
        Dim listasuelos As New ArrayList
        Dim listaefluentes As New ArrayList
        lista4 = sm.listarporficha(ficha)
        lista5 = csm.listarporsolicitud3(ficha)
        lista6 = cs.listarporsolicitud(ficha)
        lista7 = a2.listarporsolicitud(ficha)
        lista10 = spal.listarporsolicitud(ficha)
        listanutricion = sn.listarporsolicitud(ficha)
        listasuelos = ss.listarporsolicitud(ficha)
        listabl = sm.listarporficha(ficha)
        listaefluentes = na.listarporficha2(ficha)
        ' SI ES PRODUCTOS LÁCTEOS ********************************************************************************
        If tipoinforme = "Alimentos" Then
            Dim sp As New dSubproducto
            Dim lista As New ArrayList
            texto = ""
            lista = sp.listarporsolicitud(ficha)
            If Not lista Is Nothing Then
                If lista.Count > 0 Then
                    For Each sp In lista
                        If sp.ESTAFCOAGPOSITIVO = 1 Then
                            texto = texto + " - Estaf. Coag. Positivo"
                        End If
                        If sp.CF = 1 Then
                            texto = texto + " - CF"
                        End If
                        If sp.MOHOSYLEVADURAS = 1 Then
                            texto = texto + " - Mohos y levaduras"
                        End If
                        If sp.CT = 1 Then
                            texto = texto + " - Coliformes Totales"
                        End If
                        If sp.ECOLI = 1 Then
                            texto = texto + " - E. Coli"
                        End If
                        If sp.SALMONELLA = 1 Then
                            texto = texto + " - Salmonella"
                        End If
                        If sp.LISTERIASPP = 1 Then
                            texto = texto + " - Listeria spp"
                        End If
                        If sp.HUMEDAD = 1 Then
                            texto = texto + " - Humedad"
                        End If
                        If sp.MGRASA = 1 Then
                            texto = texto + " - M. Grasa"
                        End If
                        If sp.PH = 1 Then
                            texto = texto + " - pH"
                        End If
                        If sp.CLORUROS = 1 Then
                            texto = texto + " - Cloruros"
                        End If
                        If sp.PROTEINAS = 1 Then
                            texto = texto + " - Proteínas"
                        End If
                        If sp.ENTEROBACTERIAS = 1 Then
                            texto = texto + " - Enterobacterias"
                        End If
                        If sp.LISTERIAAMBIENTAL = 1 Then
                            texto = texto + " - Listeria Ambiental"
                        End If
                        If sp.ESPORANAERMESOFILO = 1 Then
                            texto = texto + " - Espor. Anaer. Mesófilos"
                        End If
                        If sp.TERMOFILOS = 1 Then
                            texto = texto + " - Termodúricos"
                        End If
                        If sp.PSICROTROFOS = 1 Then
                            texto = texto + " - Psicrótrofos"
                        End If
                        If sp.RB = 1 Then
                            texto = texto + " - RB"
                        End If
                        If sp.TABLANUTRICIONAL = 1 Then
                            texto = texto + " - Tabla nutricional"
                        End If
                        If sp.LISTERIAMONOCITOGENES = 1 Then
                            texto = texto + " - Listeria monocitógenes"
                        End If
                        If sp.CENIZAS = 1 Then
                            texto = texto + " - Cenizas"
                        End If
                    Next
                End If
            End If
            ' SI ES AGUA ********************************************************************************
        ElseIf tipoinforme = "Agua" Then
            Dim a1 As New dAgua
            texto = ""
            a1.ID = ficha
            a1 = a1.buscar()
            texto = subtipoinformeweb
            If Not a1 Is Nothing Then
                If a1.HET22 = 1 Then
                    texto = texto & " " & " - Heterotróficos 22"
                End If
                If a1.HET35 = 1 Then
                    texto = texto & " " & " - Heterotróficos 35"
                End If
                If a1.HET37 = 1 Then
                    texto = texto & " " & " - Heterotróficos 37"
                End If
                If a1.CLORO = 1 Then
                    texto = texto & " " & " - Cloro"
                End If
                If a1.CONDUCTIVIDAD = 1 Then
                    texto = texto & " " & " - Conductividad"
                End If
                If a1.PH = 1 Then
                    texto = texto & " " & " - pH"
                End If
                If a1.ECOLI = 1 Then
                    texto = texto & " " & " - Ecoli"
                End If
            End If
            ' SI ES CALIDAD DE LECHE ********************************************************************************
        ElseIf tipoinforme = "Calidad de leche" Then
            Dim rb As Integer = 0
            Dim rc As Integer = 0
            Dim comp As Integer = 0
            Dim criosc As Integer = 0
            Dim inh As Integer = 0
            Dim espor As Integer = 0
            Dim urea As Integer = 0
            Dim term As Integer = 0
            Dim psicr As Integer = 0
            Dim crioscopo As Integer = 0
            texto = ""
            If Not lista5 Is Nothing Then
                If lista5.Count > 0 Then
                    For Each csm In lista5
                        If csm.RB = 1 Then
                            rb = 1
                        End If
                        If csm.RC = 1 Then
                            rc = 1
                        End If
                        If csm.COMPOSICION = 1 Then
                            comp = 1
                        End If
                        If csm.CRIOSCOPIA = 1 Then
                            criosc = 1
                        End If
                        If csm.INHIBIDORES = 1 Then
                            inh = 1
                        End If
                        If csm.ESPORULADOS = 1 Then
                            espor = 1
                        End If
                        If csm.UREA = 1 Then
                            urea = 1
                        End If
                        If csm.TERMOFILOS = 1 Then
                            term = 1
                        End If
                        If csm.PSICROTROFOS = 1 Then
                            psicr = 1
                        End If
                        If csm.CRIOSCOPIA_CRIOSCOPO = 1 Then
                            crioscopo = 1
                        End If
                    Next
                End If
            End If
            If rb = 1 Then
                texto = texto + " - RB"
            End If
            If rc = 1 Then
                texto = texto + " - RC"
            End If
            If comp = 1 Then
                texto = texto + " - Composición"
            End If
            If criosc = 1 Then
                texto = texto + " - Crioscopía"
            End If
            If inh = 1 Then
                texto = texto + " - Inhibidores"
            End If
            If espor = 1 Then
                texto = texto + " - Esporulados"
            End If
            If urea = 1 Then
                texto = texto + " - Urea"
            End If
            If term = 1 Then
                texto = texto + " - Termófilos"
            End If
            If psicr = 1 Then
                texto = texto + " - Psicrótrofos"
            End If
            If crioscopo = 1 Then
                texto = texto + " - Crioscopía (crióscopo)"
            End If
            ' SI ES CONTROL LECHERO ********************************************************************************
        ElseIf tipoinforme = "Control Lechero" Then
            Dim rc As Integer = 0
            Dim comp As Integer = 0
            Dim urea As Integer = 0
            texto = ""
            If Not lista6 Is Nothing Then
                If lista6.Count > 0 Then
                    For Each cs In lista6
                        If cs.RC = 1 Then
                            rc = 1
                        End If
                        If cs.COMPOSICION = 1 Then
                            comp = 1
                        End If
                        If cs.UREA = 1 Then
                            urea = 1
                        End If
                    Next
                End If
            End If
            If rc = 1 Then
                texto = texto + " - RC"
            End If
            If comp = 1 Then
                texto = texto + " - Composición"
            End If
            If urea = 1 Then
                texto = texto + " - Urea"
            End If
            ' SI ANTIBIOGRAMA ********************************************************************************
        ElseIf tipoinforme = "Bacteriología y Antibiograma" Then
            Dim aislamiento As Integer = 0
            Dim antibiograma As Integer = 0
            texto = ""
            If Not lista7 Is Nothing Then
                If lista7.Count > 0 Then
                    For Each a2 In lista7
                        If a2.AISLAMIENTO = 1 Then
                            aislamiento = 1
                        End If
                        If a2.ANTIBIOGRAMA = 1 Then
                            antibiograma = 1
                        End If
                    Next
                End If
            End If
            If aislamiento = 1 Then
                texto = texto + " - Aislamiento"
            End If
            If antibiograma = 1 Then
                texto = texto + " - Antibiograma"
            End If
            ' SI ES AMBIENTAL ********************************************************************************
        ElseIf tipoinforme = "Ambiental" Then
            Dim ambs As New dAmbientalSolicitud
            Dim lista8 As ArrayList
            lista8 = ambs.listarporsolicitud(ficha)
            Dim enterobacterias As Integer = 0
            Dim listambiental As Integer = 0
            Dim listmono As Integer = 0
            Dim salmonella As Integer = 0
            Dim ecoli As Integer = 0
            Dim mohosylevaduras As Integer = 0
            Dim rb As Integer = 0
            Dim ct As Integer = 0
            Dim cf As Integer = 0
            Dim pseudomonaspp As Integer = 0
            texto = ""
            If Not lista8 Is Nothing Then
                If lista8.Count > 0 Then
                    For Each ambs In lista8
                        If ambs.ENTEROBACTERIAS = 1 Then
                            enterobacterias = 1
                        End If
                        If ambs.LISTAMBIENTAL = 1 Then
                            listambiental = 1
                        End If
                        If ambs.LISTMONO = 1 Then
                            listmono = 1
                        End If
                        If ambs.SALMONELLA = 1 Then
                            salmonella = 1
                        End If
                        If ambs.ECOLI = 1 Then
                            ecoli = 1
                        End If
                        If ambs.MOHOSYLEVADURAS = 1 Then
                            mohosylevaduras = 1
                        End If
                        If ambs.RB = 1 Then
                            rb = 1
                        End If
                        If ambs.CT = 1 Then
                            ct = 1
                        End If
                        If ambs.CF = 1 Then
                            cf = 1
                        End If
                        If ambs.PSEUDOMONASPP = 1 Then
                            pseudomonaspp = 1
                        End If
                    Next
                End If
            End If
            If enterobacterias = 1 Then
                texto = texto + " - Enterobacterias"
            End If
            If listambiental = 1 Then
                texto = texto + " - Listeria ambiental"
            End If
            If listmono = 1 Then
                texto = texto + " - Listeria monocitógenes"
            End If
            If salmonella = 1 Then
                texto = texto + " - Salmonella"
            End If
            If ecoli = 1 Then
                texto = texto + " - E. Coli"
            End If
            If mohosylevaduras = 1 Then
                texto = texto + " - Mohos y levaduras"
            End If
            If rb = 1 Then
                texto = texto + " - RB"
            End If
            If ct = 1 Then
                texto = texto + " - Coliformes totales"
            End If
            If cf = 1 Then
                texto = texto + " - CF"
            End If
            If pseudomonaspp = 1 Then
                texto = texto + " - Pseudomona spp"
            End If
            ' SI ES PARASITOLOGÍA ********************************************************************************
        ElseIf tipoinforme = "Parasitología" Then
            Dim p As New dParasitologiaSolicitud
            Dim lista9 As ArrayList
            lista9 = p.listarporsolicitud(ficha)
            Dim gastrointestinales As Integer = 0
            Dim fasciola As Integer = 0
            Dim coccidias As Integer = 0
            texto = ""
            If Not lista9 Is Nothing Then
                If lista9.Count > 0 Then
                    For Each p In lista9
                        If p.GASTROINTESTINALES = 1 Then
                            gastrointestinales = 1
                        End If
                        If p.FASCIOLA = 1 Then
                            fasciola = 1
                        End If
                        If p.COCCIDIAS = 1 Then
                            coccidias = 1
                        End If
                    Next
                End If
            End If
            If gastrointestinales = 1 Then
                texto = texto + " - Gastrointestinales"
            End If
            If fasciola = 1 Then
                texto = texto + " - Fasciola"
            End If
            If coccidias = 1 Then
                texto = texto + " - Coccidias"
            End If
            ' SI ES NUTRICIÓN ********************************************************************************
        ElseIf tipoinforme = "Nutrición" Then
            Dim mga As Integer = 0
            Dim mgb As Integer = 0
            Dim ensilados As Integer = 0
            Dim pasturas As Integer = 0
            Dim extetereo As Integer = 0
            Dim nida As Integer = 0
            texto = ""
            If Not listanutricion Is Nothing Then
                If listanutricion.Count > 0 Then
                    For Each sn In listanutricion
                        texto = texto & " // " & sn.MUESTRA & " - "
                        If sn.MGA = 1 Then
                            texto = texto & "MGA - "
                        End If
                        If sn.MGB = 1 Then
                            texto = texto & "MGB - "
                        End If
                        If sn.ENSILADOS = 1 Then
                            texto = texto & "Ensilados - "
                        End If
                        If sn.PASTURAS = 1 Then
                            texto = texto & "Pasturas - "
                        End If
                        If sn.EXTETEREO = 1 Then
                            texto = texto & "Extracto etéreo - "
                        End If
                        If sn.NIDA = 1 Then
                            texto = texto & "NIDA - "
                        End If
                    Next
                End If
            End If
            ' SI ES SUELOS ********************************************************************************
        ElseIf tipoinforme = "Suelos" Then
            Dim nitratos As Integer = 0
            Dim mineralizacion As Integer = 0
            Dim fosforobray As Integer = 0
            Dim fosforocitrico As Integer = 0
            Dim phagua As Integer = 0
            Dim phkci As Integer = 0
            Dim materiaorg As Integer = 0
            Dim potasioint As Integer = 0
            Dim sulfatos As Integer = 0
            Dim nitrogenovegetal As Integer = 0
            texto = ""
            If Not listasuelos Is Nothing Then
                If listasuelos.Count > 0 Then
                    For Each ss In listasuelos
                        texto = texto & " // " & ss.MUESTRA & " - "
                        If ss.NITRATOS = 1 Then
                            texto = texto & "Nitratos - "
                        End If
                        If ss.MINERALIZACION = 1 Then
                            texto = texto & "Mineralización - "
                        End If
                        If ss.FOSFOROBRAY = 1 Then
                            texto = texto & "Fósforo Bray I - "
                        End If
                        If ss.FOSFOROCITRICO = 1 Then
                            texto = texto & "Fósforo Ac.Cítrico - "
                        End If
                        If ss.PHAGUA = 1 Then
                            texto = texto & "pH Agua - "
                        End If
                        If ss.PHKCI = 1 Then
                            texto = texto & "pH KCI - "
                        End If
                        If ss.MATERIAORG = 1 Then
                            texto = texto & "Materia orgánica - "
                        End If
                        If ss.POTASIOINT = 1 Then
                            texto = texto & "Potasio intercambiable - "
                        End If
                        If ss.SULFATOS = 1 Then
                            texto = texto & "Sulfatos - "
                        End If
                        If ss.NITROGENOVEGETAL = 1 Then
                            texto = texto & "Nitrógeno vegetal - "
                        End If
                    Next
                End If
            End If
        End If
        '*** LISTADO DE MUESTRAS *********************************************************************************
        ' SI ES PRODUCTOS LÁCTEOS ********************************************************************************
        If tipoinforme = "Alimentos" Then
            texto2 = ""
            If Not lista4 Is Nothing Then
                If lista4.Count > 0 Then
                    For Each sm In lista4
                        texto2 = texto2 + sm.IDMUESTRA & " - "
                    Next
                End If
            End If
            ' SI ES AGUA ********************************************************************************
        ElseIf tipoinforme = "Agua" Then
            texto2 = ""
            If Not lista4 Is Nothing Then
                If lista4.Count > 0 Then
                    For Each sm In lista4
                        texto2 = texto2 + sm.IDMUESTRA & " - "
                    Next
                End If
            End If
            ' SI ES CALIDAD ********************************************************************************
        ElseIf tipoinforme = "Calidad de leche" Then
            texto2 = ""
            Dim cuenta_rb As Integer = 0
            Dim cuenta_rc As Integer = 0
            Dim cuenta_comp As Integer = 0
            Dim cuenta_criosc As Integer = 0
            Dim cuenta_inhib As Integer = 0
            Dim cuenta_espor As Integer = 0
            Dim cuenta_urea As Integer = 0
            Dim cuenta_termo As Integer = 0
            Dim cuenta_psicro As Integer = 0
            Dim cuenta_criosc_criosc As Integer = 0
            Dim cuenta_caseina As Integer = 0
            If Not lista5 Is Nothing Then
                If lista5.Count > 0 Then
                    For Each csm In lista5
                        texto2 = texto2 + csm.MUESTRA
                        If csm.RB = 1 Then
                            cuenta_rb = cuenta_rb + 1
                        End If
                        If csm.RC = 1 Then
                            cuenta_rc = cuenta_rc + 1
                        End If
                        If csm.COMPOSICION = 1 Then
                            cuenta_comp = cuenta_comp + 1
                        End If
                        If csm.CRIOSCOPIA = 1 Then
                            cuenta_criosc = cuenta_criosc + 1
                        End If
                        If csm.INHIBIDORES = 1 Then
                            cuenta_inhib = cuenta_inhib + 1
                        End If
                        If csm.ESPORULADOS = 1 Then
                            cuenta_espor = cuenta_espor + 1
                        End If
                        If csm.UREA = 1 Then
                            cuenta_urea = cuenta_urea + 1
                        End If
                        If csm.TERMOFILOS = 1 Then
                            cuenta_termo = cuenta_termo + 1
                        End If
                        If csm.PSICROTROFOS = 1 Then
                            cuenta_psicro = cuenta_psicro + 1
                        End If
                        If csm.CRIOSCOPIA_CRIOSCOPO = 1 Then
                            cuenta_criosc_criosc = cuenta_criosc_criosc + 1
                        End If
                        If csm.CASEINA = 1 Then
                            cuenta_caseina = cuenta_caseina + 1
                        End If
                        texto2 = texto2 + " - "
                    Next
                End If
            End If
            If cuenta_rb > 0 Then
                texto3 = texto3 & cuenta_rb & " RB - "
            End If
            If cuenta_rc > 0 Then
                texto3 = texto3 & cuenta_rc & " RC - "
            End If
            If cuenta_comp > 0 Then
                texto3 = texto3 & cuenta_comp & " Comp. - "
            End If
            If cuenta_criosc > 0 Then
                texto3 = texto3 & cuenta_criosc & " Criosc. - "
            End If
            If cuenta_inhib > 0 Then
                texto3 = texto3 & cuenta_inhib & " Inhib. - "
            End If
            If cuenta_espor > 0 Then
                texto3 = texto3 & cuenta_espor & " Espor. - "
            End If
            If cuenta_urea > 0 Then
                texto3 = texto3 & cuenta_urea & " Urea - "
            End If
            If cuenta_termo > 0 Then
                texto3 = texto3 & cuenta_termo & " Termof. - "
            End If
            If cuenta_psicro > 0 Then
                texto3 = texto3 & cuenta_psicro & " Psicro. - "
            End If
            If cuenta_criosc_criosc > 0 Then
                texto3 = texto3 & cuenta_criosc_criosc & " Criosc.(Crióscopo) - "
            End If
            If cuenta_caseina > 0 Then
                texto3 = texto3 & cuenta_caseina & " Caseina - "
            End If
            ' SI ES CONTROL LECHERO ********************************************************************************
        ElseIf tipoinforme = "Control Lechero" Then
            texto2 = ""
            ' SI ES ANTIBIOGRAMA ********************************************************************************
        ElseIf tipoinforme = "Bacteriología y Antibiograma" Then
            texto2 = ""
            If Not lista4 Is Nothing Then
                If lista4.Count > 0 Then
                    For Each sm In lista4
                        texto2 = texto2 + sm.IDMUESTRA & " - "
                    Next
                End If
            End If
            ' SI ES AMBIENTAL ********************************************************************************
        ElseIf tipoinforme = "Ambiental" Then
            texto2 = ""
            If Not lista4 Is Nothing Then
                If lista4.Count > 0 Then
                    For Each sm In lista4
                        texto2 = texto2 + sm.IDMUESTRA & " - "
                    Next
                End If
            End If
            ' SI ES PARASITOLOGÍA ********************************************************************************
        ElseIf tipoinforme = "Parasitología" Then
            texto2 = ""
            If Not lista4 Is Nothing Then
                If lista4.Count > 0 Then
                    For Each sm In lista4
                        texto2 = texto2 + sm.IDMUESTRA & " - "
                    Next
                End If
            End If
            ' SI ES PAL ********************************************************************************
        ElseIf tipoinforme = "PAL" Then
            texto2 = ""
            If Not lista10 Is Nothing Then
                If lista10.Count > 0 Then
                    For Each spal In lista10
                        texto2 = texto2 + spal.MATRICULA & " - "
                    Next
                End If
            End If
            Dim solpal As New dSolicitudPAL
            solpal.FICHA = ficha
            solpal = solpal.buscar
            If Not solpal Is Nothing Then
                ' x1hoja.Cells(fila, columna).Formula = "Vacas: " & solpal.VACAS & " - " & "Fecha extracción: " & solpal.FECHAEXT
            End If
            '********************************************************************************************************************
            ' SI ES BRUCELOSIS LECHE ********************************************************************************
        ElseIf tipoinforme = "Brucelosis en leche" Then
            texto2 = ""
            If Not listabl Is Nothing Then
                If listabl.Count > 0 Then
                    For Each sm In listabl
                        texto2 = texto2 + sm.IDMUESTRA & " - "
                    Next
                End If
            End If
        End If
        '********************************************************************************************************************
        If tipoinforme = "Nutrición" Or tipoinforme = "Suelos" Then
            If email <> "" And email <> "no aportado" Then
                'CONFIGURACIÓN DEL STMP 
                _SMTP.Credentials = New System.Net.NetworkCredential("notificaciones@colaveco.com.uy", "19912021Notificaciones")
                _SMTP.Host = "170.249.199.66"
                _SMTP.Port = 25
                _SMTP.EnableSsl = False


                ' CONFIGURACION DEL MENSAJE 
                _Message.[To].Add(LTrim(email))
                _Message.[To].Add("envios@colaveco.com.uy")
                'Cuenta de Correo al que se le quiere enviar el e-mail 
                _Message.From = New System.Net.Mail.MailAddress("notificaciones@colaveco.com.uy", "COLAVECO", System.Text.Encoding.UTF8)
                'Quien lo envía 
                _Message.Subject = "Solicitud de análisis"
                'Sujeto del e-mail 
                _Message.SubjectEncoding = System.Text.Encoding.UTF8
                'Codificacion 
                _Message.Body = "Ha ingresado una solicitud con el número" & " " & ficha & vbCrLf _
                & "Fecha de recepción: " & fecha & "." & vbCrLf _
                & "A nombre de: " & nombre_productor & "." & vbCrLf _
                & "Muestras ingresadas: " & nmuestras & "." & vbCrLf _
                & "Tipo de muestra: " & muestraweb & "." & vbCrLf _
                & "Análisis requerido: " & tipoinforme & "." & vbCrLf _
                & "Subtipo: " & subtipoinforme & "." & vbCrLf _
                & vbCrLf _
                & texto & vbCrLf _
                & vbCrLf _
                & "Observaciones:" & vbCrLf _
                & observaciones & vbCrLf _
                & vbCrLf _
                & "En nuestro sitio web, http://www.colaveco.com.uy/gestor, puede ver el estado de su solicitud." & vbCrLf _
                & "Gracias." & vbCrLf _
                & "COLAVECO"
                'contenido del mail 
                _Message.BodyEncoding = System.Text.Encoding.UTF8 '
                _Message.Priority = System.Net.Mail.MailPriority.Normal
                _Message.IsBodyHtml = False
                ' ADICION DE DATOS ADJUNTOS ‘
                Try
                    _SMTP.Send(_Message)
                Catch ex As System.Net.Mail.SmtpException ' MessageBox.Show(ex.ToString) 
                End Try
            End If
            email = ""
            nficha = 0
        Else
            If email <> "" And email <> "no aportado" Then
                Dim var = Mid(email, Len(email))
                If var = "," Then
                    email = Replace(email, ",", "")
                End If
                'CONFIGURACIÓN DEL STMP 
                _SMTP.Credentials = New System.Net.NetworkCredential("notificaciones@colaveco.com.uy", "19912021Notificaciones")
                _SMTP.Host = "170.249.199.66"
                _SMTP.Port = 25
                _SMTP.EnableSsl = False
                ' CONFIGURACION DEL MENSAJE 
                _Message.[To].Add(LTrim(email))
                _Message.[To].Add("envios@colaveco.com.uy")
                'Cuenta de Correo al que se le quiere enviar el e-mail 
                _Message.From = New System.Net.Mail.MailAddress("notificaciones@colaveco.com.uy", "COLAVECO", System.Text.Encoding.UTF8)
                'Quien lo envía 
                _Message.Subject = "Solicitud de análisis"
                'Sujeto del e-mail 
                _Message.SubjectEncoding = System.Text.Encoding.UTF8
                'Codificacion 
                _Message.Body = "Ha ingresado una solicitud con el número" & " " & ficha & vbCrLf _
                & "Fecha/Hora de recepción: " & fecha & "." & vbCrLf _
                & "A nombre de: " & nombre_productor & "." & vbCrLf _
                & "Muestras ingresadas: " & nmuestras & "." & vbCrLf _
                & "Tipo de muestra: " & muestraweb & "." & vbCrLf _
                & "Análisis requerido: " & tipoinforme & "." & vbCrLf _
                & "Subtipo: " & subtipoinforme & "." & vbCrLf _
                & vbCrLf _
                & texto & vbCrLf _
                & vbCrLf _
                & "Identificación de las muestras:" & vbCrLf _
                & texto2 & vbCrLf _
                & vbCrLf _
                & "Observaciones:" & vbCrLf _
                & observaciones & vbCrLf _
                & vbCrLf _
                & "En nuestro sitio web, http://www.colaveco.com.uy/gestor, puede ver el estado de su solicitud." & vbCrLf _
                & "Gracias." & vbCrLf & vbCrLf _
                & "COLAVECO" & vbCrLf _
                & "Parque El Retiro - Nueva Helvecia - Tel/Fax: 45545311/45545975/45546838" & vbCrLf _
                & "Email: colaveco@gmail.com - web: http://www.colaveco.com.uy" & vbCrLf & vbCrLf _
                & "-------------------------------------------------------------------------------------" & vbCrLf _
                & "Cuando el cliente solicite suspender el servicio ya presupuestado y en ejecución, o una parte del mismo," & vbCrLf _
                & "los costos de las actividades ya realizadas en el momento de la suspensión deberán pagarse."
                'contenido del mail 
                _Message.BodyEncoding = System.Text.Encoding.UTF8 '
                _Message.Priority = System.Net.Mail.MailPriority.Normal
                _Message.IsBodyHtml = False
                ' ADICION DE DATOS ADJUNTOS ‘
                Try
                    _SMTP.Send(_Message)
                Catch ex As System.Net.Mail.SmtpException ' MessageBox.Show(ex.ToString) 
                End Try
            End If
        End If
    End Sub
    Private Sub enviosms_sw()
        Dim num1 As String = ""
        Dim num2 As String = ""
        Dim email1 As String = ""
        Dim email2 As String = ""
        Dim sms As String = ""
        Dim sms1 As String = ""
        Dim sms2 As String = ""
        Dim cel1 As String = ""
        Dim cel2 As String = ""
        Dim largotexto As Integer = 0
        Dim celular1 As String = ""
        Dim celular2 As String = ""
        Dim _Message As New System.Net.Mail.MailMessage()
        Dim _SMTP As New System.Net.Mail.SmtpClient
        Dim texto As String = celular
        Dim cantcaracteres As Integer = Len(texto)
        If celular <> "" Then
            largotexto = celular.Length
        End If
        Dim posicion As Integer
        Dim posicion1 As Integer
        Dim posicion2 As Integer
        posicion = InStr(celular, ",")
        If posicion > 0 Then
            posicion1 = posicion - 1
            posicion2 = posicion + 1
            cel1 = Mid(celular, 1, posicion1)
            cel2 = Mid(celular, posicion2, largotexto)
            celular1 = cel1
            email = celular1
            num1 = Mid(celular1, 3, 1)
            If num1 = "9" Or num1 = "8" Or num1 = "1" Or num1 = "2" Then
                'ancel es numero  + pin
                sms1 = email & "@antelinfo.com.uy"
            ElseIf num1 = "3" Or num1 = "4" Or num1 = "5" Then
                'movistar es numero (sin 0 inicial + pin)
                If Mid(celular, 1, 1) = "0" Then
                    celular1 = celular.Remove(0, 1)
                End If
                email = celular1
                sms1 = email & "@sms.movistar.com.uy"
            ElseIf num1 = "6" Or num1 = "7" Then
                'claro es numero (sin 0 inicial sin pin)
                If Mid(celular, 1, 1) = "0" Then
                    celular2 = celular.Remove(0, 1)
                End If
                email = celular1
                sms1 = email & "@sms.ctimovil.com.uy"
            End If
            '*****************************************
            celular2 = cel2
            email2 = celular2
            num2 = Mid(celular2, 1, 1)
            If num2 = "9" Or num2 = "8" Or num2 = "1" Or num2 = "2" Then
                'ancel es numero (sin 09 inicial + pin)
                sms2 = email2 & "@antelinfo.com.uy"
            ElseIf num2 = "3" Or num2 = "4" Or num2 = "5" Then
                'movistar es numero (sin 0 inicial + pin)
                If Mid(celular2, 1, 1) = "0" Then
                    celular2 = celular2.Remove(0, 1)
                End If
                email2 = celular2
                sms2 = email2 & "@sms.movistar.com.uy"
            ElseIf num2 = "6" Or num2 = "7" Then
                'claro es numero (sin 0 inicial sin pin)
                If Mid(celular2, 1, 1) = "0" Then
                    celular2 = celular2.Remove(0, 1)
                End If
                email2 = celular2
                sms2 = email2 & "@sms.ctimovil.com.uy"
            End If
            sms = sms1 & "," & sms2
        Else
            celular2 = celular
            email = celular2
            num1 = Mid(celular2, 1, 1)
            If num1 = "9" Or num1 = "8" Or num1 = "1" Or num1 = "2" Then
                'ancel es numero (sin 09 inicial + pin)
                sms = email & "@antelinfo.com.uy"
            ElseIf num1 = "3" Or num1 = "4" Or num1 = "5" Then
                'movistar es numero (sin 0 inicial + pin)
                If Mid(celular, 1, 1) = "0" Then
                    celular2 = celular.Remove(0, 1)
                End If
                email = celular2
                sms = email & "@sms.movistar.com.uy"
            ElseIf num1 = "6" Or num1 = "7" Then
                'claro es numero (sin 0 inicial sin pin)
                If Mid(celular, 1, 1) = "0" Then
                    celular2 = celular.Remove(0, 1)
                End If
                email = celular2
                sms = email & "@sms.ctimovil.com.uy"
            End If
        End If
        Dim sa As New dSolicitudAnalisis
        Dim p As New dCliente
        Dim ti As New dTipoInforme
        Dim nombre_productor As String = ""
        Dim tipo_analisis As String = ""
        sa.ID = nficha
        sa = sa.buscar
        If Not sa Is Nothing Then
            p.ID = sa.IDPRODUCTOR
            p = p.buscar
            If Not p Is Nothing Then
                nombre_productor = p.NOMBRE
            End If
            ti.ID = sa.IDTIPOINFORME
            ti = ti.buscar
            If Not ti Is Nothing Then
                tipo_analisis = ti.NOMBRE
            End If
        End If
        If sms <> "" Then
            'CONFIGURACIÓN DEL STMP 
            _SMTP.Credentials = New System.Net.NetworkCredential("notificaciones@colaveco.com.uy", "19912021Notificaciones")
            _SMTP.Host = "170.249.199.66"
            _SMTP.Port = 25
            _SMTP.EnableSsl = False
            ' CONFIGURACION DEL MENSAJE 
            _Message.[To].Add(sms)
            _Message.[To].Add("envios@colaveco.com.uy")
            'Cuenta de Correo al que se le quiere enviar el e-mail 
            _Message.From = New System.Net.Mail.MailAddress("notificaciones@colaveco.com.uy", "COLAVECO", System.Text.Encoding.UTF8)
            'Quien lo envía 
            _Message.Subject = "Su solicitud de análisis Nº " & " " & nficha & " - " & tipo_analisis & " (" & nombre_productor & ")," & "ha ingresado correctamente al sistema. Gracias. COLAVECO"
            'Sujeto del e-mail 
            _Message.SubjectEncoding = System.Text.Encoding.UTF8
            'Codificacion 
            'contenido del mail 
            _Message.BodyEncoding = System.Text.Encoding.UTF8 '
            _Message.Priority = System.Net.Mail.MailPriority.Normal
            _Message.IsBodyHtml = False
            ' ADICION DE DATOS ADJUNTOS ‘
            'Dim _File As String = My.Application.Info.DirectoryPath & "archivo" 'archivo que se quiere adjuntar ‘
            'Dim _Attachment As New System.Net.Mail.Attachment(_File, System.Net.Mime.MediaTypeNames.Application.Octet) '
            '_Message.Attachments.Add(_Attachment) 'ENVIO 
            Try
                _SMTP.Send(_Message)
                'MessageBox.Show("Mensaje enviado!", "Correo", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As System.Net.Mail.SmtpException ' MessageBox.Show(ex.ToString) 
            End Try
        End If
        email = ""
        texto = ""
    End Sub
    Private Sub enviaremailcompras()
        Dim _Message As New System.Net.Mail.MailMessage()
        Dim _SMTP As New System.Net.Mail.SmtpClient
        Dim email As String = ""
        Dim destinatario As String = ""
        Dim c As New dCompras
        c.ID = compraid
        c = c.buscar
        If Not c Is Nothing Then
            If c.EMAIL <> "" Then
                email = Trim(c.EMAIL)
            End If
            Dim p As New dProveedores
            p.ID = c.PROVEEDOR
            p = p.buscar
            If Not p Is Nothing Then
                destinatario = p.NOMBRE
            End If
        End If
        If email <> "" And email <> "no aportado" Then
            'CONFIGURACIÓN DEL STMP 
            _SMTP.Credentials = New System.Net.NetworkCredential("laboratorio@colaveco.com.uy", "19912021Laboratorio")
            _SMTP.Host = "170.249.199.66"
            _SMTP.Port = 25
            _SMTP.EnableSsl = False
            ' CONFIGURACION DEL MENSAJE 
            _Message.[To].Add(email)
            _Message.[To].Add("envios@colaveco.com.uy")
            'Cuenta de Correo al que se le quiere enviar el e-mail 
            _Message.From = New System.Net.Mail.MailAddress("laboratorio@colaveco.com.uy", "COLAVECO", System.Text.Encoding.UTF8)
            'Quien lo envía 
            _Message.Subject = "Orden de compra" & " - " & compraid
            'Sujeto del e-mail 
            _Message.SubjectEncoding = System.Text.Encoding.UTF8
            'Codificacion 
            '_Message.Body = "Se han enviado las siguientes cajas:" & " " & ecaja1 & ", " & "por" & " " & eagencia & " " & "envío nº" & " " & eremito & ""
            _Message.Body = "Sres. de" & " " & destinatario & ", " & "por medio del presente correo adjuntamos orden de compra. Desde ya gracias. COLAVECO"
            'contenido del mail 
            _Message.BodyEncoding = System.Text.Encoding.UTF8 '
            _Message.Priority = System.Net.Mail.MailPriority.Normal
            _Message.IsBodyHtml = False
            ' ADICION DE DATOS ADJUNTOS ‘
            Dim _File As String = "\\192.168.1.10\e\NET\COMPRAS\OC\OC_" & compraid & ".xls" 'archivo que se quiere adjuntar ‘
            Dim _Attachment As New System.Net.Mail.Attachment(_File, System.Net.Mime.MediaTypeNames.Application.Octet) '
            _Message.Attachments.Add(_Attachment) 'ENVIO 
            Try
                _SMTP.Send(_Message)
                'MessageBox.Show("Correo enviado!", "Correo", MessageBoxButtons.OK, MessageBoxIcon.Information)
                marcarenvio()

            Catch ex As System.Net.Mail.SmtpException ' MessageBox.Show(ex.ToString) 
                'MessageBox.Show("Falla al enviar el correo!", "Correo", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End Try
            _File = ""
            _Attachment = Nothing
        Else
            MsgBox("No tiene dirección de correo cargada")
        End If
        email = ""
    End Sub
    Private Sub enviaremailCajasAtrasadas(ByVal clienteid As Long)
        Dim _Message As New System.Net.Mail.MailMessage()
        Dim _SMTP As New System.Net.Mail.SmtpClient
        Dim email As String = ""
        Dim destinatario As String = ""
        Dim c As New dCliente
        c.ID = clienteid
        c = c.buscar
        If Not c Is Nothing Then
            If c.EMAIL <> "" Then
                email = Trim(c.EMAIL)
            End If
            If Not c.NOMBRE <> "" Then
                destinatario = c.NOMBRE
            End If
        End If
        If email <> "" And email <> "no aportado" Then
            'CONFIGURACIÓN DEL STMP 
            _SMTP.Credentials = New System.Net.NetworkCredential("laboratorio@colaveco.com.uy", "19912021Laboratorio")
            _SMTP.Host = "170.249.199.66"
            _SMTP.Port = 25
            _SMTP.EnableSsl = False

            ' CONFIGURACION DEL MENSAJE 
            _Message.[To].Add(email)
            _Message.[To].Add("envios@colaveco.com.uy")
            'Cuenta de Correo al que se le quiere enviar el e-mail 
            _Message.From = New System.Net.Mail.MailAddress("laboratorio@colaveco.com.uy", "COLAVECO", System.Text.Encoding.UTF8)
            'Quien lo envía 
            _Message.Subject = "Orden de compra" & " - " & compraid
            'Sujeto del e-mail 
            _Message.SubjectEncoding = System.Text.Encoding.UTF8
            'Codificacion 
            _Message.Body = "Sres. de" & " " & destinatario & ", " & "por medio del presente correo adjuntamos orden de compra. Desde ya gracias. COLAVECO"
            'contenido del mail 
            _Message.BodyEncoding = System.Text.Encoding.UTF8 '
            _Message.Priority = System.Net.Mail.MailPriority.Normal
            _Message.IsBodyHtml = False
            ' ADICION DE DATOS ADJUNTOS ‘
            Dim _File As String = "\\192.168.1.10\e\NET\COMPRAS\OC\OC_" & compraid & ".xls" 'archivo que se quiere adjuntar ‘
            Dim _Attachment As New System.Net.Mail.Attachment(_File, System.Net.Mime.MediaTypeNames.Application.Octet) '
            _Message.Attachments.Add(_Attachment) 'ENVIO 
            Try
                _SMTP.Send(_Message)
                'MessageBox.Show("Correo enviado!", "Correo", MessageBoxButtons.OK, MessageBoxIcon.Information)
                marcarenvio()

            Catch ex As System.Net.Mail.SmtpException ' MessageBox.Show(ex.ToString) 
                'MessageBox.Show("Falla al enviar el correo!", "Correo", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Exit Sub
            End Try
            _File = ""
            _Attachment = Nothing
        Else
            MsgBox("No tiene dirección de correo cargada")
        End If
        email = ""
    End Sub
    Private Sub enviarcompras()
        Dim c As New dCompras
        Dim lista As New ArrayList
        lista = c.listarsinenviar
        If Not lista Is Nothing Then
            For Each c In lista
                compraid = c.ID
                enviaremailcompras()
            Next
        End If
    End Sub
    Private Sub enviarMailCajas()
        Dim ListaEC As New dEnvioCajas
        Dim lista As New ArrayList
        lista = ListaEC.buscarCajasAtrasadas
        If Not lista Is Nothing Then
            For Each c In lista
                Dim enviocaja As New dEnvioCajas
                enviocaja = enviocaja.buscarCajasConFechaPorPedido(c.ID)
                clienteid = c.IDPRODUCTOR
                enviaremailCajasAtrasadas(clienteid)
            Next
        End If
    End Sub
    Private Sub enviomailcajas(ByVal _id As Long)
        Dim p As New dPedidos
        p.ID = _id
        p = p.buscar
        Dim eagencia As String = ""
        Dim eremito As String = ""
        Dim hora As Date = Now()
        Dim horaenvio As String
        Dim horaenvio2 As Integer
        horaenvio = Format(hora, "yyyy-MM-dd HH:mm:ss")
        horaenvio2 = Mid(horaenvio, 12, 2)
        Dim _Message As New System.Net.Mail.MailMessage()
        Dim _SMTP As New System.Net.Mail.SmtpClient
        Dim c As New dCliente
        Dim pw_com As New dProductorWeb_com
        Dim env As New dEnvioCajas
        Dim ag As New dEmpresaT
        ag.ID = p.IDAGENCIA
        ag = ag.buscar
        eagencia = ag.NOMBRE
        Dim texto As String = ""
        Dim id As Long = p.IDPRODUCTOR 'CType(TextIdProductor.Text, Long)
        c.ID = id 'Val(TextIdProductor.Text)
        c = c.buscar
        If Not c Is Nothing Then
            email = ""
            email2 = ""
            If c.NOT_EMAIL_FRASCOS1 <> "" Then
                email = RTrim(c.NOT_EMAIL_FRASCOS1)
            End If
            If c.NOT_EMAIL_FRASCOS2 <> "" Then
                email2 = RTrim(c.NOT_EMAIL_FRASCOS2)
            End If
            If email = "" Then
                If email2 = "" Then
                    If c.EMAIL <> "" Then
                        email = RTrim(c.EMAIL)
                    End If
                Else
                    email = email2
                End If
            Else
                If email2 = "" Then
                    email = email
                Else
                    email = email & "," & email2
                End If
            End If
            If Not c.USUARIO_WEB = "" Then
                pw_com.USUARIO = c.USUARIO_WEB
                pw_com = pw_com.buscar
                If Not pw_com Is Nothing Then
                    celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                End If
            End If
        End If
        'eagencia = ComboAgencia.Text
        Dim idped As Long = p.ID 'CType(TextId.Text, Long)
        Dim lista As New ArrayList
        lista = env.listarporid(idped)
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each env In lista
                    texto = texto & env.IDCAJA & ", "
                    eremito = env.ENVIO
                Next
            End If
        End If
        If Not String.IsNullOrEmpty(email) And String.IsNullOrEmpty(email2) Then

                'CONFIGURACIÓN DEL STMP 
            _SMTP.Credentials = New System.Net.NetworkCredential("notificaciones@colaveco.com.uy", "19912021Notificaciones")
            _SMTP.Host = "170.249.199.66"
            _SMTP.Port = 25
            _SMTP.EnableSsl = False

                ' CONFIGURACION DEL MENSAJE 
            _Message.[To].Add(email)
            _Message.[To].Add("envios@colaveco.com.uy")
                'Cuenta de Correo al que se le quiere enviar el e-mail 
            _Message.From = New System.Net.Mail.MailAddress("notificaciones@colaveco.com.uy", "COLAVECO", System.Text.Encoding.UTF8)
                'Quien lo envía 
                _Message.Subject = "Envío de cajas"
                'Sujeto del e-mail 
                _Message.SubjectEncoding = System.Text.Encoding.UTF8
                'Codificacion 
                '_Message.Body = "Se han enviado las siguientes cajas:" & " " & ecaja1 & ", " & "por" & " " & eagencia & " " & "envío nº" & " " & eremito & ""
                If eremito <> "" Then
                    _Message.Body = "Colaveco despachó la/s siguiente/s caja/s Nº " & texto & " por agencia" & " " & eagencia & ", " & "envío nº" & " " & eremito
                Else
                    _Message.Body = "Colaveco despachó la/s siguiente/s caja/s Nº " & texto & " por agencia" & " " & eagencia
                End If
                'contenido del mail 
                _Message.BodyEncoding = System.Text.Encoding.UTF8 '
                _Message.Priority = System.Net.Mail.MailPriority.Normal
                _Message.IsBodyHtml = False
                ' ADICION DE DATOS ADJUNTOS ‘
                'Dim _File As String = My.Application.Info.DirectoryPath & "archivo" 'archivo que se quiere adjuntar ‘
                'Dim _Attachment As New System.Net.Mail.Attachment(_File, System.Net.Mime.MediaTypeNames.Application.Octet) '
                '_Message.Attachments.Add(_Attachment) 'ENVIO 
                Try
                    _SMTP.Send(_Message)
                    'MessageBox.Show("Correo enviado!", "Correo", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As System.Net.Mail.SmtpException ' MessageBox.Show(ex.ToString) 
                End Try
            Else
                'MsgBox("Este cliente no tiene correo electrónico ingresado, por lo tanto no se le envía aviso.")
        End If
            email = ""
            eagencia = ""
            eremito = ""
            texto = ""
    End Sub

    Function IsValidEmailAddress(ByVal emailAddress As String) As Boolean
        Dim valid As Boolean = True
        Try
            Dim a = New System.Net.Mail.MailAddress(emailAddress)
        Catch ex As FormatException
            valid = False
        End Try
        Return valid
    End Function
    Private Sub envio_mail_informes()
        Dim _Message As New System.Net.Mail.MailMessage()
        Dim _SMTP As New System.Net.Mail.SmtpClient
        Dim _informe As String = ""
        Dim _cliente As String = ""
        Dim _email As String = ""
        Dim _texto As String = ""
        Dim _adjunto As String = ""
        Dim c As New dCorreos
        c.FICHA = idficha
        c = c.buscar
        If Not c Is Nothing Then
            _informe = c.INFORME
            _cliente = c.CLIENTE
            _email = c.EMAIL
            _texto = c.TEXTO
            _adjunto = c.ADJUNTO
        End If
        If _email <> "" And _email <> "No aportado" And _email <> "no aportado" Then
            If _adjunto <> "" Then
                'CONFIGURACIÓN DEL STMP 
                _SMTP.Credentials = New System.Net.NetworkCredential("notificaciones@colaveco.com.uy", "19912021Notificaciones")
                _SMTP.Host = "170.249.199.66"
                _SMTP.Port = 25
                _SMTP.EnableSsl = False

                ' CONFIGURACION DEL MENSAJE 
                _Message.[To].Add(_email)
                _Message.[To].Add("envios@colaveco.com.uy")
                'Cuenta de Correo al que se le quiere enviar el e-mail 
                _Message.From = New System.Net.Mail.MailAddress("notificaciones@colaveco.com.uy", "COLAVECO", System.Text.Encoding.UTF8)
                'Quien lo envía 
                _Message.Subject = "Informe - Colaveco"
                'Sujeto del e-mail 
                _Message.SubjectEncoding = System.Text.Encoding.UTF8
                'Codificacion 
                _Message.Body = _texto
                'contenido del mail 
                _Message.BodyEncoding = System.Text.Encoding.UTF8 '
                _Message.Priority = System.Net.Mail.MailPriority.Normal
                _Message.IsBodyHtml = False
                ' ADICION DE DATOS ADJUNTOS ‘
                Dim _File As String = _adjunto 'archivo que se quiere adjuntar ‘
                Dim _Attachment As New System.Net.Mail.Attachment(_File, System.Net.Mime.MediaTypeNames.Application.Octet) '
                _Message.Attachments.Add(_Attachment) 'ENVIO 
                Try
                    _SMTP.Send(_Message)
                    Dim c2 As New dCorreos
                    c2.FICHA = idficha
                    c2.marcar_enviado()
                    'MessageBox.Show("Correo enviado!", "Correo", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As System.Net.Mail.SmtpException ' MessageBox.Show(ex.ToString) 
                End Try
            Else
                'CONFIGURACIÓN DEL STMP 
                _SMTP.Credentials = New System.Net.NetworkCredential("notificaciones@colaveco.com.uy", "19912021Notificaciones")
                _SMTP.Host = "170.249.199.66"
                _SMTP.Port = 25
                _SMTP.EnableSsl = False

                ' CONFIGURACION DEL MENSAJE 
                _Message.[To].Add(_email)
                _Message.[To].Add("envios@colaveco.com.uy")
                'Cuenta de Correo al que se le quiere enviar el e-mail 
                _Message.From = New System.Net.Mail.MailAddress("notificaciones@colaveco.com.uy", "COLAVECO", System.Text.Encoding.UTF8)
                'Quien lo envía 
                _Message.Subject = "Informe - Colaveco"
                'Sujeto del e-mail 
                _Message.SubjectEncoding = System.Text.Encoding.UTF8
                'Codificacion 
                _Message.Body = _texto
                'contenido del mail 
                _Message.BodyEncoding = System.Text.Encoding.UTF8 '
                _Message.Priority = System.Net.Mail.MailPriority.Normal
                _Message.IsBodyHtml = False
                Try
                    _SMTP.Send(_Message)
                    Dim c2 As New dCorreos
                    c2.FICHA = idficha
                    c2.marcar_enviado()
                    'MessageBox.Show("Correo enviado!", "Correo", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As System.Net.Mail.SmtpException ' MessageBox.Show(ex.ToString) 
                End Try
            End If
        End If
        _email = ""
    End Sub

#End Region

#Region "SIN-USO"

    Private Sub robot_revisar_cuentas_corrientes()
        'No se usa /hardcode
        Dim sv As New dSinVisualizacion
        Dim cliente As Long = 0
        Dim lista As New ArrayList
        Dim idficha As Long = 0
        lista = sv.listarsv
        If Not lista Is Nothing Then
            For Each sv In lista
                barra = barra + 1
                ProgressBar1.Value = barra
                If barra = 100 Then
                    barra = 0
                End If
                Dim sa As New dSolicitudAnalisis
                sa.ID = sv.FICHA
                sa = sa.buscar
                If Not sa Is Nothing Then
                    cliente = sa.IDPRODUCTOR
                    idficha = sa.ID
                End If
                '*** VERIFICA SI EL CLIENTE TIENE DEUDA ***************************************
                '****SI EL CLIENTE ES CONTADO*******************************************
                Dim cli As New dCliente
                cli.ID = cliente
                cli = cli.buscar
                If Not cli Is Nothing Then
                    If cli.PROLESA = 0 Then
                        If cli.FAC_CONTADO = 1 Then
                            Dim f As New dFacturacion
                            Dim listaF As New ArrayList
                            listaF = f.listarxficha(idficha)
                            If Not listaF Is Nothing Then
                                For Each f In listaF
                                    If f.FACTURA <> 0 And f.FACTURA <> 999999 Then
                                        Dim mc As New dMovCte
                                        mc.MCCCMP = f.FACTURA
                                        mc = mc.buscarxcomprobante
                                        If Not mc Is Nothing Then
                                            If mc.MCCPAG >= mc.MCCIMP Then
                                                Dim fecha As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                                                Dim fec As String
                                                fec = Format(fecha, "yyyy-MM-dd")
                                                Dim c As New dSinVisualizacion
                                                c.FICHA = idficha
                                                c = c.buscarxficha
                                                Dim abonado As Integer = 0
                                                c.marcarvisualizacion(fec)
                                                abonado = 2
                                                marcarweb(idficha, abonado)
                                            End If
                                        Else
                                            Dim fecha As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                                            Dim fec As String
                                            fec = Format(fecha, "yyyy-MM-dd")
                                            Dim c As New dSinVisualizacion
                                            c.FICHA = idficha
                                            c = c.buscarxficha
                                            Dim abonado As Integer = 0
                                            c.marcarvisualizacion(fec)
                                            abonado = 2
                                            marcarweb(idficha, abonado)
                                        End If
                                    End If
                                Next
                            End If
                        Else
                            Dim mc As New dMovCte
                            Dim listamc As New ArrayList
                            Dim fechaactual As Date = Now.ToString("yyyy-MM-dd")
                            Dim fechaact As String = Format(fechaactual, "yyyy-MM-dd")
                            Dim vencido As Integer = 0
                            listamc = mc.listarxcli(cliente)
                            If Not listamc Is Nothing Then
                                For Each mc In listamc
                                    Dim fechavto As Date = mc.MCCVTO
                                    Dim fecvto As String = Format(fechavto, "yyyy-MM-dd")
                                    If fecvto < fechaact Then
                                        If mc.MCCPAG < mc.MCCIMP Then
                                            Dim diferencia As Double = 0
                                            diferencia = mc.MCCIMP - mc.MCCPAG
                                            If diferencia > 100 Then
                                                vencido = 1
                                            End If
                                        End If
                                    End If
                                Next
                            End If
                            If vencido = 0 Then
                                Dim fecha As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                                Dim fec As String
                                fec = Format(fecha, "yyyy-MM-dd")
                                Dim c As New dSinVisualizacion
                                c.FICHA = idficha
                                c = c.buscarxficha
                                Dim abonado As Integer = 0
                                c.marcarvisualizacion(fec)
                                abonado = 1
                                marcarweb(idficha, abonado)
                            End If
                        End If
                    Else
                        Dim f As New dFacturacion
                        Dim listaF As New ArrayList
                        Dim _abonado As Integer = 0
                        listaF = f.listarxficha(idficha)
                        If Not listaF Is Nothing Then
                            For Each f In listaF
                                If f.FACTURA <> 0 And f.FACTURA <> 999999 Then
                                    _abonado = 1
                                End If
                            Next
                            If _abonado = 1 Then
                                Dim fecha As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                                Dim fec As String
                                fec = Format(fecha, "yyyy-MM-dd")
                                Dim c As New dSinVisualizacion
                                c.FICHA = idficha
                                c = c.buscarxficha
                                Dim abonado As Integer = 0
                                c.marcarvisualizacion(fec)
                                abonado = 2
                                marcarweb(idficha, abonado)
                            End If
                        End If
                    End If
                    '*******************************************************************************
                End If
            Next
        End If
    End Sub
    Private Sub robot_subir_ctacte()
        'No se usa /hardcode
        Dim fechaactual As Date = Now()
        Dim _fecha As String
        _fecha = Format(fechaactual, "yyyy-MM-dd")
        Dim movimientosDict As New Dictionary(Of String, List(Of dMovimientos))
        Dim movimientos As New List(Of dMovimientos)
        Dim m As New dMovCte2
        Dim listamovimientos As New ArrayList
        listamovimientos = m.listarxfecha(_fecha)
        If Not listamovimientos Is Nothing Then
            For Each m In listamovimientos
                barra = barra + 1
                ProgressBar1.Value = barra
                If barra = 100 Then
                    barra = 0
                End If
                Dim movi As New dMovimientos
                Dim _path As String = ""
                Dim path As String = "/home/colaveco/public_html/gestor/facturas/"
                movi.idnet_movimiento = m.MCCNRO
                movi.fecha_emision = m.MCCFCH
                movi.fecha_vencimiento = m.MCCVTO
                Dim tipo As String = ""
                If m.DOCCOD = "NF" Then
                    tipo = "NC"
                ElseIf m.DOCCOD = "NI" Then
                    tipo = "NC"
                ElseIf m.DOCCOD = "01" Then
                    tipo = "R"
                ElseIf m.DOCCOD = "02" Then
                    tipo = "ND"
                ElseIf m.DOCCOD = "AA" Then
                    tipo = "AA"
                ElseIf m.DOCCOD = "AD" Then
                    tipo = "AD"
                ElseIf m.DOCCOD = "CI" Then
                    tipo = "F"
                ElseIf m.DOCCOD = "FA" Then
                    tipo = "F"
                ElseIf m.DOCCOD = "FF" Then
                    tipo = "F"
                ElseIf m.DOCCOD = "NC" Then
                    tipo = "NC"
                End If
                movi.tipo = tipo
                movi.numero = m.MCCDOC
                movi.detalle = m.MCCDES
                movi.importe = m.MCCIMP
                movi.tipo_movimiento = m.MCCCOD
                movi.importe_pago = m.MCCPAG
                movi.idnet_usuario = m.CLICOD
                If m.MCCTIP = "V" Then
                    Dim f As New dFactur
                    f.FACNRO = m.MCCCMP
                    f = f.buscar
                    If Not f Is Nothing Then
                        Dim c As String = "\\192.168.1.106\apls\soft\"
                        Dim tx As String = f.FACPDF
                        If tx.Contains(c) Then
                            _path = tx.Replace(c, "")
                            factura_origen = _path
                            path = path & _path
                        End If
                    End If
                End If
                movi.path_pdf = path
                movimientos.Add(movi)
                movi = Nothing
                tipo = Nothing
                subirFacturaPdf()
            Next
            If m.MCCDOC <> 0 Then
                movimientosDict.Add("movimientos", movimientos)
                Dim parameters As String = JsonConvert.SerializeObject(movimientosDict, Formatting.None)
                Dim status As HttpStatusCode = HttpStatusCode.ExpectationFailed
                Dim response As Byte() = PostResponse("http://colaveco-gestor.herokuapp.com/factura_movimientos/migrar", "POST", parameters, status)
            End If
        End If
    End Sub
    Private Sub robot_subir_ctacte2()
        'No se usa /hardcode
        Dim un As New dUltimoNumero
        Dim ultnum As Long = 0
        un = un.buscar
        ultnum = un.CTACTE
        Dim fechaactual As Date = Now()
        Dim _fecha As String
        _fecha = Format(fechaactual, "yyyy-MM-dd")
        '_fecha = "2018-05-31"
        Dim movimientosDict As New Dictionary(Of String, List(Of dMovimientos))
        Dim movimientos As New List(Of dMovimientos)
        Dim m As New dMovCte2
        Dim listamovimientos As New ArrayList
        listamovimientos = m.listarxid(ultnum)
        If Not listamovimientos Is Nothing Then
            For Each m In listamovimientos
                barra = barra + 1
                ProgressBar1.Value = barra
                If barra = 100 Then
                    barra = 0
                End If
                Dim movi As New dMovimientos
                Dim _path As String = ""
                Dim path As String = "/home/colaveco/public_html/gestor/facturas/"
                movi.idnet_movimiento = m.MCCNRO
                movi.fecha_emision = m.MCCFCH
                movi.fecha_vencimiento = m.MCCVTO
                Dim tipo As String = ""
                If m.DOCCOD = "NF" Then
                    tipo = "NC"
                ElseIf m.DOCCOD = "NI" Then
                    tipo = "NC"
                ElseIf m.DOCCOD = "01" Then
                    tipo = "R"
                ElseIf m.DOCCOD = "02" Then
                    tipo = "ND"
                ElseIf m.DOCCOD = "AA" Then
                    tipo = "AA"
                ElseIf m.DOCCOD = "AD" Then
                    tipo = "AD"
                ElseIf m.DOCCOD = "CI" Then
                    tipo = "F"
                ElseIf m.DOCCOD = "FA" Then
                    tipo = "F"
                ElseIf m.DOCCOD = "FF" Then
                    tipo = "F"
                ElseIf m.DOCCOD = "NC" Then
                    tipo = "NC"
                End If
                movi.tipo = tipo
                movi.numero = m.MCCDOC
                movi.detalle = m.MCCDES
                movi.importe = m.MCCIMP
                movi.tipo_movimiento = m.MCCCOD
                movi.importe_pago = m.MCCPAG
                movi.idnet_usuario = m.CLICOD
                If m.MCCTIP = "V" Then
                    Dim f As New dFactur
                    f.FACNRO = m.MCCCMP
                    f = f.buscar
                    If Not f Is Nothing Then
                        Dim c As String = "\\192.168.1.106\apls\soft\"
                        Dim tx As String = f.FACPDF
                        If tx.Contains(c) Then
                            _path = tx.Replace(c, "")
                            factura_origen = _path
                            path = path & _path
                        End If
                    End If
                End If
                movi.path_pdf = path
                movimientos.Add(movi)
                movi = Nothing
                tipo = Nothing
                subirFacturaPdf()
                un.CTACTE = m.MCCNRO
                un.modificarctacte()
            Next
            If m.MCCDOC <> 0 Then
                movimientosDict.Add("movimientos", movimientos)
                Dim parameters As String = JsonConvert.SerializeObject(movimientosDict, Formatting.None)
                Dim status As HttpStatusCode = HttpStatusCode.ExpectationFailed
                Dim response As Byte() = PostResponse("http://colaveco-gestor.herokuapp.com/factura_movimientos/migrar", "POST", parameters, status)
            End If
        End If
    End Sub
    Private Sub robot_importar()
        'No se usa /hardcode
        borrarimportacionescalidad()
        borrarimportacionescontrol()
        calidadcsv()
        calidadxls()
        calidadfat()
        ibc()
        controllecherocsv()
        controllecheroxls()
        controllecherofat()
    End Sub
    Private Sub robot_preinforme_calidad(ByVal id_sol As Long)
        'No se usa /hardcode
        Dim x1app As Microsoft.Office.Interop.Excel.Application
        Dim x1libro As Microsoft.Office.Interop.Excel.Workbook
        Dim x1hoja As Microsoft.Office.Interop.Excel.Worksheet
        x1app = CType(CreateObject("Excel.Application"), Microsoft.Office.Interop.Excel.Application)
        x1libro = CType(x1app.Workbooks.Add, Microsoft.Office.Interop.Excel.Workbook)
        x1hoja = CType(x1libro.Worksheets(1), Microsoft.Office.Interop.Excel.Worksheet)
        Dim csm As New dCalidadSolicitudMuestra
        Dim i As New dIbc
        Dim sa As New dSolicitudAnalisis
        Dim pro As New dCliente
        Dim tec As New dTecnicos
        Dim lista As New ArrayList
        '*****************************
        Dim idsol As Long = id_sol 'TextFicha.Text.Trim
        sa.ID = idsol
        sa = sa.buscar
        '*****************************
        Dim fila As Integer
        Dim columna As Integer
        fila = 1
        columna = 2
        '*********************** ENCABEZADO ************************************************************************
        '***********************************************************************************************************
        x1hoja.Cells(1, 1).columnwidth = 6 '7
        x1hoja.Cells(1, 2).columnwidth = 5
        x1hoja.Cells(1, 3).columnwidth = 5
        x1hoja.Cells(1, 4).columnwidth = 5
        x1hoja.Cells(1, 5).columnwidth = 5
        x1hoja.Cells(1, 6).columnwidth = 5
        x1hoja.Cells(1, 7).columnwidth = 5
        x1hoja.Cells(1, 8).columnwidth = 5
        x1hoja.Cells(1, 9).columnwidth = 5
        x1hoja.Cells(1, 10).columnwidth = 5
        x1hoja.Cells(1, 11).columnwidth = 6 '8
        x1hoja.Cells(1, 12).columnwidth = 6
        x1hoja.Cells(1, 13).columnwidth = 6 '8
        x1hoja.Cells(1, 14).columnwidth = 5
        lista = csm.listarporsolicitud(idsol)
        fila = 15
        columna = 1
        '*** FIN DEL ENCABEZADO************************************************************************************
        '**********************************************************************************************************
        x1hoja.Cells(fila, columna).Formula = "Ident."
        x1hoja.Cells(fila, columna).Font.Bold = True
        x1hoja.Cells(fila, columna).Font.Size = 8
        x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
        x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        columna = columna + 1
        x1hoja.Cells(fila, columna).Formula = "Rc"
        x1hoja.Cells(fila, columna).Font.Bold = True
        x1hoja.Cells(fila, columna).Font.Size = 8
        x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
        x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        columna = columna + 1
        x1hoja.Cells(fila, columna).Formula = "R Bact."
        x1hoja.Cells(fila, columna).Font.Bold = True
        x1hoja.Cells(fila, columna).Font.Size = 8
        x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
        x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        columna = columna + 1
        x1hoja.Cells(fila, columna).Formula = "Gr"
        x1hoja.Cells(fila, columna).Font.Bold = True
        x1hoja.Cells(fila, columna).Font.Size = 8
        x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
        x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        columna = columna + 1
        x1hoja.Cells(fila, columna).Formula = "Pr"
        x1hoja.Cells(fila, columna).Font.Bold = True
        x1hoja.Cells(fila, columna).Font.Size = 8
        x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
        x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        columna = columna + 1
        x1hoja.Cells(fila, columna).Formula = "Lc*"
        x1hoja.Cells(fila, columna).Font.Bold = True
        x1hoja.Cells(fila, columna).Font.Size = 8
        x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
        x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        columna = columna + 1
        x1hoja.Cells(fila, columna).Formula = "ST"
        x1hoja.Cells(fila, columna).Font.Bold = True
        x1hoja.Cells(fila, columna).Font.Size = 8
        x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
        x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        columna = columna + 1
        x1hoja.Cells(fila, columna).Formula = "Cr*"
        x1hoja.Cells(fila, columna).Font.Bold = True
        x1hoja.Cells(fila, columna).Font.Size = 8
        x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
        x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        columna = columna + 1
        x1hoja.Cells(fila, columna).Formula = "MUN*"
        x1hoja.Cells(fila, columna).Font.Bold = True
        x1hoja.Cells(fila, columna).Font.Size = 8
        x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
        x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        columna = columna + 1
        x1hoja.Cells(fila, columna).Formula = "Inh"
        x1hoja.Cells(fila, columna).Font.Bold = True
        x1hoja.Cells(fila, columna).Font.Size = 8
        x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
        x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        columna = columna + 1
        x1hoja.Cells(fila, columna).Formula = "Esp.Ana.*"
        x1hoja.Cells(fila, columna).Font.Bold = True
        x1hoja.Cells(fila, columna).Font.Size = 8
        x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
        x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        columna = columna + 1
        x1hoja.Cells(fila, columna).Formula = "Psicro.*"
        x1hoja.Cells(fila, columna).Font.Bold = True
        x1hoja.Cells(fila, columna).Font.Size = 8
        x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
        x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        columna = columna + 1
        x1hoja.Cells(fila, columna).Formula = "Caseína*"
        x1hoja.Cells(fila, columna).Font.Bold = True
        x1hoja.Cells(fila, columna).Font.Size = 8
        x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
        x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        columna = columna + 1
        x1hoja.Cells(fila, columna).Formula = "Afla.M1*"
        x1hoja.Cells(fila, columna).Font.Bold = True
        x1hoja.Cells(fila, columna).Font.Size = 8
        x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
        x1hoja.Cells(fila, columna).BORDERS.color = RGB(0, 0, 0)
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        columna = 1
        fila = fila + 1
        x1hoja.Cells(1, 1).RowHeight = 18
        x1hoja.Range("A16", "A17").Merge()
        x1hoja.Range("A16", "A17").Borders.Color = RGB(0, 0, 0)
        x1hoja.Range("A16", "A17").Interior.Color = RGB(192, 192, 192)
        x1hoja.Range("A16", "A17").WrapText = True
        x1hoja.Cells(fila, columna).formula = ""
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
        x1hoja.Cells(fila, columna).Font.Size = 6
        columna = columna + 1
        x1hoja.Cells(1, 1).RowHeight = 18
        x1hoja.Range("B16", "B17").Merge()
        x1hoja.Range("B16", "B17").Borders.Color = RGB(0, 0, 0)
        x1hoja.Range("B16", "B17").Interior.Color = RGB(192, 192, 192)
        x1hoja.Range("B16", "B17").WrapText = True
        x1hoja.Cells(fila, columna).formula = "x 1.000 cel/mL"
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
        x1hoja.Cells(fila, columna).Font.Size = 6
        columna = columna + 1
        x1hoja.Cells(1, 1).RowHeight = 18
        x1hoja.Range("C16", "C17").Merge()
        x1hoja.Range("C16", "C17").Borders.Color = RGB(0, 0, 0)
        x1hoja.Range("C16", "C17").Interior.Color = RGB(192, 192, 192)
        x1hoja.Range("C16", "C17").WrapText = True
        x1hoja.Cells(fila, columna).formula = "x 1.000 eq. UFC/ml"
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
        x1hoja.Cells(fila, columna).Font.Size = 6
        columna = columna + 1
        x1hoja.Cells(1, 1).RowHeight = 18
        x1hoja.Range("D16", "D17").Merge()
        x1hoja.Range("D16", "D17").Borders.Color = RGB(0, 0, 0)
        x1hoja.Range("D16", "D17").Interior.Color = RGB(192, 192, 192)
        x1hoja.Range("D16", "D17").WrapText = True
        x1hoja.Cells(fila, columna).formula = "% peso/vol"
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
        x1hoja.Cells(fila, columna).Font.Size = 6
        columna = columna + 1
        x1hoja.Cells(1, 1).RowHeight = 18
        x1hoja.Range("E16", "E17").Merge()
        x1hoja.Range("E16", "E17").Borders.Color = RGB(0, 0, 0)
        x1hoja.Range("E16", "E17").Interior.Color = RGB(192, 192, 192)
        x1hoja.Range("E16", "E17").WrapText = True
        x1hoja.Cells(fila, columna).formula = "% peso/vol"
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
        x1hoja.Cells(fila, columna).Font.Size = 6
        columna = columna + 1
        x1hoja.Cells(1, 1).RowHeight = 18
        x1hoja.Range("F16", "F17").Merge()
        x1hoja.Range("F16", "F17").Borders.Color = RGB(0, 0, 0)
        x1hoja.Range("F16", "F17").Interior.Color = RGB(192, 192, 192)
        x1hoja.Range("F16", "F17").WrapText = True
        x1hoja.Cells(fila, columna).formula = "% peso/vol"
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
        x1hoja.Cells(fila, columna).Font.Size = 6
        columna = columna + 1
        x1hoja.Cells(1, 1).RowHeight = 18
        x1hoja.Range("G16", "G17").Merge()
        x1hoja.Range("G16", "G17").Borders.Color = RGB(0, 0, 0)
        x1hoja.Range("G16", "G17").Interior.Color = RGB(192, 192, 192)
        x1hoja.Range("G16", "G17").WrapText = True
        x1hoja.Cells(fila, columna).formula = "% peso/vol"
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
        x1hoja.Cells(fila, columna).Font.Size = 6
        columna = columna + 1
        x1hoja.Cells(1, 1).RowHeight = 18
        x1hoja.Range("H16", "H17").Merge()
        x1hoja.Range("H16", "H17").Borders.Color = RGB(0, 0, 0)
        x1hoja.Range("H16", "H17").Interior.Color = RGB(192, 192, 192)
        x1hoja.Range("H16", "H17").WrapText = True
        x1hoja.Cells(fila, columna).formula = "(ºC)"
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
        x1hoja.Cells(fila, columna).Font.Size = 6
        columna = columna + 1
        x1hoja.Cells(1, 1).RowHeight = 18
        x1hoja.Range("I16", "I17").Merge()
        x1hoja.Range("I16", "I17").Borders.Color = RGB(0, 0, 0)
        x1hoja.Range("I16", "I17").Interior.Color = RGB(192, 192, 192)
        x1hoja.Range("I16", "I17").WrapText = True
        x1hoja.Cells(fila, columna).formula = "mg/dl"
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
        x1hoja.Cells(fila, columna).Font.Size = 6
        columna = columna + 1
        x1hoja.Cells(1, 1).RowHeight = 18
        x1hoja.Range("J16", "J17").Merge()
        x1hoja.Range("J16", "J17").Borders.Color = RGB(0, 0, 0)
        x1hoja.Range("J16", "J17").Interior.Color = RGB(192, 192, 192)
        x1hoja.Range("J16", "J17").WrapText = True
        x1hoja.Cells(fila, columna).formula = ""
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
        x1hoja.Cells(fila, columna).Font.Size = 6
        columna = columna + 1
        x1hoja.Cells(1, 1).RowHeight = 18
        x1hoja.Range("K16", "K17").Merge()
        x1hoja.Range("K16", "K17").Borders.Color = RGB(0, 0, 0)
        x1hoja.Range("K16", "K17").Interior.Color = RGB(192, 192, 192)
        x1hoja.Range("K16", "K17").WrapText = True
        x1hoja.Cells(fila, columna).formula = "NMP/L"
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
        x1hoja.Cells(fila, columna).Font.Size = 6
        columna = columna + 1
        x1hoja.Cells(1, 1).RowHeight = 18
        x1hoja.Range("L16", "L17").Merge()
        x1hoja.Range("L16", "L17").Borders.Color = RGB(0, 0, 0)
        x1hoja.Range("L16", "L17").Interior.Color = RGB(192, 192, 192)
        x1hoja.Range("L16", "L17").WrapText = True
        x1hoja.Cells(fila, columna).formula = "x 1000 UFC/ml UFC/mL "
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
        x1hoja.Cells(fila, columna).Font.Size = 6
        columna = columna + 1
        x1hoja.Cells(1, 1).RowHeight = 18
        x1hoja.Range("M16", "M17").Merge()
        x1hoja.Range("M16", "M17").Borders.Color = RGB(0, 0, 0)
        x1hoja.Range("M16", "M17").Interior.Color = RGB(192, 192, 192)
        x1hoja.Range("M16", "M17").WrapText = True
        x1hoja.Cells(fila, columna).formula = ""
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
        x1hoja.Cells(fila, columna).Font.Size = 6
        columna = columna + 1
        x1hoja.Cells(1, 1).RowHeight = 18
        x1hoja.Range("N16", "N17").Merge()
        x1hoja.Range("N16", "N17").Borders.Color = RGB(0, 0, 0)
        x1hoja.Range("N16", "N17").Interior.Color = RGB(192, 192, 192)
        x1hoja.Range("N16", "N17").WrapText = True
        x1hoja.Cells(fila, columna).formula = "ppb"
        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
        x1hoja.Cells(fila, columna).VerticalAlignment = XlVAlign.xlVAlignTop
        x1hoja.Cells(fila, columna).Font.Size = 6
        columna = 1
        fila = fila + 2
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each csm In lista
                    Dim c As New dCalidad
                    c.FICHA = idsol
                    c.MUESTRA = Trim(csm.MUESTRA)
                    c = c.buscarxfichaxmuestra
                    If csm.MUESTRA <> "" Then
                        x1hoja.Cells(fila, columna).formula = Trim(csm.MUESTRA)
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    Else
                        x1hoja.Cells(fila, columna).formula = "-"
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    End If
                    If csm.RC = 1 Then
                        If Not c Is Nothing Then
                            x1hoja.Cells(fila, columna).formula = c.RC
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                            If c.RC < 100 Then
                                'contador_rc = contador_rc + 1
                            End If
                        Else
                            x1hoja.Cells(fila, columna).formula = "-"
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        End If
                    Else
                        x1hoja.Cells(fila, columna).formula = "-"
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    End If
                    If csm.RB = 1 Then
                        Dim ibc As New dIbc
                        ibc.FICHA = idsol
                        ibc.MUESTRA = Trim(csm.MUESTRA)
                        ibc = ibc.buscarxfichaxmuestra
                        If Not ibc Is Nothing Then
                            x1hoja.Cells(fila, columna).formula = ibc.RB
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        Else
                            x1hoja.Cells(fila, columna).formula = "-"
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        End If
                    Else
                        x1hoja.Cells(fila, columna).formula = "-"
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    End If
                    If csm.COMPOSICION = 1 Or csm.COMPOSICIONSUERO = 1 Then
                        If Not c Is Nothing Then
                            Dim valgrasa As Double = Val(c.GRASA)
                            If valgrasa < 2 Or valgrasa > 4.5 Then
                                x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
                            End If
                            x1hoja.Cells(fila, columna).formula = c.GRASA
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        Else
                            x1hoja.Cells(fila, columna).formula = "-"
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        End If
                    Else
                        x1hoja.Cells(fila, columna).formula = "-"
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    End If
                    If csm.COMPOSICION = 1 Or csm.COMPOSICIONSUERO = 1 Then
                        If Not c Is Nothing Then
                            Dim valproteina As Double = Val(c.PROTEINA)
                            If valproteina < 2 Or valproteina > 3.8 Then
                                x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
                            End If
                            x1hoja.Cells(fila, columna).formula = c.PROTEINA
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        Else
                            x1hoja.Cells(fila, columna).formula = "-"
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        End If
                    Else
                        x1hoja.Cells(fila, columna).formula = "-"
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    End If
                    If csm.COMPOSICION = 1 Or csm.COMPOSICIONSUERO = 1 Then
                        If Not c Is Nothing Then
                            x1hoja.Cells(fila, columna).formula = c.LACTOSA
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        Else
                            x1hoja.Cells(fila, columna).formula = "-"
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        End If
                    Else
                        x1hoja.Cells(fila, columna).formula = "-"
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    End If
                    If csm.COMPOSICION = 1 Or csm.COMPOSICIONSUERO = 1 Then
                        If Not c Is Nothing Then
                            x1hoja.Cells(fila, columna).formula = c.ST
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        Else
                            x1hoja.Cells(fila, columna).formula = "-"
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        End If
                    Else
                        x1hoja.Cells(fila, columna).formula = "-"
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    End If
                    If csm.CRIOSCOPIA = 1 Or csm.CRIOSCOPIA_CRIOSCOPO = 1 Then
                        If Not c Is Nothing Then
                            If c.CRIOSCOPIA <> -1 Then
                                Dim valcrioscopia As Double = Val(c.CRIOSCOPIA) * -1 / 1000
                                If valcrioscopia > -0.51 Or valcrioscopia < -0.54 Then
                                    x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
                                End If
                                x1hoja.Cells(fila, columna).formula = valcrioscopia.ToString("##,##0.000")
                                x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                x1hoja.Cells(fila, columna).Font.Size = 8
                                columna = columna + 1
                            Else
                                x1hoja.Cells(fila, columna).formula = "-"
                                x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                x1hoja.Cells(fila, columna).Font.Size = 8
                                columna = columna + 1
                            End If
                        Else
                            x1hoja.Cells(fila, columna).formula = "-"
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        End If
                    Else
                        x1hoja.Cells(fila, columna).formula = "-"
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    End If
                    If csm.UREA = 1 Then
                        If Not c Is Nothing Then
                            If c.UREA <> -1 Then
                                Dim valorurea As Integer
                                valorurea = c.UREA * 0.466
                                If valorurea > 20 Or valorurea < 9 Then
                                    x1hoja.Cells(fila, columna).interior.color = RGB(192, 192, 192)
                                End If
                                x1hoja.Cells(fila, columna).formula = valorurea
                                x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                x1hoja.Cells(fila, columna).Font.Size = 8
                                columna = columna + 1
                            Else
                                x1hoja.Cells(fila, columna).formula = "-"
                                x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                x1hoja.Cells(fila, columna).Font.Size = 8
                                columna = columna + 1
                            End If
                        Else
                            x1hoja.Cells(fila, columna).formula = "-"
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        End If
                    Else
                        x1hoja.Cells(fila, columna).formula = "-"
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    End If
                    Dim inh As New dInhibidores
                    inh.FICHA = idsol
                    inh.MUESTRA = Trim(csm.MUESTRA)
                    inh = inh.buscarxfichaxmuestra
                    If Not inh Is Nothing Then
                        If inh.RESULTADO = 0 Then
                            x1hoja.Cells(fila, columna).formula = "Negativo"
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 6
                            columna = columna + 1
                        Else
                            x1hoja.Cells(fila, columna).formula = "Positivo"
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 6
                            columna = columna + 1
                        End If
                    Else
                        x1hoja.Cells(fila, columna).formula = "-"
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    End If
                    'ESPORULADOS*******************************************************************************
                    Dim esp As New dEsporulados
                    esp.FICHA = idsol
                    esp.MUESTRA = Trim(csm.MUESTRA)
                    esp = esp.buscarxfichaxmuestra
                    If Not esp Is Nothing Then
                        x1hoja.Cells(fila, columna).formula = esp.RESULTADO
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    Else
                        x1hoja.Cells(fila, columna).formula = "-"
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    End If
                    'PSICROTROFOS*******************************************************************************
                    Dim psi As New dPsicrotrofos
                    psi.FICHA = idsol
                    psi.MUESTRA = Trim(csm.MUESTRA)
                    psi = psi.buscarxfichaxmuestra
                    If Not psi Is Nothing Then
                        x1hoja.Cells(fila, columna).formula = psi.PROMEDIO
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    Else
                        x1hoja.Cells(fila, columna).formula = "-"
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    End If
                    If csm.CASEINA = 1 Then
                        If Not c Is Nothing Then
                            If c.CASEINA <> -1 Then
                                Dim valorcaseina As Double
                                valorcaseina = c.CASEINA
                                x1hoja.Cells(fila, columna).formula = valorcaseina
                                x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                x1hoja.Cells(fila, columna).Font.Size = 8
                                columna = columna + 1
                            Else
                                x1hoja.Cells(fila, columna).formula = "-"
                                x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                                x1hoja.Cells(fila, columna).Font.Size = 8
                                columna = columna + 1
                            End If
                        Else
                            x1hoja.Cells(fila, columna).formula = "-"
                            x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                            x1hoja.Cells(fila, columna).Font.Size = 8
                            columna = columna + 1
                        End If
                    Else
                        x1hoja.Cells(fila, columna).formula = "-"
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = columna + 1
                    End If
                    'AFLATOXINA M1*******************************************************************************
                    Dim m As New dMicotoxinasLeche
                    m.FICHA = idsol
                    m.MUESTRA = Trim(csm.MUESTRA)
                    m = m.buscarxfichaxmuestra
                    If Not m Is Nothing Then
                        x1hoja.Cells(fila, columna).formula = m.RESULTADO
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = 1
                    Else
                        x1hoja.Cells(fila, columna).formula = "-"
                        x1hoja.Cells(fila, columna).HorizontalAlignment = XlHAlign.xlHAlignCenter
                        x1hoja.Cells(fila, columna).Font.Size = 8
                        columna = 1
                    End If
                    columna = 1
                    fila = fila + 1
                Next
                'Referencias
                fila = fila + 1
                columna = 1
            End If
        End If
        'PROTEGE LA HOJA DE EXCEL
        'x1hoja.Protect(Password:="1582782", DrawingObjects:=True, _
        'Contents:=True, Scenarios:=True)
        'GUARDA EL ARCHIVO DE EXCEL
        'x1hoja.SaveAs("\\SRVDATOS\datos (d)\NET\PREINFORMES\CALIDAD\" & idsol & ".xls")
        'x1hoja.PageSetup.CenterFooter = "Página &P de &N"
        x1hoja.PageSetup.CenterFooter = "Página &P"
        'x1hoja.SaveAs("\\192.168.1.10\e\NET\PREINFORMES\CALIDAD\" & idsol & ".xls")
        x1hoja.SaveAs("C:\PREINFORMES\CALIDAD\" & idsol & ".xls")
        'Marcar como creado
        Dim preinf As New dPreinformes
        preinf.FICHA = idsol
        preinf.marcarcreado()
        preinf = Nothing
        x1app.Visible = False
        x1libro.Close()
        x1app = Nothing
        x1libro = Nothing
        x1hoja = Nothing
    End Sub
    Private Sub robot_pre_informe_control()
        'No se usa /Hardcode
        Dim pi As New dPreinformes
        Dim lista As New ArrayList
        Dim id_sol As Long = 0
        lista = pi.listarsinmarcarcontrol
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each pi In lista
                    id_sol = pi.FICHA
                    preinforme_control(id_sol)
                Next
            End If
        End If
        pi = Nothing
        lista = Nothing
    End Sub
    Private Sub robot_subir_informes_control()
        'No se usa /hardcode
        Dim pi As New dPreinformes
        Dim lista As New ArrayList
        Dim subidoxls As Integer = 0
        Dim subidopdf As Integer = 0
        Dim subidotxt As Integer = 0
        lista = pi.listarparasubircontrol
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each pi In lista
                    Dim cifq As New dControlInformesFQ
                    Dim ficha As Long = 0
                    cifq.FICHA = pi.FICHA
                    cifq = cifq.buscarxficha
                    If Not cifq Is Nothing Then
                        If cifq.CONTROLADO = 1 Then
                            idficha = pi.FICHA
                            _tipoinforme = pi.TIPO
                            enviar_copia = pi.COPIA
                            _abonado = pi.ABONADO
                            _comentarios = pi.COMENTARIO
                            Dim sa As New dSolicitudAnalisis
                            sa.ID = idficha
                            sa = sa.buscar
                            If Not sa Is Nothing Then
                                Dim p As New dCliente
                                tipoinforme = sa.IDTIPOINFORME
                                p.ID = sa.IDPRODUCTOR
                                p = p.buscar
                                If Not p Is Nothing Then
                                    email = ""
                                    email2 = ""
                                    If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                        email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                    End If
                                    If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                        email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                    End If
                                    If email = "" Then
                                        If email2 = "" Then
                                            If p.EMAIL <> "" Then
                                                email = RTrim(p.EMAIL)
                                            End If
                                        Else
                                            email = email2
                                        End If
                                    Else
                                        If email2 = "" Then
                                            email = email
                                        Else
                                            email = email & "," & email2
                                        End If
                                    End If
                                    productorweb_com = p.USUARIO_WEB
                                    Dim pw_com As New dProductorWeb_com
                                    pw_com.USUARIO = productorweb_com
                                    pw_com = pw_com.buscar
                                    If Not pw_com Is Nothing Then
                                        idproductorweb_com = pw_com.ID
                                        'email = RTrim(pw_com.ENVIAR_EMAIL)
                                        celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                        carpeta = idproductorweb_com
                                        crea_carpeta()
                                    End If
                                    sa = Nothing
                                End If
                            End If
controlexcel:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroXls()
                            End If
                            existeXls()
                            If excel = 1 Then
                                GoTo controlexcel
                            End If
                            subidoxls = 1
controlpdf:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroPdf()
                            End If
                            existePdf()
                            If pdf = 1 Then
                                GoTo controlpdf
                            End If
                            subidopdf = 1
                            If pi.TIPO = 1 Then
controltxt:
                                If nombre_pc = "ROBOT" Then
                                    subirFicheroCsv()
                                End If
                                existeCsv()
                                If csv = 1 Then
                                    GoTo controltxt
                                End If
                                If nombre_pc = "ROBOT" Then
                                    movertxt()
                                Else
                                    movertxt_otrapc()
                                End If
                            End If
                            modificarRegistro()
                            Dim fechaactual As Date = Now()
                            Dim _fecha As String
                            _fecha = Format(fechaactual, "yyyy-MM-dd")
                            pi.marcarsubido(_fecha)
                            If tipoinforme = 15 Then
                                enviaremail()
                            End If
                            Dim s As New dSolicitudAnalisis
                            Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                            Dim fecenv As String
                            fecenv = Format(fechaenvio, "yyyy-MM-dd")
                            s.ID = idficha
                            s.actualizarfechaenvio2(fecenv)
                            s.marcar2()
                            s = Nothing
                            If subidoxls = 1 And subidopdf = 1 Then
                                If nombre_pc = "ROBOT" Then
                                    moverexcel()
                                    moverpdf()
                                Else
                                    moverexcel_otrapc()
                                    moverpdf_otrapc()
                                End If
                                ' Grabar estado de la ficha
                                Dim est As New dEstados
                                est.FICHA = idficha
                                est.ESTADO = 8
                                est.FECHA = fecenv
                                est.guardar2()
                                est = Nothing
                                '****************************
                                envio_mail_informes()
                            End If
                        Else
                        End If
                    Else
                        idficha = pi.FICHA
                        _tipoinforme = pi.TIPO
                        enviar_copia = pi.COPIA
                        _abonado = pi.ABONADO
                        _comentarios = pi.COMENTARIO
                        Dim sa As New dSolicitudAnalisis
                        sa.ID = idficha
                        sa = sa.buscar
                        If Not sa Is Nothing Then
                            Dim p As New dCliente
                            tipoinforme = sa.IDTIPOINFORME
                            p.ID = sa.IDPRODUCTOR
                            p = p.buscar
                            If Not p Is Nothing Then
                                email = ""
                                email2 = ""
                                If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                    email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                End If
                                If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                    email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                End If
                                If email = "" Then
                                    If email2 = "" Then
                                        If p.EMAIL <> "" Then
                                            email = RTrim(p.EMAIL)
                                        End If
                                    Else
                                        email = email2
                                    End If
                                Else
                                    If email2 = "" Then
                                        email = email
                                    Else
                                        email = email & "," & email2
                                    End If
                                End If
                                productorweb_com = p.USUARIO_WEB
                                Dim pw_com As New dProductorWeb_com
                                pw_com.USUARIO = productorweb_com
                                pw_com = pw_com.buscar
                                If Not pw_com Is Nothing Then
                                    idproductorweb_com = pw_com.ID
                                    'email = RTrim(pw_com.ENVIAR_EMAIL)
                                    celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                    carpeta = idproductorweb_com
                                    crea_carpeta()
                                End If
                                sa = Nothing
                            End If
                        End If
controlexcel2:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroXls()
                        End If
                        existeXls()
                        If excel = 1 Then
                            GoTo controlexcel2
                        End If
                        subidoxls = 1
controlpdf2:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroPdf()
                        End If
                        existePdf()
                        If pdf = 1 Then
                            GoTo controlpdf2
                        End If
                        subidopdf = 1
                        If pi.TIPO = 1 Then
controltxt2:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroCsv()
                            End If
                            existeCsv()
                            If csv = 1 Then
                                GoTo controltxt2
                            End If
                            If nombre_pc = "ROBOT" Then
                                movertxt()
                            Else
                                movertxt_otrapc()
                            End If
                        End If
                        modificarRegistro()
                        Dim fechaactual As Date = Now()
                        Dim _fecha As String
                        _fecha = Format(fechaactual, "yyyy-MM-dd")
                        pi.marcarsubido(_fecha)
                        If tipoinforme = 15 Then
                            enviaremail()
                        End If
                        Dim s As New dSolicitudAnalisis
                        Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                        Dim fecenv As String
                        fecenv = Format(fechaenvio, "yyyy-MM-dd")
                        s.ID = idficha
                        s.actualizarfechaenvio2(fecenv)
                        s.marcar2()
                        s = Nothing
                        If subidoxls = 1 And subidopdf = 1 Then
                            If nombre_pc = "ROBOT" Then
                                moverexcel()
                                moverpdf()
                            Else
                                moverexcel_otrapc()
                                moverpdf_otrapc()
                            End If
                            ' Grabar estado de la ficha
                            Dim est As New dEstados
                            est.FICHA = idficha
                            est.ESTADO = 8
                            est.FECHA = fecenv
                            est.guardar2()
                            est = Nothing
                            '****************************
                            envio_mail_informes()
                        End If
                    End If
                    cifq = Nothing
                Next
            End If
        End If
    End Sub
    Private Sub robot_subir_informes_agua()
        'No se usa /hadcode
        Dim pi As New dPreinformes
        Dim lista As New ArrayList
        Dim subidoxls As Integer = 0
        Dim subidopdf As Integer = 0
        Dim subidotxt As Integer = 0
        lista = pi.listarparasubiragua
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each pi In lista
                    Dim cimicro As New dControlInformesMicro
                    Dim ficha As Long = 0
                    cimicro.FICHA = pi.FICHA
                    cimicro = cimicro.buscarxficha
                    If Not cimicro Is Nothing Then
                        If cimicro.CONTROLADO = 1 Then
                            idficha = pi.FICHA
                            _tipoinforme = pi.TIPO
                            enviar_copia = pi.COPIA
                            _abonado = pi.ABONADO
                            _comentarios = pi.COMENTARIO
                            Dim sa As New dSolicitudAnalisis
                            sa.ID = idficha
                            sa = sa.buscar
                            If Not sa Is Nothing Then
                                Dim p As New dCliente
                                tipoinforme = sa.IDTIPOINFORME
                                p.ID = sa.IDPRODUCTOR
                                p = p.buscar
                                If Not p Is Nothing Then
                                    email = ""
                                    email2 = ""
                                    If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                        email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                    End If
                                    If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                        email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                    End If
                                    If email = "" Then
                                        If email2 = "" Then
                                            If p.EMAIL <> "" Then
                                                email = RTrim(p.EMAIL)
                                            End If
                                        Else
                                            email = email2
                                        End If
                                    Else
                                        If email2 = "" Then
                                            email = email
                                        Else
                                            email = email & "," & email2
                                        End If
                                    End If
                                    productorweb_com = p.USUARIO_WEB
                                    Dim pw_com As New dProductorWeb_com
                                    pw_com.USUARIO = productorweb_com
                                    pw_com = pw_com.buscar
                                    If Not pw_com Is Nothing Then
                                        idproductorweb_com = pw_com.ID
                                        'email = RTrim(pw_com.ENVIAR_EMAIL)
                                        celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                        carpeta = idproductorweb_com
                                        crea_carpeta()
                                    End If
                                    sa = Nothing
                                End If
                            End If
controlexcel:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroXls()
                            End If
                            existeXls()
                            If excel = 1 Then
                                GoTo controlexcel
                            End If
                            subidoxls = 1
controlpdf:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroPdf()
                            End If
                            existePdf()
                            If pdf = 1 Then
                                GoTo controlpdf
                            End If
                            subidopdf = 1
                            If pi.TIPO = 1 Then
controltxt:
                                If nombre_pc = "ROBOT" Then
                                    subirFicheroCsv()
                                End If
                                existeCsv()
                                If csv = 1 Then
                                    GoTo controltxt
                                End If
                                movertxt()
                            End If
                            modificarRegistro()
                            Dim fechaactual As Date = Now()
                            Dim _fecha As String
                            _fecha = Format(fechaactual, "yyyy-MM-dd")
                            pi.marcarsubido(_fecha)
                            If tipoinforme = 15 Then
                                enviaremail()
                            End If
                            Dim s As New dSolicitudAnalisis
                            Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                            Dim fecenv As String
                            fecenv = Format(fechaenvio, "yyyy-MM-dd")
                            s.ID = idficha
                            s.actualizarfechaenvio2(fecenv)
                            s.marcar2()
                            s = Nothing
                            If subidoxls = 1 And subidopdf = 1 Then
                                If nombre_pc = "ROBOT" Then
                                    moverexcel()
                                    moverpdf()
                                Else
                                    moverexcel_otrapc()
                                    moverpdf_otrapc()
                                End If
                                ' Grabar estado de la ficha
                                Dim est As New dEstados
                                est.FICHA = idficha
                                est.ESTADO = 8
                                est.FECHA = fecenv
                                est.guardar2()
                                est = Nothing
                                '****************************
                                envio_mail_informes()
                            End If
                        Else
                        End If
                    Else
                        idficha = pi.FICHA
                        _tipoinforme = pi.TIPO
                        enviar_copia = pi.COPIA
                        _abonado = pi.ABONADO
                        _comentarios = pi.COMENTARIO
                        Dim sa As New dSolicitudAnalisis
                        sa.ID = idficha
                        sa = sa.buscar
                        If Not sa Is Nothing Then
                            Dim p As New dCliente
                            tipoinforme = sa.IDTIPOINFORME
                            p.ID = sa.IDPRODUCTOR
                            p = p.buscar
                            If Not p Is Nothing Then
                                email = ""
                                email2 = ""
                                If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                    email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                End If
                                If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                    email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                End If
                                If email = "" Then
                                    If email2 = "" Then
                                        If p.EMAIL <> "" Then
                                            email = RTrim(p.EMAIL)
                                        End If
                                    Else
                                        email = email2
                                    End If
                                Else
                                    If email2 = "" Then
                                        email = email
                                    Else
                                        email = email & "," & email2
                                    End If
                                End If
                                productorweb_com = p.USUARIO_WEB
                                Dim pw_com As New dProductorWeb_com
                                pw_com.USUARIO = productorweb_com
                                pw_com = pw_com.buscar
                                If Not pw_com Is Nothing Then
                                    idproductorweb_com = pw_com.ID
                                    'email = RTrim(pw_com.ENVIAR_EMAIL)
                                    celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                    carpeta = idproductorweb_com
                                    crea_carpeta()
                                End If
                                sa = Nothing
                            End If
                        End If
controlexcel2:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroXls()
                        End If
                        existeXls()
                        If excel = 1 Then
                            GoTo controlexcel2
                        End If
                        subidoxls = 1
controlpdf2:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroPdf()
                        End If
                        existePdf()
                        If pdf = 1 Then
                            GoTo controlpdf2
                        End If
                        subidopdf = 1
                        If pi.TIPO = 1 Then
controltxt2:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroCsv()
                            End If
                            existeCsv()
                            If csv = 1 Then
                                GoTo controltxt2
                            End If
                            movertxt()
                        End If
                        modificarRegistro()
                        Dim fechaactual As Date = Now()
                        Dim _fecha As String
                        _fecha = Format(fechaactual, "yyyy-MM-dd")
                        pi.marcarsubido(_fecha)
                        If tipoinforme = 15 Then
                            enviaremail()
                        End If
                        Dim s As New dSolicitudAnalisis
                        Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                        Dim fecenv As String
                        fecenv = Format(fechaenvio, "yyyy-MM-dd")
                        s.ID = idficha
                        s.actualizarfechaenvio2(fecenv)
                        s.marcar2()
                        s = Nothing
                        If subidoxls = 1 And subidopdf = 1 Then
                            If nombre_pc = "ROBOT" Then
                                moverexcel()
                                moverpdf()
                            Else
                                moverexcel_otrapc()
                                moverpdf_otrapc()
                            End If
                            ' Grabar estado de la ficha
                            Dim est As New dEstados
                            est.FICHA = idficha
                            est.ESTADO = 8
                            est.FECHA = fecenv
                            est.guardar2()
                            est = Nothing
                            '****************************
                            envio_mail_informes()
                        End If
                    End If
                    cimicro = Nothing
                Next
            End If
        End If
    End Sub
    Private Sub robot_subir_informes_efluentes()
        'No se usa /hardcode
        Dim pi As New dPreinformes
        Dim lista As New ArrayList
        Dim subidoxls As Integer = 0
        Dim subidopdf As Integer = 0
        Dim subidotxt As Integer = 0
        lista = pi.listarparasubirefluentes
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each pi In lista
                    Dim cimicro As New dControlInformesMicro
                    Dim ficha As Long = 0
                    cimicro.FICHA = pi.FICHA
                    cimicro = cimicro.buscarxficha
                    If Not cimicro Is Nothing Then
                        If cimicro.CONTROLADO = 1 Then
                            idficha = pi.FICHA
                            _tipoinforme = pi.TIPO
                            enviar_copia = pi.COPIA
                            _abonado = pi.ABONADO
                            _comentarios = pi.COMENTARIO
                            Dim sa As New dSolicitudAnalisis
                            sa.ID = idficha
                            sa = sa.buscar
                            If Not sa Is Nothing Then
                                Dim p As New dCliente
                                tipoinforme = sa.IDTIPOINFORME
                                p.ID = sa.IDPRODUCTOR
                                p = p.buscar
                                If Not p Is Nothing Then
                                    email = ""
                                    email2 = ""
                                    If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                        email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                    End If
                                    If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                        email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                    End If
                                    If email = "" Then
                                        If email2 = "" Then
                                            If p.EMAIL <> "" Then
                                                email = RTrim(p.EMAIL)
                                            End If
                                        Else
                                            email = email2
                                        End If
                                    Else
                                        If email2 = "" Then
                                            email = email
                                        Else
                                            email = email & "," & email2
                                        End If
                                    End If
                                    productorweb_com = p.USUARIO_WEB
                                    Dim pw_com As New dProductorWeb_com
                                    pw_com.USUARIO = productorweb_com
                                    pw_com = pw_com.buscar
                                    If Not pw_com Is Nothing Then
                                        idproductorweb_com = pw_com.ID
                                        'email = RTrim(pw_com.ENVIAR_EMAIL)
                                        celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                        carpeta = idproductorweb_com
                                        crea_carpeta()
                                    End If
                                    sa = Nothing
                                End If
                            End If
controlexcel:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroXls()
                            End If
                            existeXls()
                            If excel = 1 Then
                                GoTo controlexcel
                            End If
                            subidoxls = 1
controlpdf:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroPdf()
                            End If
                            existePdf()
                            If pdf = 1 Then
                                GoTo controlpdf
                            End If
                            subidopdf = 1
                            If pi.TIPO = 1 Then
controltxt:
                                If nombre_pc = "ROBOT" Then
                                    subirFicheroCsv()
                                End If
                                existeCsv()
                                If csv = 1 Then
                                    GoTo controltxt
                                End If
                                movertxt()
                            End If
                            modificarRegistro()
                            Dim fechaactual As Date = Now()
                            Dim _fecha As String
                            _fecha = Format(fechaactual, "yyyy-MM-dd")
                            pi.marcarsubido(_fecha)
                            If tipoinforme = 15 Then
                                enviaremail()
                            End If
                            Dim s As New dSolicitudAnalisis
                            Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                            Dim fecenv As String
                            fecenv = Format(fechaenvio, "yyyy-MM-dd")
                            s.ID = idficha
                            s.actualizarfechaenvio2(fecenv)
                            s.marcar2()
                            s = Nothing
                            If subidoxls = 1 And subidopdf = 1 Then
                                If nombre_pc = "ROBOT" Then
                                    moverexcel()
                                    moverpdf()
                                Else
                                    moverexcel_otrapc()
                                    moverpdf_otrapc()
                                End If
                                ' Grabar estado de la ficha
                                Dim est As New dEstados
                                est.FICHA = idficha
                                est.ESTADO = 8
                                est.FECHA = fecenv
                                est.guardar2()
                                est = Nothing
                                '****************************
                                envio_mail_informes()
                            End If
                        Else
                        End If
                    Else
                        idficha = pi.FICHA
                        _tipoinforme = pi.TIPO
                        enviar_copia = pi.COPIA
                        _abonado = pi.ABONADO
                        _comentarios = pi.COMENTARIO
                        Dim sa As New dSolicitudAnalisis
                        sa.ID = idficha
                        sa = sa.buscar
                        If Not sa Is Nothing Then
                            Dim p As New dCliente
                            tipoinforme = sa.IDTIPOINFORME
                            p.ID = sa.IDPRODUCTOR
                            p = p.buscar
                            If Not p Is Nothing Then
                                email = ""
                                email2 = ""
                                If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                    email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                End If
                                If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                    email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                End If
                                If email = "" Then
                                    If email2 = "" Then
                                        If p.EMAIL <> "" Then
                                            email = RTrim(p.EMAIL)
                                        End If
                                    Else
                                        email = email2
                                    End If
                                Else
                                    If email2 = "" Then
                                        email = email
                                    Else
                                        email = email & "," & email2
                                    End If
                                End If
                                productorweb_com = p.USUARIO_WEB
                                Dim pw_com As New dProductorWeb_com
                                pw_com.USUARIO = productorweb_com
                                pw_com = pw_com.buscar
                                If Not pw_com Is Nothing Then
                                    idproductorweb_com = pw_com.ID
                                    'email = RTrim(pw_com.ENVIAR_EMAIL)
                                    celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                    carpeta = idproductorweb_com
                                    crea_carpeta()
                                End If
                                sa = Nothing
                            End If
                        End If
controlexcel2:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroXls()
                        End If
                        existeXls()
                        If excel = 1 Then
                            GoTo controlexcel2
                        End If
                        subidoxls = 1
controlpdf2:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroPdf()
                        End If
                        existePdf()
                        If pdf = 1 Then
                            GoTo controlpdf2
                        End If
                        subidopdf = 1
                        If pi.TIPO = 1 Then
controltxt2:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroCsv()
                            End If
                            existeCsv()
                            If csv = 1 Then
                                GoTo controltxt2
                            End If
                            movertxt()
                        End If
                        modificarRegistro()
                        Dim fechaactual As Date = Now()
                        Dim _fecha As String
                        _fecha = Format(fechaactual, "yyyy-MM-dd")
                        pi.marcarsubido(_fecha)
                        If tipoinforme = 15 Then
                            enviaremail()
                        End If
                        Dim s As New dSolicitudAnalisis
                        Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                        Dim fecenv As String
                        fecenv = Format(fechaenvio, "yyyy-MM-dd")
                        s.ID = idficha
                        s.actualizarfechaenvio2(fecenv)
                        s.marcar2()
                        s = Nothing
                        If subidoxls = 1 And subidopdf = 1 Then
                            If nombre_pc = "ROBOT" Then
                                moverexcel()
                                moverpdf()
                            Else
                                moverexcel_otrapc()
                                moverpdf_otrapc()
                            End If
                            ' Grabar estado de la ficha
                            Dim est As New dEstados
                            est.FICHA = idficha
                            est.ESTADO = 8
                            est.FECHA = fecenv
                            est.guardar2()
                            est = Nothing
                            '****************************
                            envio_mail_informes()
                        End If
                    End If
                    cimicro = Nothing
                Next
            End If
        End If
    End Sub
    Private Sub robot_subir_informes_atb()
        'No se usa /hardcode
        Dim pi As New dPreinformes
        Dim lista As New ArrayList
        Dim subidoxls As Integer = 0
        Dim subidopdf As Integer = 0
        Dim subidotxt As Integer = 0
        lista = pi.listarparasubiratb
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each pi In lista
                    Dim cimicro As New dControlInformesMicro
                    Dim ficha As Long = 0
                    cimicro.FICHA = pi.FICHA
                    cimicro = cimicro.buscarxficha
                    If Not cimicro Is Nothing Then
                        If cimicro.CONTROLADO = 1 Then
                            idficha = pi.FICHA
                            _tipoinforme = pi.TIPO
                            enviar_copia = pi.COPIA
                            _abonado = pi.ABONADO
                            _comentarios = pi.COMENTARIO
                            Dim sa As New dSolicitudAnalisis
                            sa.ID = idficha
                            sa = sa.buscar
                            If Not sa Is Nothing Then
                                Dim p As New dCliente
                                tipoinforme = sa.IDTIPOINFORME
                                p.ID = sa.IDPRODUCTOR
                                p = p.buscar
                                If Not p Is Nothing Then
                                    email = ""
                                    email2 = ""
                                    If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                        email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                    End If
                                    If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                        email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                    End If
                                    If email = "" Then
                                        If email2 = "" Then
                                            If p.EMAIL <> "" Then
                                                email = RTrim(p.EMAIL)
                                            End If
                                        Else
                                            email = email2
                                        End If
                                    Else
                                        If email2 = "" Then
                                            email = email
                                        Else
                                            email = email & "," & email2
                                        End If
                                    End If
                                    productorweb_com = p.USUARIO_WEB
                                    Dim pw_com As New dProductorWeb_com
                                    pw_com.USUARIO = productorweb_com
                                    pw_com = pw_com.buscar
                                    If Not pw_com Is Nothing Then
                                        idproductorweb_com = pw_com.ID
                                        'email = RTrim(pw_com.ENVIAR_EMAIL)
                                        celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                        carpeta = idproductorweb_com
                                        crea_carpeta()
                                    End If
                                    sa = Nothing
                                End If
                            End If
controlexcel:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroXls()
                            End If
                            existeXls()
                            If excel = 1 Then
                                GoTo controlexcel
                            End If
                            subidoxls = 1
controlpdf:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroPdf()
                            End If
                            existePdf()
                            If pdf = 1 Then
                                GoTo controlpdf
                            End If
                            subidopdf = 1
                            If pi.TIPO = 1 Then
controltxt:
                                If nombre_pc = "ROBOT" Then
                                    subirFicheroCsv()
                                End If
                                existeCsv()
                                If csv = 1 Then
                                    GoTo controltxt
                                End If
                                movertxt()
                            End If
                            modificarRegistro()
                            Dim fechaactual As Date = Now()
                            Dim _fecha As String
                            _fecha = Format(fechaactual, "yyyy-MM-dd")
                            pi.marcarsubido(_fecha)
                            If tipoinforme = 15 Then
                                enviaremail()
                            End If
                            Dim s As New dSolicitudAnalisis
                            Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                            Dim fecenv As String
                            fecenv = Format(fechaenvio, "yyyy-MM-dd")
                            s.ID = idficha
                            s.actualizarfechaenvio2(fecenv)
                            s.marcar2()
                            s = Nothing
                            If subidoxls = 1 And subidopdf = 1 Then
                                If nombre_pc = "ROBOT" Then
                                    moverexcel()
                                    moverpdf()
                                Else
                                    moverexcel_otrapc()
                                    moverpdf_otrapc()
                                End If
                                ' Grabar estado de la ficha
                                Dim est As New dEstados
                                est.FICHA = idficha
                                est.ESTADO = 8
                                est.FECHA = fecenv
                                est.guardar2()
                                est = Nothing
                                '****************************
                                envio_mail_informes()
                            End If
                        Else
                        End If
                    Else
                        idficha = pi.FICHA
                        _tipoinforme = pi.TIPO
                        enviar_copia = pi.COPIA
                        _abonado = pi.ABONADO
                        _comentarios = pi.COMENTARIO
                        Dim sa As New dSolicitudAnalisis
                        sa.ID = idficha
                        sa = sa.buscar
                        If Not sa Is Nothing Then
                            Dim p As New dCliente
                            tipoinforme = sa.IDTIPOINFORME
                            p.ID = sa.IDPRODUCTOR
                            p = p.buscar
                            If Not p Is Nothing Then
                                email = ""
                                email2 = ""
                                If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                    email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                End If
                                If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                    email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                End If
                                If email = "" Then
                                    If email2 = "" Then
                                        If p.EMAIL <> "" Then
                                            email = RTrim(p.EMAIL)
                                        End If
                                    Else
                                        email = email2
                                    End If
                                Else
                                    If email2 = "" Then
                                        email = email
                                    Else
                                        email = email & "," & email2
                                    End If
                                End If
                                productorweb_com = p.USUARIO_WEB
                                Dim pw_com As New dProductorWeb_com
                                pw_com.USUARIO = productorweb_com
                                pw_com = pw_com.buscar
                                If Not pw_com Is Nothing Then
                                    idproductorweb_com = pw_com.ID
                                    'email = RTrim(pw_com.ENVIAR_EMAIL)
                                    celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                    carpeta = idproductorweb_com
                                    crea_carpeta()
                                End If
                                sa = Nothing
                            End If
                        End If
controlexcel2:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroXls()
                        End If
                        existeXls()
                        If excel = 1 Then
                            GoTo controlexcel2
                        End If
                        subidoxls = 1
controlpdf2:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroPdf()
                        End If
                        existePdf()
                        If pdf = 1 Then
                            GoTo controlpdf2
                        End If
                        subidopdf = 1
                        If pi.TIPO = 1 Then
controltxt2:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroCsv()
                            End If
                            existeCsv()
                            If csv = 1 Then
                                GoTo controltxt2
                            End If
                            movertxt()
                        End If
                        modificarRegistro()
                        Dim fechaactual As Date = Now()
                        Dim _fecha As String
                        _fecha = Format(fechaactual, "yyyy-MM-dd")
                        pi.marcarsubido(_fecha)
                        If tipoinforme = 15 Then
                            enviaremail()
                        End If
                        Dim s As New dSolicitudAnalisis
                        Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                        Dim fecenv As String
                        fecenv = Format(fechaenvio, "yyyy-MM-dd")
                        s.ID = idficha
                        s.actualizarfechaenvio2(fecenv)
                        s.marcar2()
                        s = Nothing
                        If subidoxls = 1 And subidopdf = 1 Then
                            If nombre_pc = "ROBOT" Then
                                moverexcel()
                                moverpdf()
                            Else
                                moverexcel_otrapc()
                                moverpdf_otrapc()
                            End If
                            ' Grabar estado de la ficha
                            Dim est As New dEstados
                            est.FICHA = idficha
                            est.ESTADO = 8
                            est.FECHA = fecenv
                            est.guardar2()
                            est = Nothing
                            '****************************
                            envio_mail_informes()
                        End If
                    End If
                    cimicro = Nothing
                Next
            End If
        End If
    End Sub
    Private Sub robot_subir_informes_subproductos()
        'No se usa /hardcode
        Dim pi As New dPreinformes
        Dim lista As New ArrayList
        Dim subidoxls As Integer = 0
        Dim subidopdf As Integer = 0
        Dim subidotxt As Integer = 0
        lista = pi.listarparasubirsubproductos
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each pi In lista
                    Dim cimicro As New dControlInformesMicro
                    Dim ficha As Long = 0
                    cimicro.FICHA = pi.FICHA
                    cimicro = cimicro.buscarxficha
                    If Not cimicro Is Nothing Then
                        If cimicro.CONTROLADO = 1 Then
                            idficha = pi.FICHA
                            _tipoinforme = pi.TIPO
                            enviar_copia = pi.COPIA
                            _abonado = pi.ABONADO
                            _comentarios = pi.COMENTARIO
                            Dim sa As New dSolicitudAnalisis
                            sa.ID = idficha
                            sa = sa.buscar
                            If Not sa Is Nothing Then
                                Dim p As New dCliente
                                tipoinforme = sa.IDTIPOINFORME
                                p.ID = sa.IDPRODUCTOR
                                p = p.buscar
                                If Not p Is Nothing Then
                                    email = ""
                                    email2 = ""
                                    If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                        email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                    End If
                                    If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                        email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                    End If
                                    If email = "" Then
                                        If email2 = "" Then
                                            If p.EMAIL <> "" Then
                                                email = RTrim(p.EMAIL)
                                            End If
                                        Else
                                            email = email2
                                        End If
                                    Else
                                        If email2 = "" Then
                                            email = email
                                        Else
                                            email = email & "," & email2
                                        End If
                                    End If
                                    productorweb_com = p.USUARIO_WEB
                                    Dim pw_com As New dProductorWeb_com
                                    pw_com.USUARIO = productorweb_com
                                    pw_com = pw_com.buscar
                                    If Not pw_com Is Nothing Then
                                        idproductorweb_com = pw_com.ID
                                        'email = RTrim(pw_com.ENVIAR_EMAIL)
                                        celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                        carpeta = idproductorweb_com
                                        crea_carpeta()
                                    End If
                                    sa = Nothing
                                End If
                            End If
controlexcel:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroXls()
                            End If
                            existeXls()
                            If excel = 1 Then
                                GoTo controlexcel
                            End If
                            subidoxls = 1
controlpdf:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroPdf()
                            End If
                            existePdf()
                            If pdf = 1 Then
                                GoTo controlpdf
                            End If
                            subidopdf = 1
                            If pi.TIPO = 1 Then
controltxt:
                                If nombre_pc = "ROBOT" Then
                                    subirFicheroCsv()
                                End If
                                existeCsv()
                                If csv = 1 Then
                                    GoTo controltxt
                                End If
                                movertxt()
                            End If
                            modificarRegistro()
                            Dim fechaactual As Date = Now()
                            Dim _fecha As String
                            _fecha = Format(fechaactual, "yyyy-MM-dd")
                            pi.marcarsubido(_fecha)
                            If tipoinforme = 15 Then
                                enviaremail()
                            End If
                            Dim s As New dSolicitudAnalisis
                            Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                            Dim fecenv As String
                            fecenv = Format(fechaenvio, "yyyy-MM-dd")
                            s.ID = idficha
                            s.actualizarfechaenvio2(fecenv)
                            s.marcar2()
                            s = Nothing
                            If subidoxls = 1 And subidopdf = 1 Then
                                If nombre_pc = "ROBOT" Then
                                    moverexcel()
                                    moverpdf()
                                Else
                                    moverexcel_otrapc()
                                    moverpdf_otrapc()
                                End If
                                ' Grabar estado de la ficha
                                Dim est As New dEstados
                                est.FICHA = idficha
                                est.ESTADO = 8
                                est.FECHA = fecenv
                                est.guardar2()
                                est = Nothing
                                '****************************
                                envio_mail_informes()
                            End If
                        Else
                        End If
                    Else
                        idficha = pi.FICHA
                        _tipoinforme = pi.TIPO
                        enviar_copia = pi.COPIA
                        _abonado = pi.ABONADO
                        _comentarios = pi.COMENTARIO
                        Dim sa As New dSolicitudAnalisis
                        sa.ID = idficha
                        sa = sa.buscar
                        If Not sa Is Nothing Then
                            Dim p As New dCliente
                            tipoinforme = sa.IDTIPOINFORME
                            p.ID = sa.IDPRODUCTOR
                            p = p.buscar
                            If Not p Is Nothing Then
                                email = ""
                                email2 = ""
                                If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                    email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                End If
                                If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                    email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                End If
                                If email = "" Then
                                    If email2 = "" Then
                                        If p.EMAIL <> "" Then
                                            email = RTrim(p.EMAIL)
                                        End If
                                    Else
                                        email = email2
                                    End If
                                Else
                                    If email2 = "" Then
                                        email = email
                                    Else
                                        email = email & "," & email2
                                    End If
                                End If
                                productorweb_com = p.USUARIO_WEB
                                Dim pw_com As New dProductorWeb_com
                                pw_com.USUARIO = productorweb_com
                                pw_com = pw_com.buscar
                                If Not pw_com Is Nothing Then
                                    idproductorweb_com = pw_com.ID
                                    'email = RTrim(pw_com.ENVIAR_EMAIL)
                                    celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                    carpeta = idproductorweb_com
                                    crea_carpeta()
                                End If
                                sa = Nothing
                            End If
                        End If
controlexcel2:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroXls()
                        End If
                        existeXls()
                        If excel = 1 Then
                            GoTo controlexcel2
                        End If
                        subidoxls = 1
controlpdf2:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroPdf()
                        End If
                        existePdf()
                        If pdf = 1 Then
                            GoTo controlpdf2
                        End If
                        subidopdf = 1
                        If pi.TIPO = 1 Then
controltxt2:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroCsv()
                            End If
                            existeCsv()
                            If csv = 1 Then
                                GoTo controltxt2
                            End If
                            movertxt()
                        End If
                        modificarRegistro()
                        Dim fechaactual As Date = Now()
                        Dim _fecha As String
                        _fecha = Format(fechaactual, "yyyy-MM-dd")
                        pi.marcarsubido(_fecha)
                        If tipoinforme = 15 Then
                            enviaremail()
                        End If
                        Dim s As New dSolicitudAnalisis
                        Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                        Dim fecenv As String
                        fecenv = Format(fechaenvio, "yyyy-MM-dd")
                        s.ID = idficha
                        s.actualizarfechaenvio2(fecenv)
                        s.marcar2()
                        s = Nothing
                        If subidoxls = 1 And subidopdf = 1 Then
                            If nombre_pc = "ROBOT" Then
                                moverexcel()
                                moverpdf()
                            Else
                                moverexcel_otrapc()
                                moverpdf_otrapc()
                            End If
                            ' Grabar estado de la ficha
                            Dim est As New dEstados
                            est.FICHA = idficha
                            est.ESTADO = 8
                            est.FECHA = fecenv
                            est.guardar2()
                            est = Nothing
                            '****************************
                            envio_mail_informes()
                        End If
                    End If
                    cimicro = Nothing
                Next
            End If
        End If
    End Sub
    Private Sub robot_subir_informes_calidad()
        'No se usa /hardcode
        Dim pi As New dPreinformes
        Dim lista As New ArrayList
        Dim subidoxls As Integer = 0
        Dim subidopdf As Integer = 0
        Dim subidotxt As Integer = 0
        lista = pi.listarparasubircalidad
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each pi In lista
                    Dim cifq As New dControlInformesFQ
                    Dim cimicro As New dControlInformesMicro
                    Dim ficha As Long = 0
                    cifq.FICHA = pi.FICHA
                    cifq = cifq.buscarxficha
                    cimicro.FICHA = pi.FICHA
                    cimicro = cimicro.buscarxficha
                    Dim control As Integer = 0
                    Dim control2 As Integer = 0
                    If Not cifq Is Nothing Then
                        If cifq.CONTROLADO = 1 Then
                            control = 1
                        Else
                            control = 0
                        End If
                    Else
                        control = 1
                    End If
                    If Not cimicro Is Nothing Then
                        If cimicro.CONTROLADO = 1 Then
                            control2 = 1
                        Else
                            control2 = 0
                        End If
                    Else
                        control2 = 1
                    End If
                    If control = 1 And control2 = 1 Then
                        idficha = pi.FICHA
                        _tipoinforme = pi.TIPO
                        enviar_copia = pi.COPIA
                        _abonado = pi.ABONADO
                        _comentarios = pi.COMENTARIO
                        Dim sa As New dSolicitudAnalisis
                        sa.ID = idficha
                        sa = sa.buscar
                        If Not sa Is Nothing Then
                            Dim p As New dCliente
                            tipoinforme = sa.IDTIPOINFORME
                            p.ID = sa.IDPRODUCTOR
                            p = p.buscar
                            If Not p Is Nothing Then
                                email = ""
                                email2 = ""
                                If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                    email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                End If
                                If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                    email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                End If
                                If email = "" Then
                                    If email2 = "" Then
                                        If p.EMAIL <> "" Then
                                            email = RTrim(p.EMAIL)
                                        End If
                                    Else
                                        email = email2
                                    End If
                                Else
                                    If email2 = "" Then
                                        email = email
                                    Else
                                        email = email & "," & email2
                                    End If
                                End If
                                productorweb_com = p.USUARIO_WEB
                                Dim pw_com As New dProductorWeb_com
                                pw_com.USUARIO = productorweb_com
                                pw_com = pw_com.buscar
                                If Not pw_com Is Nothing Then
                                    idproductorweb_com = pw_com.ID
                                    'email = RTrim(pw_com.ENVIAR_EMAIL)
                                    celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                    carpeta = idproductorweb_com
                                    crea_carpeta()
                                End If
                                sa = Nothing
                            End If
                        End If
controlexcel:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroXls()
                        End If
                        existeXls()
                        If excel = 1 Then
                            GoTo controlexcel
                        End If
                        subidoxls = 1
controlpdf:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroPdf()
                        End If
                        existePdf()
                        If pdf = 1 Then
                            GoTo controlpdf
                        End If
                        subidopdf = 1
                        If pi.TIPO = 1 Then
controltxt:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroCsv()
                            End If
                            existeCsv()
                            If csv = 1 Then
                                GoTo controltxt
                            End If
                            movertxt()
                        End If
                        modificarRegistro()
                        Dim fechaactual As Date = Now()
                        Dim _fecha As String
                        _fecha = Format(fechaactual, "yyyy-MM-dd")
                        pi.marcarsubido(_fecha)
                        If tipoinforme = 15 Then
                            enviaremail()
                        End If
                        Dim s As New dSolicitudAnalisis
                        Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                        Dim fecenv As String
                        fecenv = Format(fechaenvio, "yyyy-MM-dd")
                        s.ID = idficha
                        s.actualizarfechaenvio2(fecenv)
                        s.marcar2()
                        s = Nothing
                        If subidoxls = 1 And subidopdf = 1 Then
                            If nombre_pc = "ROBOT" Then
                                moverexcel()
                                moverpdf()
                            Else
                                moverexcel_otrapc()
                                moverpdf_otrapc()
                            End If
                            ' Grabar estado de la ficha
                            Dim est As New dEstados
                            est.FICHA = idficha
                            est.ESTADO = 8
                            est.FECHA = fecenv
                            est.guardar2()
                            est = Nothing
                            '****************************
                            envio_mail_informes()
                        End If
                    End If
                    cifq = Nothing
                Next
            End If
        End If
    End Sub
    Private Sub robot_subir_informes_nutricion()
        'No se usa /hardcode
        Dim pi As New dPreinformes
        Dim lista As New ArrayList
        Dim subidoxls As Integer = 0
        Dim subidopdf As Integer = 0
        lista = pi.listarparasubirnutricion
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each pi In lista
                    Dim cinutricion As New dControlInformesNutricion
                    Dim ficha As Long = 0
                    cinutricion.FICHA = pi.FICHA
                    cinutricion = cinutricion.buscarxficha
                    If Not cinutricion Is Nothing Then
                        If cinutricion.CONTROLADO = 1 Then
                            idficha = pi.FICHA
                            _tipoinforme = pi.TIPO
                            enviar_copia = pi.COPIA
                            _abonado = pi.ABONADO
                            _comentarios = pi.COMENTARIO
                            Dim sa As New dSolicitudAnalisis
                            sa.ID = idficha
                            sa = sa.buscar
                            If Not sa Is Nothing Then
                                Dim p As New dCliente
                                tipoinforme = sa.IDTIPOINFORME
                                p.ID = sa.IDPRODUCTOR
                                p = p.buscar
                                If Not p Is Nothing Then
                                    email = ""
                                    email2 = ""
                                    If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                        email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                    End If
                                    If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                        email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                    End If
                                    If email = "" Then
                                        If email2 = "" Then
                                            If p.EMAIL <> "" Then
                                                email = RTrim(p.EMAIL)
                                            End If
                                        Else
                                            email = email2
                                        End If
                                    Else
                                        If email2 = "" Then
                                            email = email
                                        Else
                                            email = email & "," & email2
                                        End If
                                    End If
                                    productorweb_com = p.USUARIO_WEB
                                    Dim pw_com As New dProductorWeb_com
                                    pw_com.USUARIO = productorweb_com
                                    pw_com = pw_com.buscar
                                    If Not pw_com Is Nothing Then
                                        idproductorweb_com = pw_com.ID
                                        'email = RTrim(pw_com.ENVIAR_EMAIL)
                                        celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                        carpeta = idproductorweb_com
                                        crea_carpeta()
                                    End If
                                    sa = Nothing
                                End If
                            End If
controlexcel:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroXls()
                            End If
                            existeXls()
                            If excel = 1 Then
                                GoTo controlexcel
                            End If
                            subidoxls = 1
controlpdf:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroPdf()
                            End If
                            existePdf()
                            If pdf = 1 Then
                                GoTo controlpdf
                            End If
                            subidopdf = 1
                            If pi.TIPO = 1 Then
controltxt:
                                If nombre_pc = "ROBOT" Then
                                    subirFicheroCsv()
                                End If
                                existeCsv()
                                If csv = 1 Then
                                    GoTo controltxt
                                End If
                                movertxt()
                            End If
                            modificarRegistro()
                            Dim fechaactual As Date = Now()
                            Dim _fecha As String
                            _fecha = Format(fechaactual, "yyyy-MM-dd")
                            pi.marcarsubido(_fecha)
                            If tipoinforme = 15 Then
                                enviaremail()
                            End If
                            Dim s As New dSolicitudAnalisis
                            Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                            Dim fecenv As String
                            fecenv = Format(fechaenvio, "yyyy-MM-dd")
                            s.ID = idficha
                            s.actualizarfechaenvio2(fecenv)
                            s.marcar2()
                            s = Nothing
                            If subidoxls = 1 And subidopdf = 1 Then
                                If nombre_pc = "ROBOT" Then
                                    moverexcel()
                                    moverpdf()
                                Else
                                    moverexcel_otrapc()
                                    moverpdf_otrapc()
                                End If
                                ' Grabar estado de la ficha
                                Dim est As New dEstados
                                est.FICHA = idficha
                                est.ESTADO = 8
                                est.FECHA = fecenv
                                est.guardar2()
                                est = Nothing
                                '****************************
                                envio_mail_informes()
                            End If
                        Else
                        End If
                    Else
                        idficha = pi.FICHA
                        _tipoinforme = pi.TIPO
                        enviar_copia = pi.COPIA
                        _abonado = pi.ABONADO
                        _comentarios = pi.COMENTARIO
                        Dim sa As New dSolicitudAnalisis
                        sa.ID = idficha
                        sa = sa.buscar
                        If Not sa Is Nothing Then
                            Dim p As New dCliente
                            tipoinforme = sa.IDTIPOINFORME
                            p.ID = sa.IDPRODUCTOR
                            p = p.buscar
                            If Not p Is Nothing Then
                                email = ""
                                email2 = ""
                                If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                    email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                End If
                                If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                    email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                End If
                                If email = "" Then
                                    If email2 = "" Then
                                        If p.EMAIL <> "" Then
                                            email = RTrim(p.EMAIL)
                                        End If
                                    Else
                                        email = email2
                                    End If
                                Else
                                    If email2 = "" Then
                                        email = email
                                    Else
                                        email = email & "," & email2
                                    End If
                                End If
                                productorweb_com = p.USUARIO_WEB
                                Dim pw_com As New dProductorWeb_com
                                pw_com.USUARIO = productorweb_com
                                pw_com = pw_com.buscar
                                If Not pw_com Is Nothing Then
                                    idproductorweb_com = pw_com.ID
                                    'email = RTrim(pw_com.ENVIAR_EMAIL)
                                    celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                    carpeta = idproductorweb_com
                                    crea_carpeta()
                                End If
                                sa = Nothing
                            End If
                        End If
controlexcel2:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroXls()
                        End If
                        existeXls()
                        If excel = 1 Then
                            GoTo controlexcel2
                        End If
                        subidoxls = 1
controlpdf2:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroPdf()
                        End If
                        existePdf()
                        If pdf = 1 Then
                            GoTo controlpdf2
                        End If
                        subidopdf = 1
                        If pi.TIPO = 1 Then
controltxt2:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroCsv()
                            End If
                            existeCsv()
                            If csv = 1 Then
                                GoTo controltxt2
                            End If
                            movertxt()
                        End If
                        modificarRegistro()
                        Dim fechaactual As Date = Now()
                        Dim _fecha As String
                        _fecha = Format(fechaactual, "yyyy-MM-dd")
                        pi.marcarsubido(_fecha)
                        If tipoinforme = 15 Then
                            enviaremail()
                        End If
                        Dim s As New dSolicitudAnalisis
                        Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                        Dim fecenv As String
                        fecenv = Format(fechaenvio, "yyyy-MM-dd")
                        s.ID = idficha
                        s.actualizarfechaenvio2(fecenv)
                        s.marcar2()
                        s = Nothing
                        If subidoxls = 1 And subidopdf = 1 Then
                            If nombre_pc = "ROBOT" Then
                                moverexcel()
                                moverpdf()
                            Else
                                moverexcel_otrapc()
                                moverpdf_otrapc()
                            End If
                            ' Grabar estado de la ficha
                            Dim est As New dEstados
                            est.FICHA = idficha
                            est.ESTADO = 8
                            est.FECHA = fecenv
                            est.guardar2()
                            est = Nothing
                            '****************************
                            envio_mail_informes()
                        End If
                    End If
                    cinutricion = Nothing
                Next
            End If
        End If
    End Sub
    Private Sub robot_subir_informes_suelos()
        'No se usa /hardcode
        Dim pi As New dPreinformes
        Dim lista As New ArrayList
        Dim subidoxls As Integer = 0
        Dim subidopdf As Integer = 0
        Dim subidotxt As Integer = 0
        lista = pi.listarparasubirsuelos
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each pi In lista
                    Dim cisuelos As New dControlInformesSuelos
                    Dim ficha As Long = 0
                    cisuelos.FICHA = pi.FICHA
                    cisuelos = cisuelos.buscarxficha
                    If Not cisuelos Is Nothing Then
                        If cisuelos.CONTROLADO = 1 Then
                            idficha = pi.FICHA
                            _tipoinforme = pi.TIPO
                            enviar_copia = pi.COPIA
                            _abonado = pi.ABONADO
                            _comentarios = pi.COMENTARIO
                            Dim sa As New dSolicitudAnalisis
                            sa.ID = idficha
                            sa = sa.buscar
                            If Not sa Is Nothing Then
                                Dim p As New dCliente
                                tipoinforme = sa.IDTIPOINFORME
                                p.ID = sa.IDPRODUCTOR
                                p = p.buscar
                                If Not p Is Nothing Then
                                    email = ""
                                    email2 = ""
                                    If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                        email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                    End If
                                    If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                        email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                    End If
                                    If email = "" Then
                                        If email2 = "" Then
                                            If p.EMAIL <> "" Then
                                                email = RTrim(p.EMAIL)
                                            End If
                                        Else
                                            email = email2
                                        End If
                                    Else
                                        If email2 = "" Then
                                            email = email
                                        Else
                                            email = email & "," & email2
                                        End If
                                    End If
                                    productorweb_com = p.USUARIO_WEB
                                    Dim pw_com As New dProductorWeb_com
                                    pw_com.USUARIO = productorweb_com
                                    pw_com = pw_com.buscar
                                    If Not pw_com Is Nothing Then
                                        idproductorweb_com = pw_com.ID
                                        'email = RTrim(pw_com.ENVIAR_EMAIL)
                                        celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                        carpeta = idproductorweb_com
                                        crea_carpeta()
                                    End If
                                    sa = Nothing
                                End If
                            End If
controlexcel:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroXls()
                            End If
                            existeXls()
                            If excel = 1 Then
                                GoTo controlexcel
                            End If
                            subidoxls = 1
controlpdf:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroPdf()
                            End If
                            existePdf()
                            If pdf = 1 Then
                                GoTo controlpdf
                            End If
                            subidopdf = 1
                            If pi.TIPO = 1 Then
controltxt:
                                If nombre_pc = "ROBOT" Then
                                    subirFicheroCsv()
                                End If
                                existeCsv()
                                If csv = 1 Then
                                    GoTo controltxt
                                End If
                                movertxt()
                            End If
                            modificarRegistro()
                            Dim fechaactual As Date = Now()
                            Dim _fecha As String
                            _fecha = Format(fechaactual, "yyyy-MM-dd")
                            pi.marcarsubido(_fecha)
                            If tipoinforme = 15 Then
                                enviaremail()
                            End If
                            Dim s As New dSolicitudAnalisis
                            Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                            Dim fecenv As String
                            fecenv = Format(fechaenvio, "yyyy-MM-dd")
                            s.ID = idficha
                            s.actualizarfechaenvio2(fecenv)
                            s.marcar2()
                            s = Nothing
                            If subidoxls = 1 And subidopdf = 1 Then
                                If nombre_pc = "ROBOT" Then
                                    moverexcel()
                                    moverpdf()
                                Else
                                    moverexcel_otrapc()
                                    moverpdf_otrapc()
                                End If
                                ' Grabar estado de la ficha
                                Dim est As New dEstados
                                est.FICHA = idficha
                                est.ESTADO = 8
                                est.FECHA = fecenv
                                est.guardar2()
                                est = Nothing
                                '****************************
                                envio_mail_informes()
                            End If
                        Else
                        End If
                    Else
                        idficha = pi.FICHA
                        _tipoinforme = pi.TIPO
                        enviar_copia = pi.COPIA
                        _abonado = pi.ABONADO
                        _comentarios = pi.COMENTARIO
                        Dim sa As New dSolicitudAnalisis
                        sa.ID = idficha
                        sa = sa.buscar
                        If Not sa Is Nothing Then
                            Dim p As New dCliente
                            tipoinforme = sa.IDTIPOINFORME
                            p.ID = sa.IDPRODUCTOR
                            p = p.buscar
                            If Not p Is Nothing Then
                                email = ""
                                email2 = ""
                                If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                    email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                End If
                                If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                    email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                End If
                                If email = "" Then
                                    If email2 = "" Then
                                        If p.EMAIL <> "" Then
                                            email = RTrim(p.EMAIL)
                                        End If
                                    Else
                                        email = email2
                                    End If
                                Else
                                    If email2 = "" Then
                                        email = email
                                    Else
                                        email = email & "," & email2
                                    End If
                                End If
                                productorweb_com = p.USUARIO_WEB
                                Dim pw_com As New dProductorWeb_com
                                pw_com.USUARIO = productorweb_com
                                pw_com = pw_com.buscar
                                If Not pw_com Is Nothing Then
                                    idproductorweb_com = pw_com.ID
                                    'email = RTrim(pw_com.ENVIAR_EMAIL)
                                    celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                    carpeta = idproductorweb_com
                                    crea_carpeta()
                                End If
                                sa = Nothing
                            End If
                        End If
controlexcel2:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroXls()
                        End If
                        existeXls()
                        If excel = 1 Then
                            GoTo controlexcel2
                        End If
                        subidoxls = 1
controlpdf2:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroPdf()
                        End If
                        existePdf()
                        If pdf = 1 Then
                            GoTo controlpdf2
                        End If
                        subidopdf = 1
                        If pi.TIPO = 1 Then
controltxt2:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroCsv()
                            End If
                            existeCsv()
                            If csv = 1 Then
                                GoTo controltxt2
                            End If
                            movertxt()
                        End If
                        modificarRegistro()
                        Dim fechaactual As Date = Now()
                        Dim _fecha As String
                        _fecha = Format(fechaactual, "yyyy-MM-dd")
                        pi.marcarsubido(_fecha)
                        If tipoinforme = 15 Then
                            enviaremail()
                        End If
                        Dim s As New dSolicitudAnalisis
                        Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                        Dim fecenv As String
                        fecenv = Format(fechaenvio, "yyyy-MM-dd")
                        s.ID = idficha
                        s.actualizarfechaenvio2(fecenv)
                        s.marcar2()
                        s = Nothing
                        If subidoxls = 1 And subidopdf = 1 Then
                            If nombre_pc = "ROBOT" Then
                                moverexcel()
                                moverpdf()
                            Else
                                moverexcel_otrapc()
                                moverpdf_otrapc()
                            End If
                            ' Grabar estado de la ficha
                            Dim est As New dEstados
                            est.FICHA = idficha
                            est.ESTADO = 8
                            est.FECHA = fecenv
                            est.guardar2()
                            est = Nothing
                            '****************************
                            envio_mail_informes()
                        End If
                    End If
                    cisuelos = Nothing
                Next
            End If
        End If
    End Sub
    Private Sub robot_subir_informes_foliar()
        'No se usa /hardcode
        Dim pi As New dPreinformes
        Dim lista As New ArrayList
        Dim subidoxls As Integer = 0
        Dim subidopdf As Integer = 0
        Dim subidotxt As Integer = 0
        lista = pi.listarparasubirfoliar
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each pi In lista
                    Dim cisuelos As New dControlInformesSuelos
                    Dim ficha As Long = 0
                    cisuelos.FICHA = pi.FICHA
                    cisuelos = cisuelos.buscarxficha
                    If Not cisuelos Is Nothing Then
                        If cisuelos.CONTROLADO = 1 Then
                            idficha = pi.FICHA
                            _tipoinforme = pi.TIPO
                            enviar_copia = pi.COPIA
                            _abonado = pi.ABONADO
                            _comentarios = pi.COMENTARIO
                            Dim sa As New dSolicitudAnalisis
                            sa.ID = idficha
                            sa = sa.buscar
                            If Not sa Is Nothing Then
                                Dim p As New dCliente
                                tipoinforme = sa.IDTIPOINFORME
                                p.ID = sa.IDPRODUCTOR
                                p = p.buscar
                                If Not p Is Nothing Then
                                    email = ""
                                    email2 = ""
                                    If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                        email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                    End If
                                    If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                        email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                    End If
                                    If email = "" Then
                                        If email2 = "" Then
                                            If p.EMAIL <> "" Then
                                                email = RTrim(p.EMAIL)
                                            End If
                                        Else
                                            email = email2
                                        End If
                                    Else
                                        If email2 = "" Then
                                            email = email
                                        Else
                                            email = email & "," & email2
                                        End If
                                    End If
                                    productorweb_com = p.USUARIO_WEB
                                    Dim pw_com As New dProductorWeb_com
                                    pw_com.USUARIO = productorweb_com
                                    pw_com = pw_com.buscar
                                    If Not pw_com Is Nothing Then
                                        idproductorweb_com = pw_com.ID
                                        'email = RTrim(pw_com.ENVIAR_EMAIL)
                                        celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                        carpeta = idproductorweb_com
                                        crea_carpeta()
                                    End If
                                    sa = Nothing
                                End If
                            End If
controlexcel:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroXls()
                            End If
                            existeXls()
                            If excel = 1 Then
                                GoTo controlexcel
                            End If
                            subidoxls = 1
controlpdf:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroPdf()
                            End If
                            existePdf()
                            If pdf = 1 Then
                                GoTo controlpdf
                            End If
                            subidopdf = 1
                            If pi.TIPO = 1 Then
controltxt:
                                If nombre_pc = "ROBOT" Then
                                    subirFicheroCsv()
                                End If
                                existeCsv()
                                If csv = 1 Then
                                    GoTo controltxt
                                End If
                                movertxt()
                            End If
                            modificarRegistro()
                            Dim fechaactual As Date = Now()
                            Dim _fecha As String
                            _fecha = Format(fechaactual, "yyyy-MM-dd")
                            pi.marcarsubido(_fecha)
                            If tipoinforme = 15 Then
                                enviaremail()
                            End If
                            Dim s As New dSolicitudAnalisis
                            Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                            Dim fecenv As String
                            fecenv = Format(fechaenvio, "yyyy-MM-dd")
                            s.ID = idficha
                            s.actualizarfechaenvio2(fecenv)
                            s.marcar2()
                            s = Nothing
                            If subidoxls = 1 And subidopdf = 1 Then
                                If nombre_pc = "ROBOT" Then
                                    moverexcel()
                                    moverpdf()
                                Else
                                    moverexcel_otrapc()
                                    moverpdf_otrapc()
                                End If
                                ' Grabar estado de la ficha
                                Dim est As New dEstados
                                est.FICHA = idficha
                                est.ESTADO = 8
                                est.FECHA = fecenv
                                est.guardar2()
                                est = Nothing
                                '****************************
                                envio_mail_informes()
                            End If
                        Else
                        End If
                    Else
                        idficha = pi.FICHA
                        _tipoinforme = pi.TIPO
                        enviar_copia = pi.COPIA
                        _abonado = pi.ABONADO
                        _comentarios = pi.COMENTARIO
                        Dim sa As New dSolicitudAnalisis
                        sa.ID = idficha
                        sa = sa.buscar
                        If Not sa Is Nothing Then
                            Dim p As New dCliente
                            tipoinforme = sa.IDTIPOINFORME
                            p.ID = sa.IDPRODUCTOR
                            p = p.buscar
                            If Not p Is Nothing Then
                                email = ""
                                email2 = ""
                                If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                    email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                End If
                                If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                    email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                End If
                                If email = "" Then
                                    If email2 = "" Then
                                        If p.EMAIL <> "" Then
                                            email = RTrim(p.EMAIL)
                                        End If
                                    Else
                                        email = email2
                                    End If
                                Else
                                    If email2 = "" Then
                                        email = email
                                    Else
                                        email = email & "," & email2
                                    End If
                                End If
                                productorweb_com = p.USUARIO_WEB
                                Dim pw_com As New dProductorWeb_com
                                pw_com.USUARIO = productorweb_com
                                pw_com = pw_com.buscar
                                If Not pw_com Is Nothing Then
                                    idproductorweb_com = pw_com.ID
                                    'email = RTrim(pw_com.ENVIAR_EMAIL)
                                    celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                    carpeta = idproductorweb_com
                                    crea_carpeta()
                                End If
                                sa = Nothing
                            End If
                        End If
controlexcel2:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroXls()
                        End If
                        existeXls()
                        If excel = 1 Then
                            GoTo controlexcel2
                        End If
                        subidoxls = 1
controlpdf2:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroPdf()
                        End If
                        existePdf()
                        If pdf = 1 Then
                            GoTo controlpdf2
                        End If
                        subidopdf = 1
                        If pi.TIPO = 1 Then
controltxt2:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroCsv()
                            End If
                            existeCsv()
                            If csv = 1 Then
                                GoTo controltxt2
                            End If
                            movertxt()
                        End If
                        modificarRegistro()
                        Dim fechaactual As Date = Now()
                        Dim _fecha As String
                        _fecha = Format(fechaactual, "yyyy-MM-dd")
                        pi.marcarsubido(_fecha)
                        If tipoinforme = 15 Then
                            enviaremail()
                        End If
                        Dim s As New dSolicitudAnalisis
                        Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                        Dim fecenv As String
                        fecenv = Format(fechaenvio, "yyyy-MM-dd")
                        s.ID = idficha
                        s.actualizarfechaenvio2(fecenv)
                        s.marcar2()
                        s = Nothing
                        If subidoxls = 1 And subidopdf = 1 Then
                            If nombre_pc = "ROBOT" Then
                                moverexcel()
                                moverpdf()
                            Else
                                moverexcel_otrapc()
                                moverpdf_otrapc()
                            End If
                            ' Grabar estado de la ficha
                            Dim est As New dEstados
                            est.FICHA = idficha
                            est.ESTADO = 8
                            est.FECHA = fecenv
                            est.guardar2()
                            est = Nothing
                            '****************************
                            envio_mail_informes()
                        End If
                    End If
                    cisuelos = Nothing
                Next
            End If
        End If
    End Sub
    Private Sub robot_subir_informes_brucelosis()
        'No se usa /hardcode
        Dim pi As New dPreinformes
        Dim lista As New ArrayList
        Dim subidoxls As Integer = 0
        Dim subidopdf As Integer = 0
        Dim subidotxt As Integer = 0
        lista = pi.listarparasubirbrucelosis
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each pi In lista
                    idficha = pi.FICHA
                    _tipoinforme = pi.TIPO
                    enviar_copia = pi.COPIA
                    _abonado = pi.ABONADO
                    _comentarios = pi.COMENTARIO
                    Dim sa As New dSolicitudAnalisis
                    sa.ID = idficha
                    sa = sa.buscar
                    If Not sa Is Nothing Then
                        Dim p As New dCliente
                        tipoinforme = sa.IDTIPOINFORME
                        p.ID = sa.IDPRODUCTOR
                        p = p.buscar
                        If Not p Is Nothing Then
                            email = ""
                            email2 = ""
                            If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                email = RTrim(p.NOT_EMAIL_ANALISIS1)
                            End If
                            If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                            End If
                            If email = "" Then
                                If email2 = "" Then
                                    If p.EMAIL <> "" Then
                                        email = RTrim(p.EMAIL)
                                    End If
                                Else
                                    email = email2
                                End If
                            Else
                                If email2 = "" Then
                                    email = email
                                Else
                                    email = email & "," & email2
                                End If
                            End If
                            productorweb_com = p.USUARIO_WEB
                            Dim pw_com As New dProductorWeb_com
                            pw_com.USUARIO = productorweb_com
                            pw_com = pw_com.buscar
                            If Not pw_com Is Nothing Then
                                idproductorweb_com = pw_com.ID
                                'email = RTrim(pw_com.ENVIAR_EMAIL)
                                celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                carpeta = idproductorweb_com
                                crea_carpeta()
                            End If
                            sa = Nothing
                        End If
                    End If
controlexcel:
                    If nombre_pc = "ROBOT" Then
                        subirFicheroXls()
                    End If
                    existeXls()
                    If excel = 1 Then
                        GoTo controlexcel
                    End If
                    subidoxls = 1
controlpdf:
                    If nombre_pc = "ROBOT" Then
                        subirFicheroPdf()
                    End If
                    existePdf()
                    If pdf = 1 Then
                        GoTo controlpdf
                    End If
                    subidopdf = 1
                    If pi.TIPO = 1 Then
controltxt:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroCsv()
                        End If
                        existeCsv()
                        If csv = 1 Then
                            GoTo controltxt
                        End If
                        movertxt()
                    End If
                    modificarRegistro()
                    Dim fechaactual As Date = Now()
                    Dim _fecha As String
                    _fecha = Format(fechaactual, "yyyy-MM-dd")
                    pi.marcarsubido(_fecha)
                    If tipoinforme = 15 Then
                        enviaremail()
                    End If
                    Dim s As New dSolicitudAnalisis
                    Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                    Dim fecenv As String
                    fecenv = Format(fechaenvio, "yyyy-MM-dd")
                    s.ID = idficha
                    s.actualizarfechaenvio2(fecenv)
                    s.marcar2()
                    s = Nothing
                    If subidoxls = 1 And subidopdf = 1 Then
                        If nombre_pc = "ROBOT" Then
                            moverexcel()
                            moverpdf()
                        Else
                            moverexcel_otrapc()
                            moverpdf_otrapc()
                        End If
                        ' Grabar estado de la ficha
                        Dim est As New dEstados
                        est.FICHA = idficha
                        est.ESTADO = 8
                        est.FECHA = fecenv
                        est.guardar2()
                        est = Nothing
                        '****************************
                        envio_mail_informes()
                    End If
                Next
            End If
        End If
    End Sub
    Private Sub robot_subir_informes_ambiental()
        'No se usa /hardcode
        Dim pi As New dPreinformes
        Dim lista As New ArrayList
        Dim subidoxls As Integer = 0
        Dim subidopdf As Integer = 0
        Dim subidotxt As Integer = 0
        lista = pi.listarparasubirambiental
        If Not lista Is Nothing Then
            If lista.Count > 0 Then
                For Each pi In lista
                    Dim cimicro As New dControlInformesMicro
                    Dim ficha As Long = 0
                    cimicro.FICHA = pi.FICHA
                    cimicro = cimicro.buscarxficha
                    If Not cimicro Is Nothing Then
                        If cimicro.CONTROLADO = 1 Then
                            idficha = pi.FICHA
                            _tipoinforme = pi.TIPO
                            enviar_copia = pi.COPIA
                            _abonado = pi.ABONADO
                            _comentarios = pi.COMENTARIO
                            Dim sa As New dSolicitudAnalisis
                            sa.ID = idficha
                            sa = sa.buscar
                            If Not sa Is Nothing Then
                                Dim p As New dCliente
                                tipoinforme = sa.IDTIPOINFORME
                                p.ID = sa.IDPRODUCTOR
                                p = p.buscar
                                If Not p Is Nothing Then
                                    email = ""
                                    email2 = ""
                                    If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                        email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                    End If
                                    If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                        email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                    End If
                                    If email = "" Then
                                        If email2 = "" Then
                                            If p.EMAIL <> "" Then
                                                email = RTrim(p.EMAIL)
                                            End If
                                        Else
                                            email = email2
                                        End If
                                    Else
                                        If email2 = "" Then
                                            email = email
                                        Else
                                            email = email & "," & email2
                                        End If
                                    End If
                                    productorweb_com = p.USUARIO_WEB
                                    Dim pw_com As New dProductorWeb_com
                                    pw_com.USUARIO = productorweb_com
                                    pw_com = pw_com.buscar
                                    If Not pw_com Is Nothing Then
                                        idproductorweb_com = pw_com.ID
                                        'email = RTrim(pw_com.ENVIAR_EMAIL)
                                        celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                        carpeta = idproductorweb_com
                                        crea_carpeta()
                                    End If
                                    sa = Nothing
                                End If
                            End If
controlexcel:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroXls()
                            End If
                            existeXls()
                            If excel = 1 Then
                                GoTo controlexcel
                            End If
                            subidoxls = 1
controlpdf:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroPdf()
                            End If
                            existePdf()
                            If pdf = 1 Then
                                GoTo controlpdf
                            End If
                            subidopdf = 1
                            If pi.TIPO = 1 Then
controltxt:
                                If nombre_pc = "ROBOT" Then
                                    subirFicheroCsv()
                                End If
                                existeCsv()
                                If csv = 1 Then
                                    GoTo controltxt
                                End If
                                movertxt()
                            End If
                            modificarRegistro()
                            Dim fechaactual As Date = Now()
                            Dim _fecha As String
                            _fecha = Format(fechaactual, "yyyy-MM-dd")
                            pi.marcarsubido(_fecha)
                            If tipoinforme = 15 Then
                                enviaremail()
                            End If
                            Dim s As New dSolicitudAnalisis
                            Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                            Dim fecenv As String
                            fecenv = Format(fechaenvio, "yyyy-MM-dd")
                            s.ID = idficha
                            s.actualizarfechaenvio2(fecenv)
                            s.marcar2()
                            s = Nothing
                            If subidoxls = 1 And subidopdf = 1 Then
                                If nombre_pc = "ROBOT" Then
                                    moverexcel()
                                    moverpdf()
                                Else
                                    moverexcel_otrapc()
                                    moverpdf_otrapc()
                                End If
                                ' Grabar estado de la ficha
                                Dim est As New dEstados
                                est.FICHA = idficha
                                est.ESTADO = 8
                                est.FECHA = fecenv
                                est.guardar2()
                                est = Nothing
                                '****************************
                                envio_mail_informes()
                            End If
                        Else
                        End If
                    Else
                        idficha = pi.FICHA
                        _tipoinforme = pi.TIPO
                        enviar_copia = pi.COPIA
                        _abonado = pi.ABONADO
                        _comentarios = pi.COMENTARIO
                        Dim sa As New dSolicitudAnalisis
                        sa.ID = idficha
                        sa = sa.buscar
                        If Not sa Is Nothing Then
                            Dim p As New dCliente
                            tipoinforme = sa.IDTIPOINFORME
                            p.ID = sa.IDPRODUCTOR
                            p = p.buscar
                            If Not p Is Nothing Then
                                email = ""
                                email2 = ""
                                If p.NOT_EMAIL_ANALISIS1 <> "" Then
                                    email = RTrim(p.NOT_EMAIL_ANALISIS1)
                                End If
                                If p.NOT_EMAIL_ANALISIS2 <> "" Then
                                    email2 = RTrim(p.NOT_EMAIL_ANALISIS2)
                                End If
                                If email = "" Then
                                    If email2 = "" Then
                                        If p.EMAIL <> "" Then
                                            email = RTrim(p.EMAIL)
                                        End If
                                    Else
                                        email = email2
                                    End If
                                Else
                                    If email2 = "" Then
                                        email = email
                                    Else
                                        email = email & "," & email2
                                    End If
                                End If
                                productorweb_com = p.USUARIO_WEB
                                Dim pw_com As New dProductorWeb_com
                                pw_com.USUARIO = productorweb_com
                                pw_com = pw_com.buscar
                                If Not pw_com Is Nothing Then
                                    idproductorweb_com = pw_com.ID
                                    'email = RTrim(pw_com.ENVIAR_EMAIL)
                                    celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                    carpeta = idproductorweb_com
                                    crea_carpeta()
                                End If
                                sa = Nothing
                            End If
                        End If
controlexcel2:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroXls()
                        End If
                        existeXls()
                        If excel = 1 Then
                            GoTo controlexcel2
                        End If
                        subidoxls = 1
controlpdf2:
                        If nombre_pc = "ROBOT" Then
                            subirFicheroPdf()
                        End If
                        existePdf()
                        If pdf = 1 Then
                            GoTo controlpdf2
                        End If
                        subidopdf = 1
                        If pi.TIPO = 1 Then
controltxt2:
                            If nombre_pc = "ROBOT" Then
                                subirFicheroCsv()
                            End If
                            existeCsv()
                            If csv = 1 Then
                                GoTo controltxt2
                            End If
                            movertxt()
                        End If
                        modificarRegistro()
                        Dim fechaactual As Date = Now()
                        Dim _fecha As String
                        _fecha = Format(fechaactual, "yyyy-MM-dd")
                        pi.marcarsubido(_fecha)
                        If tipoinforme = 15 Then
                            enviaremail()
                        End If
                        Dim s As New dSolicitudAnalisis
                        Dim fechaenvio As Date = DateFecha.Value.ToString("yyyy-MM-dd")
                        Dim fecenv As String
                        fecenv = Format(fechaenvio, "yyyy-MM-dd")
                        s.ID = idficha
                        s.actualizarfechaenvio2(fecenv)
                        s.marcar2()
                        s = Nothing
                        If subidoxls = 1 And subidopdf = 1 Then
                            If nombre_pc = "ROBOT" Then
                                moverexcel()
                                moverpdf()
                            Else
                                moverexcel_otrapc()
                                moverpdf_otrapc()
                            End If
                            ' Grabar estado de la ficha
                            Dim est As New dEstados
                            est.FICHA = idficha
                            est.ESTADO = 8
                            est.FECHA = fecenv
                            est.guardar2()
                            est = Nothing
                            '****************************
                            envio_mail_informes()
                        End If
                    End If
                    cimicro = Nothing
                Next
            End If
        End If
    End Sub
    Private Sub robot_enviarcompras()
        'No se usa /hardcode
        Dim c As New dCompras
        Dim lista As New ArrayList
        lista = c.listarsinenviar
        If Not lista Is Nothing Then
            For Each c In lista
                compraid = c.ID
                enviaremailcompras()
            Next
        End If
    End Sub
    Private Sub robot_calcularsaldos()
        'No se usa /hardcode
        Dim s As New dSaldos
        s.eliminartodos()
        Dim m As New dMovCte
        Dim c As New dCliente
        Dim idcliente As Long = 0
        Dim debe As Double = 0
        Dim haber As Double = 0
        Dim saldo As Double = 0
        Dim listaclientes As New ArrayList
        Dim listamovimientos As New ArrayList
        listaclientes = c.listar
        If Not listaclientes Is Nothing Then
            For Each c In listaclientes
                idcliente = c.ID
                listamovimientos = m.saldosxcli(idcliente)
                If Not listamovimientos Is Nothing Then
                    For Each m In listamovimientos
                        If m.MCCCOD = 1 Then
                            debe = debe + m.MCCIMP
                        ElseIf m.MCCCOD = 2 Then
                            haber = haber + m.MCCIMP
                        End If
                    Next
                    saldo = debe - haber
                End If
                s.IDCLIENTE = idcliente
                s.SALDO = saldo
                saldo = 0
                debe = 0
                haber = 0
                s.guardar()
            Next
        End If
        subirsaldos()
    End Sub
    Private Sub robot_moverarchivossubidos()
        'No se usa /hardcode
        Dim extension As String
        Dim nombrearchivo As String = ""
        Dim numficha As Long = 0
        'Dim folder As New DirectoryInfo("\\192.168.1.10\e\NET\Informes para subir")
        Dim folder As New DirectoryInfo("C:\INFORMES PARA SUBIR")
        Dim _ficheros() As String
        '_ficheros = Directory.GetFiles("\\192.168.1.10\e\NET\Informes para subir")
        _ficheros = Directory.GetFiles("C:\INFORMES PARA SUBIR")
        If Not (_ficheros.Length > 0) Then
        Else
            For Each file As FileInfo In folder.GetFiles("*.*")
                nombrearchivo = file.Name
                extension = Microsoft.VisualBasic.Right(file.Name, 3)
                If extension = "xls" Or extension = "pdf" Or extension = "txt" Or extension = "XLS" Or extension = "PDF" Or extension = "TXT" Then
                    numficha = Mid(file.Name, 1, Len(file.Name) - 4)
                    idficha = numficha
                    Dim sa As New dSolicitudAnalisis
                    sa.ID = numficha
                    sa = sa.buscar
                    If Not sa Is Nothing Then
                        tipoinforme = sa.IDTIPOINFORME
                        If sa.MARCA = 1 Then
                            If extension = "xls" Or extension = "XLS" Then
                                moverexcel()
                            End If
                            If extension = "pdf" Or extension = "PDF" Then
                                moverpdf()
                            End If
                            If extension = "txt" Or extension = "TXT" Then
                                movertxt()
                            End If
                        End If
                    End If
                    sa = Nothing
                End If
            Next
        End If
    End Sub
    Private Sub robot_cargarpedidosweb()
        'No se usa /hardcode
        Dim pw_com As New dPedidosWeb_com
        Dim fecha As Date = Now
        Dim fec As String
        fec = Format(fecha, "yyyy-MM-dd")
        Dim lista As New ArrayList
        lista = pw_com.listar
        If Not lista Is Nothing Then
            For Each pw_com In lista
                Dim pw As New dPedidosWeb
                pw.FECHA = fec
                pw.CODIGO = pw_com.CODIGO
                pw.NOMBRE = pw_com.NOMBRE
                pw.DIRECCION = pw_com.DIRECCION
                pw.AGENCIA = pw_com.AGENCIA
                pw.TELEFONO = pw_com.TELEFONO
                pw.EMAIL = pw_com.EMAIL
                pw.CCONSERVANTE = pw_com.CCONSERVANTE
                pw.SCONSERVANTE = pw_com.SCONSERVANTE
                pw.AGUA = pw_com.AGUA
                pw.SANGRE = pw_com.SANGRE
                pw.OBSERVACIONES = pw_com.OBSERVACIONES
                pw.REALIZADO = 0
                pw.CANCELADO = 0
                pw.guardar()
                pw_com.eliminar()
                pw = Nothing
            Next
        End If
    End Sub
    Private Sub robot_cargarsolicitudesweb()
        'No se usa /hardcode
        Dim El_Ping As Boolean
        Dim eco As New System.Net.NetworkInformation.Ping
        Dim res As System.Net.NetworkInformation.PingReply
        Dim ip As Net.IPAddress
        ip = Net.IPAddress.Parse("209.62.67.162")
        res = eco.Send(ip)
        If res.Status = System.Net.NetworkInformation.IPStatus.Success Then
            El_Ping = 1
        End If
        If El_Ping = False Then
        Else
            Dim sw As New dSolicitudWeb
            Dim lista As New ArrayList
            lista = sw.listarsincargar
            If Not lista Is Nothing Then
                If lista.Count > 0 Then
                    For Each sw In lista
                        barra = barra + 1
                        ProgressBar1.Value = barra
                        If barra = 100 Then
                            barra = 0
                        End If
                        Dim s As New dSolicitudAnalisis
                        s.ID = sw.FICHA
                        id_ficha_ = s.ID
                        s = s.buscar
                        nficha = sw.FICHA
                        nmuestrasweb = s.NMUESTRAS
                        If Not s Is Nothing Then
                            Dim ti As New dTipoInforme
                            ti.ID = s.IDTIPOINFORME
                            ti = ti.buscar
                            If Not ti Is Nothing Then
                                tipoinformeweb = ti.NOMBRE
                            End If
                            Dim si As New dSubInforme
                            si.ID = s.IDSUBINFORME
                            si = si.buscar
                            If Not si Is Nothing Then
                                subtipoinformeweb = si.NOMBRE
                            End If
                            observacionesweb = s.OBSERVACIONES
                            Dim m As New dMuestras
                            m.ID = s.IDMUESTRA
                            m = m.buscar
                            If Not m Is Nothing Then
                                muestraweb = m.NOMBRE
                            End If
                            Dim p As New dCliente
                            p.ID = s.IDPRODUCTOR
                            p = p.buscar
                            If Not p Is Nothing Then
                                email = ""
                                email2 = ""
                                If p.NOT_EMAIL_MUESTRAS1 <> "" Then
                                    email = RTrim(p.NOT_EMAIL_MUESTRAS1)
                                End If
                                If p.NOT_EMAIL_MUESTRAS2 <> "" Then
                                    email2 = RTrim(p.NOT_EMAIL_MUESTRAS2)
                                End If
                                If email = "" Then
                                    If email2 = "" Then
                                        If p.EMAIL <> "" Then
                                            email = RTrim(p.EMAIL)
                                        End If
                                    Else
                                        email = email2
                                    End If
                                Else
                                    If email2 = "" Then
                                        email = email
                                    Else
                                        email = email & "," & email2
                                    End If
                                End If
                                nombreproductorweb = p.NOMBRE
                                productorweb_com = p.USUARIO_WEB
                                Dim pw_com As New dProductorWeb_com
                                pw_com.USUARIO = productorweb_com
                                pw_com = pw_com.buscar
                                If Not pw_com Is Nothing Then
                                    idproductorweb_com = pw_com.ID
                                    'email = RTrim(pw_com.ENVIAR_EMAIL)
                                    celular = Replace(pw_com.ENVIAR_SMS, " ", "")
                                    InsertarRegistro_com()
                                    sw.marcarcargado()
                                    nficha = 0
                                    email = ""
                                    email2 = ""
                                End If
                            End If
                        End If
                    Next
                End If
            End If
        End If
    End Sub

#End Region

#Region "BUTTONS"

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        subir_informes2()
    End Sub

    Private Sub ButtonImportar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        importar()
    End Sub
    Private Sub MoverArchivosSubidosToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MoverArchivosSubidosToolStripMenuItem.Click
        moverarchivossubidos()
    End Sub
    Private Sub ImportarToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImportarToolStripMenuItem.Click
        importar()
    End Sub
    Private Sub SubirInformesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SubirInformesToolStripMenuItem.Click
        ListBox1.Items.Clear()
        ListBox1.Items.Add(Now)
        ListBox1.Items.Add("Subiendo informes de control lechero")
        subir_informes_control()
        ListBox1.Items.Clear()
        ListBox1.Items.Add(Now)
        ListBox1.Items.Add("Subiendo informes de agua")
        subir_informes_agua()
        ListBox1.Items.Clear()
        ListBox1.Items.Add(Now)
        ListBox1.Items.Add("Subiendo informes de ATB")
        subir_informes_atb()
        ListBox1.Items.Clear()
        ListBox1.Items.Add(Now)
        ListBox1.Items.Add("Subiendo informes de Alimentos")
        subir_informes_subproductos()
        ListBox1.Items.Clear()
        ListBox1.Items.Add(Now)
        ListBox1.Items.Add("Subiendo informes de calidad de leche")
        subir_informes_calidad()
        ListBox1.Items.Clear()
        ListBox1.Items.Add(Now)
        ListBox1.Items.Add("Subiendo informes de nutrición")
        subir_informes_nutricion()
        ListBox1.Items.Clear()
        ListBox1.Items.Add(Now)
        ListBox1.Items.Add("Subiendo informes de suelos")
        subir_informes_suelos()
        ListBox1.Items.Clear()
        ListBox1.Items.Add(Now)
        ListBox1.Items.Add("Subiendo informes de brucelosis")
        subir_informes_brucelosis()
        ListBox1.Items.Clear()
        ListBox1.Items.Add(Now)
        ListBox1.Items.Add("Subiendo informes de ambiental")
        subir_informes_ambiental()
    End Sub
    Private Sub SubirSaldoswebToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SubirSaldoswebToolStripMenuItem.Click
        Timer1.Enabled = False
        Timer3.Enabled = False
        Timer4.Enabled = False
        calcularsaldos()
        Timer1.Enabled = True
        Timer3.Enabled = True
        Timer4.Enabled = True
    End Sub
    Private Sub SubirCtaCteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SubirCtaCteToolStripMenuItem.Click
        subir_ctacte()
    End Sub
    Private Sub RevisarCuentasCorrientesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RevisarCuentasCorrientesToolStripMenuItem.Click
        revisar_cuentas_corrientes()
    End Sub
    Private Sub CargarSolicitudesWebToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CargarSolicitudesWebToolStripMenuItem.Click
        cargarsolicitudesweb()
    End Sub
    Private Sub SubirCtasCtes2ToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SubirCtasCtes2ToolStripMenuItem.Click
        subir_ctacte2()
    End Sub
    Private Sub SubirCtasCtesToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SubirCtasCtesToolStripMenuItem.Click
        subir_ctacte()
    End Sub



#End Region
    
    Public Function PostResponse(ByVal url As String, ByVal metodo As String, ByVal content As String, ByRef statusCode As HttpStatusCode) As Byte()
        Dim responseFromServer As Byte() = Nothing
        Dim dataStream As Stream = Nothing
        Try
            Dim request As WebRequest = WebRequest.Create(url)
            request.Timeout = 120000
            request.Method = metodo
            Dim byteArray As Byte() = System.Text.Encoding.UTF8.GetBytes(content)
            request.ContentType = "application/json"
            request.ContentLength = byteArray.Length
            dataStream = request.GetRequestStream()
            dataStream.Write(byteArray, 0, byteArray.Length)
            dataStream.Close()
            Dim response As WebResponse = request.GetResponse()
            dataStream = response.GetResponseStream()
            Dim ms As New MemoryStream()
            Dim thisRead As Integer = 0
            Dim buff As Byte() = New Byte(1023) {}
            Do
                thisRead = dataStream.Read(buff, 0, buff.Length)
                If thisRead = 0 Then
                    Exit Do
                End If
                ms.Write(buff, 0, thisRead)
            Loop While True
            responseFromServer = ms.ToArray()
            dataStream.Close()
            response.Close()
            statusCode = HttpStatusCode.OK
        Catch ex As WebException
            If ex.Response IsNot Nothing Then
                dataStream = ex.Response.GetResponseStream()
                Dim reader As New StreamReader(dataStream)
                Dim resp As String = reader.ReadToEnd()
                statusCode = DirectCast(ex.Response, HttpWebResponse).StatusCode
            Else
                Dim resp As String = ""
                statusCode = HttpStatusCode.ExpectationFailed
            End If
        Catch ex As Exception
            statusCode = HttpStatusCode.ExpectationFailed
        End Try
        Return responseFromServer
    End Function
    
    Public Function GetResponse(ByVal url As String, ByVal metodo As String, ByRef statusCode As HttpStatusCode) As Byte()
        Dim responseFromServer As Byte() = Nothing
        Dim dataStream As Stream = Nothing
        Try
            Dim request As WebRequest = WebRequest.Create(url)
            request.Timeout = 120000
            request.Method = metodo
            request.ContentType = "application/json"



            request.UseDefaultCredentials = True
            request.PreAuthenticate = True
            request.Credentials = CredentialCache.DefaultCredentials

            Dim response As WebResponse = request.GetResponse()
            dataStream = response.GetResponseStream()
            Dim ms As New MemoryStream()
            Dim thisRead As Integer = 0
            Dim buff As Byte() = New Byte(1023) {}
            Do
                thisRead = dataStream.Read(buff, 0, buff.Length)
                If thisRead = 0 Then
                    Exit Do
                End If
                ms.Write(buff, 0, thisRead)
            Loop While True
            responseFromServer = ms.ToArray()
            dataStream.Close()
            response.Close()
            statusCode = HttpStatusCode.OK
        Catch ex As WebException
            If ex.Response IsNot Nothing Then
                dataStream = ex.Response.GetResponseStream()
                Dim reader As New StreamReader(dataStream)
                Dim resp As String = reader.ReadToEnd()
                statusCode = DirectCast(ex.Response, HttpWebResponse).StatusCode
            Else
                Dim resp As String = ""
                statusCode = HttpStatusCode.ExpectationFailed
            End If
        Catch ex As Exception
            statusCode = HttpStatusCode.ExpectationFailed
        End Try
        Return responseFromServer
    End Function
 
End Class
