<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormInicio
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.DateFecha = New System.Windows.Forms.DateTimePicker()
        Me.Timer3 = New System.Windows.Forms.Timer(Me.components)
        Me.Timer4 = New System.Windows.Forms.Timer(Me.components)
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.AccionesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ImportarToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MoverArchivosSubidosToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SubirInformesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SubirSaldoswebToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SubirCtaCteToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.RevisarCuentasCorrientesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CargarSolicitudesWebToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SubirCtasCtesToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SubirCtasCtes2ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.ProgressBar1 = New System.Windows.Forms.ProgressBar()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'DateFecha
        '
        Me.DateFecha.Format = System.Windows.Forms.DateTimePickerFormat.[Short]
        Me.DateFecha.Location = New System.Drawing.Point(12, 39)
        Me.DateFecha.Name = "DateFecha"
        Me.DateFecha.Size = New System.Drawing.Size(113, 20)
        Me.DateFecha.TabIndex = 27
        '
        'Timer3
        '
        Me.Timer3.Interval = 10000
        '
        'Timer4
        '
        Me.Timer4.Interval = 10000
        '
        'Timer1
        '
        Me.Timer1.Interval = 10000
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.AccionesToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(312, 24)
        Me.MenuStrip1.TabIndex = 28
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'AccionesToolStripMenuItem
        '
        Me.AccionesToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ImportarToolStripMenuItem, Me.MoverArchivosSubidosToolStripMenuItem, Me.SubirInformesToolStripMenuItem, Me.SubirSaldoswebToolStripMenuItem, Me.SubirCtaCteToolStripMenuItem, Me.RevisarCuentasCorrientesToolStripMenuItem, Me.CargarSolicitudesWebToolStripMenuItem, Me.SubirCtasCtesToolStripMenuItem, Me.SubirCtasCtes2ToolStripMenuItem})
        Me.AccionesToolStripMenuItem.Name = "AccionesToolStripMenuItem"
        Me.AccionesToolStripMenuItem.Size = New System.Drawing.Size(67, 20)
        Me.AccionesToolStripMenuItem.Text = "Acciones"
        '
        'ImportarToolStripMenuItem
        '
        Me.ImportarToolStripMenuItem.Name = "ImportarToolStripMenuItem"
        Me.ImportarToolStripMenuItem.Size = New System.Drawing.Size(210, 22)
        Me.ImportarToolStripMenuItem.Text = "Importar"
        '
        'MoverArchivosSubidosToolStripMenuItem
        '
        Me.MoverArchivosSubidosToolStripMenuItem.Name = "MoverArchivosSubidosToolStripMenuItem"
        Me.MoverArchivosSubidosToolStripMenuItem.Size = New System.Drawing.Size(210, 22)
        Me.MoverArchivosSubidosToolStripMenuItem.Text = "Mover archivos subidos"
        '
        'SubirInformesToolStripMenuItem
        '
        Me.SubirInformesToolStripMenuItem.Name = "SubirInformesToolStripMenuItem"
        Me.SubirInformesToolStripMenuItem.Size = New System.Drawing.Size(210, 22)
        Me.SubirInformesToolStripMenuItem.Text = "Subir informes"
        '
        'SubirSaldoswebToolStripMenuItem
        '
        Me.SubirSaldoswebToolStripMenuItem.Name = "SubirSaldoswebToolStripMenuItem"
        Me.SubirSaldoswebToolStripMenuItem.Size = New System.Drawing.Size(210, 22)
        Me.SubirSaldoswebToolStripMenuItem.Text = "Subir saldos (web)"
        '
        'SubirCtaCteToolStripMenuItem
        '
        Me.SubirCtaCteToolStripMenuItem.Name = "SubirCtaCteToolStripMenuItem"
        Me.SubirCtaCteToolStripMenuItem.Size = New System.Drawing.Size(210, 22)
        Me.SubirCtaCteToolStripMenuItem.Text = "Subir cta cte"
        '
        'RevisarCuentasCorrientesToolStripMenuItem
        '
        Me.RevisarCuentasCorrientesToolStripMenuItem.Name = "RevisarCuentasCorrientesToolStripMenuItem"
        Me.RevisarCuentasCorrientesToolStripMenuItem.Size = New System.Drawing.Size(210, 22)
        Me.RevisarCuentasCorrientesToolStripMenuItem.Text = "Revisar cuentas corrientes"
        '
        'CargarSolicitudesWebToolStripMenuItem
        '
        Me.CargarSolicitudesWebToolStripMenuItem.Name = "CargarSolicitudesWebToolStripMenuItem"
        Me.CargarSolicitudesWebToolStripMenuItem.Size = New System.Drawing.Size(210, 22)
        Me.CargarSolicitudesWebToolStripMenuItem.Text = "Cargar solicitudes web"
        '
        'SubirCtasCtesToolStripMenuItem
        '
        Me.SubirCtasCtesToolStripMenuItem.Name = "SubirCtasCtesToolStripMenuItem"
        Me.SubirCtasCtesToolStripMenuItem.Size = New System.Drawing.Size(210, 22)
        Me.SubirCtasCtesToolStripMenuItem.Text = "Subir ctas ctes"
        '
        'SubirCtasCtes2ToolStripMenuItem
        '
        Me.SubirCtasCtes2ToolStripMenuItem.Name = "SubirCtasCtes2ToolStripMenuItem"
        Me.SubirCtasCtes2ToolStripMenuItem.Size = New System.Drawing.Size(210, 22)
        Me.SubirCtasCtes2ToolStripMenuItem.Text = "Subir ctas ctes 2"
        '
        'Timer2
        '
        Me.Timer2.Interval = 10000
        '
        'ListBox1
        '
        Me.ListBox1.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(192, Byte), Integer))
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Location = New System.Drawing.Point(12, 94)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(288, 43)
        Me.ListBox1.TabIndex = 29
        '
        'ProgressBar1
        '
        Me.ProgressBar1.Location = New System.Drawing.Point(12, 65)
        Me.ProgressBar1.Name = "ProgressBar1"
        Me.ProgressBar1.Size = New System.Drawing.Size(288, 23)
        Me.ProgressBar1.TabIndex = 30
        '
        'FormInicio
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(312, 152)
        Me.Controls.Add(Me.ProgressBar1)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.DateFecha)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "FormInicio"
        Me.Text = "ROBOT"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DateFecha As System.Windows.Forms.DateTimePicker
    Friend WithEvents Timer3 As System.Windows.Forms.Timer
    Friend WithEvents Timer4 As System.Windows.Forms.Timer
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents AccionesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MoverArchivosSubidosToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ImportarToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SubirInformesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SubirSaldoswebToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SubirCtaCteToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Timer2 As System.Windows.Forms.Timer
    Friend WithEvents RevisarCuentasCorrientesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CargarSolicitudesWebToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents ProgressBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents SubirCtasCtes2ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SubirCtasCtesToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem

End Class
