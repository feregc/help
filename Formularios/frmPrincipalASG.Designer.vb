<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPrincipalASG
    Inherits System.Windows.Forms.Form

    'Form reemplaza a Dispose para limpiar la lista de componentes.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Requerido por el Diseñador de Windows Forms
    Private components As System.ComponentModel.IContainer

    'NOTA: el Diseñador de Windows Forms necesita el siguiente procedimiento
    'Se puede modificar usando el Diseñador de Windows Forms.  
    'No lo modifique con el editor de código.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPrincipalASG))
        Me.lv = New System.Windows.Forms.ListView()
        Me.Logtxtbox = New System.Windows.Forms.RichTextBox()
        Me.Iconos = New System.Windows.Forms.ImageList(Me.components)
        Me.tiempolog = New System.Windows.Forms.Timer(Me.components)
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.Ciclo = New System.Windows.Forms.Timer(Me.components)
        Me.SuspendLayout()
        '
        'lv
        '
        Me.lv.Location = New System.Drawing.Point(12, 277)
        Me.lv.Name = "lv"
        Me.lv.Size = New System.Drawing.Size(637, 17)
        Me.lv.TabIndex = 0
        Me.lv.UseCompatibleStateImageBehavior = False
        Me.lv.Visible = False
        '
        'Logtxtbox
        '
        Me.Logtxtbox.Location = New System.Drawing.Point(12, 12)
        Me.Logtxtbox.Name = "Logtxtbox"
        Me.Logtxtbox.Size = New System.Drawing.Size(637, 234)
        Me.Logtxtbox.TabIndex = 3
        Me.Logtxtbox.Text = ""
        '
        'Iconos
        '
        Me.Iconos.ImageStream = CType(resources.GetObject("Iconos.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.Iconos.TransparentColor = System.Drawing.Color.Transparent
        Me.Iconos.Images.SetKeyName(0, "CPApp.ico")
        Me.Iconos.Images.SetKeyName(1, "advancedsettings.png")
        Me.Iconos.Images.SetKeyName(2, "Busy.ico")
        Me.Iconos.Images.SetKeyName(3, "P.ico")
        '
        'tiempolog
        '
        Me.tiempolog.Enabled = True
        Me.tiempolog.Interval = 60000
        '
        'Button2
        '
        Me.Button2.Image = Global.Cargador_ASG.My.Resources.Resources.desktop_enhancements
        Me.Button2.Location = New System.Drawing.Point(655, 172)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(75, 74)
        Me.Button2.TabIndex = 7
        Me.ToolTip1.SetToolTip(Me.Button2, "Ver Log")
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Image = Global.Cargador_ASG.My.Resources.Resources.agt_utilities
        Me.Button1.Location = New System.Drawing.Point(655, 92)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 74)
        Me.Button1.TabIndex = 5
        Me.ToolTip1.SetToolTip(Me.Button1, "Configuración")
        Me.Button1.UseVisualStyleBackColor = True
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.Text = "ASG-BO"
        Me.NotifyIcon1.Visible = True
        '
        'Ciclo
        '
        Me.Ciclo.Enabled = True
        Me.Ciclo.Interval = 15000
        '
        'frmPrincipalASG
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(738, 256)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Logtxtbox)
        Me.Controls.Add(Me.lv)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "frmPrincipalASG"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Cargador - ASG"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents lv As System.Windows.Forms.ListView
    Friend WithEvents Logtxtbox As System.Windows.Forms.RichTextBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Iconos As System.Windows.Forms.ImageList
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents tiempolog As System.Windows.Forms.Timer
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents NotifyIcon1 As System.Windows.Forms.NotifyIcon
    Friend WithEvents Ciclo As System.Windows.Forms.Timer

End Class
