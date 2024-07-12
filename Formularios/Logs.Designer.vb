<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Logs
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.LogtxtBoxReader = New System.Windows.Forms.RichTextBox
        Me.ListaArchivos = New System.Windows.Forms.ListView
        Me.SuspendLayout()
        '
        'LogtxtBoxReader
        '
        Me.LogtxtBoxReader.Location = New System.Drawing.Point(12, 262)
        Me.LogtxtBoxReader.Name = "LogtxtBoxReader"
        Me.LogtxtBoxReader.ReadOnly = True
        Me.LogtxtBoxReader.Size = New System.Drawing.Size(637, 244)
        Me.LogtxtBoxReader.TabIndex = 4
        Me.LogtxtBoxReader.Text = ""
        '
        'ListaArchivos
        '
        Me.ListaArchivos.Location = New System.Drawing.Point(12, 12)
        Me.ListaArchivos.MultiSelect = False
        Me.ListaArchivos.Name = "ListaArchivos"
        Me.ListaArchivos.Size = New System.Drawing.Size(637, 244)
        Me.ListaArchivos.TabIndex = 5
        Me.ListaArchivos.UseCompatibleStateImageBehavior = False
        Me.ListaArchivos.View = System.Windows.Forms.View.List
        '
        'Logs
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(659, 518)
        Me.Controls.Add(Me.ListaArchivos)
        Me.Controls.Add(Me.LogtxtBoxReader)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.MaximizeBox = False
        Me.Name = "Logs"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Logs"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents LogtxtBoxReader As System.Windows.Forms.RichTextBox
    Friend WithEvents ListaArchivos As System.Windows.Forms.ListView
End Class
