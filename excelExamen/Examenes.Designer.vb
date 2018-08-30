<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class excel
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(excel))
        Me.initExcel = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'initExcel
        '
        Me.initExcel.BackColor = System.Drawing.Color.Green
        Me.initExcel.Cursor = System.Windows.Forms.Cursors.Hand
        Me.initExcel.FlatAppearance.BorderSize = 0
        Me.initExcel.ForeColor = System.Drawing.Color.White
        Me.initExcel.Location = New System.Drawing.Point(150, 50)
        Me.initExcel.Margin = New System.Windows.Forms.Padding(0)
        Me.initExcel.Name = "initExcel"
        Me.initExcel.Size = New System.Drawing.Size(100, 50)
        Me.initExcel.TabIndex = 0
        Me.initExcel.Text = "Examen Excel"
        Me.initExcel.UseVisualStyleBackColor = False
        '
        'excel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(384, 361)
        Me.Controls.Add(Me.initExcel)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "excel"
        Me.Text = "Examenes - CI"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents initExcel As Button
End Class
