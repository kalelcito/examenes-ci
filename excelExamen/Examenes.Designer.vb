﻿<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
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
        Me.initWord = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'initExcel
        '
        Me.initExcel.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.initExcel.BackColor = System.Drawing.Color.Green
        Me.initExcel.Cursor = System.Windows.Forms.Cursors.Hand
        Me.initExcel.FlatAppearance.BorderSize = 0
        Me.initExcel.ForeColor = System.Drawing.Color.White
        Me.initExcel.Location = New System.Drawing.Point(40, 40)
        Me.initExcel.Margin = New System.Windows.Forms.Padding(0)
        Me.initExcel.Name = "initExcel"
        Me.initExcel.Size = New System.Drawing.Size(200, 50)
        Me.initExcel.TabIndex = 0
        Me.initExcel.Text = "Examen Excel"
        Me.initExcel.UseVisualStyleBackColor = False
        '
        'initWord
        '
        Me.initWord.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.initWord.BackColor = System.Drawing.Color.DodgerBlue
        Me.initWord.Cursor = System.Windows.Forms.Cursors.Hand
        Me.initWord.FlatAppearance.BorderSize = 0
        Me.initWord.ForeColor = System.Drawing.Color.White
        Me.initWord.Location = New System.Drawing.Point(42, 105)
        Me.initWord.Margin = New System.Windows.Forms.Padding(0)
        Me.initWord.Name = "initWord"
        Me.initWord.Size = New System.Drawing.Size(200, 50)
        Me.initWord.TabIndex = 1
        Me.initWord.Text = "Examen Word"
        Me.initWord.UseVisualStyleBackColor = False
        Me.initWord.Visible = False
        '
        'excel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(284, 261)
        Me.Controls.Add(Me.initWord)
        Me.Controls.Add(Me.initExcel)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "excel"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Examenes - CI"
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents initExcel As Button
    Friend WithEvents initWord As Button
End Class
