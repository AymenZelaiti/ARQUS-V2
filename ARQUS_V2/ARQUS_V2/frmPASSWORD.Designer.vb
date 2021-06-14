<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmPASSWORD
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.txtPASSWRD = New System.Windows.Forms.TextBox()
        Me.btnexit = New System.Windows.Forms.Button()
        Me.btnpass = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'txtPASSWRD
        '
        Me.txtPASSWRD.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPASSWRD.Location = New System.Drawing.Point(169, 85)
        Me.txtPASSWRD.Name = "txtPASSWRD"
        Me.txtPASSWRD.Size = New System.Drawing.Size(286, 22)
        Me.txtPASSWRD.TabIndex = 1
        Me.txtPASSWRD.UseSystemPasswordChar = True
        '
        'btnexit
        '
        Me.btnexit.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnexit.Location = New System.Drawing.Point(38, 129)
        Me.btnexit.Name = "btnexit"
        Me.btnexit.Size = New System.Drawing.Size(121, 43)
        Me.btnexit.TabIndex = 3
        Me.btnexit.Text = "Annuler"
        Me.btnexit.UseVisualStyleBackColor = True
        '
        'btnpass
        '
        Me.btnpass.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnpass.Location = New System.Drawing.Point(469, 129)
        Me.btnpass.Name = "btnpass"
        Me.btnpass.Size = New System.Drawing.Size(121, 43)
        Me.btnpass.TabIndex = 2
        Me.btnpass.Text = "Valider"
        Me.btnpass.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Verdana", 9.75!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Red
        Me.Label1.Location = New System.Drawing.Point(273, 43)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(104, 16)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "Mot de Passe"
        '
        'frmPASSWORD
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(632, 215)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnpass)
        Me.Controls.Add(Me.btnexit)
        Me.Controls.Add(Me.txtPASSWRD)
        Me.Name = "frmPASSWORD"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Config Pass Word"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtPASSWRD As TextBox
    Friend WithEvents btnexit As Button
    Friend WithEvents btnpass As Button
    Friend WithEvents Label1 As Label
End Class
