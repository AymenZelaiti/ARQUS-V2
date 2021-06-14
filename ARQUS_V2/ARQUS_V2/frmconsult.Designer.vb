<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmconsult
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmconsult))
        Me.btnprevious = New System.Windows.Forms.Button()
        Me.btnaffichecrq = New System.Windows.Forms.Button()
        Me.btnnext = New System.Windows.Forms.Button()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.tbhoraire = New System.Windows.Forms.TabPage()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.tbphenomenes = New System.Windows.Forms.TabPage()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.DataGridView3 = New System.Windows.Forms.DataGridView()
        Me.DataGridView2 = New System.Windows.Forms.DataGridView()
        Me.lbldate_of_crq = New System.Windows.Forms.Label()
        Me.btntbphenomene = New System.Windows.Forms.Button()
        Me.btntbhoraire = New System.Windows.Forms.Button()
        Me.btnquitconsult = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.TabControl1.SuspendLayout()
        Me.tbhoraire.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tbphenomenes.SuspendLayout()
        CType(Me.DataGridView3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnprevious
        '
        Me.btnprevious.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnprevious.Location = New System.Drawing.Point(274, 44)
        Me.btnprevious.Name = "btnprevious"
        Me.btnprevious.Size = New System.Drawing.Size(158, 34)
        Me.btnprevious.TabIndex = 0
        Me.btnprevious.Text = "<<CRQ précedent<<"
        Me.btnprevious.UseVisualStyleBackColor = True
        '
        'btnaffichecrq
        '
        Me.btnaffichecrq.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnaffichecrq.Location = New System.Drawing.Point(441, 44)
        Me.btnaffichecrq.Name = "btnaffichecrq"
        Me.btnaffichecrq.Size = New System.Drawing.Size(158, 34)
        Me.btnaffichecrq.TabIndex = 1
        Me.btnaffichecrq.Text = "CRQ d'aujourd'huit"
        Me.btnaffichecrq.UseVisualStyleBackColor = True
        '
        'btnnext
        '
        Me.btnnext.Font = New System.Drawing.Font("Verdana", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnnext.Location = New System.Drawing.Point(608, 44)
        Me.btnnext.Name = "btnnext"
        Me.btnnext.Size = New System.Drawing.Size(158, 34)
        Me.btnnext.TabIndex = 2
        Me.btnnext.Text = ">>CRQ suivant>>"
        Me.btnnext.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.tbhoraire)
        Me.TabControl1.Controls.Add(Me.tbphenomenes)
        Me.TabControl1.Location = New System.Drawing.Point(12, 96)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(1083, 617)
        Me.TabControl1.TabIndex = 3
        '
        'tbhoraire
        '
        Me.tbhoraire.Controls.Add(Me.DataGridView1)
        Me.tbhoraire.Location = New System.Drawing.Point(4, 22)
        Me.tbhoraire.Name = "tbhoraire"
        Me.tbhoraire.Padding = New System.Windows.Forms.Padding(3)
        Me.tbhoraire.Size = New System.Drawing.Size(1075, 591)
        Me.tbhoraire.TabIndex = 0
        Me.tbhoraire.Text = "TabPage1"
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(6, 6)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(1063, 579)
        Me.DataGridView1.TabIndex = 0
        '
        'tbphenomenes
        '
        Me.tbphenomenes.Controls.Add(Me.Label3)
        Me.tbphenomenes.Controls.Add(Me.Label2)
        Me.tbphenomenes.Controls.Add(Me.DataGridView3)
        Me.tbphenomenes.Controls.Add(Me.DataGridView2)
        Me.tbphenomenes.Location = New System.Drawing.Point(4, 22)
        Me.tbphenomenes.Name = "tbphenomenes"
        Me.tbphenomenes.Padding = New System.Windows.Forms.Padding(3)
        Me.tbphenomenes.Size = New System.Drawing.Size(1075, 591)
        Me.tbphenomenes.TabIndex = 1
        Me.tbphenomenes.Text = "TabPage2"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Verdana", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.Indigo
        Me.Label3.Location = New System.Drawing.Point(466, 454)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(129, 13)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "Extêmes et Cumuls"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Verdana", 8.25!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Indigo
        Me.Label2.Location = New System.Drawing.Point(488, 15)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(90, 13)
        Me.Label2.TabIndex = 2
        Me.Label2.Text = "Phénomènes"
        '
        'DataGridView3
        '
        Me.DataGridView3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView3.Location = New System.Drawing.Point(6, 481)
        Me.DataGridView3.Name = "DataGridView3"
        Me.DataGridView3.Size = New System.Drawing.Size(1063, 104)
        Me.DataGridView3.TabIndex = 1
        '
        'DataGridView2
        '
        Me.DataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView2.Location = New System.Drawing.Point(6, 40)
        Me.DataGridView2.Name = "DataGridView2"
        Me.DataGridView2.Size = New System.Drawing.Size(1063, 391)
        Me.DataGridView2.TabIndex = 0
        '
        'lbldate_of_crq
        '
        Me.lbldate_of_crq.AutoSize = True
        Me.lbldate_of_crq.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbldate_of_crq.ForeColor = System.Drawing.Color.Teal
        Me.lbldate_of_crq.Location = New System.Drawing.Point(39, 53)
        Me.lbldate_of_crq.Name = "lbldate_of_crq"
        Me.lbldate_of_crq.Size = New System.Drawing.Size(0, 16)
        Me.lbldate_of_crq.TabIndex = 4
        '
        'btntbphenomene
        '
        Me.btntbphenomene.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btntbphenomene.Location = New System.Drawing.Point(646, 719)
        Me.btntbphenomene.Name = "btntbphenomene"
        Me.btntbphenomene.Size = New System.Drawing.Size(220, 34)
        Me.btntbphenomene.TabIndex = 5
        Me.btntbphenomene.Text = "Phénomènes/Extrêmes" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        Me.btntbphenomene.UseVisualStyleBackColor = True
        '
        'btntbhoraire
        '
        Me.btntbhoraire.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btntbhoraire.Location = New System.Drawing.Point(872, 719)
        Me.btntbhoraire.Name = "btntbhoraire"
        Me.btntbhoraire.Size = New System.Drawing.Size(175, 34)
        Me.btntbhoraire.TabIndex = 6
        Me.btntbhoraire.Text = "Données Horaires"
        Me.btntbhoraire.UseVisualStyleBackColor = True
        '
        'btnquitconsult
        '
        Me.btnquitconsult.Font = New System.Drawing.Font("Verdana", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnquitconsult.Location = New System.Drawing.Point(12, 719)
        Me.btnquitconsult.Name = "btnquitconsult"
        Me.btnquitconsult.Size = New System.Drawing.Size(158, 34)
        Me.btnquitconsult.TabIndex = 7
        Me.btnquitconsult.Text = "Quitter"
        Me.btnquitconsult.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Verdana", 12.0!, CType((System.Drawing.FontStyle.Bold Or System.Drawing.FontStyle.Italic), System.Drawing.FontStyle), System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.Coral
        Me.Label1.Location = New System.Drawing.Point(426, 9)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(196, 18)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Consultation des CRQ"
        '
        'Panel1
        '
        Me.Panel1.Location = New System.Drawing.Point(12, 96)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(1079, 28)
        Me.Panel1.TabIndex = 9
        '
        'frmconsult
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1107, 772)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btnquitconsult)
        Me.Controls.Add(Me.btntbhoraire)
        Me.Controls.Add(Me.btntbphenomene)
        Me.Controls.Add(Me.lbldate_of_crq)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.btnnext)
        Me.Controls.Add(Me.btnaffichecrq)
        Me.Controls.Add(Me.btnprevious)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmconsult"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Automatisation des Renseignements Quotidiens des Stations (Aymen Zelaiti -INM Tun" &
    "isie)"
        Me.TabControl1.ResumeLayout(False)
        Me.tbhoraire.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tbphenomenes.ResumeLayout(False)
        Me.tbphenomenes.PerformLayout()
        CType(Me.DataGridView3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DataGridView2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnprevious As Button
    Friend WithEvents btnaffichecrq As Button
    Friend WithEvents btnnext As Button
    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents tbhoraire As TabPage
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents tbphenomenes As TabPage
    Friend WithEvents lbldate_of_crq As Label
    Friend WithEvents DataGridView2 As DataGridView
    Friend WithEvents btntbphenomene As Button
    Friend WithEvents btntbhoraire As Button
    Friend WithEvents btnquitconsult As Button
    Friend WithEvents DataGridView3 As DataGridView
    Friend WithEvents Label1 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Panel1 As Panel
End Class
