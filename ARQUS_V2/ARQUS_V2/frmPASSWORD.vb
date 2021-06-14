Public Class frmPASSWORD

    Private Sub btnpass_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnpass.Click
        If txtPASSWRD.Text = "789*+" Then
            Me.Hide()
            FrmConfig.ShowDialog()
        Else
            MsgBox("Mot de passe non valide !", MsgBoxStyle.Information, vbOKOnly)

        End If
    End Sub

    Private Sub btnexit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnexit.Click
        txtPASSWRD.Text = ""
        Me.Hide()

    End Sub

    Private Sub frmPASSWORD_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        txtPASSWRD.Text = ""

    End Sub

    Protected Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim param As CreateParams = MyBase.CreateParams
            param.ClassStyle = param.ClassStyle Or &H200
            Return param
        End Get
    End Property
End Class