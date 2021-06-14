Public Class frmconsult
    Dim dd As Integer
    Dim IDstation As String
    Dim disck As String
    Dim Folder_crq As String = "Data\données station"

    Private Sub btnaffichecrq_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnaffichecrq.Click

        dd = 0
        btnnext.Enabled = False
        Dim date_to_show As String = DateTime.Now.ToString("dd/MM/yyyy")

        Try

            Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim dataSet As System.Data.DataSet
            Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
            Dim path As String = [String].Format(disck & ":\" & Folder_crq & "\" & IDstation & "crq{0}.xls", DateTime.Now.ToString("ddMMyyyy"))

            MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1;';") 'access database engine data connectivity component
            MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [horaire$C3:AC27]", MyConnection)

            dataSet = New System.Data.DataSet
            MyCommand.Fill(dataSet)
            DataGridView1.DataSource = dataSet.Tables(0)

            MyConnection.Close()
            lbldate_of_crq.Text = date_to_show
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try

        Try

            Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim dataSet As System.Data.DataSet
            Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
            Dim path As String = [String].Format(disck & ":\" & Folder_crq & "\" & IDstation & "crq{0}.xls", DateTime.Now.ToString("ddMMyyyy"))

            MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1;';") 'access database engine data connectivity component
            MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [phenomenes$C1:K500]", MyConnection)

            dataSet = New System.Data.DataSet
            MyCommand.Fill(dataSet)
            DataGridView2.DataSource = dataSet.Tables(0)

            MyConnection.Close()
            lbldate_of_crq.Text = date_to_show
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try
        Try
            Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim dataSet As System.Data.DataSet
            Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
            Dim path As String = [String].Format(disck & ":\" & Folder_crq & "\" & IDstation & "crq{0}.xls", DateTime.Now.ToString("ddMMyyyy"))

            MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1;';") 'access database engine data connectivity component
            MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [extrèmes$A1:W2]", MyConnection)
            dataSet = New System.Data.DataSet
            MyCommand.Fill(dataSet)
            DataGridView3.DataSource = dataSet.Tables(0)
            MyConnection.Close()
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try
    End Sub

    Private Sub btnquitconsult_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnquitconsult.Click

        saisietdh.Show()
        Me.Hide()
    End Sub

    Private Sub frmconsult_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        saisietdh.Hide()
        TabControl1.SelectedTab = tbhoraire
        dd = 0
        DataGridView1.EditMode = DataGridViewEditMode.EditProgrammatically
        DataGridView2.EditMode = DataGridViewEditMode.EditProgrammatically
        DataGridView3.EditMode = DataGridViewEditMode.EditProgrammatically
        btnnext.Enabled = False
        Dim date_to_show As String = DateTime.Now.ToString("dd/MM/yyyy")
        IDstation = My.Settings.IDst
        disck = My.Settings.disck

        Try

            Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim dataSet As System.Data.DataSet
            Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
            Dim path As String = [String].Format(disck & ":\" & Folder_crq & "\" & IDstation & "crq{0}.xls", DateTime.Now.ToString("ddMMyyyy"))

            MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1;';") 'access database engine data connectivity component
            MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [horaire$C3:AC27]", MyConnection)

            dataSet = New System.Data.DataSet
            MyCommand.Fill(dataSet)
            DataGridView1.DataSource = dataSet.Tables(0)

            MyConnection.Close()
            lbldate_of_crq.Text = date_to_show
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try

        Try

            Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim dataSet As System.Data.DataSet
            Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
            Dim path As String = [String].Format(disck & ":\" & Folder_crq & "\" & IDstation & "crq{0}.xls", DateTime.Now.ToString("ddMMyyyy"))

            MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1;';") 'access database engine data connectivity component
            MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [phenomenes$C1:K500]", MyConnection)

            dataSet = New System.Data.DataSet
            MyCommand.Fill(dataSet)
            DataGridView2.DataSource = dataSet.Tables(0)

            MyConnection.Close()
            lbldate_of_crq.Text = date_to_show
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try
        Try
            Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim dataSet As System.Data.DataSet
            Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
            Dim path As String = [String].Format(disck & ":\" & Folder_crq & "\" & IDstation & "crq{0}.xls", DateTime.Now.ToString("ddMMyyyy"))

            MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1;';") 'access database engine data connectivity component
            MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [extrèmes$A1:W2]", MyConnection)
            dataSet = New System.Data.DataSet
            MyCommand.Fill(dataSet)
            DataGridView3.DataSource = dataSet.Tables(0)
            MyConnection.Close()
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try

    End Sub

    Protected Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim param As CreateParams = MyBase.CreateParams
            param.ClassStyle = param.ClassStyle Or &H200
            Return param
        End Get
    End Property


    Private Sub btnprevious_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnprevious.Click

        dd -= 1


        btnnext.Enabled = True

        Dim dat1 As String = DateTime.Now.AddDays(dd).ToString("ddMMyyyy")
        Dim date_to_show As String = DateTime.Now.AddDays(dd).ToString("dd/MM/yyyy")

        Try

            Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim dataSet As System.Data.DataSet
            Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
            Dim path As String = [String].Format(disck & ":\" & Folder_crq & "\" & IDstation & "crq{0}.xls", dat1)

            MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1;';") 'access database engine data connectivity component
            MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [horaire$C3:AC27]", MyConnection)

            dataSet = New System.Data.DataSet
            MyCommand.Fill(dataSet)
            DataGridView1.DataSource = dataSet.Tables(0)

            MyConnection.Close()
            lbldate_of_crq.Text = date_to_show

        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try

        Try

            Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim dataSet As System.Data.DataSet
            Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
            Dim path As String = [String].Format(disck & ":\" & Folder_crq & "\" & IDstation & "crq{0}.xls", dat1)

            MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1;';") 'access database engine data connectivity component
            MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [phenomenes$C1:K500]", MyConnection)

            dataSet = New System.Data.DataSet
            MyCommand.Fill(dataSet)
            DataGridView2.DataSource = dataSet.Tables(0)

            MyConnection.Close()

        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try
        Try
            Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim dataSet As System.Data.DataSet
            Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
            Dim path As String = [String].Format(disck & ":\" & Folder_crq & "\" & IDstation & "crq{0}.xls", dat1)

            MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1;';") 'access database engine data connectivity component
            MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [extrèmes$A1:W2]", MyConnection)
            dataSet = New System.Data.DataSet
            MyCommand.Fill(dataSet)
            DataGridView3.DataSource = dataSet.Tables(0)
            MyConnection.Close()
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try

    End Sub

    Private Sub btnnext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnnext.Click

        dd += 1

        Dim dat1 As String = DateTime.Now.AddDays(dd).ToString("ddMMyyyy")
        Dim date_to_show As String = DateTime.Now.AddDays(dd).ToString("dd/MM/yyyy")

        Try

            Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim dataSet As System.Data.DataSet
            Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
            Dim path As String = [String].Format(disck & ":\" & Folder_crq & "\" & IDstation & "crq{0}.xls", dat1)

            MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1;';") 'access database engine data connectivity component
            MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [horaire$C3:AC27]", MyConnection)

            dataSet = New System.Data.DataSet
            MyCommand.Fill(dataSet)
            DataGridView1.DataSource = dataSet.Tables(0)

            MyConnection.Close()
            lbldate_of_crq.Text = date_to_show

        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try

        Try

            Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim dataSet As System.Data.DataSet
            Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
            Dim path As String = [String].Format(disck & ":\" & Folder_crq & "\" & IDstation & "crq{0}.xls", dat1)

            MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1;';") 'access database engine data connectivity component
            MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [phenomenes$C1:K500]", MyConnection)

            dataSet = New System.Data.DataSet
            MyCommand.Fill(dataSet)
            DataGridView2.DataSource = dataSet.Tables(0)

            MyConnection.Close()

        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try
        Try
            Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim dataSet As System.Data.DataSet
            Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
            Dim path As String = [String].Format(disck & ":\" & Folder_crq & "\" & IDstation & "crq{0}.xls", dat1)

            MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1;';") 'access database engine data connectivity component
            MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [extrèmes$A1:W2]", MyConnection)
            dataSet = New System.Data.DataSet
            MyCommand.Fill(dataSet)
            DataGridView3.DataSource = dataSet.Tables(0)
            MyConnection.Close()
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try


        If dd = Date.Now.Day Then
            btnnext.Enabled = False
        End If

    End Sub

    Private Sub btntbhoraire_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btntbhoraire.Click

        TabControl1.SelectedTab = tbhoraire

    End Sub

    Private Sub btntbphenomene_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btntbphenomene.Click

        TabControl1.SelectedTab = tbphenomenes

    End Sub

End Class


