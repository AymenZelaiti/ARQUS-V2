Imports Microsoft.Office.Interop
Imports System.Security.Permissions
Imports System.IO
Imports System.Globalization
Imports System.Data.OleDb
Imports Microsoft.SqlServer

Public Class frmrecup
    Dim numbers() As Char = {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}
    Dim numberup4() As Char = {"5", "6", "7", "8", "9", "0"}
    Dim numberdn4() As Char = {"0", "1", "2", "3", "4"}
    Dim IDstation As String
    Dim disck As String
    Const Folder_crq As String = "Data\données station"
    Const BD_crq As String = ":\Data\bdcrq\"
    Const source As String = "\aero_mes\PISTE0\"
    Dim ADRESS As String
    Dim sourceADRESS As String
    Dim config As Boolean
    Dim SmDD As Integer
    Dim SmFF As Integer
    Dim SmT As Integer
    Dim SmTd As Integer
    Dim SmHr As Integer
    Dim SmPs As Integer
    Dim SmPm As Integer
    Dim ShEw As Integer
    Dim ShRR As Integer
    Dim KTn As Integer
    Dim KhTn As Integer
    Dim KTx As Integer
    Dim KhTx As Integer
    Dim KUn As Integer
    Dim KhUn As Integer
    Dim KUx As Integer
    Dim KhUx As Integer
    Dim KRR As Integer
    Dim KIs As Integer
    Dim KRg As Integer
    Dim KTxSol As Integer
    Dim KTnSol As Integer


    Private Sub frmrecup_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        saisietdh.Hide()
        Timer1.Start()
        DataGridView1.EditMode = DataGridViewEditMode.EditProgrammatically

        IDstation = My.Settings.IDst
        disck = My.Settings.disck

        SmDD = My.Settings.SmDD
        SmFF = SmDD + 1
        SmT = My.Settings.SmT
        SmTd = SmT + 2
        SmHr = My.Settings.SmHR
        SmPs = My.Settings.SmPs
        SmPm = SmPs + 1
        ShEw = My.Settings.ShEw
        ShRR = My.Settings.ShRR

        KTx = My.Settings.KTx
        KhTx = KTx + 1
        KTn = My.Settings.KTn
        KhTn = KTn + 1
        KUx = My.Settings.KUx
        KhUx = KUx + 1
        KUn = My.Settings.KUn
        KhUn = KUn + 1
        KIs = My.Settings.KIs
        KRg = My.Settings.KRg
        KRR = My.Settings.KRR
        KTnSol = My.Settings.KTnSol
        KTxSol = KTnSol + 2

        ADRESS = My.Settings.adressCAOBS
        sourceADRESS = "\\" & ADRESS & source
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
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try


    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        lbltime.Text = Date.Now

    End Sub

    Private Sub btnquitrecup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnquitrecup.Click
        saisietdh.Show()
        Me.Hide()

    End Sub


    Private Sub btnmakecrq_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnmakecrq.Click

        Dim selecdate As String = Format(DateTimePicker2.Value, "ddMMyyyy")

        If DateTimePicker2.Value.Date > Date.Now.Date Then
            MsgBox("La date selectionnée est invalide, selectionnez une date dans le passé !", MsgBoxStyle.Exclamation, vbOKOnly)
        ElseIf DateTimePicker2.Value.Date = Date.Now.Date Then
            MsgBox("La date selectionnée est la date d'aujourd'huit, selectionnez une date dans le passé !", MsgBoxStyle.Exclamation, vbOKOnly)

        ElseIf DateTimePicker2.Value.Date < Date.Now.Date Then
            Dim xlApp As Excel.Application
            Dim NEWcrq As Excel.Workbook
            Dim hrly As Excel.Worksheet
            Dim phen As Excel.Worksheet
            Dim extr As Excel.Worksheet
            Dim FileDate As String = DateTimePicker2.Value.Date.ToString("dd MMM yyyy", CultureInfo.CreateSpecificCulture("fr-FRA"))
            Dim nomfichier As String = [String].Format(disck & ":\" & Folder_crq & "\" & IDstation & "crq{0}.xls", DateTimePicker2.Value.Date.ToString("ddMMyyyy"))
            Dim permis As New FileIOPermission(FileIOPermissionAccess.AllAccess, disck & ":\" & Folder_crq)

            If System.IO.File.Exists(nomfichier) Then
                'the file exists
                MsgBox("le CRQ du :" & FileDate & " éxiste déjà,utilizez les outils de récupération Horaire ci-déssus !", MsgBoxStyle.Exclamation, vbOKOnly)
            Else
                'the file doesn't exist
                xlApp = New Excel.Application
                NEWcrq = xlApp.Workbooks.Add()
                NEWcrq.Worksheets(1).name = "horaire"
                hrly = NEWcrq.Worksheets(1)
                NEWcrq.Worksheets(2).name = "phenomenes"
                phen = NEWcrq.Worksheets(2)
                NEWcrq.Worksheets(3).name = "extrèmes"
                extr = NEWcrq.Worksheets(3)
                hrly.Cells(3, 3).value = "heure" 'mise en forme sheet-donnees horaires
                hrly.Cells(3, 3).font.italic = True
                hrly.Cells(3, 3).font.bold = True
                hrly.Cells(3, 3).font.size = 11
                hrly.Cells(4, 3).value = "00H00"
                hrly.Cells(5, 3).value = "01H00"
                hrly.Cells(6, 3).value = "02H00"
                hrly.Cells(7, 3).value = "03H00"
                hrly.Cells(8, 3).value = "04H00"
                hrly.Cells(9, 3).value = "05H00"
                hrly.Cells(10, 3).value = "06H00"
                hrly.Cells(11, 3).value = "07H00"
                hrly.Cells(12, 3).value = "08H00"
                hrly.Cells(13, 3).value = "09H00"
                hrly.Cells(14, 3).value = "10H00"
                hrly.Cells(15, 3).value = "11H00"
                hrly.Cells(16, 3).value = "12H00"
                hrly.Cells(17, 3).value = "13H00"
                hrly.Cells(18, 3).value = "14H00"
                hrly.Cells(19, 3).value = "15H00"
                hrly.Cells(20, 3).value = "16H00"
                hrly.Cells(21, 3).value = "17H00"
                hrly.Cells(22, 3).value = "18H00"
                hrly.Cells(23, 3).value = "19H00"
                hrly.Cells(24, 3).value = "20H00"
                hrly.Cells(25, 3).value = "21H00"
                hrly.Cells(26, 3).value = "22H00"
                hrly.Cells(27, 3).value = "23H00"
                hrly.Cells(3, 4).value = "Visi."
                hrly.Cells(3, 4).font.italic = True
                hrly.Cells(3, 4).font.bold = True
                hrly.Cells(3, 4).font.size = 11
                hrly.Cells(3, 5).value = "DDD"
                hrly.Cells(3, 5).font.italic = True
                hrly.Cells(3, 5).font.bold = True
                hrly.Cells(3, 5).font.size = 11
                hrly.Cells(3, 6).value = "FF"
                hrly.Cells(3, 6).font.italic = True
                hrly.Cells(3, 6).font.bold = True
                hrly.Cells(3, 6).font.size = 11
                hrly.Cells(3, 7).value = "WW"
                hrly.Cells(3, 7).font.italic = True
                hrly.Cells(3, 7).font.bold = True
                hrly.Cells(3, 7).font.size = 8
                hrly.Cells(3, 8).value = "w1w2"
                hrly.Cells(3, 8).font.italic = True
                hrly.Cells(3, 8).font.bold = True
                hrly.Cells(3, 8).font.size = 8
                hrly.Cells(3, 9).value = "CL"
                hrly.Cells(3, 9).font.italic = True
                hrly.Cells(3, 9).font.bold = True
                hrly.Cells(3, 9).font.size = 10
                hrly.Cells(3, 10).value = "CM"
                hrly.Cells(3, 10).font.italic = True
                hrly.Cells(3, 10).font.bold = True
                hrly.Cells(3, 10).font.size = 10
                hrly.Cells(3, 11).value = "CH"
                hrly.Cells(3, 11).font.italic = True
                hrly.Cells(3, 11).font.bold = True
                hrly.Cells(3, 11).font.size = 10
                hrly.Cells(3, 12).value = "N"
                hrly.Cells(3, 12).font.italic = True
                hrly.Cells(3, 12).font.bold = True
                hrly.Cells(3, 12).font.size = 10
                hrly.Cells(3, 13).value = "h1"
                hrly.Cells(3, 13).font.italic = True
                hrly.Cells(3, 13).font.bold = True
                hrly.Cells(3, 13).font.size = 10
                hrly.Cells(3, 14).value = "n1"
                hrly.Cells(3, 14).font.italic = True
                hrly.Cells(3, 14).font.bold = True
                hrly.Cells(3, 14).font.size = 10
                hrly.Cells(3, 15).value = "C1"
                hrly.Cells(3, 15).font.italic = True
                hrly.Cells(3, 15).font.bold = True
                hrly.Cells(3, 15).font.size = 10
                hrly.Cells(3, 16).value = "h0"
                hrly.Cells(3, 16).font.italic = True
                hrly.Cells(3, 16).font.bold = True
                hrly.Cells(3, 16).font.size = 10
                hrly.Cells(3, 17).value = "n0"
                hrly.Cells(3, 17).font.italic = True
                hrly.Cells(3, 17).font.bold = True
                hrly.Cells(3, 17).font.size = 10
                hrly.Cells(3, 18).value = "C0"
                hrly.Cells(3, 18).font.italic = True
                hrly.Cells(3, 18).font.bold = True
                hrly.Cells(3, 18).font.size = 10
                hrly.Cells(3, 19).value = "h2"
                hrly.Cells(3, 19).font.italic = True
                hrly.Cells(3, 19).font.bold = True
                hrly.Cells(3, 19).font.size = 10
                hrly.Cells(3, 20).value = "n2"
                hrly.Cells(3, 20).font.italic = True
                hrly.Cells(3, 20).font.bold = True
                hrly.Cells(3, 20).font.size = 10
                hrly.Cells(3, 21).value = "C2"
                hrly.Cells(3, 21).font.italic = True
                hrly.Cells(3, 21).font.bold = True
                hrly.Cells(3, 21).font.size = 10
                hrly.Cells(3, 22).value = "T.air"
                hrly.Cells(3, 22).font.italic = True
                hrly.Cells(3, 22).font.bold = True
                hrly.Cells(3, 22).font.size = 11
                hrly.Cells(3, 23).value = "ew"
                hrly.Cells(3, 23).font.italic = True
                hrly.Cells(3, 23).font.bold = True
                hrly.Cells(3, 23).font.size = 11
                hrly.Cells(3, 24).value = "Td"
                hrly.Cells(3, 24).font.italic = True
                hrly.Cells(3, 24).font.bold = True
                hrly.Cells(3, 24).font.size = 11
                hrly.Cells(3, 25).value = "U%"
                hrly.Cells(3, 25).font.italic = True
                hrly.Cells(3, 25).font.bold = True
                hrly.Cells(3, 25).font.size = 11
                hrly.Cells(3, 26).value = "P.st°"
                hrly.Cells(3, 26).font.italic = True
                hrly.Cells(3, 26).font.bold = True
                hrly.Cells(3, 26).font.size = 11
                hrly.Cells(3, 27).value = "P.mer"
                hrly.Cells(3, 27).font.italic = True
                hrly.Cells(3, 27).font.bold = True
                hrly.Cells(3, 27).font.size = 11
                hrly.Cells(3, 28).value = "RR"
                hrly.Cells(3, 28).font.italic = True
                hrly.Cells(3, 28).font.bold = True
                hrly.Cells(3, 28).font.size = 11
                hrly.Cells(3, 29).value = "dur.RR"
                hrly.Cells(3, 29).font.italic = True
                hrly.Cells(3, 29).font.bold = True
                hrly.Cells(3, 29).font.size = 11
                hrly.Columns("C:C").ColumnWidth = 5.43
                hrly.Columns("D:D").ColumnWidth = 7
                hrly.Columns("E:E").ColumnWidth = 6
                hrly.Columns("F:F").ColumnWidth = 4
                hrly.Columns("G:G").ColumnWidth = 3.86
                hrly.Columns("H:H").ColumnWidth = 4.86
                hrly.Columns("I:I").ColumnWidth = 2.14
                hrly.Columns("J:J").ColumnWidth = 2.4
                hrly.Columns("K:K").ColumnWidth = 2.29
                hrly.Columns("L:L").ColumnWidth = 2
                hrly.Columns("M:M").ColumnWidth = 6
                hrly.Columns("N:N").ColumnWidth = 2.14
                hrly.Columns("O:O").ColumnWidth = 2.14
                hrly.Columns("P:P").ColumnWidth = 6
                hrly.Columns("Q:Q").ColumnWidth = 2.14
                hrly.Columns("R:R").ColumnWidth = 2.14
                hrly.Columns("S:S").ColumnWidth = 6
                hrly.Columns("T:T").ColumnWidth = 2.14
                hrly.Columns("U:U").ColumnWidth = 2.14
                hrly.Columns("V:V").ColumnWidth = 4.7
                hrly.Columns("W:W").ColumnWidth = 4
                hrly.Columns("X:X").ColumnWidth = 4.7
                hrly.Columns("Y:Y").ColumnWidth = 4
                hrly.Columns("Z:Z").ColumnWidth = 7
                hrly.Columns("AA:AA").ColumnWidth = 7
                hrly.Columns("AB:AB").ColumnWidth = 5
                hrly.Columns("AC:AC").ColumnWidth = 7
                hrly.Cells.NumberFormat = "@"
                For i As Integer = 4 To 27
                    hrly.Cells(i, 1).value = IDstation
                    hrly.Cells(i, 2).value = DateTimePicker2.Value.Date.ToString("ddMMyyyy", CultureInfo.CreateSpecificCulture("fr-FRA"))
                Next

                hrly.Cells(1, 2).value = "Date:"
                hrly.Cells(1, 2).font.italic = True
                hrly.Cells(1, 2).font.bold = True
                hrly.Cells(1, 2).font.size = 11
                hrly.Range("C1:E1").Merge()
                hrly.Cells(1, 3).value = FileDate
                hrly.Cells(1, 3).font.bold = True
                hrly.Range("C1:E1").HorizontalAlignment = Excel.Constants.xlCenter
                hrly.Range("G2:I2").Merge()
                hrly.Cells(2, 7).value = "Station:"
                hrly.Cells(2, 7).font.bold = True
                hrly.Cells(2, 7).font.italic = True
                hrly.Range("G2:I2").HorizontalAlignment = Excel.Constants.xlCenter
                hrly.Range("J2:L2").Merge()
                hrly.Cells(2, 10).value = IDstation
                hrly.Cells(2, 10).font.bold = True
                hrly.Range("J2:L2").HorizontalAlignment = Excel.Constants.xlCenter
                phen.Cells(1, 3).value = "Phénomène"   'mise en forme sheet-phenomenes
                phen.Cells(1, 4).value = "Code"
                phen.Cells(1, 5).value = "H.début"
                phen.Cells(1, 6).value = "H.fin"
                phen.Cells(1, 7).value = "Intensité"
                phen.Cells(1, 8).value = "Secteur"
                phen.Cells(1, 9).value = "Visi.mini"
                phen.Cells(1, 10).value = "heure"
                phen.Cells(1, 11).value = "Hauteur"
                phen.Cells.NumberFormat = "@"
                phen.Rows("1:1").Font.Bold = True
                phen.Rows("1:1").Font.size = 10
                phen.Columns("A:A").ColumnWidth = 6
                phen.Columns("B:B").ColumnWidth = 10
                phen.Columns("C:C").ColumnWidth = 12
                phen.Columns("D:D").ColumnWidth = 5
                phen.Columns("E:E").ColumnWidth = 8
                phen.Columns("F:F").ColumnWidth = 8
                phen.Columns("G:G").ColumnWidth = 8
                phen.Columns("H:H").ColumnWidth = 8
                phen.Columns("I:I").ColumnWidth = 8
                phen.Columns("J:J").ColumnWidth = 8
                phen.Columns("K:K").ColumnWidth = 8
                phen.Columns("L:L").ColumnWidth = 8
                phen.Columns("M:M").ColumnWidth = 8
                phen.Columns("N:N").ColumnWidth = 8
                phen.Columns("O:O").ColumnWidth = 8
                phen.Columns("P:P").ColumnWidth = 8
                extr.Cells.NumberFormat = "@"
                NEWcrq.SaveAs(nomfichier)
                NEWcrq.Close(SaveChanges:=True)
                hrly = Nothing
                phen = Nothing
                extr = Nothing
                NEWcrq = Nothing
                xlApp = Nothing
                releaseObject(hrly)
                releaseObject(phen)
                releaseObject(extr)
                releaseObject(NEWcrq)
                releaseObject(xlApp)
                MsgBox("Le CRQ du : " & FileDate & " est crée,vous pouvez récuperer les données manquantes", MsgBoxStyle.Information, vbOKOnly)

                Try
                    Dim MyConnection As System.Data.OleDb.OleDbConnection
                    Dim dataSet As System.Data.DataSet
                    Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
                    Dim path As String = [String].Format(disck & ":\" & Folder_crq & "\" & IDstation & "crq{0}.xls", DateTimePicker2.Value.Date.ToString("ddMMyyyy"))

                    MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1;';") 'access database engine data connectivity component
                    MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [horaire$C3:AC27]", MyConnection)
                    dataSet = New System.Data.DataSet
                    MyCommand.Fill(dataSet)
                    DataGridView1.DataSource = dataSet.Tables(0)
                    MyConnection.Close()

                Catch ex As Exception
                    MsgBox(ex.Message.ToString)
                End Try

            End If

        End If

    End Sub

    Protected Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim param As CreateParams = MyBase.CreateParams
            param.ClassStyle = param.ClassStyle Or &H200
            Return param
        End Get
    End Property

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub txtvv_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtvv.KeyPress
        If Not numbers.Contains(e.KeyChar) And Not Asc(e.KeyChar) = 8 Then
            e.Handled = True
        End If

    End Sub

    Private Sub txtww_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtww.KeyPress
        If Not numbers.Contains(e.KeyChar) And Not Asc(e.KeyChar) = 8 Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtw1w2_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtw1w2.KeyPress
        If Not numbers.Contains(e.KeyChar) And Not Asc(e.KeyChar) = 8 Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtN_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtN.KeyPress
        If Not numbers.Contains(e.KeyChar) And Not Asc(e.KeyChar) = 8 Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtcl_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtcl.KeyPress
        If Not numbers.Contains(e.KeyChar) And Not Asc(e.KeyChar) = 8 Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtcm_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtcm.KeyPress
        If Not numbers.Contains(e.KeyChar) And Not Asc(e.KeyChar) = 8 Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtch_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtch.KeyPress
        If Not numbers.Contains(e.KeyChar) And Not Asc(e.KeyChar) = 8 Then
            e.Handled = True
        End If
    End Sub

    Private Sub txth1_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txth1.KeyPress
        If Not numbers.Contains(e.KeyChar) And Not Asc(e.KeyChar) = 8 Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtn1_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtn1.KeyPress
        If Not numberup4.Contains(e.KeyChar) And Not Asc(e.KeyChar) = 8 Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtc1_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtc1.KeyPress
        If Not numbers.Contains(e.KeyChar) And Not Asc(e.KeyChar) = 8 Then
            e.Handled = True
        End If
    End Sub

    Private Sub txth0_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txth0.KeyPress
        If Not numbers.Contains(e.KeyChar) And Not Asc(e.KeyChar) = 8 Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtn0_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtn0.KeyPress
        If Not numberdn4.Contains(e.KeyChar) And Not Asc(e.KeyChar) = 8 Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtc0_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtc0.KeyPress
        If Not numbers.Contains(e.KeyChar) And Not Asc(e.KeyChar) = 8 Then
            e.Handled = True
        End If
    End Sub

    Private Sub txth2_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txth2.KeyPress
        If Not numbers.Contains(e.KeyChar) And Not Asc(e.KeyChar) = 8 Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtn2_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtn2.KeyPress
        If Not numbers.Contains(e.KeyChar) And Not Asc(e.KeyChar) = 8 Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtc2_KeyPress(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtc2.KeyPress
        If Not numbers.Contains(e.KeyChar) And Not Asc(e.KeyChar) = 8 Then
            e.Handled = True
        End If
    End Sub

    Private Sub btnbindhourly_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnbindhourly.Click

        If DateTimePicker2.Value.Date <= Date.Now.Date Then
            Dim Name_CRQ As String = [String].Format(disck & ":\" & Folder_crq & "\" & IDstation & "crq{0}.xls", DateTimePicker2.Value.Date.ToString("ddMMyyyy"))

            If System.IO.File.Exists(Name_CRQ) Then
                Dim XlApp As Excel.Application
                Dim XlWbk As Excel.Workbook
                Dim HrlyWsh As Excel.Worksheet
                XlApp = New Excel.Application
                XlWbk = XlApp.Workbooks.Open(Name_CRQ)
                HrlyWsh = XlWbk.Worksheets("horaire")
                'code here
                Dim vector() As Integer = {1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24}
                Dim fichierminute As String = "A_" & DateTimePicker2.Value.Date.ToString("MMdd") & ".xls"
                Dim fichierhoraire As String = "S_" & DateTimePicker2.Value.Date.ToString("MMdd") & ".xls"
                Dim fichmint1 As String = "A_" & DateTimePicker2.Value.Date.AddDays(-1).ToString("MMdd") & ".xls"
                Dim fichhor1 As String = "S_" & DateTimePicker2.Value.Date.AddDays(-1).ToString("MMdd") & ".xls"
                Dim localminute As String = disck & BD_crq & fichierminute
                Dim localhoraire As String = disck & BD_crq & fichierhoraire
                Dim locmin1 As String = disck & BD_crq & fichmint1
                Dim lochor1 As String = disck & BD_crq & fichhor1

                If System.IO.File.Exists(sourceADRESS & fichierminute) Then
                    File.Copy(sourceADRESS & fichierminute, localminute, True)
                    Dim minutewbk As Excel.Workbook
                    Dim minutwsh As Excel.Worksheet

                    minutewbk = XlApp.Workbooks.Open(localminute)
                    minutwsh = minutewbk.Sheets(1)
                    Dim t As Integer

                    Dim Mnt_lastRow As Integer = minutwsh.Range("A1500").End(Excel.XlDirection.xlUp).Row
                    Dim Sum_Row_chck As Integer = (Mnt_lastRow - 1) / 60
                    If vector.Contains(Sum_Row_chck) Then
                        Dim Mnt_Step As Integer
                        For i = 1 To 23
                            t = i + 4
                            Mnt_Step = (i * 60) + 1
                            minutwsh.Cells(Mnt_Step, SmDD).Copy(HrlyWsh.Cells(t, 5))
                            minutwsh.Cells(Mnt_Step, SmFF).Copy(HrlyWsh.Cells(t, 6))
                            minutwsh.Cells(Mnt_Step, SmT).Copy(HrlyWsh.Cells(t, 22))
                            minutwsh.Cells(Mnt_Step, SmTd).Copy(HrlyWsh.Cells(t, 24))
                            minutwsh.Cells(Mnt_Step, SmHr).Copy(HrlyWsh.Cells(t, 25))
                            minutwsh.Cells(Mnt_Step, SmPs).Copy(HrlyWsh.Cells(t, 26))
                            minutwsh.Cells(Mnt_Step, SmPm).Copy(HrlyWsh.Cells(t, 27))
                        Next
                        minutewbk.Close(SaveChanges:=False)
                        releaseObject(minutewbk)
                        releaseObject(minutewbk)
                    Else
                        MsgBox("Les données minutes ne sont pas complètes sur CAOBS, la recupérétion sera interrompue, UTILISEZ LA RECUPERATION HORAIRE CI-DESSUS! ", vbOKOnly, MsgBoxStyle.Exclamation)
                        minutewbk.Close(SaveChanges:=False)
                        releaseObject(minutwsh)
                        releaseObject(minutewbk)
                        XlWbk.Close(SaveChanges:=True)
                        releaseObject(HrlyWsh)
                        releaseObject(XlWbk)
                        releaseObject(XlApp)
                        Try
                            Dim MyConnection As System.Data.OleDb.OleDbConnection
                            Dim dataSet As System.Data.DataSet
                            Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
                            Dim path As String = [String].Format(disck & ":\" & Folder_crq & "\" & IDstation & "crq{0}.xls", DateTimePicker2.Value.Date.ToString("ddMMyyyy"))

                            MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1;';") 'access database engine data connectivity component
                            MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [horaire$C3:AC27]", MyConnection)
                            dataSet = New System.Data.DataSet
                            MyCommand.Fill(dataSet)
                            DataGridView1.DataSource = dataSet.Tables(0)
                            MyConnection.Close()

                        Catch ex As Exception
                            MsgBox(ex.Message.ToString)
                        End Try

                        Exit Sub

                    End If
                End If
                If System.IO.File.Exists(sourceADRESS & fichierhoraire) Then
                    File.Copy(sourceADRESS & fichierhoraire, localhoraire, True)
                    Dim horairewb As Excel.Workbook
                    Dim horairewsh As Excel.Worksheet
                    Dim r As Integer
                    horairewb = XlApp.Workbooks.Open(localhoraire)
                    horairewsh = horairewb.Sheets(1)
                    Dim horlastRow As Integer = horairewsh.Range("A30").End(Excel.XlDirection.xlUp).Row
                    Dim sum_Hor_row_Check As Integer = horlastRow - 1

                    If vector.Contains(sum_Hor_row_Check) Then
                        Dim Hr_Step As Integer
                        For i = 1 To 23
                            r = i + 4
                            Hr_Step = i + 1
                            horairewsh.Cells(Hr_Step, ShEw).Copy(HrlyWsh.Cells(r, 23))
                            horairewsh.Cells(Hr_Step, ShRR).Copy(HrlyWsh.Cells(r, 28))
                        Next
                        horairewb.Close(SaveChanges:=False)
                        releaseObject(horairewsh)
                        releaseObject(horairewb)
                    Else
                        MsgBox("Les données HORAIRES sur CAOBS ne sont pas complètes, la recupérétion sera interrompue, UTILISEZ LA RECUPERATION HORAIRE CI-DESSUS!", vbOKOnly, MsgBoxStyle.Exclamation)
                        horairewb.Close(SaveChanges:=False)
                        releaseObject(horairewsh)
                        releaseObject(horairewb)
                        XlWbk.Close(SaveChanges:=True)
                        releaseObject(HrlyWsh)
                        releaseObject(XlWbk)
                        releaseObject(XlApp)

                        Try
                            Dim MyConnection As System.Data.OleDb.OleDbConnection
                            Dim dataSet As System.Data.DataSet
                            Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
                            Dim path As String = [String].Format(disck & ":\" & Folder_crq & "\" & IDstation & "crq{0}.xls", DateTimePicker2.Value.Date.ToString("ddMMyyyy"))

                            MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1;';") 'access database engine data connectivity component
                            MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [horaire$C3:AC27]", MyConnection)
                            dataSet = New System.Data.DataSet
                            MyCommand.Fill(dataSet)
                            DataGridView1.DataSource = dataSet.Tables(0)
                            MyConnection.Close()

                        Catch ex As Exception
                            MsgBox(ex.Message.ToString)
                        End Try
                        Exit Sub

                    End If

                End If
                If System.IO.File.Exists(sourceADRESS & fichmint1) Then
                    File.Copy(sourceADRESS & fichmint1, locmin1, True)
                    Dim minutewbk_1 As Excel.Workbook
                    Dim minutwsh_1 As Excel.Worksheet
                    minutewbk_1 = XlApp.Workbooks.Open(locmin1)
                    minutwsh_1 = minutewbk_1.Sheets(1)
                    Dim Mnt_lastrow_1 As Integer = minutwsh_1.Range("A1500").End(Excel.XlDirection.xlUp).Row
                    Dim row_check = CType(minutwsh_1.Cells(Mnt_lastrow_1, 1), Excel.Range).Value
                    Dim st_row_chk As String = row_check.ToString

                    If st_row_chk = DateTimePicker2.Value.Date.ToString("dd/MM/yyyy") & " 00:00:00" OrElse st_row_chk = DateTimePicker2.Value.Date.ToString("MM/dd/yyyy") & " 00:00:00" Then
                        minutwsh_1.Cells(Mnt_lastrow_1, SmDD).Copy(HrlyWsh.Cells(4, 5))
                        minutwsh_1.Cells(Mnt_lastrow_1, SmFF).Copy(HrlyWsh.Cells(4, 6))
                        minutwsh_1.Cells(Mnt_lastrow_1, SmT).Copy(HrlyWsh.Cells(4, 22))
                        minutwsh_1.Cells(Mnt_lastrow_1, SmTd).Copy(HrlyWsh.Cells(4, 24))
                        minutwsh_1.Cells(Mnt_lastrow_1, SmHr).Copy(HrlyWsh.Cells(4, 25))
                        minutwsh_1.Cells(Mnt_lastrow_1, SmPs).Copy(HrlyWsh.Cells(4, 26))
                        minutwsh_1.Cells(Mnt_lastrow_1, SmPm).Copy(HrlyWsh.Cells(4, 27))

                        minutewbk_1.Close(SaveChanges:=False)
                        releaseObject(minutwsh_1)
                        releaseObject(minutewbk_1)
                    Else
                        HrlyWsh.Cells(4, 5) = "///"
                        HrlyWsh.Cells(4, 6) = "//./"
                        HrlyWsh.Cells(4, 22) = "//./"
                        HrlyWsh.Cells(4, 24) = "//./"
                        HrlyWsh.Cells(4, 25) = "//./"
                        HrlyWsh.Cells(4, 26) = "////./"
                        HrlyWsh.Cells(4, 27) = "////./"
                        minutewbk_1.Close(SaveChanges:=False)
                        releaseObject(minutwsh_1)
                        releaseObject(minutewbk_1)

                    End If
                End If
                If System.IO.File.Exists(sourceADRESS & fichhor1) Then
                    File.Copy(sourceADRESS & fichhor1, lochor1, True)
                    Dim horairewb_1 As Excel.Workbook
                    Dim horairewsh_1 As Excel.Worksheet
                    horairewb_1 = XlApp.Workbooks.Open(lochor1)
                    horairewsh_1 = horairewb_1.Sheets(1)
                    Dim horlastRow1 As Integer = horairewsh_1.Range("A30").End(Excel.XlDirection.xlUp).Row
                    Dim hor_row_chk = CType(horairewsh_1.Cells(horlastRow1, 1), Excel.Range).Value
                    Dim st_hor_row_chk As String = hor_row_chk.ToString
                    If st_hor_row_chk = DateTimePicker2.Value.Date.ToString("dd/MM/yyyy") & " 00:00:00" OrElse st_hor_row_chk = DateTimePicker2.Value.Date.ToString("MM/dd/yyyy") & " 00:00:00" Then
                        horairewsh_1.Cells(horlastRow1, ShEw).Copy(HrlyWsh.Cells(4, 23))
                        horairewsh_1.Cells(horlastRow1, ShRR).Copy(HrlyWsh.Cells(4, 28))
                        horairewb_1.Close(SaveChanges:=False)
                        releaseObject(horairewsh_1)
                        releaseObject(horairewb_1)
                    Else
                        HrlyWsh.Cells(4, 23) = "//./"
                        HrlyWsh.Cells(4, 28) = "//./"
                        horairewb_1.Close(SaveChanges:=False)
                        releaseObject(horairewsh_1)
                        releaseObject(horairewb_1)

                    End If
                End If
                XlWbk.Close(SaveChanges:=True)
                releaseObject(HrlyWsh)
                releaseObject(XlWbk)
                releaseObject(XlApp)

                Try
                    Dim MyConnection As System.Data.OleDb.OleDbConnection
                    Dim dataSet As System.Data.DataSet
                    Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
                    Dim path As String = Name_CRQ
                    MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1;';") 'access database engine data connectivity component
                    MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [horaire$C3:AC27]", MyConnection)
                    dataSet = New System.Data.DataSet
                    MyCommand.Fill(dataSet)
                    DataGridView1.DataSource = dataSet.Tables(0)
                    MyConnection.Close()

                Catch ex As Exception
                    MsgBox(ex.Message.ToString)
                End Try

            Else

                Dim msg As String = DateTimePicker2.Value.Date.ToString("dd/MM/yyyy")
                MsgBox("Le CRQ du : " & msg & " n'existe pas créez le d'abord", vbOKOnly, MsgBoxStyle.Exclamation)

            End If

        Else
            MsgBox("La date selectionnée appartient au futur, selectionnez une date convenable!", vbOKOnly, MsgBoxStyle.Exclamation)
        End If


    End Sub

    Private Sub btnvalidtdh_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnvalidtdh.Click

        If DateTimePicker1.Value.Date > Date.Now.Date Then
            MsgBox("La date selectionnée appartient au futur, selectionnez une date convenable!", vbOKOnly, MsgBoxStyle.Exclamation)
        ElseIf DateTimePicker1.Value.Date <= Date.Now.Date Then
            If cbohour.SelectedItem <> "" Then
                Dim File_Date As String = DateTimePicker1.Value.Date.ToString("ddMMyyyy")
                Dim Name_CRQ As String = String.Format(disck & ":\" & Folder_crq & "\" & IDstation & "crq{0}.xls", File_Date)
                If System.IO.File.Exists(Name_CRQ) Then
                    Dim XlApp As Excel.Application
                    Dim XlWbk As Excel.Workbook
                    Dim HrlyWsh As Excel.Worksheet
                    XlApp = New Excel.Application
                    XlWbk = XlApp.Workbooks.Open(Name_CRQ)
                    HrlyWsh = XlWbk.Worksheets("horaire")

                    Dim Hr As Integer = cbohour.SelectedItem + 4

                    HrlyWsh.Cells(Hr, 4).Value = txtvv.Text
                    HrlyWsh.Cells(Hr, 7).Value = txtww.Text
                    HrlyWsh.Cells(Hr, 8).Value = txtw1w2.Text
                    HrlyWsh.Cells(Hr, 9).Value = txtcl.Text
                    HrlyWsh.Cells(Hr, 10).Value = txtcm.Text
                    HrlyWsh.Cells(Hr, 11).Value = txtch.Text
                    HrlyWsh.Cells(Hr, 12).Value = txtN.Text
                    HrlyWsh.Cells(Hr, 13).Value = txth1.Text
                    HrlyWsh.Cells(Hr, 14).Value = txtn1.Text
                    HrlyWsh.Cells(Hr, 15).Value = txtc1.Text
                    HrlyWsh.Cells(Hr, 16).Value = txth0.Text
                    HrlyWsh.Cells(Hr, 17).Value = txtn0.Text
                    HrlyWsh.Cells(Hr, 18).Value = txtc0.Text
                    HrlyWsh.Cells(Hr, 19).Value = txth2.Text
                    HrlyWsh.Cells(Hr, 20).Value = txtn2.Text
                    HrlyWsh.Cells(Hr, 21).Value = txtc2.Text

                    XlWbk.Close(SaveChanges:=True)
                    releaseObject(HrlyWsh)
                    releaseObject(XlWbk)
                    releaseObject(XlApp)
                    MsgBox("La validation du Tour d'horizon de " & cbohour.SelectedItem & "Heure est éffectuée", MsgBoxStyle.Information, vbOKOnly)

                    Try
                        Dim MyConnection As System.Data.OleDb.OleDbConnection
                        Dim dataSet As System.Data.DataSet
                        Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
                        Dim path As String = Name_CRQ
                        MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1;';")
                        MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [horaire$C3:AC27]", MyConnection)
                        dataSet = New System.Data.DataSet
                        MyCommand.Fill(dataSet)
                        DataGridView1.DataSource = dataSet.Tables(0)
                        MyConnection.Close()
                    Catch ex As Exception
                        MsgBox(ex.Message.ToString)
                    End Try

                Else
                    MsgBox("Le CRQ du : " & File_Date & " n'existe pas, créez le à partir des Outils de Création ci-dessous!", vbOKOnly, MsgBoxStyle.Exclamation)

                End If

            Else
                MsgBox("Selectionnez une heure convenante à la validation du Tour d'Horizon.", vbOKOnly, MsgBoxStyle.Exclamation)

            End If

        End If


    End Sub

    Private Sub btnrecuptelem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnrecuptelem.Click

        If DateTimePicker1.Value.Date > Date.Now.Date Then

            MsgBox("Le CRQ du : " & DateTimePicker1.Value.Date & " appartient au futur ,cet outil vous permet uniquement de récupérer des données manquantes !", MsgBoxStyle.Exclamation, vbOKOnly)

        ElseIf DateTimePicker1.Value.Date <= Date.Now.Date Then

            If cbohour.Text <> "" Then
                    Dim CRQ_date As String
                    Dim Srs_date As String
                    Dim Qry_date As String

                If cbohour.SelectedItem <> 0 Then
                    CRQ_date = DateTimePicker1.Value.Date.ToString("ddMMyyyy")
                    Srs_date = DateTimePicker1.Value.Date.ToString("MMdd")
                    Qry_date = DateTimePicker1.Value.Date.ToString("dd/MM/yyyy")
                    recup_Query(CRQ_date, Srs_date, Qry_date)

                ElseIf cbohour.SelectedItem = 0 Then
                    CRQ_date = DateTimePicker1.Value.Date.ToString("ddMMyyyy")
                    Srs_date = DateTimePicker1.Value.Date.AddDays(-1).ToString("MMdd")
                    Qry_date = DateTimePicker1.Value.Date.ToString("dd/MM/yyyy")
                    recup_Query(CRQ_date, Srs_date, Qry_date)

                End If

                    Dim CRQ_Path As String = [String].Format(disck & ":\" & Folder_crq & "\" & IDstation & "crq{0}.xls", DateTimePicker1.Value.Date.ToString("ddMMyyyy"))
                    Show_on_DGV(CRQ_Path)

                ElseIf cbohour.Text = "" Then
                MsgBox("Selectionner une Heure pour la validation des télémesures !", MsgBoxStyle.Exclamation, vbOKOnly)
            End If

            End If

    End Sub

    Private Sub btnbindextr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnbindextr.Click
        If DateTimePicker2.Value.Date > Date.Now.Date Then

            MsgBox("Le CRQ du : " & DateTimePicker2.Value.Date & " appartient au futur ,cet outil vous permet uniquement de récupérer des données manquantes !", MsgBoxStyle.Exclamation, vbOKOnly)

        ElseIf DateTimePicker2.Value.Date = Date.Now.Date Then

            MsgBox("Le CRQ du : " & DateTimePicker2.Value.Date & " est celui d'aujourd'huit , utilizez le menu pricipal pour charger les Extêmes!", MsgBoxStyle.Exclamation, vbOKOnly)

        ElseIf DateTimePicker2.Value.Date < Date.Now.Date Then

            Dim ystrdycrq As Excel.Workbook
            Dim ystrdextr As Excel.Worksheet
            Dim ystrdname As String = [String].Format(disck & ":\" & Folder_crq & "\" & IDstation & "crq{0}.xls", DateTimePicker2.Value.Date.ToString("ddMMyyyy"))
            Dim ystrdStmp As String = DateTimePicker2.Value.Date.ToString("ddMMyyyy")
            Dim fichiextreme As String = "K_" & DateTimePicker2.Value.Date.ToString("MMdd") & ".xls"
            Dim sourcextreme As String = sourceADRESS & fichiextreme
            Dim localextreme As String = disck & BD_crq & fichiextreme

            Dim fichihor As String = "S_" & DateTimePicker2.Value.Date.ToString("MMdd") & ".xls"
            Dim sourhor As String = sourceADRESS & fichihor
            Dim localHor As String = disck & BD_crq & fichihor
            Dim sourceHor As String = sourceADRESS & fichihor

            If System.IO.File.Exists(ystrdname) Then

                If System.IO.File.Exists(sourcextreme) Then
                    File.Copy(sourcextreme, localextreme, True)
                    Dim app As Excel.Application
                    Dim extrwb As Excel.Workbook
                    Dim extrwsh As Excel.Worksheet
                    app = New Excel.Application
                    ystrdycrq = app.Workbooks.Open(ystrdname)
                    ystrdextr = ystrdycrq.Worksheets("extrèmes")
                    extrwb = app.Workbooks.Open(localextreme)
                    extrwsh = extrwb.Sheets(1)

                    Dim st_KhTn, st_KhTx, st_KhUn, st_KhUx As String

                    If IsNumeric(CType(extrwsh.Cells(2, KhTn), Excel.Range).Value2) Then
                        st_KhTn = (New DateTime()).AddDays(CType(extrwsh.Cells(2, KhTn), Excel.Range).Value2).ToString("HH:mm")
                    Else
                        st_KhTn = "//://"

                    End If
                    If IsNumeric(CType(extrwsh.Cells(2, KhTx), Excel.Range).Value2) Then
                        st_KhTx = (New DateTime()).AddDays(CType(extrwsh.Cells(2, KhTx), Excel.Range).Value2).ToString("HH:mm")
                    Else
                        st_KhTx = "//://"
                    End If
                    If IsNumeric(CType(extrwsh.Cells(2, KhUn), Excel.Range).Value2) Then
                        st_KhUn = (New DateTime()).AddDays(CType(extrwsh.Cells(2, KhUn), Excel.Range).Value2).ToString("HH:mm")
                    Else
                        st_KhUn = "//://"
                    End If
                    If IsNumeric(CType(extrwsh.Cells(2, KhUx), Excel.Range).Value2) Then
                        st_KhUx = (New DateTime()).AddDays(CType(extrwsh.Cells(2, KhUx), Excel.Range).Value2).ToString("HH:mm")
                    Else
                        st_KhUx = "//://"
                    End If

                    ystrdextr.Cells(2, 1).value = IDstation
                    ystrdextr.Cells(2, 2).value = ystrdStmp
                    ystrdextr.Cells(1, 1).value = "Indicatif"
                    ystrdextr.Cells(1, 2).value = "Date"
                    extrwsh.Cells(1, KTn).Copy(ystrdextr.Cells(1, 3))
                    extrwsh.Cells(2, KTn).Copy(ystrdextr.Cells(2, 3))
                    extrwsh.Cells(1, KhTn).Copy(ystrdextr.Cells(1, 4))
                    ystrdextr.Cells(2, 4).value = st_KhTn
                    extrwsh.Cells(1, KTx).Copy(ystrdextr.Cells(1, 5))
                    extrwsh.Cells(2, KTx).Copy(ystrdextr.Cells(2, 5))
                    extrwsh.Cells(1, KhTx).Copy(ystrdextr.Cells(1, 6))
                    ystrdextr.Cells(2, 6).value = st_KhTx
                    extrwsh.Cells(1, KUn).Copy(ystrdextr.Cells(1, 7))
                    extrwsh.Cells(2, KUn).Copy(ystrdextr.Cells(2, 7))
                    extrwsh.Cells(1, KhUn).Copy(ystrdextr.Cells(1, 8))
                    ystrdextr.Cells(2, 8).value = st_KhUn
                    extrwsh.Cells(1, KUx).Copy(ystrdextr.Cells(1, 9))
                    extrwsh.Cells(2, KUx).Copy(ystrdextr.Cells(2, 9))
                    extrwsh.Cells(1, KhUx).Copy(ystrdextr.Cells(1, 10))
                    ystrdextr.Cells(2, 10).value = st_KhUx
                    extrwsh.Cells(1, KRR).Copy(ystrdextr.Cells(1, 11))
                    extrwsh.Cells(2, KRR).Copy(ystrdextr.Cells(2, 11))
                    extrwsh.Cells(1, KIs).Copy(ystrdextr.Cells(1, 12))
                    extrwsh.Cells(2, KIs).Copy(ystrdextr.Cells(2, 12))
                    extrwsh.Cells(1, KRg).Copy(ystrdextr.Cells(1, 14))
                    extrwsh.Cells(2, KRg).Copy(ystrdextr.Cells(2, 14))
                    extrwsh.Cells(1, KTxSol).Copy(ystrdextr.Cells(1, 22))
                    extrwsh.Cells(2, KTxSol).Copy(ystrdextr.Cells(2, 22))
                    extrwsh.Cells(1, KTnSol).Copy(ystrdextr.Cells(1, 23))
                    extrwsh.Cells(2, KTnSol).Copy(ystrdextr.Cells(2, 23))
                    ystrdextr.Cells(1, 13).value = "Insol 1/10 heure"
                    ystrdextr.Columns("M:M").ColumnWidth = 10
                    ystrdextr.Cells(2, 13).value = Math.Round((ystrdextr.Cells(2, 12).value) / 6)
                    ystrdextr.Cells(1, 15).value = "DD vent max Inst."
                    ystrdextr.Cells(1, 16).value = "FF vent max Inst."
                    ystrdextr.Cells(1, 17).value = "Heure du max"
                    ystrdextr.Cells(1, 18).value = "DD vent max 10'"
                    ystrdextr.Cells(1, 19).value = "FF vent max 10'"
                    ystrdextr.Cells(1, 20).value = "Heure du max"
                    ystrdextr.Cells(1, 21).value = "Neige"

                    If System.IO.File.Exists(sourceHor) Then
                        File.Copy(sourceHor, localHor, True)
                        Dim HorWBK As Excel.Workbook
                        Dim HorWsh As Excel.Worksheet
                        HorWBK = app.Workbooks.Open(localHor)
                        HorWsh = HorWBK.Sheets(1)
                        Dim VntMoyTAB(2, 23) As Double
                        Dim VntInsTAB(2, 23) As Double
                        Dim MaxFFmoy As Double
                        Dim DDmaxmoy As Double
                        Dim minutMaxmoy As Double
                        Dim heuremoy As Integer

                        Dim MaxFFins As Double
                        Dim DDmaxins As Double
                        Dim minutMaxins As Double
                        Dim heureins As Integer

                        For i = 0 To 23
                            If HorWsh.Cells(i + 2, 3).text <> "//./" Then
                                VntMoyTAB(1, i) = HorWsh.Cells(i + 2, 3).value
                            End If
                        Next

                        For g = 0 To 23
                            If HorWsh.Cells(g + 2, 6).text <> "//./" Then
                                VntInsTAB(1, g) = HorWsh.Cells(g + 2, 6).value
                            End If
                        Next

                        For Each element1 As Double In VntMoyTAB
                            MaxFFmoy = Math.Max(MaxFFmoy, element1)
                        Next

                        For Each element2 As Double In VntInsTAB
                            MaxFFins = Math.Max(MaxFFins, element2)
                        Next

                        For i = 0 To 23
                            If HorWsh.Cells(i + 2, 2).text <> "///" Then
                                VntMoyTAB(0, i) = HorWsh.Cells(i + 2, 2).value
                            End If
                            If HorWsh.Cells(i + 2, 4).text <> "//" Then
                                VntMoyTAB(2, i) = HorWsh.Cells(i + 2, 4).value
                            End If
                        Next


                        For j = 0 To 23
                            If HorWsh.Cells(j + 2, 5).text <> "///" Then
                                VntInsTAB(0, j) = HorWsh.Cells(j + 2, 5).value
                            End If
                            If HorWsh.Cells(j + 2, 7).text <> "//" Then
                                VntInsTAB(2, j) = HorWsh.Cells(j + 2, 7).value
                            End If
                        Next


                        For km = 0 To 23
                            If VntMoyTAB(1, km) = MaxFFmoy Then
                                DDmaxmoy = HorWsh.Cells(km + 2, 2).value
                                minutMaxmoy = HorWsh.Cells(km + 2, 4).value
                                heuremoy = km
                            End If
                        Next km

                        For ki = 0 To 23
                            If VntInsTAB(1, ki) = MaxFFins Then
                                DDmaxins = HorWsh.Cells(ki + 2, 5).value
                                minutMaxins = HorWsh.Cells(ki + 2, 7).value
                                heureins = ki
                            End If
                        Next ki

                        Dim mm_moy As String = CInt(minutMaxmoy).ToString("D2")
                        Dim hh_moy As String = heuremoy.ToString("D2")
                        Dim mm_ins As String = CInt(minutMaxins).ToString("D2")
                        Dim hh_ins As String = heureins.ToString("D2")
                        Dim HHmm_moy As String = hh_moy & ":" & mm_moy
                        Dim HHmm_ins As String = hh_ins & ":" & mm_ins

                        ystrdextr.Cells(2, 18).value = DDmaxmoy
                        ystrdextr.Cells(2, 19).value = MaxFFmoy
                        ystrdextr.Cells(2, 20).value = HHmm_moy

                        ystrdextr.Cells(2, 15).value = DDmaxins
                        ystrdextr.Cells(2, 16).value = MaxFFins
                        ystrdextr.Cells(2, 17).value = HHmm_ins

                        Array.Clear(VntMoyTAB, 0, VntMoyTAB.Length)
                        Array.Clear(VntInsTAB, 0, VntInsTAB.Length)
                        HorWBK.Close(SaveChanges:=False)
                        releaseObject(HorWsh)
                        releaseObject(HorWBK)

                    End If

                    extrwb.Close(SaveChanges:=False)
                    ystrdycrq.Close(SaveChanges:=True)
                    releaseObject(ystrdextr)
                    releaseObject(ystrdycrq)
                    releaseObject(extrwsh)
                    releaseObject(extrwb)

                    Try
                        Dim MyConnection As System.Data.OleDb.OleDbConnection
                        Dim dataSet As System.Data.DataSet
                        Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
                        Dim path As String = [String].Format(disck & ":\" & Folder_crq & "\" & IDstation & "crq{0}.xls", DateTimePicker2.Value.Date.ToString("ddMMyyyy"))

                        MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1;';") 'access database engine data connectivity component
                        MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [extrèmes$A1:W2]", MyConnection)
                        dataSet = New System.Data.DataSet
                        MyCommand.Fill(dataSet)
                        DataGridView1.DataSource = dataSet.Tables(0)
                        MyConnection.Close()
                    Catch ex As Exception
                        MsgBox(ex.Message.ToString)
                    End Try

                    MsgBox("Les extrêmes du :" & DateTimePicker2.Value.Date & " sont chargés avec succés.", MsgBoxStyle.Information, vbOKOnly)

                Else
                    MsgBox("Le fichier extrêmes n'existe pas verifiez le chemain d'accés !", MsgBoxStyle.Exclamation, vbOKOnly)

                End If
            Else
                MsgBox("Le CRQ du : " & ystrdname & " n'existe pas, créez le d'abord !", MsgBoxStyle.Critical, vbOKOnly)

            End If

        End If

    End Sub

    Private Sub rdtoday_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) 
        DateTimePicker1.Enabled = False
    End Sub

    Private Sub rdbydate_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) 
        DateTimePicker1.Enabled = True
    End Sub

    Private Sub Show_on_DGV(File_path As String)
        Try
            Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim dataSet As System.Data.DataSet
            Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
            Dim path As String = File_path

            MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1;';") 'access database engine data connectivity component
            MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [horaire$C3:AC27]", MyConnection)
            dataSet = New System.Data.DataSet
            MyCommand.Fill(dataSet)
            DataGridView1.DataSource = dataSet.Tables(0)
            MyConnection.Close()
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try

    End Sub

    Private Sub recup_Query(CRQ_Dt As String, Srs_Dt As String, Qry_Dt As String)
        Dim xlapp As Excel.Application

        Dim CRQ As Excel.Workbook
        Dim CrqhrWsh As Excel.Worksheet

        Dim CopyMntWbk As Excel.Workbook
        Dim copyMntWsh As Excel.Worksheet
        Dim CopyHrWbk As Excel.Workbook
        Dim copyHrWsh As Excel.Worksheet

        Dim MintWbk As Excel.Workbook
        Dim mintWsh As Excel.Worksheet
        Dim HorWbk As Excel.Workbook
        Dim HorWsh As Excel.Worksheet

        Dim DD10 As String = String.Empty
        Dim FF10 As String = String.Empty
        Dim TAir As String = String.Empty
        Dim Td As String = String.Empty
        Dim Humidite As String = String.Empty
        Dim Pst As String = String.Empty
        Dim Pmer As String = String.Empty
        Dim Ew As String = String.Empty
        Dim Pluie As String = String.Empty



        Dim Name_CRQ As String = [String].Format(disck & ":\" & Folder_crq & "\" & IDstation & "crq{0}.xls", CRQ_Dt)
        If System.IO.File.Exists(Name_CRQ) Then

            Dim Copy_A_Name As String = disck & BD_crq & "TempoCopy\tmp_A_" & Srs_Dt & ".xls"
            Dim Copy_S_Name As String = disck & BD_crq & "TempoCopy\tmp_S_" & Srs_Dt & ".xls"
            Dim Mint_Name As String = disck & BD_crq & "A_" & Srs_Dt & ".xls"
            Dim Hor_Name As String = disck & BD_crq & "S_" & Srs_Dt & ".xls"
            Dim source_Mint_Name As String = sourceADRESS & "A_" & Srs_Dt & ".xls"
            Dim source_Hor_Name As String = sourceADRESS & "S_" & Srs_Dt & ".xls"

            Dim Dt As String = Qry_Dt
            Dim Hr As Integer = CInt(cbohour.SelectedItem)
            Dim stHr As String = Hr.ToString("D2")
            Dim st_Full_date As String = String.Empty
            If Hr = 0 Then
                st_Full_date = "BETWEEN #" & Dt & " " & stHr & ":00:00# AND #" & Dt & " " & stHr & ":01:00#"
            ElseIf Hr <> 0 Then
                st_Full_date = "= #" & Dt & " " & stHr & ":00:00#"
            End If

            If System.IO.File.Exists(source_Mint_Name) Then
                If System.IO.File.Exists(Copy_A_Name) Then
                    File.Delete(Copy_A_Name)
                End If
                File.Copy(source_Mint_Name, Mint_Name, True)
                xlapp = New Excel.Application
                CopyMntWbk = xlapp.Workbooks.Add()
                CopyMntWbk.Worksheets(1).name = "Mnt"
                copyMntWsh = CopyMntWbk.Worksheets("Mnt")
                MintWbk = xlapp.Workbooks.Open(Mint_Name)
                mintWsh = MintWbk.Worksheets(1)

                mintWsh.Range("A1").CurrentRegion.Copy(copyMntWsh.Range("A1"))
                MintWbk.Close(SaveChanges:=False)
                CopyMntWbk.SaveAs(Copy_A_Name)

                DD10 = CType(copyMntWsh.Cells(1, SmDD), Excel.Range).Value.ToString
                FF10 = CType(copyMntWsh.Cells(1, SmFF), Excel.Range).Value.ToString
                TAir = CType(copyMntWsh.Cells(1, SmT), Excel.Range).Value.ToString
                Td = CType(copyMntWsh.Cells(1, SmTd), Excel.Range).Value.ToString
                Humidite = CType(copyMntWsh.Cells(1, SmHr), Excel.Range).Value.ToString
                Pst = CType(copyMntWsh.Cells(1, SmPs), Excel.Range).Value.ToString
                Pmer = CType(copyMntWsh.Cells(1, SmPm), Excel.Range).Value.ToString

                CopyMntWbk.Close(SaveChanges:=True)
                Try
                    Dim MyConnection As System.Data.OleDb.OleDbConnection
                    Dim dataSet As System.Data.DataSet
                    Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
                    Dim path As String = Copy_A_Name
                    MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1;';") 'access database engine data connectivity component
                    MyCommand = New System.Data.OleDb.OleDbDataAdapter("select [Date/Heure], [" & DD10 & "], [" & FF10 & "], [" & TAir & "], [" & Td & "], [" & Humidite & "], [" & Pst & "], [" & Pmer & "] FROM [Mnt$] WHERE [Date/Heure] " & st_Full_date, MyConnection)
                    dataSet = New System.Data.DataSet
                    MyCommand.Fill(dataSet)
                    MyConnection.Close()

                    If dataSet.Tables(0) IsNot Nothing AndAlso dataSet.Tables(0).Rows.Count > 0 Then
                        Dim t As Integer = Hr + 4
                        CRQ = xlapp.Workbooks.Open(Name_CRQ)
                        CrqhrWsh = CRQ.Worksheets("horaire")
                        CrqhrWsh.Cells(t, 5).value = dataSet.Tables(0).Rows(0).Item(1).ToString
                        CrqhrWsh.Cells(t, 6).value = dataSet.Tables(0).Rows(0).Item(2).ToString
                        CrqhrWsh.Cells(t, 22).value = dataSet.Tables(0).Rows(0).Item(3).ToString
                        CrqhrWsh.Cells(t, 24).value = dataSet.Tables(0).Rows(0).Item(4).ToString
                        CrqhrWsh.Cells(t, 25).value = dataSet.Tables(0).Rows(0).Item(5).ToString
                        CrqhrWsh.Cells(t, 26).value = dataSet.Tables(0).Rows(0).Item(6).ToString
                        CrqhrWsh.Cells(t, 27).value = dataSet.Tables(0).Rows(0).Item(7).ToString
                        CRQ.Close(SaveChanges:=True)
                    Else
                        Dim t As Integer = Hr + 4
                        CRQ = xlapp.Workbooks.Open(Name_CRQ)
                        CrqhrWsh = CRQ.Worksheets("horaire")
                        CrqhrWsh.Cells(t, 5).value = "///./"
                        CrqhrWsh.Cells(t, 6).value = "//./"
                        CrqhrWsh.Cells(t, 22).value = "//./"
                        CrqhrWsh.Cells(t, 24).value = "//./"
                        CrqhrWsh.Cells(t, 25).value = "//./"
                        CrqhrWsh.Cells(t, 26).value = "////./"
                        CrqhrWsh.Cells(t, 27).value = "////./"
                        CRQ.Close(SaveChanges:=True)

                    End If

                Catch ex As Exception
                    MsgBox(ex.Message.ToString)
                End Try

                File.Delete(Copy_A_Name)

            Else
                MsgBox("source Minute non existatnte !", vbOKOnly)
            End If

            If System.IO.File.Exists(source_Hor_Name) Then
                If System.IO.File.Exists(Copy_S_Name) Then
                    File.Delete(Copy_S_Name)
                End If
                File.Copy(source_Hor_Name, Hor_Name, True)
                xlapp = New Excel.Application
                CopyHrWbk = xlapp.Workbooks.Add()
                CopyHrWbk.Worksheets(1).name = "Hor"
                copyHrWsh = CopyHrWbk.Worksheets("Hor")
                HorWbk = xlapp.Workbooks.Open(Hor_Name)
                HorWsh = HorWbk.Worksheets(1)

                HorWsh.Range("A1").CurrentRegion.Copy(copyHrWsh.Range("A1"))
                HorWbk.Close(SaveChanges:=False)
                CopyHrWbk.SaveAs(Copy_S_Name)
                Ew = CType(copyHrWsh.Cells(1, ShEw), Excel.Range).Value.ToString
                Pluie = CType(copyHrWsh.Cells(1, ShRR), Excel.Range).Value.ToString

                CopyHrWbk.Close(SaveChanges:=True)
                Try
                    Dim MyConnection As System.Data.OleDb.OleDbConnection
                    Dim dataSet As System.Data.DataSet
                    Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
                    Dim path As String = Copy_S_Name
                    MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1;';") 'access database engine data connectivity component
                    MyCommand = New System.Data.OleDb.OleDbDataAdapter("select [Date/Heure], [" & Ew & "], [" & Pluie & "] FROM [Hor$] WHERE [Date/Heure] " & st_Full_date, MyConnection)
                    dataSet = New System.Data.DataSet
                    MyCommand.Fill(dataSet)
                    MyConnection.Close()

                    If dataSet.Tables(0) IsNot Nothing AndAlso dataSet.Tables(0).Rows.Count > 0 Then
                        Dim t As Integer = Hr + 4
                        CRQ = xlapp.Workbooks.Open(Name_CRQ)
                        CrqhrWsh = CRQ.Worksheets("horaire")
                        CrqhrWsh.Cells(t, 23).value = dataSet.Tables(0).Rows(0).Item(1).ToString
                        CrqhrWsh.Cells(t, 28).value = dataSet.Tables(0).Rows(0).Item(2).ToString
                        CRQ.Close(SaveChanges:=True)
                    Else
                        Dim t As Integer = Hr + 4
                        CRQ = xlapp.Workbooks.Open(Name_CRQ)
                        CrqhrWsh = CRQ.Worksheets("horaire")
                        CrqhrWsh.Cells(t, 23).value = "//./"
                        CrqhrWsh.Cells(t, 28).value = "//./"
                        CRQ.Close(SaveChanges:=True)

                    End If

                Catch ex As Exception
                    MsgBox(ex.Message.ToString)
                End Try
                File.Delete(Copy_S_Name)
            Else
                MsgBox("Source Horaire non existatnte !", vbOKOnly, MsgBoxStyle.Exclamation)
            End If


        Else
            MsgBox("Le crq du: " & CRQ_Dt & " n'existe pas dans son repertoire, aller vers 'Création d'un CRQ manquant' pour le créer d'abord !", vbOKOnly, MsgBoxStyle.Critical)
        End If

    End Sub

End Class