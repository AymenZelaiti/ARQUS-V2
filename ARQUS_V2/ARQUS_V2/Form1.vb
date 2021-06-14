Imports Microsoft.Office.Interop
Imports System.Security.Permissions
Imports System.IO
Imports System.Globalization
Imports System.Data.OleDb
Imports Microsoft.SqlServer
Imports System.Net

Public Class saisietdh
    Dim Cor_Phen_Index As Integer
    Dim crq As Excel.Workbook
    Dim hrly As Excel.Worksheet
    Dim phen As Excel.Worksheet
    Dim extr As Excel.Worksheet
    Dim misValue As Object = System.Reflection.Missing.Value
    Dim DateOfcrq As String = DateTime.Now.ToString("ddMMyyyy")
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

    Dim MsRR As Integer
    Dim MsDI As Integer
    Dim MsRG As Integer

    Dim HrWnd As Integer
    Dim HrFFWnd As Integer
    Dim HrhWnd As Integer
    Dim HrMinT As Integer
    Dim HrhMinT As Integer
    Dim HrMxT As Integer
    Dim HrhMxT As Integer
    Dim HrMaxU As Integer
    Dim HrhMaxU As Integer
    Dim HrMinU As Integer
    Dim HrhMinU As Integer
    Dim HrRR As Integer
    Dim HrDi As Integer
    Dim HrRg As Integer

    Private Sub saisietdh_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Timer1.Start()

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

        MsDI = My.Settings.MsDI
        MsRR = My.Settings.MsRR
        MsRG = My.Settings.MsRG

        HrWnd = My.Settings.HrWnd
        HrFFWnd = HrWnd + 1
        HrhWnd = HrWnd + 2
        HrMinT = My.Settings.HrMinT
        HrhMinT = HrMinT + 1
        HrMxT = My.Settings.HrMxT
        HrhMxT = HrMxT + 1
        HrMaxU = My.Settings.HrMaxU
        HrhMaxU = HrMaxU + 1
        HrMinU = My.Settings.HrMinU
        HrhMinU = HrMinU + 1
        HrRR = My.Settings.HrRR
        HrDi = My.Settings.HrDi
        HrRg = My.Settings.HrRg

        If IDstation = "" Or disck = "" Then
            config = False
            MsgBox("Configurations manquantes,verifier le menu Config !", MsgBoxStyle.Critical, vbOKOnly)
        ElseIf IDstation <> "" And disck <> "" Then

            If System.IO.Directory.Exists(disck & ":\" & Folder_crq) Then
                config = True
                Dim FileDate As String = Date.Now.ToString("dd MMM yyyy", CultureInfo.CreateSpecificCulture("fr-FRA"))
                Dim nomfichier As String = [String].Format(disck & ":\" & Folder_crq & "\" & IDstation & "crq{0}.xls", DateTime.Now.ToString("ddMMyyyy"))
                Dim SourcePath As String = nomfichier
                Dim Filename As String = System.IO.Path.GetFileName(SourcePath) 'get the filename of the original file without the directory on it
                Dim permis As New FileIOPermission(FileIOPermissionAccess.AllAccess, disck & ":\" & Folder_crq)

                If System.IO.File.Exists(SourcePath) Then
                    'the file exists

                Else
                    Dim stamp_Date As String = Date.Now.ToString("dd MMM yyyy", CultureInfo.CreateSpecificCulture("fr-FRA"))
                    Dim stamp_Row As String = Date.Now.ToString("ddMMyyyy", CultureInfo.CreateSpecificCulture("fr-FRA"))
                    Dim Name_CRQ As String = String.Format(disck & ":\" & Folder_crq & "\" & IDstation & "crq{0}.xls", Date.Now.ToString("ddMMyyyy"))
                    New_Crq(stamp_Date, Name_CRQ, stamp_Row)
                    MsgBox("Votre CRQ est crée,vous pouvez saisir les données", MsgBoxStyle.Information, vbOKOnly)

                End If
            ElseIf Not System.IO.Directory.Exists(disck & ":\" & Folder_crq) Then
                MsgBox("Configuration manquante le dossier: " & disck & ":\" & Folder_crq & " n'existe pas!", MsgBoxStyle.Critical, vbOKOnly)
                config = False
            End If

        End If
        btncorrige.Hide()

        btnforcetelem.Enabled = False
        btnauto_stop.Enabled = False
        btnauto_start.Enabled = True
        lblauto.Hide()
        Timer2.Stop()

        If My.Settings.autoSend = True Then
            TimerSender.Start()
        ElseIf My.Settings.autoSend = False Then
            TimerSender.Stop()
        End If

    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick

        lbltime.Text = Date.Now

    End Sub

    Private Sub btnquit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnquit.Click

        If config = True Then
            Dim D As String
            D = MsgBox("Voulez vous vraiment quitter ARQUS ?", vbYesNo)

            If D = vbYes Then

                Application.Exit()
            Else
                Exit Sub
            End If
        ElseIf config = False Then
            Application.Exit()
        End If

    End Sub
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

    Private Sub btnefface_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnefface.Click

        txtvv.Text = ""
        txtww.Text = ""
        txtw1w2.Text = ""
        txtcm.Text = ""
        txtcl.Text = ""
        txtch.Text = ""
        txtN.Text = ""
        txth1.Text = ""
        txtn1.Text = ""
        txtc1.Text = ""
        txth0.Text = ""
        txtn0.Text = ""
        txtc0.Text = ""
        txth2.Text = ""
        txtn2.Text = ""
        txtc2.Text = ""
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

    Private Sub btnvalid_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnvalid.Click

        Dim a As Integer
        Dim currentdate As String = DateTime.Now.ToString("ddMMyyyy")
        Dim Name_File As String = String.Format(disck & ":\" & Folder_crq & "\" & IDstation & "crq{0}.xls", Date.Now.ToString("ddMMyyyy"))
        Dim Date_CRQ As String = Date.Now.ToString("dd MMM yyyy", CultureInfo.CreateSpecificCulture("fr-FRA"))
        Dim Date_Row As String = Date.Now.ToString("ddMMyyyy", CultureInfo.CreateSpecificCulture("fr-FRA"))

        If (Date.Now.TimeOfDay > New TimeSpan(0, 40, 0)) And (Date.Now.TimeOfDay < New TimeSpan(1, 10, 0)) Then  'la validation de 00h est manquante
            a = 5
            Valid_TdH(Name_File, Date_CRQ, Date_Row, a)

        ElseIf (Date.Now.TimeOfDay > New TimeSpan(1, 40, 0)) And (Date.Now.TimeOfDay < New TimeSpan(2, 10, 0)) Then
            a = 6
            Valid_TdH(Name_File, Date_CRQ, Date_Row, a)
        ElseIf (Date.Now.TimeOfDay > New TimeSpan(2, 40, 0)) And (Date.Now.TimeOfDay < New TimeSpan(3, 10, 0)) Then
            a = 7
            Valid_TdH(Name_File, Date_CRQ, Date_Row, a)
        ElseIf (Date.Now.TimeOfDay > New TimeSpan(3, 40, 0)) And (Date.Now.TimeOfDay < New TimeSpan(4, 10, 0)) Then
            a = 8
            Valid_TdH(Name_File, Date_CRQ, Date_Row, a)
        ElseIf (Date.Now.TimeOfDay > New TimeSpan(4, 40, 0)) And (Date.Now.TimeOfDay < New TimeSpan(5, 10, 0)) Then
            a = 9
            Valid_TdH(Name_File, Date_CRQ, Date_Row, a)
        ElseIf (Date.Now.TimeOfDay > New TimeSpan(5, 40, 0)) And (Date.Now.TimeOfDay < New TimeSpan(6, 10, 0)) Then
            a = 10
            Valid_TdH(Name_File, Date_CRQ, Date_Row, a)
        ElseIf (Date.Now.TimeOfDay > New TimeSpan(6, 40, 0)) And (Date.Now.TimeOfDay < New TimeSpan(7, 10, 0)) Then
            a = 11
            Valid_TdH(Name_File, Date_CRQ, Date_Row, a)
        ElseIf (Date.Now.TimeOfDay > New TimeSpan(7, 40, 0)) And (Date.Now.TimeOfDay < New TimeSpan(8, 10, 0)) Then
            a = 12
            Valid_TdH(Name_File, Date_CRQ, Date_Row, a)
        ElseIf (Date.Now.TimeOfDay > New TimeSpan(8, 40, 0)) And (Date.Now.TimeOfDay < New TimeSpan(9, 10, 0)) Then
            a = 13
            Valid_TdH(Name_File, Date_CRQ, Date_Row, a)
        ElseIf (Date.Now.TimeOfDay > New TimeSpan(9, 40, 0)) And (Date.Now.TimeOfDay < New TimeSpan(10, 10, 0)) Then
            a = 14
            Valid_TdH(Name_File, Date_CRQ, Date_Row, a)
        ElseIf (Date.Now.TimeOfDay > New TimeSpan(10, 40, 0)) And (Date.Now.TimeOfDay < New TimeSpan(11, 10, 0)) Then
            a = 15
            Valid_TdH(Name_File, Date_CRQ, Date_Row, a)
        ElseIf (Date.Now.TimeOfDay > New TimeSpan(11, 40, 0)) And (Date.Now.TimeOfDay < New TimeSpan(12, 10, 0)) Then
            a = 16
            Valid_TdH(Name_File, Date_CRQ, Date_Row, a)
        ElseIf (Date.Now.TimeOfDay > New TimeSpan(12, 40, 0)) And (Date.Now.TimeOfDay < New TimeSpan(13, 10, 0)) Then
            a = 17
            Valid_TdH(Name_File, Date_CRQ, Date_Row, a)
        ElseIf (Date.Now.TimeOfDay > New TimeSpan(13, 40, 0)) And (Date.Now.TimeOfDay < New TimeSpan(14, 10, 0)) Then
            a = 18
            Valid_TdH(Name_File, Date_CRQ, Date_Row, a)
        ElseIf (Date.Now.TimeOfDay > New TimeSpan(14, 40, 0)) And (Date.Now.TimeOfDay < New TimeSpan(15, 10, 0)) Then
            a = 19
            Valid_TdH(Name_File, Date_CRQ, Date_Row, a)
        ElseIf (Date.Now.TimeOfDay > New TimeSpan(15, 40, 0)) And (Date.Now.TimeOfDay < New TimeSpan(16, 10, 0)) Then
            a = 20
            Valid_TdH(Name_File, Date_CRQ, Date_Row, a)
        ElseIf (Date.Now.TimeOfDay > New TimeSpan(16, 40, 0)) And (Date.Now.TimeOfDay < New TimeSpan(17, 10, 0)) Then
            a = 21
            Valid_TdH(Name_File, Date_CRQ, Date_Row, a)
        ElseIf (Date.Now.TimeOfDay > New TimeSpan(17, 40, 0)) And (Date.Now.TimeOfDay < New TimeSpan(18, 10, 0)) Then
            a = 22
            Valid_TdH(Name_File, Date_CRQ, Date_Row, a)
        ElseIf (Date.Now.TimeOfDay > New TimeSpan(18, 40, 0)) And (Date.Now.TimeOfDay < New TimeSpan(19, 10, 0)) Then
            a = 23
            Valid_TdH(Name_File, Date_CRQ, Date_Row, a)
        ElseIf (Date.Now.TimeOfDay > New TimeSpan(19, 40, 0)) And (Date.Now.TimeOfDay < New TimeSpan(20, 10, 0)) Then
            a = 24
            Valid_TdH(Name_File, Date_CRQ, Date_Row, a)
        ElseIf (Date.Now.TimeOfDay > New TimeSpan(20, 40, 0)) And (Date.Now.TimeOfDay < New TimeSpan(21, 10, 0)) Then
            a = 25
            Valid_TdH(Name_File, Date_CRQ, Date_Row, a)
        ElseIf (Date.Now.TimeOfDay > New TimeSpan(21, 40, 0)) And (Date.Now.TimeOfDay < New TimeSpan(22, 10, 0)) Then
            a = 26
            Valid_TdH(Name_File, Date_CRQ, Date_Row, a)
        ElseIf (Date.Now.TimeOfDay > New TimeSpan(22, 40, 0)) And (Date.Now.TimeOfDay < New TimeSpan(23, 10, 0)) Then
            a = 27
            Valid_TdH(Name_File, Date_CRQ, Date_Row, a)
        ElseIf (Date.Now.TimeOfDay > New TimeSpan(23, 40, 0)) And (Date.Now.TimeOfDay < New TimeSpan(23, 59, 59)) Then  'to redefine proprely with j+1
            a = 4
            Valid_TdH(Name_File, Date_CRQ, Date_Row, a)

        ElseIf (Date.Now.TimeOfDay > New TimeSpan(0, 0, 0)) And (Date.Now.TimeOfDay < New TimeSpan(0, 10, 0)) Then
            a = 4
            Valid_TdH(Name_File, Date_CRQ, Date_Row, a)
        ElseIf Not a >= 4 And a <= 27 Then
            MsgBox("Vous êtes en dehors de l'heure de validation, validation non éffectuée", MsgBoxStyle.Exclamation, vbOKOnly)

        End If

    End Sub
    Private Sub New_Crq(Date_CRQ As String, File_Name As String, Row_Date As String)

        Dim XlApp As Excel.Application
        Dim XlWbk As Excel.Workbook
        Dim HrlyWsh As Excel.Worksheet
        Dim PhenWsh As Excel.Worksheet
        Dim ExtrWsh As Excel.Worksheet
        XlApp = New Excel.Application
        XlWbk = XlApp.Workbooks.Add
        HrlyWsh = XlWbk.Worksheets(1)
        PhenWsh = XlWbk.Worksheets(2)
        ExtrWsh = XlWbk.Worksheets(3)

        With XlWbk
            HrlyWsh.Name = "horaire"
            PhenWsh.Name = "phenomenes"
            ExtrWsh.Name = "extrèmes"

        End With

        '*********feille 1
        With HrlyWsh
            .Cells(3, 1).value = "Indicatif"
            .Cells(3, 1).font.bold = True
            .Cells(3, 1).font.size = 11
            .Cells(3, 2).value = "Date"
            .Cells(3, 2).font.bold = True
            .Cells(3, 2).font.size = 11
            .Cells(3, 3).value = "heure"
            .Cells(3, 3).font.italic = True
            .Cells(3, 3).font.bold = True
            .Cells(3, 3).font.size = 11
            .Cells(4, 3).value = "00H00"
            .Cells(5, 3).value = "01H00"
            .Cells(6, 3).value = "02H00"
            .Cells(7, 3).value = "03H00"
            .Cells(8, 3).value = "04H00"
            .Cells(9, 3).value = "05H00"
            .Cells(10, 3).value = "06H00"
            .Cells(11, 3).value = "07H00"
            .Cells(12, 3).value = "08H00"
            .Cells(13, 3).value = "09H00"
            .Cells(14, 3).value = "10H00"
            .Cells(15, 3).value = "11H00"
            .Cells(16, 3).value = "12H00"
            .Cells(17, 3).value = "13H00"
            .Cells(18, 3).value = "14H00"
            .Cells(19, 3).value = "15H00"
            .Cells(20, 3).value = "16H00"
            .Cells(21, 3).value = "17H00"
            .Cells(22, 3).value = "18H00"
            .Cells(23, 3).value = "19H00"
            .Cells(24, 3).value = "20H00"
            .Cells(25, 3).value = "21H00"
            .Cells(26, 3).value = "22H00"
            .Cells(27, 3).value = "23H00"
            .Cells(3, 4).value = "Visi."
            .Cells(3, 4).font.italic = True
            .Cells(3, 4).font.bold = True
            .Cells(3, 4).font.size = 11
            .Cells(3, 5).value = "DDD"
            .Cells(3, 5).font.italic = True
            .Cells(3, 5).font.bold = True
            .Cells(3, 5).font.size = 11
            .Cells(3, 6).value = "FF"
            .Cells(3, 6).font.italic = True
            .Cells(3, 6).font.bold = True
            .Cells(3, 6).font.size = 11
            .Cells(3, 7).value = "WW"
            .Cells(3, 7).font.italic = True
            .Cells(3, 7).font.bold = True
            .Cells(3, 7).font.size = 8
            .Cells(3, 8).value = "w1w2"
            .Cells(3, 8).font.italic = True
            .Cells(3, 8).font.bold = True
            .Cells(3, 8).font.size = 8
            .Cells(3, 9).value = "CL"
            .Cells(3, 9).font.italic = True
            .Cells(3, 9).font.bold = True
            .Cells(3, 9).font.size = 10
            .Cells(3, 10).value = "CM"
            .Cells(3, 10).font.italic = True
            .Cells(3, 10).font.bold = True
            .Cells(3, 10).font.size = 10
            .Cells(3, 11).value = "CH"
            .Cells(3, 11).font.italic = True
            .Cells(3, 11).font.bold = True
            .Cells(3, 11).font.size = 10
            .Cells(3, 12).value = "N"
            .Cells(3, 12).font.italic = True
            .Cells(3, 12).font.bold = True
            .Cells(3, 12).font.size = 10
            .Cells(3, 13).value = "h1"
            .Cells(3, 13).font.italic = True
            .Cells(3, 13).font.bold = True
            .Cells(3, 13).font.size = 10
            .Cells(3, 14).value = "n1"
            .Cells(3, 14).font.italic = True
            .Cells(3, 14).font.bold = True
            .Cells(3, 14).font.size = 10
            .Cells(3, 15).value = "C1"
            .Cells(3, 15).font.italic = True
            .Cells(3, 15).font.bold = True
            .Cells(3, 15).font.size = 10
            .Cells(3, 16).value = "h0"
            .Cells(3, 16).font.italic = True
            .Cells(3, 16).font.bold = True
            .Cells(3, 16).font.size = 10
            .Cells(3, 17).value = "n0"
            .Cells(3, 17).font.italic = True
            .Cells(3, 17).font.bold = True
            .Cells(3, 17).font.size = 10
            .Cells(3, 18).value = "C0"
            .Cells(3, 18).font.italic = True
            .Cells(3, 18).font.bold = True
            .Cells(3, 18).font.size = 10
            .Cells(3, 19).value = "h2"
            .Cells(3, 19).font.italic = True
            .Cells(3, 19).font.bold = True
            .Cells(3, 19).font.size = 10
            .Cells(3, 20).value = "n2"
            .Cells(3, 20).font.italic = True
            .Cells(3, 20).font.bold = True
            .Cells(3, 20).font.size = 10
            .Cells(3, 21).value = "C2"
            .Cells(3, 21).font.italic = True
            .Cells(3, 21).font.bold = True
            .Cells(3, 21).font.size = 10
            .Cells(3, 22).value = "T.air"
            .Cells(3, 22).font.italic = True
            .Cells(3, 22).font.bold = True
            .Cells(3, 22).font.size = 11
            .Cells(3, 23).value = "ew"
            .Cells(3, 23).font.italic = True
            .Cells(3, 23).font.bold = True
            .Cells(3, 23).font.size = 11
            .Cells(3, 24).value = "Td"
            .Cells(3, 24).font.italic = True
            .Cells(3, 24).font.bold = True
            .Cells(3, 24).font.size = 11
            .Cells(3, 25).value = "U%"
            .Cells(3, 25).font.italic = True
            .Cells(3, 25).font.bold = True
            .Cells(3, 25).font.size = 11
            .Cells(3, 26).value = "P.st°"
            .Cells(3, 26).font.italic = True
            .Cells(3, 26).font.bold = True
            .Cells(3, 26).font.size = 11
            .Cells(3, 27).value = "P.mer"
            .Cells(3, 27).font.italic = True
            .Cells(3, 27).font.bold = True
            .Cells(3, 27).font.size = 11
            .Cells(3, 28).value = "RR"
            .Cells(3, 28).font.italic = True
            .Cells(3, 28).font.bold = True
            .Cells(3, 28).font.size = 11
            .Cells(3, 29).value = "dur.RR"
            .Cells(3, 29).font.italic = True
            .Cells(3, 29).font.bold = True
            .Cells(3, 29).font.size = 11
            .Columns("C:C").ColumnWidth = 5.43
            .Columns("D:D").ColumnWidth = 7
            .Columns("E:E").ColumnWidth = 6
            .Columns("F:F").ColumnWidth = 4
            .Columns("G:G").ColumnWidth = 3.86
            .Columns("H:H").ColumnWidth = 4.86
            .Columns("I:I").ColumnWidth = 2.14
            .Columns("J:J").ColumnWidth = 2.4
            .Columns("K:K").ColumnWidth = 2.29
            .Columns("L:L").ColumnWidth = 2
            .Columns("M:M").ColumnWidth = 6
            .Columns("N:N").ColumnWidth = 2.14
            .Columns("O:O").ColumnWidth = 2.14
            .Columns("P:P").ColumnWidth = 6
            .Columns("Q:Q").ColumnWidth = 2.14
            .Columns("R:R").ColumnWidth = 2.14
            .Columns("S:S").ColumnWidth = 6
            .Columns("T:T").ColumnWidth = 2.14
            .Columns("U:U").ColumnWidth = 2.14
            .Columns("V:V").ColumnWidth = 4.7
            .Columns("W:W").ColumnWidth = 4
            .Columns("X:X").ColumnWidth = 4.7
            .Columns("Y:Y").ColumnWidth = 4
            .Columns("Z:Z").ColumnWidth = 7
            .Columns("AA:AA").ColumnWidth = 7
            .Columns("AB:AB").ColumnWidth = 5
            .Columns("AC:AC").ColumnWidth = 7
            .Cells.NumberFormat = "@"
            For i As Integer = 4 To 27
                .Cells(i, 1).value = IDstation
                .Cells(i, 2).value = Row_Date
            Next

            .Cells(1, 2).value = "Date:"
            .Cells(1, 2).font.italic = True
            .Cells(1, 2).font.bold = True
            .Cells(1, 2).font.size = 11
            .Range("C1:E1").Merge()
            .Cells(1, 3).value = Date_CRQ
            .Cells(1, 3).font.bold = True
            .Range("C1:E1").HorizontalAlignment = Excel.Constants.xlCenter
            .Range("G2:I2").Merge()
            .Cells(2, 7).value = "Station:"
            .Cells(2, 7).font.bold = True
            .Cells(2, 7).font.italic = True
            .Range("G2:I2").HorizontalAlignment = Excel.Constants.xlCenter
            .Range("J2:L2").Merge()
            .Cells(2, 10).value = IDstation
            .Cells(2, 10).font.bold = True
            .Range("J2:L2").HorizontalAlignment = Excel.Constants.xlCenter


        End With
        '*********feuille 2
        With PhenWsh
            .Cells(1, 1).value = "Indicatif"
            .Cells(1, 2).value = "Date"
            .Cells(1, 3).value = "Phénomène"   'mise en forme sheet-phenomenes
            .Cells(1, 4).value = "Code"
            .Cells(1, 5).value = "H.début"
            .Cells(1, 6).value = "H.fin"
            .Cells(1, 7).value = "Intensité"
            .Cells(1, 8).value = "Secteur"
            .Cells(1, 9).value = "Visi.mini"
            .Cells(1, 10).value = "heure"
            .Cells(1, 11).value = "Hauteur"
            .Rows("1:1").Font.Bold = True
            .Rows("1:1").Font.size = 10
            .Columns("A:A").ColumnWidth = 6
            .Columns("B:B").ColumnWidth = 10
            .Columns("C:C").ColumnWidth = 12
            .Columns("D:D").ColumnWidth = 5
            .Columns("E:E").ColumnWidth = 8
            .Columns("F:F").ColumnWidth = 8
            .Columns("G:G").ColumnWidth = 8
            .Columns("H:H").ColumnWidth = 8
            .Columns("I:I").ColumnWidth = 8
            .Columns("J:J").ColumnWidth = 8
            .Columns("K:K").ColumnWidth = 8
            .Columns("L:L").ColumnWidth = 8
            .Columns("M:M").ColumnWidth = 8
            .Columns("N:N").ColumnWidth = 8
            .Columns("O:O").ColumnWidth = 8
            .Columns("P:P").ColumnWidth = 8
            .Cells.NumberFormat = "@"
        End With
        '*********feuille 3
        With ExtrWsh
            .Cells.NumberFormat = "@"

        End With

        XlWbk.SaveAs(File_Name)
        XlWbk.Close(SaveChanges:=True)
        releaseObject(HrlyWsh)
        releaseObject(PhenWsh)
        releaseObject(ExtrWsh)
        releaseObject(XlWbk)
        releaseObject(XlApp)
    End Sub

    Private Sub btntelem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btntelem.Click
        btnforcetelem.Enabled = False
        TabControl1.SelectedTab = tbtelem

    End Sub

    Private Sub btnextr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnextr.Click

        TabControl1.SelectedTab = tbextr
        btnforcer_extrm.Enabled = False

    End Sub

    Private Sub btntbphn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btntbphn.Click

        TabControl1.SelectedTab = tbphen
        DataGridView1.EditMode = DataGridViewEditMode.EditProgrammatically
        btncorrige.Hide()
        Try

            Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim dataSet As System.Data.DataSet
            Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
            Dim path As String = [String].Format(disck & ":\" & Folder_crq & "\" & IDstation & "crq{0}.xls", DateTime.Now.ToString("ddMMyyyy"))

            MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1;';") 'access database engine data connectivity component
            MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [phenomenes$C1:K500]", MyConnection)

            dataSet = New System.Data.DataSet
            MyCommand.Fill(dataSet)
            DataGridView1.DataSource = dataSet.Tables(0)

            MyConnection.Close()
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try

    End Sub

    Private Sub btnstart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnstart.Click

        Dim startphn As String = DateTime.Now.ToString("HH:mm")
        txtstart.Text = startphn

    End Sub

    Private Sub btnstop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnstop.Click

        Dim stopphn As String = DateTime.Now.ToString("HH:mm")
        txtend.Text = stopphn

    End Sub

    Private Sub btnphvalid_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnphvalid.Click
        Dim r As Integer = 2
        Dim lastRow As Integer
        Dim currpheno As Integer
        Dim Name_CRQ As String = String.Format(disck & ":\" & Folder_crq & "\" & IDstation & "crq{0}.xls", Date.Now.ToString("ddMMyyyy"))
        Dim DateOfPhen As String = Date.Now.ToString("ddMMyyyy")
        If System.IO.File.Exists(Name_CRQ) Then
            If txtcode.Text = "" Then
                MsgBox("la case code phénomène est vide,vous devez inserer le code du phénomène", MsgBoxStyle.Exclamation, vbAbort)
                Exit Sub
            ElseIf txtcode.Text <> "" Then
                If txtstart.Text = "" And txtend.Text = "" Then
                    MsgBox("inserer debut ou fin du phénomène", MsgBoxStyle.Exclamation, vbAbort)
                    Exit Sub

                ElseIf txtstart.Text <> "" Or txtend.Text <> "" Then
                    Dim XlApp As Excel.Application
                    Dim XlWbk As Excel.Workbook
                    Dim PhenWsh As Excel.Worksheet
                    XlApp = New Excel.Application
                    XlWbk = XlApp.Workbooks.Open(Name_CRQ)
                    PhenWsh = XlWbk.Worksheets("phenomenes")
                    ''code validphen
                    lastRow = PhenWsh.Range("D500").End(Excel.XlDirection.xlUp).Row
                    currpheno = PhenWsh.Range("F500").End(Excel.XlDirection.xlUp).Row
                    If txtstart.Text <> "" And txtend.Text = "" Then
                        r = lastRow + 1
                    ElseIf txtstart.Text <> "" And txtend.Text <> "" Then
                        If txtstart.Text = PhenWsh.Cells(lastRow, 5).Value And PhenWsh.Cells(lastRow, 2).Value = txtcode.Text Then
                            r = lastRow
                        Else
                            r = lastRow + 1
                        End If

                    ElseIf txtstart.Text = "" And txtend.Text <> "" Then
                        If PhenWsh.Cells(lastRow, 4).Value = txtcode.Text Then
                            txtstart.Text = PhenWsh.Cells(lastRow, 5).Value
                            r = lastRow
                        ElseIf PhenWsh.Cells(lastRow, 4).Value <> txtcode.Text Then
                            For i As Integer = 2 To lastRow
                                If PhenWsh.Cells(i, 6).Value = "" Then
                                    If PhenWsh.Cells(i, 4).Value = txtcode.Text Then
                                        txtstart.Text = PhenWsh.Cells(i, 5).Value
                                        r = i
                                        MsgBox("vous avez mis fin à un phénomène dans la ligne:" & i, MsgBoxStyle.Information, vbOKOnly)
                                    ElseIf PhenWsh.Cells(i, 4).Value <> txtcode.Text Then
                                        For j As Integer = i + 1 To lastRow
                                            If PhenWsh.Cells(j, 6).Value = "" And PhenWsh.Cells(j, 4).Value = txtcode.Text Then
                                                txtstart.Text = PhenWsh.Cells(j, 5).Value
                                                r = j
                                                MsgBox("vous avez mis fin à un phénomène dans la ligne:" & j, MsgBoxStyle.Information, vbOKOnly)

                                            ElseIf PhenWsh.Cells(j, 6).Value = "" And PhenWsh.Cells(j, 4).Value <> txtcode.Text Then
                                                MsgBox("Vous avez plusieurs phénomènes en cours,selectionnez un Index convenant avec la ligne du Tableau,puis clicker sur Forcer!" & j, MsgBoxStyle.Exclamation, vbAbort)
                                                txtstart.Text = PhenWsh.Cells(j, 5).Value
                                                Exit Sub
                                            End If
                                        Next j
                                    End If

                                End If
                            Next i
                        End If

                    End If
                    PhenWsh.Cells(r, 1).Value = IDstation
                    PhenWsh.Cells(r, 2).Value = DateOfPhen
                    PhenWsh.Cells(r, 3).Value = cbophen.Text
                    PhenWsh.Cells(r, 4).Value = txtcode.Text
                    PhenWsh.Cells(r, 5).Value = txtstart.Text
                    PhenWsh.Cells(r, 6).Value = txtend.Text
                    PhenWsh.Cells(r, 7).Value = cbointens.SelectedItem
                    PhenWsh.Cells(r, 8).Value = cbosctr.SelectedItem
                    PhenWsh.Cells(r, 9).Value = txtvvmin.Text
                    PhenWsh.Cells(r, 10).Value = txthvvmin.Text
                    PhenWsh.Cells(r, 11).Value = txthaut.Text
                    XlWbk.Close(SaveChanges:=True)
                    releaseObject(PhenWsh)
                    releaseObject(XlWbk)
                    releaseObject(XlApp)

                End If

            End If
        End If

        connecttophenom()

    End Sub

    Protected Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim param As CreateParams = MyBase.CreateParams
            param.ClassStyle = param.ClassStyle Or &H200
            Return param
        End Get
    End Property

    Private Sub btncorrige_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btncorrige.Click
        If IsNumeric(Cor_Phen_Index) Then
            Dim index As Integer = Cor_Phen_Index
            Dim Name_CRQ As String = String.Format(disck & ":\" & Folder_crq & "\" & IDstation & "crq{0}.xls", Date.Now.ToString("ddMMyyyy"))
            Dim XlApp As Excel.Application
            Dim XlWbk As Excel.Workbook
            Dim PhenWsh As Excel.Worksheet
            XlApp = New Excel.Application
            XlWbk = XlApp.Workbooks.Open(Name_CRQ)
            PhenWsh = XlWbk.Worksheets("phenomenes")
            PhenWsh.Cells(index, 3).Value = cbophen.Text
            PhenWsh.Cells(index, 4).Value = txtcode.Text
            PhenWsh.Cells(index, 5).Value = txtstart.Text
            PhenWsh.Cells(index, 6).Value = txtend.Text
            PhenWsh.Cells(index, 7).Value = cbointens.SelectedItem
            PhenWsh.Cells(index, 8).Value = cbosctr.SelectedItem
            PhenWsh.Cells(index, 9).Value = txtvvmin.Text
            PhenWsh.Cells(index, 10).Value = txthvvmin.Text
            PhenWsh.Cells(index, 11).Value = txthaut.Text
            XlWbk.Close(SaveChanges:=True)
            releaseObject(PhenWsh)
            releaseObject(XlWbk)
            releaseObject(XlApp)
        End If
        Cor_Phen_Index = Nothing
        connecttophenom()
        btnphvalid.Show()
        btncorrige.Hide()
    End Sub

    Private Sub btnconsult_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnconsult.Click
        frmconsult.ShowDialog()

    End Sub

    Private Sub connecttophenom()

        Try

            Dim MyConnection As System.Data.OleDb.OleDbConnection
            Dim dataSet As System.Data.DataSet
            Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
            Dim path As String = [String].Format(disck & ":\" & Folder_crq & "\" & IDstation & "crq{0}.xls", DateTime.Now.ToString("ddMMyyyy"))

            MyConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 8.0;HDR=YES;IMEX=1;';") 'access database engine data connectivity component
            MyCommand = New System.Data.OleDb.OleDbDataAdapter("select * from [phenomenes$C1:K500]", MyConnection)

            dataSet = New System.Data.DataSet
            MyCommand.Fill(dataSet)
            DataGridView1.DataSource = dataSet.Tables(0)

            MyConnection.Close()
        Catch ex As Exception
            MsgBox(ex.Message.ToString)
        End Try

    End Sub

    Private Sub btnchargerhor_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnchargerhor.Click
        ADRESS = My.Settings.adressCAOBS
        sourceADRESS = "\\" & ADRESS & source

        If ADRESS <> "" Then
            If Date.Now.Hour <> 0 And Date.Now.Minute < 30 Then
                Dim Name_CRQ As String = String.Format(disck & ":\" & Folder_crq & "\" & IDstation & "crq{0}.xls", Date.Now.ToString("ddMMyyyy"))
                Dim XlApp As Excel.Application
                Dim XlWbk As Excel.Workbook
                Dim HrlyWsh As Excel.Worksheet
                XlApp = New Excel.Application
                XlWbk = XlApp.Workbooks.Open(Name_CRQ)
                HrlyWsh = XlWbk.Worksheets("horaire")

                Dim fichierminute As String = "A_" & Date.Now.ToString("MMdd") & ".xls"
                Dim fichierhoraire As String = "S_" & Date.Now.ToString("MMdd") & ".xls"

                Dim sourceminute As String = sourceADRESS & fichierminute
                Dim localminute As String = disck & BD_crq & fichierminute
                Dim localhoraire As String = disck & BD_crq & fichierhoraire

                Try
                    File.Copy(sourceADRESS & fichierminute, disck & BD_crq & fichierminute, True)

                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try

                Try
                    File.Copy(sourceADRESS & fichierhoraire, disck & BD_crq & fichierhoraire, True)

                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try

                Dim t As Integer
                Dim tr As Integer = DateTime.Now.Hour
                t = tr + 4

                If System.IO.File.Exists(localminute) Then
                    Dim minutewbk As Excel.Workbook
                    Dim minutwsh As Excel.Worksheet
                    Dim lastRow As Integer
                    minutewbk = XlApp.Workbooks.Open(localminute) ' ne pas oublier de la fermer
                    minutwsh = minutewbk.Sheets(1)
                    lastRow = minutwsh.Range("A1500").End(Excel.XlDirection.xlUp).Row

                    minutwsh.Cells(lastRow, SmDD).Copy(HrlyWsh.Cells(t, 5))
                    minutwsh.Cells(lastRow, SmFF).Copy(HrlyWsh.Cells(t, 6))
                    minutwsh.Cells(lastRow, SmT).Copy(HrlyWsh.Cells(t, 22))
                    minutwsh.Cells(lastRow, SmTd).Copy(HrlyWsh.Cells(t, 24))
                    minutwsh.Cells(lastRow, SmHr).Copy(HrlyWsh.Cells(t, 25))
                    minutwsh.Cells(lastRow, SmPs).Copy(HrlyWsh.Cells(t, 26))
                    minutwsh.Cells(lastRow, SmPm).Copy(HrlyWsh.Cells(t, 27))
                    minutewbk.Close(SaveChanges:=False)
                    minutwsh = Nothing
                    minutewbk = Nothing
                    releaseObject(minutwsh)
                    releaseObject(minutewbk)
                Else
                    MsgBox("Le fichier Horaire" & fichierminute & " n'existe pas,verifiez le chemain d'accés", MsgBoxStyle.Critical, vbOKOnly)
                End If

                If System.IO.File.Exists(localhoraire) Then
                    Dim horairewb As Excel.Workbook
                    Dim horairewsh As Excel.Worksheet
                    Dim horlastrow As Integer
                    horairewb = XlApp.Workbooks.Open(localhoraire)
                    horairewsh = horairewb.Sheets(1)
                    horlastrow = horairewsh.Range("A30").End(Excel.XlDirection.xlUp).Row
                    horairewsh.Cells(horlastrow, ShEw).Copy(HrlyWsh.Cells(t, 23))
                    horairewsh.Cells(horlastrow, ShRR).Copy(HrlyWsh.Cells(t, 28))
                    horairewb.Close(SaveChanges:=False)
                    horairewsh = Nothing
                    horairewb = Nothing
                    releaseObject(horairewsh)
                    releaseObject(horairewb)
                Else
                    MsgBox("Le fichier Horaire" & fichierhoraire & " n'existe pas,verifiez le chemain d'accés", MsgBoxStyle.Critical, vbOKOnly)
                End If

                HrlyWsh.Cells(t, 29).value = txtdurRR.Text

                txtDD.Text = HrlyWsh.Cells(t, 5).value
                txtFF.Text = HrlyWsh.Cells(t, 6).value
                txtT.Text = HrlyWsh.Cells(t, 22).value
                txtTd.Text = HrlyWsh.Cells(t, 24).value
                txtHR.Text = HrlyWsh.Cells(t, 25).value
                txtPst.Text = HrlyWsh.Cells(t, 26).value
                txtPmer.Text = HrlyWsh.Cells(t, 27).value
                txtew.Text = HrlyWsh.Cells(t, 23).value
                txtRR.Text = HrlyWsh.Cells(t, 28).value

                XlWbk.Close(SaveChanges:=True)
                releaseObject(HrlyWsh)
                releaseObject(XlWbk)
                releaseObject(XlApp)

            ElseIf Date.Now.Hour <> 0 And Date.Now.Minute > 30 Then
                MsgBox("Il n'est pas l'heure pour valider les télémesures", MsgBoxStyle.Exclamation, vbOKOnly)

            ElseIf Date.Now.Hour = 0 And Date.Now.Minute < 30 Then

                Dim Name_CRQ As String = String.Format(disck & ":\" & Folder_crq & "\" & IDstation & "crq{0}.xls", Date.Now.ToString("ddMMyyyy"))
                Dim XlApp As Excel.Application
                Dim XlWbk As Excel.Workbook
                Dim HrlyWsh As Excel.Worksheet
                XlApp = New Excel.Application
                XlWbk = XlApp.Workbooks.Open(Name_CRQ)
                HrlyWsh = XlWbk.Worksheets("horaire")
                Dim fichierminute As String = "A_" & Date.Now.AddDays(-1).ToString("MMdd") & ".xls"
                Dim fichierhoraire As String = "S_" & Date.Now.AddDays(-1).ToString("MMdd") & ".xls"
                Dim sourceminute As String = sourceADRESS & fichierminute
                Dim localminute As String = disck & BD_crq & fichierminute
                Dim localhoraire As String = disck & BD_crq & fichierhoraire

                Try
                    File.Copy(sourceADRESS & fichierminute, disck & BD_crq & fichierminute, True)

                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try

                Try
                    File.Copy(sourceADRESS & fichierhoraire, disck & BD_crq & fichierhoraire, True)

                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try

                Dim t As Integer
                Dim tr As Integer = DateTime.Now.Hour
                t = tr + 4
                If System.IO.File.Exists(localminute) Then
                    Dim minutewbk As Excel.Workbook
                    Dim minutwsh As Excel.Worksheet
                    Dim lastRow As Integer
                    minutewbk = XlApp.Workbooks.Open(localminute) ' ne pas oublier de la fermer
                    minutwsh = minutewbk.Sheets(1)
                    lastRow = minutwsh.Range("A1500").End(Excel.XlDirection.xlUp).Row
                    minutwsh.Cells(lastRow, SmDD).Copy(HrlyWsh.Cells(t, 5))
                    minutwsh.Cells(lastRow, SmFF).Copy(HrlyWsh.Cells(t, 6))
                    minutwsh.Cells(lastRow, SmT).Copy(HrlyWsh.Cells(t, 22))
                    minutwsh.Cells(lastRow, SmTd).Copy(HrlyWsh.Cells(t, 24))
                    minutwsh.Cells(lastRow, SmHr).Copy(HrlyWsh.Cells(t, 25))
                    minutwsh.Cells(lastRow, SmPs).Copy(HrlyWsh.Cells(t, 26))
                    minutwsh.Cells(lastRow, SmPm).Copy(HrlyWsh.Cells(t, 27))
                    minutewbk.Close(SaveChanges:=False)
                    minutwsh = Nothing
                    minutewbk = Nothing
                    releaseObject(minutwsh)
                    releaseObject(minutewbk)
                Else
                    MsgBox("Le fichier Horaire" & fichierminute & " n'existe pas,verifiez le chemain d'accés", MsgBoxStyle.Critical, vbOKOnly)

                End If

                If System.IO.File.Exists(localhoraire) Then
                    Dim horairewb As Excel.Workbook
                    Dim horairewsh As Excel.Worksheet
                    Dim horlastrow As Integer
                    horairewb = XlApp.Workbooks.Open(localhoraire)
                    horairewsh = horairewb.Sheets(1)
                    horlastrow = horairewsh.Range("A30").End(Excel.XlDirection.xlUp).Row
                    horairewsh.Cells(horlastrow, ShEw).Copy(HrlyWsh.Cells(t, 23))
                    horairewsh.Cells(horlastrow, ShRR).Copy(HrlyWsh.Cells(t, 28))
                    horairewb.Close(SaveChanges:=False)
                    horairewsh = Nothing
                    horairewb = Nothing
                    releaseObject(horairewsh)
                    releaseObject(horairewb)
                Else
                    MsgBox("Le fichier Horaire" & fichierhoraire & " n'existe pas,verifiez le chemain d'accés", MsgBoxStyle.Critical, vbOKOnly)

                End If

                HrlyWsh.Cells(t, 29).value = txtdurRR.Text

                txtDD.Text = hrly.HrlyWsh(t, 5).value
                txtFF.Text = hrly.HrlyWsh(t, 6).value
                txtT.Text = hrly.HrlyWsh(t, 22).value
                txtTd.Text = hrly.HrlyWsh(t, 24).value
                txtHR.Text = hrly.HrlyWsh(t, 25).value
                txtPst.Text = hrly.HrlyWsh(t, 26).value
                txtPmer.Text = hrly.HrlyWsh(t, 27).value
                txtew.Text = hrly.HrlyWsh(t, 23).value
                txtRR.Text = hrly.HrlyWsh(t, 28).value

                XlWbk.Close(SaveChanges:=True)
                releaseObject(HrlyWsh)
                releaseObject(XlWbk)
                releaseObject(XlApp)
            ElseIf Date.Now.Hour = 0 And Date.Now.Minute > 30 Then

                MsgBox("Il n'est pas l'heure pour valider les télémesures", MsgBoxStyle.Exclamation, vbOKOnly)

            End If

        ElseIf ADRESS = "" Then
            MsgBox("configuration manquante, verifiez l'adresse du CAOBS dans le menu config.!", MsgBoxStyle.Exclamation, vbOKOnly)
        End If

    End Sub

    Private Sub btncorrigetelem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btncorrigetelem.Click
        btnforcetelem.Enabled = True

    End Sub

    Private Sub btnforcetelem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnforcetelem.Click

        Dim cor As Integer = Date.Now.Hour
        Dim cor_rw As Integer = cor + 4
        Dim Name_CRQ As String = String.Format(disck & ":\" & Folder_crq & "\" & IDstation & "crq{0}.xls", Date.Now.ToString("ddMMyyyy"))
        Dim XlApp As Excel.Application
        Dim XlWbk As Excel.Workbook
        Dim HrlyWsh As Excel.Worksheet
        XlApp = New Excel.Application
        XlWbk = XlApp.Workbooks.Open(Name_CRQ)
        HrlyWsh = XlWbk.Worksheets("horaire")

        HrlyWsh.Cells(cor_rw, 5).Value = txtDD.Text
        HrlyWsh.Cells(cor_rw, 6).Value = txtFF.Text
        HrlyWsh.Cells(cor_rw, 22).Value = txtT.Text
        HrlyWsh.Cells(cor_rw, 24).Value = txtTd.Text
        HrlyWsh.Cells(cor_rw, 25).Value = txtHR.Text
        HrlyWsh.Cells(cor_rw, 26).Value = txtPst.Text
        HrlyWsh.Cells(cor_rw, 27).Value = txtPmer.Text
        HrlyWsh.Cells(cor_rw, 28).Value = txtRR.Text
        HrlyWsh.Cells(cor_rw, 23).Value = txtew.Text
        HrlyWsh.Cells(cor_rw, 29).Value = txtdurRR.Text

        XlWbk.Close(SaveChanges:=True)
        releaseObject(HrlyWsh)
        releaseObject(XlWbk)
        releaseObject(XlApp)
        MsgBox("Le forçage des valeurs de:" & cor & "heure est effectué", MsgBoxStyle.Information)
        btnforcetelem.Enabled = False

    End Sub


    Private Sub btncorrige_extrm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btncorrige_extrm.Click
        btnforcer_extrm.Enabled = True

    End Sub

    Private Sub btncharger_extrem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btncharger_extrem.Click

        ADRESS = My.Settings.adressCAOBS
        sourceADRESS = "\\" & ADRESS & source
        Dim XlApp As Excel.Application
        Dim ystrdycrq As Excel.Workbook
        Dim ystrdextr As Excel.Worksheet
        XlApp = New Excel.Application
        Dim ystrdname As String = [String].Format(disck & ":\" & Folder_crq & "\" & IDstation & "crq{0}.xls", DateTime.Now.AddDays(-1).ToString("ddMMyyyy"))
        Dim ystrdStmp As String = Date.Now.AddDays(-1).ToString("ddMMyyyy")
        Dim fichiextreme As String = "K_" & Date.Now.AddDays(-1).ToString("MMdd") & ".xls"
        Dim sourcextreme As String = sourceADRESS & fichiextreme
        Dim localextreme As String = disck & BD_crq & fichiextreme

        Dim fichihor As String = "S_" & Date.Now.AddDays(-1).ToString("MMdd") & ".xls"
        Dim sourhor As String = sourceADRESS & fichihor
        Dim localHor As String = disck & BD_crq & fichihor
        Dim sourceHor As String = sourceADRESS & fichihor
        If System.IO.File.Exists(ystrdname) Then
            If System.IO.File.Exists(sourcextreme) Then
                File.Copy(sourceADRESS & fichiextreme, disck & BD_crq & fichiextreme, True)
                Dim extrwb As Excel.Workbook
                Dim extrwsh As Excel.Worksheet
                ystrdycrq = XlApp.Workbooks.Open(ystrdname)
                ystrdextr = ystrdycrq.Worksheets("extrèmes")
                extrwb = XlApp.Workbooks.Open(localextreme) ' ne pas oublier de la fermer
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
                txtUmin.Text = ystrdextr.Cells(2, 7).text
                txtHH_Umin.Text = ystrdextr.Cells(2, 8).text
                txtUmax.Text = ystrdextr.Cells(2, 9).text
                txtHH_Umax.Text = ystrdextr.Cells(2, 10).text
                txtTmin.Text = ystrdextr.Cells(2, 3).text
                txtHH_Tmin.Text = ystrdextr.Cells(2, 4).text
                txtTmax.Text = ystrdextr.Cells(2, 5).text
                txtHH_Tmax.Text = ystrdextr.Cells(2, 6).text
                txtRR_cumul.Text = ystrdextr.Cells(2, 11).value
                txtDI.Text = ystrdextr.Cells(2, 12).value
                txtRGlob.Text = ystrdextr.Cells(2, 14).value
                txtTxSol.Text = ystrdextr.Cells(2, 22).value
                txtTnSol.Text = ystrdextr.Cells(2, 23).value

                If System.IO.File.Exists(sourceHor) Then
                    File.Copy(sourceHor, localHor, True)
                    Dim HorWBK As Excel.Workbook
                    Dim HorWsh As Excel.Worksheet
                    HorWBK = XlApp.Workbooks.Open(localHor)
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

                    txtDDmax_moy.Text = DDmaxmoy
                    txtFFmax_moy.Text = MaxFFmoy
                    txtHH_moy.Text = HHmm_moy
                    txtDDmax_ins.Text = DDmaxins
                    txtFFmax_ins.Text = MaxFFins
                    txtHH_ins.Text = HHmm_ins

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
                releaseObject(XlApp)
                MsgBox("Les étrêmes des la veille sont ajoutés avec succes.", MsgBoxStyle.Information, vbOKOnly)

            Else
                MsgBox("le ficier extrêmes n'est pas encore prêt,il sera prêt après 06H00 TU", MsgBoxStyle.Exclamation, vbOKOnly)

            End If
        Else
            MsgBox("Le CRQ du " & ystrdname & " n'existe pas,verifiez le dossier 'données station'", MsgBoxStyle.Critical, vbOKOnly)

        End If

    End Sub

    Private Sub btnforcer_extrm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnforcer_extrm.Click
        Dim XlApp As Excel.Application
        Dim ystrdycrq As Excel.Workbook
        Dim ystrdextr As Excel.Worksheet
        Dim ystrdname As String = [String].Format(disck & ":\" & Folder_crq & "\" & IDstation & "crq{0}.xls", DateTime.Now.AddDays(-1).ToString("ddMMyyyy"))
        Dim ystrdaystmp As String = Date.Now.AddDays(-1).ToString("ddMMyyyy")

        If System.IO.File.Exists(ystrdname) Then
            XlApp = New Excel.Application
            ystrdycrq = XlApp.Workbooks.Open(ystrdname)
            ystrdextr = ystrdycrq.Worksheets("extrèmes")
            ystrdextr.Cells(1, 1).value = "Indicatif"
            ystrdextr.Cells(1, 2).value = "Date"
            ystrdextr.Cells(1, 3).value = "Min Tair"
            ystrdextr.Cells(1, 4).value = "Heure du min"
            ystrdextr.Cells(1, 5).value = "Max Tair"
            ystrdextr.Cells(1, 6).value = "Heure du max"
            ystrdextr.Cells(1, 7).value = "Min HR"
            ystrdextr.Cells(1, 8).value = "Heure du min"
            ystrdextr.Cells(1, 9).value = "Max HR"
            ystrdextr.Cells(1, 10).value = "Heure du max"
            ystrdextr.Cells(1, 11).value = "Pluie"
            ystrdextr.Cells(1, 12).value = "Durée Insol."
            ystrdextr.Cells(1, 13).value = "Insol 1/10 heure"
            ystrdextr.Cells(1, 14).value = "Ray. global"
            ystrdextr.Cells(1, 15).value = "DD vent max Inst."
            ystrdextr.Cells(1, 16).value = "FF vent max Inst."
            ystrdextr.Cells(1, 17).value = "Heure du max"
            ystrdextr.Cells(1, 18).value = "FF vent max 10'"
            ystrdextr.Cells(1, 19).value = "FF vent max 10'"
            ystrdextr.Cells(1, 20).value = "Heure du max"
            ystrdextr.Cells(1, 21).value = "Neige"
            ystrdextr.Cells(1, 22).value = "Max T+10"
            ystrdextr.Cells(1, 23).value = "Min T+10"
            ystrdextr.Cells(2, 1).value = IDstation
            ystrdextr.Cells(2, 2).value = ystrdaystmp
            ystrdextr.Cells(2, 7).value = txtUmin.Text
            ystrdextr.Cells(2, 8).value = txtHH_Umin.Text
            ystrdextr.Cells(2, 9).value = txtUmax.Text
            ystrdextr.Cells(2, 10).value = txtHH_Umax.Text
            ystrdextr.Cells(2, 3).value = txtTmin.Text
            ystrdextr.Cells(2, 4).value = txtHH_Tmin.Text
            ystrdextr.Cells(2, 5).value = txtTmax.Text
            ystrdextr.Cells(2, 6).value = txtHH_Tmax.Text
            ystrdextr.Cells(2, 11).value = txtRR_cumul.Text
            ystrdextr.Cells(2, 12).value = txtDI.Text
            ystrdextr.Cells(2, 13).value = Math.Round((ystrdextr.Cells(2, 12).value) / 6)
            ystrdextr.Cells(2, 14).value = txtRGlob.Text
            ystrdextr.Cells(2, 15).value = txtDDmax_ins.Text
            ystrdextr.Cells(2, 16).value = txtFFmax_ins.Text
            ystrdextr.Cells(2, 17).value = txtHH_ins.Text
            ystrdextr.Cells(2, 18).value = txtDDmax_moy.Text
            ystrdextr.Cells(2, 19).value = txtFFmax_moy.Text
            ystrdextr.Cells(2, 20).value = txtHH_moy.Text
            ystrdextr.Cells(1, 21).value = txtSN.Text
            ystrdextr.Cells(1, 22).value = txtTxSol.Text
            ystrdextr.Cells(1, 23).value = txtTnSol.Text
            btnforcer_extrm.Enabled = False

            ystrdycrq.Close(SaveChanges:=True)
            ystrdextr = Nothing
            ystrdycrq = Nothing
            releaseObject(ystrdextr)
            releaseObject(ystrdycrq)
        Else
            MsgBox("Le CRQ de la veille" & ystrdname & " n'existe pas,verifiez le dossier 'données station'", MsgBoxStyle.Critical, vbOKOnly)
        End If

        btnforcer_extrm.Enabled = False

    End Sub

    Private Sub horaireauto()

        ADRESS = My.Settings.adressCAOBS
        sourceADRESS = "\\" & ADRESS & source
        Dim XlApp As Excel.Application
        Dim crq As Excel.Workbook
        Dim hrly As Excel.Worksheet
        Dim nomfichier As String = [String].Format(disck & ":\" & Folder_crq & "\" & IDstation & "crq{0}.xls", DateTime.Now.ToString("ddMMyyyy"))
        If System.IO.File.Exists(nomfichier) Then
            XlApp = New Excel.Application
            crq = XlApp.Workbooks.Open(nomfichier)
            hrly = crq.Worksheets("horaire")

            Dim fichierminute As String = "A_" & Date.Now.ToString("MMdd") & ".xls"
            Dim fichierhoraire As String = "S_" & Date.Now.ToString("MMdd") & ".xls"

            Dim sourceminute As String = sourceADRESS & fichierminute
            Dim localminute As String = disck & BD_crq & fichierminute
            Dim localhoraire As String = disck & BD_crq & fichierhoraire

            Try
                File.Copy(sourceADRESS & fichierminute, disck & BD_crq & fichierminute, True)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            Try
                File.Copy(sourceADRESS & fichierhoraire, disck & BD_crq & fichierhoraire, True)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            Dim t As Integer
            Dim tr As Integer = DateTime.Now.Hour
            t = tr + 4

            If System.IO.File.Exists(localminute) Then
                Dim minutewbk As Excel.Workbook
                Dim minutwsh As Excel.Worksheet
                Dim lastRow As Integer
                minutewbk = XlApp.Workbooks.Open(localminute)
                minutwsh = minutewbk.Sheets(1)
                lastRow = minutwsh.Range("A1500").End(Excel.XlDirection.xlUp).Row

                minutwsh.Cells(lastRow, SmDD).Copy(hrly.Cells(t, 5))
                minutwsh.Cells(lastRow, SmFF).Copy(hrly.Cells(t, 6))
                minutwsh.Cells(lastRow, SmT).Copy(hrly.Cells(t, 22))
                minutwsh.Cells(lastRow, SmTd).Copy(hrly.Cells(t, 24))
                minutwsh.Cells(lastRow, SmHr).Copy(hrly.Cells(t, 25))
                minutwsh.Cells(lastRow, SmPs).Copy(hrly.Cells(t, 26))
                minutwsh.Cells(lastRow, SmPm).Copy(hrly.Cells(t, 27))

                minutewbk.Close(SaveChanges:=False)

                minutwsh = Nothing
                minutewbk = Nothing

                releaseObject(minutwsh)
                releaseObject(minutewbk)

            Else
                MsgBox("Le fichier Horaire" & fichierminute & " n'existe pas,verifiez le chemain d'accés", vbOKOnly)

            End If


            If System.IO.File.Exists(localhoraire) Then

                Dim horairewb As Excel.Workbook
                Dim horairewsh As Excel.Worksheet
                Dim horlastrow As Integer
                horairewb = XlApp.Workbooks.Open(localhoraire)
                horairewsh = horairewb.Sheets(1)
                horlastrow = horairewsh.Range("A30").End(Excel.XlDirection.xlUp).Row

                horairewsh.Cells(horlastrow, ShEw).Copy(hrly.Cells(t, 23))
                horairewsh.Cells(horlastrow, ShRR).Copy(hrly.Cells(t, 28))

                horairewb.Close(SaveChanges:=False)

                horairewsh = Nothing
                horairewb = Nothing

                releaseObject(horairewsh)
                releaseObject(horairewb)

            Else
                MsgBox("Le fichier Horaire" & fichierhoraire & " n'existe pas,verifiez le chemain d'accés", vbOKOnly)

            End If

            hrly.Cells(t, 29).value = txtdurRR.Text
            txtDD.Text = hrly.Cells(t, 5).value
            txtFF.Text = hrly.Cells(t, 6).value
            txtT.Text = hrly.Cells(t, 22).value
            txtTd.Text = hrly.Cells(t, 24).value
            txtHR.Text = hrly.Cells(t, 25).value
            txtPst.Text = hrly.Cells(t, 26).value
            txtPmer.Text = hrly.Cells(t, 27).value
            txtew.Text = hrly.Cells(t, 23).value
            txtRR.Text = hrly.Cells(t, 28).value
            crq.Close(SaveChanges:=True)
            releaseObject(hrly)
            releaseObject(crq)
            releaseObject(XlApp)

        Else
            Dim stamp_Date As String = Date.Now.ToString("dd MMM yyyy", CultureInfo.CreateSpecificCulture("fr-FRA"))
            Dim stamp_Row As String = Date.Now.ToString("ddMMyyyy", CultureInfo.CreateSpecificCulture("fr-FRA"))
            Dim Name_CRQ As String = String.Format(disck & ":\" & Folder_crq & "\" & IDstation & "crq{0}.xls", Date.Now.ToString("ddMMyyyy"))
            New_Crq(stamp_Date, Name_CRQ, stamp_Row)

            XlApp = New Excel.Application
            crq = XlApp.Workbooks.Open(nomfichier)
            hrly = crq.Worksheets("horaire")

            Dim fichierminute As String = "A_" & Date.Now.ToString("MMdd") & ".xls"
            Dim fichierhoraire As String = "S_" & Date.Now.ToString("MMdd") & ".xls"

            Dim sourceminute As String = sourceADRESS & fichierminute
            Dim localminute As String = disck & BD_crq & fichierminute
            Dim localhoraire As String = disck & BD_crq & fichierhoraire

            Try
                File.Copy(sourceADRESS & fichierminute, disck & BD_crq & fichierminute, True)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            Try
                File.Copy(sourceADRESS & fichierhoraire, disck & BD_crq & fichierhoraire, True)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            Dim t As Integer
            Dim tr As Integer = DateTime.Now.Hour
            t = tr + 4

            If System.IO.File.Exists(localminute) Then
                Dim minutewbk As Excel.Workbook
                Dim minutwsh As Excel.Worksheet
                Dim lastRow As Integer
                minutewbk = XlApp.Workbooks.Open(localminute)
                minutwsh = minutewbk.Sheets(1)
                lastRow = minutwsh.Range("A1500").End(Excel.XlDirection.xlUp).Row

                minutwsh.Cells(lastRow, SmDD).Copy(hrly.Cells(t, 5))
                minutwsh.Cells(lastRow, SmFF).Copy(hrly.Cells(t, 6))
                minutwsh.Cells(lastRow, SmT).Copy(hrly.Cells(t, 22))
                minutwsh.Cells(lastRow, SmTd).Copy(hrly.Cells(t, 24))
                minutwsh.Cells(lastRow, SmHr).Copy(hrly.Cells(t, 25))
                minutwsh.Cells(lastRow, SmPs).Copy(hrly.Cells(t, 26))
                minutwsh.Cells(lastRow, SmPm).Copy(hrly.Cells(t, 27))

                minutewbk.Close(SaveChanges:=False)

                minutwsh = Nothing
                minutewbk = Nothing

                releaseObject(minutwsh)
                releaseObject(minutewbk)

            Else
                MsgBox("Le fichier Horaire" & fichierminute & " n'existe pas,verifiez le chemain d'accés", vbOKOnly)

            End If


            If System.IO.File.Exists(localhoraire) Then

                Dim horairewb As Excel.Workbook
                Dim horairewsh As Excel.Worksheet
                Dim horlastrow As Integer
                horairewb = XlApp.Workbooks.Open(localhoraire)
                horairewsh = horairewb.Sheets(1)
                horlastrow = horairewsh.Range("A30").End(Excel.XlDirection.xlUp).Row

                horairewsh.Cells(horlastrow, ShEw).Copy(hrly.Cells(t, 23))
                horairewsh.Cells(horlastrow, ShRR).Copy(hrly.Cells(t, 28))

                horairewb.Close(SaveChanges:=False)

                horairewsh = Nothing
                horairewb = Nothing

                releaseObject(horairewsh)
                releaseObject(horairewb)

            Else
                MsgBox("Le fichier Horaire" & fichierhoraire & " n'existe pas,verifiez le chemain d'accés", vbOKOnly)

            End If

            hrly.Cells(t, 29).value = txtdurRR.Text
            txtDD.Text = hrly.Cells(t, 5).value
            txtFF.Text = hrly.Cells(t, 6).value
            txtT.Text = hrly.Cells(t, 22).value
            txtTd.Text = hrly.Cells(t, 24).value
            txtHR.Text = hrly.Cells(t, 25).value
            txtPst.Text = hrly.Cells(t, 26).value
            txtPmer.Text = hrly.Cells(t, 27).value
            txtew.Text = hrly.Cells(t, 23).value
            txtRR.Text = hrly.Cells(t, 28).value
            crq.Close(SaveChanges:=True)
            releaseObject(hrly)
            releaseObject(crq)
            releaseObject(XlApp)


        End If


    End Sub

    Private Sub horaireauto_ystrday()

        ADRESS = My.Settings.adressCAOBS
        sourceADRESS = "\\" & ADRESS & source
        Dim XlApp As Excel.Application
        Dim crq As Excel.Workbook
        Dim hrly As Excel.Worksheet
        Dim nomfichier As String = [String].Format(disck & ":\" & Folder_crq & "\" & IDstation & "crq{0}.xls", DateTime.Now.ToString("ddMMyyyy"))
        If System.IO.File.Exists(nomfichier) Then
            XlApp = New Excel.Application
            crq = XlApp.Workbooks.Open(nomfichier)
            hrly = crq.Worksheets("horaire")

            Dim fichierminute As String = "A_" & Date.Now.AddDays(-1).ToString("MMdd") & ".xls"
            Dim fichierhoraire As String = "S_" & Date.Now.AddDays(-1).ToString("MMdd") & ".xls"

            Dim sourceminute As String = sourceADRESS & fichierminute
            Dim localminute As String = disck & BD_crq & fichierminute
            Dim localhoraire As String = disck & BD_crq & fichierhoraire

            Try
                File.Copy(sourceADRESS & fichierminute, disck & BD_crq & fichierminute, True)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            Try
                File.Copy(sourceADRESS & fichierhoraire, disck & BD_crq & fichierhoraire, True)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            Dim t As Integer
            Dim tr As Integer = DateTime.Now.Hour
            t = tr + 4

            If System.IO.File.Exists(localminute) Then
                Dim minutewbk As Excel.Workbook
                Dim minutwsh As Excel.Worksheet
                Dim lastRow As Integer
                minutewbk = XlApp.Workbooks.Open(localminute)
                minutwsh = minutewbk.Sheets(1)
                lastRow = minutwsh.Range("A1500").End(Excel.XlDirection.xlUp).Row

                minutwsh.Cells(lastRow, SmDD).Copy(hrly.Cells(t, 5))
                minutwsh.Cells(lastRow, SmFF).Copy(hrly.Cells(t, 6))
                minutwsh.Cells(lastRow, SmT).Copy(hrly.Cells(t, 22))
                minutwsh.Cells(lastRow, SmTd).Copy(hrly.Cells(t, 24))
                minutwsh.Cells(lastRow, SmHr).Copy(hrly.Cells(t, 25))
                minutwsh.Cells(lastRow, SmPs).Copy(hrly.Cells(t, 26))
                minutwsh.Cells(lastRow, SmPm).Copy(hrly.Cells(t, 27))

                minutewbk.Close(SaveChanges:=False)

                minutwsh = Nothing
                minutewbk = Nothing

                releaseObject(minutwsh)
                releaseObject(minutewbk)

            Else
                MsgBox("Le fichier Horaire" & fichierminute & " n'existe pas,verifiez le chemain d'accés", vbOKOnly)

            End If


            If System.IO.File.Exists(localhoraire) Then

                Dim horairewb As Excel.Workbook
                Dim horairewsh As Excel.Worksheet
                Dim horlastrow As Integer
                horairewb = XlApp.Workbooks.Open(localhoraire)
                horairewsh = horairewb.Sheets(1)
                horlastrow = horairewsh.Range("A30").End(Excel.XlDirection.xlUp).Row

                horairewsh.Cells(horlastrow, ShEw).Copy(hrly.Cells(t, 23))
                horairewsh.Cells(horlastrow, ShRR).Copy(hrly.Cells(t, 28))

                horairewb.Close(SaveChanges:=False)

                horairewsh = Nothing
                horairewb = Nothing

                releaseObject(horairewsh)
                releaseObject(horairewb)

            Else
                MsgBox("Le fichier Horaire" & fichierhoraire & " n'existe pas,verifiez le chemain d'accés", vbOKOnly)

            End If

            hrly.Cells(t, 29).value = txtdurRR.Text
            txtDD.Text = hrly.Cells(t, 5).value
            txtFF.Text = hrly.Cells(t, 6).value
            txtT.Text = hrly.Cells(t, 22).value
            txtTd.Text = hrly.Cells(t, 24).value
            txtHR.Text = hrly.Cells(t, 25).value
            txtPst.Text = hrly.Cells(t, 26).value
            txtPmer.Text = hrly.Cells(t, 27).value
            txtew.Text = hrly.Cells(t, 23).value
            txtRR.Text = hrly.Cells(t, 28).value
            crq.Close(SaveChanges:=True)
            releaseObject(hrly)
            releaseObject(crq)
            releaseObject(XlApp)

        Else
            Dim stamp_Date As String = Date.Now.ToString("dd MMM yyyy", CultureInfo.CreateSpecificCulture("fr-FRA"))
            Dim stamp_Row As String = Date.Now.ToString("ddMMyyyy", CultureInfo.CreateSpecificCulture("fr-FRA"))
            Dim Name_CRQ As String = String.Format(disck & ":\" & Folder_crq & "\" & IDstation & "crq{0}.xls", Date.Now.ToString("ddMMyyyy"))
            New_Crq(stamp_Date, Name_CRQ, stamp_Row)

            XlApp = New Excel.Application
            crq = XlApp.Workbooks.Open(nomfichier)
            hrly = crq.Worksheets("horaire")

            Dim fichierminute As String = "A_" & Date.Now.AddDays(-1).ToString("MMdd") & ".xls"
            Dim fichierhoraire As String = "S_" & Date.Now.AddDays(-1).ToString("MMdd") & ".xls"
            Dim localminute As String = disck & BD_crq & fichierminute
            Dim localhoraire As String = disck & BD_crq & fichierhoraire

            Try
                File.Copy(sourceADRESS & fichierminute, disck & BD_crq & fichierminute, True)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            Try
                File.Copy(sourceADRESS & fichierhoraire, disck & BD_crq & fichierhoraire, True)
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            Dim t As Integer
            Dim tr As Integer = DateTime.Now.Hour
            t = tr + 4

            If System.IO.File.Exists(localminute) Then
                Dim minutewbk As Excel.Workbook
                Dim minutwsh As Excel.Worksheet
                Dim lastRow As Integer
                minutewbk = XlApp.Workbooks.Open(localminute)
                minutwsh = minutewbk.Sheets(1)
                lastRow = minutwsh.Range("A1500").End(Excel.XlDirection.xlUp).Row

                minutwsh.Cells(lastRow, SmDD).Copy(hrly.Cells(t, 5))
                minutwsh.Cells(lastRow, SmFF).Copy(hrly.Cells(t, 6))
                minutwsh.Cells(lastRow, SmT).Copy(hrly.Cells(t, 22))
                minutwsh.Cells(lastRow, SmTd).Copy(hrly.Cells(t, 24))
                minutwsh.Cells(lastRow, SmHr).Copy(hrly.Cells(t, 25))
                minutwsh.Cells(lastRow, SmPs).Copy(hrly.Cells(t, 26))
                minutwsh.Cells(lastRow, SmPm).Copy(hrly.Cells(t, 27))

                minutewbk.Close(SaveChanges:=False)
                releaseObject(minutwsh)
                releaseObject(minutewbk)

            Else
                MsgBox("Le fichier Horaire" & fichierminute & " n'existe pas,verifiez le chemain d'accés", vbOKOnly)

            End If


            If System.IO.File.Exists(localhoraire) Then

                Dim horairewb As Excel.Workbook
                Dim horairewsh As Excel.Worksheet
                Dim horlastrow As Integer
                horairewb = XlApp.Workbooks.Open(localhoraire)
                horairewsh = horairewb.Sheets(1)
                horlastrow = horairewsh.Range("A30").End(Excel.XlDirection.xlUp).Row

                horairewsh.Cells(horlastrow, ShEw).Copy(hrly.Cells(t, 23))
                horairewsh.Cells(horlastrow, ShRR).Copy(hrly.Cells(t, 28))

                horairewb.Close(SaveChanges:=False)
                releaseObject(horairewsh)
                releaseObject(horairewb)

            Else
                MsgBox("Le fichier Horaire" & fichierhoraire & " n'existe pas,verifiez le chemain d'accés", vbOKOnly)

            End If

            hrly.Cells(t, 29).value = txtdurRR.Text
            txtDD.Text = hrly.Cells(t, 5).value
            txtFF.Text = hrly.Cells(t, 6).value
            txtT.Text = hrly.Cells(t, 22).value
            txtTd.Text = hrly.Cells(t, 24).value
            txtHR.Text = hrly.Cells(t, 25).value
            txtPst.Text = hrly.Cells(t, 26).value
            txtPmer.Text = hrly.Cells(t, 27).value
            txtew.Text = hrly.Cells(t, 23).value
            txtRR.Text = hrly.Cells(t, 28).value
            crq.Close(SaveChanges:=True)
            releaseObject(hrly)
            releaseObject(crq)
            releaseObject(XlApp)


        End If

    End Sub


    Private Sub btnauto_start_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnauto_start.Click

        Timer2.Start()
        btnauto_start.Enabled = False
        btnauto_stop.Enabled = True
        lblauto.Show()

    End Sub

    Private Sub btnauto_stop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnauto_stop.Click

        Timer2.Stop()
        btnauto_stop.Enabled = False
        btnauto_start.Enabled = True
        lblauto.Hide()

    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick

        If TimeString = "01:00:30" Then
            horaireauto()

        End If

        If TimeString = "02:00:30" Then
            horaireauto()

        End If

        If TimeString = "03:00:30" Then
            horaireauto()

        End If

        If TimeString = "04:00:30" Then
            horaireauto()

        End If

        If TimeString = "05:00:30" Then
            horaireauto()

        End If

        If TimeString = "06:00:30" Then
            horaireauto()

        End If

        If TimeString = "07:00:30" Then
            horaireauto()

        End If

        If TimeString = "08:00:30" Then
            horaireauto()

        End If

        If TimeString = "09:00:30" Then
            horaireauto()

        End If

        If TimeString = "10:00:30" Then
            horaireauto()

        End If

        If TimeString = "11:00:30" Then
            horaireauto()

        End If

        If TimeString = "12:00:30" Then
            horaireauto()

        End If

        If TimeString = "13:00:30" Then
            horaireauto()

        End If

        If TimeString = "14:00:30" Then
            horaireauto()

        End If

        If TimeString = "15:00:30" Then
            horaireauto()

        End If

        If TimeString = "16:00:30" Then
            horaireauto()

        End If

        If TimeString = "17:00:30" Then
            horaireauto()

        End If

        If TimeString = "18:00:30" Then
            horaireauto()

        End If

        If TimeString = "19:00:30" Then
            horaireauto()

        End If

        If TimeString = "20:00:30" Then
            horaireauto()

        End If

        If TimeString = "21:00:30" Then
            horaireauto()

        End If

        If TimeString = "22:00:30" Then
            horaireauto()

        End If

        If TimeString = "23:00:30" Then
            horaireauto()

        End If

        If TimeString = "00:00:30" Then
            horaireauto_ystrday()

        End If



    End Sub

    Private Sub btnrecup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnrecup.Click
        frmrecup.ShowDialog()

    End Sub

    Private Sub btnConfig_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConfig.Click
        frmPASSWORD.ShowDialog()

    End Sub

    Private Sub Valid_TdH(File_Name As String, File_Date As String, Row_Date As String, Hour As Integer)

        Dim Name_CRQ As String = File_Name
        Dim Date_CRQ As String = File_Date
        Dim Date_Row As String = Row_Date
        Dim index As Integer = Hour

        If System.IO.File.Exists(Name_CRQ) Then
            Dim XlApp As Excel.Application
            Dim XlWbk As Excel.Workbook
            Dim HrlyWsh As Excel.Worksheet
            Dim PhenWsh As Excel.Worksheet
            Dim ExtrWsh As Excel.Worksheet
            XlApp = New Excel.Application
            XlWbk = XlApp.Workbooks.Open(Name_CRQ)
            HrlyWsh = XlWbk.Worksheets("horaire")
            PhenWsh = XlWbk.Worksheets("phenomenes")
            ExtrWsh = XlWbk.Worksheets("extrèmes")
            'code here
            With HrlyWsh
                .Cells(index, 4).Value = txtvv.Text
                .Cells(index, 7).Value = txtww.Text
                .Cells(index, 8).Value = txtw1w2.Text
                .Cells(index, 9).Value = txtcl.Text
                .Cells(index, 10).Value = txtcm.Text
                .Cells(index, 11).Value = txtch.Text
                .Cells(index, 12).Value = txtN.Text
                .Cells(index, 13).Value = txth1.Text
                .Cells(index, 14).Value = txtn1.Text
                .Cells(index, 15).Value = txtc1.Text
                .Cells(index, 16).Value = txth0.Text
                .Cells(index, 17).Value = txtn0.Text
                .Cells(index, 18).Value = txtc0.Text
                .Cells(index, 19).Value = txth2.Text
                .Cells(index, 20).Value = txtn2.Text
                .Cells(index, 21).Value = txtc2.Text

            End With

            MsgBox("Validation Tour d'Horizon effectuée", MsgBoxStyle.Information, vbOKOnly)
            ''
            XlWbk.Close(SaveChanges:=True)
            releaseObject(HrlyWsh)
            releaseObject(PhenWsh)
            releaseObject(ExtrWsh)
            releaseObject(XlWbk)
            releaseObject(XlApp)
        Else
            Dim XlApp As Excel.Application
            Dim XlWbk As Excel.Workbook
            Dim HrlyWsh As Excel.Worksheet
            Dim PhenWsh As Excel.Worksheet
            Dim ExtrWsh As Excel.Worksheet
            XlApp = New Excel.Application
            XlWbk = XlApp.Workbooks.Add
            HrlyWsh = XlWbk.Worksheets(1)
            PhenWsh = XlWbk.Worksheets(2)
            ExtrWsh = XlWbk.Worksheets(3)
            With XlWbk
                HrlyWsh.Name = "horaire"
                PhenWsh.Name = "phenomenes"
                ExtrWsh.Name = "extrèmes"

            End With

            '*********feille 1
            With HrlyWsh
                .Cells(3, 1).value = "Indicatif"
                .Cells(3, 1).font.bold = True
                .Cells(3, 1).font.size = 11
                .Cells(3, 2).value = "Date"
                .Cells(3, 2).font.bold = True
                .Cells(3, 2).font.size = 11
                .Cells(3, 3).value = "heure"
                .Cells(3, 3).font.italic = True
                .Cells(3, 3).font.bold = True
                .Cells(3, 3).font.size = 11
                .Cells(4, 3).value = "00H00"
                .Cells(5, 3).value = "01H00"
                .Cells(6, 3).value = "02H00"
                .Cells(7, 3).value = "03H00"
                .Cells(8, 3).value = "04H00"
                .Cells(9, 3).value = "05H00"
                .Cells(10, 3).value = "06H00"
                .Cells(11, 3).value = "07H00"
                .Cells(12, 3).value = "08H00"
                .Cells(13, 3).value = "09H00"
                .Cells(14, 3).value = "10H00"
                .Cells(15, 3).value = "11H00"
                .Cells(16, 3).value = "12H00"
                .Cells(17, 3).value = "13H00"
                .Cells(18, 3).value = "14H00"
                .Cells(19, 3).value = "15H00"
                .Cells(20, 3).value = "16H00"
                .Cells(21, 3).value = "17H00"
                .Cells(22, 3).value = "18H00"
                .Cells(23, 3).value = "19H00"
                .Cells(24, 3).value = "20H00"
                .Cells(25, 3).value = "21H00"
                .Cells(26, 3).value = "22H00"
                .Cells(27, 3).value = "23H00"
                .Cells(3, 4).value = "Visi."
                .Cells(3, 4).font.italic = True
                .Cells(3, 4).font.bold = True
                .Cells(3, 4).font.size = 11
                .Cells(3, 5).value = "DDD"
                .Cells(3, 5).font.italic = True
                .Cells(3, 5).font.bold = True
                .Cells(3, 5).font.size = 11
                .Cells(3, 6).value = "FF"
                .Cells(3, 6).font.italic = True
                .Cells(3, 6).font.bold = True
                .Cells(3, 6).font.size = 11
                .Cells(3, 7).value = "WW"
                .Cells(3, 7).font.italic = True
                .Cells(3, 7).font.bold = True
                .Cells(3, 7).font.size = 8
                .Cells(3, 8).value = "w1w2"
                .Cells(3, 8).font.italic = True
                .Cells(3, 8).font.bold = True
                .Cells(3, 8).font.size = 8
                .Cells(3, 9).value = "CL"
                .Cells(3, 9).font.italic = True
                .Cells(3, 9).font.bold = True
                .Cells(3, 9).font.size = 10
                .Cells(3, 10).value = "CM"
                .Cells(3, 10).font.italic = True
                .Cells(3, 10).font.bold = True
                .Cells(3, 10).font.size = 10
                .Cells(3, 11).value = "CH"
                .Cells(3, 11).font.italic = True
                .Cells(3, 11).font.bold = True
                .Cells(3, 11).font.size = 10
                .Cells(3, 12).value = "N"
                .Cells(3, 12).font.italic = True
                .Cells(3, 12).font.bold = True
                .Cells(3, 12).font.size = 10
                .Cells(3, 13).value = "h1"
                .Cells(3, 13).font.italic = True
                .Cells(3, 13).font.bold = True
                .Cells(3, 13).font.size = 10
                .Cells(3, 14).value = "n1"
                .Cells(3, 14).font.italic = True
                .Cells(3, 14).font.bold = True
                .Cells(3, 14).font.size = 10
                .Cells(3, 15).value = "C1"
                .Cells(3, 15).font.italic = True
                .Cells(3, 15).font.bold = True
                .Cells(3, 15).font.size = 10
                .Cells(3, 16).value = "h0"
                .Cells(3, 16).font.italic = True
                .Cells(3, 16).font.bold = True
                .Cells(3, 16).font.size = 10
                .Cells(3, 17).value = "n0"
                .Cells(3, 17).font.italic = True
                .Cells(3, 17).font.bold = True
                .Cells(3, 17).font.size = 10
                .Cells(3, 18).value = "C0"
                .Cells(3, 18).font.italic = True
                .Cells(3, 18).font.bold = True
                .Cells(3, 18).font.size = 10
                .Cells(3, 19).value = "h2"
                .Cells(3, 19).font.italic = True
                .Cells(3, 19).font.bold = True
                .Cells(3, 19).font.size = 10
                .Cells(3, 20).value = "n2"
                .Cells(3, 20).font.italic = True
                .Cells(3, 20).font.bold = True
                .Cells(3, 20).font.size = 10
                .Cells(3, 21).value = "C2"
                .Cells(3, 21).font.italic = True
                .Cells(3, 21).font.bold = True
                .Cells(3, 21).font.size = 10
                .Cells(3, 22).value = "T.air"
                .Cells(3, 22).font.italic = True
                .Cells(3, 22).font.bold = True
                .Cells(3, 22).font.size = 11
                .Cells(3, 23).value = "ew"
                .Cells(3, 23).font.italic = True
                .Cells(3, 23).font.bold = True
                .Cells(3, 23).font.size = 11
                .Cells(3, 24).value = "Td"
                .Cells(3, 24).font.italic = True
                .Cells(3, 24).font.bold = True
                .Cells(3, 24).font.size = 11
                .Cells(3, 25).value = "U%"
                .Cells(3, 25).font.italic = True
                .Cells(3, 25).font.bold = True
                .Cells(3, 25).font.size = 11
                .Cells(3, 26).value = "P.st°"
                .Cells(3, 26).font.italic = True
                .Cells(3, 26).font.bold = True
                .Cells(3, 26).font.size = 11
                .Cells(3, 27).value = "P.mer"
                .Cells(3, 27).font.italic = True
                .Cells(3, 27).font.bold = True
                .Cells(3, 27).font.size = 11
                .Cells(3, 28).value = "RR"
                .Cells(3, 28).font.italic = True
                .Cells(3, 28).font.bold = True
                .Cells(3, 28).font.size = 11
                .Cells(3, 29).value = "dur.RR"
                .Cells(3, 29).font.italic = True
                .Cells(3, 29).font.bold = True
                .Cells(3, 29).font.size = 11
                .Columns("C:C").ColumnWidth = 5.43
                .Columns("D:D").ColumnWidth = 7
                .Columns("E:E").ColumnWidth = 6
                .Columns("F:F").ColumnWidth = 4
                .Columns("G:G").ColumnWidth = 3.86
                .Columns("H:H").ColumnWidth = 4.86
                .Columns("I:I").ColumnWidth = 2.14
                .Columns("J:J").ColumnWidth = 2.4
                .Columns("K:K").ColumnWidth = 2.29
                .Columns("L:L").ColumnWidth = 2
                .Columns("M:M").ColumnWidth = 6
                .Columns("N:N").ColumnWidth = 2.14
                .Columns("O:O").ColumnWidth = 2.14
                .Columns("P:P").ColumnWidth = 6
                .Columns("Q:Q").ColumnWidth = 2.14
                .Columns("R:R").ColumnWidth = 2.14
                .Columns("S:S").ColumnWidth = 6
                .Columns("T:T").ColumnWidth = 2.14
                .Columns("U:U").ColumnWidth = 2.14
                .Columns("V:V").ColumnWidth = 4.7
                .Columns("W:W").ColumnWidth = 4
                .Columns("X:X").ColumnWidth = 4.7
                .Columns("Y:Y").ColumnWidth = 4
                .Columns("Z:Z").ColumnWidth = 7
                .Columns("AA:AA").ColumnWidth = 7
                .Columns("AB:AB").ColumnWidth = 5
                .Columns("AC:AC").ColumnWidth = 7
                .Cells.NumberFormat = "@"
                For i As Integer = 4 To 27
                    .Cells(i, 1).value = IDstation
                    .Cells(i, 2).value = Date_Row
                Next

                .Cells(1, 2).value = "Date:"
                .Cells(1, 2).font.italic = True
                .Cells(1, 2).font.bold = True
                .Cells(1, 2).font.size = 11
                .Range("C1:E1").Merge()
                .Cells(1, 3).value = Date_CRQ
                .Cells(1, 3).font.bold = True
                .Range("C1:E1").HorizontalAlignment = Excel.Constants.xlCenter
                .Range("G2:I2").Merge()
                .Cells(2, 7).value = "Station:"
                .Cells(2, 7).font.bold = True
                .Cells(2, 7).font.italic = True
                .Range("G2:I2").HorizontalAlignment = Excel.Constants.xlCenter
                .Range("J2:L2").Merge()
                .Cells(2, 10).value = IDstation
                .Cells(2, 10).font.bold = True
                .Range("J2:L2").HorizontalAlignment = Excel.Constants.xlCenter


            End With
            '*********feuille 2
            With PhenWsh
                .Cells(1, 1).value = "Indicatif"
                .Cells(1, 2).value = "Date"
                .Cells(1, 3).value = "Phénomène"   'mise en forme sheet-phenomenes
                .Cells(1, 4).value = "Code"
                .Cells(1, 5).value = "H.début"
                .Cells(1, 6).value = "H.fin"
                .Cells(1, 7).value = "Intensité"
                .Cells(1, 8).value = "Secteur"
                .Cells(1, 9).value = "Visi.mini"
                .Cells(1, 10).value = "heure"
                .Cells(1, 11).value = "Hauteur"
                .Rows("1:1").Font.Bold = True
                .Rows("1:1").Font.size = 10
                .Columns("A:A").ColumnWidth = 6
                .Columns("B:B").ColumnWidth = 10
                .Columns("C:C").ColumnWidth = 12
                .Columns("D:D").ColumnWidth = 5
                .Columns("E:E").ColumnWidth = 8
                .Columns("F:F").ColumnWidth = 8
                .Columns("G:G").ColumnWidth = 8
                .Columns("H:H").ColumnWidth = 8
                .Columns("I:I").ColumnWidth = 8
                .Columns("J:J").ColumnWidth = 8
                .Columns("K:K").ColumnWidth = 8
                .Columns("L:L").ColumnWidth = 8
                .Columns("M:M").ColumnWidth = 8
                .Columns("N:N").ColumnWidth = 8
                .Columns("O:O").ColumnWidth = 8
                .Columns("P:P").ColumnWidth = 8
                .Cells.NumberFormat = "@"
            End With
            '*********feuille 3
            With ExtrWsh
                .Cells.NumberFormat = "@"

            End With
            With HrlyWsh
                .Cells(index, 4).Value = txtvv.Text
                .Cells(index, 7).Value = txtww.Text
                .Cells(index, 8).Value = txtw1w2.Text
                .Cells(index, 9).Value = txtcl.Text
                .Cells(index, 10).Value = txtcm.Text
                .Cells(index, 11).Value = txtch.Text
                .Cells(index, 12).Value = txtN.Text
                .Cells(index, 13).Value = txth1.Text
                .Cells(index, 14).Value = txtn1.Text
                .Cells(index, 15).Value = txtc1.Text
                .Cells(index, 16).Value = txth0.Text
                .Cells(index, 17).Value = txtn0.Text
                .Cells(index, 18).Value = txtc0.Text
                .Cells(index, 19).Value = txth2.Text
                .Cells(index, 20).Value = txtn2.Text
                .Cells(index, 21).Value = txtc2.Text

            End With
            XlWbk.SaveAs(Name_CRQ)
            XlWbk.Close(SaveChanges:=True)
            XlApp = Nothing
            XlWbk = Nothing
            HrlyWsh = Nothing
            PhenWsh = Nothing
            ExtrWsh = Nothing

            releaseObject(HrlyWsh)
            releaseObject(PhenWsh)
            releaseObject(ExtrWsh)
            releaseObject(XlWbk)
            releaseObject(XlApp)
            MsgBox("le CRQ du:" & File_Date & " est crée, Tdh est validé!", MsgBoxStyle.Information, vbOKOnly)
        End If

    End Sub
    Private Sub DataGridView1_CellMouseUp(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseUp
        If e.Button = MouseButtons.Right Then
            If e.RowIndex > -1 Then
                Me.DataGridView1.Rows(e.RowIndex).Selected = True
                Me.DataGridView1.CurrentCell = Me.DataGridView1.Rows(e.RowIndex).Cells(1)
                Me.ContextMenuStrip1.Show(Me.DataGridView1, e.Location)
                ContextMenuStrip1.Show(Cursor.Position)
            End If
        End If
    End Sub
    Private Sub ContextMenuStrip1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ContextMenuStrip1.Click

        If Not Me.DataGridView1.Rows(DataGridView1.CurrentRow.Index).IsNewRow Then
            Dim Name_CRQ As String = String.Format(disck & ":\" & Folder_crq & "\" & IDstation & "crq{0}.xls", Date.Now.ToString("ddMMyyyy"))
            Dim index As Integer = DataGridView1.CurrentRow.Index + 2
            Dim XlApp As Excel.Application
            Dim XlWbk As Excel.Workbook
            Dim PhenWsh As Excel.Worksheet
            XlApp = New Excel.Application
            XlWbk = XlApp.Workbooks.Open(Name_CRQ)
            PhenWsh = XlWbk.Worksheets("phenomenes")

            Cor_Phen_Index = index
            cbophen.Text = PhenWsh.Cells(index, 3).Value
            txtcode.Text = PhenWsh.Cells(index, 4).Value
            txtstart.Text = PhenWsh.Cells(index, 5).Value
            txtend.Text = PhenWsh.Cells(index, 6).Value
            cbointens.SelectedItem = PhenWsh.Cells(index, 7).Value
            cbosctr.SelectedItem = PhenWsh.Cells(index, 8).Value
            txtvvmin.Text = PhenWsh.Cells(index, 9).Value
            txthvvmin.Text = PhenWsh.Cells(index, 10).Value
            txthaut.Text = PhenWsh.Cells(index, 11).Value

            XlWbk.Close(SaveChanges:=True)
            releaseObject(PhenWsh)
            releaseObject(XlWbk)
            releaseObject(XlApp)
            btncorrige.Show()
            btnphvalid.Hide()

        End If

    End Sub
    Private Sub Make_SixMn_csv(Hour As Integer)

        ADRESS = My.Settings.adressCAOBS
        sourceADRESS = "\\" & ADRESS & source
        If ADRESS <> "" Then
            Dim csvRepertory As String = disck & BD_crq & "SmnCsv\"
            Dim intHR As Integer = Hour
            Dim csvFile As String = IDstation & "smn" & Date.Now.ToString("ddMMyyyy") & intHR.ToString & ".txt"
            Dim csvFILE_Path As String = csvRepertory & csvFile
            Dim fichierminute As String = "A_" & Date.Now.ToString("MMdd") & ".xls"
            Dim fichierSixMn As String = "M_" & Date.Now.ToString("MMdd") & ".xls"

            Dim localminute As String = disck & BD_crq & fichierminute
            Dim localSixMn As String = disck & BD_crq & fichierSixMn

            If System.IO.File.Exists(sourceADRESS & fichierminute) Then
                Try

                    File.Copy(sourceADRESS & fichierminute, localminute, True)

                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try

            End If

            If System.IO.File.Exists(sourceADRESS & fichierSixMn) Then
                Try

                    File.Copy(sourceADRESS & fichierSixMn, localSixMn, True)

                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try

            End If

            If System.IO.File.Exists(localminute) And System.IO.File.Exists(localSixMn) Then
                Dim app As Excel.Application
                Dim SixWb As Excel.Workbook
                Dim MntWb As Excel.Workbook
                Dim SixWsh As Excel.Worksheet
                Dim MntWsh As Excel.Worksheet
                app = New Excel.Application
                SixWb = app.Workbooks.Open(localSixMn)
                SixWsh = SixWb.Worksheets(1)
                MntWb = app.Workbooks.Open(localminute)
                MntWsh = MntWb.Worksheets(1)
                Dim SixMnLastRow As Integer = (intHR * 10) + 1
                Dim MntLastRow As Integer = (intHR * 60) + 1
                Dim SixMnFirstRow As Integer = SixMnLastRow - 10
                Dim MntFirstRow As Integer = MntLastRow - 60
                Dim CSV_array(9, 9) As String
                Dim indice As Integer
                Dim indice2 As Integer
                'datetime format YYYY-MM-DD hh:mm:ss
                '02/10/2020 07:06:00

                For i As Integer = 0 To 9
                    indice = 9 - i
                    CSV_array(i, 0) = Date.Now.ToString("yyyy-MM-dd") & " " & (CStr(SixWsh.Cells(SixMnLastRow - indice, 1).Value)).Substring(11, 8) 'date & time*
                    CSV_array(i, 7) = CStr(SixWsh.Cells(SixMnLastRow - indice, MsRR).Value) 'pluie
                    CSV_array(i, 8) = CStr(SixWsh.Cells(SixMnLastRow - indice, MsDI).Value) 'DI
                    CSV_array(i, 9) = CStr(SixWsh.Cells(SixMnLastRow - indice, MsRG).Value) 'RG
                Next

                For i As Integer = 1 To 10
                    indice2 = i * 6
                    CSV_array(i - 1, 1) = CStr(MntWsh.Cells(MntFirstRow + indice2, SmDD).Value) 'DD
                    CSV_array(i - 1, 2) = CStr(MntWsh.Cells(MntFirstRow + indice2, SmFF).Value) 'FF
                    CSV_array(i - 1, 3) = CStr(MntWsh.Cells(MntFirstRow + indice2, SmT).Value) 'T
                    CSV_array(i - 1, 4) = CStr(MntWsh.Cells(MntFirstRow + indice2, SmHr).Value) 'HR
                    CSV_array(i - 1, 5) = CStr(MntWsh.Cells(MntFirstRow + indice2, SmTd).Value) 'Td
                    CSV_array(i - 1, 6) = CStr(MntWsh.Cells(MntFirstRow + indice2, SmPm).Value) 'Pmer

                Next

                '****closing *****
                SixWb.Close(SaveChanges:=False)
                MntWb.Close(SaveChanges:=False)
                SixWsh = Nothing
                SixWb = Nothing
                MntWsh = Nothing
                MntWb = Nothing
                releaseObject(SixWb)
                releaseObject(SixWsh)
                releaseObject(MntWb)
                releaseObject(MntWsh)
                '****Closed*********

                If File.Exists(csvFILE_Path) Then
                    File.Delete(csvFILE_Path)
                End If

                Dim filestream As New FileStream(csvFILE_Path, FileMode.Create, FileAccess.Write)
                Dim Swriter As StreamWriter
                Swriter = New StreamWriter(filestream)
                For i As Integer = 0 To 9
                    Swriter.Write((CSV_array(i, 0).ToString) + ",")
                    For j As Integer = 1 To 8
                        If IsNumeric(CSV_array(i, j)) Then
                            Swriter.Write((CSV_array(i, j).ToString) + ",")
                        Else
                            Swriter.Write(("") + ",")
                        End If
                    Next
                    If IsNumeric(CSV_array(i, 9)) Then
                        Swriter.Write((CSV_array(i, 9).ToString))
                    Else
                        Swriter.Write("")

                    End If
                    Swriter.WriteLine()

                Next
                '***********
                Swriter.Close()
                Swriter.Dispose()
                Swriter = Nothing
                'MsgBox("file created" & csvFILE, vbOKOnly)
                Array.Clear(CSV_array, 0, CSV_array.Length)
                plotArqus_Send(csvFile, csvFILE_Path)

            End If
        End If

    End Sub
    Private Sub Make_Hourly_csv(HR As Integer)
        ADRESS = My.Settings.adressCAOBS
        sourceADRESS = "\\" & ADRESS & source

        If ADRESS <> "" Then
            Dim csvRepertory As String = disck & BD_crq & "HorCsv\"
            Dim intHR As Integer = HR
            Dim csvFile As String = IDstation & "hrl" & Date.Now.ToString("ddMMyyyy") & intHR.ToString & ".txt"
            Dim csvFILE_Path As String = csvRepertory & csvFile
            Dim fichierHoraire As String = "S_" & Date.Now.ToString("MMdd") & ".xls"
            Dim localHoraire As String = disck & BD_crq & fichierHoraire
            If System.IO.File.Exists(sourceADRESS & fichierHoraire) Then
                Try
                    File.Copy(sourceADRESS & fichierHoraire, localHoraire, True)

                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try
            End If

            If System.IO.File.Exists(localHoraire) Then
                Dim app As Excel.Application
                Dim HRWb As Excel.Workbook
                Dim HRWsh As Excel.Worksheet
                app = New Excel.Application
                HRWb = app.Workbooks.Open(localHoraire)
                HRWsh = HRWb.Worksheets(1)
                Dim HrlyRow As Integer = intHR + 1
                Dim CSV_array(14) As String
                'datetime format YYYY-MM-DD hh:mm:ss
                '02/10/2020 07:06:00
                CSV_array(0) = Date.Now.ToString("yyyy-MM-dd") & " " & (CStr(HRWsh.Cells(HrlyRow, 1).Value)).Substring(11, 8)
                CSV_array(1) = CStr(HRWsh.Cells(HrlyRow, HrWnd).Value)
                CSV_array(2) = CStr(HRWsh.Cells(HrlyRow, HrFFWnd).Value)
                CSV_array(3) = CStr(HRWsh.Cells(HrlyRow, HrhWnd).Value)
                CSV_array(4) = CStr(HRWsh.Cells(HrlyRow, HrMinT).Value)
                CSV_array(5) = CStr(HRWsh.Cells(HrlyRow, HrhMinT).Value)
                CSV_array(6) = CStr(HRWsh.Cells(HrlyRow, HrMxT).Value)
                CSV_array(7) = CStr(HRWsh.Cells(HrlyRow, HrhMxT).Value)
                CSV_array(8) = CStr(HRWsh.Cells(HrlyRow, HrMinU).Value)
                CSV_array(9) = CStr(HRWsh.Cells(HrlyRow, HrhMinU).Value)
                CSV_array(10) = CStr(HRWsh.Cells(HrlyRow, HrMaxU).Value)
                CSV_array(11) = CStr(HRWsh.Cells(HrlyRow, HrhMaxU).Value)
                CSV_array(12) = CStr(HRWsh.Cells(HrlyRow, HrRR).Value)
                CSV_array(13) = CStr(HRWsh.Cells(HrlyRow, HrDi).Value)
                CSV_array(14) = CStr(HRWsh.Cells(HrlyRow, HrRg).Value)

                '****closing *****
                HRWb.Close(SaveChanges:=False)
                HRWsh = Nothing
                HRWb = Nothing
                releaseObject(HRWb)
                releaseObject(HRWsh)
                '****Closed*********

                If File.Exists(csvFILE_Path) Then
                    File.Delete(csvFILE_Path)
                End If

                Dim filestream As New FileStream(csvFILE_Path, FileMode.Create, FileAccess.Write)

                Dim Swriter As StreamWriter
                Swriter = New StreamWriter(filestream)

                Swriter.Write((CSV_array(0).ToString) + ",")
                For j As Integer = 1 To 13
                    If IsNumeric(CSV_array(j)) Then
                        Swriter.Write((CSV_array(j).ToString) + ",")
                    Else
                        Swriter.Write(("") + ",")

                    End If

                Next

                If IsNumeric(CSV_array(14)) Then
                    Swriter.Write((CSV_array(14).ToString))
                Else
                    Swriter.Write("")

                End If

                Swriter.WriteLine()

                Swriter.Close()
                Swriter.Dispose()
                Swriter = Nothing
                Array.Clear(CSV_array, 0, CSV_array.Length)
                plotArqus_Send(csvFile, csvFILE_Path)

            End If

        End If

    End Sub

    Private Sub TimerSender_Tick(sender As Object, e As EventArgs) Handles TimerSender.Tick
        ' à faire la copmpilation de 00H D-1
        If TimeString = "00:01:00" Then
            Make_Smn_Csv_Ystrday(24)
            Make_Hrl_Csv_Ystrday(24)
        End If
        If TimeString = "01:01:00" Then
            Dim Hour_Now As Integer = Date.Now.Hour
            Make_SixMn_csv(Hour_Now)
            Make_Hourly_csv(Hour_Now)
        End If
        If TimeString = "02:01:00" Then
            Dim Hour_Now As Integer = Date.Now.Hour
            Make_SixMn_csv(Hour_Now)
            Make_Hourly_csv(Hour_Now)
        End If
        If TimeString = "03:01:00" Then
            Dim Hour_Now As Integer = Date.Now.Hour
            Make_SixMn_csv(Hour_Now)
            Make_Hourly_csv(Hour_Now)
        End If
        If TimeString = "04:01:00" Then
            Dim Hour_Now As Integer = Date.Now.Hour
            Make_SixMn_csv(Hour_Now)
            Make_Hourly_csv(Hour_Now)
        End If
        If TimeString = "05:01:00" Then
            Dim Hour_Now As Integer = Date.Now.Hour
            Make_SixMn_csv(Hour_Now)
            Make_Hourly_csv(Hour_Now)
        End If
        If TimeString = "06:01:00" Then
            Dim Hour_Now As Integer = Date.Now.Hour
            Make_SixMn_csv(Hour_Now)
            Make_Hourly_csv(Hour_Now)
        End If
        If TimeString = "07:01:00" Then
            Dim Hour_Now As Integer = Date.Now.Hour
            Make_SixMn_csv(Hour_Now)
            Make_Hourly_csv(Hour_Now)
        End If
        If TimeString = "08:01:00" Then
            Dim Hour_Now As Integer = Date.Now.Hour
            Make_SixMn_csv(Hour_Now)
            Make_Hourly_csv(Hour_Now)
        End If
        If TimeString = "09:01:00" Then
            Dim Hour_Now As Integer = Date.Now.Hour
            Make_SixMn_csv(Hour_Now)
            Make_Hourly_csv(Hour_Now)
        End If
        If TimeString = "10:01:00" Then
            Dim Hour_Now As Integer = Date.Now.Hour
            Make_SixMn_csv(Hour_Now)
            Make_Hourly_csv(Hour_Now)
        End If
        If TimeString = "11:01:00" Then
            Dim Hour_Now As Integer = Date.Now.Hour
            Make_SixMn_csv(Hour_Now)
            Make_Hourly_csv(Hour_Now)
        End If
        If TimeString = "12:01:00" Then
            Dim Hour_Now As Integer = Date.Now.Hour
            Make_SixMn_csv(Hour_Now)
            Make_Hourly_csv(Hour_Now)
        End If
        If TimeString = "13:01:00" Then
            Dim Hour_Now As Integer = Date.Now.Hour
            Make_SixMn_csv(Hour_Now)
            Make_Hourly_csv(Hour_Now)
        End If
        If TimeString = "14:01:00" Then
            Dim Hour_Now As Integer = Date.Now.Hour
            Make_SixMn_csv(Hour_Now)
            Make_Hourly_csv(Hour_Now)
        End If
        If TimeString = "15:01:00" Then
            Dim Hour_Now As Integer = Date.Now.Hour
            Make_SixMn_csv(Hour_Now)
            Make_Hourly_csv(Hour_Now)
        End If
        If TimeString = "16:01:00" Then
            Dim Hour_Now As Integer = Date.Now.Hour
            Make_SixMn_csv(Hour_Now)
            Make_Hourly_csv(Hour_Now)
        End If
        If TimeString = "17:01:00" Then
            Dim Hour_Now As Integer = Date.Now.Hour
            Make_SixMn_csv(Hour_Now)
            Make_Hourly_csv(Hour_Now)
        End If
        If TimeString = "18:01:00" Then
            Dim Hour_Now As Integer = Date.Now.Hour
            Make_SixMn_csv(Hour_Now)
            Make_Hourly_csv(Hour_Now)
        End If
        If TimeString = "19:01:00" Then
            Dim Hour_Now As Integer = Date.Now.Hour
            Make_SixMn_csv(Hour_Now)
            Make_Hourly_csv(Hour_Now)
        End If
        If TimeString = "20:01:00" Then
            Dim Hour_Now As Integer = Date.Now.Hour
            Make_SixMn_csv(Hour_Now)
            Make_Hourly_csv(Hour_Now)
        End If
        If TimeString = "21:01:00" Then
            Dim Hour_Now As Integer = Date.Now.Hour
            Make_SixMn_csv(Hour_Now)
            Make_Hourly_csv(Hour_Now)
        End If
        If TimeString = "22:01:00" Then
            Dim Hour_Now As Integer = Date.Now.Hour
            Make_SixMn_csv(Hour_Now)
            Make_Hourly_csv(Hour_Now)
        End If
        If TimeString = "23:01:00" Then
            Dim Hour_Now As Integer = Date.Now.Hour
            Make_SixMn_csv(Hour_Now)
            Make_Hourly_csv(Hour_Now)
        End If
    End Sub
    Private Sub plotArqus_Send(fileName As String, FilePath As String)
        Try
            Dim client As WebClient = New WebClient
            Dim FTPadress As String = My.Settings.PlotArqus.ToString
            Dim remotePath As String = My.Settings.remotePath.ToString '/remote/path/
            Dim Usr As String = My.Settings.ftpUser.ToString
            Dim Pswd As String = My.Settings.ftpPass.ToString
            Dim File_name As String = fileName
            Dim localPath As String = FilePath
            client.Credentials = New NetworkCredential(Usr, Pswd)
            client.UploadFile("ftp://" & FTPadress & remotePath & File_name, localPath)

        Catch ex As Exception

        End Try

    End Sub
    Private Sub Make_Smn_Csv_Ystrday(Hour As Integer)
        ADRESS = My.Settings.adressCAOBS
        sourceADRESS = "\\" & ADRESS & source
        If ADRESS <> "" Then
            Dim csvRepertory As String = disck & BD_crq & "SmnCsv\"
            Dim intHR As Integer = Hour
            Dim csvFile As String = IDstation & "smn" & Date.Now.AddDays(-1).ToString("ddMMyyyy") & intHR.ToString & ".txt"
            Dim csvFILE_Path As String = csvRepertory & csvFile
            Dim fichierminute As String = "A_" & Date.Now.AddDays(-1).ToString("MMdd") & ".xls"
            Dim fichierSixMn As String = "M_" & Date.Now.AddDays(-1).ToString("MMdd") & ".xls"

            Dim localminute As String = disck & BD_crq & fichierminute
            Dim localSixMn As String = disck & BD_crq & fichierSixMn

            If System.IO.File.Exists(sourceADRESS & fichierminute) Then
                Try

                    File.Copy(sourceADRESS & fichierminute, localminute, True)

                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try

            End If

            If System.IO.File.Exists(sourceADRESS & fichierSixMn) Then
                Try

                    File.Copy(sourceADRESS & fichierSixMn, localSixMn, True)

                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try

            End If

            If System.IO.File.Exists(localminute) And System.IO.File.Exists(localSixMn) Then
                Dim app As Excel.Application
                Dim SixWb As Excel.Workbook
                Dim MntWb As Excel.Workbook
                Dim SixWsh As Excel.Worksheet
                Dim MntWsh As Excel.Worksheet
                app = New Excel.Application
                SixWb = app.Workbooks.Open(localSixMn)
                SixWsh = SixWb.Worksheets(1)
                MntWb = app.Workbooks.Open(localminute)
                MntWsh = MntWb.Worksheets(1)
                Dim SixMnLastRow As Integer = (intHR * 10) + 1
                Dim MntLastRow As Integer = (intHR * 60) + 1
                Dim SixMnFirstRow As Integer = SixMnLastRow - 10
                Dim MntFirstRow As Integer = MntLastRow - 60
                Dim CSV_array(9, 9) As String
                Dim indice As Integer
                Dim indice2 As Integer
                'datetime format YYYY-MM-DD hh:mm:ss
                '02/10/2020 07:06:00

                For i As Integer = 0 To 9
                    indice = 9 - i
                    CSV_array(i, 0) = Date.Now.AddDays(-1).ToString("yyyy-MM-dd") & " " & (CStr(SixWsh.Cells(SixMnLastRow - indice, 1).Value)).Substring(11, 8) 'date & time*
                    CSV_array(i, 7) = CStr(SixWsh.Cells(SixMnLastRow - indice, MsRR).Value) 'pluie
                    CSV_array(i, 8) = CStr(SixWsh.Cells(SixMnLastRow - indice, MsDI).Value) 'DI
                    CSV_array(i, 9) = CStr(SixWsh.Cells(SixMnLastRow - indice, MsRG).Value) 'RG
                Next

                For i As Integer = 1 To 10
                    indice2 = i * 6
                    CSV_array(i - 1, 1) = CStr(MntWsh.Cells(MntFirstRow + indice2, SmDD).Value) 'DD
                    CSV_array(i - 1, 2) = CStr(MntWsh.Cells(MntFirstRow + indice2, SmFF).Value) 'FF
                    CSV_array(i - 1, 3) = CStr(MntWsh.Cells(MntFirstRow + indice2, SmT).Value) 'T
                    CSV_array(i - 1, 4) = CStr(MntWsh.Cells(MntFirstRow + indice2, SmHr).Value) 'HR
                    CSV_array(i - 1, 5) = CStr(MntWsh.Cells(MntFirstRow + indice2, SmTd).Value) 'Td
                    CSV_array(i - 1, 6) = CStr(MntWsh.Cells(MntFirstRow + indice2, SmPm).Value) 'Pmer

                Next

                '****closing *****
                SixWb.Close(SaveChanges:=False)
                MntWb.Close(SaveChanges:=False)
                SixWsh = Nothing
                SixWb = Nothing
                MntWsh = Nothing
                MntWb = Nothing
                releaseObject(SixWb)
                releaseObject(SixWsh)
                releaseObject(MntWb)
                releaseObject(MntWsh)
                '****Closed*********

                If File.Exists(csvFILE_Path) Then
                    File.Delete(csvFILE_Path)
                End If

                Dim filestream As New FileStream(csvFILE_Path, FileMode.Create, FileAccess.Write)
                Dim Swriter As StreamWriter
                Swriter = New StreamWriter(filestream)
                For i As Integer = 0 To 9
                    Swriter.Write((CSV_array(i, 0).ToString) + ",")
                    For j As Integer = 1 To 8
                        If IsNumeric(CSV_array(i, j)) Then
                            Swriter.Write((CSV_array(i, j).ToString) + ",")
                        Else
                            Swriter.Write(("") + ",")
                        End If
                    Next
                    If IsNumeric(CSV_array(i, 9)) Then
                        Swriter.Write((CSV_array(i, 9).ToString))
                    Else
                        Swriter.Write("")

                    End If
                    Swriter.WriteLine()

                Next
                '***********
                Swriter.Close()
                Swriter.Dispose()
                Swriter = Nothing
                Array.Clear(CSV_array, 0, CSV_array.Length)
                plotArqus_Send(csvFile, csvFILE_Path)
            End If
        End If

    End Sub
    Private Sub Make_Hrl_Csv_Ystrday(HR As Integer)
        ADRESS = My.Settings.adressCAOBS
        sourceADRESS = "\\" & ADRESS & source

        If ADRESS <> "" Then
            Dim csvRepertory As String = disck & BD_crq & "HorCsv\"
            Dim intHR As Integer = HR
            Dim csvFile As String = IDstation & "hrl" & Date.Now.AddDays(-1).ToString("ddMMyyyy") & intHR.ToString & ".txt"
            Dim csvFILE_Path As String = csvRepertory & csvFile
            Dim fichierHoraire As String = "S_" & Date.Now.AddDays(-1).ToString("MMdd") & ".xls"
            Dim localHoraire As String = disck & BD_crq & fichierHoraire
            If System.IO.File.Exists(sourceADRESS & fichierHoraire) Then
                Try
                    File.Copy(sourceADRESS & fichierHoraire, localHoraire, True)

                Catch ex As Exception
                    MsgBox(ex.ToString)
                End Try
            End If

            If System.IO.File.Exists(localHoraire) Then
                Dim app As Excel.Application
                Dim HRWb As Excel.Workbook
                Dim HRWsh As Excel.Worksheet
                app = New Excel.Application
                HRWb = app.Workbooks.Open(localHoraire)
                HRWsh = HRWb.Worksheets(1)
                Dim HrlyRow As Integer = intHR + 1
                Dim CSV_array(14) As String
                'datetime format YYYY-MM-DD hh:mm:ss
                '02/10/2020 07:06:00
                CSV_array(0) = Date.Now.AddDays(-1).ToString("yyyy-MM-dd") & " " & (CStr(HRWsh.Cells(HrlyRow, 1).Value)).Substring(11, 8)
                CSV_array(1) = CStr(HRWsh.Cells(HrlyRow, HrWnd).Value)
                CSV_array(2) = CStr(HRWsh.Cells(HrlyRow, HrFFWnd).Value)
                CSV_array(3) = CStr(HRWsh.Cells(HrlyRow, HrhWnd).Value)
                CSV_array(4) = CStr(HRWsh.Cells(HrlyRow, HrMinT).Value)
                CSV_array(5) = CStr(HRWsh.Cells(HrlyRow, HrhMinT).Value)
                CSV_array(6) = CStr(HRWsh.Cells(HrlyRow, HrMxT).Value)
                CSV_array(7) = CStr(HRWsh.Cells(HrlyRow, HrhMxT).Value)
                CSV_array(8) = CStr(HRWsh.Cells(HrlyRow, HrMinU).Value)
                CSV_array(9) = CStr(HRWsh.Cells(HrlyRow, HrhMinU).Value)
                CSV_array(10) = CStr(HRWsh.Cells(HrlyRow, HrMaxU).Value)
                CSV_array(11) = CStr(HRWsh.Cells(HrlyRow, HrhMaxU).Value)
                CSV_array(12) = CStr(HRWsh.Cells(HrlyRow, HrRR).Value)
                CSV_array(13) = CStr(HRWsh.Cells(HrlyRow, HrDi).Value)
                CSV_array(14) = CStr(HRWsh.Cells(HrlyRow, HrRg).Value)

                '****closing *****
                HRWb.Close(SaveChanges:=False)
                HRWsh = Nothing
                HRWb = Nothing
                releaseObject(HRWb)
                releaseObject(HRWsh)
                '****Closed*********

                If File.Exists(csvFILE_Path) Then
                    File.Delete(csvFILE_Path)
                End If

                Dim filestream As New FileStream(csvFILE_Path, FileMode.Create, FileAccess.Write)

                Dim Swriter As StreamWriter
                Swriter = New StreamWriter(filestream)

                Swriter.Write((CSV_array(0).ToString) + ",")
                For j As Integer = 1 To 13
                    If IsNumeric(CSV_array(j)) Then
                        Swriter.Write((CSV_array(j).ToString) + ",")
                    Else
                        Swriter.Write(("") + ",")

                    End If

                Next

                If IsNumeric(CSV_array(14)) Then
                    Swriter.Write((CSV_array(14).ToString))
                Else
                    Swriter.Write("")

                End If

                Swriter.WriteLine()

                Swriter.Close()
                Swriter.Dispose()
                Swriter = Nothing
                'MsgBox("file created" & csvFILE_Path, vbOKOnly)
                Array.Clear(CSV_array, 0, CSV_array.Length)
                plotArqus_Send(csvFile, csvFILE_Path)
            End If

        End If

    End Sub

End Class
