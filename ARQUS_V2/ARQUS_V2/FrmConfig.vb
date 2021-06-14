Imports Microsoft.Office.Interop
Imports System.IO
Imports Microsoft.Office.Interop.Excel
Public Class FrmConfig
    Private Sub btnSaveConfig_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveConfig.Click
        My.Settings.adressCAOBS = txtAdress.Text
        My.Settings.disck = txtDisck.Text
        My.Settings.IDst = txtIDstation.Text

        My.Settings.SmDD = CInt(txtIndexVnt.Text)
        My.Settings.SmT = CInt(txtIndexT.Text)
        My.Settings.SmHR = CInt(txtIndexHR.Text)
        My.Settings.SmPs = CInt(txtIndexPs.Text)
        My.Settings.ShEw = CInt(txtIndexEw.Text)
        My.Settings.ShRR = CInt(txtIndexRRhr.Text)

        My.Settings.KTx = CInt(txtindexMaxT.Text)
        My.Settings.KTn = CInt(txtIndexMinT.Text)
        My.Settings.KUn = CInt(txtIndexMinU.Text)
        My.Settings.KUx = CInt(txtIndexMaxU.Text)
        My.Settings.KIs = CInt(txtIndexIns.Text)
        My.Settings.KRg = CInt(txtIndexRG.Text)
        My.Settings.KRR = CInt(txtIndexRR.Text)
        My.Settings.KTnSol = CInt(txtindexTsol.Text)
        My.Settings.MsRR = CInt(txtIndex_RR_sixMn.Text)
        My.Settings.MsDI = CInt(txtIndex_Di_sixMn.Text)
        My.Settings.MsRG = CInt(txtIndex_RG_sixMn.Text)

        My.Settings.HrWnd = CInt(txtIndex_WND_Hor.Text)
        My.Settings.HrMinT = CInt(txtIndex_MinT_Hor.Text)
        My.Settings.HrMxT = CInt(txtIndex_MaxT_HOR.Text)
        My.Settings.HrMaxU = CInt(txt_IndexMax_Hr_HOR.Text)
        My.Settings.HrMinU = CInt(txtIndex_MinHr_HOR.Text)
        My.Settings.HrRR = CInt(txtIndexRR_HOR.Text)
        My.Settings.HrDi = CInt(txtIndexDi_HOR.Text)
        My.Settings.HrRg = CInt(txtIndexRG_HOR.Text)
        My.Settings.PlotArqus = txtPlotArqus.Text
        My.Settings.remotePath = txtRemotePath.Text
        My.Settings.ftpUser = txtFtpUser.Text
        My.Settings.ftpPass = txtFtpPasswd.Text
        My.Settings.Save()
        MsgBox("Configuration suavgardée", MsgBoxStyle.Information, vbOKOnly)

    End Sub

    Private Sub FrmConfig_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        saisietdh.Hide()

        My.Settings.Reload() 'For check
        If My.Settings.autoSend = False Then
            lblPlotArqusActivation.Hide()
            rdStopSend.Checked = True
            rdStartSend.Checked = False
        ElseIf My.Settings.autoSend = True Then
            lblPlotArqusActivation.Show()
            rdStartSend.Checked = True
            rdStopSend.Checked = False
        End If
        txtAdress.Text = My.Settings.adressCAOBS
        txtDisck.Text = My.Settings.disck
        txtIDstation.Text = My.Settings.IDst
        txtPlotArqus.Text = My.Settings.PlotArqus

        txtIndexVnt.Text = My.Settings.SmDD
        txtIndexT.Text = My.Settings.SmT
        txtIndexHR.Text = My.Settings.SmHR
        txtIndexPs.Text = My.Settings.SmPs
        txtIndexEw.Text = My.Settings.ShEw
        txtIndexRRhr.Text = My.Settings.ShRR

        txtindexMaxT.Text = My.Settings.KTx
        txtIndexMinT.Text = My.Settings.KTn
        txtIndexMaxU.Text = My.Settings.KUx
        txtIndexMinU.Text = My.Settings.KUn

        txtIndexIns.Text = My.Settings.KIs
        txtIndexRG.Text = My.Settings.KRg
        txtIndexRR.Text = My.Settings.KRR
        txtindexTsol.Text = My.Settings.KTnSol
        txtIndex_RR_sixMn.Text = My.Settings.MsRR
        txtIndex_Di_sixMn.Text = My.Settings.MsDI
        txtIndex_RG_sixMn.Text = My.Settings.MsRG

        txtIndex_WND_Hor.Text = My.Settings.HrWnd
        txtIndex_MinT_Hor.Text = My.Settings.HrMinT
        txtIndex_MaxT_HOR.Text = My.Settings.HrMxT
        txt_IndexMax_Hr_HOR.Text = My.Settings.HrMaxU
        txtIndex_MinHr_HOR.Text = My.Settings.HrMinU
        txtIndexRR_HOR.Text = My.Settings.HrRR
        txtIndexDi_HOR.Text = My.Settings.HrDi
        txtIndexRG_HOR.Text = My.Settings.HrRg

        txtPlotArqus.Text = My.Settings.PlotArqus
        txtRemotePath.Text = My.Settings.remotePath
        txtFtpUser.Text = My.Settings.ftpUser
        txtFtpPasswd.Text = My.Settings.ftpPass
    End Sub

    Private Sub btnquitconfig_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnquitconfig.Click
        Me.Hide()
        saisietdh.Show()

    End Sub
    Protected Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim param As CreateParams = MyBase.CreateParams
            param.ClassStyle = param.ClassStyle Or &H200
            Return param
        End Get
    End Property

    Private Sub btnFindT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindT.Click

        If txtDisck.Text = "" Or txtAdress.Text = "" Then
            MsgBox("Configuration manquante: l'adresse ou le nom du disck, n'est pas valide", MsgBoxStyle.Exclamation, vbOKOnly)
        ElseIf txtAdress.Text <> "" And txtDisck.Text <> "" Then
            '"\\172.17.56.3\aero_mes\PISTE0\"
            Dim adresseCAOBS As String = "\\" & txtAdress.Text & "\aero_mes\PISTE0\"
            Dim locaSPACE As String = txtDisck.Text
            Dim Tindx As Integer
            Dim fichierminute As String = "A_" & Date.Now.ToString("MMdd") & ".xls"
            Dim sourceminute As String = adresseCAOBS & fichierminute
            Dim localminute As String = locaSPACE & ":\data\bdcrq\" & fichierminute
            Try
                File.Copy(sourceminute, localminute, True)

            Catch ex As Exception
                MsgBox(ex.Message.ToString)
            End Try

            Dim app As Excel.Application
            Dim minutewbk As Excel.Workbook
            Dim minutwsh As Excel.Worksheet
            Dim lastColumn As Integer
            app = New Excel.Application
            minutewbk = app.Workbooks.Open(localminute)
            minutwsh = minutewbk.Sheets(1)
            lastColumn = minutwsh.UsedRange.Columns.Count
            For i As Integer = 1 To lastColumn
                If CType(minutwsh.Cells(1, i), Excel.Range).Value.ToString Like "TAir(Parc*)" Then
                    Tindx = i
                    txtIndexT.Text = Tindx
                    Exit For
                End If
            Next i

            minutewbk.Close(SaveChanges:=False)
            releaseObject(minutwsh)
            releaseObject(minutewbk)
            releaseObject(app)

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

    Private Sub btnFindHR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindHR.Click

        If txtDisck.Text = "" Or txtAdress.Text = "" Then
            MsgBox("Configuration manquante: l'adresse ou le nom du disck, n'est pas valide", MsgBoxStyle.Exclamation, vbOKOnly)
        ElseIf txtAdress.Text <> "" And txtDisck.Text <> "" Then
            '"\\172.17.56.3\aero_mes\PISTE0\"
            Dim adresseCAOBS As String = "\\" & txtAdress.Text & "\aero_mes\PISTE0\"
            Dim locaSPACE As String = txtDisck.Text
            Dim HRindx As Integer
            Dim fichierminute As String = "A_" & Date.Now.ToString("MMdd") & ".xls"
            Dim sourceminute As String = adresseCAOBS & fichierminute
            Dim localminute As String = locaSPACE & ":\data\bdcrq\" & fichierminute
            Try
                File.Copy(sourceminute, localminute, True)
            Catch ex As Exception
                MsgBox(ex.Message.ToString)
            End Try
            Dim app As Excel.Application
            Dim minutewbk As Excel.Workbook
            Dim minutwsh As Excel.Worksheet
            Dim lastColumn As Integer
            app = New Excel.Application
            minutewbk = app.Workbooks.Open(localminute) ' ne pas oublier de la fermer
            minutwsh = minutewbk.Sheets(1)
            lastColumn = minutwsh.UsedRange.Columns.Count
            For i As Integer = 1 To lastColumn
                If CType(minutwsh.Cells(1, i), Excel.Range).Value.ToString Like "Humidité(Parc*)" Then
                    HRindx = i
                    txtIndexHR.Text = HRindx
                    Exit For
                End If
            Next i

            minutewbk.Close(SaveChanges:=False)
            releaseObject(minutwsh)
            releaseObject(minutewbk)
            releaseObject(app)
        End If
    End Sub

    Private Sub btnFindPs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindPs.Click

        If txtDisck.Text = "" Or txtAdress.Text = "" Then
            MsgBox("Configuration manquante: l'adresse ou le nom du disck, n'est pas valide", MsgBoxStyle.Exclamation, vbOKOnly)
        ElseIf txtAdress.Text <> "" And txtDisck.Text <> "" Then
            '"\\172.17.56.3\aero_mes\PISTE0\"
            Dim adresseCAOBS As String = "\\" & txtAdress.Text & "\aero_mes\PISTE0\"
            Dim locaSPACE As String = txtDisck.Text
            Dim Psindx As Integer
            Dim fichierminute As String = "A_" & Date.Now.ToString("MMdd") & ".xls"
            Dim sourceminute As String = adresseCAOBS & fichierminute
            Dim localminute As String = locaSPACE & ":\data\bdcrq\" & fichierminute
            Try
                File.Copy(sourceminute, localminute, True)
            Catch ex As Exception
                MsgBox(ex.Message.ToString)
            End Try
            Dim app As Excel.Application
            Dim minutewbk As Excel.Workbook
            Dim minutwsh As Excel.Worksheet
            Dim lastColumn As Integer
            app = New Excel.Application
            minutewbk = app.Workbooks.Open(localminute) ' ne pas oublier de la fermer
            minutwsh = minutewbk.Sheets(1)
            lastColumn = minutwsh.UsedRange.Columns.Count
            For i As Integer = 1 To lastColumn
                If CType(minutwsh.Cells(1, i), Excel.Range).Value.ToString Like "Pression(Parc*)" Then
                    Psindx = i
                    txtIndexPs.Text = Psindx
                    Exit For
                End If
            Next i

            minutewbk.Close(SaveChanges:=False)
            releaseObject(minutwsh)
            releaseObject(minutewbk)
            releaseObject(app)
        End If
    End Sub

    Private Sub btnIndexVnt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnIndexVnt.Click

        If txtDisck.Text = "" Or txtAdress.Text = "" Then
            MsgBox("Configuration manquante: l'adresse ou le nom du disck, n'est pas valide", MsgBoxStyle.Exclamation, vbOKOnly)
        ElseIf txtAdress.Text <> "" And txtDisck.Text <> "" Then
            '"\\172.17.56.3\aero_mes\PISTE0\"
            Dim adresseCAOBS As String = "\\" & txtAdress.Text & "\aero_mes\PISTE0\"
            Dim locaSPACE As String = txtDisck.Text
            Dim Vntindx As Integer
            Dim fichierminute As String = "A_" & Date.Now.ToString("MMdd") & ".xls"
            Dim sourceminute As String = adresseCAOBS & fichierminute
            Dim localminute As String = locaSPACE & ":\data\bdcrq\" & fichierminute
            Try
                File.Copy(sourceminute, localminute, True)
            Catch ex As Exception
                MsgBox(ex.Message.ToString)
            End Try
            Dim app As Excel.Application
            Dim minutewbk As Excel.Workbook
            Dim minutwsh As Excel.Worksheet
            Dim lastColumn As Integer
            app = New Excel.Application
            minutewbk = app.Workbooks.Open(localminute) ' ne pas oublier de la fermer
            minutwsh = minutewbk.Sheets(1)
            lastColumn = minutwsh.UsedRange.Columns.Count
            For i As Integer = 1 To lastColumn
                If CType(minutwsh.Cells(1, i), Excel.Range).Value.ToString Like "DD 10'(Parc*)" Then
                    Vntindx = i
                    txtIndexVnt.Text = Vntindx
                    Exit For
                End If
            Next i

            minutewbk.Close(SaveChanges:=False)
            releaseObject(minutwsh)
            releaseObject(minutewbk)
            releaseObject(app)
        End If
    End Sub

    Private Sub btnIndexEw_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnIndexEw.Click

        If txtDisck.Text = "" Or txtAdress.Text = "" Then
            MsgBox("Configuration manquante: l'adresse ou le nom du disck, n'est pas valide", MsgBoxStyle.Exclamation, vbOKOnly)
        ElseIf txtAdress.Text <> "" And txtDisck.Text <> "" Then
            '"\\172.17.56.3\aero_mes\PISTE0\"
            Dim adresseCAOBS As String = "\\" & txtAdress.Text & "\aero_mes\PISTE0\"
            Dim locaSPACE As String = txtDisck.Text
            Dim Ewindx As Integer
            Dim fichierhoraire As String = "S_" & Date.Now.ToString("MMdd") & ".xls"
            Dim sourcehoraire As String = adresseCAOBS & fichierhoraire
            Dim localhoraire As String = locaSPACE & ":\data\bdcrq\" & fichierhoraire
            Try
                File.Copy(sourcehoraire, localhoraire, True)
            Catch ex As Exception
                MsgBox(ex.Message.ToString)
            End Try
            Dim app As Excel.Application
            Dim horairewbk As Excel.Workbook
            Dim horairewsh As Excel.Worksheet
            Dim lastColumn As Integer
            app = New Excel.Application
            horairewbk = app.Workbooks.Open(localhoraire) ' ne pas oublier de la fermer
            horairewsh = horairewbk.Sheets(1)
            lastColumn = horairewsh.UsedRange.Columns.Count
            For i As Integer = 1 To lastColumn
                If CType(horairewsh.Cells(1, i), Excel.Range).Value.ToString = "Tension vapeur d'eau" Then
                    Ewindx = i
                    txtIndexEw.Text = Ewindx
                    Exit For
                End If
            Next i

            horairewbk.Close(SaveChanges:=False)
            releaseObject(horairewsh)
            releaseObject(horairewbk)
            releaseObject(app)
        End If
    End Sub

    Private Sub btnIndexRRhr_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnIndexRRhr.Click

        If txtDisck.Text = "" Or txtAdress.Text = "" Then
            MsgBox("Configuration manquante: l'adresse ou le nom du disck, n'est pas valide", MsgBoxStyle.Exclamation, vbOKOnly)
        ElseIf txtAdress.Text <> "" And txtDisck.Text <> "" Then
            '"\\172.17.56.3\aero_mes\PISTE0\"
            Dim adresseCAOBS As String = "\\" & txtAdress.Text & "\aero_mes\PISTE0\"
            Dim locaSPACE As String = txtDisck.Text
            Dim RRindx As Integer
            Dim fichierhoraire As String = "S_" & Date.Now.ToString("MMdd") & ".xls"
            Dim sourcehoraire As String = adresseCAOBS & fichierhoraire
            Dim localhoraire As String = locaSPACE & ":\data\bdcrq\" & fichierhoraire
            Try
                File.Copy(sourcehoraire, localhoraire, True)
            Catch ex As Exception
                MsgBox(ex.Message.ToString)
            End Try
            Dim app As Excel.Application
            Dim horairewbk As Excel.Workbook
            Dim horairewsh As Excel.Worksheet
            Dim lastColumn As Integer
            app = New Excel.Application
            horairewbk = app.Workbooks.Open(localhoraire) ' ne pas oublier de la fermer
            horairewsh = horairewbk.Sheets(1)
            lastColumn = horairewsh.UsedRange.Columns.Count
            For i As Integer = 1 To lastColumn
                If CType(horairewsh.Cells(1, i), Excel.Range).Value.ToString = "Pluie" Then
                    RRindx = i
                    txtIndexRRhr.Text = RRindx
                    Exit For
                End If
            Next i

            horairewbk.Close(SaveChanges:=False)
            releaseObject(horairewsh)
            releaseObject(horairewbk)
            releaseObject(app)
        End If
    End Sub

    Private Sub btnFindMaxT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindMaxT.Click

        If txtDisck.Text = "" Or txtAdress.Text = "" Then
            MsgBox("Configuration manquante: l'adresse ou le nom du disck, n'est pas valide", MsgBoxStyle.Exclamation, vbOKOnly)
        ElseIf txtAdress.Text <> "" And txtDisck.Text <> "" Then
            '"\\172.17.56.3\aero_mes\PISTE0\"
            Dim adresseCAOBS As String = "\\" & txtAdress.Text & "\aero_mes\PISTE0\"
            Dim locaSPACE As String = txtDisck.Text
            Dim TMaxindx As Integer
            Dim fichierextreme As String = "K_" & Date.Now.AddDays(-1).ToString("MMdd") & ".xls"
            Dim sourcextreme As String = adresseCAOBS & fichierextreme
            Dim localextreme As String = locaSPACE & ":\data\bdcrq\" & fichierextreme
            Try
                File.Copy(sourcextreme, localextreme, True)
            Catch ex As Exception
                MsgBox(ex.Message.ToString)
            End Try
            Dim app As Excel.Application
            Dim extremwbk As Excel.Workbook
            Dim extremwsh As Excel.Worksheet
            Dim lastColumn As Integer
            app = New Excel.Application
            extremwbk = app.Workbooks.Open(localextreme)
            extremwsh = extremwbk.Sheets(1)
            lastColumn = extremwsh.UsedRange.Columns.Count
            For i As Integer = 1 To lastColumn
                If CType(extremwsh.Cells(1, i), Excel.Range).Value.ToString = "Max Tair" Then
                    TMaxindx = i
                    txtindexMaxT.Text = TMaxindx
                    Exit For
                End If
            Next i
            extremwbk.Close(SaveChanges:=False)
            releaseObject(extremwsh)
            releaseObject(extremwbk)
            releaseObject(app)
        End If

    End Sub

    Private Sub btnFindMinT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindMinT.Click

        If txtDisck.Text = "" Or txtAdress.Text = "" Then
            MsgBox("Configuration manquante: l'adresse ou le nom du disck, n'est pas valide", MsgBoxStyle.Exclamation, vbOKOnly)
        ElseIf txtAdress.Text <> "" And txtDisck.Text <> "" Then

            Dim adresseCAOBS As String = "\\" & txtAdress.Text & "\aero_mes\PISTE0\"
            Dim locaSPACE As String = txtDisck.Text
            Dim TMinindx As Integer
            Dim fichierextreme As String = "K_" & Date.Now.AddDays(-1).ToString("MMdd") & ".xls"
            Dim sourcextreme As String = adresseCAOBS & fichierextreme
            Dim localextreme As String = locaSPACE & ":\data\bdcrq\" & fichierextreme
            Try
                File.Copy(sourcextreme, localextreme, True)
            Catch ex As Exception
                MsgBox(ex.Message.ToString)
            End Try
            Dim app As Excel.Application
            Dim extremwbk As Excel.Workbook
            Dim extremwsh As Excel.Worksheet
            Dim lastColumn As Integer
            app = New Excel.Application
            extremwbk = app.Workbooks.Open(localextreme)
            extremwsh = extremwbk.Sheets(1)
            lastColumn = extremwsh.UsedRange.Columns.Count
            For i As Integer = 1 To lastColumn
                If CType(extremwsh.Cells(1, i), Excel.Range).Value.ToString = "Min Tair" Then
                    TMinindx = i
                    txtIndexMinT.Text = TMinindx
                    Exit For
                End If
            Next i

            extremwbk.Close(SaveChanges:=False)
            releaseObject(extremwsh)
            releaseObject(extremwbk)
            releaseObject(app)
        End If
    End Sub

    Private Sub btnFindMaxU_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindMaxU.Click

        If txtDisck.Text = "" Or txtAdress.Text = "" Then
            MsgBox("Configuration manquante: l'adresse ou le nom du disck, n'est pas valide", MsgBoxStyle.Exclamation, vbOKOnly)
        ElseIf txtAdress.Text <> "" And txtDisck.Text <> "" Then

            Dim adresseCAOBS As String = "\\" & txtAdress.Text & "\aero_mes\PISTE0\"
            Dim locaSPACE As String = txtDisck.Text
            Dim UMaxindx As Integer
            Dim fichierextreme As String = "K_" & Date.Now.AddDays(-1).ToString("MMdd") & ".xls"
            Dim sourcextreme As String = adresseCAOBS & fichierextreme
            Dim localextreme As String = locaSPACE & ":\data\bdcrq\" & fichierextreme
            Try
                File.Copy(sourcextreme, localextreme, True)
            Catch ex As Exception
                MsgBox(ex.Message.ToString)
            End Try
            Dim app As Excel.Application
            Dim extremwbk As Excel.Workbook
            Dim extremwsh As Excel.Worksheet
            Dim lastColumn As Integer
            app = New Excel.Application
            extremwbk = app.Workbooks.Open(localextreme)
            extremwsh = extremwbk.Sheets(1)
            lastColumn = extremwsh.UsedRange.Columns.Count
            For i As Integer = 1 To lastColumn
                If CType(extremwsh.Cells(1, i), Excel.Range).Value.ToString = "Max HR" Then
                    UMaxindx = i
                    txtIndexMaxU.Text = UMaxindx
                    Exit For
                End If
            Next i

            extremwbk.Close(SaveChanges:=False)
            releaseObject(extremwsh)
            releaseObject(extremwbk)
            releaseObject(app)
        End If
    End Sub

    Private Sub btnFindMinU_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindMinU.Click

        If txtDisck.Text = "" Or txtAdress.Text = "" Then
            MsgBox("Configuration manquante: l'adresse ou le nom du disck, n'est pas valide", MsgBoxStyle.Exclamation, vbOKOnly)
        ElseIf txtAdress.Text <> "" And txtDisck.Text <> "" Then

            Dim adresseCAOBS As String = "\\" & txtAdress.Text & "\aero_mes\PISTE0\"
            Dim locaSPACE As String = txtDisck.Text
            Dim UMinindx As Integer
            Dim fichierextreme As String = "K_" & Date.Now.AddDays(-1).ToString("MMdd") & ".xls"
            Dim sourcextreme As String = adresseCAOBS & fichierextreme
            Dim localextreme As String = locaSPACE & ":\data\bdcrq\" & fichierextreme
            Try
                File.Copy(sourcextreme, localextreme, True)
            Catch ex As Exception
                MsgBox(ex.Message.ToString)
            End Try
            Dim app As Excel.Application
            Dim extremwbk As Excel.Workbook
            Dim extremwsh As Excel.Worksheet
            Dim lastColumn As Integer
            app = New Excel.Application
            extremwbk = app.Workbooks.Open(localextreme)
            extremwsh = extremwbk.Sheets(1)
            lastColumn = extremwsh.UsedRange.Columns.Count
            For i As Integer = 1 To lastColumn
                If CType(extremwsh.Cells(1, i), Excel.Range).Value.ToString = "Min HR" Then
                    UMinindx = i
                    txtIndexMinU.Text = UMinindx
                    Exit For
                End If
            Next i

            extremwbk.Close(SaveChanges:=False)
            releaseObject(extremwsh)
            releaseObject(extremwbk)
            releaseObject(app)
        End If
    End Sub

    Private Sub btnFindCumulRR_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindCumulRR.Click

        If txtDisck.Text = "" Or txtAdress.Text = "" Then
            MsgBox("Configuration manquante: l'adresse ou le nom du disck, n'est pas valide", MsgBoxStyle.Exclamation, vbOKOnly)
        ElseIf txtAdress.Text <> "" And txtDisck.Text <> "" Then

            Dim adresseCAOBS As String = "\\" & txtAdress.Text & "\aero_mes\PISTE0\"
            Dim locaSPACE As String = txtDisck.Text
            Dim RRCumindx As Integer
            Dim fichierextreme As String = "K_" & Date.Now.AddDays(-1).ToString("MMdd") & ".xls"
            Dim sourcextreme As String = adresseCAOBS & fichierextreme
            Dim localextreme As String = locaSPACE & ":\data\bdcrq\" & fichierextreme
            Try
                File.Copy(sourcextreme, localextreme, True)
            Catch ex As Exception
                MsgBox(ex.Message.ToString)
            End Try
            Dim app As Excel.Application
            Dim extremwbk As Excel.Workbook
            Dim extremwsh As Excel.Worksheet
            Dim lastColumn As Integer
            app = New Excel.Application
            extremwbk = app.Workbooks.Open(localextreme)
            extremwsh = extremwbk.Sheets(1)
            lastColumn = extremwsh.UsedRange.Columns.Count
            For i As Integer = 1 To lastColumn
                If CType(extremwsh.Cells(1, i), Excel.Range).Value.ToString = "Pluie" Then
                    RRCumindx = i
                    txtIndexRR.Text = RRCumindx
                    Exit For
                End If
            Next i

            extremwbk.Close(SaveChanges:=False)
            releaseObject(extremwsh)
            releaseObject(extremwbk)
            releaseObject(app)
        End If

    End Sub

    Private Sub btnFindIns_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindIns.Click

        If txtDisck.Text = "" Or txtAdress.Text = "" Then
            MsgBox("Configuration manquante: l'adresse ou le nom du disck, n'est pas valide", MsgBoxStyle.Exclamation, vbOKOnly)
        ElseIf txtAdress.Text <> "" And txtDisck.Text <> "" Then

            Dim adresseCAOBS As String = "\\" & txtAdress.Text & "\aero_mes\PISTE0\"
            Dim locaSPACE As String = txtDisck.Text
            Dim InsKindx As Integer
            Dim fichierextreme As String = "K_" & Date.Now.AddDays(-1).ToString("MMdd") & ".xls"
            Dim sourcextreme As String = adresseCAOBS & fichierextreme
            Dim localextreme As String = locaSPACE & ":\data\bdcrq\" & fichierextreme
            Try
                File.Copy(sourcextreme, localextreme, True)
            Catch ex As Exception
                MsgBox(ex.Message.ToString)
            End Try
            Dim app As Excel.Application
            Dim extremwbk As Excel.Workbook
            Dim extremwsh As Excel.Worksheet
            Dim lastColumn As Integer
            app = New Excel.Application
            extremwbk = app.Workbooks.Open(localextreme)
            extremwsh = extremwbk.Sheets(1)
            lastColumn = extremwsh.UsedRange.Columns.Count
            For i As Integer = 1 To lastColumn
                If CType(extremwsh.Cells(1, i), Excel.Range).Value.ToString = "Durée Insol." Then
                    InsKindx = i
                    txtIndexIns.Text = InsKindx
                    Exit For
                End If
            Next i

            extremwbk.Close(SaveChanges:=False)
            releaseObject(extremwsh)
            releaseObject(extremwbk)
            releaseObject(app)
        End If
    End Sub

    Private Sub btnFindRGlo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFindRGlo.Click

        If txtDisck.Text = "" Or txtAdress.Text = "" Then
            MsgBox("Configuration manquante: l'adresse ou le nom du disck, n'est pas valide", MsgBoxStyle.Exclamation, vbOKOnly)
        ElseIf txtAdress.Text <> "" And txtDisck.Text <> "" Then

            Dim adresseCAOBS As String = "\\" & txtAdress.Text & "\aero_mes\PISTE0\"
            Dim locaSPACE As String = txtDisck.Text
            Dim RGindx As Integer
            Dim fichierextreme As String = "K_" & Date.Now.AddDays(-1).ToString("MMdd") & ".xls"
            Dim sourcextreme As String = adresseCAOBS & fichierextreme
            Dim localextreme As String = locaSPACE & ":\data\bdcrq\" & fichierextreme
            Try
                File.Copy(sourcextreme, localextreme, True)
            Catch ex As Exception
                MsgBox(ex.Message.ToString)
            End Try
            Dim app As Excel.Application
            Dim extremwbk As Excel.Workbook
            Dim extremwsh As Excel.Worksheet
            Dim lastColumn As Integer
            app = New Excel.Application
            extremwbk = app.Workbooks.Open(localextreme)
            extremwsh = extremwbk.Sheets(1)
            lastColumn = extremwsh.UsedRange.Columns.Count
            For i As Integer = 1 To lastColumn
                If CType(extremwsh.Cells(1, i), Excel.Range).Value.ToString = "Ray. global" Then
                    RGindx = i
                    txtIndexRG.Text = RGindx
                    Exit For
                End If
            Next i

            extremwbk.Close(SaveChanges:=False)
            releaseObject(extremwsh)
            releaseObject(extremwbk)
            releaseObject(app)
        End If

    End Sub

    Private Sub btnMakeFolders_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMakeFolders.Click

        Dim disck As String = txtDisck.Text

        If disck = "" Then
            MsgBox("Indiquez le nom du disk à utiliser pour ARQUS", MsgBoxStyle.Critical, vbOKOnly)
        ElseIf disck <> "" Then

            Try

                If IO.Directory.Exists(disck & ":\") Then

                    MsgBox("Le disque " & disck & " existe", MsgBoxStyle.Information, vbOKOnly)
                    If IO.Directory.Exists(disck & ":\" & "Data\") Then
                        If Not IO.Directory.Exists(disck & ":\" & "Data\données station\") Then
                            IO.Directory.CreateDirectory(disck & ":\" & "Data\données station\")
                            MsgBox("L'espace de stokage" & disck & ":\" & "Data\données station\ est créé", MsgBoxStyle.Information, vbOKOnly)
                        End If
                        If Not IO.Directory.Exists(disck & ":\" & "Data\bdcrq\") Then
                            IO.Directory.CreateDirectory(disck & ":\" & "Data\bdcrq\")
                            MsgBox("L'espace de stokage" & disck & ":\" & "Data\bdcrq\ est créé", MsgBoxStyle.Information, vbOKOnly)
                        End If

                    ElseIf Not IO.Directory.Exists(disck & ":\" & "Data\") Then
                        IO.Directory.CreateDirectory(disck & ":\" & "Data\données station\")
                        MsgBox("Le répertoire d'archivage des CRQs: 'Data\données station\' est crée", MsgBoxStyle.Information, vbOKOnly)

                        IO.Directory.CreateDirectory(disck & ":\" & "Data\bdcrq\")
                        MsgBox("Le répertoire d'archivage des données Brutes: 'Data\bdcrq\' est crée", MsgBoxStyle.Information, vbOKOnly)
                    End If

                Else

                    MsgBox("Le nom du disque n'est pas correcte ou n'existe pas sur votre ordinateur!", MsgBoxStyle.Critical, vbOKOnly)

                End If

            Catch ex As Exception

                MessageBox.Show(ex.Message)

            End Try
        End If

        If Not IO.Directory.Exists(disck & ":\" & "Data\bdcrq\TempoCopy\") Then
            IO.Directory.CreateDirectory(disck & ":\" & "Data\bdcrq\TempoCopy\")
        End If
        If Not IO.Directory.Exists(disck & ":\" & "Data\bdcrq\SmnCsv\") Then
            IO.Directory.CreateDirectory(disck & ":\" & "Data\bdcrq\SmnCsv\")
        End If
        If Not IO.Directory.Exists(disck & ":\" & "Data\bdcrq\HorCsv\") Then
            IO.Directory.CreateDirectory(disck & ":\" & "Data\bdcrq\HorCsv\")
        End If

    End Sub

    Private Sub btnFind_Tsol_Click(sender As Object, e As EventArgs) Handles btnFind_Tsol.Click

        If txtDisck.Text = "" Or txtAdress.Text = "" Then
            MsgBox("Configuration manquante: l'adresse ou le nom du disck, n'est pas valide", MsgBoxStyle.Exclamation, vbOKOnly)
        ElseIf txtAdress.Text <> "" And txtDisck.Text <> "" Then

            Dim adresseCAOBS As String = "\\" & txtAdress.Text & "\aero_mes\PISTE0\"
            Dim locaSPACE As String = txtDisck.Text
            Dim UMinindx As Integer
            Dim fichierextreme As String = "K_" & Date.Now.AddDays(-1).ToString("MMdd") & ".xls"
            Dim sourcextreme As String = adresseCAOBS & fichierextreme
            Dim localextreme As String = locaSPACE & ":\data\bdcrq\" & fichierextreme
            Try
                File.Copy(sourcextreme, localextreme, True)
            Catch ex As Exception
                MsgBox(ex.Message.ToString)
            End Try
            Dim XlApp As Excel.Application
            Dim extremwbk As Excel.Workbook
            Dim extremwsh As Excel.Worksheet
            XlApp = New Excel.Application
            Dim lastColumn As Integer
            extremwbk = XlApp.Workbooks.Open(localextreme)
            extremwsh = extremwbk.Sheets(1)
            lastColumn = extremwsh.UsedRange.Columns.Count
            For i As Integer = 1 To lastColumn
                If CType(extremwsh.Cells(1, i), Excel.Range).Value.ToString = "Min T+10" OrElse CType(extremwsh.Cells(1, i), Excel.Range).Value.ToString = "Min Tsol" Then
                    UMinindx = i
                    txtindexTsol.Text = UMinindx
                    Exit For
                End If
            Next i

            extremwbk.Close(SaveChanges:=False)
            releaseObject(extremwsh)
            releaseObject(extremwbk)
            releaseObject(XlApp)

        End If

    End Sub

    Private Sub btnFindRR_SixMn_Click(sender As Object, e As EventArgs) Handles btnFindRR_SixMn.Click
        If txtDisck.Text = "" Or txtAdress.Text = "" Then
            MsgBox("Configuration manquante: l'adresse ou le nom du disck, n'est pas valide", MsgBoxStyle.Exclamation, vbOKOnly)
        ElseIf txtAdress.Text <> "" And txtDisck.Text <> "" Then
            '"\\172.17.56.3\aero_mes\PISTE0\"
            Dim adresseCAOBS As String = "\\" & txtAdress.Text & "\aero_mes\PISTE0\"
            Dim locaSPACE As String = txtDisck.Text
            Dim RRindex As Integer
            Dim fichierSixMn As String = "M_" & Date.Now.ToString("MMdd") & ".xls"
            Dim sourceSixMn As String = adresseCAOBS & fichierSixMn
            Dim localSixMn As String = locaSPACE & ":\data\bdcrq\" & fichierSixMn
            Try
                File.Copy(sourceSixMn, localSixMn, True)

            Catch ex As Exception
                MsgBox(ex.Message.ToString)
            End Try

            Dim app As Excel.Application
            Dim SixMnwbk As Excel.Workbook
            Dim SixMnwsh As Excel.Worksheet
            Dim lastColumn As Integer
            app = New Excel.Application
            SixMnwbk = app.Workbooks.Open(localSixMn)
            SixMnwsh = SixMnwbk.Sheets(1)
            lastColumn = SixMnwsh.UsedRange.Columns.Count
            For i As Integer = 1 To lastColumn
                If CType(SixMnwsh.Cells(1, i), Excel.Range).Value.ToString = "Pluie" Then
                    RRindex = i
                    txtIndex_RR_sixMn.Text = RRindex
                    Exit For
                End If
            Next i

            SixMnwbk.Close(SaveChanges:=False)
            releaseObject(SixMnwsh)
            releaseObject(SixMnwbk)
            releaseObject(app)

        End If
    End Sub

    Private Sub btnFindDI_sixMn_Click(sender As Object, e As EventArgs) Handles btnFindDI_sixMn.Click
        If txtDisck.Text = "" Or txtAdress.Text = "" Then
            MsgBox("Configuration manquante: l'adresse ou le nom du disck, n'est pas valide", MsgBoxStyle.Exclamation, vbOKOnly)
        ElseIf txtAdress.Text <> "" And txtDisck.Text <> "" Then
            '"\\172.17.56.3\aero_mes\PISTE0\"
            Dim adresseCAOBS As String = "\\" & txtAdress.Text & "\aero_mes\PISTE0\"
            Dim locaSPACE As String = txtDisck.Text
            Dim Diindex_Smn As Integer
            Dim fichierSixMn As String = "M_" & Date.Now.ToString("MMdd") & ".xls"
            Dim sourceSixMn As String = adresseCAOBS & fichierSixMn
            Dim localSixMn As String = locaSPACE & ":\data\bdcrq\" & fichierSixMn
            Try
                File.Copy(sourceSixMn, localSixMn, True)

            Catch ex As Exception
                MsgBox(ex.Message.ToString)
            End Try

            Dim app As Excel.Application
            Dim SixMnwbk As Excel.Workbook
            Dim SixMnwsh As Excel.Worksheet
            Dim lastColumn As Integer
            app = New Excel.Application
            SixMnwbk = app.Workbooks.Open(localSixMn)
            SixMnwsh = SixMnwbk.Sheets(1)
            lastColumn = SixMnwsh.UsedRange.Columns.Count
            For i As Integer = 1 To lastColumn
                If CType(SixMnwsh.Cells(1, i), Excel.Range).Value.ToString = "Durée Insol." Then
                    Diindex_Smn = i
                    txtIndex_Di_sixMn.Text = Diindex_Smn
                    Exit For
                End If
            Next i

            SixMnwbk.Close(SaveChanges:=False)
            releaseObject(SixMnwsh)
            releaseObject(SixMnwbk)
            releaseObject(app)

        End If
    End Sub

    Private Sub btnFindRG_sixMn_Click(sender As Object, e As EventArgs) Handles btnFindRG_sixMn.Click
        If txtDisck.Text = "" Or txtAdress.Text = "" Then
            MsgBox("Configuration manquante: l'adresse ou le nom du disck, n'est pas valide", MsgBoxStyle.Exclamation, vbOKOnly)
        ElseIf txtAdress.Text <> "" And txtDisck.Text <> "" Then
            '"\\172.17.56.3\aero_mes\PISTE0\"
            Dim adresseCAOBS As String = "\\" & txtAdress.Text & "\aero_mes\PISTE0\"
            Dim locaSPACE As String = txtDisck.Text
            Dim RGindex_Smn As Integer
            Dim fichierSixMn As String = "M_" & Date.Now.ToString("MMdd") & ".xls"
            Dim sourceSixMn As String = adresseCAOBS & fichierSixMn
            Dim localSixMn As String = locaSPACE & ":\data\bdcrq\" & fichierSixMn
            Try
                File.Copy(sourceSixMn, localSixMn, True)

            Catch ex As Exception
                MsgBox(ex.Message.ToString)
            End Try

            Dim app As Excel.Application
            Dim SixMnwbk As Excel.Workbook
            Dim SixMnwsh As Excel.Worksheet
            Dim lastColumn As Integer
            app = New Excel.Application
            SixMnwbk = app.Workbooks.Open(localSixMn)
            SixMnwsh = SixMnwbk.Sheets(1)
            lastColumn = SixMnwsh.UsedRange.Columns.Count
            For i As Integer = 1 To lastColumn
                If CType(SixMnwsh.Cells(1, i), Excel.Range).Value.ToString = "Ray. global" Then
                    RGindex_Smn = i
                    txtIndex_RG_sixMn.Text = RGindex_Smn
                    Exit For
                End If
            Next i

            SixMnwbk.Close(SaveChanges:=False)
            releaseObject(SixMnwsh)
            releaseObject(SixMnwbk)
            releaseObject(app)

        End If
    End Sub

    Private Sub btnIndex_Vent_HR_Click(sender As Object, e As EventArgs) Handles btnIndex_Vent_HR.Click
        If txtDisck.Text = "" Or txtAdress.Text = "" Then
            MsgBox("Configuration manquante: l'adresse ou le nom du disck, n'est pas valide", MsgBoxStyle.Exclamation, vbOKOnly)
        ElseIf txtAdress.Text <> "" And txtDisck.Text <> "" Then
            '"\\172.17.56.3\aero_mes\PISTE0\"
            Dim adresseCAOBS As String = "\\" & txtAdress.Text & "\aero_mes\PISTE0\"
            Dim locaSPACE As String = txtDisck.Text
            Dim WNDindex_Hor As Integer
            Dim fichierHOR As String = "S_" & Date.Now.ToString("MMdd") & ".xls"
            Dim sourceHOR As String = adresseCAOBS & fichierHOR
            Dim localHOR As String = locaSPACE & ":\data\bdcrq\" & fichierHOR
            Try
                File.Copy(sourceHOR, localHOR, True)

            Catch ex As Exception
                MsgBox(ex.Message.ToString)
            End Try

            Dim app As Excel.Application
            Dim HORwbk As Excel.Workbook
            Dim HORwsh As Excel.Worksheet
            Dim lastColumn As Integer
            app = New Excel.Application
            HORwbk = app.Workbooks.Open(localHOR)
            HORwsh = HORwbk.Sheets(1)
            lastColumn = HORwsh.UsedRange.Columns.Count
            For i As Integer = 1 To lastColumn
                If CType(HORwsh.Cells(1, i), Excel.Range).Value.ToString = "Max DD inst" Then
                    WNDindex_Hor = i
                    txtIndex_WND_Hor.Text = WNDindex_Hor
                    Exit For
                End If
            Next i

            HORwbk.Close(SaveChanges:=False)
            releaseObject(HORwsh)
            releaseObject(HORwbk)
            releaseObject(app)

        End If
    End Sub

    Private Sub btnIndex_MinT_HOR_Click(sender As Object, e As EventArgs) Handles btnIndex_MinT_HOR.Click
        If txtDisck.Text = "" Or txtAdress.Text = "" Then
            MsgBox("Configuration manquante: l'adresse ou le nom du disck, n'est pas valide", MsgBoxStyle.Exclamation, vbOKOnly)
        ElseIf txtAdress.Text <> "" And txtDisck.Text <> "" Then
            '"\\172.17.56.3\aero_mes\PISTE0\"
            Dim adresseCAOBS As String = "\\" & txtAdress.Text & "\aero_mes\PISTE0\"
            Dim locaSPACE As String = txtDisck.Text
            Dim TMinindex_Hor As Integer
            Dim fichierHOR As String = "S_" & Date.Now.ToString("MMdd") & ".xls"
            Dim sourceHOR As String = adresseCAOBS & fichierHOR
            Dim localHOR As String = locaSPACE & ":\data\bdcrq\" & fichierHOR
            Try
                File.Copy(sourceHOR, localHOR, True)

            Catch ex As Exception
                MsgBox(ex.Message.ToString)
            End Try

            Dim app As Excel.Application
            Dim HORwbk As Excel.Workbook
            Dim HORwsh As Excel.Worksheet
            Dim lastColumn As Integer
            app = New Excel.Application
            HORwbk = app.Workbooks.Open(localHOR)
            HORwsh = HORwbk.Sheets(1)
            lastColumn = HORwsh.UsedRange.Columns.Count
            For i As Integer = 1 To lastColumn
                If CType(HORwsh.Cells(1, i), Excel.Range).Value.ToString = "Min TAir" Then
                    TMinindex_Hor = i
                    txtIndex_MinT_Hor.Text = TMinindex_Hor
                    Exit For
                End If
            Next i

            HORwbk.Close(SaveChanges:=False)
            releaseObject(HORwsh)
            releaseObject(HORwbk)
            releaseObject(app)

        End If
    End Sub

    Private Sub btnIndex_MaxT_HOR_Click(sender As Object, e As EventArgs) Handles btnIndex_MaxT_HOR.Click
        If txtDisck.Text = "" Or txtAdress.Text = "" Then
            MsgBox("Configuration manquante: l'adresse ou le nom du disck, n'est pas valide", MsgBoxStyle.Exclamation, vbOKOnly)
        ElseIf txtAdress.Text <> "" And txtDisck.Text <> "" Then
            '"\\172.17.56.3\aero_mes\PISTE0\"
            Dim adresseCAOBS As String = "\\" & txtAdress.Text & "\aero_mes\PISTE0\"
            Dim locaSPACE As String = txtDisck.Text
            Dim TMaxindexHOR As Integer
            Dim fichierHOR As String = "S_" & Date.Now.ToString("MMdd") & ".xls"
            Dim sourceHOR As String = adresseCAOBS & fichierHOR
            Dim localHOR As String = locaSPACE & ":\data\bdcrq\" & fichierHOR
            Try
                File.Copy(sourceHOR, localHOR, True)

            Catch ex As Exception
                MsgBox(ex.Message.ToString)
            End Try

            Dim app As Excel.Application
            Dim HORwbk As Excel.Workbook
            Dim HORwsh As Excel.Worksheet
            Dim lastColumn As Integer
            app = New Excel.Application
            HORwbk = app.Workbooks.Open(localHOR)
            HORwsh = HORwbk.Sheets(1)
            lastColumn = HORwsh.UsedRange.Columns.Count
            For i As Integer = 1 To lastColumn
                If CType(HORwsh.Cells(1, i), Excel.Range).Value.ToString = "Max TAir" Then
                    TMaxindexHOR = i
                    txtIndex_MaxT_HOR.Text = TMaxindexHOR
                    Exit For
                End If
            Next i

            HORwbk.Close(SaveChanges:=False)
            releaseObject(HORwsh)
            releaseObject(HORwbk)
            releaseObject(app)

        End If
    End Sub

    Private Sub btnIndex_MinHR_HOR_Click(sender As Object, e As EventArgs) Handles btnIndex_MinHR_HOR.Click
        If txtDisck.Text = "" Or txtAdress.Text = "" Then
            MsgBox("Configuration manquante: l'adresse ou le nom du disck, n'est pas valide", MsgBoxStyle.Exclamation, vbOKOnly)
        ElseIf txtAdress.Text <> "" And txtDisck.Text <> "" Then
            '"\\172.17.56.3\aero_mes\PISTE0\"
            Dim adresseCAOBS As String = "\\" & txtAdress.Text & "\aero_mes\PISTE0\"
            Dim locaSPACE As String = txtDisck.Text
            Dim UMinHOR As Integer
            Dim fichierHOR As String = "S_" & Date.Now.ToString("MMdd") & ".xls"
            Dim sourceHOR As String = adresseCAOBS & fichierHOR
            Dim localHOR As String = locaSPACE & ":\data\bdcrq\" & fichierHOR
            Try
                File.Copy(sourceHOR, localHOR, True)

            Catch ex As Exception
                MsgBox(ex.Message.ToString)
            End Try

            Dim app As Excel.Application
            Dim HORwbk As Excel.Workbook
            Dim HORwsh As Excel.Worksheet
            Dim lastColumn As Integer
            app = New Excel.Application
            HORwbk = app.Workbooks.Open(localHOR)
            HORwsh = HORwbk.Sheets(1)
            lastColumn = HORwsh.UsedRange.Columns.Count
            For i As Integer = 1 To lastColumn
                If CType(HORwsh.Cells(1, i), Excel.Range).Value.ToString = "Min HR" Then
                    UMinHOR = i
                    txtIndex_MinHr_HOR.Text = UMinHOR
                    Exit For
                End If
            Next i

            HORwbk.Close(SaveChanges:=False)
            releaseObject(HORwsh)
            releaseObject(HORwbk)
            releaseObject(app)

        End If
    End Sub

    Private Sub btnIndex_MaxHR_HOR_Click(sender As Object, e As EventArgs) Handles btnIndex_MaxHR_HOR.Click
        If txtDisck.Text = "" Or txtAdress.Text = "" Then
            MsgBox("Configuration manquante: l'adresse ou le nom du disck, n'est pas valide", MsgBoxStyle.Exclamation, vbOKOnly)
        ElseIf txtAdress.Text <> "" And txtDisck.Text <> "" Then
            '"\\172.17.56.3\aero_mes\PISTE0\"
            Dim adresseCAOBS As String = "\\" & txtAdress.Text & "\aero_mes\PISTE0\"
            Dim locaSPACE As String = txtDisck.Text
            Dim UMaxHOR As Integer
            Dim fichierHOR As String = "S_" & Date.Now.ToString("MMdd") & ".xls"
            Dim sourceHOR As String = adresseCAOBS & fichierHOR
            Dim localHOR As String = locaSPACE & ":\data\bdcrq\" & fichierHOR
            Try
                File.Copy(sourceHOR, localHOR, True)

            Catch ex As Exception
                MsgBox(ex.Message.ToString)
            End Try

            Dim app As Excel.Application
            Dim HORwbk As Excel.Workbook
            Dim HORwsh As Excel.Worksheet
            Dim lastColumn As Integer
            app = New Excel.Application
            HORwbk = app.Workbooks.Open(localHOR)
            HORwsh = HORwbk.Sheets(1)
            lastColumn = HORwsh.UsedRange.Columns.Count
            For i As Integer = 1 To lastColumn
                If CType(HORwsh.Cells(1, i), Excel.Range).Value.ToString = "Max HR" Then
                    UMaxHOR = i
                    txt_IndexMax_Hr_HOR.Text = UMaxHOR
                    Exit For
                End If
            Next i

            HORwbk.Close(SaveChanges:=False)
            releaseObject(HORwsh)
            releaseObject(HORwbk)
            releaseObject(app)

        End If
    End Sub

    Private Sub btnIndex_RR_HOR_Click(sender As Object, e As EventArgs) Handles btnIndex_RR_HOR.Click
        If txtDisck.Text = "" Or txtAdress.Text = "" Then
            MsgBox("Configuration manquante: l'adresse ou le nom du disck, n'est pas valide", MsgBoxStyle.Exclamation, vbOKOnly)
        ElseIf txtAdress.Text <> "" And txtDisck.Text <> "" Then
            '"\\172.17.56.3\aero_mes\PISTE0\"
            Dim adresseCAOBS As String = "\\" & txtAdress.Text & "\aero_mes\PISTE0\"
            Dim locaSPACE As String = txtDisck.Text
            Dim RRHor As Integer
            Dim fichierHOR As String = "S_" & Date.Now.ToString("MMdd") & ".xls"
            Dim sourceHOR As String = adresseCAOBS & fichierHOR
            Dim localHOR As String = locaSPACE & ":\data\bdcrq\" & fichierHOR
            Try
                File.Copy(sourceHOR, localHOR, True)

            Catch ex As Exception
                MsgBox(ex.Message.ToString)
            End Try

            Dim app As Excel.Application
            Dim HORwbk As Excel.Workbook
            Dim HORwsh As Excel.Worksheet
            Dim lastColumn As Integer
            app = New Excel.Application
            HORwbk = app.Workbooks.Open(localHOR)
            HORwsh = HORwbk.Sheets(1)
            lastColumn = HORwsh.UsedRange.Columns.Count
            For i As Integer = 1 To lastColumn
                If CType(HORwsh.Cells(1, i), Excel.Range).Value.ToString = "Pluie" Then
                    RRHor = i
                    txtIndexRR_HOR.Text = RRHor
                    Exit For
                End If
            Next i

            HORwbk.Close(SaveChanges:=False)
            releaseObject(HORwsh)
            releaseObject(HORwbk)
            releaseObject(app)

        End If
    End Sub

    Private Sub btnIndexDi_HOR_Click(sender As Object, e As EventArgs) Handles btnIndexDi_HOR.Click
        If txtDisck.Text = "" Or txtAdress.Text = "" Then
            MsgBox("Configuration manquante: l'adresse ou le nom du disck, n'est pas valide", MsgBoxStyle.Exclamation, vbOKOnly)
        ElseIf txtAdress.Text <> "" And txtDisck.Text <> "" Then
            '"\\172.17.56.3\aero_mes\PISTE0\"
            Dim adresseCAOBS As String = "\\" & txtAdress.Text & "\aero_mes\PISTE0\"
            Dim locaSPACE As String = txtDisck.Text
            Dim DiHOR As Integer
            Dim fichierHOR As String = "S_" & Date.Now.ToString("MMdd") & ".xls"
            Dim sourceHOR As String = adresseCAOBS & fichierHOR
            Dim localHOR As String = locaSPACE & ":\data\bdcrq\" & fichierHOR
            Try
                File.Copy(sourceHOR, localHOR, True)

            Catch ex As Exception
                MsgBox(ex.Message.ToString)
            End Try

            Dim app As Excel.Application
            Dim HORwbk As Excel.Workbook
            Dim HORwsh As Excel.Worksheet
            Dim lastColumn As Integer
            app = New Excel.Application
            HORwbk = app.Workbooks.Open(localHOR)
            HORwsh = HORwbk.Sheets(1)
            lastColumn = HORwsh.UsedRange.Columns.Count
            For i As Integer = 1 To lastColumn
                If CType(HORwsh.Cells(1, i), Excel.Range).Value.ToString = "Durée Insol." Then
                    DiHOR = i
                    txtIndexDi_HOR.Text = DiHOR
                    Exit For
                End If
            Next i

            HORwbk.Close(SaveChanges:=False)
            releaseObject(HORwsh)
            releaseObject(HORwbk)
            releaseObject(app)

        End If
    End Sub

    Private Sub btnIndexRG_HOR_Click(sender As Object, e As EventArgs) Handles btnIndexRG_HOR.Click
        If txtDisck.Text = "" Or txtAdress.Text = "" Then
            MsgBox("Configuration manquante: l'adresse ou le nom du disck, n'est pas valide", MsgBoxStyle.Exclamation, vbOKOnly)
        ElseIf txtAdress.Text <> "" And txtDisck.Text <> "" Then
            '"\\172.17.56.3\aero_mes\PISTE0\"
            Dim adresseCAOBS As String = "\\" & txtAdress.Text & "\aero_mes\PISTE0\"
            Dim locaSPACE As String = txtDisck.Text
            Dim RgHOR As Integer
            Dim fichierHOR As String = "S_" & Date.Now.ToString("MMdd") & ".xls"
            Dim sourceHOR As String = adresseCAOBS & fichierHOR
            Dim localHOR As String = locaSPACE & ":\data\bdcrq\" & fichierHOR
            Try
                File.Copy(sourceHOR, localHOR, True)

            Catch ex As Exception
                MsgBox(ex.Message.ToString)
            End Try

            Dim app As Excel.Application
            Dim HORwbk As Excel.Workbook
            Dim HORwsh As Excel.Worksheet
            Dim lastColumn As Integer
            app = New Excel.Application
            HORwbk = app.Workbooks.Open(localHOR)
            HORwsh = HORwbk.Sheets(1)
            lastColumn = HORwsh.UsedRange.Columns.Count
            For i As Integer = 1 To lastColumn
                If CType(HORwsh.Cells(1, i), Excel.Range).Value.ToString = "Pluie" Then
                    RgHOR = i
                    txtIndexRG_HOR.Text = RgHOR
                    Exit For
                End If
            Next i

            HORwbk.Close(SaveChanges:=False)
            releaseObject(HORwsh)
            releaseObject(HORwbk)
            releaseObject(app)

        End If
    End Sub

    Private Sub btnPlotArqusActivation_Click(sender As Object, e As EventArgs) Handles btnPlotArqusActivation.Click

        If rdStartSend.Checked Then
            My.Settings.autoSend = True
            lblPlotArqusActivation.Show()

        ElseIf rdStopSend.Checked Then
            My.Settings.autoSend = False
            lblPlotArqusActivation.Hide()

        End If
    End Sub
End Class