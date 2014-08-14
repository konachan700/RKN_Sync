Imports System.IO

Public Class Form1

    Dim RKN As ru.gov.rkn.vigruzki.OperatorRequestService
    Dim OkFlag As Boolean = False, ErrFlag As Boolean = False, StopFlag As Boolean = True, ExecFlag As Boolean = True
    Dim OkFlag2 As Boolean = False, ErrFlag2 As Boolean = False
    Dim strDocumentVersion As String = "", strWebServiceVersion As String = "", strFormatVersion As String = ""
    Dim lDDU As Long, lDDU_last As Long, lLastUpdate As Long, lLastUpdate_last As Long


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Hide()
    End Sub

    Private Sub Form1_Disposed(sender As Object, e As EventArgs) Handles Me.Disposed
        Thread1.CancelAsync()
    End Sub


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        RKN = New ru.gov.rkn.vigruzki.OperatorRequestService()
    End Sub

    Private Sub Thread1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles Thread1.DoWork
        On Error GoTo ERRLABEL
        ExecFlag = False

        'Dim retval As Long
        lLastUpdate = RKN.getLastDumpDateEx(lDDU, strWebServiceVersion, strFormatVersion, strDocumentVersion)

        OkFlag = True
        ExecFlag = True

        Threading.Thread.Sleep(1000 * 60 * NumericUpDown1.Value)

        Exit Sub

ERRLABEL:
        ErrFlag = True
        ExecFlag = True

        Threading.Thread.Sleep(1000 * 60)

        Exit Sub
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim oDateTime As System.DateTime, nDateTime As System.DateTime
        Button3.Enabled = False
        Button4.Enabled = True
        StopFlag = True
        Timer1.Enabled = True
        GenerateTextsToolStripMenuItem.Enabled = False

        Do While (StopFlag)
            System.Windows.Forms.Application.DoEvents()

            If ExecFlag = True And Thread1.IsBusy = False And (OkFlag = False And ErrFlag = False) Then
                Thread1.WorkerSupportsCancellation = True
                Thread1.RunWorkerAsync()
            End If

            If OkFlag = True Or ErrFlag = True Then
                ' Магические числа такие магические =^_^= ибо utc+4 так сделать проще всего. Ни на что в логике программы не влияет, нужно только для ведения логов.
                oDateTime = New System.DateTime((lLastUpdate * 1000L * 10L) + 621355968000000000L + (1000L * 1000L * 10L * 60 * 60 * 4), DateTimeKind.Utc)
                nDateTime = New System.DateTime((lDDU * 1000L * 10L) + 621355968000000000L + (1000L * 1000L * 10L * 60 * 60 * 4), DateTimeKind.Utc)

                ListView1.Items.Add(oDateTime.ToShortDateString + " " + oDateTime.ToLongTimeString)

                If OkFlag = True Then
                    ListView1.Items(ListView1.Items.Count - 1).SubItems.Add(nDateTime.ToShortDateString + " " + nDateTime.ToLongTimeString)
                    ListView1.Items(ListView1.Items.Count - 1).SubItems.Add("WEB:" + strWebServiceVersion + " Format:" + strFormatVersion + " Doc:" + strDocumentVersion)

                    If (lDDU_last <> lDDU) Or (lLastUpdate <> lLastUpdate_last) Then
                        Thread2.RunWorkerAsync()
                        Do While (Thread2.IsBusy)
                            System.Windows.Forms.Application.DoEvents()
                        Loop

                        If ErrFlag2 = True Then
                            ListView1.Items.Add(oDateTime.ToShortDateString + " " + oDateTime.ToLongTimeString)
                            ListView1.Items(ListView1.Items.Count - 1).SubItems.Add(nDateTime.ToShortDateString + " " + nDateTime.ToLongTimeString)
                            ListView1.Items(ListView1.Items.Count - 1).SubItems.Add("Error on SOAP interface, see log file")
                            ListView1.SelectedIndices.Clear()
                            ListView1.SelectedIndices.Add(ListView1.Items.Count - 1)
                            ErrFlag2 = False
                        End If

                        If OkFlag2 = True Then
                            ListView1.Items.Add(oDateTime.ToShortDateString + " " + oDateTime.ToLongTimeString)
                            ListView1.Items(ListView1.Items.Count - 1).SubItems.Add(nDateTime.ToShortDateString + " " + nDateTime.ToLongTimeString)
                            ListView1.Items(ListView1.Items.Count - 1).SubItems.Add("ZIP-file uploaded")
                            ListView1.Items(ListView1.Items.Count - 1).ForeColor = Color.DarkGreen
                            ListView1.Items(ListView1.Items.Count - 1).BackColor = Color.LightGreen
                            ListView1.SelectedIndices.Clear()
                            ListView1.SelectedIndices.Add(ListView1.Items.Count - 1)
                        End If

                        lDDU_last = lDDU
                        lLastUpdate_last = lLastUpdate
                    End If

                    OkFlag = False
                End If

                If ErrFlag = True Then
                    ListView1.Items(ListView1.Items.Count - 1).SubItems.Add("Connection fail")
                    ErrFlag = False
                End If
            End If

            If ListView1.Items.Count > 250 Then
                Dim strTemp1 As String = ""
                For i = 0 To ListView1.Items.Count
                    strTemp1 = strTemp1 + vbNewLine + vbTab + vbTab + vbTab + ListView1.Items(i).Text + vbTab + ListView1.Items(i).SubItems(1).Text
                Next

                ToLog(TextBox2.Text + "ok.log", "List cleared." + strTemp1)

                ListView1.Items.Clear()
            End If

            System.Windows.Forms.Application.DoEvents()
        Loop

        Button3.Enabled = True
        Timer1.Enabled = False
        GenerateTextsToolStripMenuItem.Enabled = True
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        StopFlag = False
        Button4.Enabled = False
        Thread1.CancelAsync()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        ProgressBar1.Value = ProgressBar1.Value + 1
        If ProgressBar1.Value >= 11 Then ProgressBar1.Value = 0
    End Sub

    Private Sub ToLog(file As String, s As String)
        Dim oDateTime As System.DateTime = System.DateTime.Now()
        My.Computer.FileSystem.WriteAllText(file, vbNewLine + "[" + oDateTime.ToString("yyyy-MM-dd hh:mm:ss") + "] " + vbTab + ">> " + s, True)
    End Sub

    Private Sub Thread2_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles Thread2.DoWork
        Dim oDateTime As System.DateTime = System.DateTime.Now()

        Dim txtPath As String = TextBox2.Text
        Dim dirInfo As IO.DirectoryInfo = New IO.DirectoryInfo(txtPath)
        If dirInfo.Exists = False Then
            dirInfo.Create()
        End If

        Dim dirInfoIN As IO.DirectoryInfo = New IO.DirectoryInfo(txtPath + "IN\")
        If dirInfoIN.Exists = False Then
            dirInfoIN.Create()
        End If

        Dim dirInfoOUT As IO.DirectoryInfo = New IO.DirectoryInfo(txtPath + "OUT\")
        If dirInfoOUT.Exists = False Then
            dirInfoOUT.Create()
        End If

        ToLog(txtPath + "ok.log", "Started")

        Try
            Dim txtSignProfile As String = TextBox1.Text

            Dim req As String = "<?xml version=""1.0"" encoding=""windows-1251""?>" + vbNewLine + _
                                "<request>" + vbNewLine + _
                                "<requestTime>" + oDateTime.ToString("yyyy-MM-dd") + "T" + oDateTime.ToString("hh:mm:ss") + ".000+04:00</requestTime>" + vbNewLine + _
                                "<operatorName>Регион Связь</operatorName>" + vbNewLine + _
                                "<inn>7733820832</inn>" + vbNewLine + _
                                "<ogrn>1127747104638</ogrn>" + vbNewLine + _
                                "<email>bazil.87@mail.ru</email>" + vbNewLine + _
                                "</request>" + vbNewLine

            My.Computer.FileSystem.WriteAllBytes(txtPath + "IN/rkn_text.txt", System.Text.Encoding.Default.GetBytes(req), False)
            My.Computer.FileSystem.WriteAllBytes(txtPath + "tdhelper.vbs", System.Text.Encoding.Default.GetBytes(My.Resources.tdhelper), False)

            ToLog(txtPath + "ok.log", "XML Document created")

            Dim strBat As String = "chcp 1251" + vbNewLine + "c:\Windows\System32\cscript.exe " + txtPath + "tdhelper.vbs """ + txtPath + "IN\"" """ + txtPath + "OUT\"" """ + txtSignProfile + """ ""S"" >> " + txtPath + "out.log"
            My.Computer.FileSystem.WriteAllBytes(txtPath + "run.bat", System.Text.Encoding.Default.GetBytes(strBat), False)

            'Dim pSign As Process = Process.Start(txtPath + "run.bat")

            Dim pPSI As New ProcessStartInfo()
            pPSI.WindowStyle = ProcessWindowStyle.Hidden
            pPSI.WorkingDirectory = txtPath
            'pPSI.StandardOutputEncoding = System.Text.Encoding.Default
            'pPSI.StandardErrorEncoding = System.Text.Encoding.Default
            pPSI.FileName = txtPath + "run.bat"

            Dim pSign As Process = Process.Start(pPSI)
            pSign.WaitForExit(1000 * 60)

            ToLog(txtPath + "ok.log", "XML Document signed; sign profile: " + txtSignProfile)

            Dim bytesOfText As Byte() = My.Computer.FileSystem.ReadAllBytes(txtPath + "IN/rkn_text.txt")
            Dim bytesOfSign As Byte() = My.Computer.FileSystem.ReadAllBytes(txtPath + "OUT/rkn_text.txt.sig")

            Dim strResult1 As String = "", strResult2 As String = ""

            Dim bRetval As Boolean = RKN.sendRequest(bytesOfText, bytesOfSign, "2.0", strResult1, strResult2)
            If bRetval = True Then
                ToLog(txtPath + "ok.log", "OK RKN.sendRequest >> Comment:[" + strResult1 + "] Code:[" + strResult2 + "]")

                Dim bRetval2 As Boolean = False
                Dim strResultComment As String = "", strFormatVersion As String = ""
                Dim baZIP As Byte() = New Byte() {}
                Dim iResultCode As Integer = 0

                Do While (1)
                    Threading.Thread.Sleep(1000 * 7)
                    bRetval2 = RKN.getResult(strResult2, strResultComment, baZIP, iResultCode, strFormatVersion)
                    If (bRetval2 = True) Or (iResultCode = 1) Then
                        ToLog(txtPath + "ok.log", "OK RKN.getResult >> ResultCode:[" + iResultCode.ToString + "] COMMENT:[" + strResultComment + "] File size:[" + baZIP.Count.ToString + "]")

                        My.Computer.FileSystem.WriteAllBytes(txtPath + "register.zip", baZIP, False)
                        OkFlag2 = True

                        Exit Do
                    Else
                        If iResultCode < 0 Then
                            ToLog(txtPath + "ok.log", "ERROR RKN.getResult >> ResultCode:[" + iResultCode.ToString + "] COMMENT:[" + strResultComment + "]")
                            ErrFlag2 = True
                            Exit Do
                        End If
                    End If
                Loop
            Else
                ToLog(txtPath + "ok.log", "ERROR RKN.sendRequest >> R1:[" + strResult1 + "] R2:[" + strResult2 + "]")
                ErrFlag2 = True
            End If

            Exit Sub

        Catch ER As Exception
            ToLog(txtPath + "ok.log", "ERROR MESSAGE >>" + ER.Message)
            ToLog(txtPath + "ok.log", "STACK TRACE >>" + ER.StackTrace)
            ErrFlag2 = True
            Exit Sub
        End Try
    End Sub

    Private Sub GenerateTextsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GenerateTextsToolStripMenuItem.Click
        Timer1.Enabled = True
        GenerateTextsToolStripMenuItem.Enabled = False
        Thread2.RunWorkerAsync()
        Do While (Thread2.IsBusy)
            System.Windows.Forms.Application.DoEvents()
        Loop
        GenerateTextsToolStripMenuItem.Enabled = True
        Timer1.Enabled = False
    End Sub

    Private Sub NotifyIcon1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
        Me.Show()
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        End
    End Sub
End Class
