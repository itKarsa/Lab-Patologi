Imports MySql.Data.MySqlClient
Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Imports Word = Microsoft.Office.Interop.Word
Imports System.Text

Public Class Form3

    Public Ambil_Data As String
    Public Form_Ambil_Data As String
    Private path As String = ""
    Dim pathImage As String = ""

    Private wdApp As Word.Application
    Private wdDocs As Word.Documents
    Private sPath As String = System.Windows.Forms.Application.StartupPath & "\"
    Private sFileName As String

    Public htmlMakro, htmlMikro, htmlConclu As String
    Dim kdHasil As String = ""

    Sub setColor(button As Button)
        btnDash.BackColor = SystemColors.HotTrack
        btnHasil.BackColor = SystemColors.HotTrack
        button.BackColor = Color.DodgerBlue
    End Sub

    Sub autoNoHasil()
        Dim noHasilPa As String
        Try
            Call koneksiServer()
            Dim query As String
            query = "SELECT SUBSTR(kdHasilPA,17,3) FROM t_hasilpemeriksaanpatologi ORDER BY CAST(SUBSTR(kdHasilPA,17,3) AS UNSIGNED) DESC LIMIT 1"
            cmd = New MySqlCommand(query, conn)
            dr = cmd.ExecuteReader
            dr.Read()
            If dr.HasRows Then
                dr.Read()
                noHasilPa = "HPA" + Format(Now, "ddMMyyHHmmss") + "-" + (Val(Trim(dr.Item(0).ToString)) + 1).ToString
                txtKdHasilPa.Text = noHasilPa
            Else
                noHasilPa = "HPA" + Format(Now, "ddMMyyHHmmss") + "-1"
                txtKdHasilPa.Text = noHasilPa
            End If
            conn.Close()
        Catch ex As Exception

        End Try
    End Sub

    Sub addHasil()
        Call koneksiServer()
        Try
            Dim str As String
            str = "INSERT INTO t_hasilpemeriksaanpatologi(kdHasilPA,noRM,noDaftar,
                                                         noRegistrasiPA,noPA,asalRs,
                                                         asalPengirim,lokalisasi,diagnosa,
                                                         bahan,tglTerimaSediaan,tglHasil,
                                                         userModify,dateModify) 
                   VALUES ('" & txtKdHasilPa.Text & "','" & txtNoRM.Text & "','" & Form1.txtNoReg.Text & "',
                           '" & txtNoPermintaan.Text & "','" & txtNoPaBaru.Text & "','" & txtRs.Text & "',
                           '" & txtUnitAsal.Text & "','" & txtLokalisasi.Text & "','" & txtKetKlinis.Text & "',
                           '" & txtBahan.Text & "','" & Format(dateTerimaSediaan.Value, "dd/MM/yyyy") & "','" & Format(dateTerimaSediaan.Value, "dd/MM/yyyy") & "',
                           ';" & LoginForm.txtUsername.Text & "',';" & Format(DateTime.Now, "dd/MM/yyyy HH:mm:ss") & "')"
            cmd = New MySqlCommand(str, conn)
            cmd.ExecuteNonQuery()
            'MessageBox.Show("Insert nomor hasil PA berhasil dilakukan.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Insert nomor hasil PA gagal dilakukan." & vbCrLf & ex.ToString, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

        conn.Close()
    End Sub

    Sub tampilKdHasil()
        Try
            Call koneksiServer()
            Dim str As String
            Dim cmd As MySqlCommand
            Dim dr As MySqlDataReader
            str = "SELECT kdHasilPA FROM t_hasilpemeriksaanpatologi WHERE noRegistrasiPA = '" & txtNoPermintaan.Text & "'"
            cmd = New MySqlCommand(str, conn)
            dr = cmd.ExecuteReader
            dr.Read()
            If dr.HasRows Then
                kdHasil = dr.Item("kdHasilPA").ToString
            End If
            dr.Close()
        Catch ex As Exception
        End Try
        conn.Close()
    End Sub

    Sub tampilHasil()
        Try
            Call koneksiServer()
            Dim str As String
            Dim cmd As MySqlCommand
            Dim dr As MySqlDataReader
            str = "SELECT makroskopik, mikroskopik, kesimpulan, noPA, tglTerimaSediaan, tglHasil FROM t_hasilpemeriksaanpatologi WHERE noRegistrasiPA = '" & txtNoPermintaan.Text & "'"
            cmd = New MySqlCommand(str, conn)
            dr = cmd.ExecuteReader
            dr.Read()
            If dr.HasRows Then

                txtNoPaBaru.Text = dr.Item("noPA").ToString
                dateTerimaSediaan.Value = dr.Item("tglTerimaSediaan")
                dateHasil.Value = dr.Item("tglHasil")

                If dr.IsDBNull(0) Then
                    Return
                Else
                    RichTextBox1.Rtf = System.Text.Encoding.Unicode.GetChars(dr.Item("makroskopik"))
                End If

                If dr.IsDBNull(1) Then
                    Return
                Else
                    RichTextBox2.Rtf = System.Text.Encoding.Unicode.GetChars(dr.Item("mikroskopik"))
                End If

                If dr.IsDBNull(2) Then
                    Return
                Else
                    RichTextBox3.Rtf = System.Text.Encoding.Unicode.GetChars(dr.Item("kesimpulan"))
                End If

            End If
            dr.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        conn.Close()
    End Sub

    Sub updateData()
        Call koneksiServer()
        Try
            Dim str As String = ""
            Dim cmd As MySqlCommand

            Select Case Form1.txtKdInstalasi.Text
                Case "RI"
                    str = "UPDATE t_registrasipatologiranap 
                      SET noPA = '" & txtNoPaBaru.Text & "',
                          noPALama = '" & txtNoPA.Text & "',
                          lokalisasi = '" & txtLokalisasi.Text & "',
                          diagnosaklinis = '" & txtKetKlinis.Text & "',
                          bahan = '" & txtBahan.Text & "',
                          userModify = CONCAT(userModify,';" & LoginForm.txtUsername.Text & "'),  
                          dateModify = CONCAT(dateModify,';" & Format(DateTime.Now, "dd/MM/yyyy HH:mm:ss") & "')  
                    WHERE noRegistrasiPARanap = '" & txtNoPermintaan.Text & "'"
                Case "RJ"
                    str = "UPDATE t_registrasipatologirajal
                      SET noPA = '" & txtNoPaBaru.Text & "',
                          noPALama = '" & txtNoPA.Text & "',
                          lokalisasi = '" & txtLokalisasi.Text & "',
                          diagnosaklinis = '" & txtKetKlinis.Text & "',
                          bahan = '" & txtBahan.Text & "',
                          userModify = CONCAT(userModify,';" & LoginForm.txtUsername.Text & "'),  
                          dateModify = CONCAT(dateModify,';" & Format(DateTime.Now, "dd/MM/yyyy HH:mm:ss") & "')  
                    WHERE noRegistrasiPARajal = '" & txtNoPermintaan.Text & "'"
            End Select

            'MsgBox(str)
            cmd = New MySqlCommand(str, conn)
            cmd.ExecuteNonQuery()
            MsgBox("Data berhasil diupdate.", MsgBoxStyle.Information)
        Catch ex As Exception
            MessageBox.Show(ex.ToString & " -update data registrasi-")
        End Try
        conn.Close()
    End Sub

    Sub updateDataHasil()
        Call koneksiServer()
        Try
            Dim str As String = ""
            Dim cmd As MySqlCommand

            str = "UPDATE t_hasilpemeriksaanpatologi
                      SET noPA = '" & txtNoPaBaru.Text & "',
                          tglTerimaSediaan = '" & Format(dateTerimaSediaan.Value, "dd/MM/yyyy") & "',  
                          tglHasil = '" & Format(dateHasil.Value, "dd/MM/yyyy") & "',
                          userModify = CONCAT(userModify,';" & LoginForm.txtUsername.Text & "'),  
                          dateModify = CONCAT(dateModify,';" & Format(DateTime.Now, "dd/MM/yyyy HH:mm:ss") & "')  
                    WHERE kdHasilPA = '" & kdHasil & "'"

            'MsgBox(str)
            cmd = New MySqlCommand(str, conn)
            cmd.ExecuteNonQuery()
            'MsgBox("data berhasil diupdate.", MsgBoxStyle.Information)
        Catch ex As Exception
            MessageBox.Show(ex.ToString & " -update data hasil-")
        End Try
        conn.Close()
    End Sub

    Private Sub Form3_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.WindowState = FormWindowState.Normal
        Me.StartPosition = FormStartPosition.Manual
        With Screen.PrimaryScreen.WorkingArea
            Me.SetBounds(.Left, .Top, .Width, .Height)
        End With

        'If wdDocs IsNot Nothing Then
        '    ReleaseObject(wdDocs)
        'End If
        'If wdApp IsNot Nothing Then
        '    ReleaseObject(wdApp)
        'End If

        dateTerimaSediaan.Format = DateTimePickerFormat.Custom
        dateHasil.Format = DateTimePickerFormat.Custom
        dateTerimaSediaan.CustomFormat = "dd/MM/yyyy"
        dateHasil.CustomFormat = "dd/MM/yyyy"

        pnlStats.Height = btnHasil.Height
        pnlStats.Top = btnHasil.Top
        btnHasil.BackColor = Color.DodgerBlue

        If Ambil_Data = True Then
            Select Case Form_Ambil_Data
                Case "Hasil"
                    txtInstalasi.Text = Form1.txtInstalasi.Text
                    txtNoRM.Text = Form1.txtNoRM.Text
                    txtNoReg.Text = Form1.txtNoReg.Text
                    txtNoPA.Text = Form1.txtNoPALama.Text
                    txtNoPaBaru.Text = Form1.txtNoPABaru.Text
                    txtNamaPasien.Text = Form1.txtNamaPasien.Text
                    txtAlamat.Text = Form1.txtAlamat.Text
                    txtJk.Text = Form1.txtJK.Text
                    txtTglLahir.Text = Form1.txtTglLahir.Text
                    txtUsia.Text = Form1.txtUsia.Text
                    txtRs.Text = "RSU Karsa Husada Batu"
                    txtDokter.Text = Form1.txtDokter.Text
                    txtKdDokter.Text = Form1.txtKdDokter.Text
                    txtDokterPA.Text = Form1.txtDokPA.Text
                    txtKdDokterPA.Text = Form1.txtKdDokterPA.Text
                    txtLokalisasi.Text = Form1.txtLokalisasi.Text
                    txtNoPermintaan.Text = Form1.txtNoPermintaan.Text
                    txtUnitAsal.Text = Form1.txtUnitAsal.Text
                    txtDateReg.Text = Form1.txtTglReg.Text
                    txtKetKlinis.Text = Form1.txtKlinis.Text
                    txtBahan.Text = Form1.txtBahan.Text
            End Select
        End If

        Call autoNoHasil()
        Call tampilHasil()
        Call tampilKdHasil()
        'Call tampilHasilPemeriksaanRanap()
        txtNamaPasien.Text = System.Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(txtNamaPasien.Text.ToLower)
        txtJk.Text = System.Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(txtJk.Text.ToLower)
        txtAlamat.Text = System.Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(txtAlamat.Text.ToLower)
        txtDokter.Text = "dr. " & System.Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(txtDokter.Text.Substring(4).ToLower)
        txtDokterPA.Text = "dr. " & System.Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(txtDokterPA.Text.Substring(4).ToLower)
        txtLokalisasi.Text = System.Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(txtLokalisasi.Text.ToLower)
        txtKetKlinis.Text = System.Threading.Thread.CurrentThread.CurrentCulture.TextInfo.ToTitleCase(txtKetKlinis.Text.ToLower)
    End Sub

    Private Sub btnDash_Click(sender As Object, e As EventArgs) Handles btnDash.Click
        'Form1.pnlStats.Height = Form1.btnDash.Height
        'Form1.pnlStats.Top = Form1.btnDash.Top
        'Dim btn As Button = CType(sender, Button)
        'setColor(btn)
        Form1.Show()
        Me.Close()
    End Sub

    Private Sub btnKeluar_Click(sender As Object, e As EventArgs) Handles btnKeluar.Click
        'Form1.pnlStats.Height = Form1.btnDash.Height
        'Form1.pnlStats.Top = Form1.btnDash.Top
        'Dim btn As Button = CType(sender, Button)
        'setColor(btn)
        Form1.Show()
        Me.Close()
    End Sub

    Private Sub btnCetak_Click(sender As Object, e As EventArgs) Handles btnCetak.Click
        Dim r As New SautinSoft.RtfToHtml
        r.OutputFormat = SautinSoft.RtfToHtml.eOutputFormat.HTML_5
        r.Encoding = SautinSoft.RtfToHtml.eEncoding.UTF_8
        r.TextStyle.InlineCSS = True
        r.TextStyle.Font = True

        htmlMakro = r.ConvertString(RichTextBox1.Rtf)
        htmlMikro = r.ConvertString(RichTextBox2.Rtf)
        htmlConclu = r.ConvertString(RichTextBox3.Rtf)
        'MsgBox(htmlMakro)
        viewCetakHasil.Show()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs)
        'Dim noRM, noPermintaan As String

        'If e.RowIndex = -1 Then
        '    Return
        'End If


        'noRM = DataGridView1.Rows(e.RowIndex).Cells(0).Value
        'noPermintaan = DataGridView1.Rows(e.RowIndex).Cells(0).Value

        'txtNoPermintaan.Text = noPermintaan
    End Sub

    Private Sub btnCari_Click(sender As Object, e As EventArgs) Handles btnCari.Click
        Try
            With OpenFileDialog1
                .InitialDirectory = sPath & "Template"
                .Title = "Browse Template Docx"
                .FileName = ""
                .Filter = "Word Template | *.dotx"
                .ShowDialog()
                path = .FileName
                txtTemplate.Text = path.Substring(path.LastIndexOf("\") + 1)
            End With
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
    End Sub

    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Sub createTemplate(filename As String)
        wdApp = New Word.Application
        wdDocs = wdApp.Documents

        Dim wdDoc As Word.Document = wdDocs.Add(sPath & "Template\" & txtTemplate.Text)
        Dim wdBooks As Word.Bookmarks = wdDoc.Bookmarks

        wdBooks("txtNoRM").Range.Text = txtNoRM.Text.ToString()
        wdBooks("txtNama").Range.Text = txtNamaPasien.Text.ToString()
        wdBooks("txtJK").Range.Text = txtJk.Text.ToString()
        wdBooks("txtTglLahir").Range.Text = txtTglLahir.Text.ToString()
        wdBooks("txtAlamat").Range.Text = txtAlamat.Text.ToString()
        wdBooks("txtDataPengirim").Range.Text = txtUnitAsal.Text.ToString()
        wdBooks("txtDokterPengirim").Range.Text = txtDokter.Text.ToString()
        wdBooks("txtTglTerimSediaan").Range.Text = dateTerimaSediaan.Text.ToString()
        wdBooks("txtTglHasil").Range.Text = dateHasil.Text.ToString()
        wdBooks("txtNoPa").Range.Text = txtNoPaBaru.Text.ToString()
        wdBooks("txtLokalisasi").Range.Text = txtLokalisasi.Text.ToString()
        wdBooks("txtDiagnos").Range.Text = txtKetKlinis.Text.ToString()
        wdBooks("txtBahan").Range.Text = txtBahan.Text.ToString()

        wdDoc.SaveAs2(sPath & "Hasil Bacaan\" & filename)

        ReleaseObject(wdBooks)
        wdDoc.Close(False)
        ReleaseObject(wdDoc)
        ReleaseObject(wdDocs)
        wdApp.Quit()

        txtFileName.Enabled = False

        MsgBox("File hasil pemeriksaan berhasil dibuat", MsgBoxStyle.Information, "Information")
    End Sub

    Sub IsiHasil(filename As String)
        wdApp = New Word.Application
        wdDocs = wdApp.Documents

        Dim qword As Word.Application
        Dim qdoc As Word.Document

        Dim wdDoc As Word.Document = wdDocs.Open(sPath & "Hasil Bacaan\" & filename & ".docx")
        wdDoc.SaveAs2(sPath & filename, Word.WdSaveFormat.wdFormatHTML)

        wdDoc.Close(False)
        ReleaseObject(wdDoc)
        ReleaseObject(wdDocs)
        wdApp.Quit()

        qword = CType(CreateObject("word.Application"), Application)
        qdoc = qword.Documents.Open(sPath & "Hasil Bacaan\" & filename & ".docx")
        qword.Visible = True
    End Sub

    Private Sub btnIsi_Click(sender As Object, e As EventArgs) Handles btnIsi.Click
        'If String.IsNullOrWhiteSpace(txtNoPaBaru.Text) Then
        '    MessageBox.Show("isi nama file terlebih dahulu", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        '    Exit Sub
        'End If

        'Dim konfirmasi As MsgBoxResult
        'Dim tempPath As String = ""

        'konfirmasi = MsgBox("Apakah anda yakin ingin menyimpannya ?", CType(vbQuestion + vbYesNo, MsgBoxStyle), "Save as")
        'If konfirmasi = vbYes Then
        '    sFileName = txtFileName.Text
        '    Call createTemplate(sFileName)
        '    Call IsiHasil(sFileName)
        'End If

        'Call addHasil()
        Dim isi As IsiHasil = New IsiHasil
        isi.ShowDialog()

    End Sub

    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        Call updateData()
        Call updateDataHasil()

        If kdHasil = "" Then
            Call addHasil()
        End If
    End Sub

    Private Sub ConvertWordToPDF(filename As String)
        Dim wordApplication As New Microsoft.Office.Interop.Word.Application
        Dim wordDocument As Microsoft.Office.Interop.Word.Document = Nothing
        Dim outputFilename As String

        Try
            wordDocument = wordApplication.Documents.Open(filename)
            outputFilename = System.IO.Path.ChangeExtension(filename, "pdf")

            If Not wordDocument Is Nothing Then
                wordDocument.ExportAsFixedFormat(outputFilename, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF, False, Microsoft.Office.Interop.Word.WdExportOptimizeFor.wdExportOptimizeForOnScreen, Microsoft.Office.Interop.Word.WdExportRange.wdExportAllDocument, 0, 0, Microsoft.Office.Interop.Word.WdExportItem.wdExportDocumentContent, True, True, Microsoft.Office.Interop.Word.WdExportCreateBookmarks.wdExportCreateNoBookmarks, True, True, False)
            End If
        Catch ex As Exception
            'TODO: handle exception
        Finally
            If Not wordDocument Is Nothing Then
                wordDocument.Close(False)
                wordDocument = Nothing
            End If

            If Not wordApplication Is Nothing Then
                wordApplication.Quit()
                wordApplication = Nothing
            End If
        End Try
    End Sub

    Private Sub btnSimpan_Click(sender As Object, e As EventArgs) Handles btnSimpan.Click
        'Try
        '    ConvertWordToPDF(sPath & "Hasil Bacaan\" & sFileName & ".docx")
        '    txtFileName.Text = sFileName & ".pdf"
        '    Me.Close()
        'Catch ex As Exception
        '    MsgBox(ex.Message)
        'End Try
    End Sub
End Class