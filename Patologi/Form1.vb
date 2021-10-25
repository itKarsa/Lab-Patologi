Imports MySql.Data.MySqlClient
Imports System.Deployment.Application
Public Class Form1

    Public Ambil_Data As String
    Public Form_Ambil_Data As String

    Public noTindakanPA As String
    Public caraBayar As String
    Dim idDetail As String
    Dim noTindakanPatologi As String
    Public jk As String

    Dim ind As String
    Dim status As String

    Dim separator As String = ","

    Sub setColor(button As Button)
        btnDash.BackColor = SystemColors.HotTrack
        btnTindakan.BackColor = SystemColors.HotTrack
        btnHasil.BackColor = SystemColors.HotTrack
        button.BackColor = Color.DodgerBlue
    End Sub

    Sub cariDataPasien()

        Dim query As String
        query = "SELECT * FROM vw_pasienpatologi
                  WHERE noRekamedis Like '%" & txtNoRM.Text & "%'
                  ORDER BY tglMasukPARajal DESC"

        txtNoReg.Text = ""
        txtNamaPasien.Text = ""
        txtUsia.Text = ""
        txtAlamat.Text = ""
        txtDokter.Text = ""
        txtKlinis.Text = ""
        txtTglReg.Text = ""
        txtNoPermintaan.Text = ""
        txtTglLahir.Text = ""
        noTindakanPatologi = ""

        Try
            Call koneksiServer()
            cmd = New MySqlCommand(query, conn)
            dr = cmd.ExecuteReader
            dgv1.Rows.Clear()
            Do While dr.Read
                dgv1.Rows.Add(dr.Item("noRekamedis"), dr.Item("noDaftar"), dr.Item("nmPasien"), dr.Item("kdUnitAsal"), dr.Item("unitAsal"),
                              dr.Item("tglLahir"), dr.Item("alamat"), dr.Item("noRegistrasiPARajal"), dr.Item("tglMasukPARajal"), dr.Item("statusPA"),
                              dr.Item("kdDokterPengirim"), dr.Item("namapetugasMedis"), dr.Item("kdDokterPemeriksaan"), dr.Item("dokterPemeriksa"), dr.Item("diagnosaKlinis"),
                              dr.Item("noTindakanPARajal"), dr.Item("totalTindakanPA"), dr.Item("statusPembayaran"), dr.Item("carabayar"), dr.Item("jenisKelamin"))
            Loop

            dr.Close()
            conn.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Timer5.Start()

    End Sub

    Sub tampilDataAll()
        Call koneksiServer()
        Dim query As String
        query = "SELECT * FROM vw_pasienpatologi
                  WHERE tglMasukPARajal BETWEEN '" & Format(DateTime.Now, "yyyy-MM-dd") & "' 
                    AND '" & Format(DateAdd(DateInterval.Day, 1, DateTime.Now), "yyyy-MM-dd") & "'
                  ORDER BY tglMasukPARajal DESC"
        'da = New MySqlDataAdapter("SELECT * FROM vw_pasienpatologi
        '                            WHERE tglMasukPARajal BETWEEN '2020-04-01' AND '2020-04-06'
        '                            ORDER BY tglMasukPenunjangRajal DESC", conn)
        Try
            cmd = New MySqlCommand(query, conn)
            dr = cmd.ExecuteReader
            dgv1.Rows.Clear()
            Do While dr.Read
                dgv1.Rows.Add(dr.Item("noRekamedis"), dr.Item("noDaftar"), dr.Item("nmPasien"), dr.Item("kdUnitAsal"), dr.Item("unitAsal"),
                              dr.Item("tglLahir"), dr.Item("alamat"), dr.Item("noRegistrasiPARajal"), dr.Item("tglMasukPARajal"), dr.Item("statusPA"),
                              dr.Item("kdDokterPengirim"), dr.Item("namapetugasMedis"), dr.Item("kdDokterPemeriksaan"), dr.Item("dokterPemeriksa"), dr.Item("diagnosaKlinis"),
                              dr.Item("noTindakanPARajal"), dr.Item("totalTindakanPA"), dr.Item("statusPembayaran"), dr.Item("carabayar"), dr.Item("jenisKelamin"))
            Loop

            dr.Close()
            conn.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        'Call aturDGV()
    End Sub

    Sub tampilDataSudahDitindakAll()
        Dim query As String = ""
        Select Case txtKdInstalasi.Text
            Case "RJ"
                query = "SELECT kdTarif,tindakan,PPA,
                                tarif,statusTindakan,tglMulaiLayaniPasien,
                                tglSelesaiLayaniPasien,idDetailPARajal,noTindakanPARajal 
                           FROM vw_datadetailparajal
                          WHERE noTindakanPARajal = '" & noTindakanPA & "'
                          ORDER BY RIGHT(statusTindakan,3) ASC"
            Case "RI"
                query = "SELECT kdTarif,tindakan,PPA,
                                tarif,statusTindakan,tglMulaiLayaniPasien,
                                tglSelesaiLayaniPasien,idDetailPARanap,noTindakanPARanap
                           FROM vw_datadetailparanap 
                          WHERE noTindakanPARanap = '" & noTindakanPA & "'
                          ORDER BY RIGHT(statusTindakan,3) ASC"
            Case "PASIEN LUAR"
        End Select

        Try
            Call koneksiServer()
            cmd = New MySqlCommand(query, conn)
            dr = cmd.ExecuteReader
            dgv2.Rows.Clear()

            Select Case txtKdInstalasi.Text
                Case "RJ"
                    Do While dr.Read
                        dgv2.Rows.Add(dr.Item("kdTarif"), dr.Item("tindakan"), dr.Item("PPA"),
                                       dr.Item("tarif"), dr.Item("statusTindakan"), dr.Item("tglMulaiLayaniPasien"),
                                       dr.Item("tglSelesaiLayaniPasien"), dr.Item("idDetailPARajal"), dr.Item("noTindakanPARajal"))
                    Loop
                Case "RI"
                    Do While dr.Read
                        dgv2.Rows.Add(dr.Item("kdTarif"), dr.Item("tindakan"), dr.Item("PPA"),
                                       dr.Item("tarif"), dr.Item("statusTindakan"), dr.Item("tglMulaiLayaniPasien"),
                                       dr.Item("tglSelesaiLayaniPasien"), dr.Item("idDetailPARanap"), dr.Item("noTindakanPARanap"))
                    Loop
                Case "PASIEN LUAR"
            End Select

            dr.Close()
            conn.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub

    Sub tampilJk()
        Call koneksiServer()
        Dim cmd As MySqlCommand
        Dim dr As MySqlDataReader
        Dim query As String
        query = "SELECT jenisKelamin 
                   FROM t_pasien
                  WHERE noRekamedis = '" & txtNoRM.Text & "'"
        cmd = New MySqlCommand(query, conn)

        dr = cmd.ExecuteReader

        While dr.Read
            jk = dr.GetString("jenisKelamin").ToString
        End While
    End Sub

    Sub tampilBahanPA()
        Call koneksiServer()

        Dim query As String
        query = "SELECT noPALama, noPA, bahan, lokalisasi 
                   FROM (SELECT * FROM t_registrasipatologirajal UNION SELECT * FROM t_registrasipatologiranap) AS U
                  WHERE U.noRegistrasiPARajal = '" & txtNoPermintaan.Text & "'"
        cmd = New MySqlCommand(query, conn)

        dr = cmd.ExecuteReader
        dr.Read()

        If dr.HasRows Then
            txtNoPALama.Text = dr.Item("noPALama").ToString
            txtNoPABaru.Text = dr.Item("noPA").ToString
            txtBahan.Text = dr.Item("bahan").ToString
            txtLokalisasi.Text = dr.Item("lokalisasi").ToString
        End If
    End Sub

    Sub aturDGV()
        Try
            dgv1.Columns(0).Width = 100
            dgv1.Columns(1).Width = 160
            dgv1.Columns(2).Width = 220
            dgv1.Columns(3).Width = 150
            dgv1.Columns(4).Width = 120
            dgv1.Columns(5).Width = 150
            dgv1.Columns(6).Width = 150
            dgv1.Columns(7).Width = 150
            dgv1.Columns(8).Width = 170
            dgv1.Columns(9).Width = 150
            dgv1.Columns(10).Width = 100
            dgv1.Columns(11).Width = 200
            dgv1.Columns(12).Width = 100
            dgv1.Columns(13).Width = 300
            dgv1.Columns(14).Width = 150
            dgv1.Columns(15).Width = 100
            dgv1.Columns(16).Width = 100
            dgv1.Columns(17).Width = 150
            dgv1.Columns(18).Width = 100
            dgv1.Columns(0).HeaderText = "No.RM"
            dgv1.Columns(1).HeaderText = "No.Daftar"
            dgv1.Columns(2).HeaderText = "Nama Pasien"
            dgv1.Columns(3).HeaderText = "KD.Unit Asal"
            dgv1.Columns(4).HeaderText = "Asal Ruang/Poli"
            dgv1.Columns(5).HeaderText = "Tgl.Lahir"
            dgv1.Columns(6).HeaderText = "Alamat"
            dgv1.Columns(7).HeaderText = "No.Permintaan"
            dgv1.Columns(8).HeaderText = "Tgl.Masuk PA"
            dgv1.Columns(9).HeaderText = "Status Tindakan"
            dgv1.Columns(10).HeaderText = "KD.Dokter"
            dgv1.Columns(11).HeaderText = "Dokter Pengirim"
            dgv1.Columns(12).HeaderText = "KD.Dokter"
            dgv1.Columns(13).HeaderText = "Dokter Patologi"
            dgv1.Columns(14).HeaderText = "Ket.Klinis"
            dgv1.Columns(15).HeaderText = "No.Tindakan"
            dgv1.Columns(16).HeaderText = "Total"
            dgv1.Columns(17).HeaderText = "Status Pembayaran"
            dgv1.Columns(18).HeaderText = "Cara Bayar"
            dgv1.Columns(19).HeaderText = "Jenis Kelamin"

            dgv1.Columns(1).Visible = False
            dgv1.Columns(3).Visible = False
            dgv1.Columns(5).Visible = False
            dgv1.Columns(6).Visible = False
            dgv1.Columns(7).Visible = False
            dgv1.Columns(10).Visible = False
            dgv1.Columns(12).Visible = False
            dgv1.Columns(13).Visible = False
            dgv1.Columns(14).Visible = False
            dgv1.Columns(15).Visible = False
            dgv1.Columns(16).Visible = False
            dgv1.Columns(17).Visible = True
            dgv1.Columns(18).Visible = True
            dgv1.Columns(19).Visible = False

        Catch ex As Exception

        End Try
    End Sub

    Sub updateRegistrasiPARanap()
        Call koneksiServer()

        Dim dt As String
        dt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

        Try
            Dim str As String
            str = "UPDATE t_registrasipatologiranap 
                      SET noPALama = '" & FormVerifikasi.txtPALama.Text & "',
                          noPA = '" & FormVerifikasi.txtPABaru.Text & "',
                          tglMulaiLayaniPasien = '" & dt & "', 
                          statusPA = 'DALAM TINDAKAN', 
                          kdDokterPemeriksaan = '0069' 
                    WHERE noRegistrasiPARanap = '" & txtNoPermintaan.Text & "'"

            cmd = New MySqlCommand(str, conn)
            cmd.ExecuteNonQuery()
            cmd.Parameters.Clear()
            'MsgBox("Update data Registrasi Pemeriksaan Lab berhasil dilakukan", MessageBoxIcon.Information)
        Catch ex As Exception
            MsgBox("Update data Registrasi gagal dilakukan. " & ex.Message, MessageBoxIcon.Error, "Error Registrasi Ranap")
        End Try

        conn.Close()
    End Sub

    Sub updateRegistrasiPARajal()
        Call koneksiServer()

        Dim dt As String
        dt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

        Try
            Dim str As String
            str = "UPDATE t_registrasipatologirajal 
                      SET noPALama = '" & FormVerifikasi.txtPALama.Text & "',
                          noPA = '" & FormVerifikasi.txtPABaru.Text & "',
                          tglMulaiLayaniPasien = '" & dt & "',
                          statusPA = 'DALAM TINDAKAN', 
                          kdDokterPemeriksaan = '0069' 
                    WHERE noRegistrasiPARajal = '" & txtNoPermintaan.Text & "'"

            cmd = New MySqlCommand(str, conn)
            cmd.ExecuteNonQuery()
            cmd.Parameters.Clear()
            'MsgBox("Update data Registrasi Pemeriksaan Lab berhasil dilakukan", MessageBoxIcon.Information)
        Catch ex As Exception
            MsgBox("Update data Registrasi gagal dilakukan. " & ex.Message, MessageBoxIcon.Error, "Error Registrasi Rajal")
        End Try

        conn.Close()
    End Sub

    Sub updateTglTindakan()
        Dim dt As String
        dt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

        Dim str As String = ""

        Select Case txtInstalasi.Text
            Case "RAWAT JALAN"
                str = "UPDATE t_tindakanpatologirajal
                          SET tglTindakanPA= '" & dt & "'
                        WHERE noTindakanPARajal = '" & noTindakanPA & "'"
            Case "RAWAT INAP"
                str = "UPDATE t_tindakanpatologiranap 
                          SET tglTindakanPA = '" & dt & "'
                        WHERE noTindakanPARanap = '" & noTindakanPA & "'"
            Case "PASIEN LUAR"
        End Select

        Call koneksiServer()
        Try
            cmd = New MySqlCommand(str, conn)
            cmd.ExecuteNonQuery()
            'MsgBox("Update Tanggal Tindakan Lab berhasil dilakukan", MessageBoxIcon.Information)
        Catch ex As Exception
            MsgBox("Update Tanggal gagal dilakukan. " & ex.Message, MessageBoxIcon.Error)
        End Try

        conn.Close()
    End Sub

    Sub updateDokterDetail()
        Dim str As String = ""

        Select Case txtInstalasi.Text
            Case "RAWAT JALAN"
                str = "UPDATE t_detailtindakanpatologirajal
                          SET statusTindakan = 'DALAM TINDAKAN' 
                        WHERE idDetailPARajal = '" & idDetail & "'"
            Case "RAWAT INAP"
                str = "UPDATE t_detailtindakanpatologiranap
                          SET statusTindakan = 'DALAM TINDAKAN' 
                        WHERE idDetailPARanap = '" & idDetail & "'"
            Case "PASIEN LUAR"
        End Select

        Call koneksiServer()
        Try
            cmd = New MySqlCommand(str, conn)
            cmd.ExecuteNonQuery()
            'MsgBox("Update Dokter Lab berhasil dilakukan", MessageBoxIcon.Information)
            MsgBox("Tindakan Pemeriksaan dimulai", MsgBoxStyle.Information, "Informasi")
        Catch ex As Exception
            MsgBox("Update Dokter PA gagal dilakukan. " & ex.Message, MessageBoxIcon.Error)
        End Try

        conn.Close()
    End Sub

    Sub updateTglSelesaiTindakan()
        Dim dt As String
        dt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

        Dim str As String = ""

        Select Case txtInstalasi.Text
            Case "RAWAT JALAN"
                str = "UPDATE t_registrasipatologirajal
                          SET tglSelesaiLayaniPasien= '" & dt & "'
                        WHERE noRegistrasiPARajal = '" & txtNoPermintaan.Text & "'"
            Case "RAWAT INAP"
                str = "UPDATE t_registrasipatologiranap 
                          SET tglSelesaiLayaniPasien = '" & dt & "'
                        WHERE noRegistrasiPARanap = '" & txtNoPermintaan.Text & "'"
            Case "PASIEN LUAR"
        End Select

        Call koneksiServer()
        Try
            cmd = New MySqlCommand(str, conn)
            cmd.ExecuteNonQuery()
            'MsgBox("Update Tanggal Tindakan Lab berhasil dilakukan", MessageBoxIcon.Information)
        Catch ex As Exception
            MsgBox("Update 'Tanggal Selesai' gagal dilakukan. " & ex.Message, MessageBoxIcon.Error)
        End Try

        conn.Close()
    End Sub

    Public Sub ClickMulai()

        If txtInstalasi.Text = "RAWAT JALAN" Then
            Call updateRegistrasiPARajal()
            Call updateTglTindakan()
            Call updateDokterDetail()
            FormVerifikasi.Close()
        Else
            Select Case txtInstalasi.Text
                Case "RAWAT INAP"
                    Call updateRegistrasiPARanap()
                    Call updateTglTindakan()
                    Call updateDokterDetail()
                    FormVerifikasi.Close()
                    'Case "PASIEN LUAR"
                    '    Call updateRegistrasiPALuar()
                    '    Call updateTglTindakan()
                    '    Call updateDokterDetail()
            End Select
            'Try
            '    client = New TcpClient(txtIpAddres2.Text, 8000)     'IP tujuan
            '    Dim writer As New StreamWriter(client.GetStream())
            '    writer.Write(txtIpAddress.Text)                     'IP pengirim
            '    writer.Flush()
            '    'ListBox1.Items.Add("Me:- " + TextBox1.Text)
            '    'TextBox1.Clear()
            'Catch ex As Exception
            '    MsgBox(ex.Message)
            'End Try
        End If
        Call tampilDataSudahDitindakAll()
    End Sub

    Public Sub ClickSelesai(id As String)
        Dim str As String = ""

        Select Case txtInstalasi.Text
            Case "RAWAT JALAN"
                str = "UPDATE t_detailtindakanpatologirajal 
                          SET statusTindakan = 'SELESAI'
                        WHERE idDetailPARajal = '" & id & "'"
            Case "RAWAT INAP"
                str = "UPDATE t_detailtindakanpatologiranap 
                          SET statusTindakan = 'SELESAI'
                        WHERE idDetailPARanap = '" & id & "'"
            Case "PASIEN LUAR"
                str = "UPDATE t_detailtindakanpatologi
                          SET statusTindakan = 'SELESAI'
                        WHERE noRegistrasiPAg = '" & id & "'"
        End Select

        Call koneksiServer()
        Try
            cmd = New MySqlCommand(str, conn)
            cmd.ExecuteNonQuery()
            cmd.Parameters.Clear()
            MsgBox("Pemeriksaan PA selesai dilakukan", MessageBoxIcon.Information)
        Catch ex As Exception
            MsgBox("Error" & ex.Message, MessageBoxIcon.Error)
        End Try

        Call updateTglSelesaiTindakan()
        Call tampilDataAll()
        Call tampilDataSudahDitindakAll()

        conn.Close()

        'Hasil.Ambil_Data = True
        'Hasil.Form_Ambil_Data = "Hasil"
        'Hasil.Show()
    End Sub

    Public Sub ClickHasil()
        Form3.Ambil_Data = True
        Form3.Form_Ambil_Data = "Hasil"
        Form3.Show()
        Me.Hide()
    End Sub

    Sub updateStatusRegPermintaan()
        Dim str As String = ""

        Select Case txtKdInstalasi.Text
            Case "RJ"
                str = "UPDATE t_registrasipatologirajal
                          SET statusPA = 'SELESAI'
                        WHERE noRegistrasiPARajal = '" & txtNoPermintaan.Text & "'"
            Case "RI"
                str = "UPDATE t_registrasipatologiranap 
                          SET statusPA = 'SELESAI'
                        WHERE noRegistrasiPARanap = '" & txtNoPermintaan.Text & "'"
            Case "PASIEN LUAR"
        End Select

        Call koneksiServer()
        Try
            cmd = New MySqlCommand(str, conn)
            cmd.ExecuteNonQuery()
            'MsgBox("Update Tanggal Tindakan Lab berhasil dilakukan", MessageBoxIcon.Information)
        Catch ex As Exception
            MsgBox("Update 'Status Selesai' gagal dilakukan.", MessageBoxIcon.Error)
        End Try

        conn.Close()
    End Sub

    Sub cekStatusSelesai()
        Dim rowCount As Integer = 0
        rowCount = dgv2.Rows.Count

        Dim itemCount As Integer
        For i As Integer = 0 To dgv2.Rows.Count - 1
            If dgv2.Rows(i).Cells(4).Value.ToString = "SELESAI" Then
                itemCount = itemCount + 1
            End If
        Next

        'MsgBox("Jumlah tindakan : " & rowCount)
        'MsgBox("Jumlah tindakan yg selesai : " & itemCount)

        If itemCount = rowCount Then
            Call updateStatusRegPermintaan()
            Call tampilDataAll()
            'MsgBox("Update Status")
            'Else
            'MsgBox("Masih ada tindakan yang belum terselesaikan")
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load

        Me.WindowState = FormWindowState.Normal
        Me.StartPosition = FormStartPosition.Manual
        With Screen.PrimaryScreen.WorkingArea
            Me.SetBounds(.Left, .Top, .Width, .Height)
        End With

        If ApplicationDeployment.IsNetworkDeployed Then
            Dim ver As ApplicationDeployment = ApplicationDeployment.CurrentDeployment
            Label9.Text = "Version " & ver.CurrentVersion.ToString()
        End If

        pnlStats.Height = btnDash.Height
        pnlStats.Top = btnDash.Top
        btnDash.BackColor = Color.DodgerBlue

        Dim pcname As String
        Dim ipadd As String = ""
        pcname = System.Net.Dns.GetHostName

        Dim objAddressList() As System.Net.IPAddress = System.Net.Dns.GetHostEntry("").AddressList
        For i = 0 To objAddressList.GetUpperBound(0)
            If objAddressList(i).AddressFamily = Net.Sockets.AddressFamily.InterNetwork Then
                ipadd = objAddressList(i).ToString
                'txtHostname.Text = objAddressList(i).ToString
                Exit For
            End If
        Next

        txtHostname.Text = "PC Name : " & pcname & " | IP Address : " & ipadd & " | Username : " & LoginForm.user

        Call tampilDataAll()
        Timer5.Start()
        'Call aturDGV()
    End Sub

    Private Sub dgv1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgv1.CellContentClick
        Dim noRm, noReg, namaPasien, alamat, usia, dokter1, kddokter2, dokter2, noPermin, tglReg, ketKlinis, unitAsal As String

        If e.RowIndex = -1 Then
            Return
        End If

        btnHasil.Enabled = True

        noRm = dgv1.Rows(e.RowIndex).Cells(0).Value
        noReg = dgv1.Rows(e.RowIndex).Cells(1).Value
        namaPasien = dgv1.Rows(e.RowIndex).Cells(2).Value
        unitAsal = dgv1.Rows(e.RowIndex).Cells(4).Value
        usia = dgv1.Rows(e.RowIndex).Cells(5).Value
        alamat = dgv1.Rows(e.RowIndex).Cells(6).Value
        noPermin = dgv1.Rows(e.RowIndex).Cells(7).Value
        tglReg = dgv1.Rows(e.RowIndex).Cells(8).Value
        dokter1 = dgv1.Rows(e.RowIndex).Cells(11).Value
        kddokter2 = dgv1.Rows(e.RowIndex).Cells(12).Value.ToString
        dokter2 = dgv1.Rows(e.RowIndex).Cells(13).Value.ToString
        ketKlinis = dgv1.Rows(e.RowIndex).Cells(14).Value.ToString
        noTindakanPA = dgv1.Rows(e.RowIndex).Cells(15).Value.ToString
        caraBayar = dgv1.Rows(e.RowIndex).Cells(18).Value.ToString

        txtNoRM.Text = noRm
        txtNoReg.Text = noReg
        txtNamaPasien.Text = namaPasien
        txtAlamat.Text = alamat
        txtNoPermintaan.Text = noPermin
        txtUnitAsal.Text = unitAsal
        txtTglReg.Text = tglReg
        txtTglLahir.Text = usia
        txtDokter.Text = dokter1
        txtKdDokterPA.Text = kddokter2
        txtDokPA.Text = dokter2
        txtKlinis.Text = ketKlinis

        Call tampilJk()
        Call tampilBahanPA()

        If jk = "L" Then
            txtJK.Text = "Laki-Laki"
        ElseIf jk = "P" Then
            txtJK.Text = "Perempuan"
        End If

        txtKdInstalasi.Text = txtNoPermintaan.Text.Substring(0, 2)
        Select Case txtKdInstalasi.Text
            Case "RJ"
                txtInstalasi.Text = "RAWAT JALAN"
            Case "RI"
                txtInstalasi.Text = "RAWAT INAP"
        End Select

        'dgv2.Columns.Clear()

        Call tampilDataSudahDitindakAll()
        'Call tampilBahanPA()
        'Call tampilFiksasiPA()
        'Call tampilDataLab()
        dgv2.ClearSelection()
    End Sub

    Private Sub dgv1_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgv1.CellMouseClick
        Dim noRm, noReg, namaPasien, alamat, usia, dokter1, kddokter2, dokter2, noPermin, tglReg, ketKlinis, unitAsal As String

        If e.RowIndex = -1 Then
            Return
        End If

        btnHasil.Enabled = True

        noRm = dgv1.Rows(e.RowIndex).Cells(0).Value
        noReg = dgv1.Rows(e.RowIndex).Cells(1).Value
        namaPasien = dgv1.Rows(e.RowIndex).Cells(2).Value
        unitAsal = dgv1.Rows(e.RowIndex).Cells(4).Value
        usia = dgv1.Rows(e.RowIndex).Cells(5).Value
        alamat = dgv1.Rows(e.RowIndex).Cells(6).Value
        noPermin = dgv1.Rows(e.RowIndex).Cells(7).Value
        tglReg = dgv1.Rows(e.RowIndex).Cells(8).Value
        dokter1 = dgv1.Rows(e.RowIndex).Cells(11).Value
        kddokter2 = dgv1.Rows(e.RowIndex).Cells(12).Value.ToString
        dokter2 = dgv1.Rows(e.RowIndex).Cells(13).Value.ToString
        ketKlinis = dgv1.Rows(e.RowIndex).Cells(14).Value.ToString
        noTindakanPA = dgv1.Rows(e.RowIndex).Cells(15).Value.ToString
        caraBayar = dgv1.Rows(e.RowIndex).Cells(18).Value.ToString

        txtNoRM.Text = noRm
        txtNoReg.Text = noReg
        txtNamaPasien.Text = namaPasien
        txtAlamat.Text = alamat
        txtNoPermintaan.Text = noPermin
        txtUnitAsal.Text = unitAsal
        txtTglReg.Text = tglReg
        txtTglLahir.Text = usia
        txtDokter.Text = dokter1
        txtKdDokterPA.Text = kddokter2
        txtDokPA.Text = dokter2
        txtKlinis.Text = ketKlinis

        Call tampilJk()
        Call tampilBahanPA()

        If jk = "L" Then
            txtJK.Text = "Pria"
        ElseIf jk = "P" Then
            txtJK.Text = "Wanita"
        End If

        txtKdInstalasi.Text = txtNoPermintaan.Text.Substring(0, 2)
        Select Case txtKdInstalasi.Text
            Case "RJ"
                txtInstalasi.Text = "RAWAT JALAN"
            Case "RI"
                txtInstalasi.Text = "RAWAT INAP"
        End Select

        'dgv2.Columns.Clear()

        Call tampilDataSudahDitindakAll()
        'Call tampilBahanPA()
        'Call tampilFiksasiPA()
        'Call tampilDataLab()
        dgv2.ClearSelection()
    End Sub

    Private Sub btnKeluar_Click(sender As Object, e As EventArgs) Handles btnKeluar.Click
        Dim konfirmasi As MsgBoxResult

        konfirmasi = MsgBox("Apakah anda yakin ingin keluar..?", vbQuestion + vbYesNo, "EXIT")
        If konfirmasi = vbYes Then
            Me.Close()
            LoginForm.Show()
        End If
    End Sub

    Private Sub btnDash_Click(sender As Object, e As EventArgs) Handles btnDash.Click
        pnlStats.Height = btnDash.Height
        pnlStats.Top = btnDash.Top

        Dim btn As Button = CType(sender, Button)
        setColor(btn)
    End Sub

    Private Sub btnTindakan_Click(sender As Object, e As EventArgs) Handles btnTindakan.Click
        'pnlStats.Height = btnTindakan.Height
        'pnlStats.Top = btnTindakan.Top

        'Dim btn As Button = CType(sender, Button)
        'setColor(btn)
        'Form2.Show()
        Dim tindakan As Form2 = New Form2
        Me.Hide()
        Tindakan.Ambil_Data = True
        tindakan.Form_Ambil_Data = "Tindakan"
        tindakan.Show()
    End Sub

    Private Sub btnHasil_Click(sender As Object, e As EventArgs) Handles btnHasil.Click
        'pnlStats.Height = btnHasil.Height
        'pnlStats.Top = btnHasil.Top

        'Dim btn As Button = CType(sender, Button)
        'setColor(btn)

        'For i As Integer = 0 To dgv1.Rows.Count - 1
        '    If dgv1.Rows(i).Cells(9).Value = "PERMINTAAN" Then
        '        MsgBox("Tindakan sedang dalam status 'PERMINTAAN'")
        '    ElseIf dgv1.Rows(i).Cells(9).Value = "DALAM TINDAKAN" Then
        '        MsgBox("Tindakan sedang 'DALAM TINDAKAN'")
        '    ElseIf dgv1.Rows(i).Cells(9).Value = "SELESAI" Then

        '    End If
        'Next

        Form3.Ambil_Data = True
        Form3.Form_Ambil_Data = "Hasil"
        Form3.Show()
        Me.Hide()
    End Sub

    Private Sub txtTglLahir_TextChanged(sender As Object, e As EventArgs) Handles txtTglLahir.TextChanged
        If Not String.IsNullOrEmpty(txtTglLahir.Text) Then
            Dim lahir As Date = txtTglLahir.Text
            txtUsia.Text = hitungUmur(lahir)
        Else
            Return
        End If
    End Sub

    Private Sub dgv2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgv2.CellContentClick
        Dim konfirmasi As MsgBoxResult
        Dim tindakan, dokPA As String
        idDetail = dgv2.Rows(e.RowIndex).Cells(7).Value.ToString
        tindakan = dgv2.Rows(e.RowIndex).Cells(1).Value.ToString
        dokPA = dgv2.Rows(e.RowIndex).Cells(2).Value.ToString

        If e.ColumnIndex = 9 Then
            Select Case dgv2.Rows(e.RowIndex).Cells(4).Value.ToString
                Case "PERMINTAAN"
                    'konfirmasi = MsgBox("Apakah tindakan '" & tindakan & "' akan dimulai ?", vbQuestion + vbYesNo, "Laboratorium")
                    'If konfirmasi = vbYes Then
                    'Call ClickMulai()
                    'MsgBox(tindakan & " - Memulai tindakan", MsgBoxStyle.Information)
                    FormVerifikasi.Ambil_Data = True
                    FormVerifikasi.Form_Ambil_Data = "Verifikasi"
                    FormVerifikasi.ShowDialog()
                    'End If
                Case "DALAM TINDAKAN"
                    konfirmasi = MsgBox("Apakah tindakan '" & tindakan & "' sudah selesai ?", vbQuestion + vbYesNo, "Laboratorium")
                    If konfirmasi = vbYes Then
                        Call ClickSelesai(idDetail)
                        Call cekStatusSelesai()
                        MsgBox(tindakan & " - Selesai", MsgBoxStyle.Information)
                    End If
                Case "SELESAI"
                    konfirmasi = MsgBox("Apakah pembayaran '" & tindakan & "' sudah lunas ?", vbQuestion + vbYesNo, "Laboratorium")
                    If konfirmasi = vbYes Then
                        Call ClickHasil()
                        MsgBox(tindakan & " - LUNAS", MsgBoxStyle.Information)
                    End If
            End Select
        End If

    End Sub

    Private Sub btnCari_Click(sender As Object, e As EventArgs) Handles btnCari.Click
        If txtNoRM.Text = "" Then
            MessageBox.Show("Masukkan No.RM Pasien", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            Timer5.Stop()
            Call cariDataPasien()
        End If
    End Sub

    Private Sub txtNoRM_KeyDown(sender As Object, e As KeyEventArgs) Handles txtNoRM.KeyDown
        If e.KeyCode = Keys.Enter And txtNoRM.Text = "" Then
            MessageBox.Show("Masukkan No.RM Pasien", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        ElseIf e.KeyCode = Keys.Enter Then
            Timer5.Stop()
            Call cariDataPasien()
        End If
    End Sub

    Private Sub txtDokter_TextChanged(sender As Object, e As EventArgs) Handles txtDokter.TextChanged
        Call koneksiServer()
        Try
            Dim query As String
            query = "SELECT * FROM t_tenagamedis2 WHERE namapetugasMedis = '" & txtDokter.Text & "'"
            cmd = New MySqlCommand(query, conn)
            dr = cmd.ExecuteReader

            While dr.Read
                txtKdDokter.Text = UCase(dr.GetString("kdPetugasMedis"))
            End While
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        conn.Close()
    End Sub

    Private Sub btnTambah_Click(sender As Object, e As EventArgs) Handles btnTambah.Click
        Dim jenisPasienFrm As New JenisPasien
        jenisPasienFrm.ShowDialog()
    End Sub

    Private Sub dgv1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles dgv1.CellFormatting
        dgv1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv1.ColumnHeadersDefaultCellStyle.BackColor = Color.DeepSkyBlue
        dgv1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White
        dgv1.ColumnHeadersDefaultCellStyle.Font = New Font("Tahoma", 10, FontStyle.Bold)
        dgv1.ColumnHeadersHeight = 35
        dgv1.DefaultCellStyle.Font = New Font("Tahoma", 8, FontStyle.Bold)
        dgv1.DefaultCellStyle.ForeColor = Color.Black
        dgv1.DefaultCellStyle.SelectionBackColor = Color.PaleTurquoise
        dgv1.DefaultCellStyle.SelectionForeColor = Color.Black
        dgv1.RowHeadersVisible = False
        dgv1.AllowUserToResizeRows = False
        dgv1.EnableHeadersVisualStyles = False

        dgv1.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv1.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv1.Columns(8).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv1.Columns(9).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv1.Columns(17).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv1.Columns(18).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        For i = 0 To dgv1.RowCount - 1
            If i Mod 2 = 0 Then
                dgv1.Rows(i).DefaultCellStyle.BackColor = Color.White
            Else
                dgv1.Rows(i).DefaultCellStyle.BackColor = Color.AliceBlue
            End If
        Next

        For Each column As DataGridViewColumn In dgv1.Columns
            column.SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        For i As Integer = 0 To dgv1.Rows.Count - 1
            If dgv1.Rows(i).Cells(9).Value = "PERMINTAAN" Then
                dgv1.Rows(i).Cells(9).Style.BackColor = Color.Orange
                dgv1.Rows(i).Cells(9).Style.ForeColor = Color.White
            ElseIf dgv1.Rows(i).Cells(9).Value = "DALAM TINDAKAN" Then
                dgv1.Rows(i).Cells(9).Style.BackColor = Color.Green
                dgv1.Rows(i).Cells(9).Style.ForeColor = Color.White
            ElseIf dgv1.Rows(i).Cells(9).Value = "SELESAI" Then
                dgv1.Rows(i).Cells(9).Style.BackColor = Color.Red
                dgv1.Rows(i).Cells(9).Style.ForeColor = Color.White
            End If
        Next

        For i As Integer = 0 To dgv1.Rows.Count - 1
            If dgv1.Rows(i).Cells(17).Value = "BELUM LUNAS" Then
                dgv1.Rows(i).Cells(17).Style.BackColor = Color.Orange
                dgv1.Rows(i).Cells(17).Style.ForeColor = Color.White
            ElseIf dgv1.Rows(i).Cells(17).Value = "LUNAS" Then
                dgv1.Rows(i).Cells(17).Style.BackColor = Color.Green
                dgv1.Rows(i).Cells(17).Style.ForeColor = Color.White
            End If
        Next
    End Sub

    Private Sub dgv2_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles dgv2.CellFormatting
        dgv2.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        dgv2.DefaultCellStyle.ForeColor = Color.Black
        dgv2.DefaultCellStyle.SelectionForeColor = Color.Black
        dgv2.Columns(3).DefaultCellStyle.Format = "###,###,###"
        dgv2.Columns(5).DefaultCellStyle.Format = "HH:mm"
        dgv2.Columns(6).DefaultCellStyle.Format = "HH:mm"
        dgv2.Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv2.Columns(2).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv2.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv2.Columns(4).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv2.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv2.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        dgv2.Columns(5).Visible = False
        dgv2.Columns(6).Visible = False
        dgv2.Columns(7).Visible = False

        For i = 0 To dgv2.RowCount - 1
            dgv2.Rows(i).Cells(9).Style.BackColor = Color.DodgerBlue
            dgv2.Rows(i).Cells(9).Style.ForeColor = Color.White
        Next

        For i = 0 To dgv2.RowCount - 1
            If i Mod 2 = 0 Then
                dgv2.Rows(i).DefaultCellStyle.BackColor = Color.White
            Else
                dgv2.Rows(i).DefaultCellStyle.BackColor = Color.AliceBlue
            End If
        Next

        For Each column As DataGridViewColumn In dgv2.Columns
            column.SortMode = DataGridViewColumnSortMode.NotSortable
        Next

        For i As Integer = 0 To dgv2.RowCount - 1
            If dgv2.Rows(i).Cells(4).Value.ToString = "PERMINTAAN" Then
                dgv2.Rows(i).Cells(4).Style.BackColor = Color.Orange
                dgv2.Rows(i).Cells(4).Style.ForeColor = Color.White
                'txtKdDokterPeriksa.Text = ""
                'btnTindak.Enabled = True
                'btnSelesai.Visible = False
            ElseIf dgv2.Rows(i).Cells(4).Value.ToString = "DALAM TINDAKAN" Then
                dgv2.Rows(i).Cells(4).Style.BackColor = Color.Green
                dgv2.Rows(i).Cells(4).Style.ForeColor = Color.White
                'btnTindak.Enabled = True
                'btnSelesai.Visible = True
            ElseIf dgv2.Rows(i).Cells(4).Value.ToString = "SELESAI" Then
                dgv2.Rows(i).Cells(4).Style.BackColor = Color.Red
                dgv2.Rows(i).Cells(4).Style.ForeColor = Color.White
                'dgv2.ClearSelection()
                'dgv2.Rows(i).Visible = False
                'btnTindak.Enabled = True
                'btnSelesai.Visible = False
            End If
        Next
    End Sub

    Private Sub btnCetak_Click(sender As Object, e As EventArgs) Handles btnCetak.Click
        viewCetakRincianBiaya.Ambil_Data = True
        viewCetakRincianBiaya.Form_Ambil_Data = "Cetak"
        viewCetakRincianBiaya.Show()
    End Sub

    Private Sub dgv1_CellMouseUp(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgv1.CellMouseUp
        If e.Button = Windows.Forms.MouseButtons.Right Then
            dgv1.Rows(e.RowIndex).Selected = True
            ind = e.RowIndex
            dgv1.CurrentCell = dgv1.Rows(e.RowIndex).Cells(2)
            ContextMenuStrip1.Show(dgv1, e.Location)
            ContextMenuStrip1.Show(Cursor.Position)
        End If
    End Sub

    Private Sub ContextMenuStrip1_Click(sender As Object, e As EventArgs) Handles ContextMenuStrip1.Click
        If Not dgv1.Rows(ind).IsNewRow Then
            Dim konfirmasi As MsgBoxResult
            konfirmasi = MsgBox("Apakah anda yakin ingin menghapus antrian tsb ?", vbQuestion + vbYesNo, "Konfirmasi")
            If konfirmasi = vbYes Then
                'MsgBox("Batal : " & DataGridView1.Rows(ind).Cells(7).Value.ToString)
                If dgv1.Rows(ind).Cells(9).Value.ToString = "DALAM TINDAKAN" Then
                    MsgBox("Tindakan tidak dapat dihapus, karena status 'DALAM TINDAKAN'")
                ElseIf dgv1.Rows(ind).Cells(9).Value.ToString = "SELESAI" Then
                    MsgBox("Tindakan tidak dapat dihapus, karena status 'SELESAI'")
                Else
                    Call deletePermintaan(dgv1.Rows(ind).Cells(7).Value.ToString)
                    Call deleteTindakan(dgv1.Rows(ind).Cells(15).Value.ToString)
                    Call deleteAllDetail(dgv1.Rows(ind).Cells(15).Value.ToString)
                    Call tampilDataAll()
                End If
            End If
        End If
    End Sub

    Sub deletePermintaan(noDel As String)
        Try
            Call koneksiServer()
            Dim query As String = ""

            Select Case txtKdInstalasi.Text
                Case "RJ"
                    query = "DELETE FROM t_registrasipatologirajal WHERE noRegistrasiPARajal= '" & noDel & "'"
                Case "RI"
                    query = "DELETE FROM t_registrasipatologiranap WHERE noRegistrasiPARanap= '" & noDel & "'"
                Case "IGD"
                    query = "DELETE FROM t_registrasipatologirajal WHERE noRegistrasiPARajal= '" & noDel & "'"
            End Select

            cmd = New MySqlCommand(query, conn)
            cmd.ExecuteNonQuery()
            conn.Close()
            MsgBox("Batal antrian berhasil", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error Delete Antrian")
        End Try
    End Sub

    Sub deleteTindakan(noDel As String)
        Try
            Call koneksiServer()
            Dim query As String = ""

            Select Case txtKdInstalasi.Text
                Case "RJ"
                    query = "DELETE FROM t_tindakanpatologirajal WHERE noTindakanPARajal= '" & noDel & "'"
                Case "RI"
                    query = "DELETE FROM t_tindakanpatologiranap WHERE noTindakanPARanap= '" & noDel & "'"
                Case "IGD"
                    query = "DELETE FROM t_tindakanpatologirajal WHERE noTindakanPARajal= '" & noDel & "'"
            End Select

            cmd = New MySqlCommand(query, conn)
            cmd.ExecuteNonQuery()
            conn.Close()
            'MsgBox("Batal tindakan berhasil", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error Delete Tindakan")
        End Try
    End Sub

    Sub deleteAllDetail(noDel As String)
        Try
            Call koneksiServer()
            Dim query As String = ""

            Select Case txtKdInstalasi.Text
                Case "RJ"
                    query = "DELETE FROM t_detailtindakanpatologirajal WHERE noTindakanPARajal= '" & noDel & "'"
                Case "RI"
                    query = "DELETE FROM t_detailtindakanpatologiranap WHERE noTindakanPARanap= '" & noDel & "'"
                Case "IGD"
                    query = "DELETE FROM t_detailtindakanpatologirajal WHERE noTindakanPARajal= '" & noDel & "'"
            End Select

            cmd = New MySqlCommand(query, conn)
            cmd.ExecuteNonQuery()
            conn.Close()
            'MsgBox("Batal detail berhasil", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error Delete Detail")
        End Try
    End Sub

    Private Sub DataGridView2_CellMouseUp(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgv2.CellMouseUp
        If e.Button = Windows.Forms.MouseButtons.Right Then
            dgv2.Rows(e.RowIndex).Selected = True
            ind = e.RowIndex
            dgv2.CurrentCell = dgv2.Rows(e.RowIndex).Cells(1)
            ContextMenuStrip2.Show(dgv2, e.Location)
            ContextMenuStrip2.Show(Cursor.Position)
        End If
    End Sub

    Private Sub ContextMenuStrip2_Click(sender As Object, e As EventArgs) Handles ContextMenuStrip2.Click
        If Not dgv2.Rows(ind).IsNewRow Then
            Dim konfirmasi As MsgBoxResult
            Dim tarif, noTindakan As String
            tarif = dgv2.Rows(ind).Cells(3).Value.ToString
            noTindakan = dgv2.Rows(ind).Cells(8).Value.ToString
            konfirmasi = MsgBox("Apakah anda yakin ingin menghapus tindakan tsb ?", vbQuestion + vbYesNo, "Konfirmasi")
            If konfirmasi = vbYes Then
                'MsgBox("Batal : " & DataGridView2.Rows(ind).Cells(7).Value.ToString)
                If dgv2.Rows(ind).Cells(4).Value.ToString = "DALAM TINDAKAN" Then
                    MsgBox("Tindakan tidak dapat dihapus, karena status 'DALAM TINDAKAN'")
                ElseIf dgv2.Rows(ind).Cells(4).Value.ToString = "SELESAI" Then
                    MsgBox("Tindakan tidak dapat dihapus, karena status 'SELESAI'")
                Else
                    'Call deleteDetail(dgv2.Rows(ind).Cells(7).Value.ToString)
                    Call updateAfterDelete(noTindakan, tarif)
                    Call tampilDataSudahDitindakAll()
                End If
            End If
        End If
    End Sub

    Sub deleteDetail(idDel As String)
        Try
            Call koneksiServer()
            Dim query As String = ""

            Select Case txtKdInstalasi.Text
                Case "RJ"
                    query = "DELETE FROM t_detailtindakanpatologirajal WHERE idDetailPARajal= '" & idDel & "'"
                Case "RI"
                    query = "DELETE FROM t_detailtindakanpatologiranap WHERE idDetailPARanap= '" & idDel & "'"
                Case "IGD"
                    query = "DELETE FROM t_detailtindakanpatologirajal WHERE idDetailPARajal= '" & idDel & "'"
            End Select

            cmd = New MySqlCommand(query, conn)
            cmd.ExecuteNonQuery()
            conn.Close()
            MsgBox("Hapus tindakan berhasil", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error Delete")
        End Try
    End Sub

    Sub updateAfterDelete(noTindakanDel As String, tarif As String)
        Dim total As Integer
        total = Val(totalTarif() - tarif)
        MsgBox(total)
        Call koneksiServer()
        Dim str As String = ""

        Select Case txtKdInstalasi.Text
            Case "RJ"
                str = "UPDATE t_tindakanpatologirajal
                              SET totalTindakanPA = '" & total & "'
                            WHERE noTindakanPARajal = '" & noTindakanDel & "'"
            Case "RI"
                str = "UPDATE t_tindakanpatologiranap
                              SET totalTindakanPA = '" & total & "'
                            WHERE noTindakanPARanap = '" & noTindakanDel & "'"
            Case "IGD"
                str = "UPDATE t_tindakanpatologirajal
                              SET totalTindakanPA = '" & total & "'
                            WHERE noTindakanPARajal = '" & noTindakanDel & "'"
        End Select
        MsgBox(str)
        Try
            cmd = New MySqlCommand(str, conn)
            cmd.ExecuteNonQuery()
            'MsgBox("Update dokter berhasil dilakukan", MessageBoxIcon.Information)
        Catch ex As Exception
            MsgBox("Update data gagal dilakukan.", MessageBoxIcon.Error, "Error Update After Delete")
        End Try

        conn.Close()
    End Sub

    Function totalTarif() As String
        Dim total As Long = 0
        Dim hasil As Long = 0
        For i As Integer = 0 To dgv2.Rows.Count - 1
            total = total + Val(dgv2.Rows(i).Cells(3).Value)
        Next
        hasil = total.ToString("#,##0")
        Return hasil
    End Function

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        Timer5.Start()
        'Call tampilDataAll()
        'dgv1.ClearSelection()
    End Sub

    Private Sub Timer5_Tick(sender As Object, e As EventArgs) Handles Timer5.Tick
        Call tampilDataAll()
        dgv1.ClearSelection()
    End Sub

    Private Sub btnFilter_Click(sender As Object, e As EventArgs) Handles btnFilter.Click
        Call koneksiServer()
        Dim query As String
        query = "SELECT * FROM vw_pasienpatologi
                  WHERE tglMasukPARajal BETWEEN '" & Format(DateTimePicker2.Value, "yyyy-MM-dd") & "' 
                    AND '" & Format(DateAdd(DateInterval.Day, 1, DateTimePicker3.Value), "yyyy-MM-dd") & "'
                  ORDER BY tglMasukPARajal DESC"
        'da = New MySqlDataAdapter("SELECT * FROM vw_pasienpatologi
        '                            WHERE tglMasukPARajal BETWEEN '2020-04-01' AND '2020-04-06'
        '                            ORDER BY tglMasukPenunjangRajal DESC", conn)
        Try
            cmd = New MySqlCommand(query, conn)
            dr = cmd.ExecuteReader
            dgv1.Rows.Clear()
            Do While dr.Read
                dgv1.Rows.Add(dr.Item("noRekamedis"), dr.Item("noDaftar"), dr.Item("nmPasien"), dr.Item("kdUnitAsal"), dr.Item("unitAsal"),
                              dr.Item("tglLahir"), dr.Item("alamat"), dr.Item("noRegistrasiPARajal"), dr.Item("tglMasukPARajal"), dr.Item("statusPA"),
                              dr.Item("kdDokterPengirim"), dr.Item("namapetugasMedis"), dr.Item("kdDokterPemeriksaan"), dr.Item("dokterPemeriksa"), dr.Item("diagnosaKlinis"),
                              dr.Item("noTindakanPARajal"), dr.Item("totalTindakanPA"), dr.Item("statusPembayaran"), dr.Item("carabayar"), dr.Item("jenisKelamin"))
            Loop

            dr.Close()
            conn.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
    End Sub
End Class
