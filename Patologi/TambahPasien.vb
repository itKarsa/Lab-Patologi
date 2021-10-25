Imports System.ComponentModel
Imports MySql.Data.MySqlClient
Public Class TambahPasien

    Public Ambil_Data As String
    Public Form_Ambil_Data As String
    Public jk As String = ""
    Sub autoDokter()
        Call koneksiServer()

        Using cmd As New MySqlCommand("SELECT DISTINCT namapetugasMedis FROM t_tenagamedis2 WHERE kdKelompokTenagaMedis in('ktm1') ORDER BY namapetugasMedis ASC", conn)
            da = New MySqlDataAdapter(cmd)
            Dim dt As New DataTable
            da.Fill(dt)

            txtDokter.DataSource = dt
            txtDokter.DisplayMember = "namapetugasMedis"
            txtDokter.ValueMember = "namapetugasMedis"
            txtDokter.AutoCompleteMode = AutoCompleteMode.SuggestAppend
            txtDokter.AutoCompleteSource = AutoCompleteSource.ListItems
        End Using
        conn.Close()
    End Sub

    Sub addRegPARanap()
        Dim dt As String
        dt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

        Call koneksiServer()
        Try
            Dim str As String
            str = "INSERT INTO t_registrasipatologiranap(noRegistrasiPARanap,noDaftar,
                                                         kdUnitAsal,unitAsal,kdUnit,unit,
                                                         tglMasukPARanap,statusPA,
                                                         kdDokterPengirim,noPA,
                                                         userModify,dateModify) 
                   VALUES ('" & txtNoPermintaan.Text & "','" & txtNoReg.Text & "',
                           '" & txtKdUnit.Text & "','" & txtUnitAsal.Text & "',
                           '3005','Patologi Anatomi',
                           '" & dt & "','PERMINTAAN',
                           '" & txtKdDokter.Text & "','" & txtNoPA.Text & "',
                           ';" & LoginForm.txtUsername.Text & "',';" & Format(DateTime.Now, "dd/MM/yyyy HH:mm:ss") & "')"
            cmd = New MySqlCommand(str, conn)
            cmd.ExecuteNonQuery()
            MsgBox("Insert data Permintaan PA berhasil dilakukan", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox("Insert data Permintaan PA gagal dilakukan.", MsgBoxStyle.Critical, "Error")
        End Try
        conn.Close()
    End Sub

    Sub addRegPARajal()
        Dim dt As String
        dt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

        Call koneksiServer()
        Try
            Dim str As String
            str = "INSERT INTO t_registrasipatologirajal(noRegistrasiPARajal,noDaftar,
                                                          kdUnitAsal,unitAsal,kdUnit,unit,
                                                          tglMasukPARajal,statusPA,
                                                          kdDokterPengirim,noPA,
                                                          userModify,dateModify) 
                   VALUES ('" & txtNoPermintaan.Text & "','" & txtNoReg.Text & "',
                           '" & txtKdUnit.Text & "','" & txtUnitAsal.Text & "',
                           '3005','Patologi Anatomi',
                           '" & dt & "','PERMINTAAN',
                           '" & txtKdDokter.Text & "','" & txtNoPA.Text & "',
                           ';" & LoginForm.txtUsername.Text & "',';" & Format(DateTime.Now, "dd/MM/yyyy HH:mm:ss") & "')"
            cmd = New MySqlCommand(str, conn)
            cmd.ExecuteNonQuery()
            MsgBox("Insert data Permintaan PA berhasil dilakukan", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox("Insert data Permintaan PA gagal dilakukan.", MsgBoxStyle.Critical, "Error")
        End Try
        conn.Close()
    End Sub

    Sub cariNoDaftar()

        conn.Close()
        Dim query As String = ""
        query = "SELECT * FROM t_registrasi WHERE noRekamedis LIKE '%" & txtNoRM.Text & "%' ORDER BY tglDaftar DESC LIMIT 1"
        cmd = New MySqlCommand(query, conn)
        da = New MySqlDataAdapter(cmd)

        Dim str As New DataTable
        str.Clear()
        da.Fill(str)

        txtNoReg.Text = ""
        txtNamaPasien.Text = ""
        txtUmurJk.Text = ""
        txtAlamat.Text = ""
        txtTglLahir.Text = ""
        txtDokter.Text = ""
        txtKdDokter.Text = ""
        txtNoPermintaan.Text = ""

        If str.Rows.Count() > 0 Then
            txtNoReg.Text = str.Rows(0)(0).ToString
            txtKdInst.Text = str.Rows(0)(9).ToString
        Else
            MessageBox.Show("Pasien Tidak Ada / Belum Terdaftar", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
        conn.Close()
    End Sub

    Sub cariDataPasien()
        conn.Close()
        Try
            Dim query As String = ""
            query = "SELECT * 
                       FROM vw_datapasien
                      WHERE noRekamedis Like '%" & txtNoRM.Text & "%'"
            cmd = New MySqlCommand(query, conn)
            da = New MySqlDataAdapter(cmd)

            Dim str As New DataTable
            str.Clear()
            da.Fill(str)

            txtNamaPasien.Text = ""
            txtTglLahir.Text = ""
            txtUmurJk.Text = ""
            txtAlamat.Text = ""

            If str.Rows.Count() > 0 Then
                txtNoRM.Text = str.Rows(0)(0).ToString
                txtNamaPasien.Text = str.Rows(0)(1).ToString
                txtTglLahir.Text = str.Rows(0)(4).ToString

                If str.Rows(0)(5).ToString.Contains("L") Then
                    jk = "LAKI-LAKI"
                ElseIf str.Rows(0)(5).ToString.Contains("P") Then
                    jk = "PEREMPUAN"
                End If

                txtUmurJk.Text = hitungUmur(Convert.ToDateTime(txtTglLahir.Text)) & " / " & jk
                txtAlamat.Text = str.Rows(0)(8).ToString

                Call cariInst()
                Call cariUnit()
                Call cariKelas()
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MessageBoxIcon.Exclamation, "SELECT DATA PASIEN")
        End Try
        conn.Close()
    End Sub

    Sub cariInst()
        Call koneksiServer()
        Try
            Dim cmdInst As MySqlCommand
            Dim daInst As MySqlDataAdapter
            Dim queryInst As String = ""

            queryInst = "SELECT instalasi FROM t_instalasiunit WHERE kdInstalasi = '" & txtKdInst.Text & "'"
            cmdInst = New MySqlCommand(queryInst, conn)
            daInst = New MySqlDataAdapter(cmdInst)

            Dim str As New DataTable
            str.Clear()
            daInst.Fill(str)

            txtInst.Text = ""
            If str.Rows.Count() > 0 Then
                txtInst.Text = str.Rows(0)(0).ToString
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MessageBoxIcon.Exclamation, "SELECT INSTALASI")
        End Try
        conn.Close()
    End Sub

    Sub cariUnit()
        Call koneksiServer()
        Dim cmdUnit As MySqlCommand
        Dim daUnit As MySqlDataAdapter
        Dim queryUnit As String = ""

        If txtInst.Text.Contains("IGD") Then
            queryUnit = "SELECT kdUnit, unit, namapetugasMedis FROM vw_pasienrawatjalan WHERE noDaftar = '" & txtNoReg.Text & "'"
        ElseIf txtInst.Text.Contains("RAWAT INAP") Then
            queryUnit = "SELECT kdRawatInap, rawatInap, namapetugasMedis FROM vw_pasienrawatinap WHERE noDaftar = '" & txtNoReg.Text & "'"
        ElseIf txtInst.Text.Contains("RAWAT JALAN") Then
            queryUnit = "SELECT kdUnit, unit, namapetugasMedis FROM vw_pasienrawatjalan WHERE noDaftar = '" & txtNoReg.Text & "'"
        End If

        cmdUnit = New MySqlCommand(queryUnit, conn)
        daUnit = New MySqlDataAdapter(cmdUnit)

        Dim str As New DataTable
        str.Clear()
        daUnit.Fill(str)

        txtUnitAsal.Text = ""

        If str.Rows.Count() > 0 Then
            txtKdUnit.Text = str.Rows(0)(0).ToString
            txtUnitAsal.Text = str.Rows(0)(1).ToString
            txtDokter.Text = str.Rows(0)(2).ToString
        Else
            MessageBox.Show("Pasien Tidak Ada / Belum Terdaftar pada Poli atau Ruang", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
        conn.Close()

        Call autoNoPermintaan()
    End Sub

    Sub cariKelas()
        Call koneksiServer()

        If txtInst.Text.Contains("IGD") Then
            txtKelas.Text = "KELAS I"
        ElseIf txtInst.Text.Contains("RAWAT INAP") Then
            Dim query As String = ""
            query = "SELECT kelas FROM vw_pasienrawatinap WHERE noDaftar = '" & txtNoReg.Text & "'"
            cmd = New MySqlCommand(query, conn)
            da = New MySqlDataAdapter(cmd)

            Dim str As New DataTable
            str.Clear()
            da.Fill(str)
            txtKelas.Text = ""

            If str.Rows.Count() > 0 Then
                txtKelas.Text = str.Rows(0)(0).ToString
            Else
                MessageBox.Show("Pasien Tidak Ada / Belum Terdaftar pada Poli atau Ruang", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        ElseIf txtInst.Text.Contains("RAWAT JALAN") Then
            txtKelas.Text = "KELAS III"
        End If

        conn.Close()
    End Sub

    Sub autoNoPermintaan()
        Dim noPermintaanPa As String

        Try
            Call koneksiServer()
            Dim query As String = ""
            Dim kode As String = ""

            If txtInst.Text.Contains("IGD") Then
                query = "SELECT SUBSTR(noRegistrasiPARajal,18,4) FROM t_registrasipatologirajal ORDER BY CAST(SUBSTR(noRegistrasiPARajal,18,4) AS UNSIGNED) DESC LIMIT 1"
                kode = "RJPA"
            ElseIf txtInst.Text.Contains("RAWAT INAP") Then
                query = "SELECT SUBSTR(noRegistrasiPARanap,18,4) FROM t_registrasipatologiranap ORDER BY CAST(SUBSTR(noRegistrasiPARanap,18,4) AS UNSIGNED) DESC LIMIT 1"
                kode = "RIPA"
            ElseIf txtInst.Text.Contains("RAWAT JALAN") Then
                query = "SELECT SUBSTR(noRegistrasiPARajal,18,4) FROM t_registrasipatologirajal ORDER BY CAST(SUBSTR(noRegistrasiPARajal,18,4) AS UNSIGNED) DESC LIMIT 1"
                kode = "RJPA"
            End If

            cmd = New MySqlCommand(query, conn)
            dr = cmd.ExecuteReader
            dr.Read()
            If dr.HasRows Then
                dr.Read()
                noPermintaanPa = kode + Format(Now, "ddMMyyHHmmss") + "-" + (Val(Trim(dr.Item(0).ToString)) + 1).ToString
                txtNoPermintaan.Text = noPermintaanPa
            Else
                noPermintaanPa = kode + Format(Now, "ddMMyyHHmmss") + "-1"
                txtNoPermintaan.Text = noPermintaanPa
            End If
            dr.Close()
            conn.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MessageBoxIcon.Exclamation, "NO.PERMINTAAN")
        End Try
    End Sub

    Sub autoComboPoliRuang()
        Call koneksiServer()
        Dim cmd As MySqlCommand
        Dim da As MySqlDataAdapter

        cmd = New MySqlCommand("SELECT a.*
                                FROM	(
				                            SELECT UPPER(unit) AS unit FROM t_unit WHERE kdInstalasi = 'ki1' ORDER BY unit ASC
			                            ) AS a
                                UNION ALL
                                SELECT b.*
                                FROM 	(
				                            SELECT UPPER(rawatInap) AS unit FROM t_rawatinap ORDER BY rawatInap ASC
			                            ) AS b", conn)
        da = New MySqlDataAdapter(cmd)
        Dim dt As New DataTable
        da.Fill(dt)

        txtUnitAsal.DataSource = dt
        txtUnitAsal.DisplayMember = "unit"
        txtUnitAsal.ValueMember = "unit"
        txtUnitAsal.AutoCompleteMode = AutoCompleteMode.SuggestAppend
        txtUnitAsal.AutoCompleteSource = AutoCompleteSource.ListItems
        conn.Close()
    End Sub

    Private Sub TambahPasien_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim dt As String
        dt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
        datePermintaan.Text = dt

        Call autoComboPoliRuang()
        Call autoDokter()
    End Sub

    Private Sub btnCari_Click(sender As Object, e As EventArgs) Handles btnCari.Click
        If txtNoRM.Text = "" Then
            MessageBox.Show("Masukkan No.RM Pasien", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Else
            Call cariNoDaftar()
            Call cariDataPasien()
        End If
    End Sub

    Private Sub txtNoRM_KeyDown(sender As Object, e As KeyEventArgs) Handles txtNoRM.KeyDown
        If e.KeyCode = Keys.Enter And txtNoRM.Text = "" Then
            MessageBox.Show("Masukkan No.RM Pasien", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information)
        ElseIf e.KeyCode = Keys.Enter Then
            Call cariNoDaftar()
            Call cariDataPasien()
        End If
    End Sub

    Private Sub txtDokter_TextChanged(sender As Object, e As EventArgs) Handles txtDokter.TextChanged
        Call koneksiServer()
        Try
            Dim cmdDok As MySqlCommand
            Dim drDok As MySqlDataReader
            Dim queryDok As String = ""

            queryDok = "SELECT kdPetugasMedis FROM t_tenagamedis2 WHERE namapetugasMedis = '" & txtDokter.Text & "'"
            cmdDok = New MySqlCommand(queryDok, conn)
            drDok = cmdDok.ExecuteReader

            While drDok.Read
                txtKdDokter.Text = UCase(drDok.GetString("kdPetugasMedis"))
            End While
            drDok.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MessageBoxIcon.Exclamation, "KODE DOKTER")
        End Try

        If txtDokter.Text <> "" Then
            txtDokter.BackColor = Color.White
        End If

        conn.Close()
    End Sub

    Private Sub btnTambah_Click_1(sender As Object, e As EventArgs) Handles btnTambah.Click
        If txtNoPA.Text = "" Then
            MsgBox("Mohon No. PA diisi terlebih dahulu !!", MsgBoxStyle.Exclamation)
            Me.ErrorProvider1.SetError(Me.txtNoPA, "No. PA harus diisi")
        Else
            Me.ErrorProvider1.SetError(Me.txtNoPA, "")
            Select Case txtInst.Text
                Case "IGD"
                    Call addRegPARajal()
                Case "RAWAT JALAN"
                    Call addRegPARajal()
                Case "RAWAT INAP"
                    Call addRegPARanap()
            End Select

            Form2.Ambil_Data = True
            Form2.Form_Ambil_Data = "Tindakan RS"
            Form2.Show()

            Me.Close()
            Form1.Hide()
        End If
    End Sub
End Class