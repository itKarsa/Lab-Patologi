Imports MySql.Data.MySqlClient
Public Class Form2

    Public Ambil_Data As String
    Public Form_Ambil_Data As String
    Dim stKawin As String
    Dim bahan As String()
    Dim fiksasi As String()
    Dim unit As String

    Sub tampilData()
        Call koneksiServer()
        Dim cmd As MySqlCommand
        Dim dr As MySqlDataReader
        Dim query As String

        query = "SELECT kdTarif,SUBSTR(kelompokTindakan,17,20) AS kelompokTindakan, tindakan, tarif FROM vw_tindakanpa"
        cmd = New MySqlCommand(query, conn)
        dr = cmd.ExecuteReader
        DataGridView1.Rows.Clear()

        Do While dr.Read
            DataGridView1.Rows.Add(dr.Item("kdTarif"), dr.Item("kelompokTindakan"),
                                   dr.Item("tindakan"), dr.Item("tarif"))
        Loop
        dr.Close()
        conn.Close()
    End Sub

    Sub caridata()
        Call koneksiServer()

        Dim query As String
        query = "SELECT kdTarif, SUBSTR(kelompokTindakan,17,20) AS kelompokTindakan, tindakan, tarif FROM vw_tindakanpa WHERE tindakan Like '%" & txtCari.Text & "%'"
        cmd = New MySqlCommand(query, conn)
        dr = cmd.ExecuteReader
        DataGridView1.Rows.Clear()

        Do While dr.Read
            DataGridView1.Rows.Add(dr.Item("kdTarif"), dr.Item("kelompokTindakan"),
                                   dr.Item("tindakan"), dr.Item("tarif"))
        Loop
        dr.Close()
        conn.Close()
    End Sub

    Sub loaddata()
        Call koneksiServer()

        Dim query As String = ""
        Select Case unit
            Case "RJ"
                query = "SELECT statusPerkawinan,lokalisasi,diagnosaKlinis, 
                                stadiumT,stadiumN,stadiumM, 
                                bahan,fiksasi 
                           FROM t_registrasipatologirajal
                          WHERE noRegistrasiPARajal = '" & Form1.txtNoPermintaan.Text & "'"
            Case "RI"
                query = "SELECT statusPerkawinan,lokalisasi,diagnosaKlinis, 
                                stadiumT,stadiumN,stadiumM, 
                                bahan,fiksasi 
                           FROM t_registrasipatologiranap
                          WHERE noRegistrasiPARanap = '" & Form1.txtNoPermintaan.Text & "'"
        End Select

        Try
            cmd = New MySqlCommand(query, conn)
            dr = cmd.ExecuteReader
            dr.Read()

            If dr.HasRows Then
                stKawin = dr.Item("statusPerkawinan").ToString
                txtLokalisasi.Text = dr.Item("lokalisasi").ToString
                txtDiagnos.Text = dr.Item("diagnosaKlinis").ToString
                txtStadiumT.Text = dr.Item("stadiumT").ToString
                txtStadiumN.Text = dr.Item("stadiumN").ToString
                txtStadiumM.Text = dr.Item("stadiumM").ToString
                bahan = dr.Item("bahan").ToString.Replace(" ", "").Split(",")
                fiksasi = dr.Item("fiksasi").ToString.Split(",")
            End If

            dr.Close()
            conn.Close()
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical, "Error Load Data")
        End Try

    End Sub

    Sub loadNoTindakan()

        Dim query As String = ""
        Select Case unit
            Case "RJ"
                query = "SELECT noTindakanPARajal 
                           FROM t_tindakanpatologirajal
                          WHERE noRegistrasiPARajal = '" & Form1.txtNoPermintaan.Text & "'"
            Case "RI"
                query = "SELECT noTindakanPARanap
                           FROM t_tindakanpatologiranap
                          WHERE noRegistrasiPARanap = '" & Form1.txtNoPermintaan.Text & "'"
        End Select
        'MsgBox(query)
        Try
            Call koneksiServer()
            cmd = New MySqlCommand(query, conn)
            dr = cmd.ExecuteReader
            dr.Read()

            Select Case unit
                Case "RJ"
                    If dr.HasRows Then
                        txtNoTindakan.Text = dr.Item("noTindakanPARajal").ToString
                    End If
                Case "RI"
                    If dr.HasRows Then
                        txtNoTindakan.Text = dr.Item("noTindakanPARanap").ToString
                    End If
            End Select

            dr.Close()
            conn.Close()
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical, "Error Load No.Tindakan")
        End Try

    End Sub

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

    Sub autoNoTindakan()
        Dim noTindakanPa As String
        Dim query As String = ""
        Dim kode As String = "TPA"

        If unit.Contains("RI") Then
            query = "SELECT SUBSTR(noTindakanPARanap,17,5) FROM t_tindakanpatologiranap ORDER BY CAST(SUBSTR(noTindakanPARanap,17,5) AS UNSIGNED) DESC LIMIT 1"
        ElseIf unit.Contains("RJ") Then
            query = "SELECT SUBSTR(noTindakanPARajal,17,5) FROM t_tindakanpatologirajal ORDER BY CAST(SUBSTR(noTindakanPARajal,17,5) AS UNSIGNED) DESC LIMIT 1"
        End If

        Try
            Call koneksiServer()
            cmd = New MySqlCommand(query, conn)
            dr = cmd.ExecuteReader
            dr.Read()
            If dr.HasRows Then
                dr.Read()
                noTindakanPa = kode + Format(Now, "ddMMyyHHmmss") + "-" + (Val(Trim(dr.Item(0).ToString)) + 1).ToString
                txtNoTindakan.Text = noTindakanPa
            Else
                noTindakanPa = kode + Format(Now, "ddMMyyHHmmss") + "-1"
                txtNoTindakan.Text = noTindakanPa
            End If
            dr.Close()
            conn.Close()
        Catch ex As Exception
            MsgBox(ex.Message, MessageBoxIcon.Exclamation, "NO.TINDAKAN")
        End Try
    End Sub

    Sub transferSelected()
        Call koneksiServer()

        Dim dt As New DataTable
        Dim dr As New System.Windows.Forms.DataGridViewRow
        For Each dr In Me.DataGridView1.SelectedRows
            DataGridView2.Rows.Add(1)
            DataGridView2.Rows(DataGridView2.RowCount - 1).Cells(0).Value = txtNoPermintaan.Text
            DataGridView2.Rows(DataGridView2.RowCount - 1).Cells(1).Value = txtNoTindakan.Text
            DataGridView2.Rows(DataGridView2.RowCount - 1).Cells(2).Value = dr.Cells(0).Value.ToString
            DataGridView2.Rows(DataGridView2.RowCount - 1).Cells(3).Value = dr.Cells(2).Value.ToString
            DataGridView2.Rows(DataGridView2.RowCount - 1).Cells(4).Value = dr.Cells(3).Value.ToString
            DataGridView2.Rows(DataGridView2.RowCount - 1).Cells(5).Value = datePermintaan.Text
            DataGridView2.Rows(DataGridView2.RowCount - 1).Cells(6).Value = txtKdDokter.Text
            DataGridView2.Rows(DataGridView2.RowCount - 1).Cells(7).Value = txtDokter.Text
            DataGridView2.Rows(DataGridView2.RowCount - 1).Cells(8).Value = dr.Cells(3).Value.ToString
            DataGridView2.Rows(DataGridView2.RowCount - 1).Cells(9).Value = txtNoRM.Text
            DataGridView2.Rows(DataGridView2.RowCount - 1).Cells(10).Value = txtReg.Text
            DataGridView2.Rows(DataGridView2.RowCount - 1).Cells(11).Value = "PERMINTAAN"
            DataGridView2.Update()
        Next

        For i As Integer = 0 To DataGridView2.RowCount - 1
            If DataGridView2.Rows(i).Cells(11).Value.ToString = "PERMINTAAN" Then
                DataGridView2.Rows(i).Cells(11).Style.BackColor = Color.Orange
                DataGridView2.Rows(i).Cells(11).Style.ForeColor = Color.White
            End If
        Next

        conn.Close()
    End Sub

    Sub totalTarif()
        Dim totTarif As Long = 0

        For i As Integer = 0 To DataGridView2.Rows.Count - 1
            totTarif = totTarif + Val(DataGridView2.Rows(i).Cells(4).Value)
        Next
        txtTotalTarif.Text = totTarif
    End Sub

    Private Sub PopulateCheckListBoxesBahan()
        Call koneksiServer()
        cmd = New MySqlCommand("SELECT bahanPA FROM t_bahanpa", conn)
        da = New MySqlDataAdapter(cmd)

        Using sda As MySqlDataAdapter = New MySqlDataAdapter(cmd)
            Dim dt As DataTable = New DataTable()
            sda.Fill(dt)
            CheckedListBox1.DisplayMember = "bahanPA"
            CheckedListBox1.ValueMember = "bahanPA"
            If dt.Rows.Count > 0 Then
                For i As Integer = 0 To dt.Rows.Count - 1
                    CheckedListBox1.Items.Add(CStr(dt.Rows(i).Item(0)), False)
                Next
            End If
            'AddHandler chk.CheckedChanged, AddressOf CheckBox_Checked
        End Using
    End Sub

    Private Sub PopulateCheckListBoxesFiksasi()
        Call koneksiServer()
        cmd = New MySqlCommand("SELECT fiksasi FROM t_fiksasipa", conn)
        da = New MySqlDataAdapter(cmd)

        Using sda As MySqlDataAdapter = New MySqlDataAdapter(cmd)
            Dim dt As DataTable = New DataTable()
            sda.Fill(dt)
            CheckedListBox2.DisplayMember = "fiksasi"
            CheckedListBox2.ValueMember = "fiksasi"
            If dt.Rows.Count > 0 Then
                For i As Integer = 0 To dt.Rows.Count - 1
                    CheckedListBox2.Items.Add(CStr(dt.Rows(i).Item(0)), False)
                Next
            End If
            'AddHandler chk.CheckedChanged, AddressOf CheckBox_Checked
        End Using
    End Sub

    Private Sub PopulateCheckListBoxesKanker()
        Dim str As String()
        str = {"Klinik", "RO", "Parth.Klinik", "Operasi", "Nekropsi"}

        For Each row As String In str
            CheckedListBox3.DisplayMember = "str"
            CheckedListBox3.ValueMember = "str"
            CheckedListBox3.Items.Add(row)
        Next
    End Sub

    Private Sub PopulateRadioButton()
        Dim str As String()
        str = {"Kawin", "Belum", "Duda", "Janda"}

        For Each row As String In str
            Dim rd As RadioButton = New RadioButton()
            rd.Width = 65
            rd.Margin = New Padding(0, 0, 0, 0)
            rd.Text = row.ToString()
            rd.ForeColor = Color.Black

            If rd.Text = stKawin Then
                rd.Checked = True
            End If

            AddHandler rd.CheckedChanged, AddressOf RadioButton_Checked
            FlowLayoutPanel4.Controls.Add(rd)
        Next
    End Sub

    Private Sub RadioButton_Checked(ByVal sender As Object, ByVal e As EventArgs)
        Dim rd As RadioButton = (TryCast(sender, RadioButton))
        If rd.Checked Then
            stKawin = rd.Text
            'MessageBox.Show("You selected: " & rd.Text)
        End If
    End Sub

    Sub addRegistrasiPa()
        Dim dt As String
        dt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

        Call koneksiServer()
        Dim str As String = ""
        Select Case unit
            Case "RJ"
                str = "UPDATE t_registrasipatologirajal
                          SET kelas = '" & txtKelas.Text & "',statusPerkawinan = '" & stKawin & "',
                              tglMasukPARajal = '" & dt & "',lokalisasi = '" & txtLokalisasi.Text & "',
                              diagnosaKlinis = '" & txtDiagnos.Text & "',stadiumT = '" & txtStadiumT.Text & "',
                              stadiumN = '" & txtStadiumN.Text & "',stadiumM = '" & txtStadiumM.Text & "',
                              bahan = @bahan,fiksasi = @fiksasi, 
                              riwayatKlinisSkrg = '" & txtSekarang.Text & "',diagnosaKanker = @dgKanker
                       WHERE  noRegistrasiPARajal = '" & txtNoPermintaan.Text & "'"
            Case "RI"
                str = "UPDATE t_registrasipatologiranap 
                          SET kelas = '" & txtKelas.Text & "',statusPerkawinan = '" & stKawin & "',
                              tglMasukPARanap = '" & dt & "',lokalisasi = '" & txtLokalisasi.Text & "',
                              diagnosaKlinis = '" & txtDiagnos.Text & "',stadiumT = '" & txtStadiumT.Text & "',
                              stadiumN = '" & txtStadiumN.Text & "',stadiumM = '" & txtStadiumM.Text & "',
                              bahan = @bahan,fiksasi = @fiksasi, 
                              riwayatKlinisSkrg = '" & txtSekarang.Text & "',diagnosaKanker = @dgKanker
                       WHERE  noRegistrasiPARanap = '" & txtNoPermintaan.Text & "'"
        End Select

        Try
            Using cmd As New MySqlCommand(str, conn)
                cmd.Parameters.AddWithValue("@bahan", pilihBahan)
                cmd.Parameters.AddWithValue("@fiksasi", pilihFiksasi)
                cmd.Parameters.AddWithValue("@dgKanker", pilihRiwayat)
                cmd.ExecuteNonQuery()
                cmd.Parameters.Clear()
            End Using

            MsgBox("Permintaan PA berhasil diupdate", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox("Insert data Permintaan PA gagal diupdate." & ex.Message, MsgBoxStyle.Critical, "Error Update")
            cmd.Dispose()
        End Try

        conn.Close()
    End Sub

    Sub addTindakan()

        Dim cmd As MySqlCommand
        Dim str As String = ""

        If Ambil_Data = True Then
            Select Case Form_Ambil_Data
                Case "Tindakan RS"
                    Select Case unit
                        Case "RJ"
                            str = "INSERT INTO t_tindakanpatologirajal(noTindakanPARajal,noRegistrasiPARajal,totalTindakanPA,statusPembayaran) 
                                   VALUES ('" & txtNoTindakan.Text & "','" & txtNoPermintaan.Text & "','" & txtTotalTarif.Text & "','BELUM LUNAS')"
                        Case "RI"
                            str = "INSERT INTO t_tindakanpatologiranap(noTindakanPARanap,noRegistrasiPARanap,totalTindakanPA,statusPembayaran) 
                                   VALUES ('" & txtNoTindakan.Text & "','" & txtNoPermintaan.Text & "','" & txtTotalTarif.Text & "','BELUM LUNAS')"
                    End Select
                Case "Tindakan"
                    Select Case unit
                        Case "RJ"
                            str = "UPDATE t_tindakanpatologirajal 
                                      SET totalTindakanPA = '" & txtTotalTarif.Text & "'
                                    WHERE noTindakanPARajal = '" & txtNoTindakan.Text & "'"
                        Case "RI"
                            str = "UPDATE t_tindakanpatologiranap
                                      SET totalTindakanPA = '" & txtTotalTarif.Text & "'
                                    WHERE noTindakanPARanap = '" & txtNoTindakan.Text & "'"
                    End Select
            End Select
        End If
        'MsgBox(str)
        Try
            Call koneksiServer()
            cmd = New MySqlCommand(str, conn)
            cmd.ExecuteNonQuery()
            'MsgBox("Insert Tindakan Lab berhasil dilakukan", MsgBoxStyle.Information, "Information")
        Catch ex As Exception
            MsgBox("Insert Tindakan PA gagal dilakukan." & ex.Message, MsgBoxStyle.Critical, "Error Tindakan")
        End Try

        conn.Close()
    End Sub

    Sub addDetail()

        Dim val1, val2, val3, val4, val8 As String
        Dim str As String = ""
        Select Case unit
            Case "RJ"
                str = "INSERT INTO t_detailtindakanpatologirajal
                                   (noTindakanPARajal,kdTarif,tindakan,tarif,jumlahTindakan,totalTarif,statusTindakan) 
                            VALUES (@noTindakanPA,@kdTarif,@tindakan,@tarif,'1',@totalTarif,'PERMINTAAN')"
            Case "RI"
                str = "INSERT INTO t_detailtindakanpatologiranap
                                   (noTindakanPARanap,kdTarif,tindakan,tarif,jumlahTindakan,totalTarif,statusTindakan) 
                            VALUES (@noTindakanPA,@kdTarif,@tindakan,@tarif,'1',@totalTarif,'PERMINTAAN')"
        End Select

        Try
            Call koneksiServer()
            cmd = New MySqlCommand(str, conn)
            For i As Integer = 0 To DataGridView2.Rows.Count - 1
                val1 = DataGridView2.Rows(i).Cells(1).Value
                val2 = DataGridView2.Rows(i).Cells(2).Value
                val3 = DataGridView2.Rows(i).Cells(3).Value
                val4 = DataGridView2.Rows(i).Cells(4).Value
                val8 = DataGridView2.Rows(i).Cells(4).Value

                cmd.Parameters.AddWithValue("@noTindakanPA", val1)
                cmd.Parameters.AddWithValue("@kdTarif", val2)
                cmd.Parameters.AddWithValue("@tindakan", val3)
                cmd.Parameters.AddWithValue("@tarif", val4)
                cmd.Parameters.AddWithValue("@totalTarif", val8)
                cmd.ExecuteNonQuery()
                cmd.Parameters.Clear()
            Next
            'MsgBox("Insert data Detail Tindakan Lab berhasil dilakukan", MessageBoxIcon.Information, "Information")
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Information, "Error Detail 1")
            cmd.Dispose()
        End Try

        conn.Close()

    End Sub

    Function pilihBahan() As String
        Dim bahan As New List(Of String)
        Dim noBahan As String

        For Each item As String In ListBox4.Items
            bahan.Add(item)
        Next

        noBahan = String.Join(", ", bahan.ToArray)

        Return noBahan
    End Function

    Function pilihFiksasi() As String
        Dim fiksasi As New List(Of String)
        Dim noFiksasi As String

        For Each item As String In ListBox5.Items
            fiksasi.Add(item)
        Next

        noFiksasi = String.Join(",", fiksasi.ToArray)

        Return noFiksasi
    End Function

    Function pilihRiwayat() As String
        Dim riwayat As New List(Of String)
        Dim noRiwayat As String

        For Each item As String In ListBox3.Items
            riwayat.Add(item)
        Next

        noRiwayat = String.Join(", ", riwayat.ToArray)

        Return noRiwayat
    End Function

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.WindowState = FormWindowState.Normal
        Me.StartPosition = FormStartPosition.Manual
        With Screen.PrimaryScreen.WorkingArea
            Me.SetBounds(.Left, .Top, .Width, .Height)
        End With

        pnlStats.Height = btnTindakan.Height
        pnlStats.Top = btnTindakan.Top
        btnTindakan.BackColor = Color.DodgerBlue

        tampilData()
        autoDokter()

        If Ambil_Data = True Then
            Select Case Form_Ambil_Data
                Case "Tindakan RS"
                    txtNoRM.Text = TambahPasien.txtNoRM.Text
                    txtReg.Text = TambahPasien.txtNoReg.Text
                    txtPasien.Text = TambahPasien.txtNamaPasien.Text
                    txtAlamat.Text = TambahPasien.txtAlamat.Text
                    txtUmur.Text = hitungUmur(Convert.ToDateTime(TambahPasien.txtTglLahir.Text))
                    txtJk.Text = TambahPasien.jk
                    txtRuang.Text = TambahPasien.txtUnitAsal.Text
                    txtKdRuang.Text = TambahPasien.txtKdUnit.Text
                    txtKdDokter.Text = TambahPasien.txtKdDokter.Text
                    txtDokter.Text = TambahPasien.txtDokter.Text
                    txtRS.Text = "RSU KARSA HUSADA BATU"
                    datePermintaan.Text = TambahPasien.datePermintaan.Text
                    txtNoPermintaan.Text = TambahPasien.txtNoPermintaan.Text
                    txtKelas.Text = TambahPasien.txtKelas.Text
                    Call autoNoTindakan()
                Case "Tindakan"
                    txtNoRM.Text = Form1.txtNoRM.Text
                    txtReg.Text = Form1.txtNoReg.Text
                    txtPasien.Text = Form1.txtNamaPasien.Text
                    txtAlamat.Text = Form1.txtAlamat.Text
                    txtUmur.Text = hitungUmur(Convert.ToDateTime(Form1.txtTglLahir.Text))
                    txtJk.Text = Form1.jk
                    txtRuang.Text = Form1.txtUnitAsal.Text
                    txtKdRuang.Text = Form1.txtKdUnit.Text
                    txtKdDokter.Text = Form1.txtKdDokter.Text
                    txtDokter.Text = Form1.txtDokter.Text
                    txtRS.Text = "RSU KARSA HUSADA BATU"
                    'datePermintaan.Text = Form1.datePermintaan.Text
                    txtNoPermintaan.Text = Form1.txtNoPermintaan.Text
                    'txtKelas.Text = Form1.txtKelas.Text
                    Call loaddata()
                    Call loadNoTindakan()
            End Select
        End If

        PopulateRadioButton()
        PopulateCheckListBoxesBahan()
        PopulateCheckListBoxesFiksasi()
        PopulateCheckListBoxesKanker()

        If bahan IsNot Nothing AndAlso bahan.Count > 0 Then
            'MsgBox("Ada")
            For i As Integer = 0 To CheckedListBox1.Items.Count - 1
                Dim Item As String = CType(CheckedListBox1.Items(i), String)
                For Each bhn As String In bahan
                    If Item.Equals(bhn) Then
                        CheckedListBox1.SetItemChecked(i, True)
                    End If
                Next
            Next

            For j As Integer = 0 To CheckedListBox2.Items.Count - 1
                Dim fik As String = CType(CheckedListBox2.Items(j), String)
                For Each fks As String In fiksasi
                    If fik.Equals(fks) Then
                        CheckedListBox2.SetItemChecked(j, True)
                    End If
                Next
            Next
        Else
            'MsgBox("Kosong")
            Return
        End If
    End Sub

    Private Sub CheckedListBox1_ItemCheck(sender As Object, e As ItemCheckEventArgs) Handles CheckedListBox1.ItemCheck
        'Dim item As String = CheckedListBox1.SelectedItem
        Dim item As String = CType(CheckedListBox1.Items(e.Index), String)
        Dim kdItem As String = ""

        Try
            Call koneksiServer()
            Dim query As String
            query = "SELECT * FROM t_bahanpa WHERE bahanPA = '" & item & "'"
            cmd = New MySqlCommand(query, conn)
            dr = cmd.ExecuteReader

            While dr.Read
                kdItem = UCase(dr.GetString("kdBahanPA"))
            End While

            If e.NewValue = CheckState.Checked Then
                ListBox1.Items.Add(kdItem)
                ListBox4.Items.Add(item)
            Else
                ListBox1.Items.Remove(kdItem)
                ListBox4.Items.Remove(item)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
        conn.Close()

    End Sub

    Private Sub CheckedListBox2_ItemCheck(sender As Object, e As ItemCheckEventArgs) Handles CheckedListBox2.ItemCheck
        'Dim item As String = CheckedListBox2.SelectedItem
        Dim item As String = CType(CheckedListBox2.Items(e.Index), String)
        Dim kdItem As String = ""

        Try
            Call koneksiServer()
            Dim query As String
            query = "SELECT * FROM t_fiksasipa WHERE fiksasi = '" & item & "'"
            cmd = New MySqlCommand(query, conn)
            dr = cmd.ExecuteReader

            While dr.Read
                kdItem = UCase(dr.GetString("kdFiksasi"))
            End While

            If e.NewValue = CheckState.Checked Then
                ListBox2.Items.Add(kdItem)
                ListBox5.Items.Add(item)
            Else
                ListBox2.Items.Remove(kdItem)
                ListBox5.Items.Remove(item)
            End If

        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
        conn.Close()

    End Sub

    Private Sub CheckedListBox3_ItemCheck(sender As Object, e As ItemCheckEventArgs) Handles CheckedListBox3.ItemCheck
        'Dim item As String = CheckedListBox3.SelectedItem
        Dim item As String = CType(CheckedListBox3.Items(e.Index), String)

        If e.NewValue = CheckState.Checked Then
            ListBox3.Items.Add(item)
        Else
            ListBox3.Items.Remove(item)
        End If
    End Sub


    Private Sub txtCari_KeyDown(sender As Object, e As KeyEventArgs) Handles txtCari.KeyDown
        If e.KeyCode = Keys.Enter Then
            Call caridata()
        End If
    End Sub

    Private Sub DataGridView1_KeyDown(sender As Object, e As KeyEventArgs) Handles DataGridView1.KeyDown
        If e.KeyCode = Keys.Enter And DataGridView1.CurrentCell.RowIndex >= 0 Then
            e.Handled = True
            e.SuppressKeyPress = True

            Dim row As DataGridViewRow
            row = Me.DataGridView1.Rows(DataGridView1.CurrentCell.RowIndex)

            If DataGridView1.CurrentCell.RowIndex = -1 Then
                Return
            End If

            Call transferSelected()
            Call totalTarif()
        End If
    End Sub

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting

        If e.RowIndex > 0 And e.ColumnIndex = 1 Then
            If DataGridView1.Item(1, e.RowIndex - 1).Value = e.Value Then
                e.Value = ""
            End If
        End If

        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If i Mod 2 = 0 Then
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.AliceBlue
            Else
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.White
            End If
        Next

    End Sub

    Private Sub DataGridView2_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView2.CellFormatting

        For i As Integer = 0 To DataGridView2.Rows.Count - 1
            If i Mod 2 = 0 Then
                DataGridView2.Rows(i).DefaultCellStyle.BackColor = Color.AliceBlue
            Else
                DataGridView2.Rows(i).DefaultCellStyle.BackColor = Color.White
            End If
        Next
    End Sub

    Private Sub DataGridView2_RowsAdded(sender As Object, e As DataGridViewRowsAddedEventArgs) Handles DataGridView2.RowsAdded
        Call totalTarif()
    End Sub

    Private Sub DataGridView2_RowsRemoved(sender As Object, e As DataGridViewRowsRemovedEventArgs) Handles DataGridView2.RowsRemoved
        'If DataGridView2.Rows.Count = 0 Then
        '    Me.btnSimpan.Enabled = False
        'Else
        '    Me.btnSimpan.Enabled = True
        'End If

        Call totalTarif()
    End Sub

    Private Sub DataGridView1_RowPostPaint(sender As Object, e As DataGridViewRowPostPaintEventArgs) Handles DataGridView1.RowPostPaint
        Dim dg As DataGridView = DirectCast(sender, DataGridView)
        ' Current row record
        Dim rowNumber As String = (e.RowIndex + 1).ToString()

        ' Format row based on number of records displayed by using leading zeros
        'While rowNumber.Length < dg.RowCount.ToString().Length
        '    rowNumber = "0" & rowNumber
        'End While

        ' Position text
        Dim size As SizeF = e.Graphics.MeasureString(rowNumber, Me.Font)
        If dg.RowHeadersWidth < CInt(size.Width + 20) Then
            dg.RowHeadersWidth = CInt(size.Width + 20)
        End If

        ' Use default system text brush
        Dim b As Brush = SystemBrushes.ControlText

        ' Draw row number
        e.Graphics.DrawString(rowNumber, dg.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))
    End Sub

    Private Sub DataGridView2_RowPostPaint(sender As Object, e As DataGridViewRowPostPaintEventArgs) Handles DataGridView2.RowPostPaint
        Dim dg As DataGridView = DirectCast(sender, DataGridView)
        ' Current row record
        Dim rowNumber As String = (e.RowIndex + 1).ToString()

        ' Format row based on number of records displayed by using leading zeros
        'While rowNumber.Length < dg.RowCount.ToString().Length
        '    rowNumber = "0" & rowNumber
        'End While

        ' Position text
        Dim size As SizeF = e.Graphics.MeasureString(rowNumber, Me.Font)
        If dg.RowHeadersWidth < CInt(size.Width + 20) Then
            dg.RowHeadersWidth = CInt(size.Width + 20)
        End If

        ' Use default system text brush
        Dim b As Brush = SystemBrushes.ControlText

        ' Draw row number
        e.Graphics.DrawString(rowNumber, dg.Font, b, e.RowBounds.Location.X + 15, e.RowBounds.Location.Y + ((e.RowBounds.Height - size.Height) / 2))
    End Sub

    Private Sub btnPilihOk_Click(sender As Object, e As EventArgs) Handles btnPilihOk.Click
        Call transferSelected()
        Call totalTarif()
    End Sub

    Private Sub btnPilihCancel_Click(sender As Object, e As EventArgs) Handles btnPilihCancel.Click
        Dim drDgv As New DataGridViewRow
        For Each drDgv In Me.DataGridView2.SelectedRows
            DataGridView2.Rows.Remove(drDgv)
        Next
        Call totalTarif()
    End Sub

    Private Sub txtNoPermintaan_TextChanged(sender As Object, e As EventArgs) Handles txtNoPermintaan.TextChanged
        unit = txtNoPermintaan.Text.Substring(0, 2)
    End Sub

    Private Sub btnDash_Click(sender As Object, e As EventArgs) Handles btnDash.Click
        Dim konfirmasi As MsgBoxResult

        konfirmasi = MsgBox("Apakah anda ingin menyimpan permintaan ?", vbQuestion + vbYesNo, "EXIT")
        If konfirmasi = vbYes Then
            If DataGridView2.Rows.Count = 0 Then
                MsgBox("Tindakan belum diinputkan !!!", MsgBoxStyle.Exclamation, "Warning")
            Else
                If Ambil_Data = True Then
                    Select Case Form_Ambil_Data
                        Case "Tindakan RS"
                            Call addRegistrasiPa()
                            Call addTindakan()
                            Call addDetail()
                            DataGridView2.Rows.Clear()
                            Me.Close()
                        Case "Tindakan"
                            Call addTindakan()
                            Call addDetail()
                            DataGridView2.Rows.Clear()
                            Me.Close()
                    End Select
                End If

                Form1.pnlStats.Height = Form1.btnDash.Height
                Form1.pnlStats.Top = Form1.btnDash.Top
                Form1.btnDash.BackColor = Color.DodgerBlue
                Form1.Show()
                Me.Close()
                Call Form1.tampilDataAll()
            End If
        Else
            Form1.pnlStats.Height = Form1.btnDash.Height
            Form1.pnlStats.Top = Form1.btnDash.Top
            Form1.btnDash.BackColor = Color.DodgerBlue
            Form1.Show()
            Me.Close()
        End If
    End Sub

    Private Sub btnKeluar_Click(sender As Object, e As EventArgs) Handles btnKeluar.Click
        Dim konfirmasi As MsgBoxResult

        konfirmasi = MsgBox("Apakah anda ingin menyimpan permintaan ?", vbQuestion + vbYesNo, "EXIT")
        If konfirmasi = vbYes Then
            If DataGridView2.Rows.Count = 0 Then
                MsgBox("Tindakan belum diinputkan !!!", MsgBoxStyle.Exclamation, "Warning")
            Else
                If Ambil_Data = True Then
                    Select Case Form_Ambil_Data
                        Case "Tindakan RS"
                            Call addRegistrasiPa()
                            Call addTindakan()
                            Call addDetail()
                            DataGridView2.Rows.Clear()
                            Me.Close()
                        Case "Tindakan"
                            Call addTindakan()
                            Call addDetail()
                            DataGridView2.Rows.Clear()
                            Me.Close()
                    End Select
                End If

                Form1.pnlStats.Height = Form1.btnDash.Height
                Form1.pnlStats.Top = Form1.btnDash.Top
                Form1.btnDash.BackColor = Color.DodgerBlue
                Form1.Show()
                Me.Close()
                Call Form1.tampilDataAll()
            End If
        Else
            Form1.pnlStats.Height = Form1.btnDash.Height
            Form1.pnlStats.Top = Form1.btnDash.Top
            Form1.btnDash.BackColor = Color.DodgerBlue
            Form1.Show()
            Me.Close()
        End If
    End Sub

    Private Sub rdYa_CheckedChanged(sender As Object, e As EventArgs) Handles rdYa.CheckedChanged
        CheckedListBox3.Enabled = True
        ListBox3.Items.Clear()
    End Sub

    Private Sub rdTidak_CheckedChanged(sender As Object, e As EventArgs) Handles rdTidak.CheckedChanged
        CheckedListBox3.Enabled = False
        ListBox3.Items.Add("-")
        For Each i As Integer In CheckedListBox3.CheckedIndices
            CheckedListBox3.SetItemCheckState(i, CheckState.Unchecked)
        Next
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim konfirmasi As MsgBoxResult

        konfirmasi = MsgBox("Apakah anda ingin menyimpan permintaan ?", vbQuestion + vbYesNo, "Konfirmasi")
        If konfirmasi = vbYes Then

            If Ambil_Data = True Then
                Select Case Form_Ambil_Data
                    Case "Tindakan RS"
                        Call addRegistrasiPa()
                        Call addTindakan()
                        Call addDetail()
                        DataGridView2.Rows.Clear()
                        Me.Close()
                    Case "Tindakan"
                        Call addTindakan()
                        Call addDetail()
                        DataGridView2.Rows.Clear()
                        Me.Close()
                End Select
            End If

            Form1.pnlStats.Height = Form1.btnDash.Height
            Form1.pnlStats.Top = Form1.btnDash.Top
            Form1.btnDash.BackColor = Color.DodgerBlue
            Form1.Show()
            Me.Close()
            Call Form1.tampilDataAll()
        End If
    End Sub
End Class