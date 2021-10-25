Imports MySql.Data.MySqlClient
Public Class FormVerifikasi

    Public Ambil_Data As String
    Public Form_Ambil_Data As String

    Sub autoDokter()
        Call koneksiServer()

        Using cmd As New MySqlCommand("SELECT namapetugasMedis FROM t_tenagamedis2 WHERE namapetugasMedis LIKE '%Sp. PA%'", conn)
            Using rd As MySqlDataReader = cmd.ExecuteReader
                While rd.Read
                    With txtDokPA
                        .AutoCompleteMode = AutoCompleteMode.Suggest
                        .AutoCompleteCustomSource.Add(rd.Item(0))
                        .AutoCompleteSource = AutoCompleteSource.CustomSource
                    End With
                End While
                rd.Close()
            End Using
        End Using
        conn.Close()
    End Sub

    Sub transferSelected()
        Dim row As New System.Windows.Forms.DataGridViewRow

        For Each row In Form1.dgv2.SelectedRows
            DataGridView1.Rows.Add(1)
            DataGridView1.Rows(DataGridView1.RowCount - 1).Cells(0).Value = row.Cells(1).Value.ToString
            DataGridView1.Update()
        Next
    End Sub

    Private Sub FormVerifikasi_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        DataGridView1.Rows.Clear()
        txtPALama.Text = ""
        txtPABaru.Text = ""

        If Ambil_Data = True Then
            Select Case Form_Ambil_Data
                Case "Verifikasi"
                    txtNoRM.Text = Form1.txtNoRM.Text
                    txtNamaPasien.Text = Form1.txtNamaPasien.Text
                    txtUsia.Text = Form1.txtUsia.Text
                    txtUnitAsal.Text = Form1.txtUnitAsal.Text
                    txtDokter.Text = Form1.txtDokter.Text
                    txtDokPA.Text = Form1.txtDokPA.Text
                    txtPALama.Text = Form1.txtNoPALama.Text
                    txtPABaru.Text = Form1.txtNoPABaru.Text
                    Call transferSelected()
            End Select
        End If

        txtDokPA.BackColor = Color.FromArgb(255, 112, 112)
        txtPALama.BackColor = Color.FromArgb(255, 112, 112)
        txtPABaru.BackColor = Color.FromArgb(255, 112, 112)
        DataGridView1.ClearSelection()

        Call autoDokter()
    End Sub

    Private Sub txtDokPA_KeyDown(sender As Object, e As KeyEventArgs) Handles txtDokPA.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
            If txtDokPA.Text = "" Then
                txtDokPA.BackColor = Color.FromArgb(255, 112, 112)
            End If
        End If
    End Sub

    Private Sub txtPALama_KeyDown(sender As Object, e As KeyEventArgs) Handles txtPALama.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
            If txtPALama.Text = "" Then
                txtPALama.BackColor = Color.FromArgb(255, 112, 112)
            End If
        End If
    End Sub

    Private Sub txtPABaru_KeyDown(sender As Object, e As KeyEventArgs) Handles txtPABaru.KeyDown
        If e.KeyCode = Keys.Enter Then
            SendKeys.Send("{TAB}")
            If txtPABaru.Text = "" Then
                txtPABaru.BackColor = Color.FromArgb(255, 112, 112)
            End If
        End If
    End Sub

    Private Sub txtDokPA_TextChanged(sender As Object, e As EventArgs) Handles txtDokPA.TextChanged
        If txtDokPA.Text <> "" Then
            txtDokPA.BackColor = Color.White
        End If

        Call koneksiServer()
        Try
            Dim query As String
            query = "SELECT * FROM t_tenagamedis2 WHERE namapetugasMedis = '" & txtDokPA.Text & "'"
            cmd = New MySqlCommand(query, conn)
            dr = cmd.ExecuteReader

            While dr.Read
                Form1.txtKdDokterPA.Text = UCase(dr.GetString("kdPetugasMedis"))
            End While
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        conn.Close()
    End Sub

    Private Sub txtPALama_TextChanged(sender As Object, e As EventArgs) Handles txtPALama.TextChanged
        If txtPALama.Text <> "" Then
            txtPALama.BackColor = Color.White
        End If
    End Sub

    Private Sub txtPABaru_TextChanged(sender As Object, e As EventArgs) Handles txtPABaru.TextChanged
        If txtPABaru.Text <> "" Then
            txtPABaru.BackColor = Color.White
        End If
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

    Private Sub DataGridView1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting
        For i As Integer = 0 To DataGridView1.Rows.Count - 1
            If i Mod 2 = 0 Then
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.AliceBlue
            Else
                DataGridView1.Rows(i).DefaultCellStyle.BackColor = Color.White
            End If
        Next
    End Sub

    Private Sub btnSimpan_Click(sender As Object, e As EventArgs) Handles btnSimpan.Click
        Call Form1.ClickMulai()
        Call Form1.tampilDataAll()
    End Sub
End Class