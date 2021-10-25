Imports MySql.Data.MySqlClient
Public Class DaftarPasienPA

    Sub tampilDataAll()
        Call koneksiServer()
        Dim query As String = ""
        Dim cmd As MySqlCommand
        Dim dr As MySqlDataReader

        query = "SELECT * FROM vw_pasienpatologi
                  WHERE tglMasukPARajal BETWEEN '" & Format(DateTimePicker1.Value, "yyyy-MM-dd") & "' 
                    AND '" & Format(DateAdd(DateInterval.Day, 1, DateTimePicker1.Value), "yyyy-MM-dd") & "'
               ORDER BY tglMasukPARajal DESC"

        cmd = New MySqlCommand(query, conn)
        dr = cmd.ExecuteReader
        dgv1.Rows.Clear()

        Do While dr.Read
            dgv1.Rows.Add(dr.Item("noRekamedis"), dr.Item("noDaftar"), dr.Item("nmPasien"),
                          dr.Item("unitAsal"), dr.Item("alamat"), dr.Item("tglMasukPARajal"),
                          dr.Item("statusPA"))
        Loop
        conn.Close()

    End Sub

    Private Sub DaftarPasienPA_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Call tampilDataAll()
    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        Call tampilDataAll()
    End Sub

    Private Sub dgv1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgv1.CellContentClick
        If e.ColumnIndex = 7 Then
            Form3.Ambil_Data = True
            Form3.Form_Ambil_Data = "DaftarHasil"
            Form3.Show()
        End If
    End Sub

    Private Sub dgv1_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles dgv1.CellFormatting
        dgv1.Columns(3).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv1.Columns(5).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
        dgv1.Columns(6).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

        For i = 0 To dgv1.RowCount - 1
            If i Mod 2 = 0 Then
                dgv1.Rows(i).DefaultCellStyle.BackColor = Color.White
            Else
                dgv1.Rows(i).DefaultCellStyle.BackColor = Color.AliceBlue
            End If
        Next

        For i As Integer = 0 To dgv1.Rows.Count - 1
            If dgv1.Rows(i).Cells(6).Value = "PERMINTAAN" Then
                dgv1.Rows(i).Cells(6).Style.BackColor = Color.Orange
                dgv1.Rows(i).Cells(6).Style.ForeColor = Color.White
            ElseIf dgv1.Rows(i).Cells(6).Value = "DALAM TINDAKAN" Then
                dgv1.Rows(i).Cells(6).Style.BackColor = Color.Green
                dgv1.Rows(i).Cells(6).Style.ForeColor = Color.White
            ElseIf dgv1.Rows(i).Cells(6).Value = "SELESAI" Then
                dgv1.Rows(i).Cells(6).Style.BackColor = Color.Red
                dgv1.Rows(i).Cells(6).Style.ForeColor = Color.White
            End If
        Next

        For i = 0 To dgv1.Rows.Count - 1
            dgv1.Rows(i).Cells(7).Style.BackColor = Color.DodgerBlue
            dgv1.Rows(i).Cells(7).Style.ForeColor = Color.White
        Next
    End Sub

    Private Sub dgv1_RowPostPaint(sender As Object, e As DataGridViewRowPostPaintEventArgs) Handles dgv1.RowPostPaint
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

    Private Sub btnCari_Click(sender As Object, e As EventArgs) Handles btnCari.Click

    End Sub
End Class