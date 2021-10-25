Imports System.ComponentModel
Imports MySql.Data.MySqlClient
Imports System.Text
Imports Patologi.Itenso.Rtf
Imports Patologi.Itenso.Rtf.Support
Imports Patologi.Itenso.Rtf.Converter.Html

Public Class IsiHasil

    Dim kdHasil As String = ""

    'Dim mikros As Byte() = System.Text.Encoding.ASCII.GetBytes(RichTextBox2.Rtf)
    'Dim kesimpulan As Byte() = System.Text.Encoding.ASCII.GetBytes(RichTextBox3.Rtf)

    Sub tampilKdHasil()
        Try
            Call koneksiServer()
            Dim str As String
            Dim cmd As MySqlCommand
            Dim dr As MySqlDataReader
            str = "SELECT kdHasilPA FROM t_hasilpemeriksaanpatologi WHERE noRegistrasiPA = '" & Form3.txtNoPermintaan.Text & "'"
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

    Sub updateMakros()
        Call koneksiServer()
        Try
            Dim str As String
            Dim cmd As MySqlCommand
            Dim makros As Byte() = System.Text.Encoding.Unicode.GetBytes(RichTextBox1.Rtf)
            str = "UPDATE t_hasilpemeriksaanpatologi 
                      SET makroskopik = @makro,
                          userModify = CONCAT(userModify,';" & LoginForm.txtUsername.Text & "'),
                          dateModify = CONCAT(dateModify,';" & Format(DateTime.Now, "dd/MM/yyyy HH:mm:ss") & "')  
                    WHERE kdHasilPA = '" & kdHasil & "'"
            'MsgBox(str)
            cmd = New MySqlCommand(str, conn)
            cmd.Parameters.Add("makro", MySqlDbType.Binary).Value = makros
            cmd.ExecuteNonQuery()
            MsgBox("Hasil makroskopis berhasil diupdate.", MsgBoxStyle.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message & " -Makroskopis-")
        End Try
        conn.Close()
    End Sub

    Sub updateMikros()
        Call koneksiServer()
        Try
            Dim str As String
            Dim cmd As MySqlCommand
            Dim mikros As Byte() = System.Text.Encoding.Unicode.GetBytes(RichTextBox2.Rtf)
            str = "UPDATE t_hasilpemeriksaanpatologi 
                      SET mikroskopik = @mikro,
                          userModify = CONCAT(userModify,';" & LoginForm.txtUsername.Text & "'),
                          dateModify = CONCAT(dateModify,';" & Format(DateTime.Now, "dd/MM/yyyy HH:mm:ss") & "')  
                    WHERE kdHasilPA = '" & kdHasil & "'"
            'MsgBox(str)
            cmd = New MySqlCommand(str, conn)
            cmd.Parameters.Add("mikro", MySqlDbType.Binary).Value = mikros
            cmd.ExecuteNonQuery()
            MsgBox("Hasil mikroskopis berhasil diupdate.", MsgBoxStyle.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message & " -Mikroskopis-")
        End Try
        conn.Close()
    End Sub

    Sub updateKesimpulan()
        Call koneksiServer()
        Try
            Dim str As String
            Dim cmd As MySqlCommand
            Dim conclu As Byte() = System.Text.Encoding.Unicode.GetBytes(RichTextBox3.Rtf)
            str = "UPDATE t_hasilpemeriksaanpatologi 
                      SET kesimpulan = @conclu,
                          userModify = CONCAT(userModify,';" & LoginForm.txtUsername.Text & "'),
                          dateModify = CONCAT(dateModify,';" & Format(DateTime.Now, "dd/MM/yyyy HH:mm:ss") & "')  
                    WHERE kdHasilPA = '" & kdHasil & "'"
            'MsgBox(str)
            cmd = New MySqlCommand(str, conn)
            cmd.Parameters.Add("conclu", MySqlDbType.Binary).Value = conclu
            cmd.ExecuteNonQuery()
            MsgBox("Hasil kesimpulan berhasil diupdate.", MsgBoxStyle.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message & " -Kesimpulan-")
        End Try
        conn.Close()
    End Sub

    Sub tampilHasil()
        Try
            Call koneksiServer()
            Dim str As String
            Dim cmd As MySqlCommand
            Dim dr As MySqlDataReader
            str = "SELECT makroskopik, mikroskopik, kesimpulan FROM t_hasilpemeriksaanpatologi WHERE noRegistrasiPA = '" & Form3.txtNoPermintaan.Text & "'"
            cmd = New MySqlCommand(str, conn)
            dr = cmd.ExecuteReader
            dr.Read()
            If dr.HasRows Then
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

    Private Sub IsiHasil_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TabControl1.SelectedTab = TabPage1
        Call tampilKdHasil()
        Call tampilHasil()

        'MsgBox(Form3.txtNoPermintaan.Text & " | " & kdHasil)
    End Sub

    Private Sub BulletsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BulletsToolStripMenuItem.Click
        If TabControl1.SelectedIndex = 2 Then   'Makros
            If RichTextBox1.SelectionBullet = True Then
                RichTextBox1.SelectionBullet = False
            ElseIf RichTextBox1.SelectionBullet = False Then
                RichTextBox1.SelectionBullet = True
            End If
        ElseIf TabControl1.SelectedIndex = 1 Then   'Mikros
            If RichTextBox2.SelectionBullet = True Then
                RichTextBox2.SelectionBullet = False
            ElseIf RichTextBox2.SelectionBullet = False Then
                RichTextBox2.SelectionBullet = True
            End If
        ElseIf TabControl1.SelectedIndex = 0 Then   'Kesimpulan
            If RichTextBox3.SelectionBullet = True Then
                RichTextBox3.SelectionBullet = False
            ElseIf RichTextBox3.SelectionBullet = False Then
                RichTextBox3.SelectionBullet = True
            End If
        End If
    End Sub
    Private Sub BoldsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BoldsToolStripMenuItem.Click
        With Me.RichTextBox1
            If .SelectionFont IsNot Nothing Then
                Dim currentFont As System.Drawing.Font = .SelectionFont
                Dim newFontStyle As System.Drawing.FontStyle

                If .SelectionFont.Bold = True Then
                    newFontStyle = currentFont.Style - Drawing.FontStyle.Bold
                Else
                    newFontStyle = currentFont.Style + Drawing.FontStyle.Bold
                End If

                .SelectionFont = New Drawing.Font(currentFont.FontFamily, currentFont.Size, newFontStyle)
            End If
        End With

        With Me.RichTextBox2
            If .SelectionFont IsNot Nothing Then
                Dim currentFont As System.Drawing.Font = .SelectionFont
                Dim newFontStyle As System.Drawing.FontStyle

                If .SelectionFont.Bold = True Then
                    newFontStyle = currentFont.Style - Drawing.FontStyle.Bold
                Else
                    newFontStyle = currentFont.Style + Drawing.FontStyle.Bold
                End If

                .SelectionFont = New Drawing.Font(currentFont.FontFamily, currentFont.Size, newFontStyle)
            End If
        End With

        With Me.RichTextBox3
            If .SelectionFont IsNot Nothing Then
                Dim currentFont As System.Drawing.Font = .SelectionFont
                Dim newFontStyle As System.Drawing.FontStyle

                If .SelectionFont.Bold = True Then
                    newFontStyle = currentFont.Style - Drawing.FontStyle.Bold
                Else
                    newFontStyle = currentFont.Style + Drawing.FontStyle.Bold
                End If

                .SelectionFont = New Drawing.Font(currentFont.FontFamily, currentFont.Size, newFontStyle)
            End If
        End With
    End Sub

    Private Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click
        If TabControl1.SelectedIndex = 2 Then   'Makros
            'MsgBox("Save Makroskopik")
            Call updateMakros()
        ElseIf TabControl1.SelectedIndex = 1 Then   'Mikros
            'MsgBox("Save Mikroskopik")
            Call updateMikros()
        ElseIf TabControl1.SelectedIndex = 0 Then   'Kesimpulan
            'MsgBox("Save Kesimpulan")
            Call updateKesimpulan()
        End If
    End Sub

    Private Sub SaveAllToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveAllToolStripMenuItem.Click
        Call koneksiServer()
        Try
            Dim str As String
            Dim cmd As MySqlCommand
            Dim makros As Byte() = System.Text.Encoding.Unicode.GetBytes(RichTextBox1.Rtf)
            Dim mikros As Byte() = System.Text.Encoding.Unicode.GetBytes(RichTextBox2.Rtf)
            Dim conclu As Byte() = System.Text.Encoding.Unicode.GetBytes(RichTextBox3.Rtf)
            str = "UPDATE t_hasilpemeriksaanpatologi 
                      SET makroskopik = @makros,
                          mikroskopik = @mikros,
                          kesimpulan = @conclu,
                          userModify = CONCAT(userModify,';" & LoginForm.txtUsername.Text & "'),
                          dateModify = CONCAT(dateModify,';" & Format(DateTime.Now, "dd/MM/yyyy HH:mm:ss") & "')  
                    WHERE kdHasilPA = '" & kdHasil & "'"
            'MsgBox(str)
            cmd = New MySqlCommand(str, conn)
            cmd.Parameters.Add("makros", MySqlDbType.Binary).Value = makros
            cmd.Parameters.Add("mikros", MySqlDbType.Binary).Value = mikros
            cmd.Parameters.Add("conclu", MySqlDbType.Binary).Value = conclu
            cmd.ExecuteNonQuery()
            MsgBox("Hasil berhasil diupdate.", MsgBoxStyle.Information)
        Catch ex As Exception
            MessageBox.Show(ex.Message & " -Save All-")
        End Try
        conn.Close()
    End Sub
    Private Sub TabControl1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        If TabControl1.SelectedIndex = 2 Then   'Makros
            RichTextBox1.Select()
        ElseIf TabControl1.SelectedIndex = 1 Then   'Mikros
            RichTextBox2.Select()
        ElseIf TabControl1.SelectedIndex = 0 Then   'Kesimpulan
            RichTextBox3.Select()
        End If
    End Sub

    Private Sub CopyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyToolStripMenuItem.Click
        If TabControl1.SelectedIndex = 2 Then   'Makros
            RichTextBox1.Copy()
        ElseIf TabControl1.SelectedIndex = 1 Then   'Mikros
            RichTextBox2.Copy()
        ElseIf TabControl1.SelectedIndex = 0 Then   'Kesimpulan
            RichTextBox3.Copy()
        End If
    End Sub

    Private Sub PasteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PasteToolStripMenuItem.Click
        If TabControl1.SelectedIndex = 2 Then   'Makros
            RichTextBox1.Paste()
        ElseIf TabControl1.SelectedIndex = 1 Then   'Mikros
            RichTextBox2.Paste()
        ElseIf TabControl1.SelectedIndex = 0 Then   'Kesimpulan
            RichTextBox3.Paste()
        End If
    End Sub

    Private Sub SeelctAllToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SeelctAllToolStripMenuItem.Click
        If TabControl1.SelectedIndex = 2 Then   'Makros
            RichTextBox1.SelectAll()
        ElseIf TabControl1.SelectedIndex = 1 Then   'Mikros
            RichTextBox2.SelectAll()
        ElseIf TabControl1.SelectedIndex = 0 Then   'Kesimpulan
            RichTextBox3.SelectAll()
        End If
    End Sub

    Private Sub UndoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UndoToolStripMenuItem.Click
        If TabControl1.SelectedIndex = 2 Then   'Makros
            RichTextBox1.Undo()
        ElseIf TabControl1.SelectedIndex = 1 Then   'Mikros
            RichTextBox2.Undo()
        ElseIf TabControl1.SelectedIndex = 0 Then   'Kesimpulan
            RichTextBox3.Undo()
        End If
    End Sub

    Private Sub RedoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RedoToolStripMenuItem.Click
        If TabControl1.SelectedIndex = 2 Then   'Makros
            RichTextBox1.Redo()
        ElseIf TabControl1.SelectedIndex = 1 Then   'Mikros
            RichTextBox2.Redo()
        ElseIf TabControl1.SelectedIndex = 0 Then   'Kesimpulan
            RichTextBox3.Redo()
        End If
    End Sub

    Private Sub CutStripMenuItem1_Click(sender As Object, e As EventArgs) Handles CutStripMenuItem1.Click
        If TabControl1.SelectedIndex = 2 Then   'Makros
            RichTextBox1.Cut()
        ElseIf TabControl1.SelectedIndex = 1 Then   'Mikros
            RichTextBox2.Cut()
        ElseIf TabControl1.SelectedIndex = 0 Then   'Kesimpulan
            RichTextBox3.Cut()
        End If
    End Sub

    Private Sub IsiHasil_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Form3.tampilHasil()
    End Sub

    Private ReadOnly Property ConversionText As String
        Get

            If Me.RichTextBox1.SelectionLength > 0 Then
                Return Me.RichTextBox1.SelectedRtf
            End If

            Return Me.RichTextBox1.Rtf
        End Get
    End Property ' ConversionText

    'Private Sub ToHtmlButtonClick(ByVal sender As Object, ByVal e As EventArgs)
    '    Try
    '        Dim rtfDocument As IRtfDocument = RtfInterpreterTool.BuildDoc(ConversionText)
    '        Dim htmlConverter As RtfHtmlConverter = New RtfHtmlConverter(rtfDocument)
    '        RichTextBox1.Text = htmlConverter.Convert()
    '    Catch exception As Exception
    '        MessageBox.Show(Me, "Error " & exception.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    End Try
    'End Sub ' ToHtmlButtonClick
End Class