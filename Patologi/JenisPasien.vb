Public Class JenisPasien
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Dispose()
        TambahPasien.ShowDialog()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        MsgBox("Maaf masih dalam pengembangan ..", MsgBoxStyle.Information)
        'Me.Dispose()
        'TambahPasienLuar.ShowDialog()
    End Sub
End Class