Imports MySql.Data.MySqlClient
Imports Microsoft.Reporting.WinForms
Public Class viewCetakRincianBiaya

    Public Ambil_Data As String
    Public Form_Ambil_Data As String

    Private Sub viewCetakRincianBiaya_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim noRMParam As New ReportParameter("noRM", Form1.txtNoRM.Text)
        Dim noPAParam As New ReportParameter("noPA", Form1.txtNoPABaru.Text)
        Dim namaParam As New ReportParameter("nama", Form1.txtNamaPasien.Text)
        Dim ttlParam As New ReportParameter("ttl", Form1.txtTglLahir.Text)
        Dim jkParam As New ReportParameter("jk", Form1.txtJK.Text)
        Dim alamatParam As New ReportParameter("alamat", Form1.txtAlamat.Text)
        Dim dokterParam As New ReportParameter("dokterPengirim", If(Form1.txtDokter.Text.Equals(Nothing), "-", Form1.txtDokter.Text))
        Dim bayarParam As New ReportParameter("pembayaran", Form1.caraBayar)
        Dim ruangParam As New ReportParameter("ruang", Form1.txtUnitAsal.Text)
        Dim userParam As New ReportParameter("user", LoginForm.user)

        ReportViewer1.LocalReport.SetParameters(noRMParam)
        ReportViewer1.LocalReport.SetParameters(noPAParam)
        ReportViewer1.LocalReport.SetParameters(namaParam)
        ReportViewer1.LocalReport.SetParameters(ttlParam)
        ReportViewer1.LocalReport.SetParameters(jkParam)
        ReportViewer1.LocalReport.SetParameters(alamatParam)
        ReportViewer1.LocalReport.SetParameters(dokterParam)
        ReportViewer1.LocalReport.SetParameters(bayarParam)
        ReportViewer1.LocalReport.SetParameters(ruangParam)
        ReportViewer1.LocalReport.SetParameters(userParam)

        Call koneksiServer()

        If Ambil_Data = True Then
            Select Case Form_Ambil_Data
                Case "Cetak"
                    Dim noTindakan As String
                    noTindakan = Form1.noTindakanPA

                    'MsgBox()

                    Dim dt As New DataTable
                    da = New MySqlDataAdapter("SELECT * FROM vw_pasienpatologidetail
                                                       WHERE noTindakanPARajal = '" & noTindakan & "'", conn)
                    ds = New DataSet
                    da.Fill(dt)
                    ReportViewer1.LocalReport.DataSources.Clear()
                    Dim rpt As New ReportDataSource("cetakRincian", dt)
                    ReportViewer1.LocalReport.DataSources.Add(rpt)
            End Select
        End If

        Me.ReportViewer1.SetDisplayMode(DisplayMode.PrintLayout)
        Me.ReportViewer1.RefreshReport()
    End Sub
End Class