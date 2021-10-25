Imports MySql.Data.MySqlClient
Imports Microsoft.Reporting.WinForms
Public Class viewCetakHasil

    Public Ambil_Data As String
    Public Form_Ambil_Data As String

    Private Sub viewCetakHasilRanap_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim makros As New ReportParameter("makros", Form3.htmlMakro)
        Dim mikros As New ReportParameter("mikros", Form3.htmlMikro)
        Dim kesimpulan As New ReportParameter("kesimpulan", Form3.htmlConclu)
        Dim nama As New ReportParameter("nama", Form3.txtNamaPasien.Text)
        Dim tglLahir As New ReportParameter("ttl", Form3.txtTglLahir.Text)
        Dim alamat As New ReportParameter("alamat", Form3.txtAlamat.Text)
        Dim asalRs As New ReportParameter("asalRs", Form3.txtRs.Text)
        Dim pengirim As New ReportParameter("pengirim", Form3.txtUnitAsal.Text)
        Dim dokter As New ReportParameter("dokter", Form3.txtDokter.Text)
        Dim tglSediaan As New ReportParameter("tglSediaan", Form3.dateTerimaSediaan.Text)
        Dim tglHasil As New ReportParameter("tglHasil", Form3.dateHasil.Text)
        Dim noPA As New ReportParameter("noPA", Form3.txtNoPaBaru.Text)
        Dim lokalisasi As New ReportParameter("lokalisasi", Form3.txtLokalisasi.Text)
        Dim klinis As New ReportParameter("klinis", Form3.txtKetKlinis.Text)
        Dim dokPA As New ReportParameter("dokPA", Form3.txtDokterPA.Text)
        Dim bahan As New ReportParameter("bahan", Form3.txtBahan.Text)

        ReportViewer1.LocalReport.SetParameters(makros)
        ReportViewer1.LocalReport.SetParameters(mikros)
        ReportViewer1.LocalReport.SetParameters(kesimpulan)
        ReportViewer1.LocalReport.SetParameters(nama)
        ReportViewer1.LocalReport.SetParameters(tglLahir)
        ReportViewer1.LocalReport.SetParameters(alamat)
        ReportViewer1.LocalReport.SetParameters(asalRs)
        ReportViewer1.LocalReport.SetParameters(pengirim)
        ReportViewer1.LocalReport.SetParameters(dokter)
        ReportViewer1.LocalReport.SetParameters(tglSediaan)
        ReportViewer1.LocalReport.SetParameters(tglHasil)
        ReportViewer1.LocalReport.SetParameters(noPA)
        ReportViewer1.LocalReport.SetParameters(lokalisasi)
        ReportViewer1.LocalReport.SetParameters(klinis)
        ReportViewer1.LocalReport.SetParameters(dokPA)
        ReportViewer1.LocalReport.SetParameters(bahan)
        koneksiServer()

        Dim noRm, noReg As String
        noRm = Form3.txtNoRM.Text
        noReg = Form3.txtNoPermintaan.Text

        Dim dt As New DataTable
        da = New MySqlDataAdapter("SELECT * FROM vw_cetakhasilpatologi
                                            WHERE noRegistrasiPARajal = '" & noReg & "'", conn)
        ds = New DataSet
        da.Fill(dt)
        ReportViewer1.LocalReport.DataSources.Clear()
        Dim rpt As New ReportDataSource("HasilPA", dt)
        ReportViewer1.LocalReport.DataSources.Add(rpt)


        Me.ReportViewer1.SetDisplayMode(DisplayMode.PrintLayout)
    End Sub
End Class