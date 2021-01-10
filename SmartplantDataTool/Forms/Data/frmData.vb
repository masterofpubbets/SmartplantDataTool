Public Class frmData
    Private DA As New SqlClient.SqlDataAdapter
    Private DT As New DataTable
    Private de As New EAMS.Coding.FileEncoding
    Public ModelID As Integer = 0

    Private Sub GetData()
        FileIO.FileSystem.CopyFile(Application.StartupPath & "\Libraries\Categories", Application.StartupPath & "\Categories.tmp", True)
        de.EncryptOrDecryptFile(Application.StartupPath & "\Categories.tmp", "sqlserver", EAMS.Coding.FileEncoding.CryptoAction.ActionDecrypt)
        Dim obj As New System.IO.StreamReader(Application.StartupPath & "\Catego")
        Dim Sql As String = Replace(obj.ReadToEnd, "ModelIDXXX", ModelID)
        obj.Close()
        FileIO.FileSystem.DeleteFile(Application.StartupPath & "\Catego")
        ' 
        DA = DB.ReturnDataAdapter(Sql)
        DT = New DataTable
        DA.SelectCommand.CommandTimeout = 0
        DA.Fill(DT)
        grd.DataSource = DT
    End Sub

    Private Sub BarButtonItem1_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles BarButtonItem1.ItemClick
        GetData()
    End Sub

    Private Sub frmData_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        GetData()
    End Sub

    Private Sub BarButtonItem2_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles BarButtonItem2.ItemClick
        sveFle.FileName = ""
        sveFle.Filter = "XLSX Files|*.xlsx"
        sveFle.ShowDialog()
        If sveFle.FileName = "" Then Exit Sub
        grd.ExportToXlsx(sveFle.FileName)
    End Sub
End Class