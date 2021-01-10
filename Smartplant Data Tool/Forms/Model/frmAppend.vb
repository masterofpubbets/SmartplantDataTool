Public Class frmAppend
    Private ModelDirectory As String = ""

    Private Sub SimpleButton2_Click(sender As System.Object, e As System.EventArgs) Handles SimpleButton2.Click
        Try
            opnFle.FileName = ""
            ModelDirectory = ""
            opnFle.ShowDialog()
            If opnFle.FileName = "" Then Exit Sub
            pp.Visible = True
            lblInfo.Text = ""
            ModelDirectory = FileIO.FileSystem.GetFileInfo(opnFle.FileName).DirectoryName
            lblInfo.Text = String.Format("Path: {0}{1}", opnFle.FileName, vbCrLf)
            lblInfo.Text &= String.Format("Name: {0}{1}", Replace(FileIO.FileSystem.GetFileInfo(opnFle.FileName).Name, ".vue", ""), vbCrLf)
            lblInfo.Text &= String.Format("CreationTime: {0}{1}", FileIO.FileSystem.GetFileInfo(opnFle.FileName).CreationTime, vbCrLf)
            lblInfo.Text &= String.Format("Last Write Time: {0}{1}", FileIO.FileSystem.GetFileInfo(opnFle.FileName).LastWriteTime, vbCrLf)
            lblInfo.Text &= String.Format("Size (byte): {0}{1}", Val(FileIO.FileSystem.GetFileInfo(opnFle.FileName).Length), vbCrLf)
            pp.Visible = False
        Catch ex As Exception
            pp.Visible = False
        End Try

    End Sub

    Private Sub SimpleButton1_Click(sender As System.Object, e As System.EventArgs) Handles SimpleButton1.Click
        Me.Close()
    End Sub

    Private Sub frmAppend_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        lblInfo.Text = "No Model"
        ModelDirectory = ""
    End Sub

    Private Sub SimpleButton3_Click(sender As System.Object, e As System.EventArgs) Handles SimpleButton3.Click
        Dim frm As New frmMsg() With {.MsgType = frmMsg.e_mesType.Exlamenation}
        If opnFle.FileName = "" Then
            frm.Message.Text = "No Model to Append"
            frm.ShowDialog(Me)
            frm = Nothing
            Exit Sub
        End If
        If Not FileIO.FileSystem.FileExists(String.Format("{0}\{1}", ModelDirectory, Replace(FileIO.FileSystem.GetFileInfo(opnFle.FileName).Name, ".vue", ".mdb2"))) Then
            frm.Message.Text = "Incomplete Model Data"
            frm.ShowDialog(Me)
            frm = Nothing
            Exit Sub
        End If
        If DB.ExcutResult(String.Format("select Model_Name from tblmodels where Model_Name='{0}'", Replace(FileIO.FileSystem.GetFileInfo(opnFle.FileName).Name, ".vue", ""))) <> "" Then
            frm.Message.Text = "This Model Already Exists"
            frm.ShowDialog(Me)
            frm = Nothing
            Exit Sub
        End If
        Dim frmIm As New frmImportModelData() With {.ModelDirectory = ModelDirectory, .ModelName = Replace(FileIO.FileSystem.GetFileInfo(opnFle.FileName).Name, ".vue", "")}
        frmIm.ShowDialog(Me)
        frmIm = Nothing
    End Sub
End Class