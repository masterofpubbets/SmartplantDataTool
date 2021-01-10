Public Class frmModelManage
    Private model As New SPModel

    Private Sub GetData()
        DB.Fill(lstModel, "select model_name from tblmodels order by model_name")
    End Sub
    Private Sub SimpleButton1_Click(sender As System.Object, e As System.EventArgs) Handles SimpleButton1.Click
        Me.Close()
    End Sub

    Private Sub SimpleButton2_Click(sender As System.Object, e As System.EventArgs) Handles SimpleButton2.Click
        Dim frm As New frmMsg() With {.MsgType = frmMsg.e_mesType.Exlamenation}
        If lstModel.SelectedIndex = -1 Then
            frm.Message.Text = "You Have to Select Model"
            frm.ShowDialog(Me)
            frm = Nothing
            Exit Sub
        End If
        frm.MsgType = frmMsg.e_mesType.YesNo
        frm.Message.Text = "Do You Want to Delete Selected Model"
        frm.ShowDialog(Me)
        If frm.MsgResult = frmMsg.e_msgRes.NO Then Exit Sub
        pp.Visible = True
        Application.DoEvents()
        model.RollBackImporting(lstModel.SelectedItem.ToString)
        GetData()
        pp.Visible = False
        frm = Nothing
    End Sub

    Private Sub frmModelManage_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        GetData()
    End Sub
End Class