Public Class frmDataMenu

    Private Sub GetModel()
        DB.Fill(cmbModel, "select model_name from tblmodels order by model_name")
    End Sub
    Private Sub GetCategories(ByVal ModelName As String)
        Dim ModelID As Integer = DB.ExcutResult(String.Format("select model_id from tblmodels where model_name='{0}'", ModelName))
        DB.Fill(lstCat, "select distinct file_name from tblLinkage where model_id=" & ModelID)
    End Sub
    Private Sub SimpleButton1_Click(sender As System.Object, e As System.EventArgs) Handles SimpleButton1.Click
        Me.Close()
    End Sub

    Private Sub frmDataMenu_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        GetModel()
    End Sub

    Private Sub cmbModel_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles cmbModel.SelectedIndexChanged
        If cmbModel.SelectedIndex = -1 Then
            lstCat.Items.Clear()
        Else
            GetCategories(cmbModel.SelectedItem)
        End If
    End Sub

    Private Sub SimpleButton3_Click(sender As System.Object, e As System.EventArgs) Handles SimpleButton3.Click
        Dim frm As New frmMsg() With {.MsgType = frmMsg.e_mesType.Exlamenation}
        If lstCat.SelectedIndex = -1 Then
            frm.Message.Text = "You Have to Select Category"
            frm.ShowDialog(Me)
            Exit Sub
        End If
        DB.ExcuteNoneResult("delete from tblSelCat")
        For inx As Integer = 0 To lstCat.SelectedIndices.Count - 1
            DB.ExcuteNoneResult(String.Format("insert into tblSelCat (Selected_Cat) values ('{0}')", lstCat.SelectedItems.Item(inx)))
        Next

        Dim frmV As New frmData() With {.ModelID = DB.ExcutResult(String.Format("select model_id from tblmodels where model_name='{0}'", cmbModel.SelectedItem)), .MdiParent = frmMain}
        frmV.RibbonPage1.Text = String.Format("{0} - {1}", cmbModel.SelectedItem, lstCat.SelectedItem)
        frmV.Show()
        Me.Close()
    End Sub

    Private Sub txtFind_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtFind.TextChanged
        If Len(txtFind.Text) > 0 Then
            lstCat.SelectedIndex = lstCat.FindString(txtFind.Text)
        Else
            lstCat.SelectedIndex = -1
        End If
    End Sub
End Class