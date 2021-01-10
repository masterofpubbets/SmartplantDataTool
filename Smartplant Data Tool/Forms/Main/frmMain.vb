Public Class frmMain
    Public SPRule As New SPDT

    Private Sub frmMain_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        lblUser.Caption = My.User.Name
        SPRule.DBReConnect()
        SPRule.initiate()
        'SPRule.Dec("E:\Work\Programming\Smartplant Data Tool\Refrence\Catego.txt")
    End Sub

    Private Sub lblStatus_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles lblStatus.ItemClick
        SPRule.f_DBConnection()
    End Sub

    Private Sub BarButtonItem1_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles BarButtonItem1.ItemClick
        frmAppend.ShowDialog(Me)
    End Sub

    Private Sub BarButtonItem3_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles BarButtonItem3.ItemClick
        Dim frm As New frmModelManage
        frm.ShowDialog(Me)
    End Sub

    Private Sub BarButtonItem2_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles BarButtonItem2.ItemClick
        Dim frm As New frmMsg() With {.MsgType = frmMsg.e_mesType.YesNo}
        frm.Message.Text = "Do You Want to Delete All Models"
        frm.ShowDialog(Me)
        If frm.MsgResult = frmMsg.e_msgRes.NO Then Exit Sub
        SPRule.ResetAll()
        frm.MsgType = frmMsg.e_mesType.Info
        frm.Message.Text = "All Models Have Been Deleted"
        frm.ShowDialog(Me)
    End Sub

    Private Sub BarButtonItem4_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles BarButtonItem4.ItemClick
        frmDataMenu.MdiParent = Me
        frmDataMenu.Show()
    End Sub

    Private Sub BarButtonItem5_ItemClick(sender As System.Object, e As DevExpress.XtraBars.ItemClickEventArgs) Handles BarButtonItem5.ItemClick
        DB.ShrinkLogFile()
        Dim frm As New frmMsg() With {.MsgType = frmMsg.e_mesType.Info}
        frm.Message.Text = "Database Has Been Compacted"
        frm.ShowDialog(Me)
    End Sub
End Class
