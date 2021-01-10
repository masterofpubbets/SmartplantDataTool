Public Class frmImportModelData
    Private Model As New SPModel
    Public ModelName As String = ""
    Public ModelDirectory As String = ""

    Private Sub h_Finished()
        SimpleButton1.Enabled = True
        SimpleButton2.Visible = True
        Dim frm As New frmMsg() With {.MsgType = frmMsg.e_mesType.Info}
        frm.Message.Text = "Importing Model Data Has Been Finished"
        frm.ShowDialog(Me)
        frm = Nothing
        Dim frmMI As New frmModelInfo
        Model.ModelInfo(frmMI.tree, ModelName)
        frmMI.lblModelName.Text = "Model: " & ModelName
        frmMI.ShowDialog(Me)
        frmMI = Nothing
        Me.Close()
    End Sub
   
    Private Sub h_WorkingOn()
        pp.Visible = True
    End Sub
    Private Sub h_WorkingOff()
        pp.Visible = False
    End Sub
    Private Sub h_err(ByVal m As String)
        Dim frm As New frmMsg() With {.MsgType = frmMsg.e_mesType.Critical}
        frm.Message.Text = m
        frm.ShowDialog(Me)
        frm = Nothing
        SimpleButton1.Enabled = True
        SimpleButton2.Visible = True
        pp.Visible = False
        pbAll.Position = 0
    End Sub

    Private Sub h_AllProgress(ByVal inx As Integer)
        pbAll.Position = inx
        lblAllProgress.Text = String.Format("Overall Progress: {0} / {1}", inx, "4")
    End Sub

    Private Sub SimpleButton3_Click(sender As System.Object, e As System.EventArgs)

    End Sub

    Private Sub SimpleButton1_Click(sender As System.Object, e As System.EventArgs) Handles SimpleButton1.Click
        Me.Close()
    End Sub

    Private Sub frmImportModelData_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        pbAll.Position = 0
        lblModelName.Text = "Model: " & ModelName
        AddHandler Model.ModelErr, AddressOf h_err
        AddHandler Model.OverallProgress, AddressOf h_AllProgress
        AddHandler Model.WorkingOff, AddressOf h_WorkingOff
        AddHandler Model.WorkingON, AddressOf h_WorkingOn
        AddHandler Model.FinishedImporting, AddressOf h_Finished
    End Sub

    Private Sub SimpleButton2_Click(sender As System.Object, e As System.EventArgs) Handles SimpleButton2.Click 'Start
        Try
            pbAll.Position = 0
            SimpleButton1.Enabled = False
            SimpleButton2.Visible = False
            Model.StartImporting(ModelName, ModelDirectory)
        Catch ex As Exception
            SimpleButton1.Enabled = True
            SimpleButton2.Visible = True
        End Try

    End Sub

End Class