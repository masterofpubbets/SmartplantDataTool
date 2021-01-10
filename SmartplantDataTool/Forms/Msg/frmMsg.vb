Public Class frmMsg
    Public MsgType As e_mesType
    Public MsgResult As e_msgRes
    Public unloadbyUser As Boolean = False

    Public Enum e_mesType
        Info = 1
        Critical = 2
        Exlamenation = 3
        Delete = 4
        YesNo = 5
    End Enum

    Public Enum e_msgRes
        OK = 1
        Cancel = 2
        Yes = 3
        NO = 4
    End Enum

    Private Sub frmmsg_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If e.CloseReason = CloseReason.UserClosing Then
            If Not unloadbyUser Then e.Cancel = True
        End If
    End Sub

    Private Sub frmmsg_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown, gCancel.KeyDown, gNo.KeyDown, gOK.KeyDown, gYes.KeyDown
        If e.KeyCode = Keys.Enter Then
            Select Case MsgType
                Case e_mesType.Critical
                    GradientButton1_Click(sender, e)
                Case e_mesType.Delete
                    GradientButton3_Click(sender, e)
                Case e_mesType.Exlamenation
                    GradientButton1_Click(sender, e)
                Case e_mesType.Info
                    GradientButton1_Click(sender, e)
                Case e_mesType.YesNo
                    GradientButton3_Click(sender, e)
            End Select
        End If
    End Sub

    Private Sub frmmsg_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.WindowState = FormWindowState.Normal
        FrmMain.Refresh()
        unloadbyUser = False
        Select Case MsgType
            Case e_mesType.Critical
                Message.ForeColor = Color.Maroon
                picCritical.Visible = True
                picDel.Visible = False
                picEx.Visible = False
                picInfo.Visible = False
                picYes.Visible = False
                gOK.Visible = True
                gOK.Left = 415
                gYes.Visible = False
                gNo.Visible = False
                gCancel.Visible = False
            Case e_mesType.Delete
                Message.ForeColor = Color.Indigo
                picCritical.Visible = False
                picDel.Visible = True
                picEx.Visible = False
                picInfo.Visible = False
                picYes.Visible = False
                gOK.Visible = False
                gYes.Visible = True
                gNo.Visible = True
                gCancel.Visible = False
            Case e_mesType.Exlamenation
                Message.ForeColor = Color.DarkOrange
                Me.BackColor = Color.DimGray
                picCritical.Visible = False
                picDel.Visible = False
                picEx.Visible = True
                picYes.Visible = False
                picInfo.Visible = False
                gOK.Visible = True
                gOK.Left = 415
                gYes.Visible = False
                gNo.Visible = False
                gCancel.Visible = False
            Case e_mesType.Info
                Message.ForeColor = Color.Black
                picCritical.Visible = False
                picDel.Visible = False
                picEx.Visible = False
                picYes.Visible = False
                picInfo.Visible = True
                gOK.Visible = True
                gOK.Left = 415
                gYes.Visible = False
                gNo.Visible = False
                gCancel.Visible = False
            Case e_mesType.YesNo
                Message.ForeColor = Color.Black
                picCritical.Visible = False
                picDel.Visible = False
                picEx.Visible = False
                picInfo.Visible = False
                picYes.Visible = True
                gOK.Visible = False
                gYes.Visible = True
                gNo.Visible = True
                gCancel.Visible = False
        End Select
    End Sub

    Private Sub GradientButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gOK.Click
        MsgResult = e_msgRes.OK
        unloadbyUser = True
        Me.Close()
    End Sub

    Private Sub GradientButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gYes.Click
        MsgResult = e_msgRes.Yes
        unloadbyUser = True
        Me.Close()
    End Sub

    Private Sub GradientButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gNo.Click
        MsgResult = e_msgRes.NO
        unloadbyUser = True
        Me.Close()
    End Sub

    Private Sub GradientButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles gCancel.Click
        MsgResult = e_msgRes.Cancel
        unloadbyUser = True
        Me.Close()
    End Sub
End Class