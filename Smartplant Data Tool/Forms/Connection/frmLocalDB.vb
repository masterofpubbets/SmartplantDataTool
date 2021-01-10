Public Class frmLocalDB

    Private Sub frmLocalDB_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        frmMain.SPRule.ClearGarbage()
    End Sub

    Private Sub SimpleButton1_Click(sender As System.Object, e As System.EventArgs) Handles SimpleButton1.Click
        Me.Close()
    End Sub

    Private Sub frmLocalDB_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        lblDBStatus.Visible = False
        txtServer.Text = GetSetting("TR", "Smartplant", "DBLoc", "")
    End Sub

    Private Sub SimpleButton2_Click(sender As System.Object, e As System.EventArgs) Handles SimpleButton2.Click
        Dim frm As New frmMsg() With {.MsgType = frmMsg.e_mesType.Exlamenation}
        If Trim(txtServer.Text) = "" Then
            frm.Message.Text = "Server Location Empty"
            frm.ShowDialog(Me)
            frm = Nothing
            Exit Sub
        End If
        DB.DataBaseLocation = txtServer.Text
        frm = Nothing
        Try
            SaveSetting("TR", "Smartplant", "DBLoc", txtServer.Text)
            frmMain.SPRule.DBReConnect()
            If DB.DBStatus = ConnectionState.Open Then
                lblDBStatus.Text = "Connected"
                lblDBStatus.ForeColor = Color.Green
                lblDBStatus.Visible = True
                SaveSetting("TR", "Smartplant", "DBLoc", txtServer.Text)
                Me.Close()
            Else
                DB.DataBaseName = "master"
                DB.Connect()
                If DB.DBStatus = ConnectionState.Open Then
                    If Not DB.DatabaseExists("TRSmartplant") Then
                        lblDBStatus.Text = "Database Not Found"
                        lblDBStatus.ForeColor = Color.Black
                        lblDBStatus.Visible = True
                    End If
                Else
                    lblDBStatus.Text = "Unknown Error"
                    lblDBStatus.ForeColor = Color.Red
                    lblDBStatus.Visible = True
                End If
            End If
            frm = Nothing
        Catch ex As Exception

        End Try

    End Sub
End Class