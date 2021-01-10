Public Class SPDT

#Region "inialization"
    Public Sub New()
        AddHandler DB.Connnected, AddressOf OnConnected
        AddHandler DB.Disconnected, AddressOf OnDisconnect
    End Sub
    Private Sub MainSpContainersVisble(ByVal vis As Boolean)

    End Sub
    Private Sub ButtonsStatus(ByVal st As Boolean)

    End Sub
    Private Sub CheckDBStatus()

    End Sub
    Private Sub OnConnected()
        CheckDBStatus()
        frmMain.lblStatus.Caption = "Status: Online"
        frmMain.lblStatus.Appearance.ForeColor = Color.Green
        frmMain.RibbonPage1.Visible = True
        frmMain.RibbonPage2.Visible = True
    End Sub
    Private Sub OnDisconnect()
        ButtonsStatus(False)
        frmMain.lblStatus.Caption = "Status: Offline"
        frmMain.lblStatus.Appearance.ForeColor = Color.Maroon
        frmMain.RibbonPage1.Visible = False
        frmMain.RibbonPage2.Visible = False
    End Sub
#End Region



#Region "Methods"
    Public Sub ResetAll()
        Dim model As New SPModel
        model.RollBackImporting()
    End Sub
    Public Sub ClearGarbage()
        CheckDBStatus()
        If Application.OpenForms.Count = 1 Then
            'Me.Home()
        End If
    End Sub

    Public Sub DBReConnect()
        Try
            DBConnect()
        Catch ex As Exception
        End Try
    End Sub
    Public Sub initiate()
        If DB.DBStatus = ConnectionState.Closed Then
            OnDisconnect()
        End If
    End Sub
#End Region

#Region "Queries"
    Public Sub Dec(ByVal FilePath As String)
        Dim d As New EAMS.Coding.FileEncoding
        d.EncryptOrDecryptFile(FilePath, "sqlserver", EAMS.Coding.FileEncoding.CryptoAction.ActionEncrypt)
    End Sub
#End Region
#Region "Forms Operation"
    Public Sub f_DBConnection()
        MainSpContainersVisble(False)
        frmLocalDB.MdiParent = frmMain
        frmLocalDB.Show()
        frmLocalDB.WindowState = FormWindowState.Normal
        frmMain.Refresh()
    End Sub
  
#End Region
End Class
