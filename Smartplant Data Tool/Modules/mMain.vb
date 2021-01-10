Module mMain
    Public DB As New EAMS.DataBaseTools.SQLServerTools
    Public _EXCLUDEPARA As String = ""
    Public _SheetName As String = "Sheet1"
    'Public Cache As New DBCaches


    Public Sub LoadSettings()
        DB.DataBaseLocation = GetSetting("TR", "Smartplant", "DBLoc", "")
        DB.DataBaseName = "TRSmartplant"
    End Sub
 

    Public Sub DBConnect()
        LoadSettings()
        Try
            If DB.DataBaseName <> "" Then DB.Connect()
        Catch ex As Exception
            MsgBox("Database Connection Failed")
        End Try
    End Sub
End Module
