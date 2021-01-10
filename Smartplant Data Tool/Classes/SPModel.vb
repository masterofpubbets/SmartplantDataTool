Public Class SPModel
    Public Event OverallProgress(ByVal inx As Integer)
    Public Event FileProgress(ByVal inx As Integer)
    Public Event FileCount(ByVal inx As Integer)
    Public Event ModelErr(ByVal e As String)
    Public Event WorkingON()
    Public Event WorkingOff()
    Public Event Canceled()
    Public Event FinishedImporting()
    Private DBAcc As New EAMS.DataBaseTools.AccessDBTools


#Region "Methods"
    Public Sub ModelInfo(ByRef Tree As TreeView, ByVal ModelName As String)
        Const sql As String = "SELECT [file_name] as Category,count(linkage_index) as ItemCount FROM [tblLinkage] where Model_ID=xxx group by [file_name]"
        Dim ModelID As Integer = DB.ExcutResult(String.Format("select model_id from tblmodels where model_name='{0}'", ModelName))
        Dim dt As New DataTable
        dt = DB.ReturnDataTable(Replace(sql, "xxx", ModelID))
        Tree.Nodes.Clear()
        For inx As Integer = 0 To dt.Rows.Count - 1
            Tree.Nodes.Add(dt.Rows(inx).Item("Category"), dt.Rows(inx).Item("Category"))
            Tree.Nodes(inx).Nodes.Add("Item Count: " & dt.Rows(inx).Item("ItemCount"))
            Application.DoEvents()
        Next
    End Sub

    Public Sub RollBackImporting()
        DB.ExcuteNoneResult("delete from tblmodels", 0)
    End Sub
    Public Sub RollBackImporting(ByVal ModelName As String)
        DB.ExcuteNoneResult(String.Format("exec sp_RollBackImporting '{0}'", ModelName), 0)
    End Sub
    Public Function IsModelExists(ByVal ModelName As String) As Boolean
        If DB.ExcutResult(String.Format("select Model_Name from tblmodels where Model_Name='{0}'", ModelName)) = "" Then
            Return False
        Else
            Return True
        End If
        Return True
    End Function
    Public Sub StartImporting(ByVal ModelName As String, ByVal ModelDirectory As String)
        Try
            Dim lst As New ListBox

            DB.ExcuteNoneResult(String.Format("insert into tblmodels (model_name) values ('{0}')", ModelName))
            Dim ModelID As Integer = DB.ExcutResult(String.Format("select model_id from tblmodels where model_name='{0}'", ModelName))
            DBAcc.DataBaseLocation = String.Format("{0}\{1}.mdb2", ModelDirectory, ModelName)
            DBAcc.Connect()
            If DBAcc.DBStatus <> ConnectionState.Open Then
                RaiseEvent ModelErr("Model Connection Data Error")
                Exit Sub
            End If

            '1st Files  ======================================
            Dim dt As New DataTable
            RaiseEvent WorkingON()
            Application.DoEvents()
            dt = DBAcc.ReturnDataTable(String.Format("select {0} as Model_id,[label_value_index],[label_value],[label_value_numeric] from label_values where label_value not like '%@a=%'", ModelID))
            Application.DoEvents()
            RaiseEvent OverallProgress(1)
            dt.TableName = "tblPropertiesValue"
            DB.BulkInsert(dt)
            Application.DoEvents()
            '================================================

            '2st Files  ======================================
            dt = New DataTable
            RaiseEvent WorkingON()
            Application.DoEvents()
            dt = DBAcc.ReturnDataTable(String.Format("select {0} as Model_id,[label_name_index],[label_name] from label_names", ModelID))
            Application.DoEvents()
            RaiseEvent OverallProgress(2)
            dt.TableName = "tblProprties"
            DB.BulkInsert(dt)
            Application.DoEvents()
            '================================================

            '3st Files  ======================================
            dt = New DataTable
            RaiseEvent WorkingON()
            Application.DoEvents()
            DB.Fill(lst, "select distinct label_name_index from tblProprties where model_id=" & ModelID)
            For inx As Integer = 0 To lst.Items.Count - 1
                dt = DBAcc.ReturnDataTable(String.Format("select {0} as Model_id,[linkage_index],[label_name_index],[label_value_index],[label_line_number],[extended_label] from labels where label_name_index={1}", ModelID, lst.Items(inx)))
                Application.DoEvents()
                RaiseEvent OverallProgress(3)
                dt.TableName = "tblTags"
                DB.BulkInsert(dt)
                Application.DoEvents()
            Next

            '================================================

            '4st Files  ======================================
            dt = New DataTable
            RaiseEvent WorkingON()
            Application.DoEvents()
            dt = DBAcc.ReturnDataTable(String.Format("select {0} as Model_id,[DMRSLinkage],[link_key_1],[link_key_2],[link_key_3],[link_key_4],[label_file_index],[linkage_index],[moniker],[file_name] from linkage", ModelID))
            Application.DoEvents()
            RaiseEvent OverallProgress(4)
            dt.TableName = "TblLinkage"
            DB.BulkInsert(dt)
            Application.DoEvents()
            '================================================



            RaiseEvent WorkingOff()
            RaiseEvent FinishedImporting()
        Catch ex As Exception
            RaiseEvent ModelErr(ex.Message)
        End Try
    End Sub
#End Region

End Class
