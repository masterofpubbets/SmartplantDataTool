Imports Microsoft.Office.Interop
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Drawing
Imports System.Drawing.Drawing2D
Imports system.Drawing.Imaging
Imports System.IO
Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports System.Text
Imports System
Imports System.Collections.Generic
Imports System.Threading
Imports System.Data
Imports System.Security.Cryptography
Imports System.Security
Imports System.Net, VB = Microsoft.VisualBasic
Imports System.Net.Mail
Imports System.IO.Compression





Namespace EAMS

    Namespace DataBaseTools

        Public Class SQLServerTools
            Private DB As New System.Data.SqlClient.SqlConnection
            Public DataBaseName As String = ""
            Public DataBaseLocation As String = "."
            Public UserName As String = ""
            Public Pass As String = ""
            Private SQLServerCmd As New System.Data.SqlClient.SqlCommand
            Public Event err(ByVal errName As String)
            Public Event Connnected()
            Public Event ConnnectedTOMaster()
            Public Event Disconnected()
            Public Event RestoreDatabaseComplete()
            Public Event ConnectionTerminated()
            Public Event DeattachedComplete()
            Public Event ExportProgress(ByVal inx As Integer)
            Public Event ExportingComplete()
            Public Event ProgressCount(ByVal c As Integer)


#Region "Structure"
            Public Structure st_ExcelRange
                Dim ColName As String
                Dim RangeColCount As Integer
                Dim BackColor As Color
                Dim ForeColor As Color
                Dim LineWeight As Integer
                Dim FontSize As Integer
                Dim IsBold As Boolean
                Dim ColumnWidth As Integer
                Dim BorderColor As Color
                Dim HAlign As Excel.XlHAlign
                Dim VAlign As Excel.XlHAlign
            End Structure
#End Region


#Region "Data Methods" '
            Public Sub BulkInsert(ByRef dt As DataTable)
                Dim bulkCopy As SqlBulkCopy = New SqlBulkCopy(DB)
                bulkCopy.DestinationTableName = dt.TableName
                bulkCopy.WriteToServer(dt)
            End Sub
            Public Function ReturnFieldDataType(ByVal sql As String) As String
                Dim DA As New SqlClient.SqlDataAdapter(sql, DB)
                Dim DT As New DataTable
                DA.Fill(DT)
                Return DT.Columns(0).DataType.FullName.ToString
            End Function
            Public Function ReturnDataTable(ByVal sql As String) As DataTable
                Dim DA As New SqlClient.SqlDataAdapter(sql, DB)
                Dim DT As New DataTable
                DA.SelectCommand.CommandTimeout = 4000
                DA.Fill(DT)
                Return DT
            End Function
            Public Function ReturnDataAdapter(ByVal sql As String) As SqlDataAdapter
                Dim cmd As New SqlClient.SqlCommand(sql, DB)
                Dim DA As New SqlClient.SqlDataAdapter(cmd)
                Return DA
            End Function
            Public Function ReturnCommand(ByVal sql As String) As SqlClient.SqlCommand
                Dim cmd As New SqlClient.SqlCommand(sql, DB)
                Return cmd
            End Function

            Public Sub ExportDataToExcel(ByVal ExcelFilePath As String, ByVal QueryToExport As String, ByVal SheetName As String, Optional ByVal Timeout As Integer = 15, Optional Ranges() As st_ExcelRange = Nothing)
                Try
                    Dim ex As New EAMS.OfficeAutomation.Excels
                    Dim DA As New SqlClient.SqlDataAdapter(QueryToExport, DB)
                    Dim DT As New DataTable
                    Dim inx As Integer = 0
                    Dim iny As Integer = 0
                    DA.SelectCommand.CommandTimeout = Timeout
                    DA.Fill(DT)
                    Dim RCount As Integer = DT.Rows.Count
                    If RCount = 0 Then RCount = 1
                    ex.SetRange("A1", RCount, DT.Columns.Count)
                    'ex.FormateRange
                    'Column Name
                    RaiseEvent ProgressCount(DT.Rows.Count)
                    For iny = 1 To DT.Columns.Count
                        ex.Write(1, iny, DT.Columns(iny - 1).ColumnName)
                    Next

                    ''''''''''''''''''''''''''''''''''
                    For inx = 1 To DT.Rows.Count
                        For iny = 1 To DT.Columns.Count
                            If DT.Rows(inx - 1).Item(iny - 1).ToString <> "" Then ex.Write(inx + 1, iny, DT.Rows(inx - 1).Item(iny - 1).ToString)
                            Application.DoEvents()
                            RaiseEvent ExportProgress(inx)
                        Next
                    Next
                    If IsNothing(Ranges) Then
                        Dim R(1) As EAMS.DataBaseTools.SQLServerTools.st_ExcelRange
                        R(0) = New EAMS.DataBaseTools.SQLServerTools.st_ExcelRange
                        R(1) = New EAMS.DataBaseTools.SQLServerTools.st_ExcelRange
                        '--------------------------------
                        R(0).ColName = "A1"
                        R(0).RangeColCount = DT.Columns.Count
                        R(0).BackColor = Color.FromArgb(75, 130, 199)
                        R(0).ColumnWidth = 20
                        R(0).FontSize = 9
                        R(0).ForeColor = Color.White
                        R(0).IsBold = True
                        R(0).LineWeight = 2
                        R(0).BorderColor = Color.White
                        R(0).HAlign = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                        R(0).VAlign = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
                        '--------------------------------------
                        R(1).ColName = "A2"
                        R(1).RangeColCount = DT.Columns.Count
                        R(1).BackColor = Color.White
                        R(1).ColumnWidth = 20
                        R(1).FontSize = 9
                        R(1).ForeColor = Color.Black
                        R(1).IsBold = False
                        R(1).LineWeight = 1
                        R(1).BorderColor = Color.Black
                        R(1).HAlign = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter
                        R(1).VAlign = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft
                        '-------------------------------------------
                        Ranges = R
                    End If
                    If Not IsNothing(Ranges) Then
                        For Rinx As Integer = 0 To Ranges.Count - 1
                            ex.SetRange(Ranges(Rinx).ColName, RCount + 1, Ranges(Rinx).RangeColCount)
                            ex.FormateRange(Ranges(Rinx).BackColor, Ranges(Rinx).ForeColor, Ranges(Rinx).FontSize, Ranges(Rinx).IsBold, Ranges(Rinx).ColumnWidth, Ranges(Rinx).VAlign, Ranges(Rinx).HAlign)
                            ex.CellBorder(Ranges(Rinx).BorderColor, Ranges(Rinx).LineWeight)
                        Next
                    End If

                    ex.Save(ExcelFilePath, SheetName)
                    ex.Close()
                    RaiseEvent ExportingComplete()
                Catch ex As Exception
                    RaiseEvent err(ex.Message)
                End Try
            End Sub

            Public Sub ExportDataHeaderToExcel(ByVal ExcelFilePath As String, ByVal QueryToExport As String, ByVal SheetName As String)
                Try
                    Dim ex As New EAMS.OfficeAutomation.Excels
                    Dim DA As New SqlClient.SqlDataAdapter(QueryToExport, DB)
                    Dim DT As New DataTable
                    DA.Fill(DT)
                    ex.SetRange("A1", 1, DT.Columns.Count)
                    'Column Name
                    For iny = 1 To DT.Columns.Count
                        ex.Write(1, iny, DT.Columns(iny - 1).ColumnName)
                    Next
                    ''''''''''''''''''''''''''''''''''
                    ex.FormateRange(Color.Navy, Color.White, 11, False, 15)
                    ex.Save(ExcelFilePath, SheetName)
                    ex.Close()
                    RaiseEvent ExportingComplete()
                Catch ex As Exception
                    RaiseEvent err(ex.Message)
                End Try
            End Sub

            Public Sub SaveDataTableToFile(ByVal DT As DataTable, filePath As String)
                DT.WriteXml(filePath, XmlWriteMode.WriteSchema)
            End Sub

            Public Overloads Function ExcutResult(ByVal SQL As String, Optional ByRef RowEffected As Integer = 0) As String
                Dim temp As String = ""
                Try
                    ' SaveLog(SQL)
                    Dim DA As New SqlClient.SqlDataAdapter(SQL, DB)
                    Dim DT As New DataTable
                    DA.Fill(DT)
                    RowEffected = DT.Rows.Count
                    If DT.Rows.Count <> 0 Then
                        temp = DT.Rows(0).Item(0).ToString
                    Else
                        Return ""
                    End If
                Catch ex As Exception
                    RaiseEvent err(ex.Message)
                    ErrorLog(ex.Message)
                End Try
                Return temp
            End Function
            Public Overloads Function ExcutResultFromFile(ByVal SQLTextFilePath As String, Optional ByRef RowEffected As Integer = 0) As String
                Dim temp As String = "", sql As String = ""
                Try
                    Dim obj As New System.IO.StreamReader(SQLTextFilePath)
                    sql = obj.ReadToEnd
                    obj.Close()
                    Dim DA As New SqlClient.SqlDataAdapter(sql, DB)
                    Dim DT As New DataTable
                    DA.Fill(DT)
                    RowEffected = DT.Rows.Count
                    If DT.Rows.Count <> 0 Then
                        temp = DT.Rows(0).Item(0).ToString
                    Else
                        Return ""
                    End If
                Catch ex As Exception
                    RaiseEvent err(ex.Message)
                End Try
                Return temp
            End Function
            Public Overloads Sub FillBYCount(ByRef lst As ListBox, ByVal SQLData As String)
                Dim DA As New System.Data.SqlClient.SqlDataAdapter(SQLData, DB)
                Dim DT As New DataTable
                Dim x As Integer = 0
                Dim y As Integer = 0

                Dim lsttempData As New ListBox
                Dim lsttempCount As New ListBox

                DA.Fill(DT)

                lst.Items.Clear()
                lsttempCount.Items.Clear()
                lsttempData.Items.Clear()

                For x = 0 To DT.Rows.Count - 1
                    lsttempData.Items.Add(DT.Rows(x).Item(0).ToString)
                Next

                For x = 0 To DT.Rows.Count - 1
                    lsttempCount.Items.Add(DT.Rows(x).Item(1).ToString)
                Next

                For x = 0 To lsttempData.Items.Count - 1
                    For y = 0 To Val(lsttempCount.Items(x)) - 1
                        lst.Items.Add(lsttempData.Items(x))
                    Next
                Next
            End Sub
            Public Function FillDataTable(ByVal SQL As String) As DataTable
                Dim DA As New System.Data.SqlClient.SqlDataAdapter(SQL, DB)
                Dim DT As New DataTable
                DA.Fill(DT)
                Return DT
            End Function
            Public Overloads Sub FillBYCount(ByRef lst As ListBox, ByVal SQLData As String, ByVal Count As Integer)
                Dim DA As New System.Data.SqlClient.SqlDataAdapter(SQLData, DB)
                Dim DT As New DataTable
                Dim x As Integer = 0
                Dim y As Integer = 0
                Dim lsttempData As New ListBox

                DA.Fill(DT)

                lst.Items.Clear()

                For x = 0 To DT.Rows.Count - 1
                    lsttempData.Items.Add(DT.Rows(x).Item(0).ToString)
                Next

                For x = 0 To lsttempData.Items.Count - 1
                    For y = 0 To Count - 1
                        lst.Items.Add(lsttempData.Items(x))
                    Next
                Next
            End Sub
            Public Overloads Sub FillExcuteFromFile(ByRef lst As ListBox, ByVal FilePath As String)
                Dim sql As String = ""
                Dim obj As New System.IO.StreamReader(FilePath)
                sql = obj.ReadToEnd
                obj.Close()
                Dim DA As New System.Data.SqlClient.SqlDataAdapter(sql, DB)
                Dim DT As New DataTable, x As Integer
                DA.Fill(DT)
                lst.Items.Clear()
                For x = 0 To DT.Rows.Count - 1
                    lst.Items.Add(DT.Rows(x).Item(0).ToString)
                Next
            End Sub
            Public Overloads Sub Fill(ByRef lst As ListBox, ByVal SQL As String)
                Dim DA As New System.Data.SqlClient.SqlDataAdapter(SQL, DB)
                Dim DT As New DataTable, x As Integer
                DA.SelectCommand.CommandTimeout = 0
                DA.Fill(DT)
                lst.Items.Clear()
                For x = 0 To DT.Rows.Count - 1
                    lst.Items.Add(DT.Rows(x).Item(0).ToString)
                Next
            End Sub
            Public Overloads Sub Fill(ByRef cmb As ComboBox, ByVal SQL As String)
                Dim DA As New System.Data.SqlClient.SqlDataAdapter(SQL, DB)
                Dim DT As New DataTable, x As Integer
                DA.SelectCommand.CommandTimeout = 0
                DA.Fill(DT)
                cmb.Items.Clear()
                For x = 0 To DT.Rows.Count - 1
                    cmb.Items.Add(DT.Rows(x).Item(0).ToString)
                Next
            End Sub
            Public Overloads Sub Fill(ByRef cmb As ToolStripComboBox, ByVal SQL As String)
                Dim DA As New System.Data.SqlClient.SqlDataAdapter(SQL, DB)
                Dim DT As New DataTable, x As Integer
                DA.SelectCommand.CommandTimeout = 0
                DA.Fill(DT)
                cmb.Items.Clear()
                For x = 0 To DT.Rows.Count - 1
                    cmb.Items.Add(DT.Rows(x).Item(0).ToString)
                Next
            End Sub
            Public Function GetImage(ByVal SQL As String) As System.Drawing.Image
                Dim _SqlRetVal As Object = Nothing
                Dim _Image As System.Drawing.Image = Nothing
                Try
                    Dim _SqlCommand As New System.Data.SqlClient.SqlCommand(SQL, DB)
                    _SqlRetVal = _SqlCommand.ExecuteScalar()
                    _SqlCommand.Dispose()
                    _SqlCommand = Nothing
                Catch _Exception As Exception
                    RaiseEvent err(_Exception.Message)
                    Return Nothing
                End Try

                ' convert object to image
                Try
                    Dim _ImageData(-1) As Byte
                    _ImageData = CType(_SqlRetVal, Byte())
                    Dim _MemoryStream As New System.IO.MemoryStream(_ImageData)
                    _Image = System.Drawing.Image.FromStream(_MemoryStream)
                Catch _Exception As Exception
                    Console.WriteLine(_Exception.Message)
                    Return Nothing
                End Try
                Return _Image
            End Function
            Public Overloads Sub SaveImage(ByVal ImagePath As String, ByVal TableName As String, ByVal FieldName As String, ByVal FieldKey As String, ByVal KeyString As String, Optional ByVal KeyInteger As Integer = -99)
                Try
                    Dim sql As String = ""
                    If KeyInteger = -99 Then
                        sql = "update " & TableName & " set " & FieldName & " = " & "(@BLOBData) where " & FieldKey & " = '" & KeyString & "'"
                    Else
                        sql = "update " & TableName & " set " & FieldName & " = " & "(@BLOBData) where " & FieldKey & " = " & KeyInteger
                    End If

                    Dim cmd As New SqlCommand(sql, DB)
                    Dim fsBLOBFile As New FileStream(ImagePath, FileMode.Open, FileAccess.Read)
                    Dim bytBLOBData(fsBLOBFile.Length) As [Byte]
                    fsBLOBFile.Read(bytBLOBData, 0, bytBLOBData.Length)
                    fsBLOBFile.Close()
                    Dim prm As New SqlParameter("@BLOBData", SqlDbType.VarBinary, bytBLOBData.Length, ParameterDirection.Input, False, 0, 0, Nothing, DataRowVersion.Current, bytBLOBData)
                    cmd.Parameters.Add(prm)
                    cmd.ExecuteNonQuery()
                Catch ex As Exception
                    RaiseEvent err(ex.Message)
                End Try

            End Sub
            Public Sub GetImageByte(ByRef _ImageData() As Byte, ByVal sql As String)
                Dim _SqlRetVal As Object = Nothing
                Dim _Image As System.Drawing.Image = Nothing
                Try
                    Dim _SqlCommand As New System.Data.SqlClient.SqlCommand(sql, DB)
                    _SqlRetVal = _SqlCommand.ExecuteScalar()
                    _SqlCommand.Dispose()
                    _SqlCommand = Nothing
                Catch _Exception As Exception
                    RaiseEvent err(_Exception.Message)
                End Try

                ' convert object to image
                Try
                    'Dim _ImageData(-1) As Byte
                    _ImageData = CType(_SqlRetVal, Byte())
                Catch _Exception As Exception
                End Try
            End Sub
            Public Overloads Sub SaveImage(ByRef bytBLOBData() As [Byte], ByVal TableName As String, ByVal FieldName As String, ByVal FieldKey As String, ByVal KeyString As String)
                Try
                    Dim sql As String = "update " & TableName & " set " & FieldName & " = " & "(@BLOBData) where " & FieldKey & " = '" & KeyString & "'"
                    Dim cmd As New SqlCommand(sql, DB)
                    Dim prm As New SqlParameter("@BLOBData", SqlDbType.VarBinary, bytBLOBData.Length, ParameterDirection.Input, False, 0, 0, Nothing, DataRowVersion.Current, bytBLOBData)
                    cmd.Parameters.Add(prm)
                    cmd.ExecuteNonQuery()
                Catch ex As Exception
                    RaiseEvent err(ex.Message)
                End Try

            End Sub
            Public Overloads Sub SaveImage(ByVal Pic As PictureBox, ByVal TableName As String, ByVal FieldName As String, ByVal FieldKey As String, ByVal KeyString As String)
                Try
                    Dim sql As String = ""
                    If Pic.Image Is Nothing Then
                        sql = "update " & TableName & " set " & FieldName & " = " & "null where " & FieldKey & " = '" & KeyString & "'"
                        Me.ExcuteNoneResult(sql)
                        Exit Sub
                    End If
                    Dim ms As New IO.MemoryStream
                    Pic.Image.Save(ms, Imaging.ImageFormat.Jpeg)
                    Dim bytes() As Byte = ms.GetBuffer()

                    sql = "update " & TableName & " set " & FieldName & " = " & "(@BLOBData) where " & FieldKey & " = '" & KeyString & "'"
                    Dim cmd As New SqlCommand(sql, DB)
                    Dim prm As New SqlParameter("@BLOBData", SqlDbType.VarBinary, bytes.Length, ParameterDirection.Input, False, 0, 0, Nothing, DataRowVersion.Current, bytes)
                    cmd.Parameters.Add(prm)
                    cmd.ExecuteNonQuery()
                Catch ex As Exception
                    RaiseEvent err(ex.Message)
                End Try

            End Sub
            Public Function CheckPrivliage(ByVal UserID As Integer, ByRef Form As Form, Optional ByVal UserInterfaceTableName As String = "user_inter") As Boolean
                Dim IsSave, IsEdit, IsDelete, IsPrint As Boolean
                Dim FormID As Integer = Val(Me.ExcutResult("select forms_id from forms where form_name ='" & Form.Name & "'"))
                Try
                    If Me.ExcutResult("select user_id from " & UserInterfaceTableName & " where user_id =" & UserID) = "" Then
                        Return False
                    Else
                        If Me.ExcutResult("select user_id from " & UserInterfaceTableName & " where user_id =" & UserID & " and forms_id =" & FormID) = "" Then
                            Return False
                        End If
                        If Me.ExcutResult("select issave from " & UserInterfaceTableName & " where user_id =" & UserID & " and forms_id =" & FormID) = "" Then
                            IsSave = False
                        Else
                            IsSave = CBool(Me.ExcutResult("select issave from " & UserInterfaceTableName & " where user_id =" & UserID & " and forms_id =" & FormID))
                        End If
                        If Me.ExcutResult("select isedit from " & UserInterfaceTableName & " where user_id =" & UserID & " and forms_id =" & FormID) = "" Then
                            IsEdit = False
                        Else
                            IsEdit = CBool(Me.ExcutResult("select isedit from " & UserInterfaceTableName & " where user_id =" & UserID & " and forms_id =" & FormID))
                        End If
                        If Me.ExcutResult("select isdelete from " & UserInterfaceTableName & " where user_id =" & UserID & " and forms_id =" & FormID) = "" Then
                            IsDelete = False
                        Else
                            IsDelete = CBool(Me.ExcutResult("select isdelete from " & UserInterfaceTableName & " where user_id =" & UserID & " and forms_id =" & FormID))
                        End If
                        If Me.ExcutResult("select isprint from " & UserInterfaceTableName & " where user_id =" & UserID & " and forms_id =" & FormID) = "" Then
                            IsPrint = False
                        Else
                            IsPrint = CBool(Me.ExcutResult("select isprint from " & UserInterfaceTableName & " where user_id =" & UserID & " and forms_id =" & FormID))
                        End If

                        Dim inx As Integer = 0
                        Dim iny As Integer = 0
                        Dim ts As New ToolStrip




                        Select Case Form.Controls(inx).Name
                            Case "fsSave", "fsSave2", "fsSave3", "fsSave4", "fsSave5"
                                Form.Controls(inx).Enabled = IsSave
                            Case "fsEdit", "fsEdit1", "fsEdit2", "fsEdit3", "fsEdit4"
                                Form.Controls(inx).Enabled = IsEdit
                            Case "fsDelete", "fsDelete1", "fsDelete2", "fsDelete3", "fsDelete4", "fsDelete5", "fsDelete6"
                                Form.Controls(inx).Enabled = IsDelete
                            Case "fsPrint", "fsPrint1", "fsPrint2", "fsPrint3"
                                Form.Controls(inx).Enabled = IsPrint
                        End Select


                        Return True
                    End If
                    Return False
                Catch ex As Exception
                    Return False
                End Try
                Return False
            End Function


            Public Function CheckReportPrivliage(ByVal UserID As Integer, ByRef ReportName As String, Optional ByVal UserInterfaceTableName As String = "user_inter") As Boolean
                Dim IsPrint As Boolean = False
                Dim FormID As Integer = Val(Me.ExcutResult("select forms_id from forms where form_name ='" & ReportName & "'"))
                Try
                    If Me.ExcutResult("select user_id from " & UserInterfaceTableName & " where user_id =" & UserID) = "" Then
                        Return False
                    Else
                        IsPrint = CBool(Me.ExcutResult("select isprint from " & UserInterfaceTableName & " where user_id =" & UserID & " and forms_id =" & FormID))
                        Dim inx As Integer = 0
                        Dim iny As Integer = 0
                        Dim ts As New ToolStrip
                        Return IsPrint
                    End If
                Catch ex As Exception
                    Return False
                End Try
                Return False
            End Function

            Public Sub Formatgrid(ByRef grd As DataGridView)
                Dim inx As Integer = 0
                Dim flag As Boolean = True

                For inx = 0 To grd.Rows.Count - 1
                    If flag Then
                        grd.Rows(inx).DefaultCellStyle.BackColor = Color.LightSkyBlue
                    Else
                        grd.Rows(inx).DefaultCellStyle.BackColor = Color.CadetBlue
                    End If
                    flag = Not flag
                    grd.Rows(inx).DefaultCellStyle.ForeColor = Color.Black
                    grd.Rows(inx).DefaultCellStyle.SelectionBackColor = Color.Black
                    grd.Rows(inx).DefaultCellStyle.SelectionForeColor = Color.White
                Next
                grd.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
                grd.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            End Sub
#End Region

#Region "Database Methods"

            Public Sub ShrinkLogFile()
                Dim sql As String = ""
                sql = "USE [master]" & vbCrLf
                sql &= "ALTER DATABASE dbXXX SET RECOVERY SIMPLE WITH NO_WAIT" & vbCrLf
                sql &= "use dbXXX" & vbCrLf
                sql &= "DBCC SHRINKFILE(dbXXX, 1)" & vbCrLf
                sql &= "DBCC SHRINKFILE('dbXXX_log', 0, TRUNCATEONLY)" & vbCrLf
                sql = Replace(sql, "dbXXX", Me.DataBaseName)
                Me.ExcuteNoneResult(sql)
            End Sub
            Public Sub BulkInsert(ByRef DT As DataTable, ByVal TableName As String)
                ' Perform an initial count on the destination table.
                ' Dim commandRowCount As New SqlCommand("SELECT COUNT(*) FROM " & TableName & ";", Me.DB)
                ' Dim countStart As Long = System.Convert.ToInt32(commandRowCount.ExecuteScalar())
                Using bulkCopy As SqlBulkCopy = New SqlBulkCopy(Me.DB)
                    bulkCopy.DestinationTableName = TableName
                    Try
                        bulkCopy.BatchSize = 1000
                        bulkCopy.WriteToServer(DT)
                    Catch ex As Exception
                        MsgBox(ex.Message)
                    End Try
                End Using
            End Sub
            Public Function DatabsePath() As String
                Do Until DB.State = ConnectionState.Open
                    DB.Open()
                Loop
                Return ExcutResult("select filename from dbo.sysfiles where fileid = 1")
            End Function
            Public Sub TerminateConnection(ByVal DBName)
                Dim tmp As Integer = 0
                Dim sql As String = "select spid from master..sysprocesses where dbid=db_id('" & DBName & "')"
                Me.ConnectTODatabase("master")
                Try
                    Dim DA As New SqlClient.SqlDataAdapter(sql, DB)
                    Dim DT As New DataTable
                    DA.Fill(DT)
                    For tmp = 0 To DT.Rows.Count - 1
                        ExcuteNoneResult("kill " & DT.Rows(tmp).Item(0).ToString)
                    Next
                    RaiseEvent ConnectionTerminated()
                Catch ex As Exception
                    RaiseEvent err(ex.Message)
                End Try
            End Sub
            Public Sub Connect()
                Dim SqlServerCon As String = ""
                If Pass = "" Then
                    SqlServerCon = String.Format("packet size=4096;integrated security=SSPI;data Source={0};initial catalog={1};persist security info=False", Trim(DataBaseLocation), DataBaseName)
                Else
                    SqlServerCon = "Persist Security Info=False;User ID=" & UserName & ";Initial Catalog=" & DataBaseName & ";Data Source=" & Trim(DataBaseLocation) & ";password=" & Pass
                End If
                Try
                    If (DB.State = System.Data.ConnectionState.Open Or DB.State = System.Data.ConnectionState.Connecting) Then
                        DB.Close()
                    End If
                    DB.ConnectionString = SqlServerCon
                    Do Until DB.State = ConnectionState.Open
                        DB.Open()
                    Loop
                    RaiseEvent Connnected()
                Catch ex As Exception
                    RaiseEvent err(ex.Message)
                    'frmMain.Disconnected = True
                    RaiseEvent Disconnected()
                End Try
            End Sub
            Public Sub ConnectTODatabase(ByVal DbName As String)
                DataBaseName = DbName
                Dim SQLServerCon As String = ""
                If UserName = "" Then
                    SQLServerCon = "packet size=4096;integrated security=SSPI;data Source=" & Trim(DataBaseLocation) & ";initial catalog=" & DataBaseName & ";persist security info=False"
                Else
                    SQLServerCon = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & UserName & ";Initial Catalog=" & DataBaseName & ";Data Source=" & Trim(DataBaseLocation) & ";password=" & Pass
                End If
                Try
                    If (DB.State = System.Data.ConnectionState.Open Or DB.State = System.Data.ConnectionState.Connecting) Then
                        DB.Close()
                    End If
                    DB.ConnectionString = SQLServerCon
                    DB.Open()
                    RaiseEvent ConnnectedTOMaster()
                Catch ex As Exception
                    RaiseEvent Disconnected()
                    RaiseEvent err(ex.Message)
                End Try
            End Sub
            Public Sub Close()
                Try
                    Do Until DB.State = ConnectionState.Closed
                        DB.Close()
                        DB = Nothing
                        GC.Collect()
                        DB = New SqlClient.SqlConnection
                    Loop
                    RaiseEvent Disconnected()
                Catch ex As Exception
                    RaiseEvent err(ex.Message)
                End Try
                GC.Collect()
            End Sub

            Public Sub AttachDatabase(ByVal MDFPath As String, ByVal LDFPath As String, ByVal RestoreName As String)
                Try
                    DeAttachDatabase(RestoreName)
                    Dim tmpSql As String = "sp_attach_db N'" & RestoreName & "' , N'" & MDFPath & "', N'" & LDFPath & "'"
                    ExcuteNoneResult(tmpSql)
                    RaiseEvent RestoreDatabaseComplete()
                Catch ex As Exception
                    RaiseEvent err(ex.Message)
                End Try

            End Sub

            Public Sub SwitchToDataBase(Optional ByVal DBName As String = "")
                If DBName = "" Then DBName = DataBaseName
                DataBaseName = DBName
                If Trim(DataBaseName) = "" Then
                    RaiseEvent err("DataBase Name Not Set")
                    Exit Sub
                End If
                If DB.State = ConnectionState.Closed Then
                    DB.Open()
                End If
                SQLServerCmd.Connection = DB
                SQLServerCmd.CommandType = System.Data.CommandType.Text
                SQLServerCmd.CommandText = "use " & DataBaseName
                Try
                    SQLServerCmd.ExecuteNonQuery()
                    If DB.Database.ToUpper = DataBaseName.ToUpper Then
                    Else
                        RaiseEvent err("Switched DataBase Failed")
                    End If
                Catch ex As Exception
                    RaiseEvent err("Switching Failed")
                End Try
            End Sub

            Private Sub SaveLog(ByRef q As String)
                Try
                    SQLServerCmd.Connection = DB
                    SQLServerCmd.CommandType = System.Data.CommandType.Text
                    SQLServerCmd.CommandText = "insert into _sys_log (qry,user_name) values ('" & Replace(q, "'", "", , , CompareMethod.Binary) & "','" & My.User.Name & "')"
                    SQLServerCmd.ExecuteNonQuery()
                Catch ex As Exception

                End Try
            End Sub
            Public Sub ExcuteNoneResultFromFile(ByVal Path As String, Optional ByVal Timeout As Integer = 15)
                Dim obj As New System.IO.StreamReader(Path)
                Dim temp() As String
                Dim Sql As String = obj.ReadToEnd
                obj.Close()
                SQLServerCmd.Connection = DB
                SQLServerCmd.CommandType = System.Data.CommandType.Text
                Try
                    If InStr(Sql, "go") > 0 Then
                        temp = Split(Sql, "go")
                        For inx As Integer = 0 To UBound(temp)
                            SQLServerCmd.CommandText = temp(inx)
                            SQLServerCmd.CommandTimeout = Timeout
                            SQLServerCmd.ExecuteNonQuery() 'Rslt
                        Next
                    Else
                        SQLServerCmd.CommandText = Sql
                        SQLServerCmd.CommandTimeout = Timeout
                        SQLServerCmd.ExecuteNonQuery() 'Rslt
                    End If
                Catch ex As Exception
                    ErrorLog(Path)
                    RaiseEvent err("Excuted Failed")
                    RaiseEvent err(ex.Message)
                End Try
            End Sub
            Public Function ExcuteNoneResult(ByVal Query As String, Optional ByVal Timeout As Integer = 15) As Integer
                'Dim Rslt As Integer
                'SaveLog(Query)
                SQLServerCmd.Connection = DB
                SQLServerCmd.CommandType = System.Data.CommandType.Text
                SQLServerCmd.CommandText = Query
                Try
                    'Rslt = SQLServerCmd.ExecuteNonQuery()
                    'If Rslt > 0 Then
                    SQLServerCmd.CommandTimeout = Timeout
                    Return SQLServerCmd.ExecuteNonQuery() 'Rslt
                    'Else
                    ' Return 0
                    'End If
                Catch ex As Exception
                    ErrorLog(Query)
                    RaiseEvent err("Excuted Failed")
                    RaiseEvent err(ex.Message)
                End Try
                Return -1
            End Function
            Private Sub ErrorLog(ByVal er As String)
                Try
                    Dim w As New System.IO.StringWriter
                    w.WriteLine(er)
                    IO.File.WriteAllText(String.Format("{0}\error_log\{1}.txt", Application.StartupPath, Format(Now, "yyyyMMddHHmmss")), w.ToString, System.Text.Encoding.Unicode)
                Catch ex As Exception
                End Try
            End Sub
            Public Sub DeAttachDatabase(ByVal DBName As String)
                Try
                    Close()
                    ConnectTODatabase("master")
                    Dim tmpSql As String = "sp_detach_db N'" & DBName & "' , N'true'"
                    TerminateConnection(DBName)
                    ExcuteNoneResult(tmpSql)
                    RaiseEvent DeattachedComplete()
                Catch ex As Exception
                    RaiseEvent err(ex.Message)
                End Try
            End Sub

            Public Function GetGridItemSelected(ByRef grd As DataGridView, ByVal ColumnIndex As Integer) As String
                Try
                    If grd.Rows.Count = 0 Then
                        Return ""
                    End If
                    If grd.SelectedCells(0).Selected = False Then
                        Return ""
                    Else
                        Return (grd.Rows(grd.SelectedCells(0).RowIndex).Cells(ColumnIndex).Value)
                    End If
                Catch ex As Exception
                    Return ""
                End Try
                Return ""
            End Function

            Public Overloads Sub FillGrd(ByRef grd As DataGridView, ByVal SQL As String, Optional ByVal Timeout As Integer = 15)
                Dim DT As New DataTable
                Dim DA As New SqlClient.SqlDataAdapter(SQL, DB)
                Try
                    DA.SelectCommand.CommandTimeout = Timeout
                    DA.Fill(DT)
                    grd.DataSource = DT
                Catch ex As Exception
                End Try
            End Sub
            Public Overloads Sub FillGrdExcuteFromFile(ByRef grd As DataGridView, ByVal filePath As String)
                Dim sql As String = ""
                Dim obj As New System.IO.StreamReader(filePath)
                sql = obj.ReadToEnd
                obj.Close()

                Dim DT As New DataTable
                Dim DA As New SqlClient.SqlDataAdapter(sql, DB)
                Try
                    DA.Fill(DT)
                    grd.DataSource = DT
                Catch ex As Exception
                End Try
            End Sub
            Public Overloads Function ReturnDataTableExcuteFromFile(ByVal SQLFilePath As String, Optional ByVal Timeout As Integer = 15) As DataTable
                Dim SQL As String = ""
                Dim dt As New DataTable
                Try
                    Dim obj As New System.IO.StreamReader(SQLFilePath)
                    SQL = obj.ReadToEnd
                    obj.Close()
                    Dim DA As New SqlClient.SqlDataAdapter(SQL, DB)
                    DA.SelectCommand.CommandTimeout = Timeout
                    DA.Fill(dt)
                    Return dt
                Catch ex As Exception
                    Return Nothing
                End Try
            End Function
            Public Overloads Function ReturnDataExcuteFromFile(ByVal SQLFilePath As String) As DataSet
                Dim SQL As String = ""
                Dim ds As New DataSet
                Try
                    Dim obj As New System.IO.StreamReader(SQLFilePath)
                    SQL = obj.ReadToEnd
                    obj.Close()
                    Dim DA As New SqlClient.SqlDataAdapter(SQL, DB)
                    DA.Fill(ds)
                    Return ds
                Catch ex As Exception
                    Return Nothing
                End Try
            End Function
#End Region

#Region "Properties"
            Public ReadOnly Property GetLogFileSize As Double
                Get
                    Dim sql As String = "SELECT (size * 8)/1024.0 AS size_in_mb FROM  DBXXX.sys.database_files WHERE data_space_id = 0"
                    Return Val(Me.ExcutResult(Replace(sql, "DBXXX", Me.DataBaseName)))
                End Get
            End Property
            Public ReadOnly Property GetDatabases() As Collection
                Get
                    Try
                        SwitchToDataBase("master")
                        Dim DA As New System.Data.SqlClient.SqlDataAdapter("select name from sysdatabases order by dbid", DB)
                        Dim DT As New System.Data.DataTable
                        SwitchToDataBase(DataBaseName)
                        Dim col As New Collection, x As Integer = 0
                        For x = 0 To DT.Rows.Count - 1
                            col.Add(DT.Rows(x).Item(0).ToString)
                        Next
                        Return col
                    Catch ex As Exception
                    End Try
                    Return Nothing
                End Get
            End Property
            Public ReadOnly Property DatabaseExists(ByVal DBName As String) As Boolean
                Get
                    Try
                        SwitchToDataBase("master")
                        Dim DA As New System.Data.SqlClient.SqlDataAdapter("select name from sysdatabases where name ='" & DBName & "'", DB)
                        Dim DT As New System.Data.DataTable
                        SwitchToDataBase(DataBaseName)
                        Dim col As New Collection, x As Integer = 0
                        If DT.Rows.Count = 0 Then
                            Return False
                        Else
                            Return True
                        End If
                    Catch ex As Exception
                    End Try
                    Return False
                End Get
            End Property
            Public Function DBStatus() As ConnectionState
                Return DB.State
            End Function
            Public Function GetServerInfo() As String
                Dim dt As New DataTable, inx As Integer = 0
                Dim tmp As String = ""
                dt = Me.ReturnDataTable("select serverproperty('MachineName') MachineName,serverproperty('ServerName') ServerInstanceName,replace(cast(serverproperty('Edition')as varchar),'Edition','') EditionInstalled,serverproperty('productVersion') ProductBuildLevel,serverproperty('productLevel') SPLevel,serverproperty('Collation') Collation_Type,serverproperty('IsClustered') [IsClustered?],convert(varchar,getdate(),102) QueryDate,case when  exists (select * from msdb.dbo.backupset where name like 'data protector%') then 'HPDPused' else 'NotOnDRP' end as DRP")
                For inx = 0 To dt.Columns.Count - 1
                    tmp &= dt.Columns(inx).ColumnName & ": " & dt.Rows(0).Item(inx).ToString & vbCrLf
                Next
                Return tmp
            End Function
#End Region

            '#Region "Crystal Report"
            '            Public Sub GetCrConnectionInfo(ByVal ReportSource As String, ByRef v As CrystalDecisions.Windows.Forms.CrystalReportViewer)
            '                Dim crtableLogoninfos As New CrystalDecisions.Shared.TableLogOnInfos
            '                Dim crtableLogoninfo As New CrystalDecisions.Shared.TableLogOnInfo
            '                Dim crConnectionInfo As New CrystalDecisions.Shared.ConnectionInfo()
            '                Dim CrTables As CrystalDecisions.CrystalReports.Engine.Tables
            '                Dim CrTable As CrystalDecisions.CrystalReports.Engine.Table
            '                Dim crReportDocument As New CrystalDecisions.CrystalReports.Engine.ReportDocument()
            '                crReportDocument.Load(ReportSource)
            '                crReportDocument.DataSourceConnections.Clear()

            '                CrTables = crReportDocument.Database.Tables
            '                For Each CrTable In CrTables
            '                    'crtableLogoninfo = CrTable.LogOnInfo
            '                    crConnectionInfo.Type = CrystalDecisions.Shared.ConnectionInfoType.SQL
            '                    crConnectionInfo.LogonProperties.Clear()
            '                    crConnectionInfo.AllowCustomConnection = True
            '                    'crConnectionInfo.UserID = Me.UserName
            '                    'crConnectionInfo.Password = ""
            '                    crConnectionInfo.ServerName = Me.DataBaseLocation
            '                    crConnectionInfo.DatabaseName = Me.DataBaseName
            '                    crConnectionInfo.IntegratedSecurity = True
            '                    crtableLogoninfo.ConnectionInfo = crConnectionInfo
            '                    CrTable.ApplyLogOnInfo(crtableLogoninfo)
            '                    'CrTable.Location.Substring(CrTable.Location.LastIndexOf(".") + 1)
            '                Next

            '                crReportDocument.ReportOptions.EnableSaveDataWithReport = False
            '                'v.RefreshReport()
            '                v.ReportSource = crReportDocument
            '                crReportDocument.Refresh()
            '            End Sub
            '            Public Sub CrExportToPDF(ByVal ReportSource As String, ByVal PDFFullPath As String, Optional Parameter As String = "")
            '                Dim crtableLogoninfo As New CrystalDecisions.Shared.TableLogOnInfo()
            '                Dim crConnectionInfo As New CrystalDecisions.Shared.ConnectionInfo()
            '                Dim CrTables As CrystalDecisions.CrystalReports.Engine.Tables
            '                Dim CrTable As CrystalDecisions.CrystalReports.Engine.Table
            '                Dim CrExportOptions As New CrystalDecisions.Shared.ExportOptions
            '                Dim CrDiskFileDestinationOptions As New CrystalDecisions.Shared.DiskFileDestinationOptions()
            '                Dim CrFormatTypeOptions As New CrystalDecisions.Shared.PdfRtfWordFormatOptions()
            '                Dim crReportDocument As New CrystalDecisions.CrystalReports.Engine.ReportDocument()
            '                crReportDocument.Load(ReportSource)
            '                'Get the DB Connection
            '                CrTables = crReportDocument.Database.Tables
            '                For Each CrTable In CrTables
            '                    'crtableLogoninfo = CrTable.LogOnInfo
            '                    crConnectionInfo.ServerName = Me.DataBaseLocation
            '                    crConnectionInfo.DatabaseName = Me.DataBaseName
            '                    crConnectionInfo.IntegratedSecurity = True
            '                    crtableLogoninfo.ConnectionInfo = crConnectionInfo
            '                    CrTable.ApplyLogOnInfo(crtableLogoninfo)
            '                    CrTable.Location.Substring(CrTable.Location.LastIndexOf(".") + 1)
            '                Next
            '                crReportDocument.ReportOptions.EnableSaveDataWithReport = False
            '                '
            '                CrDiskFileDestinationOptions.DiskFileName = PDFFullPath
            '                CrExportOptions = crReportDocument.ExportOptions
            '                With CrExportOptions
            '                    .ExportDestinationType = CrystalDecisions.Shared.ExportDestinationType.DiskFile
            '                    .ExportFormatType = CrystalDecisions.Shared.ExportFormatType.PortableDocFormat
            '                    .DestinationOptions = CrDiskFileDestinationOptions
            '                    .FormatOptions = CrFormatTypeOptions
            '                End With
            '                If Parameter <> "" Then
            '                    crReportDocument.SetParameterValue(0, Parameter)
            '                End If
            '                crReportDocument.Export()
            '            End Sub
            '#End Region

        End Class

        Public Class AccessDBTools
            Private DB As New System.Data.OleDb.OleDbConnection
            Public DataBaseName As String = ""
            Public DataBaseLocation As String = "."
            Public UserName As String = ""
            Public Pass As String = ""
            Private AccessCmd As New OleDb.OleDbCommand
            Public Event err(ByVal errName As String)
            Public Event Connnected()
            Public Event ConnnectedTOMaster()
            Public Event Disconnected()
            Public Event RestoreDatabaseComplete()
            Public Event ConnectionTerminated()
            Public Event DeattachedComplete()
            Public Event ExportProgress(ByVal inx As Integer)
            Public Event ExportingComplete()
            Public Event ExportingDataCount(ByVal Inx As Integer)





#Region "Data Methods" '
            Public Function ReturnFieldDataType(ByVal sql As String) As String
                Dim DA As New OleDb.OleDbDataAdapter(sql, DB)
                Dim DT As New DataTable
                DA.Fill(DT)
                Return DT.Columns(0).DataType.FullName.ToString
            End Function
            Public Function ReturnDataTable(ByVal sql As String) As DataTable
                Dim DA As New OleDb.OleDbDataAdapter(sql, DB)
                DA.SelectCommand.CommandTimeout = 400
                Dim DT As New DataTable
                DA.Fill(DT)
                Return DT
            End Function
            Public Function ReturnDataAdapter(ByVal sql As String) As OleDb.OleDbDataAdapter
                Dim cmd As New OleDb.OleDbCommand(sql, DB)
                Dim DA As New OleDb.OleDbDataAdapter(cmd)
                Return DA
            End Function
            Public Function ReturnCommand(ByVal sql As String) As OleDb.OleDbCommand
                Dim cmd As New OleDb.OleDbCommand(sql, DB)
                Return cmd
            End Function

            Public Sub SaveDataTableToFile(ByVal DT As DataTable, filePath As String)
                DT.WriteXml(filePath, XmlWriteMode.WriteSchema)
            End Sub

            Public Overloads Function ExcutResult(ByVal SQL As String, Optional ByRef RowEffected As Integer = 0) As String
                Dim temp As String = ""
                Try
                    ' SaveLog(SQL)
                    Dim DA As New OleDb.OleDbDataAdapter(SQL, DB)
                    Dim DT As New DataTable
                    DA.Fill(DT)
                    RowEffected = DT.Rows.Count
                    If DT.Rows.Count <> 0 Then
                        temp = DT.Rows(0).Item(0).ToString
                    Else
                        Return ""
                    End If
                Catch ex As Exception
                    RaiseEvent err(ex.Message)
                    ErrorLog(ex.Message)
                End Try
                Return temp
            End Function
            Public Overloads Function ExcutResultFromFile(ByVal SQLTextFilePath As String, Optional ByRef RowEffected As Integer = 0) As String
                Dim temp As String = "", sql As String = ""
                Try
                    Dim obj As New System.IO.StreamReader(SQLTextFilePath)
                    sql = obj.ReadToEnd
                    obj.Close()
                    Dim DA As New OleDb.OleDbDataAdapter(sql, DB)
                    Dim DT As New DataTable
                    DA.Fill(DT)
                    RowEffected = DT.Rows.Count
                    If DT.Rows.Count <> 0 Then
                        temp = DT.Rows(0).Item(0).ToString
                    Else
                        Return ""
                    End If
                Catch ex As Exception
                    RaiseEvent err(ex.Message)
                End Try
                Return temp
            End Function
            Public Overloads Sub FillBYCount(ByRef lst As ListBox, ByVal SQLData As String)
                Dim DA As New OleDb.OleDbDataAdapter(SQLData, DB)
                Dim DT As New DataTable
                Dim x As Integer = 0
                Dim y As Integer = 0

                Dim lsttempData As New ListBox
                Dim lsttempCount As New ListBox

                DA.Fill(DT)

                lst.Items.Clear()
                lsttempCount.Items.Clear()
                lsttempData.Items.Clear()

                For x = 0 To DT.Rows.Count - 1
                    lsttempData.Items.Add(DT.Rows(x).Item(0).ToString)
                Next

                For x = 0 To DT.Rows.Count - 1
                    lsttempCount.Items.Add(DT.Rows(x).Item(1).ToString)
                Next

                For x = 0 To lsttempData.Items.Count - 1
                    For y = 0 To Val(lsttempCount.Items(x)) - 1
                        lst.Items.Add(lsttempData.Items(x))
                    Next
                Next
            End Sub
            Public Function FillDataTable(ByVal SQL As String) As DataTable
                Dim DA As New OleDb.OleDbDataAdapter(SQL, DB)
                Dim DT As New DataTable
                DA.Fill(DT)
                Return DT
            End Function
            Public Overloads Sub FillBYCount(ByRef lst As ListBox, ByVal SQLData As String, ByVal Count As Integer)
                Dim DA As New OleDb.OleDbDataAdapter(SQLData, DB)
                Dim DT As New DataTable
                Dim x As Integer = 0
                Dim y As Integer = 0
                Dim lsttempData As New ListBox

                DA.Fill(DT)

                lst.Items.Clear()

                For x = 0 To DT.Rows.Count - 1
                    lsttempData.Items.Add(DT.Rows(x).Item(0).ToString)
                Next

                For x = 0 To lsttempData.Items.Count - 1
                    For y = 0 To Count - 1
                        lst.Items.Add(lsttempData.Items(x))
                    Next
                Next
            End Sub
            Public Overloads Sub FillExcuteFromFile(ByRef lst As ListBox, ByVal FilePath As String)
                Dim sql As String = ""
                Dim obj As New System.IO.StreamReader(FilePath)
                sql = obj.ReadToEnd
                obj.Close()
                Dim DA As New OleDb.OleDbDataAdapter(sql, DB)
                Dim DT As New DataTable, x As Integer
                DA.Fill(DT)
                lst.Items.Clear()
                For x = 0 To DT.Rows.Count - 1
                    lst.Items.Add(DT.Rows(x).Item(0).ToString)
                Next
            End Sub
            Public Overloads Sub Fill(ByRef lst As ListBox, ByVal SQL As String)
                Dim DA As New OleDb.OleDbDataAdapter(SQL, DB)
                Dim DT As New DataTable, x As Integer
                DA.Fill(DT)
                lst.Items.Clear()
                For x = 0 To DT.Rows.Count - 1
                    lst.Items.Add(DT.Rows(x).Item(0).ToString)
                Next
            End Sub
            Public Overloads Sub Fill(ByRef cmb As ComboBox, ByVal SQL As String)
                Dim DA As New OleDb.OleDbDataAdapter(SQL, DB)
                Dim DT As New DataTable, x As Integer
                DA.Fill(DT)
                cmb.Items.Clear()
                For x = 0 To DT.Rows.Count - 1
                    cmb.Items.Add(DT.Rows(x).Item(0).ToString)
                Next
            End Sub
            Public Overloads Sub Fill(ByRef cmb As ToolStripComboBox, ByVal SQL As String)
                Dim DA As New OleDb.OleDbDataAdapter(SQL, DB)
                Dim DT As New DataTable, x As Integer
                DA.Fill(DT)
                cmb.Items.Clear()
                For x = 0 To DT.Rows.Count - 1
                    cmb.Items.Add(DT.Rows(x).Item(0).ToString)
                Next
            End Sub
            Public Function GetImage(ByVal SQL As String) As System.Drawing.Image
                Dim _SqlRetVal As Object = Nothing
                Dim _Image As System.Drawing.Image = Nothing
                Try
                    Dim _SqlCommand As New OleDb.OleDbCommand(SQL, DB)
                    _SqlRetVal = _SqlCommand.ExecuteScalar()
                    _SqlCommand.Dispose()
                    _SqlCommand = Nothing
                Catch _Exception As Exception
                    RaiseEvent err(_Exception.Message)
                    Return Nothing
                End Try

                ' convert object to image
                Try
                    Dim _ImageData(-1) As Byte
                    _ImageData = CType(_SqlRetVal, Byte())
                    Dim _MemoryStream As New System.IO.MemoryStream(_ImageData)
                    _Image = System.Drawing.Image.FromStream(_MemoryStream)
                Catch _Exception As Exception
                    Console.WriteLine(_Exception.Message)
                    Return Nothing
                End Try
                Return _Image
            End Function
            Public Overloads Sub SaveImage(ByVal ImagePath As String, ByVal TableName As String, ByVal FieldName As String, ByVal FieldKey As String, ByVal KeyString As String, Optional ByVal KeyInteger As Integer = -99)
                Try
                    Dim sql As String = ""
                    If KeyInteger = -99 Then
                        sql = "update " & TableName & " set " & FieldName & " = " & "(@BLOBData) where " & FieldKey & " = '" & KeyString & "'"
                    Else
                        sql = "update " & TableName & " set " & FieldName & " = " & "(@BLOBData) where " & FieldKey & " = " & KeyInteger
                    End If

                    Dim cmd As New OleDb.OleDbCommand(sql, DB)
                    Dim fsBLOBFile As New FileStream(ImagePath, FileMode.Open, FileAccess.Read)
                    Dim bytBLOBData(fsBLOBFile.Length) As [Byte]
                    fsBLOBFile.Read(bytBLOBData, 0, bytBLOBData.Length)
                    fsBLOBFile.Close()
                    Dim prm As New SqlParameter("@BLOBData", SqlDbType.VarBinary, bytBLOBData.Length, ParameterDirection.Input, False, 0, 0, Nothing, DataRowVersion.Current, bytBLOBData)
                    cmd.Parameters.Add(prm)
                    cmd.ExecuteNonQuery()
                Catch ex As Exception
                    RaiseEvent err(ex.Message)
                End Try

            End Sub
            Public Sub GetImageByte(ByRef _ImageData() As Byte, ByVal sql As String)
                Dim _SqlRetVal As Object = Nothing
                Dim _Image As System.Drawing.Image = Nothing
                Try
                    Dim _SqlCommand As New OleDb.OleDbCommand(sql, DB)
                    _SqlRetVal = _SqlCommand.ExecuteScalar()
                    _SqlCommand.Dispose()
                    _SqlCommand = Nothing
                Catch _Exception As Exception
                    RaiseEvent err(_Exception.Message)
                End Try

                ' convert object to image
                Try
                    'Dim _ImageData(-1) As Byte
                    _ImageData = CType(_SqlRetVal, Byte())
                Catch _Exception As Exception
                End Try
            End Sub
            Public Overloads Sub SaveImage(ByRef bytBLOBData() As [Byte], ByVal TableName As String, ByVal FieldName As String, ByVal FieldKey As String, ByVal KeyString As String)
                Try
                    Dim sql As String = "update " & TableName & " set " & FieldName & " = " & "(@BLOBData) where " & FieldKey & " = '" & KeyString & "'"
                    Dim cmd As New OleDb.OleDbCommand(sql, DB)
                    Dim prm As New SqlParameter("@BLOBData", SqlDbType.VarBinary, bytBLOBData.Length, ParameterDirection.Input, False, 0, 0, Nothing, DataRowVersion.Current, bytBLOBData)
                    cmd.Parameters.Add(prm)
                    cmd.ExecuteNonQuery()
                Catch ex As Exception
                    RaiseEvent err(ex.Message)
                End Try

            End Sub
            Public Overloads Sub SaveImage(ByVal Pic As PictureBox, ByVal TableName As String, ByVal FieldName As String, ByVal FieldKey As String, ByVal KeyString As String)
                Try
                    Dim sql As String = ""
                    If Pic.Image Is Nothing Then
                        sql = "update " & TableName & " set " & FieldName & " = " & "null where " & FieldKey & " = '" & KeyString & "'"
                        Me.ExcuteNoneResult(sql)
                        Exit Sub
                    End If
                    Dim ms As New IO.MemoryStream
                    Pic.Image.Save(ms, Imaging.ImageFormat.Jpeg)
                    Dim bytes() As Byte = ms.GetBuffer()

                    sql = "update " & TableName & " set " & FieldName & " = " & "(@BLOBData) where " & FieldKey & " = '" & KeyString & "'"
                    Dim cmd As New OleDb.OleDbCommand(sql, DB)
                    Dim prm As New SqlParameter("@BLOBData", SqlDbType.VarBinary, bytes.Length, ParameterDirection.Input, False, 0, 0, Nothing, DataRowVersion.Current, bytes)
                    cmd.Parameters.Add(prm)
                    cmd.ExecuteNonQuery()
                Catch ex As Exception
                    RaiseEvent err(ex.Message)
                End Try

            End Sub
            Public Function CheckPrivliage(ByVal UserID As Integer, ByRef Form As Form, Optional ByVal UserInterfaceTableName As String = "user_inter") As Boolean
                'Dim IsSave, IsEdit, IsDelete, IsPrint As Boolean
                'Dim FormID As Integer = Val(Me.ExcutResult("select forms_id from forms where form_name ='" & Form.Name & "'"))
                'Try
                '    If Me.ExcutResult("select user_id from " & UserInterfaceTableName & " where user_id =" & UserID) = "" Then
                '        Return False
                '    Else
                '        If Me.ExcutResult("select user_id from " & UserInterfaceTableName & " where user_id =" & UserID & " and forms_id =" & FormID) = "" Then
                '            Return False
                '        End If
                '        If Me.ExcutResult("select issave from " & UserInterfaceTableName & " where user_id =" & UserID & " and forms_id =" & FormID) = "" Then
                '            IsSave = False
                '        Else
                '            IsSave = CBool(Me.ExcutResult("select issave from " & UserInterfaceTableName & " where user_id =" & UserID & " and forms_id =" & FormID))
                '        End If
                '        If Me.ExcutResult("select isedit from " & UserInterfaceTableName & " where user_id =" & UserID & " and forms_id =" & FormID) = "" Then
                '            IsEdit = False
                '        Else
                '            IsEdit = CBool(Me.ExcutResult("select isedit from " & UserInterfaceTableName & " where user_id =" & UserID & " and forms_id =" & FormID))
                '        End If
                '        If Me.ExcutResult("select isdelete from " & UserInterfaceTableName & " where user_id =" & UserID & " and forms_id =" & FormID) = "" Then
                '            IsDelete = False
                '        Else
                '            IsDelete = CBool(Me.ExcutResult("select isdelete from " & UserInterfaceTableName & " where user_id =" & UserID & " and forms_id =" & FormID))
                '        End If
                '        If Me.ExcutResult("select isprint from " & UserInterfaceTableName & " where user_id =" & UserID & " and forms_id =" & FormID) = "" Then
                '            IsPrint = False
                '        Else
                '            IsPrint = CBool(Me.ExcutResult("select isprint from " & UserInterfaceTableName & " where user_id =" & UserID & " and forms_id =" & FormID))
                '        End If

                '        Dim inx As Integer = 0
                '        Dim iny As Integer = 0
                '        Dim ts As New ToolStrip
                '        Dim bb As New DevExpress.XtraBars.Ribbon.RibbonControl
                '        For inx = 0 To Form.Controls.Count - 1

                '            If TypeOf (Form.Controls(inx)) Is ToolStrip Then
                '                ts = Form.Controls(inx)
                '                For iny = 0 To ts.Items.Count - 1
                '                    Select Case ts.Items(iny).Name
                '                        Case "tsSave", "tsSave2", "tsSave3"
                '                            ts.Items(iny).Enabled = IsSave
                '                        Case "tsEdit", "tsEdit1", "tsEdit2", "tsEdit3", "tsEdit4"
                '                            ts.Items(iny).Enabled = IsEdit
                '                        Case "tsDelete", "tsDelete1", "tsDelete2", "tsDelete3", "tsDelete4", "tsDelete5", "tsDelete6"
                '                            ts.Items(iny).Enabled = IsDelete
                '                        Case "tsPrint", "tsPrint1", "tsPrint2", "tsPrint3"
                '                            ts.Items(iny).Enabled = IsPrint
                '                    End Select
                '                Next
                '            End If

                '            If TypeOf (Form.Controls(inx)) Is DevExpress.XtraBars.Ribbon.RibbonControl Then
                '                bb = Form.Controls(inx)
                '                For iny = 0 To bb.Items.Count - 1
                '                    Select Case bb.Items(iny).Name
                '                        Case "bbSave", "bbSave2", "bbSave3"
                '                            bb.Items(iny).Visibility = Math.Abs(Val(Not IsSave))
                '                        Case "bbEdit", "bbEdit1", "bbEdit2", "bbEdit3", "bbEdit4"
                '                            bb.Items(iny).Visibility = Math.Abs(Val(Not IsEdit))
                '                        Case "bbDelete", "bbDelete1", "bbDelete2", "bbDelete3", "bbDelete4", "bbDelete5", "bbDelete6"
                '                            bb.Items(iny).Visibility = Math.Abs(Val(Not IsDelete))
                '                        Case "bbPrint", "bbPrint1", "bbPrint2", "bbPrint3"
                '                            bb.Items(iny).Visibility = Math.Abs(Val(Not IsPrint))
                '                    End Select
                '                Next
                '            End If

                '            Select Case Form.Controls(inx).Name
                '                Case "fsSave", "fsSave2", "fsSave3", "fsSave4", "fsSave5"
                '                    Form.Controls(inx).Enabled = IsSave
                '                Case "fsEdit", "fsEdit1", "fsEdit2", "fsEdit3", "fsEdit4"
                '                    Form.Controls(inx).Enabled = IsEdit
                '                Case "fsDelete", "fsDelete1", "fsDelete2", "fsDelete3", "fsDelete4", "fsDelete5", "fsDelete6"
                '                    Form.Controls(inx).Enabled = IsDelete
                '                Case "fsPrint", "fsPrint1", "fsPrint2", "fsPrint3"
                '                    Form.Controls(inx).Enabled = IsPrint
                '            End Select

                '        Next
                '        Return True
                '    End If
                '    Return False
                'Catch ex As Exception
                '    Return False
                'End Try
                'Return False
                Return True
            End Function


            Public Function CheckReportPrivliage(ByVal UserID As Integer, ByRef ReportName As String, Optional ByVal UserInterfaceTableName As String = "user_inter") As Boolean
                Dim IsPrint As Boolean = False
                Dim FormID As Integer = Val(Me.ExcutResult("select forms_id from forms where form_name ='" & ReportName & "'"))
                Try
                    If Me.ExcutResult("select user_id from " & UserInterfaceTableName & " where user_id =" & UserID) = "" Then
                        Return False
                    Else
                        IsPrint = CBool(Me.ExcutResult("select isprint from " & UserInterfaceTableName & " where user_id =" & UserID & " and forms_id =" & FormID))
                        Dim inx As Integer = 0
                        Dim iny As Integer = 0
                        Dim ts As New ToolStrip
                        Return IsPrint
                    End If
                Catch ex As Exception
                    Return False
                End Try
                Return False
            End Function

            Public Sub Formatgrid(ByRef grd As DataGridView)
                Dim inx As Integer = 0
                Dim flag As Boolean = True

                For inx = 0 To grd.Rows.Count - 1
                    If flag Then
                        grd.Rows(inx).DefaultCellStyle.BackColor = Color.LightSkyBlue
                    Else
                        grd.Rows(inx).DefaultCellStyle.BackColor = Color.CadetBlue
                    End If
                    flag = Not flag
                    grd.Rows(inx).DefaultCellStyle.ForeColor = Color.Black
                    grd.Rows(inx).DefaultCellStyle.SelectionBackColor = Color.Black
                    grd.Rows(inx).DefaultCellStyle.SelectionForeColor = Color.White
                Next
                grd.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
                grd.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
            End Sub
#End Region

#Region "Database Methods"
            Public Function DatabsePath() As String
                Do Until DB.State = ConnectionState.Open
                    DB.Open()
                Loop
                Return ExcutResult("select filename from dbo.sysfiles where fileid = 1")
            End Function
            Public Sub TerminateConnection(ByVal DBName)
                Dim tmp As Integer = 0
                Dim sql As String = "select spid from master..sysprocesses where dbid=db_id('" & DBName & "')"
                Me.ConnectTODatabase("master")
                Try
                    Dim DA As New OleDb.OleDbDataAdapter(sql, DB)
                    Dim DT As New DataTable
                    DA.Fill(DT)
                    For tmp = 0 To DT.Rows.Count - 1
                        ExcuteNoneResult("kill " & DT.Rows(tmp).Item(0).ToString)
                    Next
                    RaiseEvent ConnectionTerminated()
                Catch ex As Exception
                    RaiseEvent err(ex.Message)
                End Try
            End Sub
            Public Sub Connect()
                'Dim SqlServerCon As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & Trim(DataBaseLocation) & ";Persist Security Info=False;"
                Dim SqlServerCon As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Trim(DataBaseLocation) & ";Persist Security Info=False;"
                Try
                    If (DB.State = System.Data.ConnectionState.Open Or DB.State = System.Data.ConnectionState.Connecting) Then
                        DB.Close()
                    End If
                    DB.ConnectionString = SqlServerCon
                    Do Until DB.State = ConnectionState.Open
                        DB.Open()
                    Loop
                    RaiseEvent Connnected()
                Catch ex As Exception
                    RaiseEvent err(ex.Message)
                    RaiseEvent Disconnected()
                End Try
            End Sub
            Public Sub ConnectTODatabase(ByVal DbName As String)
                DataBaseName = DbName
                'Dim SQLServerCon As String = ""
                Dim SqlServerCon As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Trim(DataBaseLocation) & ";Persist Security Info=False;"
                Try
                    If (DB.State = System.Data.ConnectionState.Open Or DB.State = System.Data.ConnectionState.Connecting) Then
                        DB.Close()
                    End If
                    DB.ConnectionString = SQLServerCon
                    DB.Open()
                    RaiseEvent ConnnectedTOMaster()
                Catch ex As Exception
                    RaiseEvent err(ex.Message)
                End Try
            End Sub
            Public Sub Close()
                Try
                    Do Until DB.State = ConnectionState.Closed
                        DB.Close()
                        DB = Nothing
                        GC.Collect()
                        DB = New OleDb.OleDbConnection
                    Loop
                    RaiseEvent Disconnected()
                Catch ex As Exception
                    RaiseEvent err(ex.Message)
                End Try
                GC.Collect()
            End Sub

            Public Function ExcuteNoneResult(ByVal Query As String, Optional ByVal Timeout As Integer = 15) As Integer
                'Dim Rslt As Integer
                'SaveLog(Query)
                AccessCmd.Connection = DB
                AccessCmd.CommandType = System.Data.CommandType.Text
                AccessCmd.CommandText = Query
                Try
                    'Rslt = SQLServerCmd.ExecuteNonQuery()
                    'If Rslt > 0 Then
                    AccessCmd.CommandTimeout = Timeout
                    Return AccessCmd.ExecuteNonQuery() 'Rslt
                    'Else
                    ' Return 0
                    'End If
                Catch ex As Exception
                    ErrorLog(Query)
                    RaiseEvent err("Excuted Failed")
                    RaiseEvent err(ex.Message)
                End Try
                Return -1
            End Function
            Private Sub ErrorLog(ByVal er As String)
                Try
                    Dim w As New System.IO.StringWriter
                    w.WriteLine(er)
                    IO.File.WriteAllText(String.Format("{0}\error_log\{1}.txt", Application.StartupPath, Format(Now, "yyyyMMddHHmmss")), w.ToString, System.Text.Encoding.Unicode)
                Catch ex As Exception
                End Try
            End Sub
            Public Sub DeAttachDatabase(ByVal DBName As String)
                Try
                    Close()
                    ConnectTODatabase("master")
                    Dim tmpSql As String = "sp_detach_db N'" & DBName & "' , N'true'"
                    TerminateConnection(DBName)
                    ExcuteNoneResult(tmpSql)
                    RaiseEvent DeattachedComplete()
                Catch ex As Exception
                    RaiseEvent err(ex.Message)
                End Try
            End Sub

            Public Function GetGridItemSelected(ByRef grd As DataGridView, ByVal ColumnIndex As Integer) As String
                Try
                    If grd.Rows.Count = 0 Then
                        Return ""
                    End If
                    If grd.SelectedCells(0).Selected = False Then
                        Return ""
                    Else
                        Return (grd.Rows(grd.SelectedCells(0).RowIndex).Cells(ColumnIndex).Value)
                    End If
                Catch ex As Exception
                    Return ""
                End Try
                Return ""
            End Function

            Public Overloads Sub FillGrd(ByRef grd As DataGridView, ByVal SQL As String)
                Dim DT As New DataTable
                Dim DA As New OleDb.OleDbDataAdapter(SQL, DB)
                Try
                    DA.Fill(DT)
                    grd.DataSource = DT
                Catch ex As Exception
                End Try
            End Sub
            Public Overloads Sub FillGrdExcuteFromFile(ByRef grd As DataGridView, ByVal filePath As String)
                Dim sql As String = ""
                Dim obj As New System.IO.StreamReader(filePath)
                sql = obj.ReadToEnd
                obj.Close()

                Dim DT As New DataTable
                Dim DA As New OleDb.OleDbDataAdapter(sql, DB)
                Try
                    DA.Fill(DT)
                    grd.DataSource = DT
                Catch ex As Exception
                End Try
            End Sub

            Public Overloads Function ReturnDataExcuteFromFile(ByVal SQLFilePath As String) As DataSet
                Dim SQL As String = ""
                Dim ds As New DataSet
                Try
                    Dim obj As New System.IO.StreamReader(SQLFilePath)
                    SQL = obj.ReadToEnd
                    obj.Close()
                    Dim DA As New OleDb.OleDbDataAdapter(SQL, DB)
                    DA.Fill(ds)
                Catch ex As Exception
                    Return Nothing
                End Try
                Return ds
            End Function
#End Region

#Region "Properties"
            Public Function Connection() As System.Data.OleDb.OleDbConnection
                Return Me.DB
            End Function

            Public Function DBStatus() As ConnectionState
                Return DB.State
            End Function
#End Region

        End Class

        Public Class ApplicationSettings
            '==========Application Settings============ table (app_set)
            'company_name  
            'logo As Image 
            'barcode_header 
            'barcode_text1  
            'barcode_text2 
            'print_label_after_supply
            'Current_currency   
            'comp_address 
            'comp_phone  
            'comp_fax  
            'comp_web 
            'comp_mail 
            'date_format
            '=================================

            '==========Accounting Settings============ table (app_acc_settings)
            'dealer_dis_type
            'def_supp_tax
            'def_sell_tax
            'sup_tax_calc
            'sel_tax_calc
            'update_item_def
            'prt_rece_order
            'float_digit
            '=================================

#Region "Properties"
            Public Property CompanyName As String
                Get
                    Return DB.ExcutResult("select company_name from app_set where app_set_id=1")
                End Get
                Set(ByVal value As String)
                    DB.ExcuteNoneResult("update app_set set company_name='" & value & "' where app_set_id =1")
                End Set
            End Property

            Public ReadOnly Property CompanyLogo() As Image
                Get
                    Return DB.GetImage("select logo from app_set where app_set_id=1")
                End Get
            End Property

            Public WriteOnly Property SetCompanyLogo() As String
                Set(ByVal value As String)
                    DB.SaveImage(value, "app_set", "logo", "app_set_id", "", 1)
                End Set
            End Property

            Public Property BarcodeHeader As String
                Get
                    Return DB.ExcutResult("select barcode_header from app_set where app_set_id=1")
                End Get
                Set(ByVal value As String)
                    DB.ExcuteNoneResult("update app_set set barcode_header='" & value & "' where app_set_id =1")
                End Set
            End Property

            Public Property BarcodeText1 As String
                Get
                    Return DB.ExcutResult("select barcode_text1 from app_set where app_set_id=1")
                End Get
                Set(ByVal value As String)
                    DB.ExcuteNoneResult("update app_set set barcode_text1='" & value & "' where app_set_id =1")
                End Set
            End Property

            Public Property BarcodeText2 As String
                Get
                    Return DB.ExcutResult("select barcode_text2 from app_set where app_set_id=1")
                End Get
                Set(ByVal value As String)
                    DB.ExcuteNoneResult("update app_set set barcode_text2='" & value & "' where app_set_id =1")
                End Set
            End Property

            Public Property PrintLabelAfterSupply As Boolean
                Get
                    Return CBool(Val(DB.ExcutResult("select Print_Label_after_supply from app_set where app_set_id=1")))
                End Get
                Set(ByVal value As Boolean)
                    DB.ExcuteNoneResult("update app_set set barcode_text2=" & Val(value) & " where app_set_id =1")
                End Set
            End Property

            Public Property CurrentCurrency As String
                Get
                    Return DB.ExcutResult("select Currencies from Currency where curr_id='" & DB.ExcutResult("select Current_currency_id from app_set where app_set_id=1") & "'")
                End Get
                Set(ByVal value As String)
                    DB.ExcuteNoneResult("update app_set set Current_currency_id='" & DB.ExcutResult("select curr_id from Currency where Currencies='" & value & "'") & "' where app_set_id=1")
                End Set
            End Property

            Public Property CompanyAddress As String
                Get
                    Return DB.ExcutResult("select comp_address from app_set where app_set_id=1")
                End Get
                Set(ByVal value As String)
                    DB.ExcuteNoneResult("update app_set set comp_address='" & value & "' where app_set_id =1")
                End Set
            End Property

            Public Property CompanyPhone As String
                Get
                    Return DB.ExcutResult("select comp_phone from app_set where app_set_id=1")
                End Get
                Set(ByVal value As String)
                    DB.ExcuteNoneResult("update app_set set comp_phone='" & value & "' where app_set_id =1")
                End Set
            End Property

            Public Property CompanyFax As String
                Get
                    Return DB.ExcutResult("select comp_fax from app_set where app_set_id=1")
                End Get
                Set(ByVal value As String)
                    DB.ExcuteNoneResult("update app_set set comp_fax='" & value & "' where app_set_id =1")
                End Set
            End Property

            Public Property CompanyWebAddress As String
                Get
                    Return DB.ExcutResult("select comp_web from app_set where app_set_id=1")
                End Get
                Set(ByVal value As String)
                    DB.ExcuteNoneResult("update app_set set comp_web='" & value & "' where app_set_id =1")
                End Set
            End Property

            Public Property CompanyMailAddress As String
                Get
                    Return DB.ExcutResult("select comp_mail from app_set where app_set_id=1")
                End Get
                Set(ByVal value As String)
                    DB.ExcuteNoneResult("update app_set set comp_mail='" & value & "' where app_set_id =1")
                End Set
            End Property

            Public Property DateFormate As String
                Get
                    Return DB.ExcutResult("select date_format from app_set where app_set_id=1")
                End Get
                Set(ByVal value As String)
                    DB.ExcuteNoneResult("update app_set set date_format='" & value & "' where app_set_id =1")
                End Set
            End Property

            Public Property DealerDiscountTypePerItem As Boolean
                Get
                    Return CBool(DB.ExcutResult("select dealer_dis_type from app_acc_settings where acc_set_id=1"))
                End Get
                Set(ByVal value As Boolean)
                    DB.ExcuteNoneResult("update app_acc_settings set dealer_dis_type=" & Val(value) & " where acc_set_id =1")
                End Set
            End Property

            Public Property DefaultSuppTax As Double
                Get
                    Return Val(DB.ExcutResult("select def_supp_tax from app_acc_settings where acc_set_id=1"))
                End Get
                Set(ByVal value As Double)
                    DB.ExcuteNoneResult("update app_acc_settings set def_supp_tax=" & Val(value) & " where acc_set_id =1")
                End Set
            End Property

            Public Property DefaultSellTax As Double
                Get
                    Return Val(DB.ExcutResult("select def_sell_tax from app_acc_settings where acc_set_id=1"))
                End Get
                Set(ByVal value As Double)
                    DB.ExcuteNoneResult("update app_acc_settings set def_sell_tax=" & Val(value) & " where acc_set_id =1")
                End Set
            End Property

            Public Property SellTaxCalc As String
                Get
                    Return DB.ExcutResult("select sell_tax_calc from app_acc_settings where acc_set_id=1")
                End Get
                Set(ByVal value As String)
                    DB.ExcuteNoneResult("update app_acc_settings set sell_tax_calc='" & value & "' where acc_set_id =1")
                End Set
            End Property

            Public Property SuppTaxCalc As String
                Get
                    Return DB.ExcutResult("select sup_tax_calc from app_acc_settings where acc_set_id=1")
                End Get
                Set(ByVal value As String)
                    DB.ExcuteNoneResult("update app_acc_settings set sup_tax_calc='" & value & "' where acc_set_id =1")
                End Set
            End Property

            Public ReadOnly Property SuppTaxCalc(ByVal SuppId As Integer) As String
                Get
                    Return DB.ExcutResult("select tax_calc from supply where sub_id=" & SuppId)
                End Get
            End Property

            Public Property UpdateItemDefaults As Boolean
                Get
                    Return CBool(DB.ExcutResult("select update_item_def from app_acc_settings where acc_set_id=1"))
                End Get
                Set(ByVal value As Boolean)
                    DB.ExcuteNoneResult("update app_acc_settings set update_item_def=" & Val(value) & " where acc_set_id =1")
                End Set
            End Property

            Public Property PrintReceiveOrder As Boolean
                Get
                    Return CBool(DB.ExcutResult("select prt_rece_order from app_acc_settings where acc_set_id=1"))
                End Get
                Set(ByVal value As Boolean)
                    DB.ExcuteNoneResult("update app_acc_settings set prt_rece_order=" & Val(value) & " where acc_set_id =1")
                End Set
            End Property

            Public Property FloatDigit As Byte
                Get
                    Return DB.ExcutResult("select float_digit from app_acc_settings where acc_set_id=1")
                End Get
                Set(ByVal value As Byte)
                    DB.ExcuteNoneResult("update app_acc_settings set float_digit=" & Val(value) & " where acc_set_id =1")
                End Set
            End Property
#End Region

        End Class

    End Namespace

    Namespace Graphical

        Public Class GradientButton
            Inherits System.Windows.Forms.Control

            Private isMouseOver As Boolean = False
            Private isMouseDown As Boolean = False
            Private isMouseUp As Boolean = False
            Private isFocused As Boolean = False


#Region "Shape"
            Enum ButtonShape
                Ellipse = 1
                Rectangle = 2
                TriangleUp = 3
                TriangleDown = 4
                TriangleLeft = 5
                TriangleRight = 6
            End Enum

            Private mShape As ButtonShape = ButtonShape.Rectangle

            <Description("The Shape of the Button")> _
            <Category("Appearance")> _
            Property Shape() As ButtonShape
                Get
                    Return mShape
                End Get
                Set(ByVal value As ButtonShape)
                    mShape = value
                    Me.Invalidate()
                End Set
            End Property

            <Description("Set the Radius of All four Corners to the Same Value")> _
           <Category("Edges")> _
           <RefreshProperties(RefreshProperties.Repaint)> _
           <DefaultValue(1)> _
            Property All() As Integer
                Get
                    Return mAll
                End Get
                Set(ByVal Value As Integer)
                    If Value < 1 Then Value = 1
                    'If Value > Me.Height Then Value = Me.Height
                    mAll = Value
                    mUL = Value
                    mUR = Value
                    mLL = Value
                    mLR = Value
                    Me.Invalidate()
                End Set

            End Property


            <Description("Set the Radius of the Upper Left Corner")> _
            <Category("Edges")> _
           <DefaultValue(1)> _
            Property UL() As Integer
                Get
                    Return mUL
                End Get
                Set(ByVal Value As Integer)
                    If Value < 1 Then Value = 1
                    'If Value > Me.Height Then Value = Me.Height
                    mUL = Value
                    Me.Invalidate()
                End Set
            End Property

            <Description("Set the Radius of the Upper Right Corner")> _
            <Category("Edges")> _
           <DefaultValue(1)> _
            Property UR() As Integer
                Get
                    Return mUR
                End Get
                Set(ByVal Value As Integer)
                    If Value < 1 Then Value = 1
                    'If Value > Me.Height Then Value = Me.Height
                    mUR = Value
                    Me.Invalidate()
                End Set
            End Property

            <Description("Set the Radius of the Lower Left Corner")> _
           <Category("Edges")> _
          <DefaultValue(1)> _
            Property LL() As Integer
                Get
                    Return mLL
                End Get
                Set(ByVal Value As Integer)
                    If Value < 1 Then Value = 1
                    'If Value > Me.Height Then Value = Me.Height
                    mLL = Value
                    Me.Invalidate()
                End Set
            End Property

            <Description("Set the Radius of the Lower Right Corner")> _
            <Category("Edges")> _
           <DefaultValue(1)> _
            Property LR() As Integer
                Get
                    Return mLR
                End Get
                Set(ByVal Value As Integer)
                    If Value < 1 Then Value = 1
                    'If Value > Me.Height Then Value = Me.Height
                    mLR = Value
                    Me.Invalidate()
                End Set
            End Property


#End Region

            Public Enum GradientButtonStates
                Normal
                Focused
                MouseOver
                MouseDown
                Disabled
            End Enum
            Public Enum BrushType
                Linear
                Path
                Solid
            End Enum
            Public Enum GradientDirection
                BackwardDiagonal
                ForwardDiagonal
                Horizontal
                Vertical
            End Enum
            Public Enum Text_Style
                Regular
                Engrave
                Embosse
                Gradinet
                Hatch
                Reflect
                Block
                Shear
                Shadow
                Chisel
            End Enum



#Region "Variables"
            Private s As GradientButtonStates
            Private mGradientColor1 As Color = Color.White
            Private mGradientColor2 As Color = Color.White
            Private mGradientColor3 As Color = Color.Black
            Private mBrush As BrushType = BrushType.Linear
            Private mGDirection As GradientDirection = GradientDirection.Vertical
            Private mAll As Integer = 1
            Private mUL As Integer = 1
            Private mUR As Integer = 1
            Private mLL As Integer = 1
            Private mLR As Integer = 1
            Private mFPx As Decimal = 0
            Private mFPy As Decimal = 0
            Private mCPx As Decimal = 25
            Private mCPy As Decimal = 25
            Private mPos As Decimal = 0.5
            Private mImage As Image = Nothing
            Private mImgX As Integer = 10
            Private mImgY As Integer = 5
            Private mImageSize As Size = New Size(16, 16)
            Friend WithEvents AnimationTimer As System.Windows.Forms.Timer
            Dim mAngle As Short = 0
            Dim mAnimated As Boolean = False
            Dim mStyle As Text_Style = Text_Style.Regular

#End Region

#Region "Properties"

            <Description("Defines the first Color of the gradient Colors"), Category("Button Colours"), DefaultValue(GetType(Color), "Control")> _
            Property ButtonGradientColor1() As Color
                Get
                    ButtonGradientColor1 = mGradientColor1
                End Get
                Set(ByVal value As Color)
                    If Not (value.Equals(ButtonGradientColor1)) Then
                        mGradientColor1 = value
                        Me.Invalidate()
                    End If
                End Set
            End Property
            <Description("Defines the second Color of the gradient Colors"), Category("Button Colours"), DefaultValue(GetType(Color), "Control")> _
            Property ButtonGradientColor2() As Color
                Get
                    ButtonGradientColor2 = mGradientColor2
                End Get
                Set(ByVal value As Color)
                    If Not (value.Equals(ButtonGradientColor2)) Then
                        mGradientColor2 = value
                        Me.Invalidate()
                    End If
                End Set
            End Property

            <Description("Defines the Third Color of the gradient Colors"), Category("Button Colours"), DefaultValue(GetType(Color), "Control")> _
            Property ButtonGradientColor3() As Color
                Get
                    ButtonGradientColor3 = mGradientColor3
                End Get
                Set(ByVal value As Color)
                    If Not (value.Equals(ButtonGradientColor3)) Then
                        mGradientColor3 = value
                        Me.Invalidate()
                    End If
                End Set
            End Property

            <Description("Defines the Brush Style"), Category("Button Colours"), DefaultValue(1)> _
            Property ButtonBrushType() As BrushType
                Get
                    ButtonBrushType = mBrush
                End Get
                Set(ByVal value As BrushType)
                    mBrush = value
                    Me.Invalidate()
                End Set
            End Property


            <Description("Defines the Linear Gradient Direction"), Category("Button Colours"), DefaultValue(1)> _
            Property LinearGradientDirection() As GradientDirection
                Get
                    LinearGradientDirection = mGDirection
                End Get
                Set(ByVal value As GradientDirection)
                    mGDirection = value
                    Me.Invalidate()
                End Set
            End Property

            <Description("Defines the X value of the Focus Point when the Brush Style is Set to Path. Value between 0 and 1."), Category("Button Colours"), DefaultValue(0.5)> _
            Property FocusPointX() As Decimal
                Get
                    FocusPointX = mFPx
                End Get
                Set(ByVal value As Decimal)
                    If value < 0 Then value = 0
                    If value > 1 Then value = 1
                    mFPx = value
                    Me.Invalidate()
                End Set
            End Property
            <Description("Defines the Y value of the Focus Point when the Brush Style is Set to Path. Value between 0 and 1."), Category("Button Colours"), DefaultValue(0.5)> _
            Property FocusPointY() As Decimal
                Get
                    FocusPointY = mFPy
                End Get
                Set(ByVal value As Decimal)
                    If value < 0 Then value = 0
                    If value > 1 Then value = 1
                    mFPy = value
                    Me.Invalidate()
                End Set
            End Property

            <Description("Defines the X value of the Center Point when the Brush Style is Set to Path and Focus Points set to 0."), Category("Button Colours"), DefaultValue(0.5)> _
            Property CenterPointX() As Decimal
                Get
                    CenterPointX = mCPx
                End Get
                Set(ByVal value As Decimal)
                    If value < 0 Then value = 0
                    'If value > 1 Then value = 1
                    mCPx = value
                    Me.Invalidate()
                End Set
            End Property
            <Description("Defines the Y value of the Center Point when the Brush Style is Set to Path and Focus Points set to 0."), Category("Button Colours"), DefaultValue(0.5)> _
            Property CenterPointY() As Decimal
                Get
                    CenterPointY = mCPy
                End Get
                Set(ByVal value As Decimal)
                    If value < 0 Then value = 0
                    'If value > 1 Then value = 1
                    mCPy = value
                    Me.Invalidate()
                End Set
            End Property

            <Description("Defines the Position value of the 2nd Color.Value between 0 and 1."), Category("Button Colours"), DefaultValue(0.5)> _
            Property Position() As Decimal
                Get
                    Position = mPos
                End Get
                Set(ByVal value As Decimal)
                    If value < 0 Then value = 0
                    If value > 1 Then value = 1
                    mPos = value
                    Me.Invalidate()
                End Set
            End Property

            <Category("Image"), _
           Description("Defines the Button Image"), _
           DefaultValue("(None)")> _
            Public Property Image() As Image
                Get
                    Return mImage
                End Get
                Set(ByVal Value As Image)
                    mImage = Value
                    Me.Invalidate()
                End Set
            End Property


            <Category("Image"), _
           Description("Defines the X Coordinate of the Upper Left Corner of the Image Location"), _
           DefaultValue(10)> _
            Public Property ImageLocationX() As Integer
                Get
                    Return mImgX
                End Get
                Set(ByVal Value As Integer)
                    If Value < 0 Then Value = 0
                    'If Value > Me.Width - Me.ImageSize.Width Then Value = Me.Width - Me.ImageSize.Width
                    mImgX = Value
                    Me.Invalidate()
                End Set
            End Property
            <Category("Image"), _
           Description("Defines the Y Coordinate of the Upper Left Corner of the Image Location"), _
           DefaultValue(10)> _
            Public Property ImageLocationY() As Integer
                Get
                    Return mImgY
                End Get
                Set(ByVal Value As Integer)
                    If Value < 0 Then Value = 0
                    'If Value > Me.Height - Me.ImageSize.Height Then Value = Me.Height - Me.ImageSize.Height
                    mImgY = Value
                    Me.Invalidate()
                End Set
            End Property

            <Category("Image"), Description("Defines the Size of the Image"), DefaultValue("16, 16")> _
            Public Property ImageSize() As Size
                Get
                    Return mImageSize
                End Get
                Set(ByVal Value As Size)
                    mImageSize = Value
                    Me.Invalidate()
                End Set
            End Property

            <Category("Image"), Description("Defines the Image Rotation Angle. Value between 0 and 359."), DefaultValue(0)> _
            Public Property ImageRotationAngle() As Short
                Get
                    Return mAngle
                End Get
                Set(ByVal Value As Short)
                    If Value < 0 Or Value > 359 Then Value = 0
                    'If Value > 359 Then Value = 0
                    mAngle = Value
                    Me.Invalidate()
                End Set
            End Property
            <Category("Image"), Description("Defines if the Image Will Animate or not."), DefaultValue("False")> _
            Public Property ImageAnimation() As Boolean
                Get
                    Return mAnimated
                End Get
                Set(ByVal Value As Boolean)
                    mAnimated = Value
                    Me.Invalidate()
                End Set
            End Property
            <Category("Appearance"), Description("Sets the Text Style."), DefaultValue("Regular")> _
            Public Property TextStyle() As Text_Style
                Get
                    Return mStyle
                End Get
                Set(ByVal Value As Text_Style)
                    mStyle = Value
                    Me.Invalidate()
                End Set
            End Property


#End Region



#Region " Windows Form Designer generated code "

            Public Sub New()
                MyBase.New()

                'This call is required by the Windows Form Designer.
                InitializeComponent()

                'Add any initialization after the InitializeComponent() call
                MyBase.SetStyle(ControlStyles.AllPaintingInWmPaint, True)
                MyBase.SetStyle(ControlStyles.UserPaint, True)
                MyBase.SetStyle(ControlStyles.DoubleBuffer, True)
                MyBase.SetStyle(ControlStyles.ResizeRedraw, True)
                MyBase.SetStyle(ControlStyles.SupportsTransparentBackColor, True)
                'MyBase.SetStyle(ControlStyles.SupportsTransparentBackColor, True)
                'Me.BackColor = Color.Transparent
            End Sub

            'UserControl1 overrides dispose to clean up the component list.
            Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
                If disposing Then
                    If Not (components Is Nothing) Then
                        components.Dispose()
                    End If
                End If
                MyBase.Dispose(disposing)
            End Sub

            'Required by the Windows Form Designer
            Private components As System.ComponentModel.IContainer

            'NOTE: The following procedure is required by the Windows Form Designer
            'It can be modified using the Windows Form Designer.  
            'Do not modify it using the code editor.
            <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
                Me.components = New System.ComponentModel.Container
                Me.AnimationTimer = New System.Windows.Forms.Timer(Me.components)
                Me.SuspendLayout()
                '
                'AnimationTimer
                '
                Me.AnimationTimer.Interval = 500
                Me.ResumeLayout(False)

            End Sub

#End Region

            Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
                MyBase.OnPaint(e)

                'Add all painting code here

                ' Dim s As GradientButtonStates

                If Me.Enabled = False Then
                    s = GradientButtonStates.Disabled
                Else
                    If isMouseDown = False Then
                        If isMouseOver = True Then
                            s = GradientButtonStates.MouseOver
                        Else
                            If isFocused = True Then
                                s = GradientButtonStates.Focused
                            Else
                                s = GradientButtonStates.Normal
                            End If
                        End If
                    Else
                        s = GradientButtonStates.MouseDown
                    End If
                End If

                Dim sf As StringFormat

                sf = New StringFormat
                sf.LineAlignment = StringAlignment.Center
                sf.Alignment = StringAlignment.Center
                sf.Trimming = StringTrimming.EllipsisCharacter

                'draw the button shape
                Dim gp As New GraphicsPath : Dim gp2 As New GraphicsPath
                Dim rec As RectangleF = New RectangleF(0, 0, Me.Width, Me.Height)
                Dim rec2 As RectangleF = New Rectangle(0, 0, Me.Width - 1, Me.Height - 1)
                Dim rec3 As RectangleF = New Rectangle(1, 1, Me.Width - 2, Me.Height - 2)
                Dim pts() As PointF = New PointF() {}

                Select Case mShape

                    Case ButtonShape.Ellipse
                        gp.AddEllipse(rec)
                        gp2.AddEllipse(rec2)

                    Case ButtonShape.Rectangle
                        gp = DrawRoundRect(0, 0, Me.Width, Me.Height, Me.UL, Me.UR, Me.LR, Me.LL)
                        gp2 = DrawRoundRect(0, 0, Me.Width - 1, Me.Height - 1, Me.UL, Me.UR, Me.LR, Me.LL)


                    Case ButtonShape.TriangleUp
                        pts = DrawTriangle(rec, "Up")
                        gp.AddPolygon(pts)
                        pts = DrawTriangle(rec2, "Up")
                        gp2.AddPolygon(pts)

                    Case ButtonShape.TriangleDown
                        pts = DrawTriangle(rec, "Down")
                        gp.AddPolygon(pts)
                        pts = DrawTriangle(rec2, "Down")
                        gp2.AddPolygon(pts)

                    Case ButtonShape.TriangleLeft
                        pts = DrawTriangle(rec, "Left")
                        gp.AddPolygon(pts)
                        pts = DrawTriangle(rec2, "Left")
                        gp2.AddPolygon(pts)

                    Case ButtonShape.TriangleRight
                        pts = DrawTriangle(rec, "Right")
                        gp.AddPolygon(pts)
                        pts = DrawTriangle(rec2, "Right")
                        gp2.AddPolygon(pts)

                End Select

                rec2 = New RectangleF(2, 2, Me.Width - 2, Me.Height - 2)
                If Not Me.Image Is Nothing Then
                    rec3 = New Rectangle(Me.ImageLocationX + Me.ImageSize.Width, 1, Me.Width - Me.ImageLocationX - Me.ImageSize.Width - 2, Me.Height - 2)
                    rec2 = New Rectangle(Me.ImageLocationX + Me.ImageSize.Width, 2, Me.Width - Me.ImageLocationX - Me.ImageSize.Width - 2, Me.Height - 2)
                End If

                Select Case s
                    Case GradientButtonStates.MouseOver
                        e.Graphics.FillPath(DimBlenedColor(ButtonGradientColor1, ButtonGradientColor2, ButtonGradientColor3, 75, mBrush, gp), gp)
                        e.Graphics.DrawPath(New Pen(DimTheColor(ButtonGradientColor3, -40)), gp2)
                        DrawText(e.Graphics, rec3, sf)
                        If Not Me.Image Is Nothing Then
                            e.Graphics.DrawImage(mImage, DrawImageRotatedAroundCenter(New Point((mImgX + mImageSize.Width) / 2, (mImgY + mImageSize.Height) / 2), mImage, -mAngle * Math.PI / 180))
                            AnimationTimer.Enabled = True
                        End If

                    Case GradientButtonStates.MouseDown
                        e.Graphics.FillPath(DimBlenedColor(ButtonGradientColor1, ButtonGradientColor2, ButtonGradientColor3, -15, mBrush, gp), gp)
                        e.Graphics.DrawPath(New Pen(DimTheColor(ButtonGradientColor1, 200), 3), gp2)
                        DrawText(e.Graphics, rec2, sf)
                        If Not Me.Image Is Nothing Then
                            e.Graphics.DrawImage(mImage, DrawImageRotatedAroundCenter(New Point((mImgX + mImageSize.Width + 2) / 2, (mImgY + mImageSize.Height + 2) / 2), mImage, -mAngle * Math.PI / 180))
                            AnimationTimer.Enabled = False
                        End If

                    Case GradientButtonStates.Normal
                        e.Graphics.FillPath(DimBlenedColor(ButtonGradientColor1, ButtonGradientColor2, ButtonGradientColor3, 0, mBrush, gp), gp)
                        e.Graphics.DrawPath(New Pen(DimTheColor(ButtonGradientColor3, 50)), gp2)
                        DrawText(e.Graphics, rec3, sf)
                        If Not Me.Image Is Nothing Then
                            e.Graphics.DrawImage(mImage, DrawImageRotatedAroundCenter(New Point((mImgX + mImageSize.Width) / 2, (mImgY + mImageSize.Height) / 2), mImage, -mAngle * Math.PI / 180))
                            AnimationTimer.Enabled = False
                        End If

                    Case GradientButtonStates.Focused
                        e.Graphics.FillPath(DimBlenedColor(ButtonGradientColor1, ButtonGradientColor2, ButtonGradientColor3, 0, mBrush, gp), gp)
                        e.Graphics.DrawPath(New Pen(DimTheColor(ButtonGradientColor1, 200), 3), gp2)
                        DrawText(e.Graphics, rec3, sf)
                        If Not Me.Image Is Nothing Then
                            e.Graphics.DrawImage(mImage, DrawImageRotatedAroundCenter(New Point((mImgX + mImageSize.Width) / 2, (mImgY + mImageSize.Height) / 2), mImage, -mAngle * Math.PI / 180))
                            AnimationTimer.Enabled = False
                        End If

                    Case GradientButtonStates.Disabled
                        e.Graphics.FillPath(New LinearGradientBrush(rec, Color.DarkGray, Color.LightGray, LinearGradientMode.Vertical), gp)
                        e.Graphics.DrawPath(New Pen(GrayTheColor(DimTheColor(ButtonGradientColor3, 50))), gp2)
                        DrawText(e.Graphics, rec3, sf)
                        If Not Me.Image Is Nothing Then
                            e.Graphics.DrawImage(EnableDisableImage(mImage), DrawImageRotatedAroundCenter(New Point((mImgX + mImageSize.Width) / 2, (mImgY + mImageSize.Height) / 2), mImage, -mAngle * Math.PI / 180))
                        End If

                End Select

                gp.Dispose()
                gp2.Dispose()

            End Sub

            Private Function DrawImageRotatedAroundCenter(ByVal center As Point, ByVal img As Image, _
                                                     ByVal angle As Double)

                ' Think of the image as a rectangle that needs to be drawn rotated.
                ' Rotate the coordinates of the rectangle's corners.
                Dim lowerLeft As PointF = RotatePoint(New PointF(-(Me.ImageSize.Width / 2), (Me.ImageSize.Height / 2)), angle)
                Dim upperRight As PointF = RotatePoint(New PointF((Me.ImageSize.Width / 2), -(Me.ImageSize.Height / 2)), angle)
                Dim upperLeft As PointF = RotatePoint(New PointF(-(Me.ImageSize.Width / 2), -(Me.ImageSize.Height / 2)), angle)

                ' Create the points array by offsetting the coordinates with the center.
                Dim points() As PointF = {upperLeft + center, upperRight + center, lowerLeft + center}
                Return points

            End Function

            Private Function RotatePoint(ByVal p As PointF, ByVal angle As Double) As PointF

                Dim x As Integer = p.X * Math.Cos(angle) + p.Y * Math.Sin(angle)
                Dim y As Integer = -p.X * Math.Sin(angle) + p.Y * Math.Cos(angle)

                Return New PointF(x, y)

            End Function

            Private Function EnableDisableImage(ByVal img As Image) As Bitmap

                If Me.Enabled Then Return CType(img, Bitmap)
                Dim bm As Bitmap = New Bitmap(img.Width, img.Height)
                Dim g As Graphics = Graphics.FromImage(bm)
                Dim cm As ColorMatrix = New ColorMatrix(New Single()() _
                     {New Single() {0.5, 0.5, 0.5, 0, 0}, _
                    New Single() {0.5, 0.5, 0.5, 0, 0}, _
                    New Single() {0.5, 0.5, 0.5, 0, 0}, _
                    New Single() {0, 0, 0, 1, 0}, _
                    New Single() {0, 0, 0, 0, 1}})

                Dim ia As ImageAttributes = New ImageAttributes()
                ia.SetColorMatrix(cm)
                g.DrawImage(img, New Rectangle(0, 0, img.Width, img.Height), 0, 0, img.Width, img.Height, GraphicsUnit.Pixel, ia)
                g.Dispose()
                Return bm

            End Function

            Function GrayTheColor(ByVal GrayColor As Color) As Color
                Dim gray As Integer = CInt(GrayColor.R * 0.3 + GrayColor.G * 0.59 + GrayColor.B * 0.11)
                Return Color.FromArgb(GrayColor.A, gray, gray, gray)
            End Function

            Public Function DrawRoundRect(ByVal x As Single, ByVal y As Single, ByVal width As Single, ByVal height As Single, ByVal ULradius As Single _
                                          , ByVal URradius As Single, ByVal LRradius As Single, ByVal LLradius As Single)

                Dim gp As GraphicsPath = New GraphicsPath()
                If URradius > Me.Height Then URradius = Me.Height
                If LRradius > Me.Height Then LRradius = Me.Height
                If LLradius > Me.Height Then LLradius = Me.Height
                If ULradius > Me.Height Then ULradius = Me.Height


                ' top right arc
                gp.AddArc(x + width - (URradius * 2), y, URradius * 2, URradius * 2, 270, 90)
                ' bottom right arc
                gp.AddArc(x + width - (LRradius * 2), y + height - (LRradius * 2), LRradius * 2, LRradius * 2, 0, 90)
                ' bottom left arc
                gp.AddArc(x, y + height - (LLradius * 2), LLradius * 2, LLradius * 2, 90, 90)
                ' top left arc
                gp.AddArc(x, y, ULradius * 2, ULradius * 2, 180, 90)

                gp.CloseFigure()
                Return gp
                gp.Dispose()

            End Function
            Public Function DrawTriangle(ByVal rec As RectangleF, ByVal str As String)

                Dim pts() As PointF = New PointF() {}
                Select Case str
                    Case "Up"
                        pts = New PointF() { _
                            New PointF(CSng(rec.Width / 2), rec.Y), _
                            New PointF(rec.Width, rec.Y + rec.Height), _
                            New PointF(rec.X, rec.Y + rec.Height)}
                    Case "Down"
                        pts = New PointF() { _
                            New PointF(rec.X, rec.Y), _
                            New PointF(CSng(rec.Width / 2), rec.Y + rec.Height), _
                            New PointF(rec.X + rec.Width, rec.Y)}
                    Case "Left"
                        pts = New PointF() { _
                            New PointF(rec.X, CSng(rec.Y + (rec.Height / 2))), _
                            New PointF(rec.Width, rec.Y), _
                            New PointF(rec.Width, rec.Y + rec.Height)}
                    Case "Right"
                        pts = New PointF() { _
                            New PointF(rec.X, rec.Y), _
                            New PointF(rec.Width, CSng(rec.Y + (rec.Height / 2))), _
                            New PointF(rec.X, rec.Y + rec.Height)}
                End Select

                Return pts

            End Function


            Protected Overrides Sub OnMouseMove(ByVal e As System.Windows.Forms.MouseEventArgs)
                MyBase.OnMouseMove(e)

                If e.Button = MouseButtons.None Or e.Button = MouseButtons.Left Then
                    If e.Button = MouseButtons.None Then
                        isMouseDown = False
                    End If
                    isMouseOver = True
                Else
                    If Not New Rectangle(0, 0, Me.Width, Me.Height).Contains(e.X, e.Y) Then
                        isMouseOver = False
                    Else
                        isMouseOver = True
                    End If
                End If

                If e.Button = MouseButtons.Left Then
                    If Not New Rectangle(0, 0, Me.Width, Me.Height).Contains(e.X, e.Y) Then
                        isMouseDown = False
                    Else
                        isMouseDown = True
                    End If
                End If
                Me.Invalidate()
            End Sub

            Protected Overrides Sub OnMouseLeave(ByVal e As System.EventArgs)
                MyBase.OnMouseLeave(e)

                isMouseOver = False
                isMouseDown = False

                Me.Invalidate()
            End Sub

            Protected Overrides Sub OnMouseDown(ByVal e As System.Windows.Forms.MouseEventArgs)
                MyBase.OnMouseDown(e)

                If e.Button = MouseButtons.Left Then
                    If New Rectangle(0, 0, Me.Width, Me.Height).Contains(e.X, e.Y) Then
                        isMouseDown = True
                    Else
                        isMouseDown = False
                    End If

                    Me.Focus()
                End If

                Me.Invalidate()
            End Sub

            Protected Overrides Sub OnMouseUp(ByVal e As System.Windows.Forms.MouseEventArgs)

                MyBase.OnMouseUp(e)

                If e.Button = MouseButtons.Left Then
                    isMouseDown = False
                End If

                Me.Invalidate()
            End Sub

            Protected Overrides Sub OnEnter(ByVal e As System.EventArgs)
                MyBase.OnEnter(e)

                isFocused = True

                Me.Invalidate()
            End Sub

            Protected Overrides Sub OnLeave(ByVal e As System.EventArgs)
                MyBase.OnLeave(e)

                isFocused = False

                Me.Invalidate()
            End Sub

            Protected Overrides Sub OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs)
                MyBase.OnKeyDown(e)

                If e.KeyCode = Keys.Space Then
                    isMouseDown = True
                End If

                Me.Invalidate()
            End Sub

            Protected Overrides Sub OnKeyUp(ByVal e As System.Windows.Forms.KeyEventArgs)
                MyBase.OnKeyUp(e)

                If e.KeyCode = 32 Then
                    isMouseDown = False
                    MyBase.OnClick(e)
                End If

                Me.Invalidate()
            End Sub

            Protected Overrides Sub OnTextChanged(ByVal e As System.EventArgs)
                MyBase.OnTextChanged(e)
                Me.Invalidate()
            End Sub

            Protected Overrides Sub OnEnabledChanged(ByVal e As System.EventArgs)
                MyBase.OnEnabledChanged(e)
                Me.Invalidate()
            End Sub
            Function DimTheColor(ByVal DimColor As Color, ByVal DimDegree As Integer) As Color
                If DimColor = Color.Transparent Or DimDegree = 0 Then Return DimColor
                Dim ColorR As Integer = DimColor.R + DimDegree
                Dim ColorG As Integer = DimColor.G + DimDegree
                Dim ColorB As Integer = DimColor.B + DimDegree

                If ColorR > 255 Then ColorR = 255
                If ColorG > 255 Then ColorG = 255
                If ColorB > 255 Then ColorB = 255
                If ColorR < 0 Then ColorR = 0
                If ColorG < 0 Then ColorG = 0
                If ColorB < 0 Then ColorB = 0

                Return Color.FromArgb(ColorR, ColorG, ColorB)

            End Function

            Function DimBlenedColor(ByVal c1 As Color, ByVal c2 As Color, ByVal c3 As Color, ByVal DimDegree As Integer, ByVal Brush As BrushType, ByVal gp As GraphicsPath) As Brush

                Dim LGM As LinearGradientMode
                Select Case mGDirection
                    Case GradientDirection.BackwardDiagonal
                        LGM = LinearGradientMode.BackwardDiagonal
                    Case GradientDirection.ForwardDiagonal
                        LGM = LinearGradientMode.ForwardDiagonal
                    Case GradientDirection.Horizontal
                        LGM = LinearGradientMode.Horizontal
                    Case GradientDirection.Vertical
                        LGM = LinearGradientMode.Vertical
                End Select

                Select Case mBrush
                    Case BrushType.Linear
                        Dim br As LinearGradientBrush = New LinearGradientBrush(New Rectangle(0, 0, Me.Width, Me.Height), _
                        Color.White, Color.Black, LGM)
                        Dim cb As New ColorBlend
                        cb.Colors = New Color() {DimTheColor(c1, DimDegree), DimTheColor(c2, DimDegree), DimTheColor(c3, DimDegree)}
                        cb.Positions = New Single() {0, mPos, 1}
                        br.InterpolationColors = cb
                        Return br

                    Case BrushType.Path
                        Dim br As PathGradientBrush = New PathGradientBrush(gp)
                        Dim cb As New ColorBlend()
                        cb.Colors = New Color() {DimTheColor(c1, DimDegree), DimTheColor(c2, DimDegree), DimTheColor(c3, DimDegree)}
                        cb.Positions = New Single() {0, mPos, 1}
                        br.FocusScales = New PointF(mFPx, mFPy)
                        br.CenterPoint = New PointF(mCPx, mCPy)
                        br.InterpolationColors = cb
                        Return br

                    Case BrushType.Solid
                        Dim br As SolidBrush = New SolidBrush(c2)
                        Return br
                End Select
                Return Brushes.Aqua
            End Function
            Private Sub DrawText(ByVal g As Graphics, ByVal rec As RectangleF, ByVal StrFor As StringFormat)

                Select Case mStyle
                    Case Text_Style.Regular
                        If s = GradientButtonStates.Disabled Then
                            g.DrawString(Me.Text, Me.Font, Brushes.LightGray, rec, StrFor)
                            rec.X = rec.X + 2
                            rec.Y = rec.Y + 2
                            g.DrawString(Me.Text, Me.Font, Brushes.Gray, rec, StrFor)
                        Else
                            g.DrawString(Me.Text, Me.Font, New SolidBrush(Me.ForeColor), rec, StrFor)
                        End If

                    Case Text_Style.Engrave
                        If s = GradientButtonStates.Disabled Then
                            g.DrawString(Me.Text, Me.Font, Brushes.LightGray, rec, StrFor)
                            g.DrawString(Me.Text, Me.Font, Brushes.Gray, New RectangleF _
                            (rec.X - 1, rec.Y - 1, rec.Width, rec.Height), StrFor)
                        Else
                            g.DrawString(Me.Text, Me.Font, Brushes.SeaShell, rec, StrFor)
                            g.DrawString(Me.Text, Me.Font, New SolidBrush(Me.ForeColor), New RectangleF _
                            (rec.X - 1, rec.Y - 1, rec.Width, rec.Height), StrFor)
                        End If

                    Case Text_Style.Embosse
                        If s = GradientButtonStates.Disabled Then
                            g.DrawString(Me.Text, Me.Font, Brushes.Gray, New RectangleF _
                            (rec.X + 1, rec.Y + 1, rec.Width + 1, rec.Height + 1), StrFor)
                            g.DrawString(Me.Text, Me.Font, Brushes.LightGray, rec, StrFor)
                        Else
                            g.DrawString(Me.Text, Me.Font, Brushes.Black, New RectangleF _
                            (rec.X + 1, rec.Y + 1, rec.Width + 1, rec.Height + 1), StrFor)
                            g.DrawString(Me.Text, Me.Font, Brushes.SeaShell, rec, StrFor)
                        End If

                    Case Text_Style.Block
                        If s = GradientButtonStates.Disabled Then
                            Dim I As Short
                            For I = 5 To 0 Step -1
                                g.DrawString(Me.Text, Me.Font, Brushes.Gray, _
                                New RectangleF(rec.X - I, rec.Y - I, rec.Width, rec.Height), StrFor)
                            Next
                            g.DrawString(Me.Text, Me.Font, Brushes.LightGray, rec, StrFor)
                        Else
                            Dim I As Short
                            For I = 5 To 0 Step -1
                                g.DrawString(Me.Text, Me.Font, Brushes.Black, _
                                New RectangleF(rec.X - I, rec.Y - I, rec.Width, rec.Height), StrFor)
                            Next
                            g.DrawString(Me.Text, Me.Font, Brushes.SeaShell, rec, StrFor)
                        End If

                    Case Text_Style.Gradinet
                        If s = GradientButtonStates.Disabled Then
                            g.DrawString(Me.Text, Me.Font, Brushes.LightGray, rec, StrFor)
                            rec.X = rec.X + 2
                            rec.Y = rec.Y + 2
                            g.DrawString(Me.Text, Me.Font, New LinearGradientBrush _
                            (rec, Color.Gray, Color.LightGray, LinearGradientMode.Horizontal), rec, StrFor)
                        Else
                            g.DrawString(Me.Text, Me.Font, New LinearGradientBrush _
                            (rec, Color.Blue, Color.Yellow, LinearGradientMode.Horizontal), rec, StrFor)
                        End If

                    Case Text_Style.Hatch
                        If s = GradientButtonStates.Disabled Then
                            g.DrawString(Me.Text, Me.Font, Brushes.Gray, rec, StrFor)
                            rec.X = rec.X + 2
                            rec.Y = rec.Y + 2
                            g.DrawString(Me.Text, Me.Font, New HatchBrush _
                            (HatchStyle.DiagonalBrick, Color.Gray, Color.LightGray), rec, StrFor)
                        Else
                            g.DrawString(Me.Text, Me.Font, New HatchBrush _
                            (HatchStyle.HorizontalBrick, Color.White, Color.Brown), rec, StrFor)
                        End If

                    Case Text_Style.Reflect
                        Dim mState As GraphicsState
                        Dim textHeight As Single
                        Dim lineAscent As Integer
                        Dim lineSpacing As Integer
                        Dim lineHeight As Single
                        Dim txtSize As SizeF
                        Dim mirrorMatrix As Drawing2D.Matrix

                        txtSize = g.MeasureString(Me.Text, Me.Font)
                        lineAscent = Me.Font.FontFamily.GetCellAscent(Me.Font.Style)
                        lineSpacing = Me.Font.FontFamily.GetLineSpacing(Me.Font.Style)
                        lineHeight = Me.Font.GetHeight(g)
                        textHeight = lineHeight * lineAscent / lineSpacing
                        mState = g.Save()
                        mirrorMatrix = New Drawing2D.Matrix(1, 0, 0, -1, 0, (rec.Height + textHeight / 2))
                        g.Transform = mirrorMatrix
                        If s = GradientButtonStates.Disabled Then
                            g.DrawString(Me.Text, Me.Font, Brushes.LightGray, rec, StrFor)
                            g.Restore(mState)
                            mirrorMatrix = New Drawing2D.Matrix(1, 0, 0, 1, 0, -7)
                            g.Transform = mirrorMatrix
                            g.DrawString(Me.Text, Me.Font, Brushes.Gray, rec, StrFor)
                            g.Restore(mState)
                        Else
                            g.DrawString(Me.Text, Me.Font, Brushes.DarkGray, rec, StrFor)
                            g.Restore(mState)
                            mirrorMatrix = New Drawing2D.Matrix(1, 0, 0, 1, 0, -7)
                            g.Transform = mirrorMatrix
                            g.DrawString(Me.Text, Me.Font, Brushes.Black, rec, StrFor)
                            g.Restore(mState)
                        End If

                    Case Text_Style.Shadow
                        If s = GradientButtonStates.Disabled Then
                            g.DrawString(Me.Text, Me.Font, Brushes.Gray, rec, StrFor)
                            g.DrawString(Me.Text, Me.Font, Brushes.LightGray, New RectangleF _
                            (rec.X - 2, rec.Y - 2, rec.Width - 2, rec.Height - 2), StrFor)
                        Else
                            g.DrawString(Me.Text, Me.Font, Brushes.Black, rec, StrFor)
                            g.DrawString(Me.Text, Me.Font, Brushes.SeaShell, New RectangleF _
                            (rec.X - 2, rec.Y - 2, rec.Width - 2, rec.Height - 2), StrFor)
                        End If

                    Case Text_Style.Shear
                        Dim mTrans As Matrix
                        mTrans = g.Transform
                        mTrans.Shear(1, 0)
                        g.Transform = mTrans
                        If s = GradientButtonStates.Disabled Then
                            g.DrawString(Me.Text, Me.Font, Brushes.LightGray, rec, StrFor)
                            rec.X = rec.X + 2
                            rec.Y = rec.Y + 2
                            g.DrawString(Me.Text, Me.Font, Brushes.Gray, rec, StrFor)
                        Else
                            g.DrawString(Me.Text, Me.Font, Brushes.Black, rec, StrFor)
                        End If

                    Case Text_Style.Chisel
                        If s = GradientButtonStates.Disabled Then
                            g.DrawString(Me.Text, Me.Font, Brushes.LightGray, rec, StrFor)
                            rec.X = rec.X + 2
                            rec.Y = rec.Y + 2
                            g.DrawString(Me.Text, Me.Font, Brushes.Gray, rec, StrFor)
                        Else
                            g.DrawString(Me.Text, Me.Font, Brushes.Black, rec, StrFor)
                            rec.X = rec.X + 4
                            rec.Y = rec.Y + 4
                            g.DrawString(Me.Text, Me.Font, Brushes.LightGray, rec, StrFor)
                            rec.X = rec.X - 2
                            rec.Y = rec.Y - 2
                            g.DrawString(Me.Text, Me.Font, Brushes.Gray, rec, StrFor)
                        End If

                End Select


            End Sub

            Private Sub AnimationTimer_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles AnimationTimer.Tick
                If mAnimated = True Then
                    Dim I As Integer
                    For I = 0 To 359
                        Application.DoEvents()
                        mAngle = I
                        Me.Invalidate()
                    Next
                End If
            End Sub
        End Class


        Public Class SinWave
            Private Pic As PictureBox
            Private c As New EAMS.Graphical.Curve
            Private t As New EAMS.Diagonstic.Synchrounce
            Private _WaveColor As Color = Color.Red
            Private p(4) As PointF
            Private _u As Integer = 1
            Private _WaveLenth As Integer = 200
            Private _X As Integer = 0
            Private _Xi As Integer = 15
            Private _WaveType As e_WaveType = e_WaveType.LoopBack
            Private _XL As Integer = 0
            Private ImgP As New EAMS.Graphical.ImageProcessing


#Region "Types"
            Enum e_WaveType
                Continuously = 1
                LoopBack = 2
            End Enum

#End Region


#Region "Intialization"
            Public Sub New()
                AddHandler t.VirtualTimer, AddressOf DrawSinWave
                t.IntervalType = Diagonstic.Synchrounce.IntType.Seconds
                t.Interval = 1
            End Sub
#End Region

#Region "Handle Timer"
            Private Sub DrawSinWave(ByVal timming As Integer)
                Dim WaveCount As Integer = 0
                c.PenColor = _WaveColor
                For WaveCount = 1 To _u
                    p(0).X = _X
                    p(0).Y = 0
                    _X += _Xi
                    p(1).X = _X
                    p(1).Y = 200
                    _X += _Xi
                    p(2).X = _X
                    p(2).Y = 0
                    _X += _Xi
                    p(3).X = _X + _Xi
                    p(3).Y = -200
                    _X += _Xi
                    p(4).X = _X
                    p(4).Y = 0
                    c.Curve(p, False)
                    Application.DoEvents()
                    '-----\
                    Dim lbl As New Label
                    ImgP.SetText(Pic, timming, lbl.Font, Brushes.Red, _X, c.CenterPoint.Y)
                    If _WaveType = e_WaveType.LoopBack Then
                        _XL = Pic.Width / 2
                        If _X >= _XL Then
                            _X = 0
                            Pic.Refresh()
                            c.PenColor = Color.Blue
                            c.CoordinatorsInialization()
                        End If
                    End If
                    '-----
                Next
            End Sub
#End Region

#Region "Properties"
            Public Property WaveType() As e_WaveType
                Get
                    Return _WaveType
                End Get
                Set(ByVal value As e_WaveType)
                    _WaveType = value
                End Set
            End Property
            Public Property µ() As Integer
                Get
                    Return _u
                End Get
                Set(ByVal value As Integer)
                    _u = value
                End Set
            End Property
            Public Property WaveLenth() As Integer
                Get
                    Return _WaveLenth
                End Get
                Set(ByVal value As Integer)
                    _WaveLenth = value
                End Set
            End Property
            Public WriteOnly Property PictureBox() As PictureBox
                Set(ByVal value As PictureBox)
                    Pic = New PictureBox
                    Pic = value
                    c.PictureBox = value
                    c.CoordinatorsInialization()
                    _X = 0
                End Set
            End Property
            Public Property WaveColor() As Color
                Get
                    Return _WaveColor
                End Get
                Set(ByVal value As Color)
                    _WaveColor = value
                End Set
            End Property
#End Region
#Region "Methods"
            Public Sub StartSinWave()
                t.ActiveTimer()
            End Sub
            Public Sub StopSinWave()
                t.DiactiveTimer()
            End Sub
            Public Sub PauseSinWave()
                t.PauseTimer()
            End Sub
#End Region
#Region "Private Methods"
            Private Sub PlotText(ByVal x As Integer, ByVal y As Integer, ByVal Text As String, ByRef frm As Form)
                ImgP.SetText(Pic, Text, frm.Font, Brushes.Yellow, x, y)
            End Sub
#End Region
        End Class

        Public Class LinearChart
            Inherits Curve
            Private colPoints As New Collection
            Private _Font As Font
            Private imgP As New EAMS.Graphical.ImageProcessing
            Private _TextPointOn As Boolean = True

#Region "Enum"
            Enum en_DrawType
                Curve = 1
                Linear = 2
            End Enum
#End Region
#Region "Property"
            Public WriteOnly Property TextPointOn() As Boolean
                Set(ByVal value As Boolean)
                    _TextPointOn = value
                End Set
            End Property

            Public WriteOnly Property Font() As Font
                Set(ByVal value As Font)
                    _Font = value
                End Set
            End Property
#End Region
#Region "Methods"
            Public Sub AddPoint(ByVal p As PointF)
                colPoints.Add(p)
            End Sub
            Public Sub DrawPoint(ByVal DrawType As en_DrawType)
                Graphical.Curve.Pic.Refresh()
                Me.CoordinatorsInialization()
                imgP.SetText(Pic, "0,0", _Font, Brushes.White, Me.CenterPoint.X, Me.CenterPoint.Y)
                If colPoints.Count = 1 Then Exit Sub
                Dim x As Integer = 0
                Dim points(colPoints.Count - 1) As PointF
                For x = 1 To colPoints.Count
                    points(x - 1) = colPoints.Item(x)
                Next

                Select Case DrawType
                    Case en_DrawType.Curve
                        Me.Curve(points, False)
                        If _TextPointOn Then
                            For x = 0 To points.GetUpperBound(0)
                                imgP.SetText(Pic, points(x).X & "," & points(x).Y, _Font, Brushes.White, points(x).X + Me.CenterPoint.X, Me.CenterPoint.Y - points(x).Y)
                            Next
                        End If
                    Case en_DrawType.Linear
                        For x = 1 To points.GetUpperBound(0)
                            Me.Line(points(x - 1).X, points(x - 1).Y, points(x).X, points(x).Y)
                        Next
                        If _TextPointOn Then
                            For x = 0 To points.GetUpperBound(0)
                                imgP.SetText(Pic, points(x).X & "," & points(x).Y, _Font, Brushes.White, points(x).X + Me.CenterPoint.X, Me.CenterPoint.Y - points(x).Y)
                            Next
                        End If
                End Select
            End Sub
#End Region
        End Class

        Namespace FSButton

#Region " fsButton Control "



            Public Class fsButton
                Inherits System.Windows.Forms.Control

                Event ButtonThemeChanged(ByVal sender As Object, ByVal e As EventArgs)
                Event ThemeColorChanged(ByVal sender As Object, ByVal e As EventArgs)
                Event ShowTextEffectsChanged(ByVal sender As Object, ByVal e As EventArgs)

#Region " Windows Form Designer generated code "

                Public Sub New()
                    MyBase.New()

                    'This call is required by the Windows Form Designer.
                    InitializeComponent()

                    'Add any initialization after the InitializeComponent() call
                    MyBase.SetStyle(ControlStyles.AllPaintingInWmPaint, True)
                    MyBase.SetStyle(ControlStyles.UserPaint, True)
                    MyBase.SetStyle(ControlStyles.DoubleBuffer, True)
                    MyBase.SetStyle(ControlStyles.ResizeRedraw, True)
                    MyBase.SetStyle(ControlStyles.SupportsTransparentBackColor, True)
                End Sub

                'UserControl1 overrides dispose to clean up the component list.
                Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
                    If disposing Then
                        If Not (components Is Nothing) Then
                            components.Dispose()
                        End If
                    End If
                    MyBase.Dispose(disposing)
                End Sub

                'Required by the Windows Form Designer
                Private components As System.ComponentModel.IContainer

                'NOTE: The following procedure is required by the Windows Form Designer
                'It can be modified using the Windows Form Designer.  
                'Do not modify it using the code editor.
                <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
                    components = New System.ComponentModel.Container
                End Sub

#End Region

                Private isMouseOver As Boolean = False
                Private isMouseDown As Boolean = False
                Private isFocused As Boolean = False

                Private _Theme As Themes
                Private _ThemeColor As Color = cRGB(0, 102, 255)
                Private _ShowTextEffects As Boolean = True

                Public Enum Themes
                    LiquidChromeXP
                    SoftGlassXP
                    WindowsXP
                    MSNLoginButton
                    Aqua
                    Hover3D
                    OfficeXP
                    Office2003
                    Macintosh
                End Enum

                <Description("Determines whether or not to show the text effects for the selected theme."), DefaultValue(GetType(Boolean), "True")> _
                Property ShowTextEffects() As Boolean
                    Get
                        Return _ShowTextEffects
                    End Get
                    Set(ByVal Value As Boolean)
                        _ShowTextEffects = Value
                        Me.Invalidate()
                        RaiseEvent ShowTextEffectsChanged(Me, New EventArgs)
                    End Set
                End Property

                <Description("Sets the base color for the theme."), DefaultValue(GetType(Color), "0, 102, 255")> _
                Property ThemeColor() As Color
                    Get
                        Return _ThemeColor
                    End Get
                    Set(ByVal Value As Color)
                        _ThemeColor = Value
                        Me.Invalidate()
                        RaiseEvent ThemeColorChanged(Me, New EventArgs)
                    End Set
                End Property

                <Description("Controls the theme which is applied to the button."), DefaultValue(GetType(Themes), "LiquidChromeXP")> _
                Property ButtonTheme() As Themes
                    Get
                        Return _Theme
                    End Get
                    Set(ByVal Value As Themes)
                        _Theme = Value
                        Me.Invalidate()
                        RaiseEvent ButtonThemeChanged(Me, New EventArgs)
                    End Set
                End Property

                Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
                    MyBase.OnPaint(e)

                    'Add all painting code here

                    Dim s As Theme.States

                    If Me.Enabled = False Then
                        s = Theme.States.Disabled
                    Else
                        If isMouseDown = False Then
                            If isMouseOver = True Then
                                s = Theme.States.MouseOver
                            Else
                                If isFocused = True Then
                                    s = Theme.States.Focused
                                Else
                                    s = Theme.States.Normal
                                End If
                            End If
                        Else
                            s = Theme.States.MouseDown
                        End If
                    End If
                    Select Case _Theme
                        Case Themes.LiquidChromeXP
                            Dim t As New Theme.LiquidChromeXP(New Rectangle(0, 0, Me.Width, Me.Height))
                            t.DrawTheme(e, s, _ShowTextEffects)
                            t.DrawText(e, s, Me.Text, Me.Font, Me.ForeColor)
                        Case Themes.SoftGlassXP
                            Dim t As New Theme.SoftGlassXP(New Rectangle(0, 0, Me.Width, Me.Height))
                            t.DrawTheme(e, s, _ThemeColor, _ShowTextEffects)
                            t.DrawText(e, s, Me.Text, Me.Font, Me.ForeColor)
                        Case Themes.WindowsXP
                            Dim t As New Theme.WindowsXP(New Rectangle(0, 0, Me.Width, Me.Height))
                            t.DrawTheme(e, s, _ShowTextEffects)
                            t.DrawText(e, s, Me.Text, Me.Font, Me.ForeColor)
                        Case Themes.MSNLoginButton
                            Dim t As New Theme.MSNLoginButton(New Rectangle(0, 0, Me.Width, Me.Height))
                            t.DrawTheme(e, s, _ShowTextEffects, Me.BackColor)
                            t.DrawText(e, s, Me.Text, Me.Font)
                        Case Themes.Aqua
                            Dim t As New Theme.Aqua(New Rectangle(0, 0, Me.Width, Me.Height))
                            t.DrawTheme(e, s, _ShowTextEffects, _ThemeColor)
                            t.DrawText(e, s, Me.Text, Me.Font, Me.ForeColor)
                        Case Themes.Hover3D
                            Dim t As New Theme.Hover3D(New Rectangle(0, 0, Me.Width, Me.Height))
                            t.DrawTheme(e, s, _ShowTextEffects, isFocused)
                            t.DrawText(e, s, Me.Text, Me.Font, Me.ForeColor)
                        Case Themes.OfficeXP
                            Dim t As New Theme.OfficeXP(New Rectangle(0, 0, Me.Width, Me.Height))
                            t.DrawTheme(e, s, _ShowTextEffects, _ThemeColor, isFocused)
                            t.DrawText(e, s, Me.Text, Me.Font, Me.ForeColor)
                        Case Themes.Office2003
                            Dim t As New Theme.Office2003(New Rectangle(0, 0, Me.Width, Me.Height))
                            t.DrawTheme(e, s, _ShowTextEffects, _ThemeColor, isFocused)
                            t.DrawText(e, s, Me.Text, Me.Font, Me.ForeColor)
                        Case Themes.Macintosh
                            Dim t As New Theme.Macintosh(New Rectangle(0, 0, Me.Width, Me.Height))
                            t.DrawTheme(e, s, _ShowTextEffects, _ThemeColor)
                            t.DrawText(e, s, Me.Text, Me.Font, Me.ForeColor)
                    End Select

                End Sub


                Protected Overrides Sub OnMouseMove(ByVal e As System.Windows.Forms.MouseEventArgs)
                    MyBase.OnMouseMove(e)

                    If e.Button = MouseButtons.None Or e.Button = MouseButtons.Left Then
                        If e.Button = MouseButtons.None Then
                            isMouseDown = False
                        End If
                        isMouseOver = True
                    Else
                        If Not New Rectangle(0, 0, Me.Width, Me.Height).Contains(e.X, e.Y) Then
                            isMouseOver = False
                        Else
                            isMouseOver = True
                        End If
                    End If

                    If e.Button = MouseButtons.Left Then
                        If Not New Rectangle(0, 0, Me.Width, Me.Height).Contains(e.X, e.Y) Then
                            isMouseDown = False
                        Else
                            isMouseDown = True
                        End If
                    End If
                    Me.Invalidate()
                End Sub

                Protected Overrides Sub OnMouseLeave(ByVal e As System.EventArgs)
                    MyBase.OnMouseLeave(e)

                    isMouseOver = False
                    isMouseDown = False

                    Me.Invalidate()
                End Sub

                Protected Overrides Sub OnMouseDown(ByVal e As System.Windows.Forms.MouseEventArgs)
                    MyBase.OnMouseDown(e)

                    If e.Button = MouseButtons.Left Then
                        If New Rectangle(0, 0, Me.Width, Me.Height).Contains(e.X, e.Y) Then
                            isMouseDown = True
                        Else
                            isMouseDown = False
                        End If

                        Me.Focus()
                    End If

                    Me.Invalidate()
                End Sub

                Protected Overrides Sub OnMouseUp(ByVal e As System.Windows.Forms.MouseEventArgs)

                    MyBase.OnMouseUp(e)

                    If e.Button = MouseButtons.Left Then
                        isMouseDown = False
                    End If

                    Me.Invalidate()
                End Sub

                Protected Overrides Sub OnEnter(ByVal e As System.EventArgs)
                    MyBase.OnEnter(e)

                    isFocused = True

                    Me.Invalidate()
                End Sub

                Protected Overrides Sub OnLeave(ByVal e As System.EventArgs)
                    MyBase.OnLeave(e)

                    isFocused = False

                    Me.Invalidate()
                End Sub

                Protected Overrides Sub OnKeyDown(ByVal e As System.Windows.Forms.KeyEventArgs)
                    MyBase.OnKeyDown(e)

                    If e.KeyCode = Keys.Space Then
                        isMouseDown = True
                    End If

                    Me.Invalidate()
                End Sub

                Protected Overrides Sub OnKeyUp(ByVal e As System.Windows.Forms.KeyEventArgs)
                    MyBase.OnKeyUp(e)

                    If e.KeyCode = 32 Then
                        isMouseDown = False
                        MyBase.OnClick(e)
                    End If

                    Me.Invalidate()
                End Sub

                Protected Overrides Sub OnTextChanged(ByVal e As System.EventArgs)
                    MyBase.OnTextChanged(e)
                    Me.Invalidate()
                End Sub

                Protected Overrides Sub OnEnabledChanged(ByVal e As System.EventArgs)
                    MyBase.OnEnabledChanged(e)
                    Me.Invalidate()
                End Sub
            End Class

#End Region

#Region " Theme "





            Namespace Theme

                Public Enum States
                    Normal
                    Focused
                    MouseOver
                    MouseDown
                    Disabled
                End Enum

#Region " Liquid Chrome XP (Default) Theme "
                Public Class LiquidChromeXP

                    Private My As Rectangle
                    Private TextEffects As Boolean

                    Public Sub New(ByVal Owner As Rectangle)
                        MyBase.New()
                        My = Owner
                    End Sub

                    Private Sub FillRoundedRectangle(ByVal b As Brush, ByVal Radius As Integer, ByVal Rect As Rectangle, ByVal e As PaintEventArgs)
                        Dim TL, TR, BL, BR As Point 'The four corners of the rectabgle

                        'Set the values of each corner point
                        TL = New Point(Rect.Left, Rect.Top)
                        TR = New Point(Rect.Left + Rect.Width, Rect.Top)
                        BL = New Point(Rect.Left, Rect.Top + Rect.Height)
                        BR = New Point(Rect.Left + Rect.Width, Rect.Top + Rect.Height)

                        'Draws the four corner circles
                        e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias 'Changes smoothing mode to anti-alias to remove jagged edges
                        e.Graphics.FillEllipse(b, New Rectangle(TL.X, TL.Y, Radius * 2, Radius * 2)) 'Top-left circle
                        e.Graphics.FillEllipse(b, New Rectangle(BL.X, BL.Y - (Radius * 2) - 1, Radius * 2, Radius * 2)) 'Bottom-left circle
                        e.Graphics.FillEllipse(b, New Rectangle(TR.X - (Radius * 2) - 1, TR.Y, Radius * 2, Radius * 2)) 'Top-right circle
                        e.Graphics.FillEllipse(b, New Rectangle(BR.X - (Radius * 2) - 1, BR.Y - (Radius * 2) - 1, Radius * 2, Radius * 2)) 'Bottom-right circle

                        'Draws the two blocks
                        e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.Default 'Returns the smoothing mode to default for a crisp structure
                        e.Graphics.FillRectangle(b, New Rectangle(TL.X, TL.Y + Radius, Rect.Width, Rect.Height - (Radius * 2)))
                        e.Graphics.FillRectangle(b, New Rectangle(TL.X + Radius, TL.Y, Rect.Width - (Radius * 2), Rect.Height))
                    End Sub

                    Private Sub DrawBase(ByVal e As PaintEventArgs, ByVal state As States)
                        Dim b As Brush
                        If state = States.Disabled Then
                            b = New SolidBrush(cRGB(201, 199, 186))
                        Else
                            b = New SolidBrush(cRGB(35, 55, 85))
                        End If
                        FillRoundedRectangle(b, 4, New Rectangle(0, 0, My.Width, My.Height), e)
                    End Sub

                    Private Sub DrawDisabledButtonBody(ByVal e As PaintEventArgs)
                        Dim b As Brush
                        b = New SolidBrush(cRGB(245, 244, 234))
                        FillRoundedRectangle(b, 3, New Rectangle(My.Left + 1, My.Top + 1, My.Width - 2, My.Height - 2), e)
                    End Sub


                    Private Sub DrawFocusButtonBody(ByVal e As PaintEventArgs)
                        Dim c1, c2, c3 As Color
                        Dim b As SolidBrush
                        Dim lgb As System.Drawing.Drawing2D.LinearGradientBrush
                        Dim cb As New System.Drawing.Drawing2D.ColorBlend

                        b = New SolidBrush(cRGB(153, 214, 255))
                        FillRoundedRectangle(b, 3, New Rectangle(1, 1, My.Width - 2, My.Height - 2), e)

                        b = New SolidBrush(cRGB(0, 122, 204))
                        FillRoundedRectangle(b, 3, New Rectangle(2, 2, My.Width - 3, My.Height - 3), e)

                        b = New SolidBrush(cRGB(0, 153, 255))
                        FillRoundedRectangle(b, 2, New Rectangle(2, 2, My.Width - 4, My.Height - 4), e)

                        c1 = cRGB(238, 240, 241)
                        c2 = cRGB(206, 209, 214)
                        c3 = Color.White

                        lgb = New System.Drawing.Drawing2D.LinearGradientBrush(New Point(0, 1), New Point(0, My.Height - 1), c1, c2)
                        cb.Colors = New Color() {c1, c2, c3}
                        cb.Positions = New Single() {0, 0.75, 1}
                        lgb.InterpolationColors = cb
                        lgb.GammaCorrection = True

                        FillRoundedRectangle(lgb, 1, New Rectangle(3, 3, My.Width - 6, My.Height - 6), e)
                    End Sub

                    Private Sub DrawDownButtonBody(ByVal e As PaintEventArgs)
                        Dim c1, c2 As Color
                        Dim b As SolidBrush
                        Dim lgb As System.Drawing.Drawing2D.LinearGradientBrush

                        b = New SolidBrush(cRGB(137, 141, 146))
                        FillRoundedRectangle(b, 3, New Rectangle(1, 1, My.Width - 2, My.Height - 2), e)

                        b = New SolidBrush(cRGB(255, 255, 255))
                        FillRoundedRectangle(b, 3, New Rectangle(2, 2, My.Width - 3, My.Height - 3), e)


                        c1 = cRGB(170, 175, 185)
                        c2 = cRGB(230, 232, 234)

                        lgb = New System.Drawing.Drawing2D.LinearGradientBrush(New Point(0, 1), New Point(0, My.Height - 1), c1, c2)
                        lgb.GammaCorrection = True

                        FillRoundedRectangle(lgb, 1, New Rectangle(2, 2, My.Width - 4, My.Height - 4), e)

                    End Sub

                    Private Sub DrawOverButtonBody(ByVal e As PaintEventArgs)
                        Dim c1, c2, c3 As Color
                        Dim b As SolidBrush
                        Dim lgb As System.Drawing.Drawing2D.LinearGradientBrush
                        Dim cb As New System.Drawing.Drawing2D.ColorBlend

                        b = New SolidBrush(cRGB(255, 214, 153))
                        FillRoundedRectangle(b, 3, New Rectangle(1, 1, My.Width - 2, My.Height - 2), e)

                        b = New SolidBrush(cRGB(204, 122, 0))
                        FillRoundedRectangle(b, 3, New Rectangle(2, 2, My.Width - 3, My.Height - 3), e)

                        b = New SolidBrush(cRGB(255, 153, 0))
                        FillRoundedRectangle(b, 2, New Rectangle(2, 2, My.Width - 4, My.Height - 4), e)

                        c1 = cRGB(243, 244, 245)
                        c2 = cRGB(218, 221, 224)
                        c3 = Color.White

                        lgb = New System.Drawing.Drawing2D.LinearGradientBrush(New Point(0, 1), New Point(0, My.Height - 1), c1, c2)
                        cb.Colors = New Color() {c1, c2, c3}
                        cb.Positions = New Single() {0, 0.75, 1}
                        lgb.InterpolationColors = cb
                        lgb.GammaCorrection = True

                        FillRoundedRectangle(lgb, 1, New Rectangle(3, 3, My.Width - 6, My.Height - 6), e)

                    End Sub

                    Private Sub DrawButtonBody(ByVal e As PaintEventArgs)
                        Dim c1, c2, c3 As Color
                        Dim b As SolidBrush
                        Dim lgb As System.Drawing.Drawing2D.LinearGradientBrush
                        Dim cb As New System.Drawing.Drawing2D.ColorBlend

                        b = New SolidBrush(Color.White)
                        FillRoundedRectangle(b, 3, New Rectangle(1, 1, My.Width - 2, My.Height - 2), e)

                        b = New SolidBrush(cRGB(185, 185, 185))
                        FillRoundedRectangle(b, 3, New Rectangle(2, 2, My.Width - 3, My.Height - 3), e)

                        c1 = cRGB(238, 240, 241)
                        c2 = cRGB(206, 209, 214)
                        c3 = Color.White

                        lgb = New System.Drawing.Drawing2D.LinearGradientBrush(New Point(0, 1), New Point(0, My.Height - 1), c1, c3)
                        cb.Colors = New Color() {c1, c2, c3}
                        cb.Positions = New Single() {0.0, 0.75, 1.0}
                        lgb.InterpolationColors = cb
                        lgb.GammaCorrection = True

                        FillRoundedRectangle(lgb, 2, New Rectangle(2, 2, My.Width - 4, My.Height - 4), e)
                    End Sub

                    Public Sub DrawTheme(ByVal e As PaintEventArgs, ByVal state As Theme.States, ByVal textfx As Boolean)
                        DrawBase(e, state)
                        TextEffects = textfx
                        If state = Theme.States.Normal Then
                            DrawButtonBody(e)
                        ElseIf state = Theme.States.MouseOver Then
                            DrawOverButtonBody(e)
                        ElseIf state = States.Focused Then
                            DrawFocusButtonBody(e)
                        ElseIf state = States.MouseDown Then
                            DrawDownButtonBody(e)
                        ElseIf state = States.Disabled Then
                            DrawDisabledButtonBody(e)
                        End If
                    End Sub

                    Public Sub DrawText(ByVal e As PaintEventArgs, ByVal s As States, ByVal text As String, ByVal font As Font, ByVal forecolor As Color)

                        Dim sf As StringFormat
                        Dim sz As SizeF
                        Dim b As Brush

                        sf = New StringFormat
                        sf.LineAlignment = StringAlignment.Center
                        sf.Alignment = StringAlignment.Center
                        sf.Trimming = StringTrimming.EllipsisCharacter

                        sz = e.Graphics.MeasureString(text, font, New SizeF(My.Width - 2, My.Height - 2), sf)

                        text = text & "                    "

                        If s = States.MouseDown Then
                            If TextEffects = True Then
                                b = New SolidBrush(Color.FromArgb(0.5 * 255, Color.White))
                                e.Graphics.DrawString(text, font, b, New RectangleF(3, 3, My.Width - 2, My.Height - 2), sf)
                            End If
                            b = New SolidBrush(forecolor)
                            e.Graphics.DrawString(text, font, b, New RectangleF(2, 2, My.Width - 2, My.Height - 2), sf)
                        ElseIf s = States.Disabled Then
                            b = New SolidBrush(Color.FromKnownColor(KnownColor.GrayText))
                            e.Graphics.DrawString(text, font, b, New RectangleF(1, 1, My.Width - 2, My.Height - 2), sf)
                        Else
                            If TextEffects = True Then
                                b = New SolidBrush(Color.FromArgb(0.5 * 255, Color.White))
                                e.Graphics.DrawString(text, font, b, New RectangleF(2, 2, My.Width - 2, My.Height - 2), sf)
                            End If
                            b = New SolidBrush(forecolor)
                            e.Graphics.DrawString(text, font, b, New RectangleF(1, 1, My.Width - 2, My.Height - 2), sf)
                        End If

                    End Sub

                End Class
#End Region

#Region " Soft Glass XP Theme "
                Public Class SoftGlassXP

                    Private My As Rectangle
                    Private ThemeColor As Color
                    Private TextEffects As Boolean

                    Public Sub New(ByVal Owner As Rectangle)
                        MyBase.New()
                        My = Owner
                    End Sub

                    Private Sub FillRoundedRectangle(ByVal b As Brush, ByVal Radius As Integer, ByVal Rect As Rectangle, ByVal e As PaintEventArgs)
                        Dim TL, TR, BL, BR As Point 'The four corners of the rectabgle

                        'Set the values of each corner point
                        TL = New Point(Rect.Left, Rect.Top)
                        TR = New Point(Rect.Left + Rect.Width, Rect.Top)
                        BL = New Point(Rect.Left, Rect.Top + Rect.Height)
                        BR = New Point(Rect.Left + Rect.Width, Rect.Top + Rect.Height)

                        'Draws the four corner circles
                        e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias 'Changes smoothing mode to anti-alias to remove jagged edges
                        e.Graphics.FillEllipse(b, New Rectangle(TL.X, TL.Y, Radius * 2, Radius * 2)) 'Top-left circle
                        e.Graphics.FillEllipse(b, New Rectangle(BL.X, BL.Y - (Radius * 2) - 1, Radius * 2, Radius * 2)) 'Bottom-left circle
                        e.Graphics.FillEllipse(b, New Rectangle(TR.X - (Radius * 2) - 1, TR.Y, Radius * 2, Radius * 2)) 'Top-right circle
                        e.Graphics.FillEllipse(b, New Rectangle(BR.X - (Radius * 2) - 1, BR.Y - (Radius * 2) - 1, Radius * 2, Radius * 2)) 'Bottom-right circle

                        'Draws the two blocks
                        e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.Default 'Returns the smoothing mode to default for a crisp structure
                        e.Graphics.FillRectangle(b, New Rectangle(TL.X, TL.Y + Radius, Rect.Width, Rect.Height - (Radius * 2)))
                        e.Graphics.FillRectangle(b, New Rectangle(TL.X + Radius, TL.Y, Rect.Width - (Radius * 2), Rect.Height))
                    End Sub

                    Private Sub DrawBase(ByVal e As PaintEventArgs, ByVal s As States)
                        Dim b As Brush
                        If s = States.Disabled Then
                            b = New SolidBrush(cRGB(201, 199, 186))
                        Else
                            b = New SolidBrush(OpacityMix(Color.Black, ThemeColor, 45))
                        End If
                        FillRoundedRectangle(b, 4, New Rectangle(0, 0, My.Width, My.Height), e)
                    End Sub

                    Private Sub DrawDisabledButtonBody(ByVal e As PaintEventArgs)
                        Dim b As Brush
                        b = New SolidBrush(cRGB(245, 244, 234))
                        FillRoundedRectangle(b, 3, New Rectangle(My.Left + 1, My.Top + 1, My.Width - 2, My.Height - 2), e)
                    End Sub

                    Private Sub DrawFocusButtonBody(ByVal e As PaintEventArgs)
                        Dim c1, c2, c3 As Color
                        Dim b As SolidBrush
                        Dim lgb As System.Drawing.Drawing2D.LinearGradientBrush
                        Dim cb As New System.Drawing.Drawing2D.ColorBlend

                        b = New SolidBrush(SoftLightMix(SoftLightMix(SoftLightMix(ThemeColor, Color.Black, 60), Color.White, 100), Color.White, 60))
                        FillRoundedRectangle(b, 3, New Rectangle(1, 1, My.Width - 2, My.Height - 2), e)

                        b = New SolidBrush(SoftLightMix(SoftLightMix(ThemeColor, Color.White, 60), Color.Black, 60))
                        FillRoundedRectangle(b, 3, New Rectangle(2, 2, My.Width - 3, My.Height - 3), e)

                        c1 = SoftLightMix(SoftLightMix(ThemeColor, Color.Black, 60), Color.White, 100)
                        c2 = SoftLightMix(ThemeColor, cRGB(52, 52, 52), 60)
                        c3 = SoftLightMix(ThemeColor, Color.White, 60)

                        cb.Colors = New Color() {c1, c2, c3}
                        cb.Positions = New Single() {0.0, 0.25, 1.0}

                        lgb = New System.Drawing.Drawing2D.LinearGradientBrush(New Point(0, 1), New Point(0, My.Height - 1), c1, c3)
                        lgb.InterpolationColors = cb
                        lgb.GammaCorrection = True

                        FillRoundedRectangle(lgb, 2, New Rectangle(2, 2, My.Width - 4, My.Height - 4), e)
                    End Sub

                    Private Sub DrawDownButtonBody(ByVal e As PaintEventArgs)
                        Dim c1, c2, c3 As Color
                        Dim b As SolidBrush
                        Dim lgb As System.Drawing.Drawing2D.LinearGradientBrush
                        Dim cb As New System.Drawing.Drawing2D.ColorBlend

                        b = New SolidBrush(SoftLightMix(SoftLightMix(SoftLightMix(ThemeColor, Color.White, 40), Color.Black, 100), Color.Black, 50))
                        FillRoundedRectangle(b, 3, New Rectangle(1, 1, My.Width - 2, My.Height - 2), e)

                        b = New SolidBrush(SoftLightMix(SoftLightMix(ThemeColor, Color.Black, 60), Color.White, 100))
                        FillRoundedRectangle(b, 3, New Rectangle(2, 2, My.Width - 3, My.Height - 3), e)

                        c1 = SoftLightMix(SoftLightMix(ThemeColor, Color.White, 40), Color.Black, 100)
                        c2 = SoftLightMix(ThemeColor, cRGB(203, 203, 203), 60)
                        c3 = SoftLightMix(ThemeColor, Color.Black, 60)

                        cb.Colors = New Color() {c1, c2, c3}
                        cb.Positions = New Single() {0.0, 0.25, 1.0}

                        lgb = New System.Drawing.Drawing2D.LinearGradientBrush(New Point(0, 1), New Point(0, My.Height - 1), c1, c3)
                        lgb.InterpolationColors = cb
                        lgb.GammaCorrection = True

                        FillRoundedRectangle(lgb, 2, New Rectangle(2, 2, My.Width - 4, My.Height - 4), e)
                    End Sub

                    Private Sub DrawOverButtonBody(ByVal e As PaintEventArgs)
                        Dim c1, c2, c3 As Color
                        Dim b As SolidBrush
                        Dim lgb As System.Drawing.Drawing2D.LinearGradientBrush
                        Dim cb As New System.Drawing.Drawing2D.ColorBlend

                        b = New SolidBrush(SoftLightMix(SoftLightMix(SoftLightMix(SoftLightMix(ThemeColor, Color.White, 50), Color.Black, 60), Color.White, 100), Color.White, 60))
                        FillRoundedRectangle(b, 3, New Rectangle(1, 1, My.Width - 2, My.Height - 2), e)

                        b = New SolidBrush(SoftLightMix(SoftLightMix(SoftLightMix(ThemeColor, Color.White, 50), Color.White, 60), Color.Black, 60))
                        FillRoundedRectangle(b, 3, New Rectangle(2, 2, My.Width - 3, My.Height - 3), e)

                        c1 = SoftLightMix(SoftLightMix(SoftLightMix(ThemeColor, Color.White, 50), Color.Black, 60), Color.White, 100)
                        c2 = SoftLightMix(SoftLightMix(ThemeColor, Color.White, 50), cRGB(52, 52, 52), 60)
                        c3 = SoftLightMix(SoftLightMix(ThemeColor, Color.White, 50), Color.White, 60)

                        cb.Colors = New Color() {c1, c2, c3}
                        cb.Positions = New Single() {0.0, 0.25, 1.0}

                        lgb = New System.Drawing.Drawing2D.LinearGradientBrush(New Point(0, 1), New Point(0, My.Height - 1), c1, c3)
                        lgb.InterpolationColors = cb
                        lgb.GammaCorrection = True

                        FillRoundedRectangle(lgb, 2, New Rectangle(2, 2, My.Width - 4, My.Height - 4), e)
                    End Sub

                    Private Sub DrawButtonBody(ByVal e As PaintEventArgs)
                        Dim c1, c2, c3 As Color
                        Dim b As SolidBrush
                        Dim lgb As System.Drawing.Drawing2D.LinearGradientBrush
                        Dim cb As New System.Drawing.Drawing2D.ColorBlend

                        b = New SolidBrush(SoftLightMix(SoftLightMix(SoftLightMix(ThemeColor, Color.Black, 60), Color.White, 100), Color.White, 60))
                        FillRoundedRectangle(b, 3, New Rectangle(1, 1, My.Width - 2, My.Height - 2), e)

                        b = New SolidBrush(SoftLightMix(SoftLightMix(ThemeColor, Color.White, 60), Color.Black, 60))
                        FillRoundedRectangle(b, 3, New Rectangle(2, 2, My.Width - 3, My.Height - 3), e)

                        c1 = SoftLightMix(SoftLightMix(ThemeColor, Color.Black, 60), Color.White, 100)
                        c2 = SoftLightMix(ThemeColor, cRGB(52, 52, 52), 60)
                        c3 = SoftLightMix(ThemeColor, Color.White, 60)

                        cb.Colors = New Color() {c1, c2, c3}
                        cb.Positions = New Single() {0.0, 0.25, 1.0}

                        lgb = New System.Drawing.Drawing2D.LinearGradientBrush(New Point(0, 1), New Point(0, My.Height - 1), c1, c3)
                        lgb.InterpolationColors = cb
                        lgb.GammaCorrection = True

                        FillRoundedRectangle(lgb, 2, New Rectangle(2, 2, My.Width - 4, My.Height - 4), e)
                    End Sub

                    Public Sub DrawTheme(ByVal e As PaintEventArgs, ByVal state As Theme.States, ByVal tcolor As Color, ByVal TextFx As Boolean)
                        ThemeColor = tcolor

                        TextEffects = TextFx

                        DrawBase(e, state)
                        If state = Theme.States.Normal Then
                            DrawButtonBody(e)
                        ElseIf state = Theme.States.MouseOver Then
                            DrawOverButtonBody(e)
                        ElseIf state = States.Focused Then
                            DrawFocusButtonBody(e)
                        ElseIf state = States.MouseDown Then
                            DrawDownButtonBody(e)
                        ElseIf state = States.Disabled Then
                            DrawDisabledButtonBody(e)
                        End If
                    End Sub

                    Public Sub DrawText(ByVal e As PaintEventArgs, ByVal s As States, ByVal text As String, ByVal font As Font, ByVal forecolor As Color)

                        Dim sf As StringFormat
                        Dim sz As SizeF
                        Dim b As Brush

                        sf = New StringFormat
                        sf.LineAlignment = StringAlignment.Center
                        sf.Alignment = StringAlignment.Center
                        sf.Trimming = StringTrimming.EllipsisCharacter

                        sz = e.Graphics.MeasureString(text, font, New SizeF(My.Width - 2, My.Height - 2), sf)
                        text = text & "                    "

                        If s = States.MouseDown Then
                            If TextEffects = True Then
                                b = New SolidBrush(Color.FromArgb(0.3 * 255, Color.Black))
                                e.Graphics.DrawString(text, font, b, New RectangleF(3, 3, My.Width - 2, My.Height - 2), sf)
                            End If
                            b = New SolidBrush(forecolor)
                            e.Graphics.DrawString(text, font, b, New RectangleF(2, 2, My.Width - 2, My.Height - 2), sf)
                        ElseIf s = States.Disabled Then
                            b = New SolidBrush(Color.FromKnownColor(KnownColor.GrayText))
                            e.Graphics.DrawString(text, font, b, New RectangleF(1, 1, My.Width - 2, My.Height - 2), sf)
                        Else
                            If TextEffects = True Then
                                b = New SolidBrush(Color.FromArgb(0.3 * 255, Color.Black))
                                e.Graphics.DrawString(text, font, b, New RectangleF(2, 2, My.Width - 2, My.Height - 2), sf)
                            End If
                            b = New SolidBrush(forecolor)
                            e.Graphics.DrawString(text, font, b, New RectangleF(1, 1, My.Width - 2, My.Height - 2), sf)
                        End If

                    End Sub

                End Class
#End Region

#Region " Windows XP Theme "
                Public Class WindowsXP

                    Private My As Rectangle
                    Private TextEffects As Boolean

                    Public Sub New(ByVal Owner As Rectangle)
                        MyBase.New()
                        My = Owner
                    End Sub

                    Private Sub FillRoundedRectangle(ByVal b As Brush, ByVal Radius As Integer, ByVal Rect As Rectangle, ByVal e As PaintEventArgs)
                        Dim TL, TR, BL, BR As Point 'The four corners of the rectabgle

                        'Set the values of each corner point
                        TL = New Point(Rect.Left, Rect.Top)
                        TR = New Point(Rect.Left + Rect.Width, Rect.Top)
                        BL = New Point(Rect.Left, Rect.Top + Rect.Height)
                        BR = New Point(Rect.Left + Rect.Width, Rect.Top + Rect.Height)

                        'Draws the four corner circles
                        e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias 'Changes smoothing mode to anti-alias to remove jagged edges
                        e.Graphics.FillEllipse(b, New Rectangle(TL.X, TL.Y, Radius * 2, Radius * 2)) 'Top-left circle
                        e.Graphics.FillEllipse(b, New Rectangle(BL.X, BL.Y - (Radius * 2) - 1, Radius * 2, Radius * 2)) 'Bottom-left circle
                        e.Graphics.FillEllipse(b, New Rectangle(TR.X - (Radius * 2) - 1, TR.Y, Radius * 2, Radius * 2)) 'Top-right circle
                        e.Graphics.FillEllipse(b, New Rectangle(BR.X - (Radius * 2) - 1, BR.Y - (Radius * 2) - 1, Radius * 2, Radius * 2)) 'Bottom-right circle

                        'Draws the two blocks
                        e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.Default 'Returns the smoothing mode to default for a crisp structure
                        e.Graphics.FillRectangle(b, New Rectangle(TL.X, TL.Y + Radius, Rect.Width, Rect.Height - (Radius * 2)))
                        e.Graphics.FillRectangle(b, New Rectangle(TL.X + Radius, TL.Y, Rect.Width - (Radius * 2), Rect.Height))
                    End Sub

                    Private Sub DrawBase(ByVal e As PaintEventArgs, ByVal state As States)
                        Dim b As Brush

                        If state = States.Disabled Then
                            b = New SolidBrush(cRGB(201, 199, 186))
                        Else
                            b = New SolidBrush(cRGB(0, 60, 116))
                        End If

                        FillRoundedRectangle(b, 2, New Rectangle(My.Left, My.Top, My.Width, My.Height), e)
                    End Sub

                    Private Sub DrawDisabledButtonBody(ByVal e As PaintEventArgs)
                        Dim b As Brush
                        b = New SolidBrush(cRGB(245, 244, 234))
                        FillRoundedRectangle(b, 1, New Rectangle(My.Left + 1, My.Top + 1, My.Width - 2, My.Height - 2), e)
                    End Sub

                    Private Sub DrawFocusButtonBody(ByVal e As PaintEventArgs)
                        Dim c1, c2, c3 As Color
                        Dim lgb As System.Drawing.Drawing2D.LinearGradientBrush
                        Dim cb As New System.Drawing.Drawing2D.ColorBlend

                        c1 = cRGB(194, 219, 255)
                        c2 = cRGB(140, 180, 242)

                        lgb = New System.Drawing.Drawing2D.LinearGradientBrush(New Point(0, 1), New Point(0, My.Height - 1), c1, c2)
                        lgb.GammaCorrection = True

                        FillRoundedRectangle(lgb, 1, New Rectangle(My.Left + 1, My.Top + 1, My.Width - 2, My.Height - 2), e)

                        e.Graphics.DrawLine(New Pen(cRGB(206, 231, 255)), 2, 1, My.Width - 3, 1)
                        e.Graphics.DrawLine(New Pen(cRGB(105, 130, 238)), 2, My.Height - 2, My.Width - 3, My.Height - 2)

                        c1 = cRGB(255, 255, 255)
                        c2 = cRGB(244, 242, 232)
                        c3 = cRGB(220, 215, 203)

                        lgb = New System.Drawing.Drawing2D.LinearGradientBrush(New Point(My.Left, My.Top + 1), New Point(My.Left, My.Height - 1), c1, c3)

                        cb.Colors = New Color() {c1, c2, c3}
                        cb.Positions = New Single() {0.0, (My.Height - 5) / My.Height, 1.0}

                        lgb.InterpolationColors = cb
                        lgb.GammaCorrection = True

                        e.Graphics.FillRectangle(lgb, My.Left + 3, My.Top + 3, My.Width - 6, My.Height - 6)
                    End Sub

                    Private Sub DrawDownButtonBody(ByVal e As PaintEventArgs)
                        Dim c1, c2, c3 As Color
                        Dim lgb As System.Drawing.Drawing2D.LinearGradientBrush
                        Dim cb As New System.Drawing.Drawing2D.ColorBlend

                        c1 = cRGB(209, 204, 193)
                        c2 = cRGB(227, 226, 218)
                        c3 = cRGB(239, 238, 234)

                        lgb = New System.Drawing.Drawing2D.LinearGradientBrush(New Point(0, My.Top + 1), New Point(0, My.Height - 1), c1, c3)

                        cb.Colors = New Color() {c1, c2, c2, c3}
                        cb.Positions = New Single() {0.0, 4 / My.Height, (My.Height - 5) / My.Height, 1.0}

                        lgb.InterpolationColors = cb
                        lgb.GammaCorrection = True

                        FillRoundedRectangle(lgb, 1, New Rectangle(My.Left + 1, My.Top + 1, My.Width - 2, My.Height - 2), e)

                        e.Graphics.DrawLine(New Pen(c1), My.Left + 1, My.Top + 2, My.Left + 1, My.Height - 3)
                    End Sub

                    Private Sub DrawOverButtonBody(ByVal e As PaintEventArgs)
                        Dim c1, c2, c3 As Color
                        Dim lgb As System.Drawing.Drawing2D.LinearGradientBrush
                        Dim cb As New System.Drawing.Drawing2D.ColorBlend

                        c1 = cRGB(255, 218, 140)
                        c2 = cRGB(255, 184, 52)

                        lgb = New System.Drawing.Drawing2D.LinearGradientBrush(New Point(0, 1), New Point(0, My.Height - 1), c1, c2)
                        lgb.GammaCorrection = True

                        FillRoundedRectangle(lgb, 1, New Rectangle(My.Left + 1, My.Top + 1, My.Width - 2, My.Height - 2), e)

                        e.Graphics.DrawLine(New Pen(cRGB(255, 240, 207)), 2, 1, My.Width - 3, 1)
                        e.Graphics.DrawLine(New Pen(cRGB(229, 151, 0)), 2, My.Height - 2, My.Width - 3, My.Height - 2)

                        c1 = cRGB(255, 255, 255)
                        c2 = cRGB(244, 242, 232)
                        c3 = cRGB(220, 215, 203)

                        lgb = New System.Drawing.Drawing2D.LinearGradientBrush(New Point(0, 1), New Point(0, My.Height - 1), c1, c3)

                        cb.Colors = New Color() {c1, c2, c3}
                        cb.Positions = New Single() {0.0, (My.Height - 5) / My.Height, 1.0}

                        lgb.InterpolationColors = cb
                        lgb.GammaCorrection = True

                        e.Graphics.FillRectangle(lgb, My.Left + 3, My.Top + 3, My.Width - 6, My.Height - 6)
                    End Sub

                    Private Sub DrawButtonBody(ByVal e As PaintEventArgs)
                        Dim c1, c2, c3 As Color
                        Dim lgb As System.Drawing.Drawing2D.LinearGradientBrush
                        Dim cb As New System.Drawing.Drawing2D.ColorBlend

                        c1 = cRGB(255, 255, 255)
                        c2 = cRGB(244, 242, 232)
                        c3 = cRGB(220, 215, 203)

                        lgb = New System.Drawing.Drawing2D.LinearGradientBrush(New Point(0, 1), New Point(0, My.Height - 1), c1, c3)

                        cb.Colors = New Color() {c1, c2, c3}
                        cb.Positions = New Single() {0.0, (My.Height - 5) / My.Height, 1.0}

                        lgb.InterpolationColors = cb
                        lgb.GammaCorrection = True

                        FillRoundedRectangle(lgb, 1, New Rectangle(My.Left + 1, My.Top + 1, My.Width - 2, My.Height - 2), e)

                        e.Graphics.DrawLine(New Pen(c3), My.Width - 2, 2, My.Width - 2, My.Height - 3)
                    End Sub

                    Public Sub DrawTheme(ByVal e As PaintEventArgs, ByVal state As Theme.States, ByVal TextFX As Boolean)
                        DrawBase(e, state)

                        TextEffects = TextFX

                        If state = Theme.States.Normal Then
                            DrawButtonBody(e)
                        ElseIf state = Theme.States.MouseOver Then
                            DrawOverButtonBody(e)
                        ElseIf state = States.Focused Then
                            DrawFocusButtonBody(e)
                        ElseIf state = States.MouseDown Then
                            DrawDownButtonBody(e)
                        ElseIf state = States.Disabled Then
                            DrawDisabledButtonBody(e)
                        End If
                    End Sub

                    Public Sub DrawText(ByVal e As PaintEventArgs, ByVal s As States, ByVal text As String, ByVal font As Font, ByVal forecolor As Color)

                        Dim sf As StringFormat
                        Dim sz As SizeF
                        Dim b As Brush

                        sf = New StringFormat
                        sf.LineAlignment = StringAlignment.Center
                        sf.Alignment = StringAlignment.Center
                        sf.Trimming = StringTrimming.EllipsisCharacter

                        sz = e.Graphics.MeasureString(text, font, New SizeF(My.Width - 2, My.Height - 2), sf)
                        text = text & "                    "

                        If s = States.MouseDown Then
                            If TextEffects = True Then
                                b = New SolidBrush(Color.FromArgb(0.75 * 255, Color.White))
                                e.Graphics.DrawString(text, font, b, New RectangleF(3, 3, My.Width - 2, My.Height - 2), sf)
                            End If
                            b = New SolidBrush(forecolor)
                            e.Graphics.DrawString(text, font, b, New RectangleF(2, 2, My.Width - 2, My.Height - 2), sf)

                        ElseIf s = States.Disabled Then
                            b = New SolidBrush(Color.FromKnownColor(KnownColor.GrayText))
                            e.Graphics.DrawString(text, font, b, New RectangleF(1, 1, My.Width - 2, My.Height - 2), sf)
                        Else
                            If TextEffects = True Then
                                b = New SolidBrush(Color.FromArgb(0.75 * 255, Color.White))
                                e.Graphics.DrawString(text, font, b, New RectangleF(2, 2, My.Width - 2, My.Height - 2), sf)
                            End If
                            b = New SolidBrush(forecolor)
                            e.Graphics.DrawString(text, font, b, New RectangleF(1, 1, My.Width - 2, My.Height - 2), sf)
                        End If

                    End Sub

                End Class
#End Region

#Region " MSN Login Button Theme "
                Public Class MSNLoginButton

                    Private My As Rectangle
                    Private TextEffects As Boolean

                    Public Sub New(ByVal Owner As Rectangle)
                        MyBase.New()
                        My = Owner
                    End Sub

                    Private Sub DrawBase(ByVal e As PaintEventArgs, ByVal state As States, ByVal backcolor As Color)
                        Dim b As Brush
                        Dim c As Color

                        If state = States.MouseDown Then
                            c = Color.FromArgb(64, 7, 66, 131)
                            b = New SolidBrush(c)

                            e.Graphics.FillRectangle(b, My.Left, My.Top + 1, My.Width - 1, My.Height - 3)
                            e.Graphics.FillRectangle(b, My.Left + 1, My.Top, My.Width - 3, My.Height - 1)

                            c = cRGB(7, 66, 131)
                            b = New SolidBrush(c)

                            e.Graphics.FillRectangle(b, My.Left + 1, My.Top + 2, My.Width - 1, My.Height - 3)
                            e.Graphics.FillRectangle(b, My.Left + 2, My.Top + 1, My.Width - 3, My.Height - 1)
                        ElseIf state = States.Disabled Then
                            c = cRGB(117, 120, 138)
                            b = New SolidBrush(c)

                            e.Graphics.FillRectangle(b, My.Left, My.Top + 1, My.Width - 1, My.Height - 3)
                            e.Graphics.FillRectangle(b, My.Left + 1, My.Top, My.Width - 3, My.Height - 1)
                        Else
                            c = cRGB(7, 66, 131)
                            b = New SolidBrush(c)

                            e.Graphics.FillRectangle(b, My.Left, My.Top + 1, My.Width - 1, My.Height - 3)
                            e.Graphics.FillRectangle(b, My.Left + 1, My.Top, My.Width - 3, My.Height - 1)
                        End If
                    End Sub

                    Private Sub DrawDisabledButtonBody(ByVal e As PaintEventArgs)
                        Dim sb As SolidBrush
                        Dim x, y As Single

                        sb = New SolidBrush(cRGB(176, 178, 191))
                        e.Graphics.FillRectangle(sb, My.Left + 1, My.Top + 1, My.Width - 3, My.Height - 3)

                        sb = New SolidBrush(cRGB(224, 226, 236))
                        e.Graphics.FillRectangle(sb, My.Left + 2, My.Top + 2, My.Width - 5, My.Height - 5)

                        x = My.Left + 2
                        y = My.Top + My.Height - 4

                        e.Graphics.DrawLine(New Pen(cRGB(194, 197, 209)), x, y, My.Left + My.Width - 3, y)
                    End Sub

                    Private Sub DrawFocusButtonBody(ByVal e As PaintEventArgs)
                        DrawButtonBody(e)
                    End Sub

                    Private Sub DrawDownButtonBody(ByVal e As PaintEventArgs)
                        Dim c1, c2 As Color
                        Dim lgb As System.Drawing.Drawing2D.LinearGradientBrush
                        Dim sb As SolidBrush
                        Dim cb As New System.Drawing.Drawing2D.ColorBlend
                        Dim x, y, z As Single

                        sb = New SolidBrush(cRGB(123, 140, 196))
                        e.Graphics.FillRectangle(sb, My.Left + 2, My.Top + 2, My.Width - 3, My.Height - 3)

                        sb = New SolidBrush(Color.White)
                        e.Graphics.FillRectangle(sb, My.Left + 3, My.Top + 3, My.Width - 5, My.Height - 5)

                        c1 = Color.White
                        c2 = cRGB(218, 225, 252)

                        lgb = New System.Drawing.Drawing2D.LinearGradientBrush(New Point(0, My.Top + 3), New Point(0, My.Top + My.Height - 2), c1, c2)
                        cb = New System.Drawing.Drawing2D.ColorBlend
                        cb.Colors = New Color() {c1, c1, c2}
                        cb.Positions = New Single() {0, 2 / (My.Height - 3), 1}
                        lgb.InterpolationColors = cb
                        lgb.GammaCorrection = True

                        e.Graphics.FillRectangle(lgb, My.Left + 5, My.Top + 5, My.Width - 10, My.Height - 7)

                        e.Graphics.DrawLine(New Pen(cRGB(186, 201, 245)), My.Left + My.Width - 3, My.Top + 4, My.Left + My.Width - 3, My.Top + My.Height - 3)

                        x = My.Left + 3
                        y = My.Top + My.Height - 2
                        z = (My.Width - 4) / 4

                        e.Graphics.DrawLine(New Pen(cRGB(249, 172, 20)), x, y, My.Left + My.Width - 3, y)
                        x += z
                        e.Graphics.DrawLine(New Pen(cRGB(152, 53, 3)), x, y, My.Left + My.Width - 3, y)
                        x += z
                        e.Graphics.DrawLine(New Pen(cRGB(7, 80, 143)), x, y, My.Left + My.Width - 3, y)
                        x += z
                        e.Graphics.DrawLine(New Pen(cRGB(17, 147, 49)), x, y, My.Left + My.Width - 3, y)

                        x = My.Left + 3
                        y -= 1

                        e.Graphics.DrawLine(New Pen(cRGB(251, 181, 42)), x, y, My.Left + My.Width - 3, y)
                        x += z
                        e.Graphics.DrawLine(New Pen(cRGB(236, 112, 102)), x, y, My.Left + My.Width - 3, y)
                        x += z
                        e.Graphics.DrawLine(New Pen(cRGB(39, 119, 209)), x, y, My.Left + My.Width - 3, y)
                        x += z
                        e.Graphics.DrawLine(New Pen(cRGB(36, 190, 74)), x, y, My.Left + My.Width - 3, y)
                    End Sub

                    Private Sub DrawOverButtonBody(ByVal e As PaintEventArgs)
                        DrawButtonBody(e)
                    End Sub

                    Private Sub DrawButtonBody(ByVal e As PaintEventArgs)
                        Dim c1, c2 As Color
                        Dim lgb As System.Drawing.Drawing2D.LinearGradientBrush
                        Dim sb As SolidBrush
                        Dim cb As New System.Drawing.Drawing2D.ColorBlend
                        Dim x, y, z As Single

                        sb = New SolidBrush(cRGB(238, 163, 53))
                        e.Graphics.FillRectangle(sb, My.Left + 1, My.Top + 1, My.Width - 3, My.Height - 3)

                        sb = New SolidBrush(Color.White)
                        e.Graphics.FillRectangle(sb, My.Left + 2, My.Top + 2, My.Width - 5, My.Height - 5)

                        c1 = Color.White
                        c2 = cRGB(218, 225, 252)

                        lgb = New System.Drawing.Drawing2D.LinearGradientBrush(New Point(0, My.Top + 2), New Point(0, My.Top + My.Height - 3), c1, c2)
                        cb = New System.Drawing.Drawing2D.ColorBlend
                        cb.Colors = New Color() {c1, c1, c2}
                        cb.Positions = New Single() {0, 2 / (My.Height - 3), 1}
                        lgb.InterpolationColors = cb
                        lgb.GammaCorrection = True

                        e.Graphics.FillRectangle(lgb, My.Left + 4, My.Top + 4, My.Width - 10, My.Height - 7)

                        e.Graphics.DrawLine(New Pen(cRGB(186, 201, 245)), My.Left + My.Width - 4, My.Top + 3, My.Left + My.Width - 4, My.Top + My.Height - 4)

                        x = My.Left + 2
                        y = My.Top + My.Height - 3
                        z = (My.Width - 4) / 4

                        e.Graphics.DrawLine(New Pen(cRGB(249, 172, 20)), x, y, My.Left + My.Width - 3, y)
                        x += z
                        e.Graphics.DrawLine(New Pen(cRGB(152, 53, 3)), x, y, My.Left + My.Width - 3, y)
                        x += z
                        e.Graphics.DrawLine(New Pen(cRGB(7, 80, 143)), x, y, My.Left + My.Width - 3, y)
                        x += z
                        e.Graphics.DrawLine(New Pen(cRGB(17, 147, 49)), x, y, My.Left + My.Width - 3, y)

                        x = My.Left + 2
                        y -= 1

                        e.Graphics.DrawLine(New Pen(cRGB(251, 181, 42)), x, y, My.Left + My.Width - 3, y)
                        x += z
                        e.Graphics.DrawLine(New Pen(cRGB(236, 112, 102)), x, y, My.Left + My.Width - 3, y)
                        x += z
                        e.Graphics.DrawLine(New Pen(cRGB(39, 119, 209)), x, y, My.Left + My.Width - 3, y)
                        x += z
                        e.Graphics.DrawLine(New Pen(cRGB(36, 190, 74)), x, y, My.Left + My.Width - 3, y)

                    End Sub

                    Public Sub DrawTheme(ByVal e As PaintEventArgs, ByVal state As Theme.States, ByVal textfx As Boolean, ByVal backcolor As Color)
                        DrawBase(e, state, backcolor)
                        TextEffects = textfx
                        If state = Theme.States.Normal Then
                            DrawButtonBody(e)
                        ElseIf state = Theme.States.MouseOver Then
                            DrawOverButtonBody(e)
                        ElseIf state = States.Focused Then
                            DrawFocusButtonBody(e)
                        ElseIf state = States.MouseDown Then
                            DrawDownButtonBody(e)
                        ElseIf state = States.Disabled Then
                            DrawDisabledButtonBody(e)
                        End If
                    End Sub

                    Public Sub DrawText(ByVal e As PaintEventArgs, ByVal s As States, ByVal text As String, ByVal font As Font)

                        Dim sf As StringFormat
                        Dim sz As SizeF
                        Dim b As Brush

                        sf = New StringFormat
                        sf.LineAlignment = StringAlignment.Center
                        sf.Alignment = StringAlignment.Center
                        sf.Trimming = StringTrimming.EllipsisCharacter

                        sz = e.Graphics.MeasureString(text, font, New SizeF(My.Width - 2, My.Height - 2), sf)

                        text = text & "                    "

                        If s = States.MouseDown Then
                            If TextEffects = True Then
                                b = New SolidBrush(Color.FromArgb(0.25 * 255, cRGB(100, 127, 205)))
                                e.Graphics.DrawString(text, font, b, New RectangleF(My.Left + 4, My.Top + 4, My.Width - 5, My.Height - 2), sf)
                            End If
                            b = New SolidBrush(cRGB(100, 127, 205))
                            e.Graphics.DrawString(text, font, b, New RectangleF(My.Left + 3, My.Top + 3, My.Width - 5, My.Height - 2), sf)
                        ElseIf s = States.Disabled Then
                            b = New SolidBrush(cRGB(117, 120, 138))
                            e.Graphics.DrawString(text, font, b, New RectangleF(My.Left + 2, My.Top + 2, My.Width - 5, My.Height - 2), sf)
                        ElseIf s = States.MouseOver Then
                            If TextEffects = True Then
                                b = New SolidBrush(Color.FromArgb(0.25 * 255, cRGB(100, 127, 205)))
                                e.Graphics.DrawString(text, font, b, New RectangleF(My.Left + 3, My.Top + 3, My.Width - 5, My.Height - 2), sf)
                            End If
                            b = New SolidBrush(cRGB(100, 127, 205))
                            e.Graphics.DrawString(text, font, b, New RectangleF(My.Left + 2, My.Top + 2, My.Width - 5, My.Height - 2), sf)
                        Else
                            If TextEffects = True Then
                                b = New SolidBrush(Color.FromArgb(0.25 * 255, cRGB(47, 74, 143)))
                                e.Graphics.DrawString(text, font, b, New RectangleF(My.Left + 3, My.Top + 3, My.Width - 5, My.Height - 2), sf)
                            End If
                            b = New SolidBrush(cRGB(47, 74, 143))
                            e.Graphics.DrawString(text, font, b, New RectangleF(My.Left + 2, My.Top + 2, My.Width - 5, My.Height - 2), sf)
                        End If

                    End Sub

                End Class
#End Region

#Region " Aqua Theme "
                Public Class Aqua

                    Private My As Rectangle
                    Private TextEffects As Boolean
                    Private tc As Color

                    Public Sub New(ByVal Owner As Rectangle)
                        MyBase.New()
                        My = Owner
                    End Sub

                    Private Sub FillPill(ByVal b As Brush, ByVal rect As Rectangle, ByVal e As PaintEventArgs)
                        If rect.Width > rect.Height Then 'Horizontal
                            e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
                            e.Graphics.FillEllipse(b, New Rectangle(rect.Left, rect.Top, rect.Height, rect.Height))
                            e.Graphics.FillEllipse(b, New Rectangle(rect.Left + rect.Width - rect.Height, rect.Top, rect.Height, rect.Height))
                            e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.Default
                            e.Graphics.FillRectangle(b, New Rectangle(rect.Left + (rect.Height / 2), rect.Top, rect.Width - rect.Height, rect.Height))
                        ElseIf rect.Width < rect.Height Then 'Vertical
                            e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
                            e.Graphics.FillEllipse(b, New Rectangle(rect.Left, rect.Top, rect.Width, rect.Width))
                            e.Graphics.FillEllipse(b, New Rectangle(rect.Left, rect.Top + rect.Height - rect.Width, rect.Width, rect.Width))
                            e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.Default
                            e.Graphics.FillRectangle(b, New Rectangle(rect.Left, rect.Top + (rect.Width / 2), rect.Width, rect.Height - rect.Width))
                        ElseIf rect.Width = rect.Height Then 'Circle
                            e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
                            e.Graphics.FillEllipse(b, rect)
                            e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.Default
                        End If
                    End Sub

                    Private Sub DrawBase(ByVal e As PaintEventArgs, ByVal state As States)
                        Dim b As Brush
                        If state = States.Disabled Then
                            b = New SolidBrush(cRGB(201, 199, 186))
                        Else
                            b = New SolidBrush(OpacityMix(Color.Black, tc, 70))
                        End If
                        FillPill(b, New Rectangle(0, 0, My.Width, My.Height), e)
                    End Sub

                    Private Sub DrawDisabledButtonBody(ByVal e As PaintEventArgs)
                        Dim sb As SolidBrush
                        sb = New SolidBrush(cRGB(245, 244, 234))
                        FillPill(sb, New Rectangle(My.Left + 1, My.Top + 1, My.Width - 2, My.Height - 2), e)
                    End Sub


                    Private Sub DrawFocusButtonBody(ByVal e As PaintEventArgs)
                        DrawButtonBody(e)
                    End Sub

                    Private Sub DrawDownButtonBody(ByVal e As PaintEventArgs)
                        Dim lgb As System.Drawing.Drawing2D.LinearGradientBrush
                        Dim cb As New System.Drawing.Drawing2D.ColorBlend

                        cb.Colors = New Color() {OpacityMix(Color.Black, tc, 25), tc, OpacityMix(Color.Black, tc, 25)}
                        cb.Positions = New Single() {0.0, 0.5, 1.0}

                        lgb = New System.Drawing.Drawing2D.LinearGradientBrush(New Point(My.Left + 1, My.Top), New Point(My.Left + 1, My.Top + My.Height - 1), OpacityMix(Color.Black, tc, 25), OpacityMix(Color.Black, tc, 25))
                        lgb.InterpolationColors = cb

                        FillPill(lgb, New Rectangle(My.Left + 1, My.Top + 1, My.Width - 2, My.Height - 2), e)
                    End Sub

                    Private Sub DrawOverButtonBody(ByVal e As PaintEventArgs)
                        Dim c1, c2, c3, c4, c5 As Color
                        Dim lgb As System.Drawing.Drawing2D.LinearGradientBrush
                        Dim cb As New System.Drawing.Drawing2D.ColorBlend
                        Dim bc As Color

                        bc = SoftLightMix(tc, Color.White, 60)

                        c1 = OpacityMix(Color.White, SoftLightMix(bc, Color.Black, 100), 40)
                        c2 = OpacityMix(Color.White, SoftLightMix(bc, cRGB(64, 64, 64), 100), 20)
                        c3 = SoftLightMix(bc, cRGB(128, 128, 128), 100)
                        c4 = SoftLightMix(bc, cRGB(192, 192, 192), 100)
                        c5 = OverlayMix(SoftLightMix(bc, Color.White, 100), Color.White, 75)

                        cb.Colors = New Color() {c1, c2, c3, c4, c5}
                        cb.Positions = New Single() {0.0, 0.25, 0.5, 0.75, 1.0}

                        lgb = New System.Drawing.Drawing2D.LinearGradientBrush(New Point(My.Left + 1, My.Top), New Point(My.Left + 1, My.Top + My.Height - 1), c1, c5)
                        lgb.InterpolationColors = cb

                        FillPill(lgb, New Rectangle(My.Left + 1, My.Top + 1, My.Width - 2, My.Height - 2), e)

                        c2 = Color.White

                        cb.Colors = New Color() {c2, c3, c4, c5}
                        cb.Positions = New Single() {0.0, 0.5, 0.75, 1.0}

                        lgb = New System.Drawing.Drawing2D.LinearGradientBrush(New Point(My.Left + 1, My.Top), New Point(My.Left + 1, My.Top + My.Height - 1), c2, c5)
                        lgb.InterpolationColors = cb

                        FillPill(lgb, New Rectangle(My.Left + 4, My.Top + 4, My.Width - 8, My.Height - 8), e)

                    End Sub

                    Private Sub DrawButtonBody(ByVal e As PaintEventArgs)
                        Dim c1, c2, c3, c4, c5 As Color
                        Dim lgb As System.Drawing.Drawing2D.LinearGradientBrush
                        Dim cb As New System.Drawing.Drawing2D.ColorBlend

                        c1 = OpacityMix(Color.White, SoftLightMix(tc, Color.Black, 100), 40)
                        c2 = OpacityMix(Color.White, SoftLightMix(tc, cRGB(64, 64, 64), 100), 20)
                        c3 = SoftLightMix(tc, cRGB(128, 128, 128), 100)
                        c4 = SoftLightMix(tc, cRGB(192, 192, 192), 100)
                        c5 = OverlayMix(SoftLightMix(tc, Color.White, 100), Color.White, 75)

                        cb.Colors = New Color() {c1, c2, c3, c4, c5}
                        cb.Positions = New Single() {0.0, 0.25, 0.5, 0.75, 1.0}

                        lgb = New System.Drawing.Drawing2D.LinearGradientBrush(New Point(My.Left + 1, My.Top), New Point(My.Left + 1, My.Top + My.Height - 1), c1, c5)
                        lgb.InterpolationColors = cb

                        FillPill(lgb, New Rectangle(My.Left + 1, My.Top + 1, My.Width - 2, My.Height - 2), e)

                        c2 = Color.White

                        cb.Colors = New Color() {c2, c3, c4, c5}
                        cb.Positions = New Single() {0.0, 0.5, 0.75, 1.0}

                        lgb = New System.Drawing.Drawing2D.LinearGradientBrush(New Point(My.Left + 1, My.Top), New Point(My.Left + 1, My.Top + My.Height - 1), c2, c5)
                        lgb.InterpolationColors = cb

                        FillPill(lgb, New Rectangle(My.Left + 4, My.Top + 4, My.Width - 8, My.Height - 8), e)

                    End Sub

                    Public Sub DrawTheme(ByVal e As PaintEventArgs, ByVal state As Theme.States, ByVal textfx As Boolean, ByVal themecolor As Color)
                        tc = themecolor

                        DrawBase(e, state)

                        TextEffects = textfx
                        If state = Theme.States.Normal Then
                            DrawButtonBody(e)
                        ElseIf state = Theme.States.MouseOver Then
                            DrawOverButtonBody(e)
                        ElseIf state = States.Focused Then
                            DrawFocusButtonBody(e)
                        ElseIf state = States.MouseDown Then
                            DrawDownButtonBody(e)
                        ElseIf state = States.Disabled Then
                            DrawDisabledButtonBody(e)
                        End If
                    End Sub

                    Public Sub DrawText(ByVal e As PaintEventArgs, ByVal s As States, ByVal text As String, ByVal font As Font, ByVal forecolor As Color)

                        Dim sf As StringFormat
                        Dim sz As SizeF
                        Dim b As Brush

                        sf = New StringFormat
                        sf.LineAlignment = StringAlignment.Center
                        sf.Alignment = StringAlignment.Center
                        sf.Trimming = StringTrimming.EllipsisCharacter

                        sz = e.Graphics.MeasureString(text, font, New SizeF(My.Width - 2, My.Height - 2), sf)
                        text = text & "                    "

                        If s = States.MouseDown Then
                            If TextEffects = True Then
                                b = New SolidBrush(Color.FromArgb(0.35 * 255, Color.Black))
                                e.Graphics.DrawString(text, font, b, New RectangleF(3, 3, My.Width - 2, My.Height - 2), sf)
                            End If
                            b = New SolidBrush(forecolor)
                            e.Graphics.DrawString(text, font, b, New RectangleF(2, 2, My.Width - 2, My.Height - 2), sf)
                        ElseIf s = States.Disabled Then
                            b = New SolidBrush(Color.FromKnownColor(KnownColor.GrayText))
                            e.Graphics.DrawString(text, font, b, New RectangleF(1, 1, My.Width - 2, My.Height - 2), sf)
                        Else
                            If TextEffects = True Then
                                b = New SolidBrush(Color.FromArgb(0.35 * 255, Color.Black))
                                e.Graphics.DrawString(text, font, b, New RectangleF(2, 2, My.Width - 2, My.Height - 2), sf)
                            End If
                            b = New SolidBrush(forecolor)
                            e.Graphics.DrawString(text, font, b, New RectangleF(1, 1, My.Width - 2, My.Height - 2), sf)
                        End If

                    End Sub
                End Class
#End Region

#Region " 3D Hover Theme "
                Public Class Hover3D
                    Private My As Rectangle
                    Private TextEffects As Boolean
                    Private isFocused As Boolean

                    Public Sub New(ByVal Owner As Rectangle)
                        MyBase.New()
                        My = Owner
                    End Sub

                    Private Sub DrawFocusButtonBody(ByVal e As PaintEventArgs)
                        Dim p As Pen
                        p = New Pen(Color.FromArgb(0.5 * 255, Color.Black))
                        p.DashStyle = Drawing2D.DashStyle.Dot
                        e.Graphics.DrawRectangle(p, My.Left + 3, My.Top + 3, My.Width - 7, My.Height - 7)
                        e.Graphics.DrawLine(New Pen(Color.FromArgb(0.25 * 255, Color.White)), My.Left, My.Top, My.Left + My.Width, My.Top)
                        e.Graphics.DrawLine(New Pen(Color.FromArgb(0.25 * 255, Color.White)), My.Left, My.Top, My.Left, My.Top + My.Height)
                        e.Graphics.DrawLine(New Pen(Color.FromArgb(0.1 * 255, Color.Black)), My.Left, My.Top + My.Height - 1, My.Left + My.Width, My.Top + My.Height - 1)
                        e.Graphics.DrawLine(New Pen(Color.FromArgb(0.1 * 255, Color.Black)), My.Left + My.Width - 1, My.Top, My.Left + My.Width - 1, My.Top + My.Height)
                    End Sub

                    Private Sub DrawDownButtonBody(ByVal e As PaintEventArgs)
                        e.Graphics.DrawLine(New Pen(Color.FromArgb(0.15 * 255, Color.Black)), My.Left, My.Top, My.Left + My.Width, My.Top)
                        e.Graphics.DrawLine(New Pen(Color.FromArgb(0.15 * 255, Color.Black)), My.Left, My.Top, My.Left, My.Top + My.Height)
                        e.Graphics.DrawLine(New Pen(Color.FromArgb(0.65 * 255, Color.White)), My.Left, My.Top + My.Height - 1, My.Left + My.Width, My.Top + My.Height - 1)
                        e.Graphics.DrawLine(New Pen(Color.FromArgb(0.65 * 255, Color.White)), My.Left + My.Width - 1, My.Top, My.Left + My.Width - 1, My.Top + My.Height)
                        e.Graphics.FillRectangle(New SolidBrush(Color.FromArgb(0.025 * 255, Color.Black)), My.Left, My.Top, My.Width, My.Height)

                        If isFocused = True Then
                            Dim p As Pen
                            p = New Pen(Color.FromArgb(0.5 * 255, Color.Black))
                            p.DashStyle = Drawing2D.DashStyle.Dot
                            e.Graphics.DrawRectangle(p, My.Left + 3, My.Top + 3, My.Width - 7, My.Height - 7)
                        End If
                    End Sub

                    Private Sub DrawOverButtonBody(ByVal e As PaintEventArgs)
                        If isFocused = True Then
                            Dim p As Pen
                            p = New Pen(Color.FromArgb(0.5 * 255, Color.Black))
                            p.DashStyle = Drawing2D.DashStyle.Dot
                            e.Graphics.DrawRectangle(p, My.Left + 3, My.Top + 3, My.Width - 7, My.Height - 7)
                        End If
                        e.Graphics.DrawLine(New Pen(Color.FromArgb(0.75 * 255, Color.White)), My.Left, My.Top, My.Left + My.Width, My.Top)
                        e.Graphics.DrawLine(New Pen(Color.FromArgb(0.75 * 255, Color.White)), My.Left, My.Top, My.Left, My.Top + My.Height)
                        e.Graphics.DrawLine(New Pen(Color.FromArgb(0.35 * 255, Color.Black)), My.Left, My.Top + My.Height - 1, My.Left + My.Width, My.Top + My.Height - 1)
                        e.Graphics.DrawLine(New Pen(Color.FromArgb(0.35 * 255, Color.Black)), My.Left + My.Width - 1, My.Top, My.Left + My.Width - 1, My.Top + My.Height)
                    End Sub


                    Public Sub DrawTheme(ByVal e As PaintEventArgs, ByVal state As Theme.States, ByVal textfx As Boolean, ByVal focused As Boolean)
                        TextEffects = textfx
                        isFocused = focused
                        If state = Theme.States.MouseOver Then
                            DrawOverButtonBody(e)
                        ElseIf state = States.Focused Then
                            DrawFocusButtonBody(e)
                        ElseIf state = States.MouseDown Then
                            DrawDownButtonBody(e)
                        End If
                    End Sub

                    Public Sub DrawText(ByVal e As PaintEventArgs, ByVal s As States, ByVal text As String, ByVal font As Font, ByVal forecolor As Color)

                        Dim sf As StringFormat
                        Dim sz As SizeF
                        Dim b As Brush

                        sf = New StringFormat
                        sf.LineAlignment = StringAlignment.Center
                        sf.Alignment = StringAlignment.Center
                        sf.Trimming = StringTrimming.EllipsisCharacter

                        sz = e.Graphics.MeasureString(text, font, New SizeF(My.Width - 2, My.Height - 2), sf)
                        text = text & "                    "

                        If s = States.MouseDown Then
                            If TextEffects = True Then
                                b = New SolidBrush(Color.FromArgb(0.15 * 255, Color.Black))
                                e.Graphics.DrawString(text, font, b, New RectangleF(3, 3, My.Width - 2, My.Height - 2), sf)
                            End If
                            b = New SolidBrush(forecolor)
                            e.Graphics.DrawString(text, font, b, New RectangleF(2, 2, My.Width - 2, My.Height - 2), sf)
                        ElseIf s = States.Disabled Then
                            b = New SolidBrush(Color.FromKnownColor(KnownColor.GrayText))
                            e.Graphics.DrawString(text, font, b, New RectangleF(1, 1, My.Width - 2, My.Height - 2), sf)
                        Else
                            If TextEffects = True Then
                                b = New SolidBrush(Color.FromArgb(0.15 * 255, Color.Black))
                                e.Graphics.DrawString(text, font, b, New RectangleF(2, 2, My.Width - 2, My.Height - 2), sf)
                            End If
                            b = New SolidBrush(forecolor)
                            e.Graphics.DrawString(text, font, b, New RectangleF(1, 1, My.Width - 2, My.Height - 2), sf)
                        End If

                    End Sub
                End Class
#End Region

#Region " Office XP Theme "
                Public Class OfficeXP
                    Private My As Rectangle
                    Private TextEffects, isFocused As Boolean
                    Private tc As Color

                    Public Sub New(ByVal Owner As Rectangle)
                        MyBase.New()
                        My = Owner
                    End Sub

                    Private Sub DrawFocusButtonBody(ByVal e As PaintEventArgs)
                        Dim p As Pen
                        p = New Pen(Color.FromArgb(0.5 * 255, Color.Black))
                        p.DashStyle = Drawing2D.DashStyle.Dot
                        e.Graphics.DrawRectangle(p, My.Left + 3, My.Top + 3, My.Width - 7, My.Height - 7)
                    End Sub

                    Private Sub DrawDownButtonBody(ByVal e As PaintEventArgs)
                        Dim b As SolidBrush
                        b = New SolidBrush(tc)
                        e.Graphics.FillRectangle(b, My.Left, My.Top, My.Width, My.Height)

                        b = New SolidBrush(OverlayMix(OpacityMix(Color.White, tc, 75), Color.Black, 50))
                        e.Graphics.FillRectangle(b, My.Left + 1, My.Top + 1, My.Width - 2, My.Height - 2)

                        If isFocused = True Then
                            Dim p As Pen
                            p = New Pen(Color.FromArgb(0.5 * 255, Color.Black))
                            p.DashStyle = Drawing2D.DashStyle.Dot
                            e.Graphics.DrawRectangle(p, My.Left + 3, My.Top + 3, My.Width - 7, My.Height - 7)
                        End If
                    End Sub

                    Private Sub DrawOverButtonBody(ByVal e As PaintEventArgs)
                        Dim b As SolidBrush
                        b = New SolidBrush(tc)
                        e.Graphics.FillRectangle(b, My.Left, My.Top, My.Width, My.Height)

                        b = New SolidBrush(Color.FromArgb(0.75 * 255, Color.White))
                        e.Graphics.FillRectangle(b, My.Left + 1, My.Top + 1, My.Width - 2, My.Height - 2)

                        If isFocused = True Then
                            Dim p As Pen
                            p = New Pen(Color.FromArgb(0.5 * 255, Color.Black))
                            p.DashStyle = Drawing2D.DashStyle.Dot
                            e.Graphics.DrawRectangle(p, My.Left + 3, My.Top + 3, My.Width - 7, My.Height - 7)
                        End If
                    End Sub


                    Public Sub DrawTheme(ByVal e As PaintEventArgs, ByVal state As Theme.States, ByVal textfx As Boolean, ByVal themecolor As Color, ByVal focused As Boolean)
                        TextEffects = textfx
                        isFocused = focused
                        tc = themecolor
                        If state = Theme.States.MouseOver Then
                            DrawOverButtonBody(e)
                        ElseIf state = States.Focused Then
                            DrawFocusButtonBody(e)
                        ElseIf state = States.MouseDown Then
                            DrawDownButtonBody(e)
                        End If
                    End Sub

                    Public Sub DrawText(ByVal e As PaintEventArgs, ByVal s As States, ByVal text As String, ByVal font As Font, ByVal forecolor As Color)

                        Dim sf As StringFormat
                        Dim sz As SizeF
                        Dim b As Brush

                        sf = New StringFormat
                        sf.LineAlignment = StringAlignment.Center
                        sf.Alignment = StringAlignment.Center
                        sf.Trimming = StringTrimming.EllipsisCharacter

                        sz = e.Graphics.MeasureString(text, font, New SizeF(My.Width - 2, My.Height - 2), sf)
                        text = text & "                    "

                        If s = States.MouseDown Then
                            If TextEffects = True Then
                                b = New SolidBrush(Color.FromArgb(0.15 * 255, Color.Black))
                                e.Graphics.DrawString(text, font, b, New RectangleF(3, 3, My.Width - 2, My.Height - 2), sf)
                            End If
                            b = New SolidBrush(forecolor)
                            e.Graphics.DrawString(text, font, b, New RectangleF(2, 2, My.Width - 2, My.Height - 2), sf)
                        ElseIf s = States.Disabled Then
                            b = New SolidBrush(Color.FromKnownColor(KnownColor.GrayText))
                            e.Graphics.DrawString(text, font, b, New RectangleF(1, 1, My.Width - 2, My.Height - 2), sf)
                        Else
                            If TextEffects = True Then
                                b = New SolidBrush(Color.FromArgb(0.15 * 255, Color.Black))
                                e.Graphics.DrawString(text, font, b, New RectangleF(2, 2, My.Width - 2, My.Height - 2), sf)
                            End If
                            b = New SolidBrush(forecolor)
                            e.Graphics.DrawString(text, font, b, New RectangleF(1, 1, My.Width - 2, My.Height - 2), sf)
                        End If

                    End Sub
                End Class
#End Region

#Region " Office 2003 Theme "
                Public Class Office2003
                    Private My As Rectangle
                    Private TextEffects, isFocused As Boolean
                    Private tc As Color

                    Public Sub New(ByVal Owner As Rectangle)
                        MyBase.New()
                        My = Owner
                    End Sub

                    Private Sub DrawFocusButtonBody(ByVal e As PaintEventArgs)

                        Dim c1, c2, c3 As Color
                        Dim lgb As System.Drawing.Drawing2D.LinearGradientBrush
                        Dim cb As New System.Drawing.Drawing2D.ColorBlend

                        c1 = OverlayMix(OpacityMix(Color.White, tc, 55), Color.White, 60)
                        c2 = OpacityMix(Color.White, tc, 55)
                        c3 = OpacityMix(Color.Black, c2, 15)

                        cb.Colors = New Color() {c1, c2, c3}
                        cb.Positions = New Single() {0.0, 0.5, 1.0}

                        lgb = New System.Drawing.Drawing2D.LinearGradientBrush(New Point(0, My.Top), New Point(0, My.Top + My.Height), c1, c3)
                        lgb.InterpolationColors = cb

                        e.Graphics.FillRectangle(lgb, My.Left, My.Top, My.Width, My.Height)

                        Dim p As Pen
                        p = New Pen(Color.FromArgb(0.5 * 255, Color.Black))
                        p.DashStyle = Drawing2D.DashStyle.Dot
                        e.Graphics.DrawRectangle(p, My.Left + 3, My.Top + 3, My.Width - 7, My.Height - 7)

                        e.Graphics.DrawLine(New Pen(Color.FromArgb(0.25 * 255, Color.White)), My.Left, My.Top, My.Left + My.Width, My.Top)
                        e.Graphics.DrawLine(New Pen(Color.FromArgb(0.25 * 255, Color.White)), My.Left, My.Top, My.Left, My.Top + My.Height)
                        e.Graphics.DrawLine(New Pen(Color.FromArgb(0.1 * 255, Color.Black)), My.Left, My.Top + My.Height - 1, My.Left + My.Width, My.Top + My.Height - 1)
                        e.Graphics.DrawLine(New Pen(Color.FromArgb(0.1 * 255, Color.Black)), My.Left + My.Width - 1, My.Top, My.Left + My.Width - 1, My.Top + My.Height)
                    End Sub

                    Private Sub DrawDownButtonBody(ByVal e As PaintEventArgs)
                        Dim c1, c2 As Color
                        Dim lgb As System.Drawing.Drawing2D.LinearGradientBrush
                        Dim bc As Color

                        bc = OpacityMix(Color.White, OverlayMix(InvertColor(tc), Color.White, 10), 50)

                        c1 = OverlayMix(bc, Color.Black, 35)
                        c2 = OverlayMix(bc, Color.White, 25)

                        lgb = New System.Drawing.Drawing2D.LinearGradientBrush(New Point(0, My.Top), New Point(0, My.Top + My.Height), c1, c2)

                        e.Graphics.FillRectangle(lgb, My.Left, My.Top, My.Width, My.Height)

                        e.Graphics.DrawRectangle(New Pen(OpacityMix(Color.Black, OverlayMix(tc, Color.Black, 75), 50)), My.Left, My.Top, My.Width - 1, My.Height - 1)

                        If isFocused = True Then
                            Dim p As Pen
                            p = New Pen(Color.FromArgb(0.5 * 255, Color.Black))
                            p.DashStyle = Drawing2D.DashStyle.Dot
                            e.Graphics.DrawRectangle(p, My.Left + 3, My.Top + 3, My.Width - 7, My.Height - 7)
                        End If
                    End Sub

                    Private Sub DrawOverButtonBody(ByVal e As PaintEventArgs)
                        Dim c1, c2 As Color
                        Dim lgb As System.Drawing.Drawing2D.LinearGradientBrush
                        Dim bc As Color

                        bc = OpacityMix(Color.White, OverlayMix(InvertColor(tc), Color.White, 25), 50)

                        c1 = OverlayMix(bc, Color.White, 50)
                        c2 = OverlayMix(bc, Color.Black, 15)

                        lgb = New System.Drawing.Drawing2D.LinearGradientBrush(New Point(0, My.Top), New Point(0, My.Top + My.Height), c1, c2)

                        e.Graphics.FillRectangle(lgb, My.Left, My.Top, My.Width, My.Height)

                        e.Graphics.DrawRectangle(New Pen(OpacityMix(Color.Black, OverlayMix(tc, Color.Black, 75), 50)), My.Left, My.Top, My.Width - 1, My.Height - 1)

                        If isFocused = True Then
                            Dim p As Pen
                            p = New Pen(Color.FromArgb(0.5 * 255, Color.Black))
                            p.DashStyle = Drawing2D.DashStyle.Dot
                            e.Graphics.DrawRectangle(p, My.Left + 3, My.Top + 3, My.Width - 7, My.Height - 7)
                        End If
                    End Sub

                    Private Sub DrawButtonBody(ByVal e As PaintEventArgs)
                        Dim c1, c2, c3 As Color
                        Dim lgb As System.Drawing.Drawing2D.LinearGradientBrush
                        Dim cb As New System.Drawing.Drawing2D.ColorBlend

                        c1 = OverlayMix(OpacityMix(Color.White, tc, 55), Color.White, 60)
                        c2 = OpacityMix(Color.White, tc, 55)
                        c3 = OpacityMix(Color.Black, c2, 15)

                        cb.Colors = New Color() {c1, c2, c3}
                        cb.Positions = New Single() {0.0, 0.5, 1.0}

                        lgb = New System.Drawing.Drawing2D.LinearGradientBrush(New Point(0, My.Top), New Point(0, My.Top + My.Height), c1, c3)
                        lgb.InterpolationColors = cb

                        e.Graphics.FillRectangle(lgb, My.Left, My.Top, My.Width, My.Height)

                        e.Graphics.DrawLine(New Pen(Color.FromArgb(0.35 * 255, Color.White)), My.Left, My.Top, My.Left + My.Width, My.Top)
                        e.Graphics.DrawLine(New Pen(Color.FromArgb(0.35 * 255, Color.White)), My.Left, My.Top, My.Left, My.Top + My.Height)
                        e.Graphics.DrawLine(New Pen(Color.FromArgb(0.1 * 255, Color.Black)), My.Left, My.Top + My.Height - 1, My.Left + My.Width, My.Top + My.Height - 1)
                        e.Graphics.DrawLine(New Pen(Color.FromArgb(0.1 * 255, Color.Black)), My.Left + My.Width - 1, My.Top, My.Left + My.Width - 1, My.Top + My.Height)
                    End Sub

                    Public Sub DrawTheme(ByVal e As PaintEventArgs, ByVal state As Theme.States, ByVal textfx As Boolean, ByVal themecolor As Color, ByVal focused As Boolean)
                        TextEffects = textfx
                        isFocused = focused
                        tc = themecolor
                        If state = States.Normal Then
                            DrawButtonBody(e)
                        ElseIf state = Theme.States.MouseOver Then
                            DrawOverButtonBody(e)
                        ElseIf state = States.Focused Then
                            DrawFocusButtonBody(e)
                        ElseIf state = States.MouseDown Then
                            DrawDownButtonBody(e)
                        ElseIf state = States.Disabled Then
                            DrawButtonBody(e)
                        End If
                    End Sub

                    Public Sub DrawText(ByVal e As PaintEventArgs, ByVal s As States, ByVal text As String, ByVal font As Font, ByVal forecolor As Color)

                        Dim sf As StringFormat
                        Dim sz As SizeF
                        Dim b As Brush

                        sf = New StringFormat
                        sf.LineAlignment = StringAlignment.Center
                        sf.Alignment = StringAlignment.Center
                        sf.Trimming = StringTrimming.EllipsisCharacter

                        sz = e.Graphics.MeasureString(text, font, New SizeF(My.Width - 2, My.Height - 2), sf)
                        text = text & "                    "

                        If s = States.MouseDown Then
                            If TextEffects = True Then
                                b = New SolidBrush(Color.FromArgb(0.15 * 255, Color.Black))
                                e.Graphics.DrawString(text, font, b, New RectangleF(3, 3, My.Width - 2, My.Height - 2), sf)
                            End If
                            b = New SolidBrush(forecolor)
                            e.Graphics.DrawString(text, font, b, New RectangleF(2, 2, My.Width - 2, My.Height - 2), sf)
                        ElseIf s = States.Disabled Then
                            b = New SolidBrush(Color.FromKnownColor(KnownColor.GrayText))
                            e.Graphics.DrawString(text, font, b, New RectangleF(1, 1, My.Width - 2, My.Height - 2), sf)
                        Else
                            If TextEffects = True Then
                                b = New SolidBrush(Color.FromArgb(0.15 * 255, Color.Black))
                                e.Graphics.DrawString(text, font, b, New RectangleF(2, 2, My.Width - 2, My.Height - 2), sf)
                            End If
                            b = New SolidBrush(forecolor)
                            e.Graphics.DrawString(text, font, b, New RectangleF(1, 1, My.Width - 2, My.Height - 2), sf)
                        End If

                    End Sub
                End Class
#End Region

#Region " Macintosh Theme "
                Public Class Macintosh
                    Private My As Rectangle
                    Private TextEffects As Boolean
                    Private tc As Color

                    Public Sub New(ByVal Owner As Rectangle)
                        MyBase.New()
                        My = Owner
                    End Sub

                    Private Sub FillPlate(ByVal b As Brush, ByVal rect As Rectangle, ByVal e As PaintEventArgs)
                        e.Graphics.FillRectangle(b, rect.Left, rect.Top + 1, rect.Width, rect.Height - 2)
                        e.Graphics.FillRectangle(b, rect.Left + 1, rect.Top, rect.Width - 2, rect.Height)
                    End Sub

                    Private Sub DrawBase(ByVal e As PaintEventArgs)
                        Dim b As SolidBrush
                        b = New SolidBrush(OpacityMix(Color.Black, SoftLightMix(tc, Color.Black, 100), 50))
                        FillPlate(b, New Rectangle(My.Left, My.Top, My.Width, My.Height), e)
                    End Sub

                    Private Sub DrawDownButtonBody(ByVal e As PaintEventArgs)
                        Dim sb As SolidBrush
                        Dim bc, c As Color

                        bc = OpacityMix(Color.Black, SoftLightMix(tc, Color.Black, 100), 30)

                        sb = New SolidBrush(SoftLightMix(bc, Color.Black, 75))
                        FillPlate(sb, New Rectangle(My.Left + 1, My.Top + 1, My.Width - 2, My.Height - 2), e)

                        sb = New SolidBrush(bc)
                        FillPlate(sb, New Rectangle(My.Left + 2, My.Top + 2, My.Width - 3, My.Height - 3), e)

                        sb = New SolidBrush(OpacityMix(Color.White, bc, 15))
                        FillPlate(sb, New Rectangle(My.Left + 2, My.Top + 2, My.Width - 4, My.Height - 4), e)

                        sb = New SolidBrush(bc)
                        FillPlate(sb, New Rectangle(My.Left + 2, My.Top + 2, My.Width - 5, My.Height - 5), e)

                        c = SoftLightMix(bc, Color.Black, 25)
                        e.Graphics.DrawLine(New Pen(c), My.Left + 2, My.Top + 3, My.Left + 2, My.Top + My.Height - 2)
                        e.Graphics.DrawLine(New Pen(c), My.Left + 3, My.Top + 2, My.Left + My.Width - 2, My.Top + 2)
                    End Sub

                    Private Sub DrawButtonBody(ByVal e As PaintEventArgs)
                        Dim sb As SolidBrush
                        Dim c As Color

                        sb = New SolidBrush(OpacityMix(Color.Black, SoftLightMix(tc, Color.Black, 100), 30))
                        FillPlate(sb, New Rectangle(My.Left + 1, My.Top + 1, My.Width - 2, My.Height - 2), e)

                        sb = New SolidBrush(tc)
                        FillPlate(sb, New Rectangle(My.Left + 1, My.Top + 1, My.Width - 3, My.Height - 3), e)

                        sb = New SolidBrush(Color.White)
                        FillPlate(sb, New Rectangle(My.Left + 2, My.Top + 2, My.Width - 4, My.Height - 4), e)

                        sb = New SolidBrush(tc)
                        FillPlate(sb, New Rectangle(My.Left + 3, My.Top + 3, My.Width - 5, My.Height - 5), e)

                        c = SoftLightMix(tc, Color.Black, 50)
                        e.Graphics.DrawLine(New Pen(c), My.Left + 1, My.Top + My.Height - 3, My.Left + My.Width - 4, My.Top + My.Height - 3)
                        e.Graphics.DrawLine(New Pen(c), My.Left + My.Width - 3, My.Top + 1, My.Left + My.Width - 3, My.Top + My.Height - 4)
                    End Sub

                    Public Sub DrawTheme(ByVal e As PaintEventArgs, ByVal state As Theme.States, ByVal textfx As Boolean, ByVal themecolor As Color)
                        TextEffects = textfx
                        tc = themecolor

                        DrawBase(e)

                        If state = States.Normal Then
                            DrawButtonBody(e)
                        ElseIf state = Theme.States.MouseOver Then
                            DrawButtonBody(e)
                        ElseIf state = States.Focused Then
                            DrawButtonBody(e)
                        ElseIf state = States.MouseDown Then
                            DrawDownButtonBody(e)
                        ElseIf state = States.Disabled Then
                            DrawButtonBody(e)
                        End If
                    End Sub

                    Public Sub DrawText(ByVal e As PaintEventArgs, ByVal s As States, ByVal text As String, ByVal font As Font, ByVal forecolor As Color)

                        Dim sf As StringFormat
                        Dim sz As SizeF
                        Dim b As Brush

                        sf = New StringFormat
                        sf.LineAlignment = StringAlignment.Center
                        sf.Alignment = StringAlignment.Center
                        sf.Trimming = StringTrimming.EllipsisCharacter

                        sz = e.Graphics.MeasureString(text, font, New SizeF(My.Width - 2, My.Height - 2), sf)
                        text = text & "                    "

                        If s = States.MouseDown Then
                            If TextEffects = True Then
                                b = New SolidBrush(Color.FromArgb(0.5 * 255, Color.White))
                                e.Graphics.DrawString(text, font, b, New RectangleF(3, 3, My.Width - 2, My.Height - 2), sf)
                            End If
                            b = New SolidBrush(forecolor)
                            e.Graphics.DrawString(text, font, b, New RectangleF(2, 2, My.Width - 2, My.Height - 2), sf)
                        ElseIf s = States.Disabled Then
                            b = New SolidBrush(Color.FromArgb(0.5 * 255, Color.White))
                            e.Graphics.DrawString(text, font, b, New RectangleF(3, 3, My.Width - 2, My.Height - 2), sf)
                            b = New SolidBrush(Color.FromKnownColor(KnownColor.GrayText))
                            e.Graphics.DrawString(text, font, b, New RectangleF(1, 1, My.Width - 2, My.Height - 2), sf)
                        Else
                            If TextEffects = True Then
                                b = New SolidBrush(Color.FromArgb(0.5 * 255, Color.White))
                                e.Graphics.DrawString(text, font, b, New RectangleF(2, 2, My.Width - 2, My.Height - 2), sf)
                            End If
                            b = New SolidBrush(forecolor)
                            e.Graphics.DrawString(text, font, b, New RectangleF(1, 1, My.Width - 2, My.Height - 2), sf)
                        End If

                    End Sub
                End Class
#End Region

            End Namespace

#End Region

#Region " Color functions "




            'This is a very powerful module that is capable of carrying out operations
            'related to colors (including color mixing, blend modes, etc.)

            'The blend modes are based on the blend modes of Adobe Photoshop but may
            'not identically reproduce the effects.

            'The color converter allows for conversion of colors between color spaces
            'and from various color formats

            Module modColorFunctions

#Region " Basic Functions "

                'This function provides an easy way to convert RGB values to a color.
                Public Function cRGB(ByVal r As Integer, ByVal g As Integer, ByVal b As Integer) As Color
                    If r > 255 Then r = 255
                    If r < 0 Then r = 0
                    If g > 255 Then g = 255
                    If g < 0 Then g = 0
                    If b > 255 Then b = 255
                    If b < 0 Then b = 0
                    cRGB = ColorTranslator.FromWin32(RGB(r, g, b))
                End Function

                Public Function cHSB(ByVal hue As Single, ByVal saturation As Single, ByVal brightness As Single) As Color
                    If hue < 0 Then hue = 0
                    If hue > 359 Then hue = hue - 360
                    If saturation < 0 Then saturation = 0
                    If saturation > 1 Then saturation = 1
                    If brightness < 0 Then brightness = 0
                    If brightness > 1 Then brightness = 1

                    Debug.Write(vbCrLf & hue & ", " & saturation & ", " & brightness)

                    Dim v1, v2, vh As Single
                    Dim r, g, b As Integer

                    If saturation = 0 Then
                        cHSB = cRGB(brightness * 255, brightness * 255, brightness * 255)
                    Else
                        If brightness < 0.5 Then
                            v2 = brightness * (1 + saturation)
                        Else
                            v2 = (brightness + saturation) - (brightness * saturation)
                        End If

                        v1 = 2 * brightness - v2
                        vh = hue / 360

                        Debug.Write(vbCrLf & v1 & ", " & v2)

                        Debug.Write(vbCrLf & Hue2RGB(v1, v2, hue + (1 / 3)) & ", " & Hue2RGB(v1, v2, hue) & ", " & Hue2RGB(v1, v2, hue - (1 / 3)))

                        r = 255 * Hue2RGB(v1, v2, vh + (1 / 3))
                        g = 255 * Hue2RGB(v1, v2, vh)
                        b = 255 * Hue2RGB(v1, v2, vh - (1 / 3))

                        cHSB = cRGB(r, g, b)
                    End If
                End Function

                Private Function Hue2RGB(ByVal v1 As Single, ByVal v2 As Single, ByVal vH As Single) As Single
                    If (vH < 0) Then vH += 1
                    If (vH > 1) Then vH -= 1

                    If ((6 * vH) < 1) Then
                        Hue2RGB = (v1 + (v2 - v1) * 6 * vH)
                    ElseIf ((2 * vH) < 1) Then
                        Hue2RGB = (v2)
                    ElseIf ((3 * vH) < 2) Then
                        Hue2RGB = (v1 + (v2 - v1) * ((2 / 3) - vH) * 6)
                    Else
                        Hue2RGB = (v1)
                    End If
                End Function

                Public Function OpacityMix(ByVal BlendColor As Color, ByVal BaseColor As Color, ByVal Opacity As Integer) As Color
                    Dim r1, g1, b1 As Integer
                    Dim r2, g2, b2 As Integer
                    Dim r3, g3, b3 As Integer

                    r1 = BlendColor.R
                    g1 = BlendColor.G
                    b1 = BlendColor.B

                    r2 = BaseColor.R
                    g2 = BaseColor.G
                    b2 = BaseColor.B

                    r3 = ((r1 * (Opacity / 100)) + (r2 * (1 - (Opacity / 100))))
                    g3 = ((g1 * (Opacity / 100)) + (g2 * (1 - (Opacity / 100))))
                    b3 = ((b1 * (Opacity / 100)) + (b2 * (1 - (Opacity / 100))))

                    OpacityMix = cRGB(r3, g3, b3)
                End Function

                Public Function InvertColor(ByVal c As Color) As Color
                    Dim r1, g1, b1 As Integer
                    Dim r2, g2, b2 As Integer

                    r1 = c.R
                    g1 = c.G
                    b1 = c.B

                    r2 = 255 - r1
                    g2 = 255 - g1
                    b2 = 255 - b1

                    InvertColor = cRGB(r2, g2, b2)
                End Function

#End Region

#Region " Blend Modes "

                Private Function ScreenMix(ByVal BaseColor As Color, ByVal BlendColor As Color, ByVal Opacity As Integer) As Color
                    Dim r1, g1, b1 As Integer
                    Dim r2, g2, b2 As Integer
                    Dim r3, g3, b3 As Integer

                    r1 = BaseColor.R
                    g1 = BaseColor.G
                    b1 = BaseColor.B

                    r2 = BlendColor.R
                    g2 = BlendColor.G
                    b2 = BlendColor.B

                    r3 = (1 - ((1 - (r1 / 255)) * (1 - (r2 / 255)))) * 255
                    g3 = (1 - ((1 - (g1 / 255)) * (1 - (g2 / 255)))) * 255
                    b3 = (1 - ((1 - (b1 / 255)) * (1 - (b2 / 255)))) * 255

                    ScreenMix = OpacityMix(cRGB(r3, g3, b3), BaseColor, Opacity)
                End Function

                Public Function MultiplyMix(ByVal BaseColor As Color, ByVal BlendColor As Color, ByVal Opacity As Integer) As Color
                    Dim r1, g1, b1 As Integer
                    Dim r2, g2, b2 As Integer
                    Dim r3, g3, b3 As Integer

                    r1 = BaseColor.R
                    g1 = BaseColor.G
                    b1 = BaseColor.B

                    r2 = BlendColor.R
                    g2 = BlendColor.G
                    b2 = BlendColor.B

                    r3 = r1 * r2 / 255
                    g3 = g1 * g2 / 255
                    b3 = b1 * b2 / 255

                    MultiplyMix = OpacityMix(cRGB(r3, g3, b3), BaseColor, Opacity)
                End Function

                Public Function SoftLightMix(ByVal BaseColor As Color, ByVal BlendColor As Color, ByVal Opacity As Integer) As Color
                    Dim r1, g1, b1 As Integer
                    Dim r2, g2, b2 As Integer
                    Dim r3, g3, b3 As Integer

                    r1 = BaseColor.R
                    g1 = BaseColor.G
                    b1 = BaseColor.B

                    r2 = BlendColor.R
                    g2 = BlendColor.G
                    b2 = BlendColor.B

                    r3 = SoftLightMath(r1, r2)
                    g3 = SoftLightMath(g1, g2)
                    b3 = SoftLightMath(b1, b2)

                    SoftLightMix = OpacityMix(cRGB(r3, g3, b3), BaseColor, Opacity)
                End Function

                Public Function OverlayMix(ByVal BaseColor As Color, ByVal BlendColor As Color, ByVal opacity As Integer) As Color
                    Dim r1, g1, b1 As Integer
                    Dim r2, g2, b2 As Integer
                    Dim r3, g3, b3 As Integer

                    r1 = BaseColor.R
                    g1 = BaseColor.G
                    b1 = BaseColor.B

                    r2 = BlendColor.R
                    g2 = BlendColor.G
                    b2 = BlendColor.B

                    r3 = OverlayMath(BaseColor.R, BlendColor.R)
                    g3 = OverlayMath(BaseColor.G, BlendColor.G)
                    b3 = OverlayMath(BaseColor.B, BlendColor.B)

                    OverlayMix = OpacityMix(cRGB(r3, g3, b3), BaseColor, opacity)
                End Function

#End Region

#Region " Blend Mode Mathematics "

                Private Function SoftLightMath(ByVal base As Integer, ByVal blend As Integer) As Integer
                    Dim dbase As Single
                    Dim dblend As Single

                    dbase = base / 255
                    dblend = blend / 255

                    If dblend < 0.5 Then
                        SoftLightMath = ((2 * dbase * dblend) + (dbase ^ 2) * (1 - (2 * dblend))) * 255
                    Else
                        SoftLightMath = ((Math.Sqrt(dbase) * (2 * dblend - 1)) + ((2 * dbase) * (1 - dblend))) * 255
                    End If
                End Function

                Public Function OverlayMath(ByVal base As Integer, ByVal blend As Integer) As Integer
                    Dim dbase, dblend As Double

                    dbase = base / 255
                    dblend = blend / 255

                    If dbase < 0.5 Then
                        OverlayMath = (2 * dbase * dblend) * 255
                    Else
                        OverlayMath = (1 - (2 * (1 - dbase) * (1 - dblend))) * 255
                    End If
                End Function

#End Region

            End Module

#Region " Color Converter "

            Public Class ColorConversion

                Private _Color As Color

#Region " New "
                Public Sub New(ByVal inputColor As Color)
                    _Color = inputColor
                End Sub

                Public Sub New(ByVal r As Byte, ByVal g As Byte, ByVal b As Byte)
                    _Color = cRGB(r, g, b)
                End Sub

                Public Sub New(ByVal Hue As Single, ByVal Saturation As Single, ByVal Brightness As Single, ByVal hsb As Boolean)
                    _Color = cHSB(Hue, Saturation, Brightness)
                End Sub

                Public Sub New(ByVal Hexadecimal As String)
                    _Color = ColorTranslator.FromHtml(Hexadecimal)
                End Sub

                Public Sub New(ByVal Win32 As Integer)
                    _Color = ColorTranslator.FromWin32(Win32)
                End Sub

#End Region

                ReadOnly Property Color() As Color
                    Get
                        Return _Color
                    End Get
                End Property

                ReadOnly Property R() As Byte
                    Get
                        Return _Color.R
                    End Get
                End Property

                ReadOnly Property G() As Byte
                    Get
                        Return _Color.G
                    End Get
                End Property

                ReadOnly Property B() As Byte
                    Get
                        Return _Color.B
                    End Get
                End Property

                ReadOnly Property Hue() As Single
                    Get
                        Return Convert.ToInt32(_Color.GetHue)
                    End Get
                End Property

                ReadOnly Property Saturation() As Single
                    Get
                        Return _Color.GetSaturation
                    End Get
                End Property

                ReadOnly Property Brightness() As Single
                    Get
                        Return _Color.GetBrightness
                    End Get
                End Property

                ReadOnly Property Hexadecimal() As String
                    Get
                        Return "#" & FixHex(Hex(_Color.R)) & FixHex(Hex(_Color.G)) & FixHex(Hex(_Color.B))
                    End Get
                End Property

                ReadOnly Property Win32() As Int32
                    Get
                        Return ColorTranslator.ToWin32(_Color)
                    End Get
                End Property

                Private Function FixHex(ByVal hex As String) As String
                    If hex.Length < 2 Then
                        FixHex = "0" & hex
                    Else
                        FixHex = hex
                    End If
                End Function

            End Class

#End Region

#End Region

        End Namespace

        Public Class ImageCombo
            Inherits System.Windows.Forms.GroupBox
            Private TS As New ToolStrip
            Private WithEvents TSBtn As New ToolStripDropDownButton
            Private TSLbl As ToolStripLabel
            Private WithEvents LstV As New ListView
            Private ColumnHeader As ColumnHeader
            Private TSBtnClicked As Boolean = False
            Private _HeaderColor As Color = Color.SlateGray
            Private _HeaderTextColor As Color = Color.White
            Private _ListBackColor As Color = Color.White
            Private _ListForeColor As Color = Color.Black
            Private _ImgList As New ImageList
            Private _BigImgList As New ImageList
            Private _ImageType As e_ImageType = e_ImageType.Small
            Private _Height As Integer = 200

            Private _ShowItemToolTips As Boolean = True
            Private _RightToLeft As Boolean = False
            Private _HoverSelecting As Boolean = False
            Private _HotTracking As Boolean = False
            Private _GridLines As Boolean = False
            Private _HeaderText As String = "No Items"

            Event HeaderColorChanged(ByVal sender As Object, ByVal e As EventArgs)
            Event HeaderTextColorChanged(ByVal sender As Object, ByVal e As EventArgs)
            Public Event SelectedItemChange(ByVal Item As String, ByVal ImageIndex As Integer)


            Enum e_ImageType
                Small = 1
                Big = 2
            End Enum

#Region " Windows Form Designer generated code "

            Public Sub New()
                MyBase.New()
                'This call is required by the Windows Form Designer.
                InitializeComponent()

                'Add any initialization after the InitializeComponent() call
                MyBase.SetStyle(ControlStyles.AllPaintingInWmPaint, True)
                MyBase.SetStyle(ControlStyles.UserPaint, True)
                MyBase.SetStyle(ControlStyles.DoubleBuffer, True)
                MyBase.SetStyle(ControlStyles.ResizeRedraw, True)
                MyBase.SetStyle(ControlStyles.SupportsTransparentBackColor, True)
            End Sub

            'UserControl1 overrides dispose to clean up the component list.
            Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
                If disposing Then
                    If Not (components Is Nothing) Then
                        components.Dispose()
                    End If
                End If
                MyBase.Dispose(disposing)
            End Sub

            'Required by the Windows Form Designer
            Private components As System.ComponentModel.IContainer

            'NOTE: The following procedure is required by the Windows Form Designer
            'It can be modified using the Windows Form Designer.  
            'Do not modify it using the code editor.
            <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
                components = New System.ComponentModel.Container
                Me.TS = New System.Windows.Forms.ToolStrip
                Me.TSLbl = New System.Windows.Forms.ToolStripLabel
                Me.TSBtn = New System.Windows.Forms.ToolStripDropDownButton
                Me.LstV = New System.Windows.Forms.ListView
                Me.ColumnHeader = New System.Windows.Forms.ColumnHeader

                'TS
                '
                Me.TS.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TSLbl, Me.TSBtn})
                Me.TS.Location = New System.Drawing.Point(0, 0)
                Me.TS.Name = "TS"
                Me.TS.Size = New System.Drawing.Size(833, 25)
                Me.TS.TabIndex = 6
                Me.TS.Text = "ToolStrip1"
                '
                'TSLbl
                '
                Me.TSLbl.Name = "TSLbl"
                Me.TSLbl.Size = New System.Drawing.Size(45, 22)
                Me.TSLbl.Text = "No Item"
                '
                'TSBtn
                '
                Me.TSBtn.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right
                Me.TSBtn.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Text
                Me.TSBtn.ImageTransparentColor = System.Drawing.Color.Magenta
                Me.TSBtn.Name = "TSBtn"
                Me.TSBtn.Size = New System.Drawing.Size(24, 22)
                Me.TSBtn.Text = ""
                '
                'lstV
                '
                Me.LstV.Activation = System.Windows.Forms.ItemActivation.OneClick
                Me.LstV.BorderStyle = System.Windows.Forms.BorderStyle.None
                Me.LstV.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader})
                Me.LstV.FullRowSelect = True
                Me.LstV.GridLines = True
                Me.LstV.HideSelection = False
                Me.LstV.Location = New System.Drawing.Point(Me.Left + 5, Me.Top + Me.TS.Height + 20)
                Me.LstV.Name = "lstV"
                Me.LstV.ShowGroups = False
                Me.LstV.ShowItemToolTips = True
                Me.LstV.Size = New System.Drawing.Size(50, 200)
                Me.LstV.TabIndex = 3
                Me.LstV.UseCompatibleStateImageBehavior = False
                Me.LstV.View = System.Windows.Forms.View.Details
                Me.LstV.HeaderStyle = ColumnHeaderStyle.None
                Me.LstV.GridLines = False
                Me.LstV.Columns.Item(0).Width = Me.LstV.Width
                Me.LstV.FullRowSelect = True
                Me.LstV.Visible = True
                '

                Me.Controls.Add(Me.TS)
                Me.Controls.Add(Me.LstV)
            End Sub

#End Region


            Protected Overrides Sub OnPaint(ByVal e As System.Windows.Forms.PaintEventArgs)
                MyBase.OnPaint(e)
                TS.BackColor = HeaderColor
                TSLbl.ForeColor = HeaderTextColor
                LstV.BackColor = _ListBackColor
                LstV.ForeColor = _ListForeColor
                LstV.ShowItemToolTips = _ShowItemToolTips
                If _RightToLeft Then
                    TSLbl.RightToLeft = Windows.Forms.RightToLeft.Yes
                    LstV.RightToLeft = Windows.Forms.RightToLeft.Yes
                Else
                    TSLbl.RightToLeft = Windows.Forms.RightToLeft.No
                    LstV.RightToLeft = Windows.Forms.RightToLeft.No
                End If
                LstV.HoverSelection = _HoverSelecting
                LstV.HotTracking = _HotTracking
                LstV.GridLines = _GridLines
                TSLbl.Text = _HeaderText

                If TSBtnClicked Then
                    Me.Height = LstV.Height + TS.Height + TS.Top + 10
                    Me.LstV.Width = Me.Width - 10
                Else
                    Me.Height = TS.Height + TS.Top + 5
                End If
                Me.LstV.Columns.Item(0).Width = Me.LstV.Width
            End Sub

#Region "Private Events"
            Private Sub TSBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TSBtn.Click
                TSBtnClicked = Not TSBtnClicked
            End Sub
            Private Sub ListView_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles LstV.MouseLeave
                Me.Height = TS.Height + TS.Top + 5
                TSBtnClicked = False
            End Sub
            Private Sub LstV_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles LstV.SelectedIndexChanged
                If LstV.SelectedItems.Count <> 0 Then
                    TSLbl.Text = LstV.SelectedItems(0).Text
                    _HeaderText = TSLbl.Text
                    RaiseEvent SelectedItemChange(LstV.SelectedItems(0).Text, LstV.SelectedItems(0).ImageIndex)
                Else
                    TSLbl.Text = ""
                End If
                Me.Height = TS.Height + TS.Top + 5
                TSBtnClicked = False
            End Sub

#End Region

#Region "Properties"
            Public ReadOnly Property SelectedIndex() As Integer
                Get
                    If LstV.SelectedIndices.Count > 0 Then
                        Return LstV.SelectedIndices(0)
                    End If
                    Return -1
                End Get
            End Property
            Public ReadOnly Property SelectedItem() As String
                Get
                    If LstV.SelectedIndices.Count > 0 Then
                        Return LstV.SelectedItems(0).Text
                    End If
                    Return ""
                End Get
            End Property
            Public ReadOnly Property SelectedImageIndex() As Integer
                Get
                    If LstV.SelectedIndices.Count > 0 Then
                        Return LstV.SelectedItems(0).ImageIndex
                    End If
                    Return -1
                End Get
            End Property
            Public ReadOnly Property SelectedItemImage() As Image
                Get
                    If LstV.SelectedIndices.Count > 0 Then
                        Select Case ImageSize
                            Case e_ImageType.Big
                                If _BigImgList Is Nothing Then Return Nothing
                                If LstV.SelectedIndices(0) >= _BigImgList.Images.Count Then
                                    Return Nothing
                                Else
                                    Return _BigImgList.Images(LstV.SelectedItems(0).ImageIndex)
                                End If
                            Case e_ImageType.Small
                                If _ImgList Is Nothing Then Return Nothing
                                If LstV.SelectedIndices(0) >= _ImgList.Images.Count Then
                                    Return Nothing
                                Else
                                    Return _ImgList.Images(LstV.SelectedItems(0).ImageIndex)
                                End If
                        End Select
                    End If
                    Return Nothing
                End Get
            End Property
#End Region

#Region "Methods"
            Public Overloads Sub AddItem(ByVal Value As String, Optional ByVal HasImageIndex As Boolean = True)
                LstV.Items.Add(Value, LstV.Items.Count)
            End Sub
            Public Overloads Sub AddItem(ByVal Value As String, ByVal ImageIndex As Integer)
                LstV.Items.Add(Value, ImageIndex)
            End Sub
#End Region

#Region "GUI Properties"
            <System.ComponentModel.Description("The text shown in the header of the imagecombo"), System.ComponentModel.DefaultValue(GetType(String), "No Items")> _
            Property HeaderText() As String
                Get
                    Return _HeaderText
                End Get
                Set(ByVal Value As String)
                    _HeaderText = Value
                    Me.Invalidate()
                End Set
            End Property
            <System.ComponentModel.Description("Show/Hide the Tool Tips"), System.ComponentModel.DefaultValue(GetType(Boolean), "True")> _
            Property ShowItemToolTips() As Boolean
                Get
                    Return _ShowItemToolTips
                End Get
                Set(ByVal Value As Boolean)
                    _ShowItemToolTips = Value
                    Me.Invalidate()
                End Set
            End Property
            <System.ComponentModel.Description("Indicate weather the control shold draw right to left"), System.ComponentModel.DefaultValue(GetType(Boolean), "False")> _
            Property WriteRightToLeft() As Boolean
                Get
                    Return _RightToLeft
                End Get
                Set(ByVal Value As Boolean)
                    _RightToLeft = Value
                    Me.Invalidate()
                End Set
            End Property
            <System.ComponentModel.Description("Allow item to be selected when the mouse over it"), System.ComponentModel.DefaultValue(GetType(Boolean), "False")> _
            Property HoverSelecting() As Boolean
                Get
                    Return _HoverSelecting
                End Get
                Set(ByVal Value As Boolean)
                    If Not Value Then
                        If _HotTracking Then
                            _HoverSelecting = True
                        Else
                            _HoverSelecting = False
                        End If
                    Else
                        _HoverSelecting = True
                    End If
                    Me.Invalidate()
                End Set
            End Property
            <System.ComponentModel.Description("Allow item to appear as a hyperlink when the mouse over it"), System.ComponentModel.DefaultValue(GetType(Boolean), "False")> _
            Property HotTracking() As Boolean
                Get
                    Return _HotTracking
                End Get
                Set(ByVal Value As Boolean)
                    If Value Then
                        If _HoverSelecting Then
                            _HotTracking = Value
                        Else
                            _HotTracking = False
                        End If
                    Else
                        _HotTracking = Value
                    End If
                    Me.Invalidate()
                End Set
            End Property
            <System.ComponentModel.Description("Display grid lines around items show only when it in small images mode"), System.ComponentModel.DefaultValue(GetType(Boolean), "False")> _
            Property GridLines() As Boolean
                Get
                    Return _GridLines
                End Get
                Set(ByVal Value As Boolean)
                    _GridLines = Value
                    Me.Invalidate()
                End Set
            End Property
            <System.ComponentModel.Description("Sets the List fore color for the ImageCombo"), System.ComponentModel.DefaultValue(GetType(Color), "0,0,0")> _
            Property ListForeColor() As Color
                Get
                    Return _ListForeColor
                End Get
                Set(ByVal Value As Color)
                    _ListForeColor = Value
                    Me.Invalidate()
                End Set
            End Property
            <System.ComponentModel.Description("Sets the List back color for the ImageCombo"), System.ComponentModel.DefaultValue(GetType(Color), "255,255,255")> _
            Property ListBackColor() As Color
                Get
                    Return _ListBackColor
                End Get
                Set(ByVal Value As Color)
                    _ListBackColor = Value
                    Me.Invalidate()
                End Set
            End Property
            <System.ComponentModel.Description("Sets the Header color for the ImageCombo"), System.ComponentModel.DefaultValue(GetType(Color), "112,128,144")> _
            Property HeaderColor() As Color
                Get
                    Return _HeaderColor
                End Get
                Set(ByVal Value As Color)
                    _HeaderColor = Value
                    Me.Invalidate()
                    RaiseEvent HeaderColorChanged(Me, New EventArgs)
                End Set
            End Property
            <System.ComponentModel.Description("Sets the Text Header color for the ImageCombo"), System.ComponentModel.DefaultValue(GetType(Color), "255,255,255")> _
            Property HeaderTextColor() As Color
                Get
                    Return _HeaderTextColor
                End Get
                Set(ByVal Value As Color)
                    _HeaderTextColor = Value
                    Me.Invalidate()
                    RaiseEvent HeaderTextColorChanged(Me, New EventArgs)
                End Set
            End Property
            <System.ComponentModel.Description("Sets the Small ImageList for the ImageCombo")> _
            Public Property SmallImagList() As ImageList
                Get
                    Return _ImgList
                End Get
                Set(ByVal value As ImageList)
                    _ImgList = value
                    If Not value Is Nothing Then LstV.SmallImageList = value
                End Set
            End Property
            <System.ComponentModel.Description("Sets the Big ImageList for the ImageCombo")> _
            Public Property BigImagList() As ImageList
                Get
                    Return _BigImgList
                End Get
                Set(ByVal value As ImageList)
                    _BigImgList = value
                    If Not value Is Nothing Then LstV.LargeImageList = value
                End Set
            End Property
            <System.ComponentModel.Description("Sets the Big ImageList for the ImageCombo"), System.ComponentModel.DefaultValue(GetType(e_ImageType), "Small")> _
            Public Property ImageSize() As e_ImageType
                Get
                    Return _ImageType
                End Get
                Set(ByVal value As e_ImageType)
                    _ImageType = value
                    Select Case value
                        Case e_ImageType.Big
                            LstV.View = View.LargeIcon
                        Case e_ImageType.Small
                            LstV.View = View.Details
                    End Select
                End Set
            End Property
            <System.ComponentModel.Description("Sets the Height for the ImageCombo"), System.ComponentModel.DefaultValue(GetType(Integer), "200")> _
            Public Property ScrolledHieght() As Integer
                Get
                    Return _Height
                End Get
                Set(ByVal value As Integer)
                    _Height = value
                    LstV.Height = value
                End Set
            End Property
#End Region

        End Class

        Public Class ProgressBar
            Inherits Label
            Enum directionEnum
                LeftToRight = 0
                RightToLeft = 1
                TopToBottom = 2
                BottomToTop = 3
            End Enum
            ' The defaults
            Dim direction As directionEnum = directionEnum.LeftToRight
            Dim pBvalue As Integer = 0
            Dim maxvalue As Integer = 100
            Dim minvalue As Integer = 0
            Dim incrementalStep As Integer = 10
            Dim color1 As Color = Color.Blue
            Dim color2 As Color = Color.Blue

            Public Sub New()
                MyBase.New()
                Me.BackColor = Color.Transparent
            End Sub
            Public Property Maximum() As Integer
                Get
                    Return maxvalue
                End Get
                Set(ByVal Value As Integer)
                    If Value = minvalue Then Exit Property
                    maxvalue = Value
                    ' Reset the value
                    pBvalue = minvalue
                End Set
            End Property
            Public Property Minimum() As Integer
                Get
                    Return minvalue
                End Get
                Set(ByVal Value As Integer)
                    If Value = maxvalue Then Exit Property
                    minvalue = Value
                    ' Reset the values
                    pBvalue = minvalue
                End Set
            End Property
            Public Property [Step]() As Integer
                Get
                    Return incrementalStep
                End Get
                Set(ByVal Value As Integer)
                    If Value > (maxvalue - minvalue) Then Exit Property
                    incrementalStep = Value
                End Set
            End Property
            Public Property Value() As Integer
                Get
                    Return pBvalue
                End Get
                Set(ByVal Value As Integer)
                    If Value > maxvalue Or Value < minvalue Then Exit Property
                    pBvalue = Value
                    ' Trigger the Paint event
                    Me.Invalidate(New Region(New RectangleF(0, 0, Me.Width, Me.Height)))
                End Set
            End Property
            Public Property GradientColor1() As Color
                Get
                    Return color1
                End Get
                Set(ByVal Value As Color)
                    color1 = Value
                End Set
            End Property
            Public Property GradientColor2() As Color
                Get
                    Return color2
                End Get
                Set(ByVal Value As Color)
                    color2 = Value
                End Set
            End Property
            Public Property PGBarColor() As Color
                Get
                    Return color1
                End Get
                Set(ByVal Value As Color)
                    color1 = Value
                    color2 = Value
                End Set
            End Property
            Public Property PGBarDirection() As directionEnum
                Get
                    Return direction
                End Get
                Set(ByVal Value As directionEnum)
                    direction = Value
                End Set
            End Property
            Public Sub increment(ByVal incrementalValue As Integer)
                incrementValue(incrementalValue)
            End Sub
            Public Sub performStep()
                incrementValue(incrementalStep)
            End Sub
            Private Sub incrementValue(ByVal Value As Integer)
                Select Case Value
                    Case Is > 0
                        If pBvalue < maxvalue Then
                            If pBvalue + Value <= maxvalue Then
                                pBvalue += Value
                            Else
                                pBvalue = maxvalue
                            End If
                        End If
                    Case Is < 0
                        If pBvalue > minvalue Then
                            If pBvalue + Value >= minvalue Then
                                pBvalue += Value
                            Else
                                pBvalue = minvalue
                            End If
                        End If
                End Select
                ' Trigger the Paint event
                Me.Invalidate(New Region(New RectangleF(0, 0, Me.Width, Me.Height)))
            End Sub
            Private Sub drawBar(ByVal grfx As Graphics)
                Dim x1 As Integer = 0
                Dim y1 As Integer = 0
                Dim x2 As Integer = CInt((pBvalue / (maxvalue - minvalue)) * Me.Width)
                Dim y2 As Integer = Me.Height
                Dim lgm As LinearGradientMode = LinearGradientMode.Horizontal
                Select Case direction
                    Case directionEnum.RightToLeft
                        x1 = CInt(Me.Width * (1 - (pBvalue / (maxvalue - minvalue))))
                        x2 = Me.Width
                    Case directionEnum.TopToBottom
                        x2 = Me.Width
                        y2 = CInt((pBvalue / (maxvalue - minvalue)) * Me.Height)
                        lgm = LinearGradientMode.Vertical
                    Case directionEnum.BottomToTop
                        y1 = CInt(Me.Height * (1 - (pBvalue / (maxvalue - minvalue))))
                        x2 = Me.Width
                        lgm = LinearGradientMode.Vertical
                End Select
                If color1.Equals(color2) Then
                    Dim drwBrush As SolidBrush = New SolidBrush(color1)
                    grfx.FillRectangle(drwBrush, x1, y1, x2, y2)
                    drwBrush.Dispose()
                Else
                    Dim drwbrush As LinearGradientBrush = New LinearGradientBrush(grfx.ClipBounds, color1, color2, lgm)
                    grfx.FillRectangle(drwbrush, x1, y1, x2, y2)
                    drwbrush.Dispose()
                End If
            End Sub
            Private Sub ProgressBarPaint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint
                Dim grfx As Graphics = e.Graphics()
                drawBar(grfx)
            End Sub
        End Class

        Public Class TextBoxFocusColor
            Inherits TextBox
            Private _color As Color
            Private _fColor As Color

            Public Property FocusColor() As Color
                Get
                    Return _fColor
                End Get
                Set(ByVal Value As Color)
                    _fColor = Value
                End Set
            End Property

            Protected Overrides Sub OnGotFocus(e As System.EventArgs)
                MyBase.OnGotFocus(e)
                _color = Me.BackColor
                Me.BackColor = _fColor
            End Sub

            Protected Overrides Sub OnLostFocus(e As System.EventArgs)
                MyBase.OnLostFocus(e)
                Me.BackColor = _color
            End Sub
        End Class


        Public Class Paint

#Region "Strucure"
            Private Shared P As sp_Plate
            Private Shared Frm As Form
            Private Shared FFlood As Boolean
            Structure sp_Point
                Dim Piont As Integer
                Dim Color As Color
            End Structure
            Structure sp_Plate
                Dim Form As Form
                Dim PictureBox As PictureBox
                Dim TopLeftPoint, TopRightPoint As sp_Point
                Dim DownLeftPoint, DownRightPoint As sp_Point
                Dim CenterColor As Color
            End Structure
            Structure sp_Object_Property_PathGradients
                Event err()
                Public Overloads Sub Draw(ByVal Palet As sp_Plate)
                    FFlood = False
                    P = New sp_Plate
                    P = Palet
                    If Not P.Form Is Nothing Then
                        Frm = New Form
                        Frm = P.Form
                    End If
                    PDraw()
                End Sub
                Public Overloads Sub Draw(ByVal Palet As sp_Plate, ByVal FillFlood As Boolean)
                    P = New sp_Plate
                    P = Palet
                    If Not P.Form Is Nothing Then
                        Frm = New Form
                        Frm = P.Form
                    End If
                    FFlood = FillFlood
                    PDraw()
                End Sub
                Public Overloads Sub DrawPic(ByVal Palet As sp_Plate, ByVal FillFlood As Boolean)
                    P = New sp_Plate
                    P = Palet
                    Frm = New Form
                    Frm = P.Form
                    FFlood = FillFlood
                    PDrawPic()
                End Sub
            End Structure
#End Region

#Region "Properties"
            Public Property DrawPathGradients() As sp_Object_Property_PathGradients
                Get

                End Get
                Set(ByVal Value As sp_Object_Property_PathGradients)

                End Set
            End Property
#End Region

#Region "Private Methods"
            Private Shared Sub PDrawforResize(ByVal sender As Object, ByVal e As System.EventArgs)
                Try
                    Dim G As Graphics
                    If Not P.Form Is Nothing Then
                        G = P.Form.CreateGraphics
                        If FFlood Then
                            P.TopRightPoint.Piont = P.Form.Width
                            P.TopLeftPoint.Piont = 0
                            P.DownRightPoint.Piont = P.Form.Width
                            P.DownLeftPoint.Piont = P.Form.Height
                        End If
                    Else
                        G = P.PictureBox.CreateGraphics
                        If FFlood Then
                            P.TopRightPoint.Piont = P.PictureBox.Width
                            P.TopLeftPoint.Piont = 0
                            P.DownRightPoint.Piont = P.PictureBox.Width
                            P.DownLeftPoint.Piont = P.PictureBox.Height
                        End If
                    End If
                    Dim path As New GraphicsPath
                    path.AddLine(New Point(P.TopLeftPoint.Piont, P.TopLeftPoint.Piont), New Point(P.TopRightPoint.Piont, P.TopLeftPoint.Piont))
                    path.AddLine(New Point(P.TopRightPoint.Piont, P.TopLeftPoint.Piont), New Point(P.TopRightPoint.Piont, P.DownLeftPoint.Piont))
                    path.AddLine(New Point(P.TopRightPoint.Piont, P.DownLeftPoint.Piont), New Point(P.TopLeftPoint.Piont, P.DownLeftPoint.Piont))
                    Dim pathBrush As New PathGradientBrush(path)
                    pathBrush.CenterColor = P.CenterColor
                    Dim surroundColors() As Color = {P.TopLeftPoint.Color, P.TopRightPoint.Color, P.DownRightPoint.Color, P.DownLeftPoint.Color}
                    pathBrush.SurroundColors = surroundColors
                    G.FillPath(pathBrush, path)
                    G = Nothing
                    path = Nothing
                    surroundColors = Nothing
                    pathBrush = Nothing
                    GC.Collect()
                Catch ex As Exception
                End Try
            End Sub
            Private Shared Sub PDrawPicforResize(ByVal sender As Object, ByVal e As System.EventArgs)
                Try
                    Dim G As Graphics
                    G = P.PictureBox.CreateGraphics
                    If FFlood Then
                        P.TopRightPoint.Piont = P.PictureBox.Width
                        P.TopLeftPoint.Piont = 0
                        P.DownRightPoint.Piont = P.PictureBox.Width
                        P.DownLeftPoint.Piont = P.PictureBox.Height
                    End If
                    Dim path As New GraphicsPath
                    path.AddLine(New Point(P.TopLeftPoint.Piont, P.TopLeftPoint.Piont), New Point(P.TopRightPoint.Piont, P.TopLeftPoint.Piont))
                    path.AddLine(New Point(P.TopRightPoint.Piont, P.TopLeftPoint.Piont), New Point(P.TopRightPoint.Piont, P.DownLeftPoint.Piont))
                    path.AddLine(New Point(P.TopRightPoint.Piont, P.DownLeftPoint.Piont), New Point(P.TopLeftPoint.Piont, P.DownLeftPoint.Piont))
                    Dim pathBrush As New PathGradientBrush(path)
                    pathBrush.CenterColor = P.CenterColor
                    Dim surroundColors() As Color = {P.TopLeftPoint.Color, P.TopRightPoint.Color, P.DownRightPoint.Color, P.DownLeftPoint.Color}
                    pathBrush.SurroundColors = surroundColors
                    G.FillPath(pathBrush, path)
                    G = Nothing
                    path = Nothing
                    surroundColors = Nothing
                    pathBrush = Nothing
                    GC.Collect()
                Catch ex As Exception
                End Try
            End Sub
            Private Shared Sub PDraw(Optional ByVal sender As Object = Nothing, Optional ByVal e As System.Windows.Forms.PaintEventArgs = Nothing)
                Try
                    Dim G As Graphics
                    If Not P.Form Is Nothing Then
                        G = P.Form.CreateGraphics

                        If FFlood Then
                            P.TopRightPoint.Piont = P.Form.Width
                            P.TopLeftPoint.Piont = 0
                            P.DownRightPoint.Piont = P.Form.Width
                            P.DownLeftPoint.Piont = P.Form.Height
                        End If
                    Else
                        G = P.PictureBox.CreateGraphics
                        If FFlood Then
                            P.TopRightPoint.Piont = P.PictureBox.Width
                            P.TopLeftPoint.Piont = 0
                            P.DownRightPoint.Piont = P.PictureBox.Width
                            P.DownLeftPoint.Piont = P.PictureBox.Height
                        End If
                    End If

                    Dim path As New GraphicsPath
                    path.AddLine(New Point(P.TopLeftPoint.Piont, P.TopLeftPoint.Piont), New Point(P.TopRightPoint.Piont, P.TopLeftPoint.Piont))
                    path.AddLine(New Point(P.TopRightPoint.Piont, P.TopLeftPoint.Piont), New Point(P.TopRightPoint.Piont, P.DownLeftPoint.Piont))
                    path.AddLine(New Point(P.TopRightPoint.Piont, P.DownLeftPoint.Piont), New Point(P.TopLeftPoint.Piont, P.DownLeftPoint.Piont))
                    Dim pathBrush As New PathGradientBrush(path)
                    pathBrush.CenterColor = P.CenterColor
                    Dim surroundColors() As Color = {P.TopLeftPoint.Color, P.TopRightPoint.Color, P.DownRightPoint.Color, P.DownLeftPoint.Color}
                    pathBrush.SurroundColors = surroundColors
                    G.FillPath(pathBrush, path)
                    G = Nothing
                    path = Nothing
                    surroundColors = Nothing
                    pathBrush = Nothing
                    GC.Collect()
                    AddHandler Frm.Paint, AddressOf PDraw
                    AddHandler Frm.Resize, AddressOf PDrawforResize
                Catch ex As Exception
                End Try
            End Sub
            Private Shared Sub PDrawPic(Optional ByVal sender As Object = Nothing, Optional ByVal e As System.Windows.Forms.PaintEventArgs = Nothing)
                Try
                    Dim G As Graphics
                    G = P.PictureBox.CreateGraphics

                    If FFlood Then
                        P.TopRightPoint.Piont = P.PictureBox.Width
                        P.TopLeftPoint.Piont = 0
                        P.DownRightPoint.Piont = P.PictureBox.Width
                        P.DownLeftPoint.Piont = P.PictureBox.Height
                    End If
                    Dim path As New GraphicsPath
                    path.AddLine(New Point(P.TopLeftPoint.Piont, P.TopLeftPoint.Piont), New Point(P.TopRightPoint.Piont, P.TopLeftPoint.Piont))
                    path.AddLine(New Point(P.TopRightPoint.Piont, P.TopLeftPoint.Piont), New Point(P.TopRightPoint.Piont, P.DownLeftPoint.Piont))
                    path.AddLine(New Point(P.TopRightPoint.Piont, P.DownLeftPoint.Piont), New Point(P.TopLeftPoint.Piont, P.DownLeftPoint.Piont))
                    Dim pathBrush As New PathGradientBrush(path)
                    pathBrush.CenterColor = P.CenterColor
                    Dim surroundColors() As Color = {P.TopLeftPoint.Color, P.TopRightPoint.Color, P.DownRightPoint.Color, P.DownLeftPoint.Color}
                    pathBrush.SurroundColors = surroundColors
                    G.FillPath(pathBrush, path)
                    G = Nothing
                    path = Nothing
                    surroundColors = Nothing
                    pathBrush = Nothing
                    GC.Collect()
                    AddHandler Frm.Resize, AddressOf PDrawPicforResize
                    AddHandler Frm.Paint, AddressOf PDrawPic
                Catch ex As Exception
                End Try
            End Sub
#End Region


        End Class

        Public Class Curve
            Dim G As Graphics
            Friend Shared Pic As System.Windows.Forms.PictureBox
            Private Shared Xe, Ye As Integer
            Private Shared PBackColor As Color = Color.White
            Private Shared PnColor As Color = Color.Yellow
            Private Shared CorPnColor As Color = Color.Red
            Private Shared Pen As New Pen(PnColor)
            Private Shared dotpointType As en_DotPoint = en_DotPoint.Center
            Private PenV As Integer = 1
            Private PLineType As en_PenLineType = en_PenLineType.Solid

#Region "ENUM"
            Enum en_DotPoint
                Center = 0
                TopLeft = 1
                TopRight = 2
                DownLeft = 3
                DownRight = 4
                Custome = 5
            End Enum
            Enum en_PenLineType
                Dash = 1
                Dot = 2
                Solid = 0
                DashDotDot = 4
                DashDot = 3
            End Enum
#End Region

#Region "Properties"
            Public Property PenVolum() As Integer
                Get
                    Return PenV
                End Get
                Set(ByVal Value As Integer)
                    PenV = Value
                    Pen = New Pen(PnColor, Value)
                End Set
            End Property
            Public Property PenLineType() As en_PenLineType
                Get
                    Return PLineType
                End Get
                Set(ByVal Value As en_PenLineType)
                    PLineType = Value
                    Pen.DashStyle = Value
                End Set
            End Property
            Public Property TheDotPoint(Optional ByVal Xdot As Integer = 0, Optional ByVal Ydot As Integer = 0) As en_DotPoint
                Get
                    Return dotpointType
                End Get
                Set(ByVal Value As en_DotPoint)
                    dotpointType = Value
                    Select Case Value
                        Case en_DotPoint.Custome
                            Xe = Xdot
                            Ye = Ydot
                        Case Else
                            If Pic Is Nothing Then Exit Select
                            Xe = Pic.Width
                            Ye = Pic.Height
                    End Select
                End Set
            End Property
            Private ReadOnly Property VirtualCenterPoint() As System.Drawing.Point
                Get
                    Dim p As New System.Drawing.Point
                    p.X = Xe
                    p.Y = Ye
                    Return p
                End Get
            End Property
            Public ReadOnly Property CenterPoint() As System.Drawing.Point
                Get
                    Dim p As Point
                    Select Case dotpointType
                        Case en_DotPoint.Center
                            p.X = Math.Abs(Pic.Width / 2)
                            p.Y = Math.Abs(Pic.Height / 2)
                        Case en_DotPoint.Custome
                            p.X = Xe
                            p.Y = Ye
                        Case en_DotPoint.DownLeft
                            p.X = 0
                            p.Y = Pic.Height
                        Case en_DotPoint.DownRight
                            p.X = Pic.Width
                            p.Y = Pic.Height
                        Case en_DotPoint.TopLeft
                            p.X = 0
                            p.Y = 0
                        Case en_DotPoint.TopRight
                            p.X = Pic.Height
                            p.Y = 0
                    End Select
                    Return p
                End Get
            End Property
            Public Property BackColor() As Color
                Get
                    Return PBackColor
                End Get
                Set(ByVal Value As Color)
                    PBackColor = Value
                    Pic.BackColor = Value
                End Set
            End Property
            Public Property PenColor() As Color
                Get
                    Return PnColor
                End Get
                Set(ByVal Value As Color)
                    PnColor = Value
                    Pen.Color = Value
                End Set
            End Property
            Public Property CordinatorColor() As Color
                Get
                    Return CorPnColor
                End Get
                Set(ByVal value As Color)
                    CorPnColor = value
                End Set
            End Property
            Public WriteOnly Property PictureBox() As System.Windows.Forms.PictureBox
                Set(ByVal Value As System.Windows.Forms.PictureBox)
                    Pic = New System.Windows.Forms.PictureBox
                    Pic = Value
                    G = Pic.CreateGraphics
                End Set
            End Property
#End Region
#Region "Internal Methods"
            Private Overloads Function LinearTransformation(ByVal X As Integer, ByVal Y As Integer) As System.Drawing.Point
                Dim p As New System.Drawing.Point
                Select Case dotpointType
                    Case en_DotPoint.Center
                        p.X = Math.Abs(X + Pic.Width / 2)
                        p.Y = Math.Abs(Y - Pic.Height / 2)
                    Case en_DotPoint.Custome
                        p.X = Math.Abs(X + Xe)
                        p.Y = Math.Abs(Y - Ye)
                    Case en_DotPoint.DownLeft
                        p.X = Math.Abs(X)
                        p.Y = Math.Abs(Ye - Y)
                    Case en_DotPoint.DownRight
                        p.X = Math.Abs(Xe - X)
                        p.Y = Math.Abs(Ye - Y)
                    Case en_DotPoint.TopLeft
                        p.X = X
                        p.Y = Y
                    Case en_DotPoint.TopRight
                        p.X = Math.Abs(Xe - X)
                        p.Y = Math.Abs(Y)
                End Select
                Return p
            End Function
            Private Overloads Function LinearTransformation(ByVal pF As System.Drawing.PointF) As System.Drawing.Point
                Dim p As New System.Drawing.Point
                Select Case dotpointType
                    Case en_DotPoint.Center
                        p.X = Math.Abs(pF.X + Pic.Width / 2)
                        p.Y = Math.Abs(pF.Y - Pic.Height / 2)
                    Case en_DotPoint.Custome
                        p.X = Math.Abs(pF.X + Xe)
                        p.Y = Math.Abs(pF.Y - Ye)
                    Case en_DotPoint.DownLeft
                        p.X = Math.Abs(pF.X)
                        p.Y = Math.Abs(Ye - pF.Y)
                    Case en_DotPoint.DownRight
                        p.X = Math.Abs(Xe - pF.X)
                        p.Y = Math.Abs(Ye - pF.Y)
                    Case en_DotPoint.TopLeft
                        p.X = pF.X
                        p.Y = pF.Y
                    Case en_DotPoint.TopRight
                        p.X = Math.Abs(Xe - pF.X)
                        p.Y = Math.Abs(pF.Y)
                End Select
                Return p
            End Function
#End Region
#Region "Methods"
            Public Sub CoordinatorsInialization()
                If Pic Is Nothing Then Exit Sub
                Pen.Color = CorPnColor
                Select Case dotpointType
                    Case en_DotPoint.Center
                        G.DrawLine(Pen, CSng(Pic.Width / 2), Pic.Height, CSng(Pic.Width / 2), 0)
                        G.DrawLine(Pen, 0, CSng(Pic.Height / 2), Pic.Width, CSng(Pic.Height / 2))
                    Case en_DotPoint.Custome
                        G.DrawLine(Pen, Xe, 0, Xe, CSng(Pic.Height))
                        G.DrawLine(Pen, 0, Ye, CSng(Pic.Width), Ye)
                    Case en_DotPoint.DownLeft
                        G.DrawLine(Pen, 0, CSng(Pic.Height) - 1, CSng(Pic.Width), CSng(Pic.Height) - 1)
                        G.DrawLine(Pen, 0, CSng(Pic.Height) - 1, 0, 0)
                    Case en_DotPoint.DownRight
                        G.DrawLine(Pen, 0, CSng(Pic.Height) - 1, CSng(Pic.Width), CSng(Pic.Height) - 1)
                        G.DrawLine(Pen, CSng(Pic.Width) - 1, CSng(Pic.Height), CSng(Pic.Width) - 1, 0)
                    Case en_DotPoint.TopLeft
                        G.DrawLine(Pen, 0, CSng(Pic.Height) - 1, 0, 0)
                        G.DrawLine(Pen, 0, 0, CSng(Pic.Width), 0)
                    Case en_DotPoint.TopRight
                        G.DrawLine(Pen, 0, 0, CSng(Pic.Width), 0)
                        G.DrawLine(Pen, CSng(Pic.Width) - 1, CSng(Pic.Height), CSng(Pic.Width) - 1, 0)
                End Select
            End Sub
            Public Sub Line(ByVal X1 As Integer, ByVal Y1 As Integer, ByVal X2 As Integer, ByVal Y2 As Integer)
                Pen.Color = PnColor
                Dim Pdot As New System.Drawing.Point
                Dim Pend As New System.Drawing.Point
                Pdot = LinearTransformation(X1, Y1)
                Pend = LinearTransformation(X2, Y2)
                G.DrawLine(Pen, Pdot, Pend)
            End Sub
            Public Sub Curve(ByVal Points() As System.Drawing.PointF, Optional ByVal Closed As Boolean = False)
                Dim inx As Integer
                Dim ptemp As New Point
                Pen.Color = PnColor
                For inx = 0 To UBound(Points)
                    ptemp = LinearTransformation(Points(inx))
                    Points(inx) = New PointF
                    Points(inx).X = ptemp.X
                    Points(inx).Y = ptemp.Y
                Next
                If Closed Then
                    G.DrawClosedCurve(Pen, Points)
                Else
                    G.DrawCurve(Pen, Points, 1)
                End If
            End Sub
            Public Sub BeziersCurve(ByVal Points() As System.Drawing.PointF)
                Dim inx As Integer
                Dim ptemp As New Point
                Pen.Color = PnColor
                For inx = 0 To UBound(Points)
                    ptemp = LinearTransformation(Points(inx))
                    Points(inx) = New PointF
                    Points(inx).X = ptemp.X
                    Points(inx).Y = ptemp.Y
                Next
                If UBound(Points) < 3 Then Exit Sub
                G.DrawBezier(Pen, Points(0).X, Points(0).Y, Points(1).X, Points(1).Y, Points(2).X, Points(2).Y, Points(3).X, Points(3).Y)
            End Sub
#End Region

        End Class

        Public Class ScreenCapture

            ' The ScreenCapture class allows you to take screenshots (printscreens)
            ' of the desktop or of individual windows.
            '
            ' Usage:
            '
            ' PictureBox1.Image = ScreenCapture.GrabScreen()
            ' PictureBox1.Image = ScreenCapture.GrabActiveWindow()
            ' PictureBox1.Image = ScreenCapture.GrabWindow(SomeHwnd)
            '
            ' PictureBox1.Image = ScreenCapture.GrabScreen(X, Y, Width, Height)
            ' PictureBox1.Image = ScreenCapture.GrabScreen(Rect)
            ' PictureBox1.Image = ScreenCapture.GrabScreen(Location, Size)


            Public Sub SaveFormToImage(ByRef frm As Form, ByVal ImagePath As String, ByVal ImageType As Imaging.ImageFormat)
                Dim bounds As Rectangle = frm.Bounds
                Dim bstyle As FormBorderStyle = frm.FormBorderStyle
                Dim pt As Point = frm.PointToScreen(bounds.Location)
                Dim bitmap As New Bitmap(frm.Width - 50, frm.Height - 50)
                Dim g As Graphics = Graphics.FromImage(bitmap)
                '---

                '---
                frm.FormBorderStyle = Windows.Forms.FormBorderStyle.None
                g.CopyFromScreen(New Point(frm.Left, frm.Top), Point.Empty, bounds.Size)
                bitmap.Save(ImagePath, ImageType)
                frm.FormBorderStyle = bstyle
                '---

                '---
            End Sub

#Region "Constants"

            Private Const HORZRES As Integer = 8
            Private Const VERTRES As Integer = 10
            Private Const SRCCOPY = &HCC0020
            Private Const SRCINVERT = &H660046

            Private Const USE_SCREEN_WIDTH = -1
            Private Const USE_SCREEN_HEIGHT = -1

#End Region

#Region "API's"

            Private Structure RECT
                Public Left As Int32
                Public Top As Int32
                Public Right As Int32
                Public Bottom As Int32
            End Structure

            Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, ByVal lpInitData As String) As Integer
            Private Declare Function CreateCompatibleDC Lib "GDI32" (ByVal hDC As Integer) As Integer
            Private Declare Function DeleteDC Lib "GDI32" (ByVal hDC As Integer) As Integer
            Private Declare Function GetWindowDC Lib "user32" Alias "GetWindowDC" (ByVal hwnd As Long) As Integer
            Private Declare Function ReleaseDC Lib "user32" Alias "ReleaseDC" (ByVal hwnd As Long, ByVal hdc As Long) As Long
            Private Declare Function GetDeviceCaps Lib "gdi32" Alias "GetDeviceCaps" (ByVal hdc As Integer, ByVal nIndex As Integer) As Integer
            Private Declare Function CreateCompatibleBitmap Lib "GDI32" (ByVal hDC As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer) As Integer
            Private Declare Function SelectObject Lib "GDI32" (ByVal hDC As Integer, ByVal hObject As Integer) As Integer
            Private Declare Function DeleteObject Lib "GDI32" (ByVal hObj As Integer) As Integer
            Private Declare Function BitBlt Lib "GDI32" (ByVal hDestDC As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, ByVal SrcX As Integer, ByVal SrcY As Integer, ByVal Rop As Integer) As Integer
            Private Declare Function GetForegroundWindow Lib "user32" Alias "GetForegroundWindow" () As Integer
            Private Declare Function IsWindow Lib "user32" Alias "IsWindow" (ByVal hwnd As Integer) As Long
            Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hWnd As Integer, ByRef lpRect As RECT) As Int32

#End Region

#Region "Enum"
            Public Enum e_imageFormat
                jpg = 1
                bmp = 2
                gif = 3
            End Enum
#End Region

#Region "Public Methods."

            Public Sub GrabScreenToFile(ByVal Path As String, Optional ByVal ImgFormat As e_imageFormat = e_imageFormat.jpg)
                Dim pic As New PictureBox
                pic.Image = GrabScreen(0, 0, USE_SCREEN_WIDTH, USE_SCREEN_HEIGHT)
                Select Case ImgFormat
                    Case e_imageFormat.bmp
                        pic.Image.Save(Path, System.Drawing.Imaging.ImageFormat.Bmp)
                    Case e_imageFormat.jpg
                        pic.Image.Save(Path, System.Drawing.Imaging.ImageFormat.Jpeg)
                    Case e_imageFormat.gif
                        pic.Image.Save(Path, System.Drawing.Imaging.ImageFormat.Gif)
                End Select
            End Sub

            Public Shared Function GrabScreen() As Bitmap

                Return GrabScreen(0, 0, USE_SCREEN_WIDTH, USE_SCREEN_HEIGHT)

            End Function

            Public Shared Function GrabScreen(ByVal Rect As Rectangle) As Bitmap

                Return GrabScreen(Rect.X, Rect.Y, Rect.Width, Rect.Height)

            End Function

            Public Shared Function GrabScreen(ByVal Location As System.Drawing.Point, ByVal Size As System.Drawing.Size) As Bitmap

                Return GrabScreen(Location.X, Location.Y, Size.Width, Size.Height)

            End Function

            Public Shared Function GrabScreen(ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer) As Bitmap

                Dim hDesktopDC As Integer
                Dim hOffscreenDC As Integer
                Dim hBitmap As Integer
                Dim hOldBmp As Integer
                Dim MyBitmap As Bitmap = Nothing

                ' Get the desktop device context.
                hDesktopDC = CreateDC("DISPLAY", "", "", "")
                If hDesktopDC Then
                    ' Adjust width and height.
                    If Width = USE_SCREEN_WIDTH Then
                        Width = GetDeviceCaps(hDesktopDC, HORZRES)
                    End If
                    If Height = USE_SCREEN_HEIGHT Then
                        Height = GetDeviceCaps(hDesktopDC, VERTRES)
                    End If
                    ' Create an offscreen device context.
                    hOffscreenDC = CreateCompatibleDC(hDesktopDC)
                    If hOffscreenDC Then
                        ' Create a bitmap for our offscreen device context.
                        hBitmap = CreateCompatibleBitmap(hDesktopDC, Width, Height)
                        If hBitmap Then
                            ' Copy the image and create an instance of the Bitmap class.
                            hOldBmp = SelectObject(hOffscreenDC, hBitmap)
                            BitBlt(hOffscreenDC, 0, 0, Width, Height, hDesktopDC, X, Y, SRCCOPY)
                            MyBitmap = Bitmap.FromHbitmap(New IntPtr(hBitmap))
                            ' Clean up.
                            DeleteObject(SelectObject(hOffscreenDC, hOldBmp))
                        End If
                        DeleteDC(hOffscreenDC)
                    End If
                    DeleteDC(hDesktopDC)
                End If
                ' Return our Bitmap instance.
                Return MyBitmap

            End Function

            Public Shared Function GrabActiveWindow() As Bitmap

                Return GrabWindow(GetForegroundWindow())

            End Function

            Public Shared Function GrabWindow(ByVal hWnd As Int32) As Bitmap

                Dim hWindowDC As Long
                Dim hOffscreenDC As Long
                Dim rec As RECT
                Dim nWidth As Long
                Dim nHeight As Long
                Dim hBitmap As Long
                Dim hOldBmp As Long
                Dim MyBitmap As Bitmap = Nothing

                ' Verify if a valid window handle was provided.
                If hWnd <> 0 And IsWindow(hWnd) Then
                    ' Get the window's device context.
                    hWindowDC = GetWindowDC(hWnd)
                    If hWindowDC Then
                        ' Get width and height.
                        If GetWindowRect(hWnd, rec) Then
                            nWidth = rec.Right - rec.Left
                            nHeight = rec.Bottom - rec.Top
                            ' Create an offscreen device context.
                            hOffscreenDC = CreateCompatibleDC(hWindowDC)
                            If hOffscreenDC Then
                                ' Create a bitmap for our offscreen device context.
                                hBitmap = CreateCompatibleBitmap(hWindowDC, nWidth, nHeight)
                                If hBitmap Then
                                    ' Copy the image and create an instance of the Bitmap class.
                                    hOldBmp = SelectObject(hOffscreenDC, hBitmap)
                                    BitBlt(hOffscreenDC, 0, 0, nWidth, nHeight, hWindowDC, 0, 0, SRCCOPY)
                                    MyBitmap = Bitmap.FromHbitmap(New IntPtr(hBitmap))
                                    ' Clean up.
                                    DeleteObject(SelectObject(hOffscreenDC, hOldBmp))
                                End If
                                DeleteDC(hOffscreenDC)
                            End If
                        End If
                        ReleaseDC(hWnd, hWindowDC)
                    End If
                End If
                ' Return our Bitmap instance.
                Return MyBitmap

            End Function

#End Region

        End Class

        Public Class ImageProcessing
            Private G As Graphics

#Region "Methods"
            Public Sub Rotate(ByRef ImageSource As Image, Optional ByVal Direction As RotateFlipType = RotateFlipType.Rotate180FlipY)
                Try
                    ImageSource.RotateFlip(Direction)
                Catch ex As Exception
                End Try
            End Sub

            Public Sub Clear(ByRef PicSource As PictureBox)
                G = PicSource.CreateGraphics
                G.Clear(PicSource.BackColor)
            End Sub
            Public Sub SetText(ByRef PicSource As PictureBox, ByVal Text As String, ByVal TextFont As System.Drawing.Font, ByVal ColorBrush As Brush, ByVal X As Integer, ByVal Y As Integer)
                G = PicSource.CreateGraphics
                If TextFont Is Nothing Then Exit Sub
                G.DrawString(Text, TextFont, ColorBrush, X, Y)
            End Sub
#End Region
        End Class

        Class ObjectMover

#Region "API"

            <DllImport("User32.dll")> _
            Private Shared Function SetForegroundWindow(ByVal hWnd As IntPtr) As Boolean
            End Function

            <DllImport("user32.dll")> _
            Private Shared Function ReleaseCapture() As Integer
            End Function

            <DllImport("user32.dll")> _
            Private Shared Function SendMessage(ByVal hwnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As Integer, ByVal lParam As Integer) As Integer
            End Function

#End Region

#Region "Functions"

            Public Shared Sub FocusObject(ByVal Handle As IntPtr)
                SetForegroundWindow(Handle)
            End Sub

            Public Shared Sub MoveObject(ByVal Handle As IntPtr)
                ReleaseCapture()
                SendMessage(Handle, 161, 2, 0)
            End Sub

#End Region

        End Class

        Class PopUpForm
            Private frm As Form
            Private _Location As e_Location = e_Location.DownRight
            Private ifstart As Boolean = False
            Private _X As Integer = 1, _Y As Integer = 1
            Private t As New EAMS.Diagonstic.Synchrounce
            Private tDown As New EAMS.Diagonstic.Synchrounce
            Private tAuto As EAMS.Diagonstic.Synchrounce
            Private _Step As Integer = 5
            Private _Vis As e_Visualization = e_Visualization.Scrolled
            Private _XFlag As Boolean = False, _YFlag As Boolean = False
            Public Event PopUpFinished()
            Public Event PopDownFinished()
            Private _Auto As Boolean = False
            Private _WaitTime As Integer = 10


#Region "Enum"
            Enum e_Location
                TopLeft = 1
                TopRight = 2
                DownRight = 3
                DownLeft = 4
                CenterHeight = 5
                CenterWidth = 6
                CenterHeightWidth = 7
            End Enum
            Enum e_Visualization
                Scrolled = 1
                Flashing = 2
            End Enum
#End Region
#Region "Properties"
            Public Property WaitTimeForDisappearMillisecond() As Integer
                Get
                    Return _WaitTime
                End Get
                Set(ByVal value As Integer)
                    _WaitTime = value
                End Set
            End Property
            Public Property AutoPopDown() As Boolean
                Get
                    Return _Auto
                End Get
                Set(ByVal value As Boolean)
                    _Auto = value
                End Set
            End Property
            Public Property StepSpeed() As Integer
                Get
                    Return _Step
                End Get
                Set(ByVal value As Integer)
                    _Step = value
                End Set
            End Property
            Public Property PopupLocation() As e_Location
                Get
                    Return _Location
                End Get
                Set(ByVal value As e_Location)
                    _Location = value
                End Set
            End Property
            Public Property VisualizationType() As e_Visualization
                Get
                    Return _Vis
                End Get
                Set(ByVal value As e_Visualization)
                    _Vis = value
                End Set
            End Property
            Public Property Form() As Form
                Get
                    If frm Is Nothing Then
                        Return Nothing
                    End If
                    Return frm
                End Get
                Set(ByVal value As Form)
                    frm = New Form
                    frm = value
                    frm.ShowInTaskbar = False
                    frm.StartPosition = FormStartPosition.Manual
                End Set
            End Property
#End Region
#Region "Methods"
            Public Sub New()
                AddHandler t.VirtualTimer, AddressOf _PopUp
                AddHandler tDown.VirtualTimer, AddressOf _PopUpDown
                AddHandler PopUpFinished, AddressOf _AutoPopDown
                t.IntervalType = Diagonstic.Synchrounce.IntType.Millisecond
                t.Interval = 1
                tDown.IntervalType = Diagonstic.Synchrounce.IntType.Millisecond
                tDown.Interval = 1
            End Sub
            Public Sub PopUp()
                If frm Is Nothing Then Exit Sub
                ifstart = False
                t.ActiveTimer()
            End Sub
            Public Sub UnPopUp()
                If frm Is Nothing Then Exit Sub
                ifstart = False
                tDown.ActiveTimer()
            End Sub
#End Region
#Region "Private Methods"
            Private Sub _PopUp(ByVal timming As Integer)
                If timming = 0 Then Exit Sub
                Select Case _Vis
                    Case e_Visualization.Scrolled
                        _Scrolled()
                    Case e_Visualization.Flashing
                        Flashing()
                End Select
            End Sub
            Private Sub _PopUpDown(ByVal timming As Integer)
                If timming = 0 Then Exit Sub
                Select Case _Vis
                    Case e_Visualization.Scrolled
                        _ScrolledDown()
                    Case e_Visualization.Flashing
                        If Not ifstart Then
                            _X = 100
                            ifstart = True
                        End If
                        _FlashingDown()
                End Select
            End Sub
            Private Sub Flashing()
                If Not ifstart Then
                    Select Case _Location
                        Case e_Location.CenterHeight
                            frm.Left = Convert.ToInt32((Screen.PrimaryScreen.Bounds.Width / 2) - (frm.Width / 2))
                            frm.Top = Convert.ToInt32((Screen.PrimaryScreen.Bounds.Height / 2) - (frm.Height / 2))
                            frm.Opacity = 1 / 100
                            _X = 1
                            frm.TopMost = True
                            frm.Show()
                            frm.Visible = False
                        Case e_Location.CenterWidth
                            frm.Left = Convert.ToInt32((Screen.PrimaryScreen.Bounds.Width / 2) - (frm.Width / 2))
                            frm.Top = Convert.ToInt32((Screen.PrimaryScreen.Bounds.Height / 2) - (frm.Height / 2))
                            frm.Opacity = 1 / 100
                            _X = 1
                            frm.TopMost = True
                            frm.Show()
                            frm.Visible = False
                        Case e_Location.CenterHeightWidth
                            frm.Left = Convert.ToInt32((Screen.PrimaryScreen.Bounds.Width / 2) - (frm.Width / 2))
                            frm.Top = Convert.ToInt32((Screen.PrimaryScreen.Bounds.Height / 2) - (frm.Height / 2))
                            frm.Opacity = 1 / 100
                            _X = 1
                            frm.TopMost = True
                            frm.Show()
                            frm.Visible = False
                        Case e_Location.DownLeft
                            frm.Left = 6
                            frm.Top = Screen.PrimaryScreen.Bounds.Height - frm.Height - 10
                            frm.Opacity = 1 / 100
                            _X = 1
                            frm.TopMost = True
                            frm.Show()
                            frm.Visible = False
                        Case e_Location.DownRight
                            frm.Left = Screen.PrimaryScreen.Bounds.Width - frm.Width - 5
                            frm.Top = Screen.PrimaryScreen.Bounds.Height - frm.Height - 10
                            frm.Opacity = 1 / 100
                            _X = 1
                            frm.TopMost = True
                            frm.Show()
                            frm.Visible = False
                        Case e_Location.TopLeft
                            frm.Left = 6
                            frm.Top = 6
                            frm.Opacity = 1 / 100
                            _X = 1
                            frm.TopMost = True
                            frm.Show()
                            frm.Visible = False
                        Case e_Location.TopRight
                            frm.Left = Screen.PrimaryScreen.Bounds.Width - frm.Width - 5
                            frm.Top = 6
                            frm.Opacity = 1 / 100
                            _X = 1
                            frm.TopMost = True
                            frm.Show()
                            frm.Visible = False
                    End Select
                    frm.Opacity = 1 / 100
                    ifstart = True
                End If
                _Flashing()
            End Sub
            Private Sub _Scrolled()
                Select Case _Location
                    Case e_Location.DownLeft
                        frm.Left = 6
                        Select Case ifstart
                            Case False
                                _Y = Screen.PrimaryScreen.Bounds.Height
                                ifstart = True
                                frm.Visible = False
                                frm.Show()
                                frm.Top = Screen.PrimaryScreen.Bounds.Height
                                frm.TopMost = True
                            Case Else
                                _PopDown()
                        End Select

                    Case e_Location.DownRight
                        frm.Left = Screen.PrimaryScreen.Bounds.Width - frm.Width - 5
                        Select Case ifstart
                            Case False
                                _Y = Screen.PrimaryScreen.Bounds.Height
                                ifstart = True
                                frm.Visible = False
                                frm.Show()
                                frm.Top = Screen.PrimaryScreen.Bounds.Height
                                frm.TopMost = True
                            Case Else
                                _PopDown()
                        End Select
                    Case e_Location.TopLeft
                        frm.Left = 6
                        frm.Top = 1
                        Select Case ifstart
                            Case False
                                _Y = frm.Height
                                frm.Height = 1
                                ifstart = True
                                frm.Visible = False
                                frm.Show()
                                frm.TopMost = True
                            Case Else
                                _PopTop()
                        End Select
                    Case e_Location.TopRight
                        frm.Left = Screen.PrimaryScreen.Bounds.Width - frm.Width - 5
                        frm.Top = 1
                        Select Case ifstart
                            Case False
                                _Y = frm.Height
                                frm.Height = 1
                                ifstart = True
                                frm.Visible = False
                                frm.Show()
                                frm.TopMost = True
                            Case Else
                                _PopTop()
                        End Select
                    Case e_Location.CenterHeight
                        Select Case ifstart
                            Case False
                                _Y = frm.Height
                                frm.Height = 1
                                ifstart = True
                                frm.Visible = False
                                frm.Show()
                                frm.TopMost = True
                            Case Else
                                frm.Left = Convert.ToInt32((Screen.PrimaryScreen.Bounds.Width / 2) - (frm.Width / 2))
                                frm.Top = Convert.ToInt32((Screen.PrimaryScreen.Bounds.Height / 2) - (frm.Height / 2))
                                _PopTop()
                        End Select
                    Case e_Location.CenterWidth
                        Select Case ifstart
                            Case False
                                _Y = frm.Width
                                frm.Width = 1
                                ifstart = True
                                frm.Visible = False
                                frm.Show()
                                frm.TopMost = True
                            Case Else
                                frm.Left = Convert.ToInt32((Screen.PrimaryScreen.Bounds.Width / 2) - (frm.Width / 2))
                                frm.Top = Convert.ToInt32((Screen.PrimaryScreen.Bounds.Height / 2) - (frm.Height / 2))
                                _PopWidth()
                        End Select
                    Case e_Location.CenterHeightWidth
                        Select Case ifstart
                            Case False
                                _X = frm.Width
                                _Y = frm.Height
                                frm.Width = 1
                                frm.Height = 1
                                ifstart = True
                                frm.Visible = False
                                frm.Show()
                                frm.TopMost = True
                            Case Else
                                frm.Left = Convert.ToInt32((Screen.PrimaryScreen.Bounds.Width / 2) - (frm.Width / 2))
                                frm.Top = Convert.ToInt32((Screen.PrimaryScreen.Bounds.Height / 2) - (frm.Height / 2))
                                _PopWidthHeight()
                        End Select
                End Select
            End Sub
            Private Sub _ScrolledDown()
                Select Case _Location
                    Case e_Location.DownLeft, e_Location.DownRight
                        _UnPopDown()
                    Case e_Location.TopLeft, e_Location.TopRight
                        _UnPopTop()
                    Case e_Location.CenterHeight
                        If ifstart Then
                            _UnPopCenterHeight()
                        Else
                            _Y = frm.Height
                            ifstart = True
                        End If
                    Case e_Location.CenterWidth
                        If ifstart Then
                            _UnPopCenterWidth()
                        Else
                            _X = frm.Width
                            ifstart = True
                        End If
                    Case e_Location.CenterHeightWidth
                        If ifstart Then
                            _UnPopCenterHeightWidth()
                        Else
                            _X = frm.Width
                            _Y = frm.Height
                            ifstart = True
                        End If
                End Select
            End Sub

#Region "Flashing Methods"
            Private Sub _Flashing()
                frm.Visible = True
                If frm.Opacity >= 1 Then
                    frm.Opacity = 1
                    t.DiactiveTimer()
                    frm.Refresh()
                    RaiseEvent PopUpFinished()
                Else
                    _X += _Step
                    frm.Opacity += _X / 100
                End If
                Application.DoEvents()
            End Sub
#End Region
#Region "Scrolled Methods"
            Private Sub _PopDown()
                frm.Visible = True
                frm.Top -= _Step
                If frm.Top < (_Y - frm.Height) Then
                    frm.Top = Convert.ToInt32(Screen.PrimaryScreen.Bounds.Height - frm.Height)
                    t.DiactiveTimer()
                    RaiseEvent PopUpFinished()
                End If
                Application.DoEvents()
            End Sub
            Private Sub _PopTop()
                frm.Visible = True
                frm.Height += _Step
                If frm.Height >= _Y Then
                    frm.Height = _Y
                    t.DiactiveTimer()
                    RaiseEvent PopUpFinished()
                End If
                Application.DoEvents()
            End Sub
            Private Sub _PopWidth()
                frm.Visible = True
                frm.Width += _Step
                If frm.Width >= _Y Then
                    frm.Width = _Y
                    t.DiactiveTimer()
                    RaiseEvent PopUpFinished()
                End If
                Application.DoEvents()
            End Sub
            Private Sub _PopWidthHeight()
                frm.Visible = True
                If frm.Height >= _Y Then
                    frm.Height = _Y
                    If _XFlag Then
                        t.DiactiveTimer()
                        RaiseEvent PopUpFinished()
                    Else
                        _YFlag = True
                    End If
                Else
                    frm.Height += _Step
                End If

                If frm.Width >= _X Then
                    frm.Width = _X
                    If _YFlag Then
                        t.DiactiveTimer()
                        RaiseEvent PopUpFinished()
                    Else
                        _XFlag = True
                    End If
                Else
                    frm.Width += _Step
                End If
                Application.DoEvents()
            End Sub
#End Region
#Region "UnScrolled Method"
            Private Sub _UnPopDown()
                frm.Top += _Step
                If frm.Top >= Screen.PrimaryScreen.Bounds.Height Then
                    frm.Top = Screen.PrimaryScreen.Bounds.Height - 10
                    tDown.DiactiveTimer()
                    frm.Visible = False
                    RaiseEvent PopDownFinished()
                End If
                Application.DoEvents()
            End Sub
            Private Sub _UnPopTop()
                frm.Top -= _Step
                If frm.Top <= 1 - frm.Height Then
                    frm.Top = 1 - frm.Height
                    tDown.DiactiveTimer()
                    frm.Visible = False
                    RaiseEvent PopDownFinished()
                End If
                Application.DoEvents()
            End Sub
            Private Sub _UnPopCenterHeight()
                frm.Left = Convert.ToInt32((Screen.PrimaryScreen.Bounds.Width / 2) - (frm.Width / 2))
                frm.Top = Convert.ToInt32((Screen.PrimaryScreen.Bounds.Height / 2) - (frm.Height / 2))
                If frm.Height - _Step < 0 Then
                    frm.Height = 1
                    tDown.DiactiveTimer()
                    frm.Visible = False
                    frm.Height = _Y
                    RaiseEvent PopDownFinished()
                Else
                    frm.Height -= _Step
                End If
                Application.DoEvents()
            End Sub
            Private Sub _UnPopCenterWidth()
                frm.Left = Convert.ToInt32((Screen.PrimaryScreen.Bounds.Width / 2) - (frm.Width / 2))
                frm.Top = Convert.ToInt32((Screen.PrimaryScreen.Bounds.Height / 2) - (frm.Height / 2))
                If frm.Width - _Step < 0 Then
                    frm.Width = 1
                    tDown.DiactiveTimer()
                    frm.Visible = False
                    frm.Width = _X
                    RaiseEvent PopDownFinished()
                Else
                    frm.Width -= _Step
                End If
                Application.DoEvents()
            End Sub
            Private Sub _UnPopCenterHeightWidth()
                frm.Left = Convert.ToInt32((Screen.PrimaryScreen.Bounds.Width / 2) - (frm.Width / 2))
                frm.Top = Convert.ToInt32((Screen.PrimaryScreen.Bounds.Height / 2) - (frm.Height / 2))
                If frm.Width - _Step < 0 Then
                    frm.Width = 1
                    If _YFlag Then
                        tDown.DiactiveTimer()
                        frm.Visible = False
                        frm.Width = _X
                        frm.Height = _Y
                        RaiseEvent PopDownFinished()
                    Else
                        _XFlag = True
                    End If
                Else
                    frm.Width -= _Step
                End If

                If frm.Height - _Step < 0 Then
                    frm.Height = 1
                    If _XFlag Then
                        tDown.DiactiveTimer()
                        frm.Visible = False
                        frm.Height = _Y
                        frm.Width = _X
                        RaiseEvent PopDownFinished()
                    Else
                        _YFlag = True
                    End If
                Else
                    frm.Height -= _Step
                End If
                Application.DoEvents()
            End Sub
#End Region
#Region "UnFlashing"
            Private Sub _FlashingDown()
                If _X <= 1 Then
                    frm.Opacity = 1 / 100
                    frm.Visible = False
                    tDown.DiactiveTimer()
                    RaiseEvent PopDownFinished()
                Else
                    _X -= 1
                    frm.Opacity = _X / 100
                End If
                Application.DoEvents()
            End Sub
#End Region
#End Region
#Region "Handle Auto PopDown"
            Private Sub _AutoPopDown()
                If _Auto Then
                    tAuto = New EAMS.Diagonstic.Synchrounce
                    AddHandler tAuto.VirtualTimer, AddressOf _WaitTimerTic
                    tAuto.IntervalType = Diagonstic.Synchrounce.IntType.Millisecond
                    tAuto.Interval = _WaitTime
                    tAuto.ActiveTimer()
                End If
            End Sub
            Private Sub _WaitTimerTic(ByVal timming As Integer)
                If timming = 0 Then Exit Sub
                tAuto.DiactiveTimer()
                Me.UnPopUp()
            End Sub
#End Region

        End Class

        Public Class BlinkForm
            <DllImport("user32.dll", entrypoint:="FlashWindow")> _
            Public Shared Function FlashWindow(ByVal hwnd As Integer, ByVal bInvert As Integer) As Integer

            End Function
        End Class
    End Namespace

    Namespace Diagonstic
        Public Class Synchrounce
            Private Timer As New System.Windows.Forms.Timer
            Private Int As Integer, TType As IntType = IntType.Minute
            Private TCounter, VCounter As Double
            Public Event RealTimer(ByVal Timming As Integer)
            Public Event VirtualTimer(ByVal Timming As Integer)

#Region "Enum"
            Enum IntType
                Seconds = 0
                Minute = 1
                Hours = 2
                Millisecond = 3
            End Enum
#End Region
#Region "Initializing"
            Public Sub New(Optional ByVal Interval As Integer = 1, Optional ByVal IntervalType As IntType = IntType.Minute)
                Int = Interval
                TType = IntervalType
                Timer.Enabled = False
                Timer.Interval = 1
                AddHandler Timer.Tick, AddressOf GetQuery
            End Sub
#End Region

#Region "Property"
            Public Property Interval() As Integer
                Get
                    Return Int
                End Get
                Set(ByVal value As Integer)
                    Int = value
                End Set
            End Property
            Public Property IntervalType() As IntType
                Get
                    Return TType
                End Get
                Set(ByVal value As IntType)
                    TType = value
                End Set
            End Property
#End Region

#Region "Methods"
            Public Sub ActiveTimer()
                Timer.Enabled = True
            End Sub
            Public Sub DiactiveTimer()
                If Timer.Enabled Then Timer.Enabled = False
                TCounter = 0
                RaiseEvent RealTimer(0)
                RaiseEvent VirtualTimer(0)
            End Sub
            Public Sub PauseTimer()
                Timer.Enabled = False
            End Sub
#End Region

#Region "Internal Methods"
            Private Sub GetQuery(ByVal Sender As Object, ByVal e As System.EventArgs)
                Application.DoEvents()
                RaiseEvent RealTimer(TCounter)
                Select Case TType
                    Case IntType.Millisecond
                        TCounter += 1
                        If TCounter >= Int Then
                            TCounter = 0
                            VCounter += 1
                            RaiseEvent VirtualTimer(VCounter)
                        End If
                    Case IntType.Seconds
                        TCounter += 1
                        If TCounter >= Int * 60 Then
                            TCounter = 0
                            VCounter += 1
                            RaiseEvent VirtualTimer(VCounter)
                        End If
                    Case IntType.Minute
                        TCounter += 1
                        If TCounter >= Int * 60 * 60 Then
                            TCounter = 0
                            VCounter += 1
                            RaiseEvent VirtualTimer(VCounter)
                        End If
                    Case IntType.Hours
                        TCounter += 1
                        If TCounter >= Int * 60 * 60 * 60 Then
                            TCounter = 0
                            VCounter += 1
                            RaiseEvent VirtualTimer(VCounter)
                        End If
                End Select
            End Sub
#End Region
        End Class

        Public Class Processes
            Private Shared ProcColName, ProcColID, ProcColSize As Collection

#Region "Structure"
            Structure psItems
                Event Errors()
                Public Function ItemName(ByVal Index As Integer) As String
                    Try
                        Index += 1
                        If ProcColName.Count <> 0 Then
                            If Index > ProcColName.Count Then Index = ProcColName.Count
                            Return ProcColName.Item(Index)
                        Else
                            Return ""
                        End If
                    Catch ex As Exception
                        RaiseEvent Errors()
                    End Try
                    Return ""
                End Function
                Public Function ItemID(ByVal Index As Integer) As String
                    Try
                        Index += 1
                        If ProcColID.Count <> 0 Then
                            If Index > ProcColID.Count Then Index = ProcColID.Count
                            Return ProcColID.Item(Index)
                        Else
                            Return ""
                        End If
                    Catch ex As Exception
                        RaiseEvent Errors()
                    End Try
                    Return ""
                End Function
                Public Function ItemMemorySize(ByVal Index As Integer) As String
                    Try
                        Index += 1
                        If ProcColSize.Count <> 0 Then
                            If Index > ProcColSize.Count Then Index = ProcColSize.Count
                            Return ProcColSize.Item(Index)
                        Else
                            Return ""
                        End If
                    Catch ex As Exception
                        RaiseEvent Errors()
                    End Try
                    Return ""
                End Function
            End Structure
#End Region
#Region "Oprations"
            Public Overloads Sub KillProcess(ByVal ProcessName As String)
                Dim Process1() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcesses()
                Dim Process2() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcesses()
                Dim ct1, ct2 As Integer
                Dim blnUnique As Boolean
                For ct1 = 0 To Process2.GetUpperBound(0)
                    If Process2(ct1).ProcessName = ProcessName Then
                        blnUnique = True
                        For ct2 = 0 To Process1.GetUpperBound(0)
                            If Process1(ct2).ProcessName = ProcessName Then
                                Process1(ct2).Kill()
                            End If
                        Next
                        If blnUnique = True Then
                            Exit For
                        End If
                    End If
                Next
            End Sub
            Public Overloads Sub KillProcess(ByVal ProcessID As Integer)
                Dim Process1() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcesses()
                Dim Process2() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcesses()
                Dim ct1, ct2 As Integer
                Dim blnUnique As Boolean
                For ct1 = 0 To Process2.GetUpperBound(0)
                    If Process2(ct1).Id = ProcessID Then
                        blnUnique = True
                        For ct2 = 0 To Process1.GetUpperBound(0)
                            If Process1(ct2).Id = ProcessID Then
                                Process1(ct2).Kill()
                            End If
                        Next
                        If blnUnique = True Then
                            Exit For
                        End If
                    End If
                Next
            End Sub
#End Region
#Region "Properties"
            Public ReadOnly Property Count() As Integer
                Get
                    Dim Process1() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcesses()
                    Return Process1.GetUpperBound(0) + 1
                End Get
            End Property

            Public Overloads ReadOnly Property BasePriority(ByVal ProcessName As String) As Integer
                Get
                    Try
                        Dim Process1() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcesses()
                        Dim Process2() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcesses()
                        Dim ct1, ct2 As Integer
                        Dim blnUnique As Boolean
                        For ct1 = 0 To Process2.GetUpperBound(0)
                            If Process2(ct1).ProcessName = ProcessName Then
                                blnUnique = True
                                For ct2 = 0 To Process1.GetUpperBound(0)
                                    If Process1(ct2).ProcessName = ProcessName Then
                                        Return (Process1(ct2).BasePriority)
                                    End If
                                Next
                                If blnUnique = True Then
                                    Exit For
                                End If
                            End If
                        Next
                    Catch ex As Exception

                    End Try
                    Return 0
                End Get
            End Property
            Public Overloads ReadOnly Property BasePriority(ByVal ProcessID As Integer) As Integer
                Get
                    Try
                        Dim Process1() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcesses()
                        Dim Process2() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcesses()
                        Dim ct1, ct2 As Integer
                        Dim blnUnique As Boolean
                        For ct1 = 0 To Process2.GetUpperBound(0)
                            If Process2(ct1).Id = ProcessID Then
                                blnUnique = True
                                For ct2 = 0 To Process1.GetUpperBound(0)
                                    If Process1(ct2).Id = ProcessID Then
                                        Return (Process1(ct2).BasePriority)
                                    End If
                                Next
                                If blnUnique = True Then
                                    Exit For
                                End If
                            End If
                        Next
                    Catch ex As Exception

                    End Try
                    Return 0
                End Get
            End Property
            Public Overloads ReadOnly Property MachineName(ByVal ProcessName As String) As String
                Get
                    Try
                        Dim Process1() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcesses()
                        Dim Process2() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcesses()
                        Dim ct1, ct2 As Integer
                        Dim blnUnique As Boolean
                        For ct1 = 0 To Process2.GetUpperBound(0)
                            If Process2(ct1).ProcessName = ProcessName Then
                                blnUnique = True
                                For ct2 = 0 To Process1.GetUpperBound(0)
                                    If Process1(ct2).ProcessName = ProcessName Then
                                        Return (Process1(ct2).MachineName)
                                    End If
                                Next
                                If blnUnique = True Then
                                    Exit For
                                End If
                            End If
                        Next
                    Catch ex As Exception

                    End Try
                    Return ""
                End Get
            End Property
            Public Overloads ReadOnly Property MachineName(ByVal ProcessID As Integer) As String
                Get
                    Try
                        Dim Process1() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcesses()
                        Dim Process2() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcesses()
                        Dim ct1, ct2 As Integer
                        Dim blnUnique As Boolean
                        For ct1 = 0 To Process2.GetUpperBound(0)
                            If Process2(ct1).Id = ProcessID Then
                                blnUnique = True
                                For ct2 = 0 To Process1.GetUpperBound(0)
                                    If Process1(ct2).Id = ProcessID Then
                                        Return (Process1(ct2).MachineName)
                                    End If
                                Next
                                If blnUnique = True Then
                                    Exit For
                                End If
                            End If
                        Next
                    Catch ex As Exception

                    End Try
                    Return ""
                End Get
            End Property
            Public Overloads ReadOnly Property MainWindowTitle(ByVal ProcessName As String) As String
                Get
                    Try
                        Dim Process1() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcesses()
                        Dim Process2() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcesses()
                        Dim ct1, ct2 As Integer
                        Dim blnUnique As Boolean
                        For ct1 = 0 To Process2.GetUpperBound(0)
                            If Process2(ct1).ProcessName = ProcessName Then
                                blnUnique = True
                                For ct2 = 0 To Process1.GetUpperBound(0)
                                    If Process1(ct2).ProcessName = ProcessName Then
                                        Return (Process1(ct2).MainWindowTitle)
                                    End If
                                Next
                                If blnUnique = True Then
                                    Exit For
                                End If
                            End If
                        Next
                    Catch ex As Exception

                    End Try
                    Return ""
                End Get
            End Property
            Public Overloads ReadOnly Property MainWindowTitle(ByVal ProcessID As Integer) As String
                Get
                    Try
                        Dim Process1() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcesses()
                        Dim Process2() As System.Diagnostics.Process = System.Diagnostics.Process.GetProcesses()
                        Dim ct1, ct2 As Integer
                        Dim blnUnique As Boolean
                        For ct1 = 0 To Process2.GetUpperBound(0)
                            If Process2(ct1).Id = ProcessID Then
                                blnUnique = True
                                For ct2 = 0 To Process1.GetUpperBound(0)
                                    If Process1(ct2).Id = ProcessID Then
                                        Return (Process1(ct2).MainWindowTitle)
                                    End If
                                Next
                                If blnUnique = True Then
                                    Exit For
                                End If
                            End If
                        Next
                    Catch ex As Exception

                    End Try
                    Return ""
                End Get
            End Property

#End Region
        End Class
    End Namespace

    Namespace StringFunctions

        Public Class Common

#Region "Enumeration"
            Enum enDirectio
                Left = 1
                Right = 2
            End Enum
            Enum enChars
                Number = 1
                BigLetter = 2
                SmallLetter = 4
                HelpChar = 8
                OperationChar = 16
                UpperChar = 32
                LowerChar = 64
                Space = 128
                Null = 256
                Dot = 512
                DoubleQoute = 1024
                SingleQoute = 2048
                CommaChars = 4096
            End Enum
#End Region
#Region "Internal Methods"
            Private Shared Function SHorof(ByVal X As Long) As String
                '         ÊÍæíá ÇáÃÑÞÇã Åáì äÕæÕ ÍÑÝíÉ
                Dim C As String = X.ToString("000000000000")


                Dim C1 As Short
                C1 = Short.Parse(C.Substring(11, 1))

                Dim Letter1 As String = ""
                Select Case C1
                    Case Is = 1 : Letter1 = "æÇÍÏ"
                    Case Is = 2 : Letter1 = "ÇËäÇä"
                    Case Is = 3 : Letter1 = "ËáÇËÉ"
                    Case Is = 4 : Letter1 = "ÇÑÈÚÉ"
                    Case Is = 5 : Letter1 = "ÎãÓÉ"
                    Case Is = 6 : Letter1 = "ÓÊÉ"
                    Case Is = 7 : Letter1 = "ÓÈÚÉ"
                    Case Is = 8 : Letter1 = "ËãÇäíÉ"
                    Case Is = 9 : Letter1 = "ÊÓÚÉ"
                End Select

                Dim C2 As Short
                C2 = Short.Parse(C.Substring(10, 1))
                Dim Letter2 As String = ""
                Select Case C2
                    Case Is = 1 : Letter2 = "ÚÔÑ"
                    Case Is = 2 : Letter2 = "ÚÔÑæä"
                    Case Is = 3 : Letter2 = "ËáÇËæä"
                    Case Is = 4 : Letter2 = "ÇÑÈÚæä"
                    Case Is = 5 : Letter2 = "ÎãÓæä"
                    Case Is = 6 : Letter2 = "ÓÊæä"
                    Case Is = 7 : Letter2 = "ÓÈÚæä"
                    Case Is = 8 : Letter2 = "ËãÇäæä"
                    Case Is = 9 : Letter2 = "ÊÓÚæä"
                End Select

                If Letter1 <> "" And C2 > 1 Then Letter2 = Letter1 & " æ" & Letter2
                If Letter2 = "" Then Letter2 = Letter1
                If C1 = 0 And C2 = 1 Then Letter2 = Letter2 & "É"
                If C1 = 1 And C2 = 1 Then Letter2 = "ÇÍÏì ÚÔÑ"
                If C1 = 2 And C2 = 1 Then Letter2 = "ÇËäì ÚÔÑ"
                If C1 > 2 And C2 = 1 Then Letter2 = Letter1 & " " & Letter2
                Dim C3 As Short

                C3 = Short.Parse(C.Substring(9, 1))
                Dim Letter3 As String = ""
                Select Case C3
                    Case Is = 1 : Letter3 = "ãÇÆÉ"
                    Case Is = 2 : Letter3 = "ãÆÊÇä"
                    Case Is > 2
                        Letter3 = SHorof(C3).Substring(0, SHorof(C3).Length - 1) & "ãÇÆÉ"
                End Select
                If Letter3 <> "" And Letter2 <> "" Then Letter3 = Letter3 & " æ" & Letter2
                If Letter3 = "" Then Letter3 = Letter2

                Dim C4 As Short
                C4 = Short.Parse(C.Substring(6, 3))
                Dim Letter4 As String = ""
                Select Case C4
                    Case Is = 1 : Letter4 = "ÇáÝ"
                    Case Is = 2 : Letter4 = "ÇáÝÇä"
                    Case 3 To 10
                        Letter4 = SHorof(C4) + " ÂáÇÝ"
                    Case Is > 10
                        Letter4 = SHorof(C4) + " ÇáÝ"
                End Select
                If Letter4 <> "" And Letter3 <> "" Then Letter4 = Letter4 & " æ" & Letter3
                If Letter4 = "" Then Letter4 = Letter3

                Dim C5 As Short
                C5 = Short.Parse(C.Substring(3, 3))

                Dim Letter5 As String = ""
                Select Case C5
                    Case Is = 1 : Letter5 = "ãáíæä"
                    Case Is = 2 : Letter5 = "ãáíæäÇä"
                    Case 3 To 10
                        Letter5 = SHorof(C5) + " ãáÇííä"
                    Case Is > 10
                        Letter5 = SHorof(C5) + " ãáíæä"
                End Select
                If Letter5 <> "" And Letter4 <> "" Then Letter5 = Letter5 & " æ" & Letter4
                If Letter5 = "" Then Letter5 = Letter4

                Dim C6 As Short
                C6 = Short.Parse(C.Substring(0, 3))
                Dim Letter6 As String = ""
                Select Case C6
                    Case Is = 1 : Letter6 = "ãáíÇÑ"
                    Case Is = 2 : Letter6 = "ãáíÇÑÇä"
                    Case Is > 2
                        Letter6 = SHorof(C6) + " ãáíÇÑ"
                End Select
                If Letter6 <> "" And Letter5 <> "" Then Letter6 = Letter6 & " æ" & Letter5
                If Letter6 = "" Then Letter6 = Letter5
                SHorof = Letter6
            End Function
            Private Shared Function Get3Num(ByVal Number As Short) As String
                If Number <= 0 Then Return ""
                Dim StrNum As String = Number.ToString("000")
                Dim TempStr As String = "", RetStr As String = ""


                If StrNum.Substring(0, 1) <> "0" Then
                    Select Case StrNum.Substring(0, 1)
                        Case "1" : TempStr = "one"
                        Case "2" : TempStr = "two"
                        Case "3" : TempStr = "three"
                        Case "4" : TempStr = "four"
                        Case "5" : TempStr = "five"
                        Case "6" : TempStr = "six"
                        Case "7" : TempStr = "seven"
                        Case "8" : TempStr = "eight"
                        Case "9" : TempStr = "nine"
                    End Select
                    TempStr += " hundred "
                End If
                RetStr = TempStr

                TempStr = ""
                If StrNum.Substring(1, 1) = "1" Then
                    Select Case StrNum.Substring(2, 1)
                        Case "0" : TempStr = "ten"
                        Case "1" : TempStr = "eleven"
                        Case "2" : TempStr = "twelve"
                        Case "3" : TempStr = "thirteen"
                        Case "4" : TempStr = "fourteen"
                        Case "5" : TempStr = "fifteen"
                        Case "6" : TempStr = "sixteen"
                        Case "7" : TempStr = "seventeen"
                        Case "8" : TempStr = "eighteen"
                        Case "9" : TempStr = "nineteen"
                    End Select
                Else

                    Select Case StrNum.Substring(1, 1)
                        Case "2" : TempStr = "twenty "
                        Case "3" : TempStr = "thirty "
                        Case "4" : TempStr = "forty "
                        Case "5" : TempStr = "fifty "
                        Case "6" : TempStr = "sixty "
                        Case "7" : TempStr = "seventy "
                        Case "8" : TempStr = "eighty "
                        Case "9" : TempStr = "ninety "
                    End Select

                    Select Case StrNum.Substring(2, 1)
                        Case "1" : TempStr += "one"
                        Case "2" : TempStr += "two"
                        Case "3" : TempStr += "three"
                        Case "4" : TempStr += "four"
                        Case "5" : TempStr += "five"
                        Case "6" : TempStr += "six"
                        Case "7" : TempStr += "seven"
                        Case "8" : TempStr += "eight"
                        Case "9" : TempStr += "nine"
                    End Select

                End If
                RetStr += TempStr
                Return RetStr


            End Function
#Region "Remove Chars"
            Private Function RemoveCommaChar(ByVal Value As String) As String
                Dim X As Integer, t As New System.Text.StringBuilder
                For X = 0 To Len(Value) - 1
                    Select Case Asc(CharOfString(Value, X))
                        Case 33, 34, 35, 39, 0, 41, 44, 59, 58, 96, 126, 63, 38
                        Case Else
                            t.Append(CharOfString(Value, X))
                    End Select
                Next
                Return t.ToString
            End Function
            Private Function RemoveNumber(ByVal Value As String) As String
                Dim X As Integer, t As New System.Text.StringBuilder
                For X = 0 To Len(Value) - 1
                    Select Case Asc(CharOfString(Value, X))
                        Case 48 To 57
                        Case Else
                            t.Append(CharOfString(Value, X))
                    End Select
                Next
                Return t.ToString
            End Function
            Private Function RemoveNull(ByVal Value As String) As String
                Dim X As Integer, t As New System.Text.StringBuilder
                For X = 0 To Len(Value) - 1
                    Select Case Asc(CharOfString(Value, X))
                        Case 0
                        Case Else
                            t.Append(CharOfString(Value, X))
                    End Select
                Next
                Return t.ToString
            End Function
            Private Function RemoveSpace(ByVal Value As String) As String
                Dim X As Integer, t As New System.Text.StringBuilder
                For X = 0 To Len(Value) - 1
                    Select Case (CharOfString(Value, X))
                        Case " "
                        Case Else
                            t.Append(CharOfString(Value, X))
                    End Select
                Next
                Return t.ToString
            End Function
            Private Function RemoveBigLetter(ByVal Value As String) As String
                Dim X As Integer, t As New System.Text.StringBuilder
                For X = 0 To Len(Value) - 1
                    Select Case Asc(CharOfString(Value, X))
                        Case 65 To 90
                        Case Else
                            t.Append(CharOfString(Value, X))
                    End Select
                Next
                Return t.ToString
            End Function
            Private Function RemoveSmallLetter(ByVal Value As String) As String
                Dim X As Integer, t As New System.Text.StringBuilder
                For X = 0 To Len(Value) - 1
                    Select Case Asc(CharOfString(Value, X))
                        Case 79 To 122
                        Case Else
                            t.Append(CharOfString(Value, X))
                    End Select
                Next
                Return t.ToString
            End Function
            Private Function RemoveDot(ByVal Value As String) As String
                Dim X As Integer, t As New System.Text.StringBuilder
                For X = 0 To Len(Value) - 1
                    Select Case (CharOfString(Value, X))
                        Case "."
                        Case Else
                            t.Append(CharOfString(Value, X))
                    End Select
                Next
                Return t.ToString
            End Function
            Private Function RemoveHelperChar(ByVal Value As String) As String
                Dim X As Integer, t As New System.Text.StringBuilder
                For X = 0 To Len(Value) - 1
                    Select Case Asc(CharOfString(Value, X))
                        Case 33 To 41
                        Case 123 To 125
                        Case 95, 96, 63, 64
                        Case 91 To 93
                        Case Else
                            t.Append(CharOfString(Value, X))
                    End Select
                Next
                Return t.ToString
            End Function
            Private Function RemoveOperationChar(ByVal Value As String) As String
                Dim X As Integer, t As New System.Text.StringBuilder
                For X = 0 To Len(Value) - 1
                    Select Case Asc(CharOfString(Value, X))
                        Case 42, 43, 45, 47, 60, 61, 62, 94
                        Case Else
                            t.Append(CharOfString(Value, X))
                    End Select
                Next
                Return t.ToString
            End Function
            Private Function RemoveUpperChar(ByVal Value As String) As String
                Dim X As Integer, t As New System.Text.StringBuilder
                For X = 0 To Len(Value) - 1
                    Select Case Asc(CharOfString(Value, X))
                        Case 1 To 32
                        Case Else
                            t.Append(CharOfString(Value, X))
                    End Select
                Next
                Return t.ToString
            End Function
            Private Function RemoveLowerChar(ByVal Value As String) As String
                Dim X As Integer, t As New System.Text.StringBuilder
                For X = 0 To Len(Value) - 1
                    Select Case Asc(CharOfString(Value, X))
                        Case 126 To 255
                        Case Else
                            t.Append(CharOfString(Value, X))
                    End Select
                Next
                Return t.ToString
            End Function
            Private Function RemoveDoubleQoute(ByVal Value As String) As String
                Dim X As Integer, t As New System.Text.StringBuilder
                For X = 0 To Len(Value) - 1
                    Select Case Asc(CharOfString(Value, X))
                        Case 34
                        Case Else
                            t.Append(CharOfString(Value, X))
                    End Select
                Next
                Return t.ToString
            End Function
            Private Function RemoveSingleQoute(ByVal Value As String) As String
                Dim X As Integer, t As New System.Text.StringBuilder
                For X = 0 To Len(Value) - 1
                    Select Case Asc(CharOfString(Value, X))
                        Case 39
                        Case Else
                            t.Append(CharOfString(Value, X))
                    End Select
                Next
                Return t.ToString
            End Function
#End Region
#End Region

#Region "Methods"
            Public Shared Function Arabic_NumToText(ByVal X As Double, Optional ByVal MainUnit As String = "Ìäíå", Optional ByVal SmallUnit As String = "ÞÑÔ") As String

                If X <= 0 Then Return ""
                If X > 999999999999 Then Return ""

                Dim N As Long = Long.Parse(X.ToString("000000000000.00").Substring(0, 12))

                Dim B As Short
                B = Short.Parse(X.ToString("000000000000.00").Substring(13, 2))

                Dim R As String = SHorof(N)

                Dim Result As String = ""
                If R <> "" And B > 0 Then Result = R & " " & MainUnit & " æ " & B & " " & SmallUnit
                If R <> "" And B = 0 Then Result = R & " " & MainUnit
                If R = "" And B <> 0 Then Result = B & " " & SmallUnit
                Arabic_NumToText = Result

            End Function
            Public Shared Function English_NumToText(ByVal Number As Double, Optional ByVal MainUnit As String = "Pound", Optional ByVal SmallUnit As String = "Pets") As String

                If Number <= 0 Then Return ""
                If Number > 999999999999 Then Return ""

                Dim StrNum As String = Number.ToString("000000000000.00")
                Dim TempStr As String, RetStr As String

                Dim StrBillion As String = StrNum.Substring(0, 3)
                Dim StrMillion As String = StrNum.Substring(3, 3)
                Dim StrThousand As String = StrNum.Substring(6, 3)
                Dim StrHandred As String = StrNum.Substring(9, 3)


                TempStr = Get3Num(Short.Parse(StrBillion))
                If TempStr <> "" Then TempStr += " Billion"
                RetStr = TempStr

                TempStr = ""
                TempStr = Get3Num(Short.Parse(StrMillion))
                If TempStr <> "" Then TempStr += " Million"
                If RetStr <> "" Then RetStr += " , " & TempStr Else RetStr = TempStr

                TempStr = ""
                TempStr = Get3Num(Short.Parse(StrThousand))
                If TempStr <> "" Then TempStr += " Thousand"
                If RetStr <> "" Then RetStr += " , " & TempStr Else RetStr = TempStr

                TempStr = ""
                TempStr = Get3Num(Short.Parse(StrHandred))
                If RetStr <> "" Then RetStr += " , " & TempStr Else RetStr = TempStr


                RetStr += " " & MainUnit

                If Short.Parse(StrNum.Substring(13, 2)) > 0 Then RetStr += " and " & StrNum.Substring(13, 2) & " " & SmallUnit

                RetStr = RetStr.Trim
                RetStr = RetStr.Substring(0, 1).ToUpper & RetStr.Substring(1)

                Return RetStr

            End Function
            Public Function RemoveMiddleSpace(ByVal Exp As String) As String
                Dim temp() As String
                Dim t As String
                t = ""
                Dim x As Integer
                temp = Split(Exp, " ")
                For x = 0 To UBound(temp)
                    If Trim(temp(x)) <> "" Then
                        t = t & Trim(temp(x))
                    End If
                Next x
                RemoveMiddleSpace = Trim(t)
            End Function
            Public Function RemoveOverMiddleSpace(ByVal str1 As String) As String
                Dim temp() As String
                Dim t As String
                t = ""
                Dim x As Integer
                temp = Split(str1, " ")
                For x = 0 To UBound(temp)
                    If Trim(temp(x)) <> "" Then
                        t = t & Trim(temp(x)) & " "
                    End If
                Next x
                RemoveOverMiddleSpace = Trim(t)
            End Function
            Public Overloads Function Equal(ByVal str1 As String, ByVal str2 As String) As Boolean
                Dim X As Int64
                If Len(str1) = Len(str2) Then
                    For X = 0 To Len(str1) - 1
                        If Asc(str1.Chars(X)) <> Asc(str2.Chars(X)) Then
                            Return False
                        End If
                    Next
                Else
                    Return False
                End If
                Return True
            End Function
            Public Overloads Function Equal(ByVal str1 As String, ByVal str2 As String, ByVal IgnoreCase As Boolean) As Boolean
                Dim X As Int64
                If IgnoreCase Then
                    str1 = str1.ToLower
                    str2 = str2.ToLower
                End If
                If Len(str1) = Len(str2) Then
                    For X = 0 To Len(str1) - 1
                        If Asc(str1.Chars(X)) <> Asc(str2.Chars(X)) Then
                            Return False
                        End If
                    Next
                Else
                    Return False
                End If
                Return True
            End Function
            Public Function CharOfString(ByVal Value As String, ByVal Index As Integer) As String
                Try
                    If Index < -1 Then Return ""
                    If Index > Len(Value) - 1 Then Return ""
                    If Value <> "" Then
                        Return Value.Chars(Index)
                    End If
                Catch ex As Exception
                    Return ""
                End Try
                Return ""
            End Function
            Public Function Invers(ByVal Value As String) As String
                Dim X As Integer, temp(Len(Value)) As String, temp2 As String = ""
                For X = 0 To Len(Value) - 1
                    temp(X) = Value.Chars(X)
                Next
                For X = Len(Value) - 1 To 0 Step -1
                    temp2 &= Value.Chars(X)
                Next
                Return temp2
            End Function
            Public Function FormatDigits(ByVal Value As String, ByVal Digits As Integer, Optional ByVal ComplateChar As String = "", Optional ByVal Direction As enDirectio = enDirectio.Left) As String
                Dim x, y As Integer, temp As String = ""
                Select Case Len(Value)
                    Case Is = Digits
                        Return Value
                    Case Is < Digits
                        For x = 1 To Digits - Len(Value)
                            temp &= ComplateChar
                        Next
                        Select Case Direction
                            Case enDirectio.Left
                                temp &= Value
                            Case enDirectio.Right
                                temp = Value & temp
                        End Select
                    Case Is > Digits
                        Select Case Direction
                            Case enDirectio.Left
                                For x = 0 To Digits - 1
                                    temp &= Value.Chars(x)
                                Next
                            Case enDirectio.Right
                                y = Len(Value) - 1
                                For x = Digits - 1 To 0 Step -1
                                    temp &= Value.Chars(y)
                                    y -= 1
                                Next
                                temp = Invers(temp)
                        End Select
                End Select
                Return temp
            End Function
            Public Overloads Function SubString(ByVal Value As String, ByVal SubLen As Integer, Optional ByVal Direction As enDirectio = enDirectio.Left) As String
                If SubLen >= Len(Value) Then
                    Return ""
                End If
                Dim X, Y As Integer, temp(Len(Value))
                Dim temp2 As String = ""
                For X = 0 To Len(Value) - 1
                    temp(X) = Value.Chars(X)
                Next
                Select Case Direction
                    Case enDirectio.Left
                        Y = UBound(temp) - 1
                        For X = 1 To Len(Value) - SubLen
                            temp2 &= temp(Y)
                            Y -= 1
                        Next
                        temp2 = Invers(temp2)
                    Case enDirectio.Right
                        For X = 1 To Len(Value) - SubLen
                            temp2 &= temp(X - 1)
                        Next
                End Select
                Return temp2
            End Function
            Public Overloads Function SubString(ByVal Value As String, ByVal Subs As String) As String
                Dim BeginIndex As Integer = Value.IndexOf(Subs)
                Dim temp() As String, X As Integer
                If BeginIndex = -1 Then
                    Return Value
                End If
                temp = Split(Value, Subs)
                Value = ""
                For X = 0 To UBound(temp)
                    If Not Equal(Subs, temp(X)) Then
                        Value &= temp(X)
                    End If
                Next
                Return Value
            End Function
            Public Function PartString(ByVal Value As String, ByVal PartDelimer As String, ByVal PartIndex As Integer) As String
                If Value = "" Or PartIndex <= -1 Or PartDelimer = "" Then
                    Return ""
                End If
                Dim temp() As String
                temp = Split(Value, PartDelimer)
                If PartIndex > UBound(temp) Then
                    Return ""
                End If
                Return temp(PartIndex)
            End Function
            Public Function RemoveString(ByVal Value As String, ByVal RemoveType As enChars) As String
                Select Case RemoveType
                    Case enChars.Number
                        Return RemoveNumber(Value)
                    Case enChars.Null
                        Return RemoveNull(Value)
                    Case enChars.Space
                        Return RemoveSpace(Value)
                    Case enChars.BigLetter
                        Return RemoveBigLetter(Value)
                    Case enChars.SmallLetter
                        Return RemoveSmallLetter(Value)
                    Case enChars.Dot
                        Return RemoveDot(Value)
                    Case enChars.HelpChar
                        Return RemoveHelperChar(Value)
                    Case enChars.OperationChar
                        Return RemoveOperationChar(Value)
                    Case enChars.UpperChar
                        Return RemoveUpperChar(Value)
                    Case enChars.LowerChar
                        Return RemoveLowerChar(Value)
                    Case enChars.DoubleQoute
                        Return RemoveDoubleQoute(Value)
                    Case enChars.SingleQoute
                        Return RemoveSingleQoute(Value)
                    Case enChars.CommaChars
                        Return RemoveCommaChar(Value)
                    Case enChars.Null + enChars.Space + enChars.HelpChar + enChars.UpperChar + enChars.LowerChar
                        Dim temp As String
                        temp = RemoveNull(Value)
                        temp = RemoveSpace(temp)
                        temp = RemoveHelperChar(temp)
                        temp = RemoveUpperChar(temp)
                        temp = RemoveLowerChar(temp)
                        Return temp
                    Case enChars.DoubleQoute + enChars.SingleQoute
                        Return RemoveDoubleQoute(RemoveSingleQoute(Value))
                End Select
                Return ""
            End Function
            Public Function GetFileName(ByVal Path As String, Optional WithoutExtention As Boolean = False, Optional ByVal PathSeparator As String = "\") As String
                Try
                    Dim temp() As String = Split(Path, PathSeparator)
                    If Not WithoutExtention Then
                        Return temp(UBound(temp))
                    Else
                        temp = Split(temp(UBound(temp)), ".")
                        Return temp(0)
                    End If
                Catch ex As Exception
                End Try
                Return ""
            End Function
#End Region

        End Class

        Public Class StringsFunction

#Region "Enumeration"
            Enum enDirectio
                Left = 1
                Right = 2
            End Enum
            Enum enChars
                Number = 1
                BigLetter = 2
                SmallLetter = 4
                HelpChar = 8
                OperationChar = 16
                UpperChar = 32
                LowerChar = 64
                Space = 128
                Null = 256
                Dot = 512
                DoubleQoute = 1024
                SingleQoute = 2048
                CommaChars = 4096
            End Enum
#End Region

#Region "Remove Chars"
            Private Function RemoveNumber(ByVal Value As String) As String
                Dim X As Integer, t As New System.Text.StringBuilder
                For X = 0 To Len(Value) - 1
                    Select Case Asc(CharOfString(Value, X))
                        Case 48 To 57
                        Case Else
                            t.Append(CharOfString(Value, X))
                    End Select
                Next
                Return t.ToString
            End Function
            Private Function RemoveNull(ByVal Value As String) As String
                Dim X As Integer, t As New System.Text.StringBuilder
                For X = 0 To Len(Value) - 1
                    Select Case Asc(CharOfString(Value, X))
                        Case 0
                        Case Else
                            t.Append(CharOfString(Value, X))
                    End Select
                Next
                Return t.ToString
            End Function
            Private Function RemoveSpace(ByVal Value As String) As String
                Dim X As Integer, t As New System.Text.StringBuilder
                For X = 0 To Len(Value) - 1
                    Select Case (CharOfString(Value, X))
                        Case " "
                        Case Else
                            t.Append(CharOfString(Value, X))
                    End Select
                Next
                Return t.ToString
            End Function
            Private Function RemoveBigLetter(ByVal Value As String) As String
                Dim X As Integer, t As New System.Text.StringBuilder
                For X = 0 To Len(Value) - 1
                    Select Case Asc(CharOfString(Value, X))
                        Case 65 To 90
                        Case Else
                            t.Append(CharOfString(Value, X))
                    End Select
                Next
                Return t.ToString
            End Function
            Private Function RemoveSmallLetter(ByVal Value As String) As String
                Dim X As Integer, t As New System.Text.StringBuilder
                For X = 0 To Len(Value) - 1
                    Select Case Asc(CharOfString(Value, X))
                        Case 79 To 122
                        Case Else
                            t.Append(CharOfString(Value, X))
                    End Select
                Next
                Return t.ToString
            End Function
            Private Function RemoveDot(ByVal Value As String) As String
                Dim X As Integer, t As New System.Text.StringBuilder
                For X = 0 To Len(Value) - 1
                    Select Case (CharOfString(Value, X))
                        Case "."
                        Case Else
                            t.Append(CharOfString(Value, X))
                    End Select
                Next
                Return t.ToString
            End Function
            Private Function RemoveHelperChar(ByVal Value As String) As String
                Dim X As Integer, t As New System.Text.StringBuilder
                For X = 0 To Len(Value) - 1
                    Select Case Asc(CharOfString(Value, X))
                        Case 33 To 41
                        Case 123 To 125
                        Case 95, 96, 63, 64
                        Case 91 To 93
                        Case Else
                            t.Append(CharOfString(Value, X))
                    End Select
                Next
                Return t.ToString
            End Function
            Private Function RemoveOperationChar(ByVal Value As String) As String
                Dim X As Integer, t As New System.Text.StringBuilder
                For X = 0 To Len(Value) - 1
                    Select Case Asc(CharOfString(Value, X))
                        Case 42, 43, 45, 47, 60, 61, 62, 94
                        Case Else
                            t.Append(CharOfString(Value, X))
                    End Select
                Next
                Return t.ToString
            End Function
            Private Function RemoveCommaChar(ByVal Value As String) As String
                Dim X As Integer, t As New System.Text.StringBuilder
                For X = 0 To Len(Value) - 1
                    Select Case Asc(CharOfString(Value, X))
                        Case 33, 34, 35, 39, 0, 41, 44, 59, 58, 96, 126, 63, 38
                        Case Else
                            t.Append(CharOfString(Value, X))
                    End Select
                Next
                Return t.ToString
            End Function
            Private Function RemoveUpperChar(ByVal Value As String) As String
                Dim X As Integer, t As New System.Text.StringBuilder
                For X = 0 To Len(Value) - 1
                    Select Case Asc(CharOfString(Value, X))
                        Case 1 To 32
                        Case Else
                            t.Append(CharOfString(Value, X))
                    End Select
                Next
                Return t.ToString
            End Function
            Private Function RemoveLowerChar(ByVal Value As String) As String
                Dim X As Integer, t As New System.Text.StringBuilder
                For X = 0 To Len(Value) - 1
                    Select Case Asc(CharOfString(Value, X))
                        Case 126 To 255
                        Case Else
                            t.Append(CharOfString(Value, X))
                    End Select
                Next
                Return t.ToString
            End Function
            Private Function RemoveDoubleQoute(ByVal Value As String) As String
                Dim X As Integer, t As New System.Text.StringBuilder
                For X = 0 To Len(Value) - 1
                    Select Case Asc(CharOfString(Value, X))
                        Case 34
                        Case Else
                            t.Append(CharOfString(Value, X))
                    End Select
                Next
                Return t.ToString
            End Function
            Private Function RemoveSingleQoute(ByVal Value As String) As String
                Dim X As Integer, t As New System.Text.StringBuilder
                For X = 0 To Len(Value) - 1
                    Select Case Asc(CharOfString(Value, X))
                        Case 39
                        Case Else
                            t.Append(CharOfString(Value, X))
                    End Select
                Next
                Return t.ToString
            End Function
#End Region

#Region "Methods"
            Public Function GetFileContent(ByVal FilePath As String) As String
                Try
                    Dim obj As New System.IO.StreamReader(FilePath)
                    Dim sql As String = ""
                    Sql = obj.ReadToEnd
                    obj.Close()
                    Return sql
                Catch ex As Exception
                    Return ""
                End Try
            End Function
            Public Function RemoveMiddleSpace(ByVal Exp As String) As String
                Dim temp() As String
                Dim t As String
                t = ""
                Dim x As Integer
                temp = Split(Exp, " ")
                For x = 0 To UBound(temp)
                    If Trim(temp(x)) <> "" Then
                        t = t & Trim(temp(x))
                    End If
                Next x
                RemoveMiddleSpace = Trim(t)
            End Function
            Public Function RemoveOverMiddleSpace(ByVal str1 As String) As String
                Dim temp() As String
                Dim t As String
                t = ""
                Dim x As Integer
                temp = Split(str1, " ")
                For x = 0 To UBound(temp)
                    If Trim(temp(x)) <> "" Then
                        t = t & Trim(temp(x)) & " "
                    End If
                Next x
                RemoveOverMiddleSpace = Trim(t)
            End Function
            Public Overloads Function Equal(ByVal str1 As String, ByVal str2 As String) As Boolean
                Dim X As Int64
                If Len(str1) = Len(str2) Then
                    For X = 0 To Len(str1) - 1
                        If Asc(str1.Chars(X)) <> Asc(str2.Chars(X)) Then
                            Return False
                        End If
                    Next
                Else
                    Return False
                End If
                Return True
            End Function
            Public Overloads Function Equal(ByVal str1 As String, ByVal str2 As String, ByVal IgnoreCase As Boolean) As Boolean
                Dim X As Int64
                If IgnoreCase Then
                    str1 = str1.ToLower
                    str2 = str2.ToLower
                End If
                If Len(str1) = Len(str2) Then
                    For X = 0 To Len(str1) - 1
                        If Asc(str1.Chars(X)) <> Asc(str2.Chars(X)) Then
                            Return False
                        End If
                    Next
                Else
                    Return False
                End If
                Return True
            End Function
            Public Function CharOfString(ByVal Value As String, ByVal Index As Integer) As String
                Try
                    If Index < -1 Then Return ""
                    If Index > Len(Value) - 1 Then Return ""
                    If Value <> "" Then
                        Return Value.Chars(Index)
                    End If
                Catch ex As Exception
                    Return ""
                End Try
                Return ""
            End Function
            Public Function Invers(ByVal Value As String) As String
                Dim X As Integer, temp(Len(Value)) As String, temp2 As String = ""
                For X = 0 To Len(Value) - 1
                    temp(X) = Value.Chars(X)
                Next
                For X = Len(Value) - 1 To 0 Step -1
                    temp2 &= Value.Chars(X)
                Next
                Return temp2
            End Function
            Public Function FormatDigits(ByVal Value As String, ByVal Digits As Integer, Optional ByVal ComplateChar As String = "", Optional ByVal Direction As enDirectio = enDirectio.Left) As String
                Dim x, y As Integer, temp As String = ""
                Select Case Len(Value)
                    Case Is = Digits
                        Return Value
                    Case Is < Digits
                        For x = 1 To Digits - Len(Value)
                            temp &= ComplateChar
                        Next
                        Select Case Direction
                            Case enDirectio.Left
                                temp &= Value
                            Case enDirectio.Right
                                temp = Value & temp
                        End Select
                    Case Is > Digits
                        Select Case Direction
                            Case enDirectio.Left
                                For x = 0 To Digits - 1
                                    temp &= Value.Chars(x)
                                Next
                            Case enDirectio.Right
                                y = Len(Value) - 1
                                For x = Digits - 1 To 0 Step -1
                                    temp &= Value.Chars(y)
                                    y -= 1
                                Next
                                temp = Invers(temp)
                        End Select
                End Select
                Return temp
            End Function
            Public Overloads Function SubString(ByVal Value As String, ByVal SubLen As Integer, Optional ByVal Direction As enDirectio = enDirectio.Left) As String
                If SubLen >= Len(Value) Then
                    Return ""
                End If
                Dim X, Y As Integer, temp(Len(Value)) As String
                Dim temp2 As String = ""
                For X = 0 To Len(Value) - 1
                    temp(X) = Value.Chars(X)
                Next
                Select Case Direction
                    Case enDirectio.Left
                        Y = UBound(temp) - 1
                        For X = 1 To Len(Value) - SubLen
                            temp2 &= temp(Y)
                            Y -= 1
                        Next
                        temp2 = Invers(temp2)
                    Case enDirectio.Right
                        For X = 1 To Len(Value) - SubLen
                            temp2 &= temp(X - 1)
                        Next
                End Select
                Return temp2
            End Function
            Public Overloads Function SubString(ByVal Value As String, ByVal Subs As String) As String
                Dim BeginIndex As Integer = Value.IndexOf(Subs)
                Dim temp() As String, X As Integer
                If BeginIndex = -1 Then
                    Return Value
                End If
                temp = Split(Value, Subs)
                Value = ""
                For X = 0 To UBound(temp)
                    If Not Equal(Subs, temp(X)) Then
                        Value &= temp(X)
                    End If
                Next
                Return Value
            End Function
            Public Function PartString(ByVal Value As String, ByVal PartDelimer As String, ByVal PartIndex As Integer) As String
                If Value = "" Or PartIndex <= -1 Or PartDelimer = "" Then
                    Return ""
                End If
                Dim temp() As String
                temp = Split(Value, PartDelimer)
                If PartIndex > UBound(temp) Then
                    Return ""
                End If
                Return temp(PartIndex)
            End Function
            Public Function RemoveString(ByVal Value As String, ByVal RemoveType As enChars) As String
                Select Case RemoveType
                    Case enChars.Number
                        Return RemoveNumber(Value)
                    Case enChars.Null
                        Return RemoveNull(Value)
                    Case enChars.Space
                        Return RemoveSpace(Value)
                    Case enChars.BigLetter
                        Return RemoveBigLetter(Value)
                    Case enChars.SmallLetter
                        Return RemoveSmallLetter(Value)
                    Case enChars.Dot
                        Return RemoveDot(Value)
                    Case enChars.HelpChar
                        Return RemoveHelperChar(Value)
                    Case enChars.OperationChar
                        Return RemoveOperationChar(Value)
                    Case enChars.UpperChar
                        Return RemoveUpperChar(Value)
                    Case enChars.LowerChar
                        Return RemoveLowerChar(Value)
                    Case enChars.DoubleQoute
                        Return RemoveDoubleQoute(Value)
                    Case enChars.SingleQoute
                        Return RemoveSingleQoute(Value)
                    Case enChars.CommaChars
                        Return RemoveCommaChar(Value)
                    Case enChars.Null + enChars.Space + enChars.HelpChar + enChars.UpperChar + enChars.LowerChar
                        Dim temp As String
                        temp = RemoveNull(Value)
                        temp = RemoveSpace(temp)
                        temp = RemoveHelperChar(temp)
                        temp = RemoveUpperChar(temp)
                        temp = RemoveLowerChar(temp)
                        Return temp
                    Case enChars.DoubleQoute + enChars.SingleQoute
                        Return RemoveDoubleQoute(RemoveSingleQoute(Value))
                End Select
                Return ""
            End Function
            Public Function GetFileName(ByVal Path As String, Optional ByVal PathSeparator As String = "\") As String
                Try
                    Dim temp() As String = Split(Path, PathSeparator)
                    Dim x As Integer = 0
                    temp = Split(temp(UBound(temp)), ".")
                    Return temp(0)
                Catch ex As Exception
                End Try
                Return ""
            End Function

            Public Shared Function GetAppPath() As String
                Dim l_intCharPos As Integer = 0, l_intReturnPos As Integer
                Dim l_strAppPath As String

                l_strAppPath = System.Reflection.Assembly.GetExecutingAssembly.Location()

                While (1)
                    l_intCharPos = InStr(l_intCharPos + 1, l_strAppPath, "\", CompareMethod.Text)
                    If l_intCharPos = 0 Then
                        If Right(Mid(l_strAppPath, 1, l_intReturnPos), 1) <> "\" Then
                            Return Mid(l_strAppPath, 1, l_intReturnPos) & "\"
                        Else
                            Return Mid(l_strAppPath, 1, l_intReturnPos)
                        End If
                        Exit Function
                    End If
                    l_intReturnPos = l_intCharPos
                End While
                Return ""
            End Function
#End Region



        End Class

    End Namespace


    Namespace Multimedia
        Friend Class Interop
            <DllImport("kernel32.dll", SetLastError:=True, CharSet:=CharSet.Auto)> _
            Public Shared Function GetShortPathName(ByVal longPath As String, <MarshalAs(UnmanagedType.LPTStr)> ByVal ShortPath As StringBuilder, <MarshalAs(UnmanagedType.U4)> ByVal bufferSize As Integer) As Integer
            End Function
        End Class
        Public Class VedioPlayer
            Private Declare Function mciSendString Lib "winmm" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Integer, ByVal hwndCallback As Integer) As Integer
            Const WS_CHILD As Integer = &H40000000

            Public Enum AudioChannel
                LeftOnly = 1
                RightOnly = 2
                LeftRight = 3
            End Enum

            Private AB_IsOpen As Boolean

            Private RetVal As Integer
            Private mWidth, mHeight As Single


            Public Sub Open(ByVal mFile As String, ByVal pic As PictureBox)
                ' ÝÊÍ ÇáãáÝ
                Dim CommandString As String
                Dim mShortName As String
                Dim Ext, MedTp As String
                mFile = mFile.Trim

                Ext = IO.Path.GetExtension(mFile)
                'Ext = mFile.Substring(mFile.Length - 3, 3)

                Select Case Ext
                    Case "cda"
                        MedTp = "videodisc"
                        'Case "avi"
                        '    MedTp = "avivideo"
                    Case Else
                        MedTp = "MPEGVideo"
                End Select

                Dim sb As New StringBuilder(1024)
                RetVal = Interop.GetShortPathName(mFile, sb, 1024)
                If RetVal <> 0 Then
                    mShortName = sb.ToString()
                Else
                    mShortName = mFile
                End If

                CommandString = "Open " & mShortName & " type " & MedTp & " alias AVIFile parent " & CStr(pic.Handle.ToInt32) & " style " & CStr(WS_CHILD)
                RetVal = mciSendString(CommandString, vbNullString, 0, 0)
                If RetVal = 0 Then AB_IsOpen = True

                ' áãÚÑÝÉ ÇáÚÑÖ æÇáÅÑÊÝÇÚ ááãáÝ

                Dim RetString As String = New String(" ", 255)
                Dim res() As String

                CommandString = "where AVIFile destination"
                RetVal = mciSendString(CommandString, RetString, Len(RetString), 0)

                RetString = Left(RetString, InStr(RetString & vbNullChar, vbNullChar) - 1)
                res = Split(RetString)
                mWidth = CSng(res(2))
                mHeight = CSng(res(3))

            End Sub

            Public Sub Close()
                ' áÛáÞ ÇáãáÝ
                RetVal = mciSendString("close AVIFile", CStr(0), 0, 0)
                If RetVal = 0 Then AB_IsOpen = False
                mWidth = 0
                mHeight = 0
            End Sub

            Public Sub Play()
                ' ÊÔÛíá ÇáãáÝ
                RetVal = mciSendString("play AVIFile", CStr(0), 0, 0)
            End Sub

            Public Sub Pause()
                ' ááÅäÊÙÇÑ ÇáãÄÞÊ
                RetVal = mciSendString("pause AVIFile", CStr(0), 0, 0)
            End Sub

            Public Sub [Stop]()
                ' ááÊæÞÝ æÇáÚæÏÉ áÃæá ÇáãáÝ
                RetVal = mciSendString("stop AVIFile", CStr(0), 0, 0)
                RetVal = mciSendString("seek AVIFile to start", CStr(0), 0, 0)
            End Sub

            Public ReadOnly Property TotalTime() As Integer
                Get
                    ' áãÚÑÝÉ Øæá ÇáÝÊÑÉ ÇáÒãäíÉ ááãáÝ
                    Dim strReturn As String = New String(" ", 255)
                    RetVal = mciSendString("set AVIFile time format milliseconds", CStr(0), 0, 0)
                    RetVal = mciSendString("status AVIFile length", strReturn, 255, 0)
                    Return Integer.Parse(strReturn)
                End Get
            End Property


            Public Property CurrentPosition() As Integer
                Get
                    ' áãÚÑÝÉ ãßÇä ÇáÊÔÛíá ÇáÍÇáí ÏÇÎá ÇáãáÝ
                    Dim strReturn As String = New String(" ", 255)
                    RetVal = mciSendString("set AVIFile time format milliseconds", CStr(0), 0, 0)
                    RetVal = mciSendString("status AVIFile position", strReturn, 255, 0)
                    Return Integer.Parse(strReturn)
                End Get
                Set(ByVal Value As Integer)
                    ' ááÅäÊÞÇá Åáì äÞØÉ ÏÇÎá ÇáãáÝ
                    Dim RetVal As Integer
                    RetVal = mciSendString("set AVIFile time format milliseconds", CStr(0), 0, 0)
                    RetVal = mciSendString("seek AVIFile to " & Value, CStr(0), 0, 0)
                End Set
            End Property

            WriteOnly Property Channel() As AudioChannel
                Set(ByVal mChannel As AudioChannel)
                    ' áÊÍÏíÏ ãÎÑÌ ÇáÕæÊ
                    ' ÇáÓãÇÚÉ ÇáíÓÑì ¡ Çáíãäì Ãã ÇáÅËäíä
                    Select Case mChannel
                        Case AudioChannel.LeftOnly ' LEFT_ONLY
                            RetVal = mciSendString("set AVIFile audio right off", CStr(0), 0, 0)
                            RetVal = mciSendString("set AVIFile audio left on", CStr(0), 0, 0)
                        Case AudioChannel.RightOnly ' RIGHT_ONLY
                            RetVal = mciSendString("set AVIFile audio right on", CStr(0), 0, 0)
                            RetVal = mciSendString("set AVIFile audio left off", CStr(0), 0, 0)
                        Case AudioChannel.LeftRight ' RIGHT_LEFT
                            RetVal = mciSendString("set AVIFile audio right on", CStr(0), 0, 0)
                            RetVal = mciSendString("set AVIFile audio left on", CStr(0), 0, 0)
                    End Select

                End Set

            End Property

            WriteOnly Property Volume() As Short
                Set(ByVal m_Volume As Short)
                    ' áÖÈØ ÇáÕæÊ
                    Dim Ret As Integer
                    Ret = mciSendString("setaudio AVIFile volume to " & Str(m_Volume), CStr(0), 0, 0)
                End Set

            End Property

            Public Sub Fill(ByVal mFill As Boolean, ByVal pic As PictureBox)
                ' áÅÚÇÏÉ ÊÍÌíã ÇáãáÝ ÏÇÎá ÇáäÇÝÐÉ
                If mFill = False Then Exit Sub
                RetVal = mciSendString("Put AVIFile window at 0 0 " & CStr(pic.Width) & " " & CStr(pic.Height), vbNullString, 0, 0)
            End Sub

            ReadOnly Property PlayHieght() As Single
                Get
                    ' áãÚÑÝÉ ÇáÅÑÊÝÇÚ ÇáÐí íÚãá Úáíå ÇáãáÝ ÇáÂä
                    Return mHeight
                End Get
            End Property
            ReadOnly Property PlayWidth() As Single
                Get
                    ' áãÚÑÝÉ ÇáÚÑÖ ÇáÐí íÚãá Úáíå ÇáãáÝ ÇáÂä
                    Return mWidth
                End Get
            End Property

        End Class
        Public Class AudioPlayer

            'Send MCi Commands
            Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Integer, ByVal hwndCallback As Integer) As Integer
            'Get MCI Error Description
            Private Declare Function mciGetErrorString Lib "winmm.dll" Alias "mciGetErrorStringA" (ByVal dwError As Integer, ByVal lpstrBuffer As String, ByVal uLength As Integer) As Integer

            Dim m_strFilename, MedTp, Alias_Renamed As String
            Dim m_lngTimeFormat, m_lngMediaType As Integer
            Dim m_CurrentPositionInSeconds, m_Channel As Integer
            Dim m_Volume As Short

            '=================================================================
            '                       CONSTANTS - SELF DEFINED
            '
            '=================================================================

            Private Enum TimeFormats
                FRAMES = 300
                MILISECONDS = 301
                MSF = 302
                TMSF = 303
                HMS = 304
                SAMPLES = 305
                BYTES = 306
            End Enum
            Public Enum MediaType
                CD_AUDIO = 100
                SYNTHESIZER = 101
                MP3 = 102
                VIDEO_CD = 103
                WAVEAUDIO = 104
                MPEG_VIDEO = 105
                AVI_VIDEO = 106
            End Enum

            'Channels
            Public Enum Channels
                RIGHT_LEFT = 200
                RIGHT_ONLY = 201
                LEFT_ONLY = 202
            End Enum
            Public Enum MediaStatus
                TOTAL_FRAMES = 400
                TOTAL_TIME_IN_MILISECONDS = 401
                CURRENT_POSITION_IN_FRAMES = 402
                CURRENT_POSITION_IN_MILISECONDS = 403
            End Enum
            Public Enum CurrentState
                READY = 500
                PLAYING = 501
                PAUSED = 502
                END_OF_FILE = 503
                MCI_ERROR = 504
            End Enum

            ' member variable for State property
            Private m_State As Integer

            Public Sub Open(ByVal strFilename As String)
                m_strFilename = strFilename.Trim
                'Exit this function if the file doesn't exist
                If IO.File.Exists(m_strFilename) = False Then Exit Sub

                FindFormat()

                Select Case m_lngMediaType
                    Case MediaType.CD_AUDIO
                        MedTp = "cdaudio"
                        Alias_Renamed = "CD"
                    Case MediaType.MP3
                        MedTp = "MPEGVideo"
                        Alias_Renamed = "MP3"
                    Case MediaType.SYNTHESIZER
                        MedTp = "sequencer"
                        Alias_Renamed = "MIDI"
                    Case MediaType.WAVEAUDIO
                        MedTp = "waveaudio"
                        Alias_Renamed = "WAV"
                    Case Else
                        MedTp = "MPEGVideo"
                        Alias_Renamed = "MPEG"
                End Select

                Dim sb As New StringBuilder(1024)
                Dim RetVal As Integer = Interop.GetShortPathName(m_strFilename, sb, 1024)
                If RetVal <> 0 Then
                    m_strFilename = sb.ToString()
                End If

                If m_lngMediaType = MediaType.CD_AUDIO Then
                    'Ext = LCase(Right(m_strFilename, 6))
                    'dr = LCase(Left(Ext, 2))
                    'dr = Right(dr, 1)

                    mciSendString("open " & Chr(34) & m_strFilename & Chr(34) & " type cdaudio alias CD wait shareable", CStr(0), 0, 0)
                    mciSendString("set CD time format tmsf wait", CStr(0), 0, 0)

                    m_lngTimeFormat = TimeFormats.TMSF

                    Exit Sub

                End If

                'Open the file
                RetVal = mciSendString("open " & Chr(34) & m_strFilename & Chr(34) & " type " & Trim(MedTp) & " alias " & Trim(Alias_Renamed), CStr(0), 0, 0)
                If RetVal <> 0 Then
                    MsgBox(GetError(RetVal))
                    Exit Sub
                End If

                'Set default time format
                RetVal = mciSendString("set " & Trim(Alias_Renamed) & " time format milliseconds", CStr(0), 0, 0)
                m_lngTimeFormat = TimeFormats.MILISECONDS
            End Sub
            Public Sub CloseAll()
                Dim Ret As Integer
                Ret = mciSendString("close all", CStr(0), 0, 0)
            End Sub
            Public Sub Close()
                Dim Ret As Integer
                Ret = mciSendString("close " & Alias_Renamed, CStr(0), 0, 0)

            End Sub
            Private Sub FindFormat()
                Dim TempPath As String
                Dim Ext As String

                TempPath = m_strFilename
                Ext = IO.Path.GetExtension(TempPath).ToLower

                Select Case Ext
                    Case "mp3"
                        m_lngMediaType = MediaType.MP3
                    Case "mid"
                        m_lngMediaType = MediaType.SYNTHESIZER
                    Case "wav"
                        m_lngMediaType = MediaType.WAVEAUDIO
                    Case Else
                        m_lngMediaType = MediaType.MPEG_VIDEO
                End Select
            End Sub
            ReadOnly Property GetMediaType() As String
                Get
                    GetMediaType = Alias_Renamed
                End Get
            End Property

            ReadOnly Property MCIState() As Integer
                Get
                    MCIState = m_State
                End Get
            End Property
            Public ReadOnly Property TotalTime() As Integer
                Get
                    Dim strReturn As String = New String(" ", 100)
                    Dim RetVal As Integer
                    RetVal = mciSendString("set " & Alias_Renamed & " time format milliseconds", 0, 0, 0)
                    RetVal = mciSendString("status " & Alias_Renamed & " length", strReturn, 100, 0)
                    TotalTime = Val(strReturn)

                End Get
            End Property
            Public Property CurrentPosition() As Integer
                Get
                    ' áãÚÑÝÉ ãßÇä ÇáÊÔÛíá ÇáÍÇáí ÏÇÎá ÇáãáÝ
                    Dim strReturn As String = New String(" ", 100)
                    Dim RetVal As Integer
                    RetVal = mciSendString("set " & Alias_Renamed & " time format milliseconds", CStr(0), 0, 0)
                    RetVal = mciSendString("status " & Alias_Renamed & " position", strReturn, 255, 0)
                    CurrentPosition = CInt(strReturn)
                End Get
                Set(ByVal Value As Integer)
                    ' ááÅäÊÞÇá Åáì äÞØÉ ÏÇÎá ÇáãáÝ
                    Dim RetVal As Integer
                    RetVal = mciSendString("set " & Alias_Renamed & " time format milliseconds", CStr(0), 0, 0)
                    RetVal = mciSendString("seek " & Alias_Renamed & " to " & Value, CStr(0), 0, 0)
                End Set
            End Property
            Public Property Volume() As Short
                Get
                    Volume = m_Volume
                End Get
                Set(ByVal Value As Short)
                    m_Volume = Value
                    SetVolume()
                End Set
            End Property
            Public Property Channel() As Channels
                Get
                    Channel = m_Channel
                End Get
                Set(ByVal Value As Channels)
                    m_Channel = Value
                    SetChannel()
                End Set
            End Property
            Public Sub Play()
                Dim Ret As Integer
                Ret = mciSendString("play " & Alias_Renamed, CStr(0), 0, 0)
                m_State = CurrentState.PLAYING
            End Sub
            Public Sub [Stop]()
                Dim Ret As Integer
                Ret = mciSendString("stop " & Alias_Renamed, CStr(0), 0, 0)
            End Sub
            Public Sub Pause()
                Dim Ret As Integer
                Ret = mciSendString("pause " & Alias_Renamed, CStr(0), 0, 0)
                m_State = CurrentState.PAUSED
            End Sub
            Public Sub Rewind()
                Dim Ret As Integer
                Ret = mciSendString("seek " & Alias_Renamed & " to start", CStr(0), 0, 0)
            End Sub
            Public Sub SetMediaPos(ByVal lngPosInMiliseconds As Integer)
                Dim Ret As Integer
                Ret = mciSendString("set " & Alias_Renamed & " time format milliseconds", CStr(0), 0, 0)
                Ret = mciSendString("seek " & Alias_Renamed & " to " & lngPosInMiliseconds, CStr(0), 0, 0)
            End Sub
            Public Sub SetVolume(Optional ByRef strChannel As String = "")
                Dim Ret As Integer
                If strChannel = "" Then
                    Ret = mciSendString("setaudio " & Alias_Renamed & " volume to " & Str(m_Volume), CStr(0), 0, 0)
                Else
                    Ret = mciSendString("setaudio " & Alias_Renamed & " channel " & strChannel & " volume to " & Str(m_Volume), CStr(0), 0, 0)
                End If
            End Sub
            Public Sub New()
                MyBase.New()
                'Set default value
                m_Volume = 500
                m_State = CurrentState.READY
            End Sub
            Protected Overrides Sub Finalize()
                Me.Close()
                MyBase.Finalize()
            End Sub
            Public Sub SetChannel()
                Dim Ret As Integer
                Select Case m_Channel
                    Case Channels.LEFT_ONLY
                        Ret = mciSendString("set " & Alias_Renamed & " audio right off", CStr(0), 0, 0)
                        Ret = mciSendString("set " & Alias_Renamed & " audio left on", CStr(0), 0, 0)
                    Case Channels.RIGHT_ONLY
                        Ret = mciSendString("set " & Alias_Renamed & " audio right on", CStr(0), 0, 0)
                        Ret = mciSendString("set " & Alias_Renamed & " audio left off", CStr(0), 0, 0)
                    Case Channels.RIGHT_LEFT
                        Ret = mciSendString("set " & Alias_Renamed & " audio right on", CStr(0), 0, 0)
                        Ret = mciSendString("set " & Alias_Renamed & " audio left on", CStr(0), 0, 0)
                End Select

            End Sub
            Private Function GetError(ByRef ErrorCode As Integer) As String
                Dim lngError As Integer
                Dim strError As String
                Dim lngRet As Integer
                On Error Resume Next
                'Get the error
                strError = New String(" ", 129)
                lngError = mciGetErrorString(ErrorCode, strError, strError.Length - 1)
                If lngError = 1 Then
                    strError = Left(strError, InStr(strError, Chr(0)))
                Else
                    strError = "Unknown MCI Error"
                End If
                GetError = strError
            End Function
        End Class

        Public Class MicRecording


            Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Integer, ByVal hwndCallback As Integer) As Integer

            Dim lSamples, lRet, lBits, lChannels As Integer
            Dim iBlockAlign As Short
            Dim lBytesPerSec As Integer

            'For Images
            Dim J As Integer = 1
            Public Enum SoundFormats
                Mono_6kbps_8_Bit
                Mono_8kbps_8_Bit
                Mono_11kbps_8_Bit
                Mono_16kbps_8_Bit
                Mono_22kbps_8_Bit
                Mono_32kbps_8_Bit
                Mono_44kbps_8_Bit
                Mono_48kbps_8_Bit
                Stereo_24kbps_16_Bit
                Stereo_32kbps_16_Bit
                Stereo_44kbps_16_Bit
                Stereo_64kbps_16_Bit
                Stereo_88kbps_16_Bit
                Stereo_128kbps_16_Bit
                Stereo_176kbps_16_Bit
                Stereo_192kbps_16_Bit
            End Enum

            Public Enum MyState
                Idle
                Recording
                Paused
            End Enum

            Private xState As MyState
            Public ReadOnly Property State() As MyState
                Get
                    State = xState
                End Get
            End Property

            Private _SoundFormat As SoundFormats
            Public Property SoundFormat() As SoundFormats
                Get
                    SoundFormat = _SoundFormat
                End Get
                Set(ByVal Value As SoundFormats)
                    _SoundFormat = Value
                End Set
            End Property

            Private Sub GetSoundFormat()
                If _SoundFormat = SoundFormats.Mono_6kbps_8_Bit Then
                    lSamples = 6000 : lBits = 8 : lChannels = 1
                ElseIf _SoundFormat = SoundFormats.Mono_8kbps_8_Bit Then
                    lSamples = 8000 : lBits = 8 : lChannels = 1
                ElseIf _SoundFormat = SoundFormats.Mono_11kbps_8_Bit Then
                    lSamples = 11025 : lBits = 8 : lChannels = 1
                ElseIf _SoundFormat = SoundFormats.Mono_16kbps_8_Bit Then
                    lSamples = 16000 : lBits = 8 : lChannels = 1
                ElseIf _SoundFormat = SoundFormats.Mono_22kbps_8_Bit Then
                    lSamples = 22050 : lBits = 8 : lChannels = 1
                ElseIf _SoundFormat = SoundFormats.Mono_32kbps_8_Bit Then
                    lSamples = 32000 : lBits = 8 : lChannels = 1
                ElseIf _SoundFormat = SoundFormats.Mono_44kbps_8_Bit Then
                    lSamples = 44100 : lBits = 8 : lChannels = 1
                ElseIf _SoundFormat = SoundFormats.Mono_48kbps_8_Bit Then
                    lSamples = 48000 : lBits = 8 : lChannels = 1
                ElseIf _SoundFormat = SoundFormats.Stereo_24kbps_16_Bit Then
                    lSamples = 6000 : lBits = 16 : lChannels = 2
                ElseIf _SoundFormat = SoundFormats.Stereo_32kbps_16_Bit Then
                    lSamples = 8000 : lBits = 16 : lChannels = 2
                ElseIf _SoundFormat = SoundFormats.Stereo_44kbps_16_Bit Then
                    lSamples = 11025 : lBits = 16 : lChannels = 2
                ElseIf _SoundFormat = SoundFormats.Stereo_64kbps_16_Bit Then
                    lSamples = 16000 : lBits = 16 : lChannels = 2
                ElseIf _SoundFormat = SoundFormats.Stereo_88kbps_16_Bit Then
                    lSamples = 22050 : lBits = 16 : lChannels = 2
                ElseIf _SoundFormat = SoundFormats.Stereo_128kbps_16_Bit Then
                    lSamples = 32000 : lBits = 16 : lChannels = 2
                ElseIf _SoundFormat = SoundFormats.Stereo_176kbps_16_Bit Then
                    lSamples = 44100 : lBits = 16 : lChannels = 2
                ElseIf _SoundFormat = SoundFormats.Stereo_192kbps_16_Bit Then
                    lSamples = 48000 : lBits = 16 : lChannels = 2
                End If
                iBlockAlign = lChannels * lBits / 8
                lBytesPerSec = lSamples * iBlockAlign
            End Sub
            Private FName As String
            Public Property FileName() As String
                Get
                    FileName = FName
                End Get
                Set(ByVal Value As String)
                    FName = Value
                End Set
            End Property


            Public Function StartRecord() As Boolean
                Call GetSoundFormat()
                On Error GoTo ER
                If FName = "" Then GoTo ER
                Dim i As Integer
                i = mciSendString("open new type waveaudio alias capture", vbNullString, 0, 0)
                i = mciSendString("set capture samplespersec " & lSamples & " channels " & lChannels & " bitspersample " & lBits & " alignment " & iBlockAlign & " bytespersec " & lBytesPerSec, vbNullString, 0, 0)
                i = mciSendString("record capture", vbNullString, 0, 0)

                xState = MyState.Recording
                StartRecord = True
                Exit Function
ER:
                StartRecord = False
                Dim SSVoiceRecordControlExec As New ArgumentException("File Name Not Specified.")
                Throw SSVoiceRecordControlExec
            End Function

            Public Function StopRecord() As Boolean
                Dim i As Integer
                Try
                    If FName = "" Then Return False
                    i = mciSendString("save capture " & FName, vbNullString, 0, 0)
                    i = mciSendString("close capture", vbNullString, 0, 0)
                    xState = MyState.Idle
                    StopRecord = True
                Catch ex As Exception
                    i = mciSendString("close capture", vbNullString, 0, 0)
                    StopRecord = False
                    Dim SSVoiceRecordControlExec As New ArgumentException("File Name Not Specified.")
                    Throw SSVoiceRecordControlExec
                End Try
            End Function

            ''Closing Recording But Not Saved
            Public Sub CloseRecord()
                Dim i As Integer
                i = mciSendString("close capture", vbNullString, 0, 0)
                xState = MyState.Idle
            End Sub

            Private Sub Class_Initialize_Renamed()
                xState = MyState.Idle
            End Sub
            Private Sub Class_Terminate_Renamed()
                StopRecord()
            End Sub
            Protected Overrides Sub Finalize()
                Class_Terminate_Renamed()
                MyBase.Finalize()
            End Sub

            Public Function PauseRecord() As Boolean
                On Error GoTo ER
                If FName = "" Then GoTo ER
                Dim RS As String
                Dim cb, I As Integer
                RS = Space(128)
                If xState = MyState.Paused Then
                    I = mciSendString("record capture", vbNullString, 0, 0)
                    xState = MyState.Recording
                ElseIf xState = MyState.Recording Then
                    I = mciSendString("pause capture", vbNullString, 0, 0)
                    xState = MyState.Paused
                End If
                PauseRecord = True
                Exit Function
ER:
                PauseRecord = False
                Dim SSVoiceRecordControlExec As New ArgumentException("File Name Not Specified.")
                Throw SSVoiceRecordControlExec
            End Function

        End Class

    End Namespace

    Namespace Mathimatics
        Public Class CommonFunction
            Public Function Trunc(ByVal Value As Double) As Double
                Dim temp, b() As String
                temp = Value.ToString
                b = Split(temp, ".")
                temp = b(0)
                Return Val(temp)
            End Function
            Public Function Div(ByVal Value_1 As Double, ByVal Value_2 As Double) As Double
                Dim temp, b() As String
                Value_1 = Value_1 / Value_2
                temp = Value_1.ToString
                b = Split(temp, ".")
                If UBound(b) = 0 Then Return 0
                temp = b(1)
                Return Val(temp)
            End Function
            Public Function DivComp(ByVal Value_1 As Double, ByVal Value_2 As Double) As Integer
                Dim temp As Integer = 0
                If Div(Value_1, Value_2) = 0 Then Return 0
                Do Until Div(Value_1, Value_2) = 0
                    temp += 1
                    Value_1 -= 1
                Loop
                Return temp
            End Function
            Public Function Odd(ByVal Value As Int64) As Boolean
                Dim x As Double
                x = Value / 2
                If Trunc(x) = x Then
                    Return False
                Else
                    Return True
                End If
            End Function
            Public Function Even(ByVal Value As Int64) As Boolean
                Dim x As Double
                x = Value / 2
                If Trunc(x) = x Then
                    Return True
                Else
                    Return False
                End If
            End Function
        End Class

        Public Class Convertor
            Inherits EAMS.Mathimatics.CommonFunction

            Public Function DecToBase(ByVal Number As Double, Optional ByVal Base As Byte = 2) As String
                If Number = 0 Then Return "0"
                Dim Neg As Boolean = False
                If Number < 0 Then
                    Neg = True
                    Number = -1 * Number
                End If
                Dim temp As String = ""
                Do While Number >= Base
                    If Div(Number, Base) = 0 Then
                        temp = "0" & temp
                    Else
                        If DivComp(Number, Base) > 9 Then
                            temp = Chr(DivComp(Number, Base) + 55) & temp
                        Else
                            temp = DivComp(Number, Base) & temp
                        End If
                    End If
                    Number = Trunc(Number / Base)
                Loop
                If DivComp(Number, Base) > 9 Then
                    temp = Chr(DivComp(Number, Base) + 55) & temp
                Else
                    temp = DivComp(Number, Base) & CStr((temp))
                End If

                If Neg Then Return temp & -1
                Return temp
            End Function
        End Class
    End Namespace

    Namespace DataStructure
        Namespace Tree
            Public Class BinaryTree
                Private Shared T(0) As t_Child
                Private Shared T0 As String
                Private Shared NodePointerR As Integer = 0
                Private Shared NodePointerL As Integer = 0
                Private Shared NodeLevel As Integer = 0
                Private Shared NodeLevelPointer As Integer = 0
                Private Shared TN As Double
                Private Shared FeedBackflag As Boolean = False

#Region "Enum"
                Enum et_NodeType
                    Left = 1
                    Right = 2
                End Enum

#End Region

#Region "Structure"
                Structure t_Node
                    Dim Value As System.Text.StringBuilder
                    Dim NextNode As Integer
                End Structure
                Structure t_Child
                    Public LeftNode, RightNode As t_Node
                    Dim Level As Integer
                End Structure
                Structure t_Nodes
                    Private NodeDirectionflag, NodeDirectionflag2 As Boolean
                    Event Err()
                    Public Overloads Sub Add(ByVal Value As String, ByVal NodeDirection As et_NodeType)
                        Select Case NodeDirection
                            Case et_NodeType.Left
                                T(NodePointerL).LeftNode.Value = New System.Text.StringBuilder
                                T(NodePointerL).LeftNode.Value.Append(Value)
                                NodePointerL += 1
                                T(NodePointerL - 1).LeftNode.NextNode = NodePointerL
                            Case et_NodeType.Right
                                T(NodePointerR).RightNode.Value = New System.Text.StringBuilder
                                T(NodePointerR).RightNode.Value.Append(Value)
                                NodePointerR += 1
                                T(NodePointerR - 1).RightNode.NextNode = NodePointerR
                        End Select
                    End Sub
                    Private Sub FeedBack(ByVal Value As String)
                        If FeedBackflag Then
                            FeedBackflag = False
                            NodeLevelPointer -= 2
                            FillLeft(Value)
                        Else
                            FeedBackflag = True
                            NodeLevelPointer -= 1
                            FillLeft(Value)
                        End If

                    End Sub
                    Private Sub FillRight(ByVal Value As String)
                        T(NodePointerR).RightNode.Value = New System.Text.StringBuilder
                        T(NodePointerR).RightNode.Value.Append(Value)
                        NodePointerR += 1
                        T(NodePointerR - 1).RightNode.NextNode = NodePointerR
                        T(NodePointerR - 1).Level = NodeLevelPointer
                        NodeLevelPointer += 1
                    End Sub
                    Private Sub FillLeft(ByVal Value As String)
                        T(NodePointerL).LeftNode.Value = New System.Text.StringBuilder
                        T(NodePointerL).LeftNode.Value.Append(Value)
                        NodePointerL += 1
                        T(NodePointerL - 1).LeftNode.NextNode = NodePointerL
                        T(NodePointerL - 1).Level = NodeLevelPointer
                        NodeLevelPointer += 1
                    End Sub
                    Public Overloads Sub Add(ByVal Value As String)
                        If NodeLevelPointer > NodeLevel Then
                            FeedBack(Value)
                        Else
                            FillRight(Value)
                        End If
                    End Sub
                    Public Sub jj()
                        Dim x, y As Integer
                        For x = 0 To NodePointerR - 1
                            'If T(x).Level = x Then
                            MsgBox(T(y).RightNode.Value.ToString)
                            MsgBox("Level == " & T(y).Level.ToString)
                            y = T(x).LeftNode.NextNode
                            ' End If
                        Next
                        MsgBox("left")
                        y = 0
                        For x = 0 To NodePointerL - 1
                            'If T(x).Level = x Then
                            MsgBox(T(y).LeftNode.Value.ToString)
                            MsgBox("Level == " & T(y).Level.ToString)
                            y = T(x).LeftNode.NextNode
                            ' End If
                        Next
                    End Sub
                    Public Overloads Function Child(ByVal ChildIndex As Integer, ByVal NodeDirection As et_NodeType) As String
                        Dim X, Y As Integer
                        Select Case NodeDirection
                            Case et_NodeType.Left
                                If ChildIndex >= NodePointerL Then ChildIndex = NodePointerL - 1
                                For X = 0 To ChildIndex - 1
                                    Y = T(Y).LeftNode.NextNode
                                Next
                                Return T(Y).LeftNode.Value.ToString
                            Case et_NodeType.Right
                                If ChildIndex >= NodePointerR Then ChildIndex = NodePointerR - 1
                                For X = 0 To ChildIndex - 1
                                    Y = T(Y).RightNode.NextNode
                                Next
                                Return T(Y).RightNode.Value.ToString
                        End Select
                        Return ""
                    End Function
                    Public Overloads Function Child(ByVal ChildIndex As Integer) As String
                        Dim X, Y As Integer
                        Select Case NodeDirectionflag2
                            Case False
                                NodeDirectionflag2 = Not NodeDirectionflag2
                                If ChildIndex >= NodePointerL Then ChildIndex = NodePointerL - 1
                                For X = 0 To ChildIndex - 1
                                    Y = T(Y).LeftNode.NextNode
                                Next
                                Return T(Y).LeftNode.Value.ToString
                            Case True
                                NodeDirectionflag2 = Not NodeDirectionflag2
                                If ChildIndex >= NodePointerR Then ChildIndex = NodePointerR - 1
                                For X = 0 To ChildIndex - 1
                                    Y = T(Y).RightNode.NextNode
                                Next
                                Return T(Y).RightNode.Value.ToString
                        End Select
                        Return ""
                    End Function
                    Public Function Count(Optional ByVal NodeDirection As et_NodeType = et_NodeType.Right) As Integer
                        Select Case NodeDirection
                            Case et_NodeType.Left
                                Return NodePointerL
                            Case et_NodeType.Right
                                Return NodePointerR
                        End Select
                        Return 0
                    End Function
                End Structure

#End Region
#Region "Intial"
                Public Sub New(Optional ByVal NodeName As String = "", Optional ByVal MaxLevel As Integer = 10)
                    If NodeName <> "" Then
                        T0 = NodeName
                    End If
                    NodeLevel = MaxLevel
                    TN = TotalNodes(MaxLevel)
                    ReDim T(TN)
                    Dim inx As Integer
                    For inx = 0 To TN
                        T(inx) = New t_Child
                    Next
                End Sub
#End Region


#Region "Internal Methods"
                Private Function TotalNodes(ByVal NumberofLevel As Integer) As Double
                    If NumberofLevel <= 1 Then Return 2
                    Return (2 ^ NumberofLevel) + TotalNodes(NumberofLevel - 1)
                End Function
#End Region
#Region "Properties"
                Public Property Childs() As t_Nodes
                    Get

                    End Get
                    Set(ByVal Value As t_Nodes)

                    End Set
                End Property
                Public ReadOnly Property FindChild(ByVal NodeValue As String, Optional ByVal NodeDirection As et_NodeType = et_NodeType.Right, Optional ByVal IgnoreCase As Boolean = True) As Integer
                    Get
                        Dim X As Integer, M As New Mathimatics.CommonFunction, Str As New StringFunctions.Common
                        Select Case NodeDirection
                            Case et_NodeType.Left
                                If M.Odd(NodePointerL - 1) Then
                                    For X = 0 To (NodePointerL - 2) / 2
                                        If Str.Equal(T(X).RightNode.Value.ToString, NodeValue, IgnoreCase) Then
                                            Return X
                                        End If
                                    Next
                                    For X = X To (NodePointerL - 1)
                                        If Str.Equal(T(X).RightNode.Value.ToString, NodeValue, IgnoreCase) Then
                                            Return X
                                        End If
                                    Next
                                    Return -1
                                Else
                                    For X = 0 To (NodePointerL - 1) / 2
                                        If Str.Equal(T(X).RightNode.Value.ToString, NodeValue, IgnoreCase) Then
                                            Return X
                                        End If
                                    Next
                                    For X = X To (NodePointerL - 1)
                                        If Str.Equal(T(X).RightNode.Value.ToString, NodeValue, IgnoreCase) Then
                                            Return X
                                        End If
                                    Next
                                    Return -1
                                End If
                            Case et_NodeType.Right
                                If M.Odd(NodePointerR - 1) Then
                                    For X = 0 To (NodePointerR - 2) / 2
                                        If Str.Equal(T(X).RightNode.Value.ToString, NodeValue, IgnoreCase) Then
                                            Return X
                                        End If
                                    Next
                                    For X = X To (NodePointerR - 1)
                                        If Str.Equal(T(X).RightNode.Value.ToString, NodeValue, IgnoreCase) Then
                                            Return X
                                        End If
                                    Next
                                    Return -1
                                Else
                                    For X = 0 To (NodePointerR - 1) / 2
                                        If Str.Equal(T(X).RightNode.Value.ToString, NodeValue, IgnoreCase) Then
                                            Return X
                                        End If
                                    Next
                                    For X = X To (NodePointerR - 1)
                                        If Str.Equal(T(X).RightNode.Value.ToString, NodeValue, IgnoreCase) Then
                                            Return X
                                        End If
                                    Next
                                    Return -1
                                End If
                        End Select
                        Return 0
                    End Get
                End Property
#End Region

#Region "Methods"
                Public Sub DrawTree(ByRef Tree As TreeView)


                End Sub
#End Region

#Region "Dispose"
                Protected Overrides Sub Finalize()
                    MyBase.Finalize()
                End Sub
#End Region
            End Class

        End Namespace
    End Namespace

    Namespace DateTime
        Public Class CommonFunction
            Public Function GetWeekNoOfDate(ByVal dte As Date) As Integer
                Return DatePart(DateInterval.WeekOfYear, dte, FirstDayOfWeek.System, FirstWeekOfYear.System)
            End Function
        End Class
    End Namespace

    Namespace Barcode

        Public Class BarCode
#Region "Strucutr"
            Structure st_PrintBarCode
                Public Event Err(ByVal msg As String)

                Public Function Code128(ByVal TheText As String, Optional ByVal CodeLetter As String = "A") As Image
                    ' TheText متغير خاص بالنص المراد تشفيره
                    ' CodeLetter متغير خاص بالفئة المراد استخدامها

                    Dim Binaryz As String = "" 'متغير سيحمل النص بعد تحويله إلى باينرى
                    Dim I As Integer
                    Dim NumCode As Integer 'متغير سيحمل  قيمة حساب النص التكميلى
                    If CodeLetter = "A" Or CodeLetter = "a" Then
                        NumCode = 103
                        Binaryz = "00101111011"
                    End If
                    If CodeLetter = "B" Or CodeLetter = "b" Then
                        NumCode = 104
                        Binaryz = "00101101111"
                    End If
                    If CodeLetter = "C" Or CodeLetter = "c" Then
                        NumCode = 105
                        Binaryz = "00101100011"
                    End If
                    ' الكود التالى سيقوم باسناد قيمة الحرف بالباينرى حسب الجدول الخاص بالكود 128
                    For I = 1 To Len(TheText)
                        NumCode = NumCode + ((Asc(Mid(TheText, I, 1)) - 32) * I)
                        Select Case Asc(Mid(TheText, I, 1))
                            Case 32
                                Binaryz = Binaryz & "00100110011"
                            Case 33
                                Binaryz = Binaryz & "00110010011"
                            Case 34
                                Binaryz = Binaryz & "00110011001"
                            Case 35
                                Binaryz = Binaryz & "01101100111"
                            Case 36
                                Binaryz = Binaryz & "01101110011"
                            Case 37
                                Binaryz = Binaryz & "01110110011"
                            Case 38
                                Binaryz = Binaryz & "01100110111"
                            Case 39
                                Binaryz = Binaryz & "01100111011"
                            Case 40
                                Binaryz = Binaryz & "01110011011"
                            Case 41
                                Binaryz = Binaryz & "00110110111"
                            Case 42
                                Binaryz = Binaryz & "00110111011"
                            Case 43
                                Binaryz = Binaryz & "00111011011"
                            Case 44
                                Binaryz = Binaryz & "01001100011"
                            Case 45
                                Binaryz = Binaryz & "01100100011"
                            Case 46
                                Binaryz = Binaryz & "01100110001"
                            Case 47
                                Binaryz = Binaryz & "01000110011"
                            Case 48
                                Binaryz = Binaryz & "01100010011"
                            Case 49
                                Binaryz = Binaryz & "01100011001"
                            Case 50
                                Binaryz = Binaryz & "00110001101"
                            Case 51
                                Binaryz = Binaryz & "00110100011"
                            Case 52
                                Binaryz = Binaryz & "00110110001"
                            Case 53
                                Binaryz = Binaryz & "00100011011"
                            Case 54
                                Binaryz = Binaryz & "00110001011"
                            Case 55
                                Binaryz = Binaryz & "00010010001"
                            Case 56
                                Binaryz = Binaryz & "00010110011"
                            Case 57
                                Binaryz = Binaryz & "00011010011"
                            Case 58
                                Binaryz = Binaryz & "00011011001"
                            Case 59
                                Binaryz = Binaryz & "00010011011"
                            Case 60
                                Binaryz = Binaryz & "00011001011"
                            Case 61
                                Binaryz = Binaryz & "00011001101"
                            Case 62
                                Binaryz = Binaryz & "00100100111"
                            Case 63
                                Binaryz = Binaryz & "00100111001"
                            Case 64
                                Binaryz = Binaryz & "00111001001"
                            Case 65
                                Binaryz = Binaryz & "01011100111"
                            Case 66
                                Binaryz = Binaryz & "01110100111"
                            Case 67
                                Binaryz = Binaryz & "01110111001"
                            Case 68
                                Binaryz = Binaryz & "01001110111"
                            Case 69
                                Binaryz = Binaryz & "01110010111"
                            Case 70
                                Binaryz = Binaryz & "01110011101"
                            Case 71
                                Binaryz = Binaryz & "00101110111"
                            Case 72
                                Binaryz = Binaryz & "00111010111"
                            Case 73
                                Binaryz = Binaryz & "00111011101"
                            Case 74
                                Binaryz = Binaryz & "01001000111"
                            Case 75
                                Binaryz = Binaryz & "01001110001"
                            Case 76
                                Binaryz = Binaryz & "01110010001"
                            Case 77
                                Binaryz = Binaryz & "01000100111"
                            Case 78
                                Binaryz = Binaryz & "01000111001"
                            Case 79
                                Binaryz = Binaryz & "01110001001"
                            Case 80
                                Binaryz = Binaryz & "00010001001"
                            Case 81
                                Binaryz = Binaryz & "00101110001"
                            Case 82
                                Binaryz = Binaryz & "00111010001"
                            Case 83
                                Binaryz = Binaryz & "00100010111"
                            Case 84
                                Binaryz = Binaryz & "00100011101"
                            Case 85
                                Binaryz = Binaryz & "00100010001"
                            Case 86
                                Binaryz = Binaryz & "00010100111"
                            Case 87
                                Binaryz = Binaryz & "00010111001"
                            Case 88
                                Binaryz = Binaryz & "00011101001"
                            Case 89
                                Binaryz = Binaryz & "00010010111"
                            Case 90
                                Binaryz = Binaryz & "00010011101"
                            Case 91
                                Binaryz = Binaryz & "00011100101"
                            Case 92
                                Binaryz = Binaryz & "00010000101"
                            Case 93
                                Binaryz = Binaryz & "00110111101"
                            Case 94
                                Binaryz = Binaryz & "00001110101"
                            Case 95
                                Binaryz = Binaryz & "01011001111"
                            Case 96
                                Binaryz = Binaryz & "01011110011"
                            Case 97
                                Binaryz = Binaryz & "01101001111"
                            Case 98
                                Binaryz = Binaryz & "01101111001"
                            Case 99
                                Binaryz = Binaryz & "01111010011"
                            Case 100
                                Binaryz = Binaryz & "01111011001"
                            Case 101
                                Binaryz = Binaryz & "01001101111"
                            Case 102
                                Binaryz = Binaryz & "01001111011"
                            Case 103
                                Binaryz = Binaryz & "01100101111"
                            Case 104
                                Binaryz = Binaryz & "01100111101"
                            Case 105
                                Binaryz = Binaryz & "01111001011"
                            Case 106
                                Binaryz = Binaryz & "01111001101"
                            Case 107
                                Binaryz = Binaryz & "00111101101"
                            Case 108
                                Binaryz = Binaryz & "00110101111"
                            Case 109
                                Binaryz = Binaryz & "00001000101"
                            Case 110
                                Binaryz = Binaryz & "00111101011"
                            Case 111
                                Binaryz = Binaryz & "01110000101"
                            Case 112
                                Binaryz = Binaryz & "01011000011"
                            Case 113
                                Binaryz = Binaryz & "01101000011"
                            Case 114
                                Binaryz = Binaryz & "01101100001"
                            Case 115
                                Binaryz = Binaryz & "01000011011"
                            Case 116
                                Binaryz = Binaryz & "01100001011"
                            Case 117
                                Binaryz = Binaryz & "01100001101"
                            Case 118
                                Binaryz = Binaryz & "00001011011"
                            Case 119
                                Binaryz = Binaryz & "00001101011"
                            Case 120
                                Binaryz = Binaryz & "00001101101"
                            Case 121
                                Binaryz = Binaryz & "00100100001"
                            Case 122
                                Binaryz = Binaryz & "00100001001"
                            Case 123
                                Binaryz = Binaryz & "00001001001"
                            Case 124
                                Binaryz = Binaryz & "01010000111"
                            Case 125
                                Binaryz = Binaryz & "01011100001"
                            Case 126
                                Binaryz = Binaryz & "01110100001"
                            Case 127
                                Binaryz = Binaryz & "01000010111"
                            Case 128
                                Binaryz = Binaryz & "01000011101"
                            Case 129
                                Binaryz = Binaryz & "00001010111"
                            Case 130
                                Binaryz = Binaryz & "00001011101"
                            Case 131
                                Binaryz = Binaryz & "01000100001"
                            Case 132
                                Binaryz = Binaryz & "01000010001"
                            Case 133
                                Binaryz = Binaryz & "00010100001"
                            Case 134
                                Binaryz = Binaryz & "00001010001"
                            Case 135
                                Binaryz = Binaryz & "00101111011"
                            Case 136
                                Binaryz = Binaryz & "00101101111"
                            Case 137
                                Binaryz = Binaryz & "00101100011"
                            Case 138
                                Binaryz = Binaryz & "0011100010100"
                        End Select
                    Next
                    NumCode = NumCode Mod 103
                    ' الكود التالى لمعرفة الحرف المراد اضافتة لاستكمال النص
                    Select Case NumCode
                        Case 0
                            Binaryz = Binaryz & "00100110011"
                        Case 1
                            Binaryz = Binaryz & "00110010011"
                        Case 2
                            Binaryz = Binaryz & "00110011001"
                        Case 3
                            Binaryz = Binaryz & "01101100111"
                        Case 4
                            Binaryz = Binaryz & "01101110011"
                        Case 5
                            Binaryz = Binaryz & "01110110011"
                        Case 6
                            Binaryz = Binaryz & "01100110111"
                        Case 7
                            Binaryz = Binaryz & "01100111011"
                        Case 8
                            Binaryz = Binaryz & "01110011011"
                        Case 9
                            Binaryz = Binaryz & "00110110111"
                        Case 10
                            Binaryz = Binaryz & "00110111011"
                        Case 11
                            Binaryz = Binaryz & "00111011011"
                        Case 12
                            Binaryz = Binaryz & "01001100011"
                        Case 13
                            Binaryz = Binaryz & "01100100011"
                        Case 14
                            Binaryz = Binaryz & "01100110001"
                        Case 15
                            Binaryz = Binaryz & "01000110011"
                        Case 16
                            Binaryz = Binaryz & "01100010011"
                        Case 17
                            Binaryz = Binaryz & "01100011001"
                        Case 18
                            Binaryz = Binaryz & "00110001101"
                        Case 19
                            Binaryz = Binaryz & "00110100011"
                        Case 20
                            Binaryz = Binaryz & "00110110001"
                        Case 21
                            Binaryz = Binaryz & "00100011011"
                        Case 22
                            Binaryz = Binaryz & "00110001011"
                        Case 23
                            Binaryz = Binaryz & "00010010001"
                        Case 24
                            Binaryz = Binaryz & "00010110011"
                        Case 25
                            Binaryz = Binaryz & "00011010011"
                        Case 26
                            Binaryz = Binaryz & "00011011001"
                        Case 27
                            Binaryz = Binaryz & "00010011011"
                        Case 28
                            Binaryz = Binaryz & "00011001011"
                        Case 29
                            Binaryz = Binaryz & "00011001101"
                        Case 30
                            Binaryz = Binaryz & "00100100111"
                        Case 31
                            Binaryz = Binaryz & "00100111001"
                        Case 32
                            Binaryz = Binaryz & "00111001001"
                        Case 33
                            Binaryz = Binaryz & "01011100111"
                        Case 34
                            Binaryz = Binaryz & "01110100111"
                        Case 35
                            Binaryz = Binaryz & "01110111001"
                        Case 36
                            Binaryz = Binaryz & "01001110111"
                        Case 37
                            Binaryz = Binaryz & "01110010111"
                        Case 38
                            Binaryz = Binaryz & "01110011101"
                        Case 39
                            Binaryz = Binaryz & "00101110111"
                        Case 40
                            Binaryz = Binaryz & "00111010111"
                        Case 41
                            Binaryz = Binaryz & "00111011101"
                        Case 42
                            Binaryz = Binaryz & "01001000111"
                        Case 43
                            Binaryz = Binaryz & "01001110001"
                        Case 44
                            Binaryz = Binaryz & "01110010001"
                        Case 45
                            Binaryz = Binaryz & "01000100111"
                        Case 46
                            Binaryz = Binaryz & "01000111001"
                        Case 47
                            Binaryz = Binaryz & "01110001001"
                        Case 48
                            Binaryz = Binaryz & "00010001001"
                        Case 49
                            Binaryz = Binaryz & "00101110001"
                        Case 50
                            Binaryz = Binaryz & "00111010001"
                        Case 51
                            Binaryz = Binaryz & "00100010111"
                        Case 52
                            Binaryz = Binaryz & "00100011101"
                        Case 53
                            Binaryz = Binaryz & "00100010001"
                        Case 54
                            Binaryz = Binaryz & "00010100111"
                        Case 55
                            Binaryz = Binaryz & "00010111001"
                        Case 56
                            Binaryz = Binaryz & "00011101001"
                        Case 57
                            Binaryz = Binaryz & "00010010111"
                        Case 58
                            Binaryz = Binaryz & "00010011101"
                        Case 59
                            Binaryz = Binaryz & "00011100101"
                        Case 60
                            Binaryz = Binaryz & "00010000101"
                        Case 61
                            Binaryz = Binaryz & "00110111101"
                        Case 62
                            Binaryz = Binaryz & "00001110101"
                        Case 63
                            Binaryz = Binaryz & "01011001111"
                        Case 64
                            Binaryz = Binaryz & "01011110011"
                        Case 65
                            Binaryz = Binaryz & "01101001111"
                        Case 66
                            Binaryz = Binaryz & "01101111001"
                        Case 67
                            Binaryz = Binaryz & "01111010011"
                        Case 68
                            Binaryz = Binaryz & "01111011001"
                        Case 69
                            Binaryz = Binaryz & "01001101111"
                        Case 70
                            Binaryz = Binaryz & "01001111011"
                        Case 71
                            Binaryz = Binaryz & "01100101111"
                        Case 72
                            Binaryz = Binaryz & "01100111101"
                        Case 73
                            Binaryz = Binaryz & "01111001011"
                        Case 74
                            Binaryz = Binaryz & "01111001101"
                        Case 75
                            Binaryz = Binaryz & "00111101101"
                        Case 76
                            Binaryz = Binaryz & "00110101111"
                        Case 77
                            Binaryz = Binaryz & "00001000101"
                        Case 78
                            Binaryz = Binaryz & "00111101011"
                        Case 79
                            Binaryz = Binaryz & "01110000101"
                        Case 80
                            Binaryz = Binaryz & "01011000011"
                        Case 81
                            Binaryz = Binaryz & "01101000011"
                        Case 82
                            Binaryz = Binaryz & "01101100001"
                        Case 83
                            Binaryz = Binaryz & "01000011011"
                        Case 84
                            Binaryz = Binaryz & "01100001011"
                        Case 85
                            Binaryz = Binaryz & "01100001101"
                        Case 86
                            Binaryz = Binaryz & "00001011011"
                        Case 87
                            Binaryz = Binaryz & "00001101011"
                        Case 88
                            Binaryz = Binaryz & "00001101101"
                        Case 89
                            Binaryz = Binaryz & "00100100001"
                        Case 90
                            Binaryz = Binaryz & "00100001001"
                        Case 91
                            Binaryz = Binaryz & "00001001001"
                        Case 92
                            Binaryz = Binaryz & "01010000111"
                        Case 93
                            Binaryz = Binaryz & "01011100001"
                        Case 94
                            Binaryz = Binaryz & "01110100001"
                        Case 95
                            Binaryz = Binaryz & "01000010111"
                        Case 96
                            Binaryz = Binaryz & "01000011101"
                        Case 97
                            Binaryz = Binaryz & "00001010111"
                        Case 98
                            Binaryz = Binaryz & "00001011101"
                        Case 99
                            Binaryz = Binaryz & "01000100001"
                        Case 100
                            Binaryz = Binaryz & "01000010001"
                        Case 101
                            Binaryz = Binaryz & "00010100001"
                        Case 102
                            Binaryz = Binaryz & "00001010001"
                    End Select
                    Binaryz = Binaryz & "0011100010100" ' انهاء الكود باضافة الباينرى الخاص بايقاف جميع الاكواد

                    ' انشاء صورة عرضها عدد حروف الباينرى المستخدم
                    Dim bmp As Bitmap = New Bitmap(Len(Binaryz), 60, System.Drawing.Imaging.PixelFormat.Format24bppRgb)
                    Dim z As String ' متغير لمعرفة لون الخط 
                    Dim GraphZ As Graphics = Graphics.FromImage(bmp)
                    Dim RectZ As Rectangle = New Rectangle(0, 0, bmp.Width, bmp.Height) ' مستطيل بحجم الصورة لاعطاء الخلفية باللون الابيض
                    ' فرشاه لدهان المستطيل السابق باللون الابيض
                    Dim myBrush As Brush = New Drawing.Drawing2D.LinearGradientBrush(RectZ, Color.White, Color.White, Drawing.Drawing2D.LinearGradientMode.ForwardDiagonal)
                    ' دهان المستطيل السابق باللون الابيض
                    GraphZ.FillRectangle(myBrush, RectZ)
                    '  رسم خطوط الباركود
                    Dim PenZ As Pen
                    Dim point1 As Point ' نقطة بداية الخط
                    Dim point2 As Point ' نقطة نهاية الخط
                    For I = 1 To Len(Binaryz)
                        z = Mid(Binaryz, I, 1)
                        If z = "0" Then
                            PenZ = New Pen(Color.Black, 1)
                            point1 = New Point(I, 0)
                            point2 = New Point(I, 40)
                            GraphZ.DrawLine(PenZ, point1, point2)
                        Else
                            PenZ = New Pen(Color.White, 1)
                            point1 = New Point(I, 0)
                            point2 = New Point(I, 40)
                            GraphZ.DrawLine(PenZ, point1, point2)
                        End If
                    Next
                    ' رسم النص المراد ترميزه اسفل الكود
                    ' GraphZ.DrawString(TheText, New Font("times new roman", 12, FontStyle.Bold), New SolidBrush(Color.DarkBlue), 20, 40)
                    ' ارجاع الصورة النهائية للدالة
                    Code128 = bmp
                End Function
            End Structure
#End Region
#Region "Property"
            Public Property Barcode() As st_PrintBarCode
                Get

                End Get
                Set(ByVal value As st_PrintBarCode)

                End Set
            End Property
#End Region

        End Class

        Public Class PrintLabel
            Private Shared _UHeader As String = ""
            Private Shared _Ubarcode As Image = Nothing
            Private Shared _UText1 As String = ""
            Private Shared _UText2 As String = ""

            Private Shared _DHeader As String = ""
            Private Shared _Dbarcode As Image = Nothing
            Private Shared _DText1 As String = ""
            Private Shared _DText2 As String = ""

#Region "Structures"
            Public Structure s_UpperSide
                Public Property Header As String
                    Get
                        Return _UHeader
                    End Get
                    Set(ByVal value As String)
                        _UHeader = value
                    End Set
                End Property

                Public Property Barcode As Image
                    Get
                        Return _Ubarcode
                    End Get
                    Set(ByVal value As Image)
                        _Ubarcode = value
                    End Set
                End Property

                Public Property Text1 As String
                    Get
                        Return _UText1
                    End Get
                    Set(ByVal value As String)
                        _UText1 = value
                    End Set
                End Property

                Public Property Text2 As String
                    Get
                        Return _UText2
                    End Get
                    Set(ByVal value As String)
                        _UText2 = value
                    End Set
                End Property

            End Structure

            Public Structure s_DownSide
                Public Property Header As String
                    Get
                        Return _DHeader
                    End Get
                    Set(ByVal value As String)
                        _DHeader = value
                    End Set
                End Property

                Public Property Barcode As Image
                    Get
                        Return _Dbarcode
                    End Get
                    Set(ByVal value As Image)
                        _Dbarcode = value
                    End Set
                End Property

                Public Property Text1 As String
                    Get
                        Return _DText1
                    End Get
                    Set(ByVal value As String)
                        _DText1 = value
                    End Set
                End Property

                Public Property Text2 As String
                    Get
                        Return _DText2
                    End Get
                    Set(ByVal value As String)
                        _DText2 = value
                    End Set
                End Property

            End Structure
#End Region

#Region "Properties"
            Public Property LabelUpperSide As s_UpperSide
                Get

                End Get
                Set(ByVal value As s_UpperSide)

                End Set
            End Property

            Public Property LabelDownSide As s_DownSide
                Get

                End Get
                Set(ByVal value As s_DownSide)

                End Set
            End Property
#End Region
            Public Enum e_printerType
                BIXOLON = 1
            End Enum
            Public Enum e_LabelDimentions
                _35X25_mm = 1
            End Enum

            Public Sub Print(ByRef e As System.Drawing.Printing.PrintPageEventArgs, ByVal PrinterType As e_printerType, ByVal LabelDimentions As e_LabelDimentions)
                Select Case PrinterType
                    Case e_printerType.BIXOLON

                        Select Case LabelDimentions
                            Case e_LabelDimentions._35X25_mm
                                'Dimentions of the printer must be width=3.5 cm x 2.5 cm
                                'Options No of copies=1
                                'Advanced setup =Gap   , Every 1 label

                                Dim prFont1 As New Font("Verdana", 4, FontStyle.Regular, GraphicsUnit.Point)
                                Dim prFont2 As New Font("Verdana", 6, FontStyle.Regular, GraphicsUnit.Point)
                                Dim rec1 As Rectangle = New Rectangle(4, 2, 28, 2)
                                Dim rec2 As Rectangle = New Rectangle(16, 8, 18, 3)
                                Dim rec3 As Rectangle = New Rectangle(2, 8, 14, 3)

                                Dim prFont3 As New Font("Verdana", 4, FontStyle.Regular, GraphicsUnit.Point)
                                Dim prFont4 As New Font("Verdana", 6, FontStyle.Regular, GraphicsUnit.Point)
                                Dim rec4 As Rectangle = New Rectangle(4, 14, 28, 2)
                                Dim rec5 As Rectangle = New Rectangle(16, 20, 18, 3)
                                Dim rec6 As Rectangle = New Rectangle(2, 20, 14, 3)

                                e.Graphics.PageUnit = GraphicsUnit.Millimeter

                                e.Graphics.DrawString(_UHeader, prFont1, Brushes.Black, rec1) 'Company Name
                                e.Graphics.DrawImage(_Ubarcode, 0, 4, 34, 4)  'BARCODE
                                e.Graphics.DrawString(_UText1, prFont2, Brushes.Black, rec3) 'Barcode text
                                e.Graphics.DrawString(_UText2, prFont2, Brushes.Black, rec2) 'Other text

                                If Not _Dbarcode Is Nothing Then
                                    'Other side
                                    e.Graphics.DrawString(_DHeader, prFont3, Brushes.Black, rec4) 'Company Name
                                    e.Graphics.DrawImage(_Dbarcode, 0, 16, 34, 4)  'BARCODE
                                    e.Graphics.DrawString(_DText1, prFont4, Brushes.Black, rec6) 'Barcode text
                                    e.Graphics.DrawString(_DText2, prFont4, Brushes.Black, rec5) 'Other text
                                End If

                        End Select

                End Select
            End Sub
        End Class

        Public Class Steaker
            Public Event CardsCounts(ByVal count As Integer)
            Public Event PrintCardsProgress(ByVal inx As Integer)
            Public Event PrintComplete()
            Private _steakerperpage As Integer = 10
            Private _steakerperline As Integer = 2
            Private Shared _StWidth As Integer = 85
            Private Shared _StHieght As Integer = 55


#Region "Structure"
            Public Structure _SteakerDimention
                Event err()
                Public Property Width() As Integer
                    Get
                        Return _StWidth
                    End Get
                    Set(ByVal value As Integer)
                        _StWidth = value
                    End Set
                End Property
                Public Property Hieght() As Integer
                    Get
                        Return _StHieght
                    End Get
                    Set(ByVal value As Integer)
                        _StHieght = value
                    End Set
                End Property
            End Structure
#End Region

#Region "Properties"
            Public Property SteakerDimention() As _SteakerDimention
                Get

                End Get
                Set(ByVal value As _SteakerDimention)

                End Set
            End Property
            Public Property SteakerPerPage() As Integer
                Get
                    Return _steakerperpage
                End Get
                Set(ByVal value As Integer)
                    _steakerperpage = value
                End Set
            End Property
            Public Property SteakerPerLine() As Integer
                Get
                    Return _steakerperline
                End Get
                Set(ByVal value As Integer)
                    _steakerperline = value
                End Set
            End Property
#End Region

#Region "Methods"
            Public Sub PrintSteaker(ByVal CardsFolder As String, ByVal DistFolder As String)
                Dim bmap As Bitmap = New Bitmap(800, 1200), x As Integer = 0
                Dim iniX As Integer = 5, iniY = 5, iniWidth = _StWidth, iniHieght = _StHieght
                Dim CardPic_rec As New Rectangle(iniX, iniY, iniWidth, iniHieght)  'dimention of the card
                Dim PageName As String = "", PageCount As Integer = 1
                Dim g As Graphics = Graphics.FromImage(bmap)
                Dim IfNewPage As Boolean = True
                Dim PicCol As New Collection

                For x = 0 To My.Computer.FileSystem.GetFiles(CardsFolder, FileIO.SearchOption.SearchAllSubDirectories).Count - 1
                    If InStr((My.Computer.FileSystem.GetFiles(CardsFolder, FileIO.SearchOption.SearchAllSubDirectories).Item(x).ToString), "thumb", CompareMethod.Text) = 0 Then
                        PicCol.Add(My.Computer.FileSystem.GetFiles(CardsFolder, FileIO.SearchOption.SearchAllSubDirectories).Item(x))
                    End If
                    Application.DoEvents()
                Next


                g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                g.PageUnit = GraphicsUnit.Millimeter
                g.FillRectangle(Brushes.White, 0, 0, 800, 1200)

                If My.Computer.FileSystem.GetFiles(CardsFolder, FileIO.SearchOption.SearchAllSubDirectories).Count = 0 Then
                    Exit Sub
                End If

                CardPic_rec.Width = iniWidth
                CardPic_rec.Height = iniHieght
                RaiseEvent CardsCounts(PicCol.Count)
                For x = 0 To PicCol.Count - 1
                    RaiseEvent PrintCardsProgress(x)
                    CardPic_rec.X = iniX
                    CardPic_rec.Y = iniY

                    PageName = "Page " & PageCount.ToString
                    '------------------------------
                    If ((x Mod _steakerperline) = 0) Then
                        iniX = 5
                        iniY += iniHieght + 5
                    Else
                        iniX += iniWidth + 5
                    End If
                    If IfNewPage Then
                        iniX = 5
                        iniY = 5
                        IfNewPage = False
                    End If
                    CardPic_rec.X = iniX
                    CardPic_rec.Y = iniY
                    '-----------------------------
                    If ((x Mod _steakerperpage) = 0) And (x <> 0) Then   'Here the check of the count of cards per page  (due to card dimention)
                        bmap.Save(DistFolder & "\" & PageName & ".jpg", Imaging.ImageFormat.Jpeg)
                        PageCount += 1
                        IfNewPage = True
                        bmap = New Bitmap(800, 1200)
                        g = Graphics.FromImage(bmap)
                        g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                        g.PageUnit = GraphicsUnit.Millimeter
                        g.FillRectangle(Brushes.White, 0, 0, 800, 1200)
                        iniX = 5
                        iniY = 5
                        CardPic_rec.X = iniX
                        CardPic_rec.Y = iniY
                        IfNewPage = False
                    End If
                    'If InStr((My.Computer.FileSystem.GetFiles(CardsFolder, FileIO.SearchOption.SearchAllSubDirectories).Item(x).ToString), "thumb", CompareMethod.Text) = 0 Then
                    g.DrawImage(Image.FromFile(PicCol.Item(x + 1)), CardPic_rec)
                    Application.DoEvents()
                    ' End If


                    Application.DoEvents()
                Next
                bmap.Save(DistFolder & "\" & PageName & ".jpg", Imaging.ImageFormat.Jpeg)
                RaiseEvent PrintComplete()
            End Sub
#End Region





        End Class
    End Namespace

    Namespace Modem

        Public Class PhoneModem
            Private _number As String = ""
            Private _date As String = ""
            Private _time As String = ""
            Private _temp(1) As String
            Private str As New EAMS.StringFunctions.StringsFunction
            Private WithEvents com As New System.IO.Ports.SerialPort
            Public Event Incoming(ByVal num As String)
            Public Event er(ByVal m As String)
            Public Event ModemSupportCallerID()
            Public Event Ringing()
            Private Buffer As New System.Windows.Forms.TextBox
            Public canceled As Boolean = False
            Public ComPort As Byte = 3
            Private demoThread As Thread = Nothing
            Delegate Sub SetTextCallback(ByVal [text] As String)

#Region "Property"
            Public ReadOnly Property Number() As String
                Get
                    Return _number
                End Get
            End Property
            Public ReadOnly Property IncomingDate() As String
                Get
                    Return _date
                End Get
            End Property
            Public ReadOnly Property IncomingTime() As String
                Get
                    Return _time
                End Get
            End Property

#End Region

#Region "Modem Handler"
#Region "Serial Port"
            Private Sub setTextSafeBtn_Click(ByVal sender As Object, ByVal e As EventArgs) Handles com.DataReceived
                Me.demoThread = New Thread(New ThreadStart(AddressOf Me.ThreadProcSafe))
                Me.demoThread.Start()
            End Sub
            Private Sub ThreadProcSafe()
                Me.SetText(com.ReadExisting)
            End Sub
            Private Sub SetText(ByVal [text] As String)
                If Buffer.InvokeRequired Then
                    Dim d As New SetTextCallback(AddressOf SetText)
                    Buffer.Invoke(d, New Object() {[text]})
                Else
                    Buffer.Text = [text]
                End If
                com_OnComm()
            End Sub
#End Region
            Private Sub com_OnComm()
                If InStr(Buffer.Text, "ok", CompareMethod.Text) Then
                    RaiseEvent ModemSupportCallerID()
                    Exit Sub
                End If
                If InStr(Buffer.Text, "DATE", CompareMethod.Text) > 0 Then
                    canceled = False
                    ImcomingCall(Buffer.Text)
                    Buffer.Text = ""
                    Exit Sub
                End If
                If InStr(Buffer.Text, "RING", CompareMethod.Text) Then
                    If Not canceled Then
                        RaiseEvent Ringing()
                        Buffer.Text = ""
                    End If
                End If
            End Sub
#End Region
#Region "Method"
            Private Sub ImcomingCall2(ByVal value As String) 'for local central
                _date = Trim(str.SubString(str.PartString(value, "=", 1), "TIME"))
                _time = Trim(str.SubString(str.PartString(value, "=", 2), "NMBR"))
                _number = Trim(str.PartString(value, "=", 3))
                _date = _date.Chars(0) & _date.Chars(1) & "/" & _date.Chars(2) & _date.Chars(3)
                _time = _time.Chars(0) & _time.Chars(1) & ":" & _time.Chars(2) & _time.Chars(3)
                RaiseEvent Incoming(_number)
            End Sub
            Private Sub ImcomingCall(ByVal value As String)  'for private central
                If value = "" Then Exit Sub
                _date = str.PartString(str.PartString(value, vbCrLf, 1), "=", 1)
                _time = str.PartString(str.PartString(value, vbCrLf, 3), "=", 1)
                _number = str.PartString(str.PartString(value, vbCrLf, 5), "=", 1)
                _date = _date.Chars(0) & _date.Chars(1) & "/" & _date.Chars(2) & _date.Chars(3)
                _time = _time.Chars(0) & _time.Chars(1) & ":" & _time.Chars(2) & _time.Chars(3)
                RaiseEvent Incoming(_number)
            End Sub
#End Region
#Region "Public Methods"
            Public Sub ModemIntilization()
                Try
                    If com.IsOpen = True Then com.Close()
                    com.PortName = "COM" & ComPort
                    com.BaudRate = 115200
                    com.DataBits = 8
                    com.Parity = IO.Ports.Parity.None
                    com.StopBits = IO.Ports.StopBits.One
                    com.DtrEnable = True
                    com.RtsEnable = True
                    com.ReadBufferSize = 8000
                    com.Open()
                    com.Write("AT#CID=1" & vbCrLf)
                    com.Write("AT+VCID=1" & vbCrLf)
                    com.Write("AT%CCID=1" & vbCrLf)
                    com.Write("AT#CC1" & vbCrLf)
                    com.Write("AT*ID1" & vbCrLf)
                    '

                Catch ex As Exception
                    RaiseEvent er("There's an error while connecting to the modem")
                End Try
            End Sub
            Public Sub SelectVoiceMode()
                If com.IsOpen Then
                    com.Write("AT+FCLASS=8" & vbCrLf)
                End If
            End Sub
            Public Sub Dial(ByVal Number As String)
                If com.IsOpen Then
                    com.Write("ATDT" & Number & vbCrLf)
                End If
            End Sub
            Public Sub AnswerInVoiceMode()
                If com.IsOpen Then
                    com.Write("ATA" & vbCrLf)
                End If
            End Sub
            Public Sub ExitVoiceMode()
                If com.IsOpen Then
                    com.Write("ATZ" & vbCrLf)
                End If
            End Sub
            Public Sub HangUp()
                If com.IsOpen Then
                    com.Write("ATH" & vbCrLf)
                End If
            End Sub
            Public Sub SendWaveFileOld(ByVal fleName As String)
                If com.IsOpen Then
                    '
                    com.Write("ATQ0V1E0&D2X4S0=0" & vbCrLf)
                    com.Write("ATM1L0S7=50" & vbCrLf)
                    com.Write("ATS0=0" & vbCrLf)
                    com.Write("AT+FCLASS=8;+IFC=2,2" & vbCrLf)
                    Thread.Sleep(100)
                    com.Write("AT+VRN=0;+VIT=0;+VEM=1" & vbCrLf)
                    Thread.Sleep(100)
                    com.Write("AT+VSM=132,8000" & vbCrLf)
                    com.Write("AT+VGT=130" & vbCrLf)
                    Thread.Sleep(100)
                    com.Write("AT+FLO=2" & vbCrLf)
                    com.Write("AT+VSD=128,0" & vbCrLf)
                    Thread.Sleep(100)
                    com.Write("AT+VLS=1" & vbCrLf)
                    Thread.Sleep(100)
                    com.Write("AT+VTX" & vbCrLf)

                    Dim sw As Boolean = False
                    Dim buffer(2000) As Byte
                    Dim strm As New FileStream(fleName, FileMode.Open)
                    Dim ms As New MemoryStream
                    Dim count As Integer = ms.Read(buffer, 44, buffer.Length - 44)
                    Dim bdr As New BinaryReader(strm)
                    While Not sw
                        Dim bt(512) As Byte
                        bt = bdr.ReadBytes(512)
                        If bt.Length = 0 Then
                            sw = True
                            Exit While
                        Else
                            com.Write(bt, 0, bt.Length)
                            Thread.Sleep(10)
                        End If
                    End While
                    strm.Close()
                    strm.Dispose()
                    HangUp()
                End If
            End Sub


#End Region


        End Class
    End Namespace

    Namespace Coding
        Public Class EncodeString
            Public Function Encode(ByVal word As String) As String
                Dim inx As Integer = 0
                Dim chrASC As Integer = 0
                Dim fCode As String = ""
                For inx = 0 To word.Length - 1
                    chrASC = Asc(word.Chars(inx)) + inx + 13
                    If chrASC > 255 Then chrASC = 255 - chrASC
                    fCode &= Chr(chrASC)
                Next
                Return fCode
            End Function
        End Class

        Public Class FileEncoding
            'Dim strFileToEncrypt As String
            ' Dim strFileToDecrypt As String
            Dim fsInput As System.IO.FileStream
            Dim fsOutput As System.IO.FileStream

            Public Event MaximumProgressValue(ByVal v As Integer)
            Public Event ProgressValue(ByVal v As Integer)
            Public Event ProgressComplete()
            Public Event Info(ByVal msg As String)

            Public Enum CryptoAction
                'Define the enumeration for CryptoAction.
                ActionEncrypt = 1
                ActionDecrypt = 2
            End Enum

            Private Function CreateKey(ByVal strPassword As String) As Byte()
                'Convert strPassword to an array and store in chrData.
                Dim chrData() As Char = strPassword.ToCharArray
                'Use intLength to get strPassword size.
                Dim intLength As Integer = chrData.GetUpperBound(0)
                'Declare bytDataToHash and make it the same size as chrData.
                Dim bytDataToHash(intLength) As Byte

                'Use For Next to convert and store chrData into bytDataToHash.
                For i As Integer = 0 To chrData.GetUpperBound(0)
                    bytDataToHash(i) = CByte(Asc(chrData(i)))
                Next

                'Declare what hash to use.
                Dim SHA512 As New System.Security.Cryptography.SHA512Managed
                'Declare bytResult, Hash bytDataToHash and store it in bytResult.
                Dim bytResult As Byte() = SHA512.ComputeHash(bytDataToHash)
                'Declare bytKey(31).  It will hold 256 bits.
                Dim bytKey(31) As Byte

                'Use For Next to put a specific size (256 bits) of 
                'bytResult into bytKey. The 0 To 31 will put the first 256 bits
                'of 512 bits into bytKey.
                For i As Integer = 0 To 31
                    bytKey(i) = bytResult(i)
                Next

                Return bytKey 'Return the key.
            End Function

            Private Function CreateIV(ByVal strPassword As String) As Byte()
                'Convert strPassword to an array and store in chrData.
                Dim chrData() As Char = strPassword.ToCharArray
                'Use intLength to get strPassword size.
                Dim intLength As Integer = chrData.GetUpperBound(0)
                'Declare bytDataToHash and make it the same size as chrData.
                Dim bytDataToHash(intLength) As Byte

                'Use For Next to convert and store chrData into bytDataToHash.
                For i As Integer = 0 To chrData.GetUpperBound(0)
                    bytDataToHash(i) = CByte(Asc(chrData(i)))
                Next

                'Declare what hash to use.
                Dim SHA512 As New System.Security.Cryptography.SHA512Managed
                'Declare bytResult, Hash bytDataToHash and store it in bytResult.
                Dim bytResult As Byte() = SHA512.ComputeHash(bytDataToHash)
                'Declare bytIV(15).  It will hold 128 bits.
                Dim bytIV(15) As Byte

                'Use For Next to put a specific size (128 bits) of 
                'bytResult into bytIV. The 0 To 30 for bytKey used the first 256 bits.
                'of the hashed password. The 32 To 47 will put the next 128 bits into bytIV.
                For i As Integer = 32 To 47
                    bytIV(i - 32) = bytResult(i)
                Next

                Return bytIV 'return the IV
            End Function

            Private Function MakeDesFile(ByVal sourceFile As String, ByVal Direction As CryptoAction) As String
                Dim strOutputDecrypt As String = ""
                Dim strOutputEncrypt As String = ""
                Dim iPosition As Integer = 0
                Dim i As Integer = 0

                Select Case Direction
                    Case CryptoAction.ActionEncrypt
                        'Get the position of the last "\" in the OpenFileDialog.FileName path.
                        '-1 is when the character your searching for is not there.
                        'IndexOf searches from left to right.
                        While sourceFile.IndexOf("\"c, i) <> -1
                            iPosition = sourceFile.IndexOf("\"c, i)
                            i = iPosition + 1
                        End While

                        'Assign strOutputFile to the position after the last "\" in the path.
                        'This position is the beginning of the file name.
                        strOutputEncrypt = sourceFile.Substring(iPosition + 1)
                        'Assign S the entire path, ending at the last "\".
                        Dim S As String = sourceFile.Substring(0, iPosition + 1)
                        'Replace the "." in the file extension with "_".
                        strOutputEncrypt = strOutputEncrypt.Replace("."c, "_"c)
                        'The final file name.  XXXXX.encrypt
                        Return S + strOutputEncrypt + ".encrypt"

                    Case CryptoAction.ActionDecrypt
                        'Get the position of the last "\" in the OpenFileDialog.FileName path.
                        '-1 is when the character your searching for is not there.
                        'IndexOf searches from left to right.

                        While sourceFile.IndexOf("\"c, i) <> -1
                            iPosition = sourceFile.IndexOf("\"c, i)
                            i = iPosition + 1
                        End While

                        'strOutputFile = the file path minus the last 8 characters (.encrypt)
                        strOutputDecrypt = sourceFile.Substring(0, sourceFile.Length - 8)
                        'Assign S the entire path, ending at the last "\".
                        Dim S As String = sourceFile.Substring(0, iPosition + 1)
                        'Assign strOutputFile to the position after the last "\" in the path.
                        strOutputDecrypt = strOutputDecrypt.Substring((iPosition + 1))
                        'Replace "_" with "."
                        Return S + strOutputDecrypt.Replace("_"c, "."c)
                End Select
                Return ""
            End Function


            Public Sub EncryptOrDecryptFile(ByVal strInputFile As String, ByVal Password As String, ByVal Direction As CryptoAction)
                Dim bytKey() As Byte
                Dim bytIV() As Byte
                Dim strOutputFile As String = MakeDesFile(strInputFile, Direction)

                Try 'In case of errors.
                    bytKey = CreateKey(Password)
                    'Send the password to the CreateIV function.
                    bytIV = CreateIV(Password)

                    'Setup file streams to handle input and output.
                    fsInput = New System.IO.FileStream(strInputFile, FileMode.Open, _
                                                       FileAccess.Read)
                    fsOutput = New System.IO.FileStream(strOutputFile, FileMode.OpenOrCreate, _
                                                        FileAccess.Write)
                    fsOutput.SetLength(0) 'make sure fsOutput is empty

                    'Declare variables for encrypt/decrypt process.
                    Dim bytBuffer(4096) As Byte 'holds a block of bytes for processing
                    Dim lngBytesProcessed As Long = 0 'running count of bytes processed
                    Dim lngFileLength As Long = fsInput.Length 'the input file's length
                    Dim intBytesInCurrentBlock As Integer 'current bytes being processed
                    Dim csCryptoStream As CryptoStream = Nothing
                    'Declare your CryptoServiceProvider.
                    Dim cspRijndael As New System.Security.Cryptography.RijndaelManaged
                    'Setup Progress Bar

                    RaiseEvent ProgressValue(0)
                    RaiseEvent MaximumProgressValue(100)

                    'Determine if ecryption or decryption and setup CryptoStream.
                    Select Case Direction
                        Case CryptoAction.ActionEncrypt
                            csCryptoStream = New CryptoStream(fsOutput, _
                            cspRijndael.CreateEncryptor(bytKey, bytIV), _
                            CryptoStreamMode.Write)

                        Case CryptoAction.ActionDecrypt
                            csCryptoStream = New CryptoStream(fsOutput, _
                            cspRijndael.CreateDecryptor(bytKey, bytIV), _
                            CryptoStreamMode.Write)
                    End Select

                    'Use While to loop until all of the file is processed.
                    While lngBytesProcessed < lngFileLength
                        'Read file with the input filestream.
                        intBytesInCurrentBlock = fsInput.Read(bytBuffer, 0, 4096)
                        'Write output file with the cryptostream.
                        csCryptoStream.Write(bytBuffer, 0, intBytesInCurrentBlock)
                        'Update lngBytesProcessed
                        lngBytesProcessed = lngBytesProcessed + CLng(intBytesInCurrentBlock)
                        'Update Progress Bar
                        RaiseEvent ProgressValue(CInt((lngBytesProcessed / lngFileLength) * 100))

                    End While

                    'Close FileStreams and CryptoStream.
                    csCryptoStream.Close()
                    fsInput.Close()
                    fsOutput.Close()

                    'If encrypting then delete the original unencrypted file.
                    If Direction = CryptoAction.ActionEncrypt Then
                        Dim fileOriginal As New FileInfo(strInputFile)
                        fileOriginal.Delete()
                    End If

                    'If decrypting then delete the encrypted file.
                    If Direction = CryptoAction.ActionDecrypt Then
                        Dim fileEncrypted As New FileInfo(strInputFile)
                        fileEncrypted.Delete()
                    End If

                    'Update the user when the file is done.
                    Dim Wrap As String = Chr(13) + Chr(10)
                    If Direction = CryptoAction.ActionEncrypt Then
                        RaiseEvent Info("Encryption Complete" + Wrap + Wrap + _
                                "Total bytes processed = " + _
                                lngBytesProcessed.ToString)

                        'Update the progress bar and textboxes.
                        RaiseEvent ProgressValue(0)
                        RaiseEvent ProgressComplete()
                    Else
                        'Update the user when the file is done.
                        RaiseEvent Info("Decryption Complete" + Wrap + Wrap + _
                               "Total bytes processed = " + _
                                lngBytesProcessed.ToString)

                        'Update the progress bar and textboxes.
                        RaiseEvent ProgressValue(0)
                        RaiseEvent ProgressComplete()
                    End If


                    'Catch file not found error.
                Catch When Err.Number = 53 'if file not found
                    RaiseEvent Info("Please check to make sure the path and filename" + _
                            "are correct and if the file exists.")

                    'Catch all other errors. And delete partial files.
                Catch
                    fsInput.Close()
                    fsOutput.Close()

                    If Direction = CryptoAction.ActionDecrypt Then
                        Dim fileDelete As New FileInfo(strOutputFile)
                        fileDelete.Delete()
                        RaiseEvent ProgressValue(0)


                        RaiseEvent Info("Please check to make sure that you entered the correct" + _
                                "password.")
                    Else
                        Dim fileDelete As New FileInfo(strOutputFile)
                        fileDelete.Delete()

                        RaiseEvent ProgressValue(0)


                        RaiseEvent Info("This file cannot be encrypted.")

                    End If

                End Try
            End Sub
        End Class
    End Namespace

    Namespace Mailling
        'For Gmail receive and send
        ' Dim m As New Mailling.POP3
        'm.Connect("pop.gmail.com", "payloadc@gmail.com", "payloadc@2013", 995, True)
        'MsgBox(m.IsConnected)
        'Dim s As New Mailling.SMTP
        'MsgBox(Mailling.SMTP.SendEMail("payloadc@gmail.com", "eslam@hts-egypt.com;", "test dot net", "hellow me", True, "smtp.gmail.com", , , , , 587, True, "payloadc@gmail.com", "payloadc@2013"))


        Public Class SMTP
            Private Smtp_Server As New SmtpClient()

            '*******************************************************************************
            ' Function Name   : BrainDeadSimpleEmailSend
            ' Purpose         : Send super simple email message (works with most SMTP servers)
            '===============================================================================
            'NOTES: strFrom   : Full email address of who is sending the email. ie, David Dingus <daviddingus@att.net>
            '       strTo     : Full email address of who to send the email to. ie, "Bubba Dingus" <bob.dingus@cox.com>
            '       strSubject: Brief text regarding what the email concerns.
            '       strBody   : text that comprises the message body of the email.
            '       smtpHost  : This is the email host you are using for sending emails, such
            '                 : as "smtp.comcast.net", "authsmtp.juno.com", etc.
            '*******************************************************************************
            Public Shared Sub BrainDeadSimpleEmailSend(ByVal strFrom As String, _
                                                ByVal strTo As String, _
                                                ByVal strSubject As String, _
                                                ByVal strBody As String, _
                                                ByVal smtpHost As String)
                Dim smtpEmail As New Mail.SmtpClient(smtpHost)      'create new SMTP client using TCP Port 25
                smtpEmail.Send(strFrom, strTo, strSubject, strBody) 'send email
            End Sub

            '*******************************************************************************
            ' Function Name   : QuickiEMail
            ' Purpose         : Send a simple email message (but packed with a lot of muscle)
            '===============================================================================
            'NOTES: strFrom   : Full email address of who is sending the email. ie, David Dingus <daviddingus@att.net>
            '       strTo     : Full email address of who to send the email to. ie, "Bubba Dingus" <bob.dingus@cox.com>
            '       strSubject: Brief text regarding what the email concerns.
            '       strBody   : text that comprises the message body of the email.
            '       smtpHost  : This is the email host you are using for sending emails, such
            '                 : as "smtp.gmail.com", "smtp.comcast.net", "authsmtp.juno.com", etc.
            '       smtpPort  : TCP Communications Port to use. Most servers default to 25, though 465 (SSL) or 587 (TLS) are becoming popular.
            '       usesSLL   : If this value is TRUE, then use SSL/TLS Authentication protocol for secure communications.
            '      SSLUsername: If usesSLL is True, this is the username to use for creating a credential. Leave blank if the same as strFrom.
            '      SSLPassword: If usesSLL is True, this is the password to use for creating a credential. If this field and SSLUsername
            '                 : are blank, then default credentials will be used (only works on local, intranet servers).
            '       SSLDomain : If creating a credential when a specific domain is required, set this parameter, otherwise, leave it blank.
            '*******************************************************************************

            Public Shared Function QuickiEMail(ByVal strFrom As String, _
                                       ByVal strTo As String, _
                                       ByVal strSubject As String, _
                                       ByVal strBody As String, _
                                       ByVal smtpHost As String, _
                              Optional ByVal smtpPort As Integer = 25, _
                              Optional ByVal usesSSL As Boolean = False, _
                              Optional ByVal SSLUsername As String = vbNullString, _
                              Optional ByVal SSLPassword As String = vbNullString, _
                              Optional ByVal SSLDomain As String = vbNullString) As Boolean
                Try
                    Dim smtpEmail As New Mail.SmtpClient(smtpHost, smtpPort)            'create new SMTP client
                    smtpEmail.EnableSsl = usesSSL                                       'true if SSL Authentication required
                    If usesSSL Then                                                     'SSL authentication required?
                        If Len(SSLUsername) = 0 AndAlso Len(SSLPassword) = 0 Then       'if both SSLUsername and SSLPassword are blank...
                            smtpEmail.UseDefaultCredentials = True                      'use default credentials
                        Else                                                            'otherwise, we must create a new credential
                            If Not CBool(Len(SSLUsername)) Then                         'if SSLUsername is blank, use strFrom
                                smtpEmail.Credentials = New NetworkCredential(strFrom, SSLPassword, SSLDomain)
                            Else
                                smtpEmail.Credentials = New NetworkCredential(SSLUsername, SSLPassword, SSLDomain)
                            End If
                        End If
                    End If
                    smtpEmail.Send(strFrom, strTo, strSubject, strBody)                 'send email using text/plain content type and QuotedPrintable encoding
                Catch e As Exception                                                    'if error, report it
                    MsgBox(e.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation, "Mail Send Error")
                    Return False                                                        'return a failure flag
                End Try
                Return True                                                             'if no error, then return a success flag
            End Function

            '*******************************************************************************
            ' Function Name   : SendEMail
            ' Purpose         : Send a more complex email message
            '===============================================================================
            'NOTES: strFrom   : Full email address of who is sending the email. ie, David Dingus <daviddingus@att.net>
            '       strTo     : Full email address of who to send the email to. ie, "Bubba Dingus" <bob.dingus@cox.com>
            '                 : If multiple recipients, separate each full email address using a semicolon (;)
            '       strSubject: Brief text regarding what the email concerns.
            '       strBody   : text that comprises the message body of the email. May be raw text or HTML code.
            '       IsHTML    : True if the strBody data is HTML, or the type of data that would be contained within an HTML Body block.
            '       smtpHost  : This is the email host you are using for sending emails, such
            '                 : as "smtp.gmail.com", "smtp.comcast.net", "authsmtp.juno.com", etc.
            '       AltView   : A System.Net.Mail.AlternateView object, such as Rich Text or HTML.
            '                 : If need be, set AltView.ContentType.MediaType and AltView.TransferEncoding to properly format the AlternateView.
            '                 : For example: AltView.ContentType.MediaType = Mime.MediaTypeNames.Text.Rtf
            '                 :              AltView.TransferEncoding = Mime.TransferEncoding.SevenBit
            '       StrCC     : Send "carbon copies" of email to this or these recipients.
            '                 : If multiple recipients, separate each full email address using a semicolon (;)
            '       strBcc    : Blind Carbon Copy. Hide this or these recipients from view by others.
            '                 : If multiple recipients, separate each full email address using a semicolon (;)
            '   strAttachments: A single filepath, or a list of filepaths to send to the recipient.
            '                 : If multiple attachments, separate each filepath using a semicolon (;) (C:\my data\win32.txt; c:\jokes.rtf)
            '                 : The contents of the attachments will be encoded and sent.
            '                 : If you wish to send the attachment by specifying content type (MediaType) and content transfer encoding
            '                 : (Encoding), then follow the attachment name with the MediaType and optional encoding (default is 
            '                 : application/octet-stream,Base64) by placing them within parentheses, and separated by a comma. For example:
            '                 :     C:\My Files\API32.txt (text/plain, SevenBit); C:\telnet.exe (application/octet-stream, Base64)
            '                 :         Where:  The MediaType is determined from the System.Net.Mime.MediaTypeNames class, which
            '                 :                 can specify Application, Image, or Text lists. For example, the above content type,
            '                 :                 "text\plain", was defined by acquiring System.Net.Mime.MediaTypeNames.Text.Plain.
            '                 :         The second parameter, Encoding, is determined by the following the values specified by the
            '                 :         System.Net.Mime.TrasperEncoding enumeration:
            '                 :                 QuotedPrintable   (acquired by System.Net.Mime.TransferEncoding.QuotedPrintable.ToString)
            '                 :                 Base64            (acquired by System.Net.Mime.TransferEncoding.Base64.ToString)
            '                 :                 SevenBit          (acquired by System.Net.Mime.TransferEncoding.SevenBit.ToString)
            '       smtpPort  : TCP Communications Port to use. Most servers default to 25.
            '       usesSLL   : If this value is TRUE, then use SSL Authentication protocol for secure communications.
            '      SSLUsername: If usesSLL is True, this is the username to use for creating a credential. Leave blank if the same as strFrom.
            '      SSLPassword: If usesSLL is True, this is the password to use for creating a credential. If this field and SSLUsername
            '                 : are blank, then default credentials will be used (only works on local, intranet servers).
            '       SSLDomain : If creating a credential when a specific domain is required, set this parameter, otherwise, leave it blank.
            '*******************************************************************************
            Public Shared Function SendEMail(ByVal strFrom As String, _
                                       ByVal strTo As String, _
                                       ByVal strSubject As String, _
                                       ByVal strBody As String, _
                                       ByVal IsHTML As Boolean, _
                                       ByVal smtpHost As String, _
                              Optional ByVal AltView As Mail.AlternateView = Nothing, _
                              Optional ByVal strCC As String = vbNullString, _
                              Optional ByVal strBcc As String = vbNullString, _
                              Optional ByVal strAttachments As String = vbNullString, _
                              Optional ByVal smtpPort As Integer = 25, _
                              Optional ByVal usesSSL As Boolean = False, _
                              Optional ByVal SSLUsername As String = vbNullString, _
                              Optional ByVal SSLPassword As String = vbNullString, _
                              Optional ByVal SSLDomain As String = vbNullString) As Boolean

                Dim Email As New Mail.MailMessage               'create a new mail message
                With Email
                    .From = New Mail.MailAddress(strFrom)       'add FROM to mail message (must be a Mail Address object)
                    '-------------------------------------------
                    Dim Ary() As String = Split(strTo, ";")     'add TO to mail message (possible list of email addresses; separated each with ";")
                    For Idx As Integer = 0 To UBound(Ary)
                        If Len(Trim(Ary(Idx))) <> 0 Then .To.Add(Trim(Ary(Idx))) 'add each TO recipent (primary recipients)
                    Next
                    '-------------------------------------------
                    .Subject = strSubject                       'add SUBJECT text line to mail message
                    '-------------------------------------------
                    .Body = strBody                             'add BODY text of email to mail message.
                    .IsBodyHtml = IsHTML                        'indicate if the message body is actually HTML text.
                    .IsBodyHtml = True
                    '-------------------------------------------
                    If AltView IsNot Nothing Then               'if an alternate view of plaint text message is defined...
                        .AlternateViews.Add(AltView)            'add the alternate view
                    End If
                    '-------------------------------------------
                    If CBool(Len(strCC)) Then                   'add CC (Carbon Copy) email addresses to mail message 
                        Ary = Split(strCC, ";")                 '(possible list of email addresses, separated each with ";")
                        For Idx As Integer = 0 To UBound(Ary)
                            If Len(Trim(Ary(Idx))) <> 0 Then .CC.Add(Trim(Ary(Idx))) 'add each recipent
                        Next
                    End If
                    '-------------------------------------------
                    If CBool(Len(strBcc)) Then                  'add Bcc (Blind Carbon Copy) email addresses to mail message 
                        Ary = Split(strBcc, ";")                '(possible list of email addresses; separated each with ";")
                        For Idx As Integer = 0 To UBound(Ary)
                            If Len(Trim(Ary(Idx))) <> 0 Then .Bcc.Add(Trim(Ary(Idx))) 'add each recipent (hidden recipents)
                        Next
                    End If
                    '-------------------------------------------
                    If CBool(Len(strAttachments)) Then                                  'add any attachments to mail message 
                        Ary = Split(strAttachments, ";")                                '(possible list of file paths, separated each with ";")
                        For Idx As Integer = 0 To UBound(Ary)                           'process each attachment
                            Dim attach As String = Trim(Ary(Idx))                       'get attachment data
                            If Len(attach) <> 0 Then                                    'if an attachment present...
                                Dim I As Integer = InStr(attach, "(")                   'check for formatting instructions
                                If CBool(I) Then                                        'formatting present?
                                    Dim Fmt As String                                   'yes, so set up format cache
                                    Fmt = Mid(attach, I + 1, Len(attach) - I - 1)       'get format data
                                    attach = Trim(VB.Left(attach, I - 1))               'strip format data from the attachment path
                                    Dim Atch As New Mail.Attachment(attach)             'create a new attachment
                                    Dim fmts() As String = Split(Fmt, ",")              'break formatting up
                                    For I = 0 To UBound(fmts)                           'process each format specification
                                        Fmt = Trim(fmts(I))                             'grab a format instruction
                                        If CBool(Len(Fmt)) Then                         'data defined?
                                            Select Case I                               'yes, so determine which type of instruction to process
                                                Case 0                                  'index 0 specified MediaType
                                                    Atch.ContentType.MediaType = Fmt    'set media type to attachment
                                                Case 1                                  'index 1 specifes Encoding
                                                    Select Case LCase(Fmt)              'check the encoding types and process accordingly
                                                        Case "quotedprintable", "quoted-printable"
                                                            Atch.TransferEncoding = Mime.TransferEncoding.QuotedPrintable
                                                        Case "sevenbit", "7bit"
                                                            Atch.TransferEncoding = Mime.TransferEncoding.SevenBit
                                                        Case Else
                                                            Atch.TransferEncoding = Mime.TransferEncoding.Base64
                                                    End Select
                                            End Select
                                        End If
                                    Next
                                    .Attachments.Add(Atch)                              'add attachment to email
                                Else
                                    .Attachments.Add(New Mail.Attachment(attach)) 'add filepath (if no format specified, encoded in effiecient Base64)
                                End If
                            End If
                        Next
                    End If
                End With
                '-----------------------------------------------------------------------
                'now open the email server...
                Try
                    Dim SmtpEmail As New Mail.SmtpClient(smtpHost, smtpPort)            'create new SMTP client on the SMTP server
                    SmtpEmail.EnableSsl = usesSSL                                       'true if SSL Authentication required
                    If usesSSL Then                                                     'SSL authentication required?
                        If Len(SSLUsername) = 0 AndAlso Len(SSLPassword) = 0 Then       'if both SSLUsername and SSLPassword are blank...
                            SmtpEmail.UseDefaultCredentials = True                      'use default credentials
                        Else                                                            'otherwise, we must create a new credential
                            If Not CBool(Len(SSLUsername)) Then                         'if SSLUsername is blank, use strFrom
                                SmtpEmail.Credentials = New NetworkCredential(strFrom, SSLPassword, SSLDomain)
                            Else
                                SmtpEmail.Credentials = New NetworkCredential(SSLUsername, SSLPassword, SSLDomain)
                            End If
                        End If
                    End If
                    SmtpEmail.Send(Email)                                               'finally, send the email...
                Catch e As Exception                                                    'if error, report it
                    MsgBox(e.Message, MsgBoxStyle.OkOnly Or MsgBoxStyle.Exclamation, "Mail Error")
                    Return False                                                        'return failure flag
                End Try
                Return True                                                             'return success flag
            End Function
        End Class


        Public Class POP3
            Inherits Sockets.TcpClient                      'this class shall inherit all the functionality of a TC/IP Client

            Dim Stream As Sockets.NetworkStream             'non-SSL stream object
            Dim UsesSSL As Boolean = False                  'True if SLL authentication required
            Dim SslStream As Net.Security.SslStream             'set to SSL stream supporting SSL authentication if UsesSSL is True
            Dim SslStreamDisposed As Boolean = False        'true if we disposed of SSL Stream object
            Public LastLineRead As String = vbNullString    'copy of the last response line read from the TCP server

            '*******************************************************************************
            ' Sub Name   : Connect          (This is the first the we do with a POP3 object)
            ' Purpose    : Connect to the server using the Server, User Name, Password,
            '            : and a flag indicating if SSL authentication is required
            '            :
            ' Returns    : Nothing
            '            :
            ' Typical TelNet I/O:
            'telnet mail.domain.net 110                     (submit)
            '+OK POP3 mail.domain.net v2011.83 server ready
            'USER myusername                                (submit)
            '+OK User name accepted, password please
            'PASS mysecretpassword                          (submit)
            '+OK Mailbox open, 3 messages                   (the server locks and opens the appropriate maildrop)
            '*******************************************************************************
            Public Overloads Sub Connect(ByVal Server As String, _
                                         ByVal Username As String, _
                                         ByVal Password As String, _
                                         Optional ByVal InPort As Integer = 110, _
                                         Optional ByVal UseSSL As Boolean = False)

                If Connected Then Disconnect() '            'check underlying boolean flag to see if we are presently connected,
                'and if so, disconnect that session
                UsesSSL = UseSSL                            'set flag True or False for SSL authentication
                MyBase.Connect(Server, InPort)              'now connect to the server via our base class
                Stream = MyBase.GetStream                   'before we can check for a response, we first have to set up a non-SSL stream
                If UsesSSL Then                             'do we also need to use SSL authentication?
                    SslStream = _
                    New Net.Security.SslStream(Stream)          'yes, so build an SSL stream object on top of the non-SSL Network Stream
                    SslStream.AuthenticateAsClient(Server)  'add authentication as a client to the server
                End If

                If Not CheckResponse() Then Exit Sub '      'exit if an error was encountered

                If CBool(Len(Username)) Then                'if the username is defined (some servers will reject submissions)
                    Me.Submit("USER " & Username & vbCrLf)  'submit user name
                    If Not CheckResponse() Then Exit Sub '  'exit if an error was encountered
                End If

                If CBool(Len(Password)) Then                'if the password is defined (some servers will reject submissions)
                    Me.Submit("PASS " & Password & vbCrLf)  'submit password
                    If Not CheckResponse() Then Exit Sub '  'exit if an error was encountered
                End If
            End Sub

            '*******************************************************************************
            ' Function Name : CheckResponse
            ' Purpose       : Check the response to a POP3 command
            '               :
            ' Returns       : Boolean flag. True = Success, False = Failure
            '               :
            ' NOTE          : All status responses from the server begin with:
            '               : +OK                     (OK; Success, or request granted)
            '            or : -ERR                    (NAGATIVE; error)
            '*******************************************************************************
            Public Function CheckResponse() As Boolean
                If Not IsConnected() Then Return False '    'exit if not in TRANSACTION mode
                LastLineRead = Me.Response                  'check response (and save response line)
                If (Left(LastLineRead, 3) <> "+OK") Then    'OK?
                    Throw New POP3Exception(LastLineRead)   'no, so throw an exception
                    Return False                            'return failure flag
                End If
                Return True                                 'else return success flag
            End Function

            '*******************************************************************************
            ' Function Name : IsConnected
            ' Purpose       : Return connected to Server state, throw error if not
            '               :
            ' Returns       : Boolean Flag. True  if connected to server
            '               :
            '*******************************************************************************
            Public Function IsConnected() As Boolean
                If Not Connected Then                   'if not connected, throw an exception
                    Throw New POP3Exception("Not Connected to an POP3 Server.")
                    Return False                        'return failure flag
                End If
                Return True                             'Indicate that we are in the TRANSACTION state)
            End Function

            '*******************************************************************************
            ' Function Name : Response 
            ' Purpose       : get response from server (read from the mail stream into a buffer)
            '               :
            ' Returns       : string of data from the server
            '               :
            ' NOTE          : If a dataSize value  > 1 is supplied, then those number of bytes will be streamed in.
            '               : Otherwise, the data will be read in a line at a time, and end with the line end code (Linefeed (vbLf) 10 decimal)
            '*******************************************************************************
            Public Function Response(Optional ByVal dataSize As Integer = 1) As String
                Dim enc As New ASCIIEncoding                                'medium for ASCII representation of Unicode characters
                Dim ServerBufr() As Byte                                    'establish buffer
                Dim Index As Integer = 0                                    'init server buffer index and character counter
                If dataSize > 1 Then                                        'did invoker specify a data length to read?
                    '-------------------------------------------------------
                    ReDim ServerBufr(dataSize - 1)                          'size to dataSize to read as a single stream block (allow for 0 index)
                    Dim dtsz As Integer = dataSize
                    Dim sz As Integer                                       'variable to store actual number of bytes read from the stream
                    Do While Index < dataSize                               'while we have not read the entire message...
                        If UsesSSL Then                                     'process through SSL Stream if secure stream
                            sz = SslStream.Read(ServerBufr, Index, dtsz)    'read a server-defined block of data from SSLstream
                        Else                                                'else process through general TCP Stream
                            sz = Stream.Read(ServerBufr, Index, dtsz)       'read a server-defined block of data from Network Stream
                        End If
                        If sz = 0 Then Return vbNullString '                'we lost data, so we could not read the string
                        Index += sz                                         'bump index for data count actually read
                        dtsz -= sz                                          'drop amount left in buffer
                    Loop
                Else '------------------------------------------------------
                    ReDim ServerBufr(255)                                  'initially dimension buffer to 256 bytes (including 0 offset)
                    Do
                        If UsesSSL Then                                     'process through SSL Stream if secure stream
                            ServerBufr(Index) = CByte(SslStream.ReadByte)   'read a byte from SSLstream
                        Else                                                'else process through general TCP Stream
                            ServerBufr(Index) = CByte(Stream.ReadByte)      'read a byte from Network stream
                        End If
                        If ServerBufr(Index) = -1 Then Exit Do '            'end of stream if -1 encountered
                        Index += 1                                          'bump our offset index and counter
                        If ServerBufr(Index - 1) = 10 Then Exit Do '        'done with line if Newline code (10; Linefeed) read in
                        If Index > UBound(ServerBufr) Then                  'if the index points past end of buffer...
                            ReDim Preserve ServerBufr(Index + 255)          'then bump buffer another 256 bytes (Inc Index), but keep existing data
                        End If
                    Loop                                                    'loop until line read in
                End If
                Return enc.GetString(ServerBufr, 0, Index)                  'decode from a byte array into a string and return the string
            End Function

            '*******************************************************************************
            ' Sub Name : Submit
            ' Purpose  : Submit a request to the server
            '          :
            ' Returns  : Nothing
            '          :
            ' NOTE     : Command name must be in UPPERCASE, such as "PASS pw1Smorf".
            '          : "pass pw1Smorf" would not be acceptable, though some servers do allow for this, we should never assume it.
            '*******************************************************************************
            Public Sub Submit(ByVal message As String)
                Dim enc As New ASCIIEncoding                            'medium for ASCII representation of Unicode characters
                Dim WriteBuffer() As Byte = enc.GetBytes(message)       'converts the submitted string into to a sequence of bytes
                If UsesSSL Then                                         'using SSL authentication?
                    SslStream.Write(WriteBuffer, 0, WriteBuffer.Length) 'yes, so write SSL buffer using the SslStream object
                Else
                    Stream.Write(WriteBuffer, 0, WriteBuffer.Length)    'else write to Network buffer using the non-SSL object
                End If
            End Sub

            '*******************************************************************************
            ' Sub Name : Disconnect          (This is the last the we do with a POP3 object)
            ' Purpose  : Disconnect from the server and have it enter the UPDATE mode
            '          :
            ' Returns  : Nothing
            '          :
            ' Typical telNet I/O:
            'QUIT           (submit)
            '+OK Sayonara
            '
            ' NOTE:  When the client issues the QUIT command from the TRANSACTION state,
            '        the POP3 session enters the UPDATE state. (Note that if the client
            '        issues the QUIT command from the AUTHORIZATION state, the POP3
            '        session terminates but does NOT enter the UPDATE state.)
            '
            '        If a session terminates for some reason other than a client-issued
            '        QUIT command, the POP3 session does NOT enter the UPDATE state and
            '        MUST NOT remove any messages from the maildrop.
            '
            '        The POP3 server removes all messages marked as deleted from the
            '        maildrop and replies as to the status of this operation. If there
            '        is an error, such as a resource shortage, encountered while removing
            '        messages, the maildrop may result in having some or none of the
            '        messages marked as deleted be removed. In no case may the server
            '        remove any messages not marked as deleted.
            '
            '        Whether the removal was successful or not, the server then releases
            '        any exclusive-access lock on the maildrop and closes the TCP connection.
            '*******************************************************************************
            Public Sub Disconnect()
                Me.Submit("QUIT" & vbCrLf)  'submit quit request
                CheckResponse()             'check response
                If UsesSSL Then             'SSL authentication used?
                    SslStream.Dispose()     'dispose of created SSL stream object if so
                    SslStreamDisposed = True
                End If
            End Sub

            '*******************************************************************************
            ' Function Name : Statistics
            ' Purpose       : Get the number of email messages and the total size as any integer array
            '               :
            ' Returns       : 2-selement interger array.
            '               :   Element(0) is the number of user email messages on the server
            '               :   Element(1) is the total bytes of all messages taken up on the server
            '               :
            ' Typical telNet I/O:
            'STAT               (submit)
            '+OK 3 16487        (3 records (emails/messages) totaling 16487 bytes (octets))
            '*******************************************************************************
            Public Function Statistics() As Integer()
                If Not IsConnected() Then Return Nothing '          'exit if not in TRANSACTION mode
                Me.Submit("STAT" & vbCrLf)                          'submit Statistics request
                LastLineRead = Me.Response                          'check response
                If (Left(LastLineRead, 3) <> "+OK") Then            'OK?
                    Throw New POP3Exception(LastLineRead)           'no, so throw an exception
                    Return Nothing                                  'return failure flag
                End If
                Dim msgInfo() As String = Split(LastLineRead, " "c) 'separate by spaces, which divide its fields
                Dim Result(1) As Integer
                Result(0) = Integer.Parse(msgInfo(1))               'get the number of emails
                Result(1) = Integer.Parse(msgInfo(2))               'get the size of the email messages
                Return Result
            End Function

            '*******************************************************************************
            ' Function Name : List
            ' Purpose       : Get the drop listing from the maildrop
            '               :
            ' Returns       : Any Arraylist of POP3Message objects
            '               :
            ' Typical telNet I/O:
            'LIST            (submit)
            '+OK Mailbox scan listing follows
            '1 2532          (record index and size in bytes)
            '2 1610
            '3 12345
            '.               (end of records terminator)
            '*******************************************************************************
            Public Function List() As ArrayList
                If Not IsConnected() Then Return Nothing '          'exit if not in TRANSACTION mode

                Me.Submit("LIST" & vbCrLf)                          'submit List request
                If Not CheckResponse() Then Return Nothing '        'check for a response, but if an error, return nothing
                '
                'get a list of emails waiting on the server for the authenticated user
                '
                Dim retval As New ArrayList                         'set aside message list storage
                Do
                    Dim response As String = Me.Response            'check response
                    If (response = "." & vbCrLf) Then               'done with list?
                        Exit Do                                     'yes
                    End If
                    Dim msg As New POP3Message                      'establish a new message
                    Dim msgInfo() As String = Split(response, " "c) 'separate by spaces, which divide its fields
                    msg.MailID = Integer.Parse(msgInfo(0))          'get the list item number
                    msg.ByteCount = Integer.Parse(msgInfo(1))           'get the size of the email message
                    msg.Retrieved = False                           'indicate its message body is not yet retreived
                    retval.Add(msg)                                 'add a new entry into the retrieval list
                Loop
                Return retval                                       'return the list
            End Function

            '*******************************************************************************
            ' Function Name : GetHeader
            ' Purpose       : Grab the email header and optionally a number of lines from the body
            '               :
            ' Returns       : Gets the Email header of the selected email. If an integer value is
            '               : provided, that number of body lines will be returned. The returned
            '               : object is the submitted POP3Message.
            '               :
            ' Typical telNet I/O:
            'TOP 1 0            (submit request for record 1's message header only, 0=no lines of body)
            '+OK Top of message follows
            ' xxxxx             (header for current record is transmitted)
            '.                  (end of record terminator)
            '
            'TOP 1 10           (submit request for record 1's message header plus 10 lines of body data)
            '+OK Top of message follows
            ' xxxxx             (header for current record is transmitted)
            ' xxxxx             (first 10 lines of body)
            '.                  (end of record terminator)
            '*******************************************************************************
            Public Function GetHeader(ByRef msg As POP3Message, Optional ByVal BodyLines As Integer = 0) As POP3Message
                If Not IsConnected() Then Return Nothing '  'exit if not in TRANSACTION mode
                Me.Submit("TOP " & msg.MailID.ToString & " " & BodyLines.ToString & vbCrLf)
                If Not CheckResponse() Then Return Nothing ''check for a response, but if an error, return nothing
                msg.Message = vbNullString                  'erase current contents of the message, if any
                '
                'now process message data by binding the lines into a single string
                '
                Do
                    Dim response As String = Me.Response    'grab message line
                    If response = "." & vbCrLf Then         'end of data?
                        Exit Do                             'yes, done with the loop if so
                    End If
                    msg.Message &= response                 'else build message by appending the new line
                Loop
                Return msg                                  'return new filled Message object
            End Function

            '*******************************************************************************
            ' Function Name : Retrieve
            ' Purpose       : Retrieve email from POP3 server for the provided POP3Message object
            '               :
            ' Returns       : The submitted POP3 Message object with its Message property filled,
            '               : and its ByteCount property properly fitted to the message size.
            '               :
            ' NOTE          : Some email servers are set up to automatically delete an email once
            '               : it is retrieved from the server. Outlook, Outlook Express, and
            '               : Windows Mail do this. It is an option under Juno and Gmail. So, if we
            '               : do not submit a POP3 QUIT (the Disconnect() method), but just close
            '               : out the POP3 object, the message(s) will not be deleted.
            '               : Even so, most Windows-based server-processors will add an additional
            '               : CR for each LF, but the reported email size does not account for them.
            '               : So we must retreive more data to account for this.
            '
            ' Typical telNet I/O:
            'RETR 1             (submit request to retrive record index 1 (cannot be an index marked for deletion))
            '+OK 2532 octets    (an octet is a fancy term for a 8-bit byte)
            ' xxxx              (message header and message are retreived)
            '.                  (end of record terminator)
            '*******************************************************************************
            Public Function Retrieve(ByRef msg As POP3Message) As POP3Message
                If Not IsConnected() Then Return Nothing '          'exit if not in TRANSACTION mode
                Me.Submit("RETR " & msg.MailID.ToString & vbCrLf)   'issue request for indicated message number
                If Not CheckResponse() Then Return Nothing '        'check for a response, but if an error, return nothing
                msg.Message = Me.Response(msg.ByteCount)            'grab message line
                'the stream reader automatically convers the NewLine code, vbLf, to vbCrLf, so the files is not yet
                'fully read. For example, a files that was 233 lines will therefore have 233 more characters not
                'yet read from the files when it has reached its reported data size. So we will scan these in.
                'But even if this was not the case, the trailing "." & vbCrLf is still pending...
                Do
                    Dim S As String = Response()                    'grab more data
                    If S = "." & vbCrLf Then                        'end of data?
                        Exit Do                                     'If so, then exit loop
                    End If
                    msg.Message &= S                                'else tack data to end of message
                Loop                                                'keep trying
                msg.ByteCount = Len(msg.Message)                    'ensure full size updated
                Return msg                                          'return new message object
            End Function

            '*******************************************************************************
            ' Sub Name : Delete
            ' Purpose  : Delete an email
            '          :
            ' Returns  : Nothing
            '          :
            ' NOTE     : Some email servers are set up to automatically delete an email once
            '          : it is retrieved from the server. Outlook, Outlook Express, and
            '          : Windows Mail do this. It is an option under Juno and Gmail.
            '
            ' Typical telNet I/O:
            'DELE 1             (submit request to delete record index 1)
            '+OK Message deleted
            '*******************************************************************************
            Public Sub Delete(ByVal msgHdr As POP3Message)
                If Not IsConnected() Then Exit Sub '                    'exit if not in TRANSACTION mode
                Me.Submit("DELE " & msgHdr.MailID.ToString & vbCrLf)    'submit Delete request
                CheckResponse()                                         'check response
            End Sub

            '*******************************************************************************
            ' Sub Name : Reset
            ' Purpose  : Reset any deletion (automatic or manual) of all email from
            '          : the current session.
            '          :
            ' Returns  : Nothing
            '          :
            ' Typical telNet I/O:
            'RSET               (submit)
            '+OK Reset state
            '*******************************************************************************
            Public Sub Reset()
                If Not IsConnected() Then Exit Sub ''exit if not in TRANSACTION mode
                Me.Submit("RSET" & vbCrLf)          'submit Reset request
                CheckResponse()                     'check response
            End Sub

            '*******************************************************************************
            ' Function Name : NOOP (No Operation)
            ' Purpose       : Does nothing. Juts gets a position response from the server
            '               :
            ' Returns       : Boolean flag. False if disconnected, else True if connected.
            '               :
            ' NOTE          : This NO OPERATION command is useful when you have a server that
            '               : automatically disconnects after a certain idle period of activity.
            '               : This command can be issued by a timer that also monitors users
            '               : inactivity, and issues a NOOP to reset the server timer.
            '               :
            ' Typical telNet I/O:
            'NOOP               (submit)
            '+OK
            '*******************************************************************************
            Public Function NOOP() As Boolean
                If Not IsConnected() Then Return False 'exit if not in TRANSACTION mode
                Me.Submit("NOOP")
                Return CheckResponse()
            End Function

            '*******************************************************************************
            ' Function Name : Finalize
            ' Purpose       : remove SSL Stream object if not removed
            '*******************************************************************************
            Protected Overrides Sub Finalize()
                If SslStream IsNot Nothing AndAlso Not SslStreamDisposed Then   'SSL Stream object Disposed?
                    SslStream.Dispose()                                         'no, so do it
                End If
                MyBase.Finalize()                                               'then do normal finalization
            End Sub
        End Class

        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        '-------------------------------------------------------------------------------
        ' Class Name : POP3Message
        ' Purpose    : POP3 message data
        '-------------------------------------------------------------------------------
        Public Class POP3Message
            Public MailID As Integer = 0                'message number
            Public ByteCount As Integer = 0             'length of message in bytes
            Public Retrieved As Boolean = False         'flag indicating if the message has been retrieved
            Public Message As String = vbNullString     'the text of the message

            Public Overrides Function ToString() As String
                Return Message
            End Function
        End Class

        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
        '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

        '-------------------------------------------------------------------------------
        ' Class Name : POP3Exception
        ' Purpose    : process exception
        ' NOTE       : This is a normal exception, but we wrap it to give it an identy
        '            : that can be associated with our POP3 class
        '-------------------------------------------------------------------------------
        Public Class POP3Exception
            Inherits ApplicationException

            Public Sub New(ByVal str As String)
                MyBase.New(str)
            End Sub
        End Class


    End Namespace

    Namespace KeyboadDriver
        Public Class KeyboadDriver
            Private keyPressed As Object
            Private charCount As Int32
            Private lineLimit As Int32 = 69
            Private addKey As Object
            Private tmr As New System.Windows.Forms.Timer
            Public Event GetKey(ByVal key As String)


            Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vkey As Integer) As Short
            Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Integer) As Short

            Public Sub StartLog()
                tmr.Enabled = True
            End Sub

            Public Sub StopLog()
                tmr.Enabled = False
            End Sub

            Public Function getCapslock() As Boolean
                'return or set the caps lock toggle
                getCapslock = CBool(GetKeyState(System.Windows.Forms.Keys.Capital) And 1)
            End Function

            Public Function getShift() As Boolean
                'check to see if the shift key is pressed
                getShift = CBool(GetAsyncKeyState(System.Windows.Forms.Keys.ShiftKey))
            End Function


            Public Sub New()
                AddHandler tmr.Tick, AddressOf Timer_Tick
                tmr.Interval = 50
                tmr.Enabled = False
            End Sub

            Private Sub Timer_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs)
                HandleKey()
            End Sub

            Private Sub HandleKey()
                On Error Resume Next
                Dim i As Object = Nothing
                'check Enter key
                keyPressed = GetAsyncKeyState(13)
                If keyPressed = -32767 Then
                    charCount = 0
                    addKey = vbCrLf
                    GoTo KeyFound
                End If

                'check for backspace
                keyPressed = GetAsyncKeyState(8)
                If keyPressed = -32767 Then
                    addKey = "[bs]"
                    charCount += 4
                    GoTo KeyFound
                End If

                'check for space bar
                keyPressed = GetAsyncKeyState(32)
                If keyPressed = -32767 Then
                    addKey = " "
                    GoTo KeyFound
                    charCount += 1
                End If

                'check for colon/semicolon
                keyPressed = GetAsyncKeyState(186)
                If keyPressed = -32767 Then
                    If getShift() = False Then
                        addKey = ";"
                    Else
                        addKey = ":"
                    End If
                    GoTo KeyFound
                    charCount += 1
                End If

                'check for =/+
                keyPressed = GetAsyncKeyState(187)
                If keyPressed = -32767 Then
                    If getShift() = False Then
                        addKey = "="
                    Else
                        addKey = "+"
                    End If
                    GoTo KeyFound
                    charCount += 1
                End If

                'check for ,/<
                keyPressed = GetAsyncKeyState(188)
                If keyPressed = -32767 Then
                    If getShift() = False Then
                        addKey = ","
                    Else
                        addKey = "<"
                    End If
                    GoTo KeyFound
                    charCount += 1
                End If

                'check for -/_
                keyPressed = GetAsyncKeyState(189)
                If keyPressed = -32767 Then
                    If getShift() = False Then
                        addKey = "-"
                    Else
                        addKey = "_"
                    End If
                    GoTo KeyFound
                    charCount += 1
                End If

                'check for ./>
                keyPressed = GetAsyncKeyState(190)
                If keyPressed = -32767 Then
                    If getShift() = False Then
                        addKey = "."
                    Else
                        addKey = ">"
                    End If
                    GoTo KeyFound
                    charCount += 1
                End If

                'check for //?
                keyPressed = GetAsyncKeyState(191)
                If keyPressed = -32767 Then
                    If getShift() = False Then
                        addKey = "/"
                    Else
                        addKey = "?"
                    End If
                    GoTo KeyFound
                    charCount += 1
                End If

                'check for `/~
                keyPressed = GetAsyncKeyState(192)
                If keyPressed = -32767 Then
                    If getShift() = False Then
                        addKey = "`"
                    Else
                        addKey = "~"
                    End If
                    GoTo KeyFound
                    charCount += 1
                End If

                'check for 0/)
                keyPressed = GetAsyncKeyState(96)
                If keyPressed = -32767 Then
                    If getShift() = False Then
                        addKey = "0"
                    Else
                        addKey = ")"
                    End If
                    GoTo KeyFound
                    charCount += 1
                End If

                'check for 1/!
                keyPressed = GetAsyncKeyState(97)
                If keyPressed = -32767 Then
                    If getShift() = False Then
                        addKey = "1"
                    Else
                        addKey = "!"
                    End If
                    GoTo KeyFound
                    charCount += 1
                End If

                'check for 2/@
                keyPressed = GetAsyncKeyState(98)
                If keyPressed = -32767 Then
                    If getShift() = False Then
                        addKey = "2"
                    Else
                        addKey = "@"
                    End If
                    GoTo KeyFound
                    charCount += 1
                End If

                'check for 3/#
                keyPressed = GetAsyncKeyState(99)
                If keyPressed = -32767 Then
                    If getShift() = False Then
                        addKey = "3"
                    Else
                        addKey = "#"
                    End If
                    GoTo KeyFound
                    charCount += 1
                End If

                'check for 4/$
                keyPressed = GetAsyncKeyState(100)
                If keyPressed = -32767 Then
                    If getShift() = False Then
                        addKey = "4"
                    Else
                        addKey = "$"
                    End If
                    GoTo KeyFound
                    charCount += 1
                End If

                'check for 5/%
                keyPressed = GetAsyncKeyState(101)
                If keyPressed = -32767 Then
                    If getShift() = False Then
                        addKey = "5"
                    Else
                        addKey = "%"
                    End If
                    GoTo KeyFound
                    charCount += 1
                End If

                'check for 6/^
                keyPressed = GetAsyncKeyState(102)
                If keyPressed = -32767 Then
                    If getShift() = False Then
                        addKey = "6"
                    Else
                        addKey = "7"
                    End If
                    GoTo KeyFound
                    charCount += 1
                End If

                'check for 7/&
                keyPressed = GetAsyncKeyState(103)
                If keyPressed = -32767 Then
                    If getShift() = False Then
                        addKey = "7"
                    Else
                        addKey = "&"
                    End If
                    GoTo KeyFound
                    charCount += 1
                End If

                'check for 8/*
                keyPressed = GetAsyncKeyState(104)
                If keyPressed = -32767 Then
                    If getShift() = False Then
                        addKey = "8"
                    Else
                        addKey = "*"
                    End If
                    GoTo KeyFound
                    charCount += 1
                End If

                'check for 9/(
                keyPressed = GetAsyncKeyState(105)
                If keyPressed = -32767 Then
                    If getShift() = False Then
                        addKey = "9"
                    Else
                        addKey = "("
                    End If
                    GoTo KeyFound
                    charCount += 1
                End If

                'other num/special chars
                keyPressed = GetAsyncKeyState(106)
                If keyPressed = -32767 Then
                    If getShift() = False Then
                        addKey = "*"
                        charCount += 1
                    Else
                        addKey = ""
                    End If
                    GoTo KeyFound
                End If

                keyPressed = GetAsyncKeyState(107)
                If keyPressed = -32767 Then
                    If getShift() = False Then
                        addKey = "+"
                    Else
                        addKey = "="
                    End If
                    GoTo KeyFound
                    charCount += 1
                End If

                keyPressed = GetAsyncKeyState(108)
                If keyPressed = -32767 Then
                    addKey = ""
                    GoTo KeyFound
                End If

                keyPressed = GetAsyncKeyState(109)
                If keyPressed = -32767 Then
                    If getShift() = False Then
                        addKey = "-"
                    Else
                        addKey = "_"
                    End If
                    GoTo KeyFound
                    charCount += 1
                End If

                keyPressed = GetAsyncKeyState(110)
                If keyPressed = -32767 Then
                    If getShift() = False Then
                        addKey = "."
                    Else
                        addKey = ">"
                    End If
                    GoTo KeyFound
                    charCount += 1
                End If

                keyPressed = GetAsyncKeyState(111)
                If keyPressed = -32767 Then
                    addKey = "/"
                    GoTo KeyFound
                    charCount += 1
                End If

                keyPressed = GetAsyncKeyState(2)
                If keyPressed = -32767 Then
                    If getShift() = False Then
                        addKey = "/"
                    Else
                        addKey = "?"
                    End If
                    GoTo KeyFound
                    charCount += 1
                End If

                keyPressed = GetAsyncKeyState(220)
                If keyPressed = -32767 Then
                    If getShift() = False Then
                        addKey = "\"
                    Else
                        addKey = "|"
                    End If
                    GoTo KeyFound
                    charCount += 1
                End If

                keyPressed = GetAsyncKeyState(222)
                If keyPressed = -32767 Then
                    If getShift() = False Then
                        addKey = "'"
                    Else
                        addKey = Chr(34)
                    End If
                    GoTo KeyFound
                    charCount += 1
                End If

                keyPressed = GetAsyncKeyState(221)
                If keyPressed = -32767 Then
                    If getShift() = False Then
                        addKey = "]"
                    Else
                        addKey = "}"
                    End If
                    GoTo KeyFound
                    charCount += 1
                End If

                keyPressed = GetAsyncKeyState(219)
                If keyPressed = -32767 Then
                    If getShift() = False Then
                        addKey = "["
                    Else
                        addKey = "{"
                    End If
                    GoTo KeyFound
                    charCount += 1
                End If

                'check for a-z upper and lower case
                For i = 65 To 128
                    keyPressed = GetAsyncKeyState(i)
                    If keyPressed = -32767 Then
                        If getShift() = False Then
                            If getCapslock() = True Then
                                addKey = UCase(Chr(i))
                            Else
                                addKey = LCase(Chr(i))
                            End If
                        Else
                            If getCapslock() = False Then
                                addKey = UCase(Chr(i))
                            Else
                                addKey = LCase(Chr(i))
                            End If
                        End If
                        GoTo KeyFound
                        charCount += 1
                    End If
                Next i

                For i = 48 To 57
                    keyPressed = GetAsyncKeyState(i)
                    If keyPressed = -32767 Then
                        If getShift() = True Then
                            Select Case Val(Chr(i))
                                Case 1
                                    addKey = "!"
                                Case 2
                                    addKey = "@"
                                Case 3
                                    addKey = "#"
                                Case 4
                                    addKey = "$"
                                Case 5
                                    addKey = "%"
                                Case 6
                                    addKey = "^"
                                Case 7
                                    addKey = "&"
                                Case 8
                                    addKey = "*"
                                Case 9
                                    addKey = "("
                                Case 0
                                    addKey = ")"
                            End Select
                        Else
                            addKey = Chr(i)
                        End If
                        GoTo KeyFound
                        charCount += 1
                    End If
                Next i

                System.Windows.Forms.Application.DoEvents()
                Exit Sub

                'keyfound 
KeyFound:
                If charCount > lineLimit Then
                    charCount = 0
                    addKey &= vbCrLf
                End If
                If addKey <> "" Then RaiseEvent GetKey(addKey)
                System.Windows.Forms.Application.DoEvents()
            End Sub
        End Class
    End Namespace

    Namespace Mac_YPS
        Public Class Mac_YPS

        End Class
    End Namespace

    Namespace MSOffice
        Public Class Excel
            Public Function GetSheetNames(ByVal Path As String) As List(Of String)
                Dim objA As New Microsoft.Office.Interop.Excel.Application
                Dim ls As New List(Of String)
                objA.Workbooks.Open(Path, False, True, , , , , , , False)
                For Each objSht As Microsoft.Office.Interop.Excel._Worksheet In objA.Sheets
                    ls.Add(objSht.Name)
                Next
                objA.Workbooks(1).Close(False)
                objA.Workbooks.Close()
                objA = Nothing
                Return ls
            End Function

            Public Function GetSheetData(ByVal WorkBookPath As String, ByVal SheetName As String) As DataTable
                Dim cn As System.Data.OleDb.OleDbConnection
                Dim cmd As System.Data.OleDb.OleDbDataAdapter
                Dim ds As New DataSet
                Dim dt As New DataTable
                Dim dg As New DataGridView

                Try
                    cn = New System.Data.OleDb.OleDbConnection(String.Format("provider=Microsoft.ACE.OLEDB.12.0;data source={0};Extended Properties=Excel 12.0;", WorkBookPath))


                    ' Select the data from Sheet1 of the workbook.
                    cmd = New System.Data.OleDb.OleDbDataAdapter(String.Format("select * from [{0}$]", SheetName), cn)
                    cn.Open()
                    cmd.Fill(ds)
                    cn.Close()
                    dg.DataSource = ds.DefaultViewManager
                    dg.Refresh()
                    cmd.Fill(dt)
                    dg.DataSource = dt
                    Return dt
                Catch ex As Exception
                    Return Nothing
                End Try


                Return Nothing
            End Function
        End Class

    End Namespace


    Namespace FilesSystem
        Public Class CompressionSnippet

            Public Shared Sub Main()
                Dim path As String = "test.txt"

                ' Create the text file if it doesn't already exist.
                If Not File.Exists(path) Then
                    Console.WriteLine("Creating a new test.txt file")
                    Dim text() As String = {"This is a test text file.", _
                        "This file will be compressed and written to the disk.", _
                        "Once the file is written, it can be decompressed", _
                        "imports various compression tools.", _
                        "The GZipStream and DeflateStream class use the same", _
                        "compression algorithms, the primary difference is that", _
                        "the GZipStream class includes a cyclic redundancy check", _
                        "that can be useful for detecting data corruption.", _
                        "One other side note: both the GZipStream and DeflateStream", _
                        "classes operate on streams as opposed to file-based", _
                        "compression data is read on a byte-by-byte basis, so it", _
                        "is not possible to perform multiple passes to determine the", _
                        "best compression method. Already compressed data can actually", _
                        "increase in size if compressed with these classes."}

                    File.WriteAllLines(path, text)
                End If

                Console.WriteLine("Contents of {0}", path)
                Console.WriteLine(File.ReadAllText(path))

                CompressFile(path)
                Console.WriteLine()

                UncompressFile(path + ".gz")
                Console.WriteLine()

                Console.WriteLine("Contents of {0}", path + ".gz.txt")
                Console.WriteLine(File.ReadAllText(path + ".gz.txt"))

            End Sub

            Public Shared Sub CompressFile(ByVal path As String)
                Dim sourceFile As FileStream = File.OpenRead(path)
                Dim destinationFile As FileStream = File.Create(path + ".gz")

                Dim buffer(sourceFile.Length) As Byte
                sourceFile.Read(Buffer, 0, Buffer.Length)

                Using output As New GZipStream(destinationFile, _
                    CompressionMode.Compress)

                    Console.WriteLine("Compressing {0} to {1}.", sourceFile.Name, _
                        destinationFile.Name, False)

                    output.Write(buffer, 0, buffer.Length)
                End Using

                ' Close the files.
                sourceFile.Close()
                destinationFile.Close()
            End Sub

            Public Shared Sub UncompressFile(ByVal path As String)
                Dim sourceFile As FileStream = File.OpenRead(path)
                Dim destinationFile As FileStream = File.Create(path + ".txt")

                ' Because the uncompressed size of the file is unknown, 
                ' we are imports an arbitrary buffer size.
                Dim buffer(4096) As Byte
                Dim n As Integer

                Using input As New GZipStream(sourceFile, _
                    CompressionMode.Decompress, False)

                    Console.WriteLine("Decompressing {0} to {1}.", sourceFile.Name, _
                        destinationFile.Name)

                    n = input.Read(buffer, 0, buffer.Length)
                    destinationFile.Write(buffer, 0, n)
                End Using

                ' Close the files.
                sourceFile.Close()
                destinationFile.Close()
            End Sub
        End Class


        Public Class FileSearch
            Private st As New EAMS.StringFunctions.StringsFunction

            Public Sub Search(ByRef lstFullName As ListBox, ByRef lstName As ListBox, ByVal SourcePath As String, ByVal Suffix As String, Optional RemoveSuffix As Boolean = True)
                Dim SourceDir As DirectoryInfo = New DirectoryInfo(SourcePath)
                Dim pathIndex As Integer = SourcePath.LastIndexOf("\")

                ' the source directory must exist, otherwise throw an exception
                lstFullName.Items.Clear()
                lstName.Items.Clear()
                If SourceDir.Exists Then
                    Dim SubDir As DirectoryInfo
                    For Each SubDir In SourceDir.GetDirectories()
                        'Console.WriteLine(SubDir.Name)
                        Search(lstFullName, lstName, SubDir.FullName, Suffix)
                    Next


                    For Each childFile As FileInfo In SourceDir.GetFiles("*", SearchOption.AllDirectories).Where(Function(file) file.Extension.ToLower = "." & Suffix)
                        If RemoveSuffix Then
                            lstName.Items.Add(Replace(childFile.Name, Suffix, ""))
                        Else
                            lstName.Items.Add(childFile.Name)
                        End If

                        lstFullName.Items.Add(childFile.FullName)


                    Next
                Else
                    Throw New DirectoryNotFoundException("Source directory does not exist: " + SourceDir.FullName)
                End If

            End Sub

        End Class


        Public Class ReadWriteFile
            Public Function ReadTextFile(ByVal Path As String) As String
                Try
                    Dim obj As New System.IO.StreamReader(Path)
                    Return obj.ReadToEnd
                Catch ex As Exception
                    Return ""
                End Try
                Return ""
            End Function

        End Class
    End Namespace

    Namespace OfficeAutomation
        Public Class Excels
            Private objApp As New Excel.Application
            Private objBooks As Excel.Workbooks
            Private objBook As Excel._Workbook
            Private objSheets As Excel.Sheets
            Private objSheet As Excel._Worksheet
            Private objRange As Excel.Range

            Public Sub New()
                objBooks = objApp.Workbooks
                objBook = objBooks.Add
                objSheets = objBook.Worksheets
                objSheet = objSheets(1)
            End Sub

            Public Sub Refresh()
                objBook.RefreshAll()
                'objBook.Save()
            End Sub

            Public Sub SetRange(ByVal ColumnName As String, ByVal rows As Integer, ByVal columns As Integer)
                objRange = objSheet.Range(ColumnName, Reflection.Missing.Value)
                objRange = objRange.Resize(rows, columns)
            End Sub

            Public Sub FormateRange(ByVal IntColor As Color, ByVal FontColor As Color, Optional FontSize As Byte = 11, Optional IsBold As Boolean = False, Optional ColumnWidth As Byte = 0, Optional VAlign As Excel.XlVAlign = Excel.XlHAlign.xlHAlignLeft, Optional HAlign As Excel.XlHAlign = Excel.XlHAlign.xlHAlignLeft)
                Try
                    objRange.Interior.Color = IntColor
                    objRange.Font.Color = FontColor
                    objRange.Font.Bold = IsBold
                    objRange.Font.Size = FontSize
                    objRange.VerticalAlignment = VAlign
                    objRange.HorizontalAlignment = HAlign
                    If ColumnWidth > 0 Then objRange.ColumnWidth = ColumnWidth
                Catch ex As Exception

                End Try

            End Sub

            Public Sub CellBorder(ByVal cr As Color, Weight As Integer)
                objRange.Cells.Borders.Color = cr
                objRange.Cells.Borders.Weight = Weight
            End Sub

            Public Overloads Sub Write(ByVal value(,) As String)
                objRange.Value = value
            End Sub
            Public Overloads Sub Write(ByVal value(,) As Byte)
                objRange.Value = value
            End Sub
            Public Overloads Sub Write(ByVal value(,) As Integer)
                objRange.Value = value
            End Sub

            Public Overloads Sub Write(ByVal RowInx As Integer, ByVal ColInx As Integer, value As String)
                objRange(RowInx, ColInx) = value
            End Sub

            Public Sub Save(ByVal FilePath As String, ByVal SheetName As String)
                objSheet.Name = SheetName
                objBook.SaveAs(FilePath)
            End Sub

            Public Sub SaveBulkDataGrid(ByRef grd As DataGridView)
                Me.SetRange("A1", grd.Rows.Count + 1, grd.Columns.Count + 1)
                With objSheet
                    grd.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
                    grd.SelectAll()
                    Clipboard.SetDataObject(grd.GetClipboardContent())
                    .Paste()
                    Clipboard.Clear()
                End With
            End Sub

            Public Sub Close()
                objBook.Close()
                objBooks.Close()
            End Sub

            Protected Overrides Sub Finalize()
                objSheet = Nothing
                objSheets = Nothing
                objBook = Nothing
                objBooks = Nothing
                objApp = Nothing
                GC.Collect()
                MyBase.Finalize()
            End Sub
        End Class

    End Namespace

    Namespace WindowsForms
        Public Class Tree_View
            Private s As New EAMS.StringFunctions.Common
            Private lst As New ListBox


            Private Sub _CollectListViewCheckedItem(ByRef nds As TreeNodeCollection)
                If Not IsNothing(nds) Then
                    For Each anode As TreeNode In nds
                        If anode.Checked Then
                            lst.Items.Add(anode.FullPath)
                        End If
                        _CollectListViewCheckedItem(anode.Nodes)
                    Next
                End If
            End Sub
            Private Sub populateFilesAndFolders(parentNode As TreeNode, startingPath As String, ByVal SearchPatern As String)
                Dim inspectDirectoryInfo As IO.DirectoryInfo = New IO.DirectoryInfo(startingPath)
                For Each directoryInfoItem As IO.DirectoryInfo In inspectDirectoryInfo.GetDirectories
                    Dim directoryTreeNode As New TreeNode() With {.Tag = directoryInfoItem.FullName, .Text = directoryInfoItem.Name}
                    parentNode.Nodes.Add(directoryTreeNode)
                    populateFilesAndFolders(directoryTreeNode, directoryInfoItem.FullName, SearchPatern)
                Next
                For Each fileItem As IO.FileInfo In inspectDirectoryInfo.GetFiles(SearchPatern)
                    Dim fileNode As New TreeNode() With {.Tag = fileItem.FullName, .Text = s.SubString(s.SubString(fileItem.Name, ".bin"), ".fld")}
                    parentNode.Nodes.Add(fileNode)
                Next
            End Sub
            Public Function CollectListViewCheckedItem(ByRef tv As TreeView) As List(Of String)
                CollectListViewCheckedItem = New List(Of String)
                lst.Items.Clear()
                _CollectListViewCheckedItem(tv.Nodes)
                For inx As Integer = 0 To lst.Items.Count - 1
                    CollectListViewCheckedItem.Add(lst.Items(inx))
                Next
                Return CollectListViewCheckedItem
            End Function
            Public Sub LoadFoldersIntoTree(ByRef tv As TreeView, ByVal ParentName As String, ByVal ParentPath As String, ByVal SearchPatern As String)
                tv.Nodes.Clear()
                Dim ndParent As TreeNode = tv.Nodes.Add(ParentName)
                ndParent.Tag = ParentPath
                populateFilesAndFolders(tv.Nodes.Item(0), tv.Nodes.Item(0).Tag.ToString, SearchPatern)
            End Sub
            Public Function IsItemExists(ByRef Nds As TreeNode, Item As String) As Boolean
                For Each nd As TreeNode In Nds.Nodes
                    If Item = nd.Name Then
                        Return True
                    End If
                Next
                Return False
            End Function
        End Class
    End Namespace




End Namespace

#Region "Ref"
'DECLARE db_cursor CURSOR FAST_FORWARD FOR SELECT name, age, color FROM table; 
'DECLARE @myName VARCHAR(256);
'DECLARE @myAge INT;
'DECLARE @myFavoriteColor VARCHAR(40);
'OPEN db_cursor;
'FETCH NEXT FROM db_cursor INTO @myName, @myAge, @myFavoriteColor;
'WHILE @@FETCH_STATUS = 0  
'BEGIN  

'       --Do stuff with scalar values

'       FETCH NEXT FROM db_cursor INTO @myName, @myAge, @myFavoriteColor;
'END;
'CLOSE db_cursor;
'DEALLOCATE db_cursor;

'Private Sub ApplayFormat()
'    Dim styleFormatCondition1 As StyleFormatCondition = New DevExpress.XtraGrid.StyleFormatCondition()
'    styleFormatCondition1.Appearance.BackColor = System.Drawing.Color.Red
'    styleFormatCondition1.Appearance.Options.UseBackColor = True
'    styleFormatCondition1.ApplyToRow = True
'    styleFormatCondition1.Condition = DevExpress.XtraGrid.FormatConditionEnum.Expression
'    styleFormatCondition1.Expression = "[Remaining] > 1"
'    GridView3.FormatConditions.Add(styleFormatCondition1)
'    GridView3.RefreshData()
'    GridView3.RefreshEditor(True)
'End Sub


'Private Sub GridView3_RowCellStyle(sender As Object, e As DevExpress.XtraGrid.Views.Grid.RowCellStyleEventArgs) Handles GridView3.RowCellStyle
'    Dim View As Views.Grid.GridView = sender
'    If e.Column.FieldName = "New Quantity" Then
'        Dim RNew As String = View.GetRowCellDisplayText(e.RowHandle, View.Columns("New Quantity"))
'        Dim RRemaining As String = View.GetRowCellDisplayText(e.RowHandle, View.Columns("Remaining"))
'        If RNew >= RRemaining Then
'            e.Appearance.BackColor = Color.Maroon
'            e.Appearance.BackColor2 = Color.MistyRose
'        End If
'        If Not IsNumeric(RNew) Then
'            e.Appearance.BackColor = Color.DarkGoldenrod
'            e.Appearance.BackColor2 = Color.Gold
'        End If
'    End If
'End Sub

'groupindex
'GridView3.Columns("Subcontractor").GroupIndex = 1

'getselected grid item
'Dim rh As Integer = GridView3.GetSelectedRows(0)

'Public Sub GetData()
'    Dim obj As New System.IO.StreamReader(Application.StartupPath & "\Binaries\ItemStepProduction.dll")
'    q1 = obj.ReadToEnd
'    obj.Close()
'    grd.DataSource = DB.ReturnDataTable(String.Format("{0} having Item_ID={1} order by step_no", q1, ItemID))
'    GridView3.Columns(0).OptionsColumn.AllowEdit = False
'    GridView3.Columns(1).OptionsColumn.AllowEdit = False
'    GridView3.Columns(2).OptionsColumn.AllowEdit = False
'    GridView3.Columns(4).OptionsColumn.AllowEdit = False
'    GridView3.Columns(5).OptionsColumn.AllowEdit = False
'    GridView3.Columns(3).OptionsColumn.AllowEdit = False
'    GridView3.Columns(0).AppearanceCell.BackColor = Color.LightGray
'    GridView3.Columns(1).AppearanceCell.BackColor = Color.LightGray
'    GridView3.Columns(2).AppearanceCell.BackColor = Color.LightGray
'    GridView3.Columns(3).AppearanceCell.BackColor = Color.LightGray
'    GridView3.Columns(4).AppearanceCell.BackColor = Color.LightGray
'    GridView3.Columns(5).AppearanceCell.BackColor = Color.LightGray
'    GridView3.Columns(6).AppearanceCell.BackColor = Color.FromArgb(194, 241, 194)
'End Sub
#End Region

#Region "Play Frequency"
'Private Declare Function Beep Lib "kernel32" (ByVal soundFrequency As Int32, ByVal soundDuration As Int32) As Int32
#End Region

#Region "Daily Tracking"
'--- DAILY EICA PRODUCTION

'declare @col as nvarchar(max)
'declare @result as nvarchar(max)
'declare @startdate as char(10)
'declare @enddate as char(10)
'declare @startdate0 as date
'declare @enddate0 as date

'select @enddate0 =  [tmp_date] from [tblTMP] where [tmp_id]=1--convert(char(10),cast(getdate() as smalldatetime),111)
'select @startdate0 = dateadd(d,-6,@enddate0) --'04/11/2015'
'set @enddate =  convert(char(10),cast(@enddate0 as smalldatetime),111)
'set @startdate = convert(char(10),cast(@startdate0 as smalldatetime),111)


'select @col = STUFF((select distinct ',' + quotename(x.Date)
'FROM (select DATE FROM (select convert(char(10),dateadd(d,(x.a + (10 * x.b) + (100 * x.c) + (1000 * x.d)),cast(@startdate as smalldatetime)),111) as Date
'FROM (select * FROM (select 0 as a UNION ALL select 1 UNION ALL select 2 UNION ALL select 3 UNION ALL select 4 UNION ALL select 5 UNION ALL select 6 UNION ALL select 7 UNION ALL select 8 UNION ALL select 9) as a CROSS JOIN
'(select 0 as b UNION ALL select 1 UNION ALL select 2 UNION ALL select 3 UNION ALL select 4 UNION ALL select 5 UNION ALL select 6 UNION ALL select 7 UNION ALL select 8 UNION ALL select 9) as b CROSS JOIN
'(select 0 as c UNION ALL select 1 UNION ALL select 2 UNION ALL select 3 UNION ALL select 4 UNION ALL select 5 UNION ALL select 6 UNION ALL select 7 UNION ALL select 8 UNION ALL select 9) as c CROSS JOIN
'(select 0 as d UNION ALL select 1 UNION ALL select 2 UNION ALL select 3 UNION ALL select 4 UNION ALL select 5 UNION ALL select 6 UNION ALL select 7 UNION ALL select 8 UNION ALL select 9) as d) as x) as x1
'where date between @startdate and @enddate) as x for xml path (''),TYPE).value('.' ,'NVARCHAR(MAX)'),1,1,'')

'set @result = '

'select Type,[Activity Name]
',rtrim([Site Unit]) as [Site Unit]
',Unit,Scope
','+ @col +' FROM 

'--******************************************************************
'(---Electrical Cable Tray 
'select ''Scope'' as Daily,''U'' + rtrim(unit) + '' - '' + ''Cable Tray Erected'' as [Activity Name],
'rtrim(unit) as [Site Unit]
',sum([EL_Tray_Length]) as ''Pulled'',''Electrical'' as [Type],''LM'' as [Unit] FROM tblEleCableTray 
'where active=1
'group by unit
'--******************************************************************
') as X

'PIVOT 
'(sum (Pulled) for Daily in ('+ @col +',Scope)
')as X1
'order by [Site Unit],Type
''  
'execute(@result)
#End Region


#Region "Dynamic Pivot With Selection and Order"
'IF OBJECT_ID('tempdb..#t') IS NOT NULL
'DROP TABLE #t
'declare @cat nvarchar (50)
'select @cat='CIRE'
'create table #t
'(
'[Cat_Name] nvarchar (50)
')
'insert into #t values (@cat)



'declare @result as nvarchar(max)
'declare @col as nvarchar(max)
'select @col = STUFF((select  ',' + quotename(x.step_name)

'FROM (select distinct Step_Name,step_no FROM vItems where Cat_Name in (select cat_name from #t) ) as x order by Step_No

'for xml path ('') ,TYPE).value('.' ,'NVARCHAR(MAX)'),1,1,'') 


'set @result = 'SELECT Family_Name,[Cat_Name],[Catalog_Name], ' + @col + ' from 
'            (
'                select Family_Name,[Cat_Name],[Catalog_Name],Step_Percentage,step_name

'                from vItems
'				where Cat_Name in (select cat_name from #t)

'           ) x

'            pivot 
'            (
'                 sum(Step_Percentage)
'                for step_name in (' + @col + ')
'            ) p 
'		order by Family_Name,[Cat_Name],[Catalog_Name]
''
'execute(@result)

#End Region


'#Region "Shrink DB"
'USE [master]
'GO
'ALTER DATABASE P6DB SET RECOVERY SIMPLE WITH NO_WAIT
'GO
'use TRP6Audit
'DBCC SHRINKFILE (TRP6Audit, 1)
'DBCC SHRINKFILE('TRPCSDB_log', 0, TRUNCATEONLY)
'go
'#End Region


#Region "Temp Table"
'IF OBJECT_ID('tempdb..#Days') IS NOT NULL
'DROP TABLE #Days
'create table #DayName
'(
'DName nvarchar (50)
',DayID int
')
#End Region