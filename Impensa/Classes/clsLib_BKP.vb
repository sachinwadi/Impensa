#Region "References"
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Reflection
Imports System.Drawing.Drawing2D
Imports System.Drawing
Imports System.Text
Imports System.Data.OleDb
Imports System.IO
Imports System.Collections.Generic
Imports System.Globalization
Imports excel = Microsoft.Office.Interop.Excel
#End Region

Public Class clsLib
#Region "Enums"
    Public Enum En_SummaryType
        Monthly = 0
        Yearly = 1
        RunningTotals = 2
        Variance = 3
    End Enum
#End Region   

#Region "Variables"
    Private Shared _cmbListing As Object
    Private Shared _cmbSelectChart As Object
    Private Shared _cmbChartType As Object
    Private Shared _dtpTo As String
    Private Shared _dtpFrom As String
    Private Shared _SearchStr As String
    Private Shared _Categories As String
    Private Shared _IsDataAvailable As Boolean
    Private Shared _SortByAmount As Boolean
    Private Shared _IsLoginDetailsChanged As Boolean
    Private Shared _ThresholdMonth As Date
    Private Shared _ThresholdCurrentMonthIndex As Integer
    Private Shared _SummaryType As Integer
    Private Shared _ImpensaFont As Font
    Private Shared _ChkLBYearsItemsList As String
    Private Shared _ExportPDFProcessID As Int64
    Private Shared _ChartDataSetAmtTotal As List(Of String)
    Private Shared _CategoryColIndex As Int32
    Private Shared _CategoriesList As Dictionary(Of Int32, String)
    'Private Shared _LogFileProcessID As Int64
    Private Shared _ImportSucceessCnt As Int32
    Private Shared _ImportFailedCnt As Int32
    Private Shared _ImportSkippedCnt As Int32
    Private Shared _TotalImportCnt As Int32
    Private Shared _ImportSucceessAndFailCnt As Int32
    Private Shared _AssemblyLocation As String
    Private Shared _LogTimeStamp As String
#End Region

#Region "Properties"

    Public Shared Property EnableImport() As Boolean
        Get
            Return My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Impensa", "Enable Import", Nothing)
        End Get
        Set(ByVal value As Boolean)
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Impensa", "Enable Import", value)
        End Set
    End Property

    Public Shared Property LogTimeStamp() As String
        Get
            Return My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Impensa", "LogTimeStamp", Nothing)
        End Get
        Set(ByVal value As String)
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Impensa", "LogTimeStamp", value)
        End Set
    End Property

    Public Shared Property LastFailedCnt() As Int32
        Get
            Return My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Impensa", "LastFailedCnt", Nothing)
        End Get
        Set(ByVal value As Int32)
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Impensa", "LastFailedCnt", value)
        End Set
    End Property

    Public Shared Property AssemblyLocation() As String
        Get
            Return My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run", "Impensa", Nothing)
        End Get
        Set(ByVal value As String)
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Run", "Impensa", value)
        End Set
    End Property

    Public Shared Property ImportFileTimeStamp() As String
        Get
            Return My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Impensa", "ImportFileTimeStamp", Nothing)
        End Get
        Set(ByVal value As String)
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Impensa", "ImportFileTimeStamp", value)
        End Set
    End Property

    Public Shared Property TotalImportCnt() As Int32
        Get
            Return _TotalImportCnt
        End Get
        Set(ByVal value As Int32)
            _TotalImportCnt = value
        End Set
    End Property

    Public Shared Property ImportSucceessCnt() As Int32
        Get
            Return _ImportSucceessCnt
        End Get
        Set(ByVal value As Int32)
            _ImportSucceessCnt = value
        End Set
    End Property

    Public Shared Property ImportFailedCnt() As Int32
        Get
            Return _ImportFailedCnt
        End Get
        Set(ByVal value As Int32)
            _ImportFailedCnt = value
        End Set
    End Property

    Public Shared Property ImportSucceessAndFailCnt() As Int32
        Get
            Return _ImportSucceessAndFailCnt
        End Get
        Set(ByVal value As Int32)
            _ImportSucceessAndFailCnt = value
        End Set
    End Property

    Public Shared Property ImportSkippedCnt() As Int32
        Get
            Return _ImportSkippedCnt
        End Get
        Set(ByVal value As Int32)
            _ImportSkippedCnt = value
        End Set
    End Property

    'Public Shared Property LogFileProcessID() As Int64
    '    Get
    '        Return My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Impensa", "LogFileProcessID", Nothing)
    '    End Get
    '    Set(ByVal value As Int64)
    '        My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Impensa", "LogFileProcessID", value)
    '    End Set
    'End Property

    Public Shared Property CategoriesList() As Dictionary(Of Int32, String)
        Get
            Return _CategoriesList
        End Get
        Set(ByVal value As Dictionary(Of Int32, String))
            _CategoriesList = value
        End Set
    End Property

    Public Shared Property CategoryColIndex() As Int32
        Get
            Return _CategoryColIndex
        End Get
        Set(ByVal value As Int32)
            _CategoryColIndex = value
        End Set
    End Property

    Public Shared Property ChartDataSetAmtTotal() As List(Of String)
        Get
            Return _ChartDataSetAmtTotal
        End Get
        Set(ByVal value As List(Of String))
            _ChartDataSetAmtTotal = value
        End Set
    End Property

    Public Shared Property ExportPDFProcessID() As Int64
        Get
            Return My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Impensa", "ExportPDFProcessID", Nothing)
        End Get
        Set(ByVal value As Int64)
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Impensa", "ExportPDFProcessID", value)
        End Set
    End Property

    Public Shared Property ChkLBYearsItemsList() As String
        Get
            Return _ChkLBYearsItemsList
        End Get
        Set(ByVal value As String)
            _ChkLBYearsItemsList = value
        End Set
    End Property

    Public Shared Property ImpensaFont() As Font
        Get
            Return _ImpensaFont
        End Get
        Set(ByVal value As Font)
            _ImpensaFont = value
        End Set
    End Property

    Public Shared ReadOnly Property Title() As String
        Get
            Return Assembly.GetExecutingAssembly.GetCustomAttributes(GetType(AssemblyTitleAttribute), False)(0).Title
        End Get
    End Property

    Public Shared Property LastACRefreshDate() As Date
        Get
            Return My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Impensa", "LastACRefreshDate", Nothing)
        End Get
        Set(ByVal value As Date)
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Impensa", "LastACRefreshDate", value)
        End Set
    End Property

    Public Shared Property HighlightDetail() As String
        Get
            Return My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Impensa", "HighlightDetail", Nothing)
        End Get
        Set(ByVal value As String)
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Impensa", "HighlightDetail", value)
        End Set
    End Property

    Public Shared Property HighlightSummMonthly() As String
        Get
            Return My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Impensa", "HighlightSummMonthly", Nothing)
        End Get
        Set(ByVal value As String)
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Impensa", "HighlightSummMonthly", value)
        End Set
    End Property

    Public Shared Property HighlightSummYearly() As String
        Get
            Return My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Impensa", "HighlightSummYearly", Nothing)
        End Get
        Set(ByVal value As String)
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Impensa", "HighlightSummYearly", value)
        End Set
    End Property

    Public Shared Property RecordKeepingStartDate() As String
        Get
            Return My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Impensa", "Record Keeping Start Date", Nothing)
        End Get
        Set(ByVal value As String)
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Impensa", "Record Keeping Start Date", Format(CDate(value), "dd/MM/yyyy"))
        End Set
    End Property

    Public Shared Property CSVBackupPath() As String
        Get
            Return My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Impensa", "Backup Path", Nothing)
        End Get
        Set(ByVal value As String)
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Impensa", "Backup Path", value)
        End Set
    End Property

    Public Shared Property ConnStr() As String
        Get
            Return My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Impensa", "ConnStr", Nothing)
        End Get
        Set(ByVal value As String)
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Impensa", "ConnStr", value)
        End Set
    End Property

    Public Shared Property ShowReminder() As Boolean
        Get
            Return My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Impensa", "Show Reminder", Nothing)
        End Get
        Set(ByVal value As Boolean)
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Impensa", "Show Reminder", value)
        End Set
    End Property

    Public Shared Property ReminderText() As String
        Get
            Return My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Impensa", "Reminder Text", Nothing)
        End Get
        Set(ByVal value As String)
            My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Impensa", "Reminder Text", value)
        End Set
    End Property

    Public Shared Property ListingCombo() As Object
        Get
            Return _cmbListing
        End Get
        Set(ByVal value As Object)
            _cmbListing = value
        End Set
    End Property

    Public Shared Property SelectChartCombo() As Object
        Get
            Return _cmbSelectChart
        End Get
        Set(ByVal value As Object)
            _cmbSelectChart = value
        End Set
    End Property

    Public Shared Property ChartTypeCombo() As Object
        Get
            Return _cmbChartType
        End Get
        Set(ByVal value As Object)
            _cmbChartType = value
        End Set
    End Property

    Public Shared Property SearchStr() As String
        Get
            Return _SearchStr
        End Get
        Set(ByVal value As String)
            _SearchStr = value
        End Set
    End Property

    Public Shared Property dtpFrom() As String
        Get
            Return _dtpFrom
        End Get
        Set(ByVal value As String)
            _dtpFrom = value
        End Set
    End Property

    Public Shared Property dtpTo() As String
        Get
            Return _dtpTo
        End Get
        Set(ByVal value As String)
            _dtpTo = value
        End Set
    End Property

    Public Shared Property Categories() As String
        Get
            Return _Categories
        End Get
        Set(ByVal value As String)
            _Categories = value
        End Set
    End Property

    Public Shared Property IsDataAvailable() As Boolean
        Get
            Return _IsDataAvailable
        End Get
        Set(ByVal value As Boolean)
            _IsDataAvailable = value
        End Set
    End Property

    Public Shared Property SortByAmount() As Boolean
        Get
            Return _SortByAmount
        End Get
        Set(ByVal value As Boolean)
            _SortByAmount = value
        End Set
    End Property

    Public Shared Property IsLoginDetailsChanged() As Boolean
        Get
            Return _IsLoginDetailsChanged
        End Get
        Set(ByVal value As Boolean)
            _IsLoginDetailsChanged = value
        End Set
    End Property

    Public Shared Property ThresholdMonth() As Date
        Get
            Return _ThresholdMonth
        End Get
        Set(ByVal value As Date)
            _ThresholdMonth = value
        End Set
    End Property

    Public Shared Property ThresholdCurrentMonthIndex() As Integer
        Get
            Return _ThresholdCurrentMonthIndex
        End Get
        Set(ByVal value As Integer)
            _ThresholdCurrentMonthIndex = value
        End Set
    End Property

    Public Shared Property SummaryType() As En_SummaryType
        Get
            Return _SummaryType
        End Get
        Set(ByVal value As En_SummaryType)
            _SummaryType = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Shared Function GetConnection() As SqlConnection
        Dim Conn As SqlConnection = New SqlConnection(ConnStr)
        Conn.Open()
        Return Conn
    End Function

    Public Shared Sub ImpensaAlert(ByVal P_Text As String, ByVal P_Style As MsgBoxStyle)
        Call MsgBox(P_Text, P_Style, Title)
    End Sub

    Public Shared Function ImpensaActionAlert(ByVal P_Text As String, ByVal P_Style As MessageBoxButtons, ByVal P_Icon As MessageBoxIcon) As Windows.Forms.DialogResult
        Return MessageBox.Show(P_Text, Title, P_Style, P_Icon)
    End Function

    Public Shared Sub DataImport()
        Call ExcelDataImport()
    End Sub

    'Public Shared Sub CSVDataImport()
    '    Try
    '        Dim dt As New DataTable
    '        Dim OleDBAdapter As OleDbDataAdapter
    '        Dim DataFile = CSVBackupPath + "\Impensa.csv"

    '        If File.Exists(DataFile) Then

    '            Using OleDBConn = New OleDbConnection("Provider=Microsoft.Jet.OleDb.4.0;Data Source=" & CSVBackupPath & ";Extended Properties=""Text; HDR=YES; FMT=Delimited""")
    '                OleDBConn.Open()
    '                OleDBAdapter = New OleDbDataAdapter("SELECT * FROM Impensa.csv", OleDBConn)
    '                OleDBAdapter.Fill(dt)
    '                OleDBConn.Close()
    '            End Using

    '            If dt.Rows.Count > 0 Then
    '                Dim da As SqlDataAdapter
    '                Dim dcb As SqlCommandBuilder
    '                Dim FileText As String() = File.ReadAllLines(DataFile)

    '                TotalImportCnt = dt.Rows.Count

    '                dt.Columns.Add("hKey", GetType(Int64), Nothing)
    '                dt.Columns.Add("iCategory", GetType(Int32), Nothing)

    '                dt.Columns("hKey").SetOrdinal(0)
    '                dt.Columns("Date").SetOrdinal(1)
    '                dt.Columns("iCategory").SetOrdinal(2)
    '                dt.Columns("Amount").SetOrdinal(3)
    '                dt.Columns("Notes").SetOrdinal(4)

    '                If Not LogFileProcessID = Nothing Then
    '                    If Process.GetProcesses.Any(Function(x) x.Id = LogFileProcessID) Then
    '                        Process.GetProcessById(LogFileProcessID).Kill()
    '                        LogFileProcessID = Nothing
    '                        Threading.Thread.Sleep(100)
    '                    End If
    '                End If

    '                Using fs = New FileStream(Path.GetTempPath & "\ImpensaImport.log", FileMode.Append, FileAccess.Write, FileShare.Read)
    '                    Using Writer = New StreamWriter(fs)
    '                        LogTimeStamp = DateTime.Now.ToString
    '                        Writer.WriteLine(LogTimeStamp)
    '                        For Each dr As DataRow In dt.Rows

    '                            If Right(FileText(dt.Rows.IndexOf(dr) + 1), 1) = "," Then
    '                                FileText(dt.Rows.IndexOf(dr) + 1) = FileText(dt.Rows.IndexOf(dr) + 1).ToString.Substring(0, Len(FileText(dt.Rows.IndexOf(dr) + 1)) - 1)
    '                            End If

    '                            If Not IIf(dr("Skip") Is DBNull.Value, "N", dr("Skip")) = "Y" Then
    '                                Dim IsSuccess As Boolean = True
    '                                For Each dc As DataColumn In dt.Columns
    '                                    If dc.ColumnName = "Date" Then
    '                                        If Not Date.TryParse(dr(dc).ToString, Nothing) Then
    '                                            Writer.WriteLine("Record No. #" & dt.Rows.IndexOf(dr) + 1 & " Failed to import. Invalid date format.")
    '                                            IsSuccess = False
    '                                        End If
    '                                    ElseIf dc.ColumnName = "Category" Then
    '                                        If Not CategoriesList.ContainsValue(LCase(dr(dc))) Then
    '                                            Writer.WriteLine("Record No. #" & dt.Rows.IndexOf(dr) + 1 & " Failed to import. Invalid category.")
    '                                            IsSuccess = False
    '                                        Else
    '                                            dr("iCategory") = CategoriesList.FirstOrDefault(Function(x) x.Value = LCase(dr(dc))).Key
    '                                        End If
    '                                    ElseIf dc.ColumnName = "Amount" Then
    '                                        If Not Double.TryParse(dr(dc).ToString, Nothing) Then
    '                                            Writer.WriteLine("Record No. #" & dt.Rows.IndexOf(dr) + 1 & " Failed to import. Invalid amount format.")
    '                                            IsSuccess = False
    '                                        End If
    '                                    ElseIf dc.ColumnName = "Notes" Then
    '                                        If dr(dc).ToString.Length > 500 Then
    '                                            Writer.WriteLine("Record No. #" & dt.Rows.IndexOf(dr) + 1 & " Failed to import. Notes' length exceeds 500 characters.")
    '                                            IsSuccess = False
    '                                        End If
    '                                    End If
    '                                Next

    '                                If IsSuccess = True Then
    '                                    Writer.WriteLine("Record No. #" & dt.Rows.IndexOf(dr) + 1 & " successfully imported.")

    '                                    If dr("Notes") Is DBNull.Value Then
    '                                        FileText(dt.Rows.IndexOf(dr) + 1) = FileText(dt.Rows.IndexOf(dr) + 1) + ",,Y"
    '                                    Else
    '                                        FileText(dt.Rows.IndexOf(dr) + 1) = FileText(dt.Rows.IndexOf(dr) + 1) + ",Y"
    '                                    End If

    '                                    ImportSucceessCnt += 1
    '                                Else
    '                                    ImportFailedCnt += 1
    '                                    dr.Delete()
    '                                End If
    '                            Else
    '                                Writer.WriteLine("Record No. #" & dt.Rows.IndexOf(dr) + 1 & " skipped. Record was already imported.")
    '                                dr.Delete()
    '                                ImportSkippedCnt += 1
    '                            End If
    '                        Next

    '                        dt.AcceptChanges()
    '                        If dt.Rows.Count > 0 Then
    '                            File.WriteAllLines(DataFile, FileText)
    '                        End If
    '                        For Each dr As DataRow In dt.Rows
    '                            dr.SetAdded()
    '                        Next

    '                        Writer.WriteLine(New String("=", 50))

    '                        Using Connection = GetConnection()
    '                            da = New SqlDataAdapter("SELECT hKey, dtDate [Date], iCategory, dAmount [Amount], sNotes [Notes] FROM tbl_ExpenditureDet", Connection)
    '                            dcb = New SqlCommandBuilder(da)
    '                            da.Update(dt)
    '                        End Using

    '                        Call SaveNotes(dt)
    '                    End Using
    '                End Using
    '            End If
    '        End If
    '        ImportFileTimeStamp = File.GetLastWriteTime(DataFile)
    '        LastFailedCnt = ImportFailedCnt
    '    Catch ex As Exception
    '        ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
    '    End Try
    'End Sub

    Public Shared Sub ExcelDataImport()
        Dim DataFile = CSVBackupPath + "\Impensa.xlsm"
        Dim dt As New DataTable
        Dim OleDBAdapter As OleDbDataAdapter
        Dim ExcelApp As excel.Application = Nothing
        Dim ExcelWorkBook As excel.Workbook = Nothing
        Dim ExcelWorkSheet As excel.Worksheet = Nothing
        Dim Workbooks As excel.Workbooks = Nothing

        Try

            Threading.Thread.Sleep(1000)

            Using OleDBConn = New OleDbConnection("provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & DataFile & ";Extended Properties=""Excel 12.0 Macro; HDR=YES; IMEX=1""")
                OleDBConn.Open()
                OleDBAdapter = New OleDbDataAdapter("SELECT * FROM [Sheet1$A:D] WHERE (Date IS NOT NULL AND Category IS NOT NULL AND Amount IS NOT NULL) AND Date < DATE()", OleDBConn)
                OleDBAdapter.Fill(dt)
            End Using

            If dt.Rows.Count > 0 Then
                Dim da As SqlDataAdapter
                Dim dcb As SqlCommandBuilder
                Dim dv As New DataView

                TotalImportCnt = dt.Rows.Count

                dt.Columns.Add("hKey", GetType(Int64), Nothing)

                dt.Columns("hKey").SetOrdinal(0)
                dt.Columns("Date").SetOrdinal(1)
                dt.Columns("Category").SetOrdinal(2)
                dt.Columns("Amount").SetOrdinal(3)
                dt.Columns("Notes").SetOrdinal(4)

                For Each dr As DataRow In dt.Rows
                    dr.SetAdded()
                Next

                Using Connection = GetConnection()
                    da = New SqlDataAdapter("SELECT hKey, dtDate [Date], sCategory [Category], dAmount [Amount], sNotes [Notes] FROM Temp_ImportData", Connection)
                    dcb = New SqlCommandBuilder(da)
                    da.Update(dt)

                    dt.AcceptChanges()

                    dt = New DataTable
                    da = New SqlDataAdapter("EXECUTE sp_ImportData", Connection)
                    da.Fill(dt)
                End Using

                dv = dt.DefaultView
                dv.RowFilter = Nothing

                dv.RowFilter = "sImportComments LIKE '%fail%'"
                ImportFailedCnt = dv.ToTable.Rows.Count

                dv.RowFilter = "sImportComments LIKE '%skip%'"
                ImportSkippedCnt = dv.ToTable.Rows.Count

                dv.RowFilter = "sImportComments LIKE '%success%'"
                ImportSucceessCnt = dv.ToTable.Rows.Count

                ImportSucceessAndFailCnt = ImportSucceessCnt + ImportFailedCnt

                dt = dv.ToTable

                Call SaveNotes(dt) 'Add notes of successful records to AutoComplete Directory

                dv.RowFilter = Nothing
                dt = dv.ToTable

                ''''Start: Format Import Excel
                ExcelApp = New excel.Application
                Workbooks = ExcelApp.Workbooks
                ExcelWorkBook = Workbooks.Open(DataFile)
                ExcelWorkSheet = ExcelWorkBook.Sheets(1)

                ExcelWorkSheet.Unprotect("Hotmail@123")

                For Each dr As DataRow In dt.Rows
                    If InStr(LCase(dr("sImportComments")), "success") > 0 Then
                        ExcelWorkSheet.Range("F" & (dt.Rows.IndexOf(dr) + 2)).Value = "Success"
                        'ExcelWorkSheet.Range("A" & (dt.Rows.IndexOf(dr) + 2), "D" & (dt.Rows.IndexOf(dr) + 2)).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Green)
                    ElseIf InStr(LCase(dr("sImportComments")), "fail") > 0 Then
                        ExcelWorkSheet.Range("F" & (dt.Rows.IndexOf(dr) + 2)).Value = "Failed"
                        'ExcelWorkSheet.Range("A" & (dt.Rows.IndexOf(dr) + 2), "D" & (dt.Rows.IndexOf(dr) + 2)).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red)
                    ElseIf InStr(LCase(dr("sImportComments")), "skip") > 0 Then
                        ExcelWorkSheet.Range("F" & (dt.Rows.IndexOf(dr) + 2)).Value = "Skipped"
                        'ExcelWorkSheet.Range("A" & (dt.Rows.IndexOf(dr) + 2), "D" & (dt.Rows.IndexOf(dr) + 2)).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange)
                    End If

                    If Date.TryParse(dr("Date").ToString, Nothing) Then
                        ExcelWorkSheet.Range("G" & (dt.Rows.IndexOf(dr) + 2)).Value = Left(CDate(dr("Date")), 10)
                    Else
                        ExcelWorkSheet.Range("G" & (dt.Rows.IndexOf(dr) + 2)).Value = dr("Date")
                    End If

                    ExcelWorkSheet.Range("H" & (dt.Rows.IndexOf(dr) + 2)).Value = dr("Category")

                    If Double.TryParse(dr("Amount").ToString, Nothing) Then
                        ExcelWorkSheet.Range("I" & (dt.Rows.IndexOf(dr) + 2)).Value = Format(CDbl(dr("Amount")), "#,##0.00")
                    Else
                        ExcelWorkSheet.Range("I" & (dt.Rows.IndexOf(dr) + 2)).Value = dr("Amount")
                    End If

                    ExcelWorkSheet.Range("J" & (dt.Rows.IndexOf(dr) + 2)).Value = dr("Notes")
                Next

                ExcelWorkSheet.Protect("Hotmail@123")
                ExcelWorkBook.Save()
                ''''End: Format Import Excel
            End If

            ImportFileTimeStamp = File.GetLastWriteTime(DataFile)
            LastFailedCnt = ImportFailedCnt
        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        Finally
            ExcelWorkBook.Close()
            Workbooks.Close()
            ExcelApp.Quit()
            System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelWorkSheet)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelWorkBook)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(Workbooks)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelApp)
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Sub

    'Public Shared Sub GetNonObseletCategoriesList()
    '    Dim dir As New Dictionary(Of Int32, String)
    '    Dim da As SqlDataAdapter
    '    Dim dt As DataTable
    '    Try
    '        Using Connection = GetConnection()
    '            dt = New DataTable
    '            da = New SqlDataAdapter("SELECT hKey, LOWER(sCategory) sCategory FROM tbl_CategoryList WHERE ISNULL(IsObsolete,0) = 0", Connection)
    '            da.Fill(dt)
    '        End Using

    '        If dt.Rows.Count > 0 Then
    '            For Each dr As DataRow In dt.Rows
    '                dir.Add(dr("hKey"), dr("sCategory"))
    '            Next
    '        End If

    '    Catch ex As Exception
    '        ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
    '    End Try

    '    CategoriesList = dir
    'End Sub

    Public Shared Sub SaveNotes(ByVal dtGrid As DataTable)
        Dim StrCommand As New StringBuilder
        Dim Cmd As New SqlCommand

        Try
            Cmd = New SqlCommand
            If Not dtGrid Is Nothing Then
                Using Connection = GetConnection()
                    Cmd.Connection = Connection
                    For Each dr As DataRow In dtGrid.Rows
                        StrCommand.Length = 0
                        StrCommand.AppendFormat("MERGE tbl_Notes T1 USING (SELECT '{0}' sNotes)T2 ON T1.sNotes = T2.sNotes WHEN MATCHED THEN UPDATE SET T1.dtLastUsed = GETDATE() WHEN NOT MATCHED THEN INSERT(sNotes) VALUES (T2.sNotes);", dr("Notes").ToString.Replace("'", "''"))
                        Cmd.CommandText = StrCommand.ToString
                        Cmd.ExecuteNonQuery()
                    Next
                End Using
            End If
        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub '5
#End Region
End Class

Public Class MergedCell
    Inherits DataGridViewTextBoxCell

    Private _LeftColIndex As Int32 = 0
    Private _RightColIndex As Int32 = 0

    Public Property LeftColIndex() As Int32
        Get
            Return _LeftColIndex
        End Get
        Set(ByVal value As Int32)
            _LeftColIndex = value
        End Set
    End Property

    Public Property RightColIndex() As Int32
        Get
            Return _RightColIndex
        End Get
        Set(ByVal value As Int32)
            _RightColIndex = value
        End Set
    End Property

    Protected Overrides Sub Paint(ByVal graphics As Graphics, ByVal clipBounds As Rectangle, ByVal cellBounds As Rectangle, _
                                  ByVal rowIndex As Int32, ByVal cellState As DataGridViewElementStates, ByVal value As Object, _
                                  ByVal formattedValue As Object, ByVal errorText As String, ByVal cellStyle As DataGridViewCellStyle, _
                                  ByVal advancedBorderStyle As DataGridViewAdvancedBorderStyle, ByVal paintParts As DataGridViewPaintParts)

        Try
            Dim MergeIndex As Int32 = ColumnIndex - _LeftColIndex
            Dim Width As Int32 = 0
            Dim WidthLeft As Int32 = 0
            Dim strText As String


            Dim pen As New Pen(Brushes.Black)

            'Draw the background
            graphics.FillRectangle(New SolidBrush(Color.FromArgb(235, 241, 222)), cellBounds)

            'Draw the separator for rows
            graphics.DrawLine(New Pen(New SolidBrush(SystemColors.ActiveBorder)), cellBounds.Left, cellBounds.Bottom - 1, cellBounds.Right, cellBounds.Bottom - 1)


            ''Draw the right vertical line for the cell
            'If ColumnIndex = _RightColIndex Then
            '    graphics.DrawLine(New Pen(New SolidBrush(SystemColors.ControlDark)), cellBounds.Right - 1, cellBounds.Top, cellBounds.Right - 1, cellBounds.Bottom)
            'End If

            'Draw the text
            Dim rectDest As RectangleF = RectangleF.Empty
            Dim sf As StringFormat = New StringFormat()

            sf.Alignment = StringAlignment.Center
            sf.LineAlignment = StringAlignment.Center
            sf.Trimming = StringTrimming.EllipsisCharacter

            'Determine the total width of the merged cell
            For i As Int32 = _LeftColIndex To _RightColIndex
                Width += Me.OwningRow.Cells(i).Size.Width
            Next

            'Determine the width before the current cell.
            For i As Int32 = _LeftColIndex To ColumnIndex - 1
                WidthLeft += Me.OwningRow.Cells(i).Size.Width
            Next

            'Retrieve the text to be displayed
            strText = Me.OwningRow.Cells(_LeftColIndex).Value.ToString()
            rectDest = New RectangleF(cellBounds.Left - WidthLeft, cellBounds.Top, Width, cellBounds.Height)
            graphics.DrawString(strText, New Font("Microsoft Sans Serif", 14, FontStyle.Regular), Brushes.Green, rectDest, sf)

        Catch ex As Exception
            clsLib.ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub
End Class
