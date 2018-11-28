#Region "References"
Imports System.Data.SqlClient
Imports System.Windows.Forms.DataVisualization.Charting
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.IO
Imports System.Globalization
Imports Impensa.clsLib
Imports iTextSharp.text
Imports iTextSharp.text.pdf
#End Region

Public Class frmMain
#Region "Enums"
    Private Enum En_CSVCreationFrequency
        Monthly = 0
        Annually = 1
        Adhoc = 2
    End Enum

    Private Enum En_ToDate
        MTD = 0
        YTD = 1
    End Enum

    Private Enum En_Period
        CurrentMonth = 0
        CurrentYear = 1
        BookStart = 2
    End Enum

    Private Enum En_BudgetBuckets
        OverBudgetCats = 1
        AtParCats = 2
        UnderBudgetCats = 3
    End Enum
#End Region
    
#Region "Variables"
    Private da As SqlDataAdapter
    Private dt As DataTable
    Private Cmd As SqlCommand
    Private StrLstCategories As String = ""
    Private dtp As DateTimePicker
    Private WithEvents cmbMonth As ComboBox
    Private StrClosedYrs As String = ""
    Private SearchStr As String = ""
    Private DefaultSearchStr As String = "Multiple Keywords Seperated By Comma"
    Private Writer As StreamWriter
    Private Reader As StreamReader
    Private AutoList As New AutoCompleteStringCollection
    Private Delegate Sub Del_CloseSpash()
    Private dtDetailGrid As DataTable
    Private WithEvents DgvTextBox As DataGridViewTextBoxEditingControl
    Private WithEvents DgvComboBox As DataGridViewComboBoxEditingControl
    Private Dgv As DataGridView
    Private ShowSearchAlert As Boolean = True
    Private LastTabIndex As Integer = 0
    Private DoNotChkRowAdded As Boolean = True
    Private CallSearchFunction As Boolean = False
    Private HighlightDetail_Orig As String
    Private HighlightSummMonthly_Orig As String
    Private HighlightSummYearly_Orig As String
    Private RunTotalMonth As String = ""
    Private RunTotRowIndex As Int32 = 0
    Private CmbMonthCurrentIndex As Int32 = -1
    Private ChartColorScheme As Integer = 0
    Private Rnd As New Random
    Private tt As ToolTip
    Private bClearCmbPeriod As Boolean = False
    Private Highlighter As System.Drawing.Color = Color.LightBlue
    Private dtBudget As DataTable
    Private dtThresholdData As DataTable
    Private dtShowUntil As Date
    Private FirstTimeLoadInBkg As Boolean
#End Region

#Region "Methods"

#Region "Expenditure Details"

    Private Sub StartImportService()
        If Not CSVBackupPath = Nothing AndAlso EnableImport Then
            If File.Exists(CSVBackupPath + "\Impensa.xlsm") Then
                If Not BGWorker.IsBusy Then
                    BGWorker.RunWorkerAsync()
                End If
            Else
                ImpensaAlert("File Impensa.xlsm could not be found at the location mentioned in the settings. An attempt would be made to resume Impensa data import service on application restart.", MsgBoxStyle.Exclamation)
                EnableImport = False
            End If
        End If
    End Sub

    Private Function GetComboBoxColumn_Category() As DataGridViewComboBoxColumn
        Dim ColCombo As New DataGridViewComboBoxColumn
        Try
            Using Connection = GetConnection()
                da = New SqlDataAdapter("SELECT hKey 'iCategory', CASE IsObsolete WHEN 0 THEN sCategory ELSE sCategory + '*' END sCategory FROM tbl_CategoryList ORDER BY sCategory", Connection)
                dt = New DataTable
                da.Fill(dt)
            End Using
            ColCombo.DataPropertyName = "iCategory"
            ColCombo.HeaderText = "Category"
            ColCombo.Name = "iCategory"
            ColCombo.DataSource = dt
            ColCombo.ValueMember = "iCategory"
            ColCombo.DisplayMember = "sCategory"
            ColCombo.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing
        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try

        Return ColCombo
    End Function '19

    Private Sub PopulateExpenditureDetailGrid()
        Dim strSQL As String = ""
        Dim dc_Category As DataGridViewComboBoxColumn
        Dim dc_DelChk As New DataGridViewCheckBoxColumn

        Try
            DataGridExpDet.ReadOnly = False
            DataGridExpDet.DataSource = Nothing
            DoNotChkRowAdded = True

            Label15.Text = "Loading Details..."
            Application.DoEvents()

            StrClosedYrs = BuildOpenOrClosedYrsStr(1) 'List Of Closed Years
            dc_Category = GetComboBoxColumn_Category()
            DataGridExpDet.Columns.Add(dc_Category)

            dc_DelChk.Name = "bDelete"
            dc_DelChk.HeaderText = "Del"
            dc_DelChk.DataPropertyName = "bDelete"
            dc_DelChk.AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
            DataGridExpDet.Columns.Add(dc_DelChk)

            Using Connection = GetConnection()
                dtDetailGrid = New DataTable
                strSQL = "EXEC sp_GetExpenditureDetails '" & dtpFrom & "', '" & dtpTo & "', '" & Categories & "'"
                da = New SqlDataAdapter(strSQL, Connection)
                da.Fill(dtDetailGrid)
            End Using

            DataGridExpDet.DataSource = dtDetailGrid
            DataGridExpDet.DefaultCellStyle.WrapMode = DataGridViewTriState.True
            DataGridExpDet.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
            DataGridExpDet.AllowUserToAddRows = False
            DataGridExpDet.AllowUserToResizeRows = False
            DataGridExpDet.RowHeadersVisible = False

            DataGridExpDet.Columns("hKey").Visible = False
            DataGridExpDet.Columns("IsReadOnly").Visible = False
            DataGridExpDet.Columns("IsDummy").Visible = False
            DataGridExpDet.Columns("sCategory").Visible = False
            DataGridExpDet.Columns("IsDummyRowAdded").Visible = False

            DataGridExpDet.Columns("bDelete").DisplayIndex = 0
            DataGridExpDet.Columns("iCategory").DisplayIndex = 3
            DataGridExpDet.Columns("Notes").DisplayIndex = 5

            DataGridExpDet.Columns("Amount").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            DataGridExpDet.Columns("Amount").DefaultCellStyle.Format = "#,##0.00"
            DataGridExpDet.Columns("Amount").HeaderText = "Amount (Rs.)"
            DataGridExpDet.Columns("Amount").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight

            DataGridExpDet.Columns("Date").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            DataGridExpDet.Columns("Date").DefaultCellStyle.Format = "dd/MM/yyyy"

            DataGridExpDet.Columns("Notes").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft

            tslblRecdCnt.Text = "Total Records Displayed: #" & dtDetailGrid.Select("hKey IS NOT NULL").ToArray.Count
            tslblSeperator2.Text = "||"
            tslblGridTotal.Text = "Grid Total: Rs. 0.00"
            tslblSeperator1.Text = "||"

            Call FormatDetailGrid()

            If Not dtp Is Nothing Then
                dtp.MinDate = dtpRecdKeeping.Value.Date
                dtp.MaxDate = DateTime.Today
            End If

            DoNotChkRowAdded = False
            chkShowAllDet.Enabled = False
        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try

    End Sub '7

    Private Sub FormatDetailGrid()
        Dim TextBoxCell As DataGridViewTextBoxCell

        For i As Integer = 0 To DataGridExpDet.Columns.Count - 1
            DataGridExpDet.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
            DataGridExpDet.Columns(i).HeaderCell.Style.Font = New System.Drawing.Font(ImpensaFont, FontStyle.Bold)
            DataGridExpDet.Columns(i).DefaultCellStyle.Font = ImpensaFont
        Next

        For Each dr As DataGridViewRow In DataGridExpDet.Rows
            If Not dr.Cells("Date").Value Is DBNull.Value Then
                If (dr.Cells("Date").Value = "TOTAL" Or dr.Cells("Date").Value = "GRAND TOTAL") Then
                    TextBoxCell = New DataGridViewTextBoxCell()
                    DataGridExpDet("iCategory", dr.Index) = TextBoxCell
                    TextBoxCell = Nothing
                    TextBoxCell = New DataGridViewTextBoxCell()
                    DataGridExpDet("Notes", dr.Index) = TextBoxCell
                    TextBoxCell = Nothing
                    TextBoxCell = New DataGridViewTextBoxCell()
                    DataGridExpDet("bDelete", dr.Index) = TextBoxCell
                    dr.DefaultCellStyle.BackColor = Color.Yellow
                    dr.DefaultCellStyle.Font = New System.Drawing.Font(ImpensaFont, FontStyle.Bold)
                    dr.ReadOnly = True

                    If dr.Cells("Date").Value = "GRAND TOTAL" Then
                        dr.DefaultCellStyle.ForeColor = Color.Red
                        tslblGridTotal.Text = "Grid Total: Rs. " & Format(dr.Cells("Amount").Value, "#,##0.00")
                    End If
                End If

                If Not IsDBNull(dr.Cells("Amount").Value) Then
                    If Not (dr.Cells("Date").Value = "TOTAL" OrElse dr.Cells("Date").Value = "GRAND TOTAL") Then
                        dr.Cells("Amount").Style.BackColor = Nothing
                        If dr.Cells("Amount").Value >= HighlightDetail Then
                            dr.Cells("Amount").Style.BackColor = Highlighter
                        End If
                    End If
                    If dr.Cells("Amount").Value < 0 Then
                        dr.DefaultCellStyle.ForeColor = Color.Gray
                    End If
                End If

            End If

            If Convert.ToBoolean(dr.Cells("IsReadOnly").Value) Then
                dr.ReadOnly = True
            End If

            If Convert.ToBoolean(dr.Cells("IsDummy").Value) Then
                dr.DefaultCellStyle.BackColor = Color.LightCyan
            End If
        Next
    End Sub '34

    Private Sub SaveExpenses()
        Dim dtGrid As New DataTable
        Dim dc As SqlCommandBuilder
        Dim lst As New List(Of Integer)
        Dim RowCounter As Integer = 0

        Try
            dtGrid = DirectCast(DataGridExpDet.DataSource, DataTable).GetChanges
            If Not dtGrid Is Nothing Then
                For Each dr As DataRow In dtGrid.Rows
                    If dr.RowState = DataRowState.Modified And dr("hKey") Is DBNull.Value Then
                        dr.AcceptChanges()
                        dr.SetAdded()
                    End If

                    If dr("hKey") Is DBNull.Value And dr("Amount") Is DBNull.Value Then
                        lst.Add(RowCounter)
                    End If

                    RowCounter += 1

                    If dr("Notes").ToString.Trim.EndsWith("#") Then
                        dr("Notes") = dr("Notes").Substring(0, dr("Notes").LastIndexOf("#"))
                    End If
                Next

                For Each lstItem As Integer In lst
                    dtGrid.Rows.RemoveAt(lstItem)
                Next

                Using Connection = GetConnection()
                    da = New SqlDataAdapter("SELECT hKey, dtDate [Date], iCategory, dAmount [Amount], sNotes [Notes] FROM tbl_ExpenditureDet", Connection)
                    dc = New SqlCommandBuilder(da)


                    For Each dr As DataRow In dtGrid.Rows
                        If Not (IsDBNull(dr.Item("bDelete"))) Then
                            If (Convert.ToBoolean(dr.Item("bDelete"))) Then
                                dr.Delete()
                            End If
                        End If
                    Next

                    da.Update(dtGrid)
                End Using
            End If

            Call RefreshGrids()

            If Not String.IsNullOrEmpty(CSVBackupPath) Then
                Call CreateMonthlyCSV()
                Call CreateAdhocCSV()
            End If

            ImpensaAlert("Your Changes Have Been Saved.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly)
        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub '1

    Private Sub PopulateSearchResults()
        Dim strSQL As String = ""
        Dim dc_Category As DataGridViewComboBoxColumn
        Dim dc_DelChk As New DataGridViewCheckBoxColumn

        Try
            DoNotChkRowAdded = True
            DataGridExpDet.DataSource = Nothing
            DataGridExpDet.Columns.Clear()

            dc_Category = GetComboBoxColumn_Category()
            DataGridExpDet.Columns.Add(dc_Category)

            dc_DelChk.Name = "bDelete"
            dc_DelChk.HeaderText = "Del"
            dc_DelChk.DataPropertyName = "bDelete"
            dc_DelChk.AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
            DataGridExpDet.Columns.Add(dc_DelChk)

            Label15.Text = "Loading Search Results..."
            Application.DoEvents()

            Using Connection = GetConnection()
                dtDetailGrid = New DataTable
                strSQL = "EXEC sp_SearchExpenses '" & dtpFrom & "', '" & dtpTo & "', '" & Categories & "', '" & SearchStr & "'"
                da = New SqlDataAdapter(strSQL, Connection)
                da.Fill(dtDetailGrid)
            End Using

            DataGridExpDet.DataSource = dtDetailGrid
            tslblRecdCnt.Text = "Total Records Displayed: #" & dtDetailGrid.Select("hKey IS NOT NULL").ToArray.Count
            tslblGridTotal.Text = "Grid Total: Rs. 0.00"

            DataGridExpDet.DefaultCellStyle.WrapMode = DataGridViewTriState.True
            DataGridExpDet.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            DataGridExpDet.AllowUserToAddRows = False
            DataGridExpDet.AllowUserToResizeRows = False

            DataGridExpDet.Columns("Sort").Visible = False
            DataGridExpDet.Columns("hKey").Visible = False
            DataGridExpDet.Columns("IsDummy").Visible = False
            DataGridExpDet.Columns("IsDummyRowAdded").Visible = False
            DataGridExpDet.Columns("IsReadOnly").Visible = False

            DataGridExpDet.Columns("bDelete").DisplayIndex = 0
            DataGridExpDet.Columns("iCategory").DisplayIndex = 4
            DataGridExpDet.Columns("Amount").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            DataGridExpDet.Columns("Amount").DefaultCellStyle.Format = "#,##0.00"
            DataGridExpDet.Columns("Date").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft
            DataGridExpDet.Columns("Date").DefaultCellStyle.Format = "dd/MM/yyyy"
            DataGridExpDet.Columns("Amount").HeaderText = "Amount (Rs.)"
            DataGridExpDet.Columns("Amount").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
            DataGridExpDet.Columns("Notes").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleLeft

            Call FormatSearchResults()

            DoNotChkRowAdded = False
        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub '15

    Private Sub FormatSearchResults()
        Dim TextBoxCell As DataGridViewTextBoxCell

        For i As Integer = 0 To DataGridExpDet.Columns.Count - 1
            DataGridExpDet.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
            DataGridExpDet.Columns(i).HeaderCell.Style.Font = New System.Drawing.Font(ImpensaFont, FontStyle.Bold)
            DataGridExpDet.Columns(i).DefaultCellStyle.Font = ImpensaFont
        Next

        For Each dr As DataGridViewRow In DataGridExpDet.Rows
            If dr.Cells("Sort").Value = 1 Then
                dr.DefaultCellStyle.Font = New System.Drawing.Font(ImpensaFont, FontStyle.Bold)
                dr.DefaultCellStyle.BackColor = Color.LightPink
                dr.DefaultCellStyle.ForeColor = Color.Blue
            ElseIf dr.Cells("Sort").Value = 3 Then
                dr.DefaultCellStyle.Font = New System.Drawing.Font(ImpensaFont, FontStyle.Bold)
                dr.DefaultCellStyle.BackColor = Color.Yellow
                If dr.Cells("Notes").Value Is DBNull.Value Then
                    dr.DefaultCellStyle.ForeColor = Color.Red
                    tslblGridTotal.Text = "Grid Total: Rs. " & Format(dr.Cells("Amount").Value, "#,##0.00")
                End If
            End If

            If dr.Cells("hKey").Value Is DBNull.Value Then
                TextBoxCell = New DataGridViewTextBoxCell()
                DataGridExpDet("iCategory", dr.Index) = TextBoxCell
                TextBoxCell = Nothing
                TextBoxCell = New DataGridViewTextBoxCell()
                DataGridExpDet("bDelete", dr.Index) = TextBoxCell
            End If

            If Convert.ToBoolean(dr.Cells("IsReadOnly").Value) Then
                dr.ReadOnly = True
            End If

            If Not IsDBNull(dr.Cells("Amount").Value) Then
                If Not (dr.Cells("Date").Value = "TOTAL" OrElse dr.Cells("Date").Value = "GRAND TOTAL") Then
                    dr.Cells("Amount").Style.BackColor = Nothing
                    If dr.Cells("Amount").Value >= HighlightDetail Then
                        dr.Cells("Amount").Style.BackColor = Highlighter
                    End If
                End If
            End If
        Next
    End Sub '35

#Region "Notes Methods"

    Private Sub SetNotesFieldAutoCompleteSource()
        Dim Reader As SqlDataReader
        Try
            Using Connection = GetConnection()
                Cmd = New SqlCommand("SELECT sNotes [Notes] FROM tbl_Notes", Connection)
                Reader = Cmd.ExecuteReader

                While Reader.Read
                    AutoList.Add(Reader.GetValue(0))
                End While

                Reader.Close()
            End Using

        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub '41

    Private Sub RefreshAutoCompleteDictionary()
        Try
            TabControl1.Enabled = False
            Panel5.BringToFront()
            Panel5.Visible = True
            Label15.Text = "Rebuilding AutoComplete Dictionary..."
            Application.DoEvents()

            Using Connection = GetConnection()
                Cmd = New SqlCommand("DELETE N FROM tbl_Notes N LEFT JOIN(SELECT MAX(dtDate) dtDate, sNotes From tbl_ExpenditureDet GROUP BY sNotes)X ON N.sNotes = X.sNotes WHERE DATEADD(YY,1,N.dtLastUsed) <= CONVERT(DATE, GETDATE())", Connection)
                Cmd.ExecuteNonQuery()
            End Using

            LastACRefreshDate = Today.Date
        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        Finally
            Panel5.SendToBack()
            Panel5.Visible = False
            TabControl1.Enabled = True
        End Try
    End Sub '44

#End Region
#End Region

#Region "Expenditure Summary"

    Private Sub PopulateExpenditureSummaryGrid()
        Dim strSQL As String = ""
        Dim procName As String = ""
        Dim dtGridSummary As DataTable
        Try
            DataGridExpSumm.DataSource = Nothing
            DataGridExpSumm.Controls.Clear()
            Label15.Text = "Loading Summary..."
            Application.DoEvents()

            If SummaryType = En_SummaryType.Monthly Then
                procName = "sp_GetExpenditureSummary_Monthly"
            ElseIf SummaryType = En_SummaryType.Yearly OrElse SummaryType = En_SummaryType.Variance Then
                procName = "sp_GetExpenditureSummary_Yearly"
            ElseIf SummaryType = En_SummaryType.RunningTotals Then
                procName = "sp_GetExpenditureSummary_RunningTotals"
            End If

            Using Connection = GetConnection()
                If SummaryType = En_SummaryType.RunningTotals Then
                    strSQL = "Execute " & procName & " '" & dtpFrom & "', '" & dtpTo & "', " & cmbCatListRunTot.SelectedItem.key & ", " & Month(Today) & ", '" & SearchStr & "', 0"
                ElseIf SummaryType = En_SummaryType.Variance Then
                    strSQL = "Execute " & procName & " '" & dtpFrom & "', '" & dtpTo & "', '" & Categories & "', '" & SearchStr & "', '" & ChkLBYearsItemsList & "'"
                Else
                    strSQL = "Execute " & procName & " '" & dtpFrom & "', '" & dtpTo & "', '" & Categories & "', '" & SearchStr & "'"
                End If
                dtGridSummary = New DataTable
                da = New SqlDataAdapter(strSQL, Connection)
                da.Fill(dtGridSummary)
            End Using

            If dtGridSummary.Rows.Count > 0 Then

                If TabControl1.SelectedTab.Name = "TabExpSumm" Then
                    tslblRecdCnt.Text = "Total Records Displayed: #" & dtGridSummary.Select("Sort = 1").ToArray.Count
                End If

                If SummaryType = En_SummaryType.RunningTotals Then
                    Using Connection = GetConnection()
                        da = New SqlDataAdapter("SELECT Y.iMonth 'iCategory', Datename(M, CONVERT(DATE, '1900-' + CONVERT(VARCHAR, Y.iMonth) + '-01')) 'Category' FROM (SELECT TOP(12) ROW_NUMBER() OVER(ORDER BY NAME) iMonth FROM sys.objects)Y", Connection)
                        dt = New DataTable
                        da.Fill(dt)
                    End Using

                    cmbMonth = New ComboBox
                    cmbMonth.DataSource = dt
                    cmbMonth.ValueMember = "iCategory"
                    cmbMonth.DisplayMember = "Category"
                    cmbMonth.DropDownStyle = ComboBoxStyle.DropDownList
                    cmbMonth.ForeColor = Color.FromArgb(31, 73, 125)
                    cmbMonth.BackColor = Color.FromArgb(197, 217, 241)
                    cmbMonth.Font = New System.Drawing.Font(ImpensaFont, FontStyle.Bold)
                    cmbMonth.Visible = False
                    DataGridExpSumm.Controls.Add(cmbMonth)
                End If

                DataGridExpSumm.DataSource = dtGridSummary
                DataGridExpSumm.AllowUserToAddRows = False
                DataGridExpSumm.AllowUserToResizeRows = False
                DataGridExpSumm.AutoResizeColumns()
                DataGridExpSumm.DefaultCellStyle.WrapMode = DataGridViewTriState.True

                If dtGridSummary.Columns.Count > 10 Then
                    DataGridExpSumm.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells
                Else
                    DataGridExpSumm.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
                End If

                CategoryColIndex = DataGridExpSumm.Columns("Category").Index

                Dim ColTotal As Double = 0

                For Each dr As DataGridViewRow In DataGridExpSumm.Rows
                    ColTotal = 0

                    If SummaryType = En_SummaryType.Variance Then
                        ColTotal = Math.Abs(dr.Cells(CategoryColIndex + 1).Value - dr.Cells(CategoryColIndex + 2).Value)
                    Else
                        For i = (CategoryColIndex + 1) To DataGridExpSumm.Columns.Count - 2
                            If Not dr.Cells(i).Value Is DBNull.Value Then
                                ColTotal += dr.Cells(i).Value
                            End If
                        Next
                    End If


                    If Not dr.Cells("Sort").Value = 3 Then
                        dr.Cells("TOTAL").Value = ColTotal
                    Else
                        dr.Cells("TOTAL").Value = DBNull.Value
                    End If

                    If dr.Cells("Sort").Value = 3 Then
                        RunTotRowIndex = dr.Index
                    End If

                    If SummaryType = En_SummaryType.RunningTotals AndAlso dr.Cells("Sort").Value = 4 Then
                        RunTotalMonth = dr.Cells("Category").Value
                    End If
                Next

                If SummaryType = En_SummaryType.RunningTotals Then
                    Dim StartMergeColIndex As Int32 = 2
                    Dim MC As MergedCell

                    DataGridExpSumm.Columns("Category").HeaderText = "Month"

                    For i As Int32 = StartMergeColIndex To DataGridExpSumm.Columns.Count - 1
                        DataGridExpSumm.Rows(RunTotRowIndex).Cells(i) = New MergedCell()
                        MC = CType(DataGridExpSumm.Rows(RunTotRowIndex).Cells(i), MergedCell)
                        MC.LeftColIndex = StartMergeColIndex
                        MC.RightColIndex = (DataGridExpSumm.Columns.Count - 1)
                    Next

                    DataGridExpSumm.Rows(RunTotRowIndex).Cells(StartMergeColIndex).Value = "RUNNING TOTAL"
                End If

                If DataGridExpSumm.Columns.Count > 0 Then
                    If SummaryType = En_SummaryType.Variance Then
                        DataGridExpSumm.Columns("TOTAL").HeaderText = "Variance"
                    Else
                        DataGridExpSumm.Columns("TOTAL").HeaderText = "Total"
                    End If
                    DataGridExpSumm.Columns("Sort").Visible = False
                    DataGridExpSumm.Columns("iCategory").Visible = False
                End If
            Else
                If TabControl1.SelectedTab.Name = "TabExpSumm" Then
                    tslblRecdCnt.Text = "Total Records Displayed: #0"
                End If
            End If

        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub '8

    Private Sub FormatSummaryGrid()
        Try
            If DataGridExpSumm.Rows.Count > 0 Then

                Label15.Text = "Styling Summary Grid..."
                Application.DoEvents()

                If SummaryType = En_SummaryType.Variance Then
                    For Each dr As DataGridViewRow In DataGridExpSumm.Rows
                        dr.Cells("TOTAL").Style.BackColor = Color.LightYellow

                        If dr.Cells(CategoryColIndex + 1).Value > dr.Cells(CategoryColIndex + 2).Value Then
                            dr.Cells("TOTAL").Style.ForeColor = Color.Green
                        Else
                            dr.Cells("TOTAL").Style.ForeColor = Color.Red
                        End If

                        dr.Cells("TOTAL").Style.Font = New System.Drawing.Font(ImpensaFont, FontStyle.Bold)

                        If Not dr.Cells("Category").Value Is DBNull.Value Then
                            If (dr.Cells("Category").Value = "TOTAL") Then
                                dr.DefaultCellStyle.BackColor = Color.Yellow

                                If dr.Cells(CategoryColIndex + 1).Value > dr.Cells(CategoryColIndex + 2).Value Then
                                    dr.DefaultCellStyle.ForeColor = Color.Green
                                    dr.Cells("TOTAL").Style.ForeColor = Color.Green
                                Else
                                    dr.DefaultCellStyle.ForeColor = Color.Red
                                    dr.Cells("TOTAL").Style.ForeColor = Color.Red
                                End If

                                dr.DefaultCellStyle.Font = New System.Drawing.Font(ImpensaFont, FontStyle.Bold)
                                dr.Cells("TOTAL").Style.BackColor = Color.Yellow
                            End If
                        End If
                    Next
                Else
                    DataGridExpSumm.Columns("TOTAL").HeaderCell.Style.ForeColor = Color.Purple

                    For Each dr As DataGridViewRow In DataGridExpSumm.Rows
                        dr.Cells("TOTAL").Style.BackColor = Color.LightYellow
                        dr.Cells("TOTAL").Style.ForeColor = Color.Purple
                        dr.Cells("TOTAL").Style.Font = New System.Drawing.Font(ImpensaFont, FontStyle.Bold)

                        If Not dr.Cells("Category").Value Is DBNull.Value Then
                            If (dr.Cells("Category").Value = "TOTAL") Then
                                dr.DefaultCellStyle.BackColor = Color.Yellow
                                dr.DefaultCellStyle.ForeColor = Color.Red
                                dr.DefaultCellStyle.Font = New System.Drawing.Font(ImpensaFont, FontStyle.Bold)

                                dr.Cells("TOTAL").Style.ForeColor = Color.Red
                                dr.Cells("TOTAL").Style.BackColor = Color.Yellow
                            End If
                        End If

                        If dr.Cells("Sort").Value = 4 Then
                            dr.DefaultCellStyle.ForeColor = Color.FromArgb(31, 73, 125)
                            dr.DefaultCellStyle.BackColor = Color.FromArgb(197, 217, 241)
                            dr.DefaultCellStyle.Font = New System.Drawing.Font(ImpensaFont, FontStyle.Bold)
                        End If
                        If dr.Cells("Sort").Value = 1 Then
                            If dr.Cells("Category").Value = RunTotalMonth Then
                                dr.DefaultCellStyle.ForeColor = Color.Brown
                                dr.DefaultCellStyle.Font = New System.Drawing.Font(ImpensaFont, FontStyle.Bold)
                            End If
                        End If
                    Next
                End If

                If DataGridExpSumm.Columns.Count > 0 Then
                    DataGridExpSumm.Columns("Category").Frozen = True
                End If

                Dim RowCnt As Int32 = IIf(SummaryType = En_SummaryType.RunningTotals, DataGridExpSumm.Rows.Count - 4, DataGridExpSumm.Rows.Count - 2)

                For i As Integer = 0 To RowCnt
                    For j As Integer = (CategoryColIndex + 1) To DataGridExpSumm.Columns.Count - 2
                        DataGridExpSumm.Rows(i).Cells(j).Style.ForeColor = Nothing
                        If Not DataGridExpSumm.Rows(i).Cells(j).Value Is DBNull.Value Then
                            DataGridExpSumm.Rows(i).Cells(j).Style.BackColor = Nothing
                            If DataGridExpSumm.Rows(i).Cells(j).Value > IIf(SummaryType = En_SummaryType.Monthly OrElse SummaryType = En_SummaryType.RunningTotals, HighlightSummMonthly, HighlightSummYearly) Then
                                DataGridExpSumm.Rows(i).Cells(j).Style.BackColor = Highlighter
                            ElseIf DataGridExpSumm.Rows(i).Cells(j).Value = 0 Then
                                DataGridExpSumm.Rows(i).Cells(j).Style.ForeColor = Color.Gray
                            End If
                        End If
                    Next
                Next

                If Not SummaryType = En_SummaryType.RunningTotals Then
                    Call GetBudgets()

                    If dtBudget.Rows.Count > 0 Then
                        Call CheckSummaryActualVsBudget()
                    End If
                End If
            End If

        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        Finally
            btnSave.Enabled = True
        End Try
    End Sub '33

    Private Sub CheckSummaryActualVsBudget()
        Try
            Dim CategoryColIndex As Integer = DataGridExpSumm.Columns("Category").Index
            Dim TotalColIndex As Integer = DataGridExpSumm.Columns("Total").Index
            Dim ApplyColorFormatting As Boolean

            If dtBudget.Rows.Count > 0 Then
                For ColId As Integer = CategoryColIndex + 1 To TotalColIndex - 1
                    ApplyColorFormatting = True

                    If SummaryType = En_SummaryType.Monthly Then
                        If Month(CDate(DataGridExpSumm.Columns(ColId).HeaderText)) = Month(CDate(dtpFrom)) _
                                AndAlso Year(CDate(DataGridExpSumm.Columns(ColId).HeaderText)) = Year(CDate(dtpFrom)) _
                                AndAlso CDate(dtpFrom).Day <> 1 Then
                            ApplyColorFormatting = False
                        ElseIf Month(CDate(DataGridExpSumm.Columns(ColId).HeaderText)) = Month(CDate(dtpTo)) _
                                AndAlso Year(CDate(DataGridExpSumm.Columns(ColId).HeaderText)) = Year(CDate(dtpTo)) _
                                AndAlso CDate(dtpTo).Day <> New DateTime(CDate(dtpTo).Year, CDate(dtpTo).Month, Date.DaysInMonth(CDate(dtpTo).Year, CDate(dtpTo).Month)).Day _
                                AndAlso Month(CDate(dtpTo)) <> Today.Month Then
                            ApplyColorFormatting = False
                        End If
                    ElseIf SummaryType = En_SummaryType.Yearly Then
                        If CDate(dtpFrom) > New DateTime(DataGridExpSumm.Columns(ColId).HeaderText, 1, 1) Then
                            ApplyColorFormatting = False
                        ElseIf CDate(dtpTo) < New DateTime(DataGridExpSumm.Columns(ColId).HeaderText, 12, 31) AndAlso CDate(dtpTo).Year <> Today.Year Then
                            ApplyColorFormatting = False
                        End If
                    End If


                    For RowId As Integer = 0 To DataGridExpSumm.Rows.Count - 1
                        If ApplyColorFormatting Then
                            If DataGridExpSumm("Category", RowId).Value = dtBudget.Rows(RowId)("Category") Then
                                If ColId > CategoryColIndex Then
                                    If Not dtBudget.Rows(RowId)(ColId) Is DBNull.Value Then
                                        If dtBudget.Rows(RowId)(ColId) > 0 Then
                                            DataGridExpSumm(ColId, RowId).Style.ForeColor = Color.Black
                                            If DataGridExpSumm(ColId, RowId).Value = dtBudget.Rows(RowId)(ColId) AndAlso DataGridExpSumm(ColId, RowId).Value > 0 Then
                                                DataGridExpSumm(ColId, RowId).Style.ForeColor = Color.Orange
                                            ElseIf DataGridExpSumm(ColId, RowId).Value > dtBudget.Rows(RowId)(ColId) Then
                                                DataGridExpSumm(ColId, RowId).Style.ForeColor = Color.Red
                                            ElseIf DataGridExpSumm(ColId, RowId).Value < dtBudget.Rows(RowId)(ColId) Then
                                                DataGridExpSumm(ColId, RowId).Style.ForeColor = Color.Green
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                Next
            End If
        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub GetBudgets()
        Dim strSQL As String = ""

        If SearchStr Is Nothing Then
            If CallSearchFunction Then
                Call SetDefaultPropValue(True)
            Else
                Call SetDefaultPropValue()
            End If

            If SummaryType = En_SummaryType.Monthly Then
                strSQL = "Execute sp_GetExpenditureSummaryBudget_Monthly '" & dtpFrom & "', '" & dtpTo & "', '" & Categories & "'"
            ElseIf SummaryType = En_SummaryType.Yearly OrElse SummaryType = En_SummaryType.Variance Then
                strSQL = "Execute sp_GetExpenditureSummaryBudget_Yearly '" & dtpFrom & "', '" & dtpTo & "', '" & Categories & "', '" & ChkLBYearsItemsList & "'"
            End If

            Using Connection = GetConnection()
                dtBudget = New DataTable
                da = New SqlDataAdapter(strSQL, Connection)
                da.Fill(dtBudget)
            End Using
        End If
    End Sub

    Private Sub ExportGridToPDF()
        Try
            btnSave.Enabled = False

            TabControl1.Enabled = False
            Panel5.BringToFront()
            Panel5.Visible = True
            Label15.Text = "Exporting Grid To PDF..."
            Application.DoEvents()
            
            'Creating iTextSharp Table from the DataTable data
            Dim pdfTable As New PdfPTable(DataGridExpSumm.ColumnCount - 2)
            pdfTable.DefaultCell.Padding = 3
            pdfTable.DefaultCell.BorderWidth = 1
            pdfTable.HorizontalAlignment = Element.ALIGN_LEFT
            pdfTable.WidthPercentage = 100.0F

            'Adding Header row
            For Each column As DataGridViewColumn In DataGridExpSumm.Columns
                If DataGridExpSumm.Columns(column.Index).Visible Then
                    Dim cell As New PdfPCell(New Phrase(New Chunk(column.HeaderText, FontFactory.GetFont(ImpensaFont.FontFamily.ToString, Nothing, True, ImpensaFont.Size, DataGridExpSumm.Columns(column.Index).HeaderCell.InheritedStyle.Font.Style, New iTextSharp.text.BaseColor(DataGridExpSumm.Columns(column.Index).HeaderCell.InheritedStyle.ForeColor)))))
                    cell.BackgroundColor = New iTextSharp.text.BaseColor(DataGridExpSumm.BackgroundColor.ToArgb)
                    cell.VerticalAlignment = Element.ALIGN_MIDDLE

                    Select Case DataGridExpSumm.Columns(column.Index).HeaderCell.InheritedStyle.Alignment
                        Case DataGridViewContentAlignment.MiddleLeft : cell.HorizontalAlignment = Element.ALIGN_LEFT
                        Case DataGridViewContentAlignment.MiddleRight : cell.HorizontalAlignment = Element.ALIGN_RIGHT
                        Case Else : cell.HorizontalAlignment = Element.ALIGN_LEFT
                    End Select

                    pdfTable.AddCell(cell)
                End If
            Next

            Dim CellVal As String
            'Adding DataRow
            For Each row As DataGridViewRow In DataGridExpSumm.Rows
                For Each cell As DataGridViewCell In row.Cells
                    If DataGridExpSumm.Columns(cell.ColumnIndex).Visible Then
                        If Not cell.Value Is DBNull.Value Then
                            CellVal = IIf(IsNumeric(cell.Value), Format(cell.Value, "#,##0.00").ToString, cell.Value.ToString())
                            Dim PDFCellDet As New PdfPCell(New Phrase(New Chunk(CellVal, FontFactory.GetFont(ImpensaFont.FontFamily.ToString, Nothing, True, ImpensaFont.Size, cell.InheritedStyle.Font.Style, New iTextSharp.text.BaseColor(cell.InheritedStyle.ForeColor.ToArgb)))))
                            PDFCellDet.BackgroundColor = New iTextSharp.text.BaseColor(cell.InheritedStyle.BackColor.ToArgb)
                            PDFCellDet.VerticalAlignment = Element.ALIGN_MIDDLE

                            Select Case cell.InheritedStyle.Alignment
                                Case DataGridViewContentAlignment.MiddleLeft : PDFCellDet.HorizontalAlignment = Element.ALIGN_LEFT
                                Case DataGridViewContentAlignment.MiddleRight : PDFCellDet.HorizontalAlignment = Element.ALIGN_RIGHT
                                Case Else : PDFCellDet.HorizontalAlignment = Element.ALIGN_LEFT
                            End Select

                            pdfTable.AddCell(PDFCellDet)
                        End If
                    End If
                Next
            Next

            If pdfTable.Rows.Count > 0 Then
                'Exporting to PDF
                Dim FileName As String = Path.GetTempPath & "Summary_" & cmbSummaryType.Text & ".pdf"

                If Not ExportPDFProcessID = Nothing Then
                    If Process.GetProcesses.Any(Function(x) x.Id = ExportPDFProcessID) Then
                        Process.GetProcessById(ExportPDFProcessID).Kill()
                        ExportPDFProcessID = Nothing
                        Threading.Thread.Sleep(100)
                    End If
                End If

                If File.Exists(FileName) Then File.Delete(FileName)

                Using stream As New FileStream(FileName, FileMode.Create)
                    Dim pdfDoc As New Document(PageSize.A2, 10.0F, 10.0F, 0.0F, 0.0F)
                    PdfWriter.GetInstance(pdfDoc, stream)
                    pdfDoc.Open()
                    pdfDoc.Add(New Phrase(cmbSummaryType.Text & " Summary"))
                    pdfDoc.Add(pdfTable)
                    pdfDoc.Close()
                    stream.Close()
                End Using

                Dim proc As New Process
                proc.StartInfo.FileName = FileName
                proc.Start()
                ExportPDFProcessID = proc.Id
            End If
        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        Finally
            Panel5.SendToBack()
            Panel5.Visible = False
            btnSave.Enabled = True
            TabControl1.Enabled = True
        End Try
    End Sub
#End Region

#Region "Monthly Budget/Threshold"

    Private Sub PopulateThresholdMonthCombo()
        cmbThrMonth.DataSource = Nothing
        cmbThrMonth.Items.Clear()
        cmbThrMonth.DataSource = mdlChart2.PopulateListingCombo
        cmbThrMonth.DisplayMember = "Value"
        cmbThrMonth.ValueMember = "Key"
        cmbThrMonth.SelectedIndex = -1
    End Sub '13

    Private Sub CheckCurrentMonthThresholds()
        Dim Reader As SqlDataReader
        Dim dtThr As New DataTable
        Dim StrCommand As String = ""
        Try
            Me.TopMost = True
            StrCommand = "SELECT TOP(1) COUNT(*) cnt, dtMonth FROM tbl_ExpThresholds WHERE dtMonth BETWEEN '" & Format(DateSerial(Today.Year, (Today.Month - 1), 1), "yyyy-MM-dd") & "' AND '" & Format(DateSerial(Today.Year, Today.Month, 1), "yyyy-MM-dd") & "' AND TAmount > 0 GROUP BY dtMonth ORDER BY dtMonth DESC"
            Using Connection = GetConnection()
                Cmd = New SqlCommand(StrCommand, GetConnection())
                Reader = Cmd.ExecuteReader

                If Reader.Read Then
                    If Not Reader.GetValue(1) = Format(DateSerial(Today.Year, Today.Month, 1), "yyyy-MM-dd") Then
                        ImpensaAlert("Last month's budgeted amounts are being copied over to this month. To change the budgeted amounts for this month go to ""Monthly Budget"" Tab", MsgBoxStyle.OkOnly + MsgBoxStyle.Information)
                    End If
                Else
                    ImpensaAlert("So far monthly budget has not been set earlier. Set the monthly budget in ""Monthly Budget"" Tab", MsgBoxStyle.OkOnly + MsgBoxStyle.Information)
                End If
                Reader.Close()
            End Using
        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        Finally
            Me.TopMost = False
        End Try
    End Sub '39

    Private Function GetThresholdData(ByVal P_Month As Date) As DataTable
        Dim dt As New DataTable
        Dim Params As SqlParameter
        Dim Comm As SqlCommand
        Dim Da As SqlDataAdapter
        Dim dtFrom As Date
        Dim dtTo As Date
        Try
            If P_Month = CDate("2100-01-01") Then 'Entire Selected Period
                dtFrom = New DateTime(Year(dtpFrom), Month(dtpFrom), 1)
                dtTo = New DateTime(Year(dtpTo), Month(dtpTo), 1).AddMonths(1).AddDays(-1)
            ElseIf P_Month = New DateTime(Year(P_Month), 12, 31) Then 'Entire Year
                If New DateTime(Year(P_Month), 1, 1) < dtpFrom Then
                    dtFrom = New DateTime(Year(dtpFrom), Month(dtpFrom), 1)
                Else
                    dtFrom = New DateTime(Year(P_Month), 1, 1)
                End If

                If dtpTo < New DateTime(Year(P_Month), 12, 31) Then
                    dtTo = New DateTime(Year(dtpTo), Month(dtpTo), 1).AddMonths(1).AddDays(-1)
                Else
                    dtTo = New DateTime(Year(P_Month), 12, 31)
                End If
            ElseIf Month(P_Month) = Month(CDate(dtpFrom)) AndAlso Year(P_Month) = Year(CDate(dtpFrom)) Then 'Starting Month 
                dtFrom = P_Month
                dtTo = New DateTime(CDate(P_Month).Year, CDate(P_Month).Month, 1).AddMonths(1).AddDays(-1)
            ElseIf Month(P_Month) = Month(CDate(dtpTo)) AndAlso Year(P_Month) = Year(CDate(dtpTo)) Then 'Last Month
                dtFrom = P_Month
                dtTo = New DateTime(CDate(P_Month).Year, CDate(P_Month).Month, 1).AddMonths(1).AddDays(-1)

            Else
                dtFrom = P_Month
                dtTo = New DateTime(P_Month.Year, P_Month.Month, 1).AddMonths(1).AddDays(-1)
            End If
            Using Connection = GetConnection()
                Comm = New SqlCommand("sp_GetExpensesThresholds", Connection)
                Comm.CommandType = CommandType.StoredProcedure

                Params = New SqlParameter("@P_FromDate", dtFrom)
                Params.Direction = ParameterDirection.Input
                Params.DbType = DbType.Date
                Comm.Parameters.Add(Params)

                Params = New SqlParameter("@P_ToDate", dtTo)
                Params.Direction = ParameterDirection.Input
                Params.DbType = DbType.Date
                Comm.Parameters.Add(Params)

                Da = New SqlDataAdapter(Comm)
                Da.Fill(dt)
            End Using
        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try

        Return dt
    End Function '40

    Private Sub PopulateThresholdGrid(Optional ByVal P_BudgetBuckets As En_BudgetBuckets = 0)
        Dim dr As DataRow
        Dim dv As New DataView
        Dim ThrAmtTotal As Double = 0
        Dim SpentAmtTotal As Double = 0
        Dim DiffAmtTotal As Double = 0
        Try
            DataGridThrLimits.DataSource = Nothing
            DataGridThrLimits.Columns.Clear()

            Label15.Text = "Loading Monthly Budget..."
            Application.DoEvents()

            Using Connection = GetConnection()
                dtThresholdData = New DataTable
                dtThresholdData = GetThresholdData(ThresholdMonth)
            End Using

            dv = dtThresholdData.DefaultView

            dv.RowFilter = Nothing

            If P_BudgetBuckets = En_BudgetBuckets.OverBudgetCats Then
                dv.RowFilter = "DifferenceSign > 0"
            ElseIf P_BudgetBuckets = En_BudgetBuckets.AtParCats Then
                dv.RowFilter = "DifferenceSign = 0"
            ElseIf P_BudgetBuckets = En_BudgetBuckets.UnderBudgetCats Then
                dv.RowFilter = "DifferenceSign < 0"
            End If

            dtThresholdData = dv.ToTable
            dr = dtThresholdData.NewRow

            For Each dr1 As DataRow In dtThresholdData.Rows
                ThrAmtTotal += dr1("TAmount")
                SpentAmtTotal += dr1("SAmount")
                DiffAmtTotal += dr1("DifferenceSign")
            Next

            dr("CateGory") = "TOTAL"
            dr("TAmount") = ThrAmtTotal
            dr("SAmount") = SpentAmtTotal
            dr("Difference") = DiffAmtTotal
            dr("IsReadOnly") = 1
            dtThresholdData.Rows.Add(dr)

            DataGridThrLimits.DataSource = dtThresholdData
            dtThresholdData.AcceptChanges()

            dv.RowFilter = Nothing

            dv.RowFilter = " Category <> 'TOTAL' AND DifferenceSign > 0"
            tsMenuOBCats.Text = "Over Budget: #" & dv.ToTable.Rows.Count

            dv.RowFilter = "Category <> 'TOTAL' AND DifferenceSign = 0"
            tsMenuAPCats.Text = "At Par: #" & dv.ToTable.Rows.Count

            dv.RowFilter = "Category <> 'TOTAL' AND DifferenceSign < 0"
            tsMenuUBCats.Text = "Under Budget: #" & dv.ToTable.Rows.Count

            DataGridThrLimits.Columns("TAmount").HeaderText = "Budgeted Amount (Rs.)"
            DataGridThrLimits.Columns("SAmount").HeaderText = "Amount Spent(Rs.)"
            DataGridThrLimits.Columns("DifferenceSign").Visible = False

            DataGridThrLimits.AllowUserToAddRows = False
            DataGridThrLimits.AllowUserToResizeRows = False
            DataGridThrLimits.AutoResizeColumns()
            DataGridThrLimits.DefaultCellStyle.WrapMode = DataGridViewTriState.True
            DataGridThrLimits.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill

            For i As Integer = 0 To DataGridThrLimits.Columns.Count - 1
                DataGridThrLimits.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                DataGridThrLimits.Columns(i).HeaderCell.Style.Font = New System.Drawing.Font(ImpensaFont, FontStyle.Bold)
                DataGridThrLimits.Columns(i).DefaultCellStyle.Font = ImpensaFont

                DataGridThrLimits.Columns("TAmount").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
                DataGridThrLimits.Columns("TAmount").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                DataGridThrLimits.Columns("TAmount").DefaultCellStyle.Format = "#,##0.00"

                DataGridThrLimits.Columns("SAmount").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
                DataGridThrLimits.Columns("SAmount").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                DataGridThrLimits.Columns("SAmount").DefaultCellStyle.Format = "#,##0.00"

                DataGridThrLimits.Columns("Difference").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
                DataGridThrLimits.Columns("Difference").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
                DataGridThrLimits.Columns("Difference").DefaultCellStyle.Format = "#,##0.00"

            Next

            DataGridThrLimits.Columns("hKey").Visible = False
            DataGridThrLimits.Columns("iCateGory").Visible = False
            DataGridThrLimits.Columns("dtMonth").Visible = False
            DataGridThrLimits.Columns("IsReadOnly").Visible = False

            DataGridThrLimits.Columns("Category").ReadOnly = True
            DataGridThrLimits.Columns("sAmount").ReadOnly = True
            DataGridThrLimits.Columns("Difference").ReadOnly = True

            Call FormatDataGridThrLimits()
            Call BuildExpenseTicker(En_ToDate.YTD)
            Call BuildExpenseTicker(En_ToDate.MTD)
        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub '14

    Private Sub SaveThresholds()
        Dim dc As SqlCommandBuilder
        Try
            dt = New DataTable
            dt = DirectCast(DataGridThrLimits.DataSource, DataTable).GetChanges
            If Not dt Is Nothing Then
                Using Connection = GetConnection()

                    da = New SqlDataAdapter("SELECT hKey, dtMonth, iCategory, TAmount FROM tbl_ExpThresholds", Connection)
                    dc = New SqlCommandBuilder(da)
                    da.Update(dt)

                    Call PopulateThresholdGrid()
                    ImpensaAlert("Your Changes Have Been Saved.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly)
                End Using
            End If
        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub '6

    Private Sub FormatDataGridThrLimits()
        For Each dr1 As DataGridViewRow In DataGridThrLimits.Rows
            If (dr1.Cells("Category").Value = "TOTAL") Then
                dr1.DefaultCellStyle.BackColor = Color.Yellow
                dr1.DefaultCellStyle.ForeColor = Color.Blue
                dr1.DefaultCellStyle.Font = New System.Drawing.Font(ImpensaFont, FontStyle.Bold)
            End If

            If dr1.Cells("TAmount").Value >= 0 Then
                If dr1.Cells("SAmount").Value > dr1.Cells("TAmount").Value Then
                    If dr1.Cells("Category").Value <> "TOTAL" Then
                        dr1.Cells("Category").Style.ForeColor = Color.Red
                    End If
                    dr1.Cells("Difference").Style.ForeColor = Color.Red
                ElseIf dr1.Cells("SAmount").Value = dr1.Cells("TAmount").Value AndAlso dr1.Cells("SAmount").Value > 0 Then
                    If dr1.Cells("Category").Value <> "TOTAL" Then
                        dr1.Cells("Category").Style.ForeColor = Color.Orange
                    End If
                    dr1.Cells("Difference").Style.ForeColor = Color.Orange
                ElseIf dr1.Cells("TAmount").Value > 0 OrElse dr1.Cells("SAmount").Value > 0 Then
                    If dr1.Cells("Category").Value <> "TOTAL" Then
                        dr1.Cells("Category").Style.ForeColor = Color.Green
                    End If
                    dr1.Cells("Difference").Style.ForeColor = Color.Green
                End If
            End If
        Next
    End Sub '32
#End Region

#Region "Graphical Analysis"

    Private Sub PopulateChartTypeCombo()
        Dim dir As New Dictionary(Of Integer, String)
        Dim binding As BindingSource

        For Each iChart As Integer In [Enum].GetValues(GetType(SeriesChartType))
            If InStr("Line, Column, Pie", [Enum].GetName(GetType(SeriesChartType), iChart).ToString) > 0 Then
                dir.Add(iChart, [Enum].GetName(GetType(SeriesChartType), iChart).ToString)
            End If
        Next

        Dim sorted = From pair In dir Order By pair.Value
        Dim sortedDictionary = sorted.ToDictionary(Function(p) p.Key, Function(p) p.Value)

        binding = New BindingSource(sortedDictionary, Nothing)

        cmbChartType.DataSource = binding
        cmbChartType.DisplayMember = "Value"
        cmbChartType.ValueMember = "Key"
        cmbChartType.SelectedValue = Convert.ToInt32(SeriesChartType.Column)
    End Sub '11

    Private Sub PopulateSelectYearCombo(Optional ByVal P_IsYrClosed As Boolean = 0)
        Dim dt As New DataTable
        Dim Comm As SqlCommand
        Dim Params As SqlParameter

        Try
            Using Conn = GetConnection()
                Comm = New SqlCommand("sp_PopulateSelectYearCombo", Conn)
                Comm.CommandType = CommandType.StoredProcedure

                Params = New SqlParameter("@P_StartDate", dtpRecdKeeping.Value)
                Params.Direction = ParameterDirection.Input
                Params.DbType = DbType.Date
                Comm.Parameters.Add(Params)

                Params = New SqlParameter("@P_IsYrClosed", P_IsYrClosed)
                Params.Direction = ParameterDirection.Input
                Params.DbType = DbType.Boolean
                Comm.Parameters.Add(Params)

                da = New SqlDataAdapter(Comm)
                da.Fill(dt)
            End Using

            cmbSelectYear.DataSource = dt
            cmbSelectYear.DisplayMember = dt.Columns(0).ToString
            cmbSelectYear.ValueMember = dt.Columns(0).ToString
            cmbSelectYear.SelectedIndex = -1

        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub '12

    Private Sub BuildYearsListString()
        Dim strYearsList As String = ""
        Dim YrFrom As Int64
        Dim YrTo As Int64

        If SelectChartCombo = "Chart BKSDTD" Then
            YrFrom = Year(CDate(RecordKeepingStartDate))
            YrTo = Today.Year
        Else
            YrFrom = Year(CDate(dtpFrom))
            YrTo = Year(CDate(dtpTo))
        End If

        If SummaryType = En_SummaryType.Variance Then
            chkLBVarComparision.Items.Clear()
        Else
            chkLBYears.Items.Clear()
        End If

        For i As Integer = 0 To YrTo - YrFrom
            If SummaryType = En_SummaryType.Variance Then
                chkLBVarComparision.Items.Add(YrFrom + i, True)
            Else
                chkLBYears.Items.Add(YrFrom + i, True)
            End If

            strYearsList = strYearsList & CStr(YrFrom + i) & ","
        Next

        ChkLBYearsItemsList = strYearsList.Substring(0, Len(strYearsList) - 1)
    End Sub


    Public Sub DisplayGraph(ByVal Chart_Analysis As System.Windows.Forms.DataVisualization.Charting.Chart)
        Dim dt As New DataTable
        Dim ChartTitle As String = ""
        Dim dv As New DataView
        Dim dtFrom As Date
        Dim dtTo As Date
        Dim lst As New List(Of String)
        Dim bShowLabel = True
        Try
            bShowLabel = IIf(chkShowLabel.Checked, True, False)

            Call ResetChartControl(Chart_Analysis)

            If SelectChartCombo = "Chart 1" Then
                dt = mdlChart1.GetChartData(dtpFrom, dtpTo, ListingCombo.key, SearchStr)
            ElseIf SelectChartCombo = "Chart 2" Then
                If ListingCombo.key = CDate("2100-01-01") Then
                    dtFrom = dtpFrom
                    dtTo = dtpTo
                ElseIf Month(ListingCombo.key) = Month(CDate(dtpFrom)) AndAlso Year(ListingCombo.key) = Year(CDate(dtpFrom)) Then
                    dtFrom = dtpFrom
                    dtTo = New DateTime(CDate(dtpFrom).Year, CDate(dtpFrom).Month, 1).AddMonths(1).AddDays(-1)
                ElseIf ListingCombo.key = New DateTime(Year(ListingCombo.key), 12, 31) Then
                    If New DateTime(Year(ListingCombo.key), 1, 1) < dtpFrom Then
                        dtFrom = dtpFrom
                    Else
                        dtFrom = New DateTime(Year(ListingCombo.key), 1, 1)
                    End If
                    If dtpTo < New DateTime(Year(ListingCombo.key), 12, 31) Then
                        dtTo = dtpTo
                    Else
                        dtTo = New DateTime(Year(ListingCombo.key), 12, 31)
                    End If
                ElseIf Month(ListingCombo.key) = Month(CDate(dtpTo)) AndAlso Year(ListingCombo.key) = Year(CDate(dtpTo)) Then
                    dtFrom = ListingCombo.key
                    dtTo = dtpTo
                Else
                    dtFrom = ListingCombo.key
                    dtTo = New DateTime(ListingCombo.key.Year, ListingCombo.key.Month, 1).AddMonths(1).AddDays(-1)
                End If

                dt = mdlChart2.GetChartData(dtFrom, dtTo, SearchStr)
            ElseIf SelectChartCombo = "Chart MTD" Then
                ChkLBYearsItemsList = Today.Year.ToString
                dt = mdlChart2.GetChartData(New DateTime(Today.Year, Today.Month, 1), Today.Date, Nothing)
                dv = dt.DefaultView
                dv.RowFilter = "Amount > 0"
                dt = dv.ToTable
            ElseIf SelectChartCombo = "Chart YTD" Then
                ChkLBYearsItemsList = Today.Year.ToString
                dt = mdlChart2.GetChartData(New DateTime(Today.Year, 1, 1), New DateTime(Today.Year, 12, 31), Nothing)
                dv = dt.DefaultView
                dv.RowFilter = "Amount > 0"
                dt = dv.ToTable
            ElseIf SelectChartCombo = "Chart BKSDTD" Then
                Call BuildYearsListString()
                dt = mdlChart2.GetChartData(CDate(RecordKeepingStartDate).ToString("yyyy-MM-dd"), Today.Date, Nothing)
                dv = dt.DefaultView
                dv.RowFilter = "Amount > 0"
                dt = dv.ToTable
            ElseIf SelectChartCombo = "Chart 3A" Then
                dt = mdlChart3.GetChartData(dtpFrom, dtpTo, ListingCombo.key, SearchStr, chkPeriodLevel.Checked)
            ElseIf SelectChartCombo = "Chart 3B" Then
                dt = mdlChart3.GetChartData_3B(dtpFrom, dtpTo, ListingCombo.key, SearchStr, chkPeriodLevel.Checked)
            ElseIf SelectChartCombo = "Chart 4" Then
                dt = mdlChart4.GetChartData(dtpFrom, dtpTo, ListingCombo.key, SearchStr, chkPeriodLevel.Checked)
            End If

            For i As Int32 = 1 To dt.Columns.Count - 1
                If (SelectChartCombo = "Chart 3A" OrElse SelectChartCombo = "Chart 3B") Then
                    lst.Add(dt.Columns(i).ColumnName & " - Rs." & Format(dt.Compute("SUM([" & dt.Columns(i).ColumnName & "])", ""), "#,##0.00"))
                Else
                    lst.Add("Rs." & Format(dt.Compute("SUM([" & dt.Columns(i).ColumnName & "])", ""), "#,##0.00"))
                End If

            Next

            ChartDataSetAmtTotal = lst

            If dt.Rows.Count > 0 Then
                IsDataAvailable = True
                If SortByAmount Then
                    dv = dt.DefaultView
                    If (SelectChartCombo = "Chart 3A" OrElse SelectChartCombo = "Chart 3B") AndAlso Year(dtpFrom) = Year(dtpTo) Then
                        dv.Sort = dt.Columns(1).ColumnName & " " & cmbSort.SelectedItem
                    Else
                        dv.Sort = "Amount " & IIf((SelectChartCombo = "Chart MTD" OrElse SelectChartCombo = "Chart YTD" OrElse SelectChartCombo = "Chart BKSDTD"), "Asc", cmbSort.SelectedItem)
                    End If
                    dt = dv.ToTable
                End If
            Else
                IsDataAvailable = False
                ImpensaAlert("No Data Available.", MsgBoxStyle.Information)
                Exit Sub
            End If

            If InStr("Chart MTD, Chart YTD, Chart BKSDTD", SelectChartCombo) = 0 Then
                If cmbListing.Enabled Then
                    If InStr(ListingCombo.Value, "Year") > 0 AndAlso Not SelectChartCombo = "Chart 3B" Then
                        ChartTitle = ListingCombo.Value.ToString.Substring(0, 10)
                    Else
                        ChartTitle = ListingCombo.Value
                    End If
                Else
                    ChartTitle = SelectChartCombo
                End If
            Else
                If SelectChartCombo = "Chart MTD" Then
                    ChartTitle = MonthName(DateTime.Today.Month) & " - " & Year(Today)
                ElseIf SelectChartCombo = "Chart YTD" Then
                    ChartTitle = " Year - " & Year(Today)
                Else
                    ChartTitle = CDate(RecordKeepingStartDate).ToString("dd/MM/yyyy", CultureInfo.InvariantCulture) & " - " & Today.Date.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture)
                End If
            End If

            Chart_Analysis.Titles.Add(ChartTitle)
            Chart_Analysis.Titles(0).Alignment = ContentAlignment.MiddleCenter
            Chart_Analysis.Titles(0).Font = New System.Drawing.Font(ImpensaFont.FontFamily, 18, FontStyle.Bold)
            Chart_Analysis.Titles(0).ForeColor = Color.Blue
            Chart_Analysis.Titles(0).BackColor = Color.Yellow
            Chart_Analysis.Titles(0).BorderColor = Color.Black

            Dim ChartSubTitle As String = "Total: "

            For Each item As String In lst
                ChartSubTitle = ChartSubTitle & item & "; "
            Next

            Chart_Analysis.Titles.Add(ChartSubTitle)
            Chart_Analysis.Titles(1).Alignment = ContentAlignment.MiddleRight
            Chart_Analysis.Titles(1).Font = New System.Drawing.Font(ImpensaFont.FontFamily, 10, FontStyle.Bold)
            Chart_Analysis.Titles(1).ForeColor = Color.Blue

            If InStr("Chart MTD, Chart YTD, Chart BKSDTD", SelectChartCombo) = 0 Then
                Chart_Analysis.Legends("Legend1").Enabled = ShowLegends(SelectChartCombo, ChartTypeCombo.value)
            Else
                Chart_Analysis.Legends("Legend1").Enabled = True
            End If
            Chart_Analysis.Legends("Legend1").TextWrapThreshold = 0

            For i As Integer = 1 To dt.Columns.Count - 1
                Chart_Analysis.Series.Add(dt.Columns(i).ColumnName)
                If InStr("Chart MTD, Chart YTD, Chart BKSDTD", SelectChartCombo) = 0 Then
                    Chart_Analysis.Series(dt.Columns(i).ColumnName).ChartType = ChartTypeCombo.Key
                Else
                    Chart_Analysis.Series(dt.Columns(i).ColumnName).ChartType = SeriesChartType.Pie
                End If

                Chart_Analysis.Series(dt.Columns(i).ColumnName).Points.DataBind(dt.DefaultView, dt.Columns(0).ColumnName, dt.Columns(i).ColumnName, "")
                Chart_Analysis.Series(dt.Columns(i).ColumnName).SmartLabelStyle.Enabled = True
                Chart_Analysis.Series(dt.Columns(i).ColumnName).SmartLabelStyle.CalloutStyle = LabelCalloutStyle.None
                Chart_Analysis.Series(dt.Columns(i).ColumnName).SmartLabelStyle.CalloutLineAnchorCapStyle = LineAnchorCapStyle.None
                Chart_Analysis.Series(dt.Columns(i).ColumnName).SmartLabelStyle.CalloutLineColor = Color.Transparent
                Chart_Analysis.Series(dt.Columns(i).ColumnName).IsValueShownAsLabel = bShowLabel
                Chart_Analysis.Series(dt.Columns(i).ColumnName)("PieLabelStyle") = "Disabled"
                Chart_Analysis.Series(dt.Columns(i).ColumnName).Font = New System.Drawing.Font(ImpensaFont, FontStyle.Regular)
                Chart_Analysis.Series(dt.Columns(i).ColumnName).LabelFormat = "#,##0.00"

                If SelectChartCombo = "Chart 3B" Then
                    Chart_Analysis.Series(dt.Columns(i).ColumnName).IsValueShownAsLabel = False
                End If

                If InStr("Line", [Enum].GetName(GetType(SeriesChartType), Chart_Analysis.Series(dt.Columns(i).ColumnName).ChartType).ToString) > 0 Then
                    Chart_Analysis.Series(dt.Columns(i).ColumnName).BorderWidth = 2
                    Chart_Analysis.Series(dt.Columns(i).ColumnName).MarkerStyle = MarkerStyle.Circle
                    Chart_Analysis.Series(dt.Columns(i).ColumnName).MarkerSize = 5
                ElseIf InStr("Pie", [Enum].GetName(GetType(SeriesChartType), Chart_Analysis.Series(dt.Columns(i).ColumnName).ChartType).ToString) > 0 Then
                    Chart_Analysis.Series(dt.Columns(i).ColumnName).Label = "#VALX --> Rs. #VALY{#,##0.00} (#PERCENT)"
                End If
            Next

            Chart_Analysis.ChartAreas.Add("Area 1")
            Chart_Analysis.ChartAreas("Area 1").Position.Auto = True
            Chart_Analysis.ChartAreas("Area 1").AxisX.Title = dt.Columns(0).ColumnName
            Chart_Analysis.ChartAreas("Area 1").AxisX.TitleFont = New System.Drawing.Font(ImpensaFont, FontStyle.Bold)
            Chart_Analysis.ChartAreas("Area 1").AxisX.Interval = 1
            If cmbChartType.SelectedItem.Value = "Line" Then
                Chart_Analysis.ChartAreas("Area 1").AxisX.MajorGrid.LineWidth = 1
                Chart_Analysis.ChartAreas("Area 1").AxisX.MajorGrid.LineDashStyle = DataVisualization.Charting.ChartDashStyle.Dash
                Chart_Analysis.ChartAreas("Area 1").AxisX.MajorGrid.LineColor = Color.LightGray
            Else
                Chart_Analysis.ChartAreas("Area 1").AxisX.MajorGrid.LineWidth = 0
            End If
            Chart_Analysis.ChartAreas("Area 1").AxisX.LabelStyle.Font = New System.Drawing.Font(ImpensaFont, FontStyle.Regular)
            Chart_Analysis.ChartAreas("Area 1").AxisX.LabelStyle.Angle = 90

            Chart_Analysis.ChartAreas("Area 1").AxisY.Title = "Amount"
            Chart_Analysis.ChartAreas("Area 1").AxisY.TitleFont = New System.Drawing.Font(ImpensaFont, FontStyle.Bold)
            Chart_Analysis.ChartAreas("Area 1").AxisY.MajorGrid.LineWidth = 1
            Chart_Analysis.ChartAreas("Area 1").AxisY.MajorGrid.LineDashStyle = ChartDashStyle.Dash
            Chart_Analysis.ChartAreas("Area 1").AxisY.MajorGrid.LineColor = Color.LightGray
            Chart_Analysis.ChartAreas("Area 1").AxisY.LabelStyle.Font = New System.Drawing.Font(ImpensaFont, FontStyle.Regular)

            For Each s As Series In Chart_Analysis.Series
                For Each p As DataPoint In s.Points

                    For i As Integer = 0 To (p.YValues.Length - 1)
                        If p.YValues(i) = 0 Then
                            p.IsValueShownAsLabel = False
                            p.LabelForeColor = Color.Transparent
                        Else
                            If SelectChartCombo = "Chart 3A" OrElse SelectChartCombo = "Chart 3B" Then
                                p.ToolTip = s.Name & "-" & p.AxisLabel & ": Rs. " & "#VAL{#,##0.00}"
                            Else
                                p.ToolTip = p.AxisLabel & ": Rs. " & "#VAL{#,##0.00} (#PERCENT)"
                            End If
                        End If
                    Next
                Next
            Next

            Chart_Analysis.Palette = ChartColorScheme

        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub '26

    Private Function ShowLegends(ByVal ChartName As String, ByVal ChartType As String) As Boolean
        If InStr("Chart 1, Chart 2, Chart 4", ChartName) > 0 And InStr("Pie, Doughnut, Pyramid", ChartType) > 0 Then
            Return True
        ElseIf ChartName = "Chart 3A" OrElse ChartName = "Chart 3B" Then
            Return True
        Else
            Return False
        End If
    End Function '27

    Private Sub SetChartColorPalette()
        ChartColorScheme = Nothing
        Do While (ChartColorScheme = Nothing OrElse ChartColorScheme = ChartColorPalette.Light OrElse ChartColorScheme = ChartColorPalette.Grayscale)
            ChartColorScheme = Rnd.Next([Enum].GetValues(GetType(ChartColorPalette)).Cast(Of Integer).Min, ([Enum].GetValues(GetType(ChartColorPalette)).Cast(Of Integer).Max + 1))
        Loop
    End Sub '45
#End Region

#Region "Categories"

    Private Sub PopulateCategoryGrid()
        Dim dc As New DataGridViewCheckBoxColumn

        Try
            Label15.Text = "Loading Categories..."
            Application.DoEvents()

            DataGridCatList.DataSource = Nothing

            dc.Name = "bDelete"
            dc.HeaderText = "Delete"
            dc.DataPropertyName = "bDelete"
            dc.AutoSizeMode = DataGridViewAutoSizeColumnMode.ColumnHeader
            DataGridCatList.Columns.Add(dc)

            Using Connection = GetConnection()
                dt = New DataTable
                da = New SqlDataAdapter("SELECT C.*,CASE WHEN (SELECT COUNT(*) FROM tbl_ExpenditureDet E WHERE C.hKey = E.iCategory) > 0 THEN 0 ELSE -1 END CanDelete, 0 'bDelete' FROM tbl_CategoryList C ORDER BY C.sCategory", Connection)
                da.Fill(dt)
            End Using

            DataGridCatList.DataSource = dt
            DataGridCatList.DefaultCellStyle.WrapMode = DataGridViewTriState.True
            DataGridCatList.Columns("bDelete").DisplayIndex = 0
            DataGridCatList.Columns("hKey").Visible = False
            DataGridCatList.Columns("CanDelete").Visible = False
            DataGridCatList.RowHeadersVisible = False
            DataGridCatList.Columns("sCategory").HeaderText = "Category"
            DataGridCatList.Columns("sCategory").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            CType(DataGridCatList.Columns("sCategory"), DataGridViewTextBoxColumn).MaxInputLength = 100
            DataGridCatList.Columns("sNotes").HeaderText = "Notes"
            DataGridCatList.Columns("sNotes").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
            CType(DataGridCatList.Columns("sNotes"), DataGridViewTextBoxColumn).MaxInputLength = 500
            DataGridCatList.Columns("IsObsolete").HeaderText = "Do Not Use"
            DataGridCatList.Columns("IsObsolete").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridCatList.Columns("IsObsolete").DisplayIndex = 6
            DataGridCatList.Columns("bNotify").HeaderText = "Notify Unpaid"
            DataGridCatList.Columns("bNotify").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            DataGridCatList.Columns("iMonthlyOccurrences").HeaderText = "Monthly Occurrences"
            DataGridCatList.Columns("iMonthlyOccurrences").HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            CType(DataGridCatList.Columns("iMonthlyOccurrences"), DataGridViewTextBoxColumn).MaxInputLength = 2
            DataGridCatList.Columns("iMonthlyOccurrences").DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter

            For i As Integer = 0 To DataGridCatList.Columns.Count - 1
                DataGridCatList.Columns(i).SortMode = DataGridViewColumnSortMode.NotSortable
                DataGridCatList.Columns(i).HeaderCell.Style.Font = New System.Drawing.Font(ImpensaFont, FontStyle.Bold)
            Next

            DataGridCatList.DefaultCellStyle.WrapMode = DataGridViewTriState.True
        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub '9

    Private Sub SaveCategories()
        Dim dc As SqlCommandBuilder
        Try
            dt = New DataTable
            dt = DirectCast(DataGridCatList.DataSource, DataTable).GetChanges

            Using Connection = GetConnection()
                da = New SqlDataAdapter("SELECT * FROM tbl_CategoryList", Connection)
                dc = New SqlCommandBuilder(da)
                Cmd = New SqlCommand()
                Cmd.Connection = Connection
                If Not dt Is Nothing Then
                    For Each dr As DataRow In dt.Rows
                        If dr("IsObsolete") Is DBNull.Value Then
                            dr("IsObsolete") = 0
                        End If

                        If Not (IsDBNull(dr.Item("bDelete"))) Then
                            If (Convert.ToBoolean(dr.Item("bDelete"))) Then
                                Cmd.CommandText = "DELETE FROM tbl_ExpThresholds WHERE iCategory = " & dr("hKey")
                                Cmd.ExecuteNonQuery()
                                dr.Delete()
                            End If
                        End If
                    Next
                    da.Update(dt)
                End If

                ImpensaAlert("Your Changes Have Been Saved.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly)
                DirectCast(DataGridCatList.DataSource, DataTable).AcceptChanges()
                Call ResetFilters()
            End Using
        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub '2

#End Region

#Region "Settings"

    Private Sub GetConfig()
        Try
            If AssemblyLocation Is Nothing Then
                AssemblyLocation = System.Reflection.Assembly.GetEntryAssembly().Location & " -startup"
            End If

            If HighlightDetail Is Nothing Then
                HighlightDetail = 1000
                txtHighlightDet.Text = 1000
            Else
                txtHighlightDet.Text = HighlightDetail
                HighlightDetail_Orig = HighlightDetail
            End If

            If HighlightSummMonthly Is Nothing Then
                HighlightSummMonthly = 3000
                txtHighlightSummMth.Text = 3000
            Else
                txtHighlightSummMth.Text = HighlightSummMonthly
                HighlightSummMonthly_Orig = HighlightSummMonthly
            End If

            If HighlightSummYearly Is Nothing Then
                HighlightSummYearly = 25000
                txtHighlightSummYr.Text = 25000
            Else
                txtHighlightSummYr.Text = HighlightSummYearly
                HighlightSummYearly_Orig = HighlightSummYearly
            End If

            If RecordKeepingStartDate Is Nothing Then
                RecordKeepingStartDate = CDate("01/01/2012")
                dtpRecdKeeping.Value = Format(CDate("01/01/2012"), "dd/MM/yyyy")
            Else
                dtpRecdKeeping.Value = Format(CDate(RecordKeepingStartDate), "dd/MM/yyyy")
            End If

            txtCSVBackupPath.Text = CSVBackupPath
            chkShowReminder.Checked = ShowReminder
            txtReminder.Text = ReminderText

            If ShowReminder Then
                txtReminder.Enabled = True
            Else
                txtReminder.Enabled = False
            End If

            chkStartImport.Checked = EnableImport

        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub '18

    Private Sub SaveConfig()
        Try
            HighlightDetail = txtHighlightDet.Text
            HighlightSummMonthly = txtHighlightSummMth.Text
            HighlightSummYearly = txtHighlightSummYr.Text
            HighlightDetail_Orig = txtHighlightDet.Text
            HighlightSummMonthly_Orig = txtHighlightSummMth.Text
            HighlightSummYearly_Orig = txtHighlightSummYr.Text

            If Not String.IsNullOrEmpty(txtCSVBackupPath.Text) Then
                CSVBackupPath = txtCSVBackupPath.Text
            End If
            If rbCloseYr.Checked Or rbOpenYr.Checked Then
                Call SaveEOY()
                cmbSelectYear.SelectedItem = Nothing
            End If

            ShowReminder = chkShowReminder.Checked
            ReminderText = txtReminder.Text
            EnableImport = chkStartImport.Checked

            'Call AlertTicker(ShowReminder)

            ImpensaAlert("Your Changes Have Been Saved.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly)

            If Not RecordKeepingStartDate = dtpRecdKeeping.Value Then
                RecordKeepingStartDate = dtpRecdKeeping.Value
                Call ResetFilters(P_bCallRefreshGrids:=False)
            End If

            Call StartImportService()
        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub '3
#End Region

#Region "End Of Year"

    Private Sub CheckOpenYears()
        Call PopulateSelectYearCombo()
        Dim StrOpenYrs = BuildOpenOrClosedYrsStr(0)

        If Not StrOpenYrs Is Nothing Then
            Me.TopMost = True
            'Start reminding to close the previous year once the first 15th days of new year are passed
            If DateDiff(DateInterval.Day, CDate("31-12-" & Year(DateTime.Today) - 1), DateTime.Today) > 15 Then
                Dim OpenYrs() As String = BuildOpenOrClosedYrsStr(0).Split(",") 'List Of Open Years
                For i As Integer = 0 To OpenYrs.Count - 2
                    If ImpensaActionAlert("Year " & OpenYrs(i) & " is still OPEN. Do you want to close this year now?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                        Using Connection = GetConnection()
                            Cmd = New SqlCommand("UPDATE tbl_EOY SET IsYrClosed = 1 WHERE Year# = " & OpenYrs(i), Connection)
                            Cmd.ExecuteNonQuery()

                            da = New SqlDataAdapter("SELECT E.dtDate [Date], C.sCategory, E.dAmount [Amount], E.sNotes [Notes] FROM tbl_ExpenditureDet E INNER JOIN tbl_CategoryList C ON C.hKey = E.iCategory WHERE YEAR (dtDate) = " & OpenYrs(i), Connection)
                            dt = New DataTable
                            da.Fill(dt)

                            If dt.Rows.Count > 0 Then
                                Call CreateCSV(dt, "Data_" & OpenYrs(i) & "Y#.csv", En_CSVCreationFrequency.Annually)
                            End If
                        End Using
                    End If
                Next
            End If
        End If
        Me.TopMost = False
    End Sub '31

    Private Function BuildOpenOrClosedYrsStr(ByVal P_IsYearClosed As Boolean) As String
        '0 --> Open Year
        '1 --> Closed Year
        Dim StrYrList As String = ""
        Dim StrSQL As String = ""
        Try
            StrSQL = "Select CONVERT(CHAR(4), Year#) +',' FROM tbl_EOY WHERE ISNULL(IsYrClosed,0) = " & IIf(P_IsYearClosed = True, 1, 0) & " FOR XML PATH('')"
            Using Connection = GetConnection()
                Cmd = New SqlCommand(StrSQL, Connection)
                StrYrList = Cmd.ExecuteScalar
            End Using
        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try
        Return StrYrList
    End Function '16

    Private Sub SaveEOY()
        Dim StrCommand As String = ""

        Try
            If Not cmbSelectYear.SelectedValue Is Nothing Then

                StrCommand = "UPDATE tbl_EOY SET IsYrClosed = " & IIf(rbCloseYr.Checked, 1, 0) & " WHERE Year# = " & cmbSelectYear.SelectedValue

                Using Connection = GetConnection()
                    Cmd = New SqlCommand(StrCommand, Connection)
                    Cmd.ExecuteNonQuery()

                    If rbCloseYr.Checked Then
                        StrCommand = "SELECT E.dtDate [Date], C.sCategory, E.dAmount [Amount], E.sNotes [Notes] FROM tbl_ExpenditureDet E INNER JOIN tbl_CategoryList C ON C.hKey = E.iCategory AND C.IsObsolete = 0 WHERE YEAR (dtDate) = " & cmbSelectYear.SelectedValue
                        da = New SqlDataAdapter(StrCommand, Connection)
                        dt = New DataTable
                        da.Fill(dt)

                        If dt.Rows.Count > 0 Then
                            Call CreateCSV(dt, "Data_" & cmbSelectYear.SelectedValue & "Y#.csv", En_CSVCreationFrequency.Annually)
                        End If
                    End If
                End Using
            End If

            rbCloseYr.Checked = False
            rbOpenYr.Checked = False
            cmbSelectYear.Enabled = False

        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub '4
#End Region

#Region "Reset"
    Private Sub ResetFilters(Optional ByVal P_bCallRefreshGrids As Boolean = True)
        CallSearchFunction = False
        dtPickerTo.MaxDate = Today.Date
        dtPickerFrom.MinDate = dtpRecdKeeping.Value
        dtPickerFrom.Value = Now.AddDays((Now.Day - 1) * -1).Date 'First Day Of Current Month
        dtPickerTo.Value = DateTime.Today
        Call PopulatePeriodCombo()
        cmbPeriod.SelectedIndex = En_Period.CurrentMonth
        txtSearch.Text = DefaultSearchStr
        txtSearch.ForeColor = Color.DarkGray
        'Call GetNonObseletCategoriesList()
        Call PopulateLstCategoryFilter()
        If P_bCallRefreshGrids Then
            Call RefreshGrids()
        End If
    End Sub '22

    Private Sub ResetAllTabs()
        DataGridExpDet.Columns.Clear()
        DataGridExpSumm.Columns.Clear()
        DataGridCatList.Columns.Clear()
        DataGridThrLimits.Columns.Clear()
        cmbSelectChart.SelectedIndex = -1
        cmbThrMonth.SelectedIndex = -1
        cmbListing.DataSource = Nothing
        cmbListing.Items.Clear()
        cmbChartType.SelectedValue = Convert.ToInt32(SeriesChartType.Column)
        chkSort.Checked = False
        cmbSort.SelectedIndex = 0
        chkPeriodLevel.Checked = False
        chkPeriodLevel.Enabled = False
        chkPeriodLevel.Checked = False
        cmbThrMonth.SelectedIndex = -1
        Label8.Visible = False
        Label8.Text = String.Empty
        tslblGridTotal.Text = String.Empty
        tslblSeperator1.Text = String.Empty
        tslblRecdCnt.Text = String.Empty
        tslblSeperator2.Text = String.Empty
        cmbSummaryType.SelectedIndex = 0
        cmbCatListRunTot.SelectedIndex = -1
        cmbCatListRunTot.Enabled = False
        SummaryType = En_SummaryType.Monthly
        HighlightDetail = IIf(HighlightDetail_Orig Is Nothing, 1000, HighlightDetail_Orig)
        HighlightSummMonthly = IIf(HighlightSummMonthly_Orig Is Nothing, 3000, HighlightSummMonthly_Orig)
        HighlightSummYearly = IIf(HighlightSummYearly_Orig Is Nothing, 25000, HighlightSummYearly_Orig)
        txtHighlight.Text = HighlightDetail
        chkLBYears.Items.Clear()
        chkLBVarComparision.Items.Clear()
        Call ResetChartControl(Chart_Analysis)
        'Call ShowErrorsAndExceptions()
    End Sub '23

    Private Sub ResetChartControl(ByVal Chart_Analysis As System.Windows.Forms.DataVisualization.Charting.Chart)
        Chart_Analysis.Titles.Clear()
        Chart_Analysis.Series.Clear()
        Chart_Analysis.ChartAreas.Clear()
        ChartDataSetAmtTotal = Nothing
    End Sub '24
#End Region

#Region "CSV"
    Private Sub CreateCSV(ByVal P_dt As DataTable, ByVal P_FileName As String, ByVal P_Frequency As En_CSVCreationFrequency)
        Dim CSVHeader As String = "Date, Category, Amount, Notes"
        Dim CurrentLine As String = ""
        Dim FilePath As String = CSVBackupPath + "\" + P_FileName

        Try
            If File.Exists(FilePath) Then
                If P_Frequency = En_CSVCreationFrequency.Annually Or P_Frequency = En_CSVCreationFrequency.Adhoc Then
                    File.Delete(FilePath)
                Else
                    Exit Sub
                End If
            End If

            File.Create(FilePath).Close()

            Reader = New StreamReader(FilePath)

            If Reader.EndOfStream Then
                Reader.Close()
                Writer = New StreamWriter(FilePath, True)
                Writer.WriteLine(CSVHeader)
            Else
                While Not Reader.EndOfStream
                    CurrentLine = Reader.ReadLine.ToString
                    If CurrentLine.ToString.Contains(CSVHeader) Then
                        Reader.Close()
                        Writer = New StreamWriter(FilePath, True)
                        Writer.WriteLine(CSVHeader)
                    End If
                    Exit While
                End While
            End If

            If Not Writer Is Nothing Then Writer.Close()
            Writer = New StreamWriter(FilePath, True)

            For Each dr As DataRow In P_dt.Rows
                Writer.WriteLine(dr.Item("Date") & "," & dr.Item("sCategory") & "," & dr.Item("Amount") & "," & dr.Item("Notes"))
            Next
        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        Finally
            If Not Writer Is Nothing Then Writer.Close()
            If Not Reader Is Nothing Then Reader.Close()
        End Try
    End Sub '28

    Private Sub CreateMonthlyCSV()
        Dim StrCommand As String = ""

        StrCommand = "SELECT E.dtDate [Date], C.sCategory, E.dAmount [Amount], E.sNotes [Notes] FROM tbl_ExpenditureDet E INNER JOIN tbl_CategoryList C ON C.hKey = E.iCategory WHERE MONTH(dtDate) = " & (DatePart(DateInterval.Month, DateTime.Today) - 1)
        Using Connection = GetConnection()
            da = New SqlDataAdapter(StrCommand, Connection)
            dt = New DataTable
            da.Fill(dt)
        End Using

        If dt.Rows.Count > 0 Then
            Call CreateCSV(dt, "Data_" & DateAndTime.MonthName((DatePart(DateInterval.Month, DateTime.Today) - 1)) & Year(DateTime.Today) & "M#.csv", En_CSVCreationFrequency.Monthly)
        End If
    End Sub '29

    Private Sub CreateAdhocCSV()
        Dim strMonthList As String = ""
        Dim StrCommand As String = ""
        Dim dtGrid = DirectCast(DataGridExpDet.DataSource, DataTable).GetChanges
        If Not dtGrid Is Nothing Then
            For Each drGrid As DataRow In dtGrid.Rows
                If DatePart(DateInterval.Month, CDate(drGrid("Date"))) < DatePart(DateInterval.Month, DateTime.Today) Then
                    Dim FilePath As String = CSVBackupPath & "Data_" & DateAndTime.MonthName(DatePart(DateInterval.Month, CDate(drGrid("Date")))) _
                    & DatePart(DateInterval.Year, CDate(drGrid("Date"))) & "M#.csv"

                    If InStr(strMonthList, DatePart(DateInterval.Month, CDate(drGrid("Date")))) = 0 Then
                        strMonthList = strMonthList + DatePart(DateInterval.Month, CDate(drGrid("Date"))) + ","

                        StrCommand = "SELECT E.dtDate [Date], C.sCategory, E.dAmount [Amount], E.sNotes [Notes] FROM tbl_ExpenditureDet E INNER JOIN tbl_CategoryList C " & _
                                     "ON C.hKey = E.iCategory WHERE MONTH(dtDate) = " & DatePart(DateInterval.Month, CDate(drGrid("Date")))

                        Using Connection = GetConnection()
                            da = New SqlDataAdapter(StrCommand, Connection)
                            dt = New DataTable
                            da.Fill(dt)
                        End Using

                        If dt.Rows.Count > 0 Then
                            Call CreateCSV(dt, "Data_" & DateAndTime.MonthName(DatePart(DateInterval.Month, CDate(drGrid("Date")))) _
                                               & DatePart(DateInterval.Year, CDate(drGrid("Date"))) & "M#.csv", En_CSVCreationFrequency.Adhoc)
                        End If
                    End If
                End If
            Next
        End If
    End Sub '30
#End Region

#Region "Ticker"
    
    Private Sub ShowErrorsAndExceptions()
        Dim AlertText As String = Nothing

        If LastFailedCnt > 0 Then
            AlertText = LastFailedCnt & " RECORD(S) FAILED TO IMPORT IN LAST ATTEMPT MADE AT " & LogTimeStamp
            lblAlertText.ForeColor = Color.Red
        ElseIf ImportExceptionOccurred Then
            AlertText = "ERROR OCCURRED DURING LAST IMPORT ATTEMPT. PLEASE CHECK ERROR LOG PRESENT AT IMPORT PATH."
            lblAlertText.ForeColor = Color.Red
        ElseIf Not String.IsNullOrEmpty(ReminderText) Then
            AlertText = IIf(ShowReminder, ReminderText, String.Empty)
            lblAlertText.ForeColor = Color.White
        End If

        If Not String.IsNullOrEmpty(AlertText) Then
            lblAlertText.Width = AlertText.Length
            'lblAlertText.Left = Me.Width
            Panel10.Visible = True
            tmrAlert.Enabled = True
        Else
            Panel10.Visible = False
            tmrAlert.Enabled = False
        End If

        If Not String.IsNullOrEmpty(AlertText) Then
            lblAlertText.Text = Microsoft.VisualBasic.Strings.Replace(AlertText, vbNewLine, "  ||  ")
        End If

    End Sub

    Private Sub BuildExpenseTicker(ByVal P_ToDateFactor As En_ToDate)
        Dim Params As SqlParameter
        Dim Comm As SqlCommand
        Dim strTicker As String = ""
        Try
            Using Connection = GetConnection()
                Comm = New SqlCommand("sp_GetExpensesTicker", Connection)
                Comm.CommandType = CommandType.StoredProcedure

                Params = New SqlParameter("@P_FromDate", IIf(P_ToDateFactor = En_ToDate.MTD, New DateTime(Today.Year, Today.Month, 1), New DateTime(Today.Year, 1, 1)))
                Params.Direction = ParameterDirection.Input
                Params.DbType = DbType.Date
                Comm.Parameters.Add(Params)

                Params = New SqlParameter("@P_ToDate", IIf(P_ToDateFactor = En_ToDate.MTD, New DateTime(Today.Year, Today.Month, 1).AddMonths(1).AddDays(-1), New DateTime(Today.Year, 12, 31)))
                Params.Direction = ParameterDirection.Input
                Params.DbType = DbType.Date
                Comm.Parameters.Add(Params)

                dt = New DataTable
                da = New SqlDataAdapter(Comm)
                da.Fill(dt)
            End Using

            For Each dr As DataRow In dt.Rows
                strTicker = strTicker & dr("Category") & " - Rs. " & Format(dr("Amount"), "#,##0.00") & " / Rs. " & Format(dr("TAmount"), "#,##0.00") & ";" & Space(5)
            Next

            If P_ToDateFactor = En_ToDate.MTD Then
                RchTB_MTDTicker.Text = strTicker
                Dim g As Graphics = RchTB_MTDTicker.CreateGraphics
                Dim sz As SizeF = TextRenderer.MeasureText(g, RchTB_MTDTicker.Text, RchTB_MTDTicker.Font, RchTB_MTDTicker.ClientSize, TextFormatFlags.SingleLine)
                'To keep both YTD and MTD tickers in sync, setting MTD ticker length = YTD ticker length
                RchTB_MTDTicker.Width = RchTB_YTDTicker.Width

            Else
                RchTB_YTDTicker.Text = strTicker
                Dim g As Graphics = RchTB_YTDTicker.CreateGraphics
                Dim sz As SizeF = TextRenderer.MeasureText(g, RchTB_YTDTicker.Text, RchTB_YTDTicker.Font, RchTB_YTDTicker.ClientSize, TextFormatFlags.SingleLine)
                RchTB_YTDTicker.Width = CInt(Math.Ceiling(sz.Width))
            End If

            If P_ToDateFactor = En_ToDate.MTD Then
                Call SetTickerTextColor(RchTB_MTDTicker, dt)
            Else
                Call SetTickerTextColor(RchTB_YTDTicker, dt)
            End If
        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub '37

    Private Sub SetTickerTextColor(ByVal P_RchTB As RichTextBox, ByVal dt As DataTable)
        For Each dr As DataRow In dt.Rows
            If P_RchTB.Find(dr("Category") & " - Rs. " & Format(dr("Amount"), "#,##0.00") & " / Rs. " & Format(dr("TAmount"), "#,##0.00") & ";") > -1 Then
                If dr("TAmount") = 0 AndAlso dr("Amount") = 0 Then
                    P_RchTB.SelectionColor = Color.Black
                Else
                    If dr("IsOverBudget") = 1 Then
                        P_RchTB.SelectionColor = Color.Red
                    ElseIf dr("IsOverBudget") = 0 Then
                        P_RchTB.SelectionColor = Color.Orange
                    ElseIf dr("IsOverBudget") = -1 Then
                        P_RchTB.SelectionColor = Color.Green
                    End If
                End If
            End If
        Next
    End Sub '38
#End Region

#Region "Filters"

    Private Sub PopulateLstCategoryFilter()
        Try
            Using Connection = GetConnection()
                dt = New DataTable
                da = New SqlDataAdapter("SELECT sCategory FROM tbl_CategoryList ORDER BY sCategory", Connection)
                da.Fill(dt)
            End Using

            LstCategory.Items.Clear()

            For Each dr As DataRow In dt.Rows
                LstCategory.Items.Add(dr.Item("sCategory").ToString)
            Next
        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub '10

#Region "Period Combo Methods"
    Private Sub PopulatePeriodCombo()
        cmbPeriod.Items.Clear()
        cmbPeriod.Items.Add(New KeyValuePair(Of Integer, String)(0, "Current Month (MTD)"))
        cmbPeriod.Items.Add(New KeyValuePair(Of Integer, String)(1, "Current Year (YTD)"))
        cmbPeriod.Items.Add(New KeyValuePair(Of Integer, String)(2, CDate(RecordKeepingStartDate).ToString("dd/MM/yyyy", CultureInfo.InvariantCulture) & " - " & Today.Date.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture)))
        cmbPeriod.DisplayMember = "Value"
        cmbPeriod.ValueMember = "Key"
        cmbPeriod.SelectedIndex = 0
    End Sub '46 

    Private Sub SetPeriodComboIndex()
        If dtPickerFrom.Value.Date = Now.AddDays((Now.Day - 1) * -1).Date AndAlso dtPickerTo.Value.Date = Today.Date Then
            cmbPeriod.SelectedIndex = En_Period.CurrentMonth
        ElseIf dtPickerFrom.Value.Date = New DateTime(Year(Today), 1, 1) AndAlso dtPickerTo.Value.Date = Today.Date Then
            cmbPeriod.SelectedIndex = En_Period.CurrentYear
        ElseIf dtPickerFrom.Value.Date = RecordKeepingStartDate AndAlso dtPickerTo.Value.Date = Today.Date Then
            cmbPeriod.SelectedIndex = En_Period.BookStart
        Else
            cmbPeriod.SelectedIndex = -1
        End If
    End Sub '47
#End Region
#End Region

    Protected Overrides Sub SetVisibleCore(ByVal value As Boolean)
        Dim Args As String() = Environment.GetCommandLineArgs
        If Array.IndexOf(Args, "Startup") > -1 Then
            If Not Me.IsHandleCreated Then
                FirstTimeLoadInBkg = True
                Me.CreateHandle()
                value = False
                NotifyIcon.BalloonTipText = "Impensa is running in background."
                NotifyIcon.ShowBalloonTip(100)
            Else
                FirstTimeLoadInBkg = False
            End If
        End If
        MyBase.SetVisibleCore(value)

        Call StartImportService()
    End Sub

    Private Sub GetConnectionInfo()
        Try
            If ConnStr Is Nothing Then
                Call CloseSplash()
                ImpensaAlert("No Connection details found. Please provide database connection details.", MsgBoxStyle.Information)
                frmLogin.ShowDialog()
                If ConnStr Is Nothing Then
                    End
                End If
            Else
                Dim TestConn As New SqlConnection(ConnStr)
                TestConn.Open()
                If TestConn.State = ConnectionState.Open Then
                    TestConn.Close()
                End If
            End If
        Catch sqlEx As SqlException
            ImpensaAlert("Database connection could not be established", MsgBoxStyle.Critical)
            frmLogin.ShowDialog()
        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub '17

    Private Sub ShowTDLableDetails()
        Dim Reader As SqlDataReader
        Using Connection = GetConnection()
            Cmd = New SqlCommand("SELECT * FROM dbo.Fn_GetTDTotal('" & CDate(RecordKeepingStartDate).ToString("yyyy-MM-dd") & "')", Connection)
            Reader = Cmd.ExecuteReader

            If Reader.HasRows Then
                Reader.Read()
                tslblMTD.Text = "MTD Total: Rs. " & Format(Reader.GetValue(0), "#,##0.00")
                tslblYTD.Text = "YTD Total: Rs. " & Format(Reader.GetValue(1), "#,##0.00")
                tslblBKSDTD.Text = "(" & CDate(RecordKeepingStartDate).ToString("dd/MM/yyyy", CultureInfo.InvariantCulture) & " - " & Today.Date.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture) & ") Total: Rs. " & Format(Reader.GetValue(2), "#,##0.00")
            End If
            Reader.Close()
        End Using
    End Sub '25

    Private Function ShowUnsavedDataWarning(ByVal sMessage As String) As Windows.Forms.DialogResult
        Dim Result As Windows.Forms.DialogResult
        If (TabControl1.TabPages(LastTabIndex).Name = "TabExpDet" AndAlso (Not DirectCast(DataGridExpDet.DataSource, DataTable).GetChanges Is Nothing)) OrElse _
        (TabControl1.TabPages(LastTabIndex).Name = "TabThreshold" AndAlso (Not DirectCast(DataGridThrLimits.DataSource, DataTable).GetChanges Is Nothing)) OrElse _
        (TabControl1.TabPages(LastTabIndex).Name = "TabCatList" AndAlso (Not DirectCast(DataGridCatList.DataSource, DataTable).GetChanges Is Nothing)) Then
            If ImpensaActionAlert(sMessage, MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) = Windows.Forms.DialogResult.Cancel Then
                Result = Windows.Forms.DialogResult.Cancel
            Else
                Result = Windows.Forms.DialogResult.OK
            End If
        End If

        Return Result
    End Function '42

    Private Sub RefreshGrids()
        Try
            TabControl1.Enabled = False
            Panel5.BringToFront()
            Panel5.Visible = True
            Panel1.Enabled = False
            btnSave.Enabled = False
            Application.DoEvents()

            If Not CallSearchFunction Then
                chkShowAllDet.Visible = False
                Call SetDefaultPropValue()
            End If

            If TabControl1.SelectedIndex <> 0 Then TabControl1.SelectedIndex = 0

            Call ResetAllTabs()
            Call ShowTDLableDetails()
            Call PopulateCategoryGrid()
            Call PopulateThresholdMonthCombo()
            Call PopulateUnpaidBillsCurrentMonth()
            Call PopulateUnpaidBillsPreviousMonth()

            If ThresholdCurrentMonthIndex <> -1 Then
                ThresholdMonth = cmbThrMonth.Items(ThresholdCurrentMonthIndex).Key
                cmbThrMonth.SelectedIndex = ThresholdCurrentMonthIndex
                Call PopulateThresholdGrid()
            End If

            If Not SearchStr Is Nothing OrElse CallSearchFunction Then
                Call PopulateSearchResults()
            Else
                Call PopulateExpenditureDetailGrid()
            End If

            If CallSearchFunction Then
                Call SetDefaultPropValue()
            End If

            Call BuildExpenseTicker(En_ToDate.MTD)
            Call BuildExpenseTicker(En_ToDate.YTD)
            Call PopulateExpenditureSummaryGrid()
        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        Finally
            Panel5.SendToBack()
            Panel5.Visible = False
            TabControl1.Enabled = True
            Panel1.Enabled = True
            btnSave.Enabled = True
        End Try

    End Sub '20

    Private Sub CloseSplash()
        If Splash.InvokeRequired Then
            Splash.TopMost = False
            Splash.SendToBack()
            Splash.Invoke(New Del_CloseSpash(AddressOf CloseSplash))
        Else
            Splash.TopMost = False
            Splash.SendToBack()
            Splash.Close()
        End If
        Me.TopMost = False
    End Sub '21

    Private Sub SetDefaultPropValue(Optional ByVal P_CallSearchFunction As Boolean = False)
        dtpFrom = dtPickerFrom.Value.Date.ToString("yyyy-MM-dd")
        dtpTo = dtPickerTo.Value.Date.ToString("yyyy-MM-dd")
        StrLstCategories = ""
        CallSearchFunction = P_CallSearchFunction

        For i As Integer = 0 To LstCategory.CheckedItems.Count - 1
            StrLstCategories = StrLstCategories & LstCategory.CheckedItems(i).ToString & ","
        Next

        If StrLstCategories <> "" Then
            StrLstCategories = StrLstCategories.Substring(0, Len(StrLstCategories) - 1)
        End If

        Categories = StrLstCategories

        SearchStr = IIf(txtSearch.Text.Equals(DefaultSearchStr), Nothing, Trim(txtSearch.Text))
    End Sub '43

    Private Sub ExitApplication()
        HighlightDetail = IIf(HighlightDetail_Orig Is Nothing, 1000, HighlightDetail_Orig)
        HighlightSummMonthly = IIf(HighlightSummMonthly_Orig Is Nothing, 3000, HighlightSummMonthly_Orig)
        HighlightSummYearly = IIf(HighlightSummYearly_Orig Is Nothing, 25000, HighlightSummYearly_Orig)

        NotifyIcon.Icon = Nothing
        NotifyIcon.Visible = False
        NotifyIcon.Dispose()
        Application.Exit()
    End Sub

    Private Sub DataImport()
        ImportFailedCnt = 0
        LastFailedCnt = 0
        ImportSkippedCnt = 0
        ImportSucceessCnt = 0
        TotalImportCnt = 0
        ImportSucceessAndFailCnt = 0

        Call clsLib.DataImport()

        If ImportSucceessAndFailCnt > 0 Then
            NotifyIcon.BalloonTipText = TotalImportCnt & " Records Found." & vbCrLf & "IMPORT STATUS: Success: " & ImportSucceessCnt & " | " & "Failed: " & ImportFailedCnt & _
                                        " | " & "Skipped: " & ImportSkippedCnt

            NotifyIcon.ShowBalloonTip(100)
        End If
    End Sub

    Private Sub PopulateUnpaidBillsCurrentMonth()
        Try
            Using Connection = GetConnection()
                dt = New DataTable
                da = New SqlDataAdapter("EXEC sp_GetUnpaidBillsCurrentMonth", Connection)
                da.Fill(dt)
            End Using

            LstUnpaidBillsCurrentMonth.Items.Clear()

            If dt.Rows.Count = 0 Then
                Dim dr As DataRow = dt.NewRow()
                dr("UnpaidCategory") = "No Unpaid Bills"
                LstUnpaidBillsCurrentMonth.Items.Add(dr.Item("UnpaidCategory").ToString)
            Else
                UnpaidBills = True
            End If


            For Each dr As DataRow In dt.Rows
                LstUnpaidBillsCurrentMonth.Items.Add(dr.Item("UnpaidCategory").ToString)
            Next

        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try

    End Sub

    Private Sub PopulateUnpaidBillsPreviousMonth()
        Try
            Using Connection = GetConnection()
                dt = New DataTable
                da = New SqlDataAdapter("EXEC sp_GetUnpaidBillsPreviousMonth", Connection)
                da.Fill(dt)
            End Using

            LstUnpaidBillsPrevMonth.Items.Clear()

            If dt.Rows.Count = 0 Then
                Dim dr As DataRow = dt.NewRow()
                dr("UnpaidCategory") = "No Unpaid Bills"
                LstUnpaidBillsPrevMonth.Items.Add(dr.Item("UnpaidCategory").ToString)
            Else
                UnpaidBills = True
            End If


            For Each dr As DataRow In dt.Rows
                LstUnpaidBillsPrevMonth.Items.Add(dr.Item("UnpaidCategory").ToString)
            Next

        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try

    End Sub

    Private Sub InsertCategoryPrevMonthOccurrences()
        Try
            Using Connection = GetConnection()
                Cmd = New SqlCommand()
                Cmd.Connection = Connection
                Cmd.CommandText = "EXEC sp_InsertCategoryPrevMonthOccurrences"
                Cmd.ExecuteNonQuery()
            End Using
        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

#End Region

#Region "Control Events"

#Region "frmMain Events"

    Private Sub frmMain_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            DataGridExpDet.Controls.Clear()

            Label15.Text = String.Empty

            Call GetConnectionInfo()
            Call GetConfig()
            dtPickerTo.MaxDate = Today.Date
            dtPickerFrom.MinDate = dtpRecdKeeping.Value
            dtPickerFrom.Value = Now.AddDays((Now.Day - 1) * -1).Date 'First Day Of Current Month
            dtPickerTo.Value = DateTime.Today
            dtpFrom = dtPickerFrom.Value.Date.ToString("yyyy-MM-dd")
            dtpTo = dtPickerTo.Value.Date.ToString("yyyy-MM-dd")
            dtpRecdKeeping.MaxDate = DateTime.Today
            SummaryType = En_SummaryType.Monthly
            txtHighlight.Text = HighlightDetail
            ImpensaFont = New System.Drawing.Font("Microsoft Sans Serif", 10)
            cmbSelectChart.Items.Clear()
            cmbSelectChart.Items.AddRange(New String() {"Chart 1", "Chart 2", "Chart 3A", "Chart 3B", "Chart 4"})

            cmbSort.Items.Clear()
            cmbSort.Items.Add("Asc")
            cmbSort.Items.Add("Desc")
            cmbSort.SelectedIndex = 0

            cmbSummaryType.Items.Clear()
            cmbSummaryType.Items.Add(New KeyValuePair(Of Integer, String)(0, "Monthly"))
            cmbSummaryType.Items.Add(New KeyValuePair(Of Integer, String)(1, "Yearly"))
            cmbSummaryType.Items.Add(New KeyValuePair(Of Integer, String)(2, "Running Totals"))

            If CDate(dtpFrom) = New DateTime(Year(CDate(dtpFrom)), 1, 1) AndAlso CDate(dtpTo) = New DateTime(Year(CDate(dtpTo)), 12, 31) AndAlso Year(CDate(dtpFrom)) <> Year(CDate(dtpTo)) Then
                cmbSummaryType.Items.Add(New KeyValuePair(Of Integer, String)(3, "Variance Comparision"))
            End If

            cmbSummaryType.DisplayMember = "Value"
            cmbSummaryType.ValueMember = "Key"
            cmbSummaryType.SelectedIndex = 0

            Call PopulatePeriodCombo()

            cmbCatListRunTot.DataSource = mdlChart1.PopulateListingCombo
            cmbCatListRunTot.DisplayMember = "Value"
            cmbCatListRunTot.ValueMember = "Key"
            cmbCatListRunTot.SelectedIndex = -1

            dtp = New DateTimePicker
            dtp.Font = ImpensaFont
            dtp.Format = DateTimePickerFormat.Custom
            dtp.CustomFormat = "dd/MM/yyyy"
            dtp.Visible = False

            DataGridExpDet.Controls.Add(dtp)

            Call CheckOpenYears()
            Call CheckCurrentMonthThresholds()

            If (CDate(LastUsedTimeStamp).Month <> DateTime.Today.Month) Then
                Call InsertCategoryPrevMonthOccurrences()
            End If

            If CStr(LastACRefreshDate) Is Nothing OrElse LastACRefreshDate <= New DateTime(Today.Year, 1, 1) Then
                Call RefreshAutoCompleteDictionary()
            End If

            Call PopulateLstCategoryFilter()
            Call PopulateChartTypeCombo()
            Call BuildYearsListString()
            Call SetNotesFieldAutoCompleteSource()
            Call RefreshGrids()

            RchTB_MTDTicker.Left = Me.Width
            RchTB_MTDTicker.Visible = True
            RchTB_YTDTicker.Left = Me.Width
            RchTB_YTDTicker.Visible = True
            tmrTicker.Start()
            LastUsedTimeStamp = DateTime.Now
        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub frmMain_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If e.CloseReason = CloseReason.UserClosing Then
            e.Cancel = True
            Me.Hide()
            NotifyIcon.BalloonTipText = "Impensa is running in background."
            NotifyIcon.ShowBalloonTip(100)
        End If
    End Sub

    Private Sub frmMain_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        Panel5.Parent = Panel3
        Panel5.Left = (Panel5.Parent.ClientSize.Width - Panel5.Width) / 2
        Panel5.Top = (Panel5.Parent.ClientSize.Height - Panel5.Height) / 2
        pbCredit.Left = (pbCredit.Parent.ClientSize.Width - pbCredit.Width) / 2
        Panel9.Width = Me.Width
        Panel10.Width = Me.Width
    End Sub
#End Region

#Region "Button Events"
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        If TabControl1.SelectedTab.Name = "TabExpDet" Then
            Call SaveNotes(DirectCast(DataGridExpDet.DataSource, DataTable).GetChanges)
            Call SaveExpenses()
        ElseIf TabControl1.SelectedTab.Name = "TabCatList" Then
            Call SaveCategories()
        ElseIf TabControl1.SelectedTab.Name = "TabFunctions" Then
            Call SaveConfig()
            Call RefreshGrids()
        ElseIf TabControl1.SelectedTab.Name = "TabThreshold" Then
            Call SaveThresholds()
        ElseIf TabControl1.SelectedTab.Name = "TabExpSumm" Then
            Call ExportGridToPDF()
        End If
    End Sub

    Private Sub btnGo_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGo.Click
        Dim strYearsList As String = ""

        SelectChartCombo = cmbSelectChart.SelectedItem
        ListingCombo = cmbListing.SelectedItem
        ChartTypeCombo = cmbChartType.SelectedItem

        If cmbSort.Enabled Then
            SortByAmount = True
        Else
            SortByAmount = False
        End If

        For Each item As Object In chkLBYears.CheckedItems
            strYearsList = strYearsList & item & ","
        Next

        ChkLBYearsItemsList = strYearsList.Substring(0, Len(strYearsList) - 1)

        Call SetChartColorPalette()
        Call DisplayGraph(Chart_Analysis)
    End Sub

    Private Sub btnSubmit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSubmit.Click
        Try
            If ShowUnsavedDataWarning("You have unsaved data in the grid(s). Navigating away or refreshing grid may cause you to loose unsaved data. Continue?") = Windows.Forms.DialogResult.Cancel Then
                Exit Sub
            End If
            Call ResetAllTabs()
            Call SetDefaultPropValue()
            Call BuildYearsListString()

            If cmbSummaryType.Items.Count = 4 Then
                If cmbSummaryType.Items(3).Equals("Variance Comparision") Then
                    cmbSummaryType.Items.RemoveAt(3)
                End If
            End If

            If CDate(dtpFrom) = New DateTime(Year(CDate(dtpFrom)), 1, 1) AndAlso CDate(dtpTo) = New DateTime(Year(CDate(dtpTo)), 12, 31) AndAlso Year(CDate(dtpFrom)) <> Year(CDate(dtpTo)) AndAlso Not cmbSummaryType.Items.Contains(New KeyValuePair(Of Integer, String)(3, "Variance Comparision")) Then
                cmbSummaryType.Items.Add(New KeyValuePair(Of Integer, String)(3, "Variance Comparision"))
            End If

            Call RefreshGrids()
        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub btnBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse.Click
        Dim fbd As New FolderBrowserDialog
        fbd.ShowDialog()
        txtCSVBackupPath.Text = fbd.SelectedPath
    End Sub

    Private Sub btnHighlight_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHighlight.Click
        If Not txtHighlight Is Nothing Then
            If TabControl1.SelectedTab.Name = "TabExpDet" Then
                HighlightDetail = txtHighlight.Text
                If CallSearchFunction Then
                    Call FormatSearchResults()
                Else
                    Call FormatDetailGrid()
                End If

            ElseIf TabControl1.SelectedTab.Name = "TabExpSumm" Then
                If SummaryType = En_SummaryType.Monthly OrElse SummaryType = En_SummaryType.RunningTotals Then
                    HighlightSummMonthly = txtHighlight.Text
                ElseIf SummaryType = En_SummaryType.Yearly Then
                    HighlightSummYearly = txtHighlight.Text
                End If
                Call FormatSummaryGrid()
            End If
        End If
    End Sub

    Private Sub btnVarCompare_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnVarCompare.Click
        Dim strYearsList As String = ""

        For Each item As Object In chkLBVarComparision.CheckedItems
            strYearsList = strYearsList & item & ","
        Next

        ChkLBYearsItemsList = strYearsList.Substring(0, Len(strYearsList) - 1)

        TabControl1.Enabled = False
        Panel5.BringToFront()
        Panel5.Visible = True
        Call PopulateExpenditureSummaryGrid()
        Call FormatSummaryGrid()
        Panel5.SendToBack()
        Panel5.Visible = False
        TabControl1.Enabled = True

    End Sub
#End Region

#Region "DataGrid Events"
#Region "DataGridExpDet Events"
    Private Sub DataGridExpDet_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DataGridExpDet.CellBeginEdit
        Try
            If (DataGridExpDet.Focused And DataGridExpDet.CurrentCell.ColumnIndex = DataGridExpDet.Columns("Date").Index) Then
                dtp.Location = DataGridExpDet.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, False).Location
                dtp.Size = DataGridExpDet.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, False).Size
                dtp.Visible = True

                If Not (IsDBNull(DataGridExpDet.CurrentCell.Value)) Then
                    dtp.Value = Convert.ToDateTime(DataGridExpDet.CurrentCell.Value)
                Else
                    dtp.Value = DateTime.Today
                End If
            Else
                dtp.Visible = False
            End If
        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub DataGridExpDet_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridExpDet.CellEndEdit
        Try
            dtp.Visible = False

            If (DataGridExpDet.Focused And DataGridExpDet.CurrentCell.ColumnIndex = DataGridExpDet.Columns("Date").Index) Then
                DataGridExpDet.CurrentCell.Value = Microsoft.VisualBasic.Strings.Replace(dtp.Value.Date, "-", "/")
            End If

            If CBool(DataGridExpDet.Rows(DataGridExpDet.CurrentCell.RowIndex).Cells("IsDummy").Value) Then
                Dim RowIndex As Integer = DataGridExpDet.CurrentCell.RowIndex

                For i As Integer = RowIndex + 1 To DataGridExpDet.Rows.Count - 1
                    If (DataGridExpDet.Rows(i).Cells("iCategory").Value Is DBNull.Value _
                    And DataGridExpDet.Rows(i).Cells("Amount").Value Is DBNull.Value _
                    And DataGridExpDet.Rows(i).Cells("Notes").Value Is DBNull.Value) Then
                        DataGridExpDet.Rows(i).Cells("Date").Value = DataGridExpDet.Rows(RowIndex).Cells("Date").Value
                        dtp.Value = CDate(DataGridExpDet.Rows(RowIndex).Cells("Date").Value)

                        Dim drv As DataRowView = DirectCast(DataGridExpDet.Rows(i).DataBoundItem, DataRowView)
                        Dim dr As DataRow = DirectCast(drv.Row, DataRow)
                        dr.AcceptChanges()
                    End If
                Next
            End If

            DgvTextBox = Nothing
            DgvComboBox = Nothing
        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub DataGridExpDet_CellLeave(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridExpDet.CellLeave
        If DataGridExpDet.CurrentCell.ColumnIndex = DataGridExpDet.Columns("Notes").Index Then
            If Not String.IsNullOrEmpty(Trim(DataGridExpDet.CurrentCell.EditedFormattedValue)) Then
                If Not AutoList.Contains(DataGridExpDet.CurrentCell.EditedFormattedValue) Then
                    AutoList.Add(DataGridExpDet.CurrentCell.EditedFormattedValue)
                End If
            End If
        End If

        If DataGridExpDet.CurrentCell.ColumnIndex = DataGridExpDet.Columns("Amount").Index And Not String.IsNullOrEmpty(DataGridExpDet(DataGridExpDet.CurrentCell.ColumnIndex, DataGridExpDet.CurrentCell.RowIndex).EditedFormattedValue) Then
            Dim st As New DataGridViewCellStyle
            st.ForeColor = Color.Gray
            If DataGridExpDet(DataGridExpDet.CurrentCell.ColumnIndex, DataGridExpDet.CurrentCell.RowIndex).EditedFormattedValue < 0 Then
                DataGridExpDet.Rows(DataGridExpDet.CurrentCell.RowIndex).DefaultCellStyle.ForeColor = Color.Gray
            End If
        End If
    End Sub

    Private Sub DataGridExpDet_CellValidating(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellValidatingEventArgs) Handles DataGridExpDet.CellValidating
        If DataGridExpDet.CurrentCell.ColumnIndex = DataGridExpDet.Columns("Date").Index Then
            DataGridExpDet.CancelEdit()

            If InStr(StrClosedYrs, CStr(Year(dtp.Value))) > 0 Then
                ImpensaAlert("Cannot add entry to the year " & CStr(Year(dtp.Value)) & ". This year has already been closed.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly)
                e.Cancel = True
            End If
        End If
    End Sub

    Private Sub DataGridExpDet_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridExpDet.CellValueChanged
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            DataGridExpDet(e.ColumnIndex, e.RowIndex).Style.BackColor = Color.LightCyan
        End If
    End Sub

    Private Sub DataGridExpDet_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridExpDet.DataError
        If (e.Context _
                = (DataGridViewDataErrorContexts.Formatting Or DataGridViewDataErrorContexts.PreferredSize)) Then
            e.ThrowException = False
        End If
    End Sub

    Private Sub DataGridExpDet_EditingControlShowing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles DataGridExpDet.EditingControlShowing

        Dgv = DataGridExpDet

        If Dgv.CurrentCell.ColumnIndex = Dgv.Columns("Amount").Index Then
            If DgvTextBox Is Nothing And TypeOf e.Control Is DataGridViewTextBoxEditingControl Then
                DgvTextBox = CType(e.Control, DataGridViewTextBoxEditingControl)
            End If
        End If

        If Dgv.CurrentCell.ColumnIndex = Dgv.Columns("iCategory").Index Then
            If DgvComboBox Is Nothing And TypeOf e.Control Is DataGridViewComboBoxEditingControl Then
                DgvComboBox = CType(e.Control, DataGridViewComboBoxEditingControl)
            End If
        End If


        If Dgv.CurrentCell.ColumnIndex = Dgv.Columns("Notes").Index Then
            Dim TBEditCtrl As DataGridViewTextBoxEditingControl
            TBEditCtrl = CType(e.Control, DataGridViewTextBoxEditingControl)
            TBEditCtrl.AutoCompleteSource = AutoCompleteSource.CustomSource
            TBEditCtrl.AutoCompleteCustomSource = AutoList
            TBEditCtrl.AutoCompleteMode = AutoCompleteMode.SuggestAppend
            TBEditCtrl.Multiline = False
            TBEditCtrl.MaxLength = 500
        End If
    End Sub

    Private Sub DataGridExpDet_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DataGridExpDet.KeyDown
        Try
            If (DataGridExpDet.Focused And DataGridExpDet.CurrentCell.ColumnIndex = DataGridExpDet.Columns("Date").Index) Then
                If e.KeyCode = Keys.Delete Then
                    dtp.Format = DateTimePickerFormat.Custom
                    dtp.CustomFormat = ""
                    DataGridExpDet.CurrentCell.Value = DBNull.Value
                End If
            End If
        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub DataGridExpDet_RowLeave(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridExpDet.RowLeave
        Dim dr As DataRow
        If DataGridExpDet("IsDummy", e.RowIndex).Value And DataGridExpDet("IsDummyRowAdded", e.RowIndex).Value = 0 Then
            If (DataGridExpDet("hKey", e.RowIndex).Value Is DBNull.Value And Not DataGridExpDet("Date", e.RowIndex).EditedFormattedValue Is DBNull.Value And Not String.IsNullOrEmpty(DataGridExpDet("iCategory", e.RowIndex).EditedFormattedValue) And Not String.IsNullOrEmpty(DataGridExpDet("Amount", e.RowIndex).EditedFormattedValue)) Then
                dr = dtDetailGrid.NewRow

                dr("Date") = Microsoft.VisualBasic.Strings.Replace(DateTime.Today, "-", "/")
                dr("IsReadOnly") = 0
                dr("IsDummy") = 1
                dr("IsDummyRowAdded") = 0
                dtDetailGrid.Rows.InsertAt(dr, dtDetailGrid.Rows.Count - 2)


                DataGridExpDet("IsDummyRowAdded", e.RowIndex).Value = 1
            End If
        End If
    End Sub

    Private Sub DataGridExpDet_RowsAdded(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowsAddedEventArgs) Handles DataGridExpDet.RowsAdded
        If Not DoNotChkRowAdded Then
            DataGridExpDet.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.LightCyan
        End If
    End Sub
#End Region

#Region "DataGripExpSumm Events"

    Private Sub DataGridExpSumm_CellFormatting(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellFormattingEventArgs) Handles DataGridExpSumm.CellFormatting
        DataGridExpSumm.Columns(e.ColumnIndex).SortMode = DataGridViewColumnSortMode.NotSortable
        DataGridExpSumm.Columns(e.ColumnIndex).HeaderCell.Style.Font = New System.Drawing.Font(ImpensaFont, FontStyle.Bold)
        DataGridExpSumm.Columns(e.ColumnIndex).DefaultCellStyle.Font = ImpensaFont
        If DataGridExpSumm.Columns(e.ColumnIndex).HeaderCell.Value <> "Category" Then
            DataGridExpSumm.Columns(e.ColumnIndex).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleRight
            DataGridExpSumm.Columns(e.ColumnIndex).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            DataGridExpSumm.Columns(e.ColumnIndex).DefaultCellStyle.Format = "#,##0.00"
            If DataGridExpSumm.Columns(e.ColumnIndex).Name = DateAndTime.MonthName(DatePart(DateInterval.Month, DateTime.Today)) & "-" & Year(DateTime.Today) Then
                DataGridExpSumm.Columns(e.ColumnIndex).DefaultCellStyle.ForeColor = Color.Blue
                DataGridExpSumm.EnableHeadersVisualStyles = False
                DataGridExpSumm.Columns(e.ColumnIndex).HeaderCell.Style.ForeColor = Color.Blue
            End If
        End If
    End Sub

    Private Sub DataGridExpSumm_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridExpSumm.DataError
        If (e.Context _
                = (DataGridViewDataErrorContexts.Formatting Or DataGridViewDataErrorContexts.PreferredSize)) Then
            e.ThrowException = False
        End If
    End Sub

    Private Sub DataGridExpSumm_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridExpSumm.CellDoubleClick
        Dim v_dtFrom As String = dtPickerFrom.Value.Date.ToString("yyyy-MM-dd")
        Dim v_dtTo As String = dtPickerTo.Value.Date.ToString("yyyy-MM-dd")

        If DataGridExpSumm.Cursor = Cursors.Hand Then
            If SummaryType = En_SummaryType.Monthly Then
                If CStr(CDate(v_dtFrom).Year) & "-" & CStr(CDate(v_dtFrom).Month) = CStr(CDate(v_dtTo).Year) & "-" & CStr(CDate(v_dtTo).Month) Then
                    'No change to v_dtFrom, v_dtTo
                ElseIf Month(CDate(DataGridExpSumm.Columns(e.ColumnIndex).HeaderText)) = Month(CDate(v_dtFrom)) AndAlso Year(CDate(DataGridExpSumm.Columns(e.ColumnIndex).HeaderText)) = Year(CDate(v_dtFrom)) Then 'First Month: 'No Change to v_dtFrom
                    'No change to v_dtFrom
                    v_dtTo = New DateTime(CDate(DataGridExpSumm.Columns(e.ColumnIndex).HeaderText).Year, CDate(DataGridExpSumm.Columns(e.ColumnIndex).HeaderText).Month, 1).AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd")
                ElseIf Month(CDate(DataGridExpSumm.Columns(e.ColumnIndex).HeaderText)) = Month(CDate(v_dtTo)) AndAlso Year(CDate(DataGridExpSumm.Columns(e.ColumnIndex).HeaderText)) = Year(CDate(v_dtTo)) Then 'Last Month: 'No Change to v_dtTo
                    v_dtFrom = New DateTime(CDate(DataGridExpSumm.Columns(e.ColumnIndex).HeaderText).Year, CDate(DataGridExpSumm.Columns(e.ColumnIndex).HeaderText).Month, 1).ToString("yyyy-MM-dd")
                    'No Change to v_dtTo
                Else
                    v_dtFrom = New DateTime(CDate(DataGridExpSumm.Columns(e.ColumnIndex).HeaderText).Year, CDate(DataGridExpSumm.Columns(e.ColumnIndex).HeaderText).Month, 1).ToString("yyyy-MM-dd")
                    v_dtTo = New DateTime(CDate(DataGridExpSumm.Columns(e.ColumnIndex).HeaderText).Year, CDate(DataGridExpSumm.Columns(e.ColumnIndex).HeaderText).Month, 1).AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd")
                End If
            ElseIf SummaryType = En_SummaryType.Yearly Then
                If CStr(CDate(v_dtFrom).Year) & "-" & CStr(CDate(v_dtFrom).Month) = CStr(CDate(v_dtTo).Year) & "-" & CStr(CDate(v_dtTo).Month) Then
                    'No change to v_dtFrom, v_dtTo
                ElseIf DataGridExpSumm.Columns(e.ColumnIndex).HeaderText = Year(CDate(v_dtFrom)) Then 'First Year: 'No change to v_dtFrom
                    v_dtTo = New DateTime(DataGridExpSumm.Columns(e.ColumnIndex).HeaderText, 12, 31).ToString("yyyy-MM-dd")
                ElseIf DataGridExpSumm.Columns(e.ColumnIndex).HeaderText = Year(CDate(v_dtTo)) Then 'Last Year: 'No Change to v_dtTo
                    v_dtFrom = New DateTime(DataGridExpSumm.Columns(e.ColumnIndex).HeaderText, 1, 1).ToString("yyyy-MM-dd")
                Else
                    v_dtFrom = New DateTime(DataGridExpSumm.Columns(e.ColumnIndex).HeaderText, 1, 1).ToString("yyyy-MM-dd")
                    v_dtTo = New DateTime(DataGridExpSumm.Columns(e.ColumnIndex).HeaderText, 12, 31).ToString("yyyy-MM-dd")
                End If
            ElseIf SummaryType = En_SummaryType.RunningTotals Then
                If CStr(CDate(v_dtFrom).Year) & "-" & CStr(CDate(v_dtFrom).Month) = CStr(CDate(v_dtTo).Year) & "-" & CStr(CDate(v_dtTo).Month) Then
                    'No change to v_dtFrom, v_dtTo
                ElseIf DataGridExpSumm("iCategory", e.RowIndex).Value = Month(CDate(v_dtFrom)) AndAlso DataGridExpSumm.Columns(e.ColumnIndex).HeaderText = Year(CDate(v_dtFrom)) Then 'First Year: 'No change to v_dtFrom
                    'No change to v_dtFrom
                    v_dtTo = New DateTime(DataGridExpSumm.Columns(e.ColumnIndex).HeaderText, DataGridExpSumm("iCategory", e.RowIndex).Value, 1).AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd")
                ElseIf DataGridExpSumm("iCategory", e.RowIndex).Value = Month(CDate(v_dtTo)) AndAlso DataGridExpSumm.Columns(e.ColumnIndex).HeaderText = Year(CDate(v_dtTo)) Then 'Last Year: 'No change to v_dtTo
                    'No change to v_dtTo
                    v_dtFrom = New DateTime(DataGridExpSumm.Columns(e.ColumnIndex).HeaderText, DataGridExpSumm("iCategory", e.RowIndex).Value, 1).ToString("yyyy-MM-dd")
                Else
                    v_dtFrom = New DateTime(DataGridExpSumm.Columns(e.ColumnIndex).HeaderText, DataGridExpSumm("iCategory", e.RowIndex).Value, 1).ToString("yyyy-MM-dd")
                    v_dtTo = New DateTime(DataGridExpSumm.Columns(e.ColumnIndex).HeaderText, DataGridExpSumm("iCategory", e.RowIndex).Value, 1).AddMonths(1).AddDays(-1).ToString("yyyy-MM-dd")
                End If
            ElseIf SummaryType = En_SummaryType.Variance Then
                v_dtFrom = New DateTime(DataGridExpSumm.Columns(e.ColumnIndex).HeaderText, 1, 1).ToString("yyyy-MM-dd")
                v_dtTo = New DateTime(DataGridExpSumm.Columns(e.ColumnIndex).HeaderText, 12, 31).ToString("yyyy-MM-dd")
            End If

            If SummaryType = En_SummaryType.RunningTotals Then
                Categories = cmbCatListRunTot.SelectedItem.value
            Else
                Categories = DataGridExpSumm("Category", e.RowIndex).Value
            End If

            dtpFrom = v_dtFrom
            dtpTo = v_dtTo

            CallSearchFunction = True
            Call PopulateSearchResults()

            TabControl1.SelectedTab = TabExpDet

            chkShowAllDet.Visible = True
            chkShowAllDet.Enabled = True
            chkShowAllDet.Checked = False
        End If
    End Sub

    Private Sub DataGridExpSumm_CellMouseEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridExpSumm.CellMouseEnter
        If Panel5.Visible = True Then Exit Sub

        DataGridExpSumm.Cursor = Cursors.Default

        If e.ColumnIndex >= 0 AndAlso e.RowIndex >= 0 Then
            If DataGridExpSumm("Sort", e.RowIndex).Value = 1 Then
                If DataGridExpSumm.Columns(e.ColumnIndex).Index > DataGridExpSumm.Columns("Category").Index AndAlso DataGridExpSumm.Columns(e.ColumnIndex).Index < (DataGridExpSumm.Columns.Count - 1) Then

                    If SearchStr Is Nothing Then
                        If Not SummaryType = En_SummaryType.RunningTotals Then
                            If Not dtBudget.Rows(e.RowIndex)(e.ColumnIndex) Is Nothing Then
                                If DataGridExpSumm(e.ColumnIndex, e.RowIndex).Value > 0 AndAlso dtBudget.Rows(e.RowIndex)(e.ColumnIndex) Is DBNull.Value Then
                                    If DataGridExpSumm("Sort", e.RowIndex).Value = 1 Then
                                        DataGridExpSumm.Cursor = Cursors.Hand
                                        DataGridExpSumm(e.ColumnIndex, e.RowIndex).ToolTipText = "Double click this cell to get the detailed breakdown of this summary." & vbNewLine & "Budgeted Amount: Not Defined"
                                    Else
                                        DataGridExpSumm(e.ColumnIndex, e.RowIndex).ToolTipText = "Budgeted Amount: Not Defined"
                                    End If
                                ElseIf Not dtBudget.Rows(e.RowIndex)(e.ColumnIndex) Is DBNull.Value Then
                                    If DataGridExpSumm(e.ColumnIndex, e.RowIndex).Value > 0 Then
                                        If DataGridExpSumm("Sort", e.RowIndex).Value = 1 Then
                                            DataGridExpSumm.Cursor = Cursors.Hand
                                            DataGridExpSumm(e.ColumnIndex, e.RowIndex).ToolTipText = "Double click this cell to get the detailed breakdown of this summary." & vbNewLine & "Budgeted Amount: " & Format(dtBudget.Rows(e.RowIndex)(e.ColumnIndex), "#,##0.00")
                                        Else
                                            DataGridExpSumm(e.ColumnIndex, e.RowIndex).ToolTipText = "Budgeted Amount: " & Format(dtBudget.Rows(e.RowIndex)(e.ColumnIndex), "#,##0.00")
                                        End If
                                    ElseIf DataGridExpSumm(e.ColumnIndex, e.RowIndex).Value = 0 Then
                                        DataGridExpSumm(e.ColumnIndex, e.RowIndex).ToolTipText = "Budgeted Amount: " & Format(dtBudget.Rows(e.RowIndex)(e.ColumnIndex), "#,##0.00")
                                    End If
                                ElseIf DataGridExpSumm(e.ColumnIndex, e.RowIndex).Value = 0 Then
                                    DataGridExpSumm(e.ColumnIndex, e.RowIndex).ToolTipText = "Budgeted Amount: Not Defined"
                                ElseIf DataGridExpSumm(e.ColumnIndex, e.RowIndex).Value > 0 Then
                                    DataGridExpSumm.Cursor = Cursors.Hand
                                    DataGridExpSumm(e.ColumnIndex, e.RowIndex).ToolTipText = "Double click this cell to get the detailed breakdown of this summary."
                                End If
                            End If
                        Else
                            If DataGridExpSumm(e.ColumnIndex, e.RowIndex).Value > 0 Then
                                DataGridExpSumm.Cursor = Cursors.Hand
                                DataGridExpSumm(e.ColumnIndex, e.RowIndex).ToolTipText = "Double click this cell to get the detailed breakdown of this summary."
                            End If
                        End If
                    Else
                        If DataGridExpSumm(e.ColumnIndex, e.RowIndex).Value > 0 Then
                            DataGridExpSumm.Cursor = Cursors.Hand
                            DataGridExpSumm(e.ColumnIndex, e.RowIndex).ToolTipText = "Double click this cell to get the detailed breakdown of this summary."
                        End If
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub DataGridExpSumm_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DataGridExpSumm.CellBeginEdit
        If Not (SummaryType = En_SummaryType.RunningTotals AndAlso DataGridExpSumm("Sort", e.RowIndex).Value = 4 AndAlso DataGridExpSumm.CurrentCell.ColumnIndex = DataGridExpSumm.Columns("Category").Index) Then
            e.Cancel = True
            Exit Sub
        End If

        If SummaryType = En_SummaryType.RunningTotals Then
            If (DataGridExpSumm.Focused And DataGridExpSumm.CurrentCell.ColumnIndex = DataGridExpSumm.Columns("Category").Index) Then
                cmbMonth.Location = DataGridExpSumm.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, False).Location
                cmbMonth.Size = DataGridExpSumm.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, False).Size
                cmbMonth.Visible = True

                If Not (IsDBNull(DataGridExpSumm.CurrentCell.Value)) Then
                    cmbMonth.SelectedItem = DataGridExpSumm.CurrentCell.Value
                Else
                    cmbMonth.SelectedItem = Nothing
                End If
            Else
                cmbMonth.Visible = False
            End If
        End If
    End Sub

    Private Sub DataGridExpSumm_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridExpSumm.CellEndEdit
        If SummaryType = En_SummaryType.RunningTotals Then
            If (DataGridExpSumm.Focused And DataGridExpSumm.CurrentCell.ColumnIndex = DataGridExpSumm.Columns("Category").Index) Then
                DataGridExpSumm.CurrentCell.Value = CType(cmbMonth.SelectedItem, DataRowView).Row.ItemArray(1)
                cmbMonth.Visible = False
            End If
        End If
    End Sub

    Private Sub DataGridExpSumm_EditingControlShowing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles DataGridExpSumm.EditingControlShowing
        If SummaryType = En_SummaryType.RunningTotals And DataGridExpSumm.CurrentCell.RowIndex = RunTotRowIndex + 1 Then
            If CmbMonthCurrentIndex = -1 Then
                cmbMonth.SelectedIndex = Month("1900-" & RunTotalMonth & "-01") - 1
            Else
                cmbMonth.SelectedIndex = CmbMonthCurrentIndex
            End If
        End If
    End Sub


#End Region

#Region "DataGridThrLimits Events"
    Private Sub DataGridThrLimits_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DataGridThrLimits.CellBeginEdit
        If DataGridThrLimits.CurrentCell.ColumnIndex = DataGridThrLimits.Columns("TAmount").Index Then
            If DataGridThrLimits("IsReadOnly", e.RowIndex).Value Then
                e.Cancel = True
            End If
        End If
    End Sub

    Private Sub DataGridThrLimits_CellEndEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridThrLimits.CellEndEdit
        DgvTextBox = Nothing
    End Sub

    Private Sub DataGridThrLimits_EditingControlShowing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles DataGridThrLimits.EditingControlShowing
        Dgv = DataGridThrLimits

        If Dgv.CurrentCell.ColumnIndex = Dgv.Columns("TAmount").Index Then
            If DgvTextBox Is Nothing And TypeOf e.Control Is DataGridViewTextBoxEditingControl Then
                DgvTextBox = CType(e.Control, DataGridViewTextBoxEditingControl)
            End If
        End If
    End Sub

    Private Sub DataGridThrLimits_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridThrLimits.CellValueChanged
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            DataGridThrLimits(e.ColumnIndex, e.RowIndex).Style.BackColor = Color.LightCyan
        End If
    End Sub

#End Region

#Region "DataGridCatList Events"
    Private Sub DataGridCatList_CellMouseEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridCatList.CellMouseEnter
        If e.ColumnIndex = DataGridCatList.Columns("bDelete").Index AndAlso e.RowIndex >= 0 Then

            If DataGridCatList("CanDelete", e.RowIndex).Value = 0 AndAlso Not DataGridCatList("sCategory", e.RowIndex).Value Is Nothing Then
                DataGridCatList(e.ColumnIndex, e.RowIndex).ToolTipText = "Can't delete this category. There are expenses against this category."
            End If
        End If
    End Sub

    Private Sub DataGridCatList_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridCatList.CellValueChanged
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            DataGridCatList(e.ColumnIndex, e.RowIndex).Style.BackColor = Color.LightCyan
        End If
    End Sub

    Private Sub DataGridCatList_EditingControlShowing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles DataGridCatList.EditingControlShowing
        If DataGridCatList.CurrentCell.ColumnIndex = DataGridCatList.Columns("iMonthlyOccurrences").Index Then
            AddHandler CType(e.Control, TextBox).KeyPress, AddressOf DataGridCatList_MonthlyOccurrences_TextBox_keyPress
        End If
    End Sub
#End Region
#End Region

#Region "TextBox Events"
    Private Sub txtSearch_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSearch.Click

        If ShowSearchAlert Then
            ImpensaAlert("""Search"" function may not show itemised amounts for older data which doesn’t have" & _
                         " notes in the form ""Item1(Amount1), Item2(Amount2),...ItemN(AmountN)"" format. For older data, aggregate amount" & _
                         " of all the items mentioned together will be shown instead.", MsgBoxStyle.OkOnly + MsgBoxStyle.Information)
            ShowSearchAlert = False
        End If

        If txtSearch.Text.Equals(DefaultSearchStr) Then
            txtSearch.Clear()
            txtSearch.ForeColor = Color.Black
        End If
    End Sub

    Private Sub txtSearch_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSearch.LostFocus
        If String.IsNullOrEmpty(Trim(txtSearch.Text)) OrElse txtSearch.Text.Equals(DefaultSearchStr) Then
            txtSearch.Text = DefaultSearchStr
            txtSearch.ForeColor = Color.DarkGray
        End If
    End Sub

    Private Sub txtHighlight_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtHighlight.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtHighlight_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHighlight.TextChanged
        Dim digitsOnly As Regex = New Regex("[^\d]")
        txtHighlight.Text = digitsOnly.Replace(txtHighlight.Text, "")
    End Sub

    Private Sub txtHighlightDet_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtHighlightDet.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtHighlightDet_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHighlightDet.TextChanged
        Dim digitsOnly As Regex = New Regex("[^\d]")
        txtHighlightDet.Text = digitsOnly.Replace(txtHighlightDet.Text, "")
    End Sub

    Private Sub txtHighlightSummMth_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtHighlightSummMth.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtHighlightSummMth_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHighlightSummMth.TextChanged
        Dim digitsOnly As Regex = New Regex("[^\d]")
        txtHighlightSummMth.Text = digitsOnly.Replace(txtHighlightSummMth.Text, "")
    End Sub

    Private Sub txtHighlightSummYr_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles txtHighlightSummYr.KeyPress
        If Not Char.IsNumber(e.KeyChar) AndAlso Not Char.IsControl(e.KeyChar) Then
            e.Handled = True
        End If
    End Sub

    Private Sub txtHighlightSummYr_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtHighlightSummYr.TextChanged
        Dim digitsOnly As Regex = New Regex("[^\d]")
        txtHighlightSummYr.Text = digitsOnly.Replace(txtHighlightSummYr.Text, "")
    End Sub
#End Region

#Region "ComboBox Events"
    Private Sub cmbMonth_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbMonth.SelectionChangeCommitted
        If SummaryType = En_SummaryType.RunningTotals AndAlso DataGridExpSumm.CurrentCell.ColumnIndex = DataGridExpSumm.Columns("Category").Index AndAlso DataGridExpSumm.CurrentCell.RowIndex = RunTotRowIndex + 1 Then
            Dim strSQL As String = ""
            Dim ColTotal As Double = 0
            If Not cmbMonth Is Nothing Then
                CmbMonthCurrentIndex = cmbMonth.SelectedIndex

                For i As Integer = 0 To 11
                    If i = CmbMonthCurrentIndex Then
                        DataGridExpSumm.Rows(i).DefaultCellStyle.ForeColor = Color.Brown
                        DataGridExpSumm.Rows(i).DefaultCellStyle.Font = New System.Drawing.Font(ImpensaFont, FontStyle.Bold)
                    Else
                        DataGridExpSumm.Rows(i).DefaultCellStyle.ForeColor = Color.Black
                        DataGridExpSumm.Rows(i).DefaultCellStyle.Font = ImpensaFont
                    End If
                Next

                Using Connection = GetConnection()
                    strSQL = "Execute sp_GetExpenditureSummary_RunningTotals '" & dtpFrom & "', '" & dtpTo & "', " & cmbCatListRunTot.SelectedItem.key & ", " & (cmbMonth.SelectedIndex + 1) & ", '" & SearchStr & "', 1"
                    dt = New DataTable
                    da = New SqlDataAdapter(strSQL, Connection)
                    da.Fill(dt)
                End Using

                For i As Int32 = 3 To DataGridExpSumm.Columns.Count - 1
                    If Not dt.Rows(0).Item(i) Is DBNull.Value Then
                        ColTotal += dt.Rows(0).Item(i)
                    End If
                    DataGridExpSumm(i, RunTotRowIndex + 1).Value = dt.Rows(0).Item(i)
                Next
                DataGridExpSumm("TOTAL", RunTotRowIndex + 1).Value = ColTotal

            End If
        End If
    End Sub

    Private Sub cmbSelectChart_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbSelectChart.SelectedIndexChanged
        SelectChartCombo = cmbSelectChart.SelectedItem

        btnGo.Enabled = True
        Label8.Visible = True
        chkLBYears.Enabled = True

        If cmbSelectChart.SelectedIndex = -1 Then
            btnGo.Enabled = False
        End If

        cmbListing.Enabled = True
        chkSort.Enabled = True
        If SelectChartCombo = "Chart 1" Then
            chkPeriodLevel.Enabled = False
            chkPeriodLevel.Checked = False
            Label8.Text = "Plot CATEGORY on [Month(X-Axis) Vs. Amount(Y-Axis)]"
            cmbListing.DataSource = mdlChart1.PopulateListingCombo
        ElseIf SelectChartCombo = "Chart 2" Then
            chkPeriodLevel.Enabled = False
            chkPeriodLevel.Checked = False
            chkLBYears.Enabled = False
            Label8.Text = "Plot PERIOD on [Category(X-Axis) Vs. Amount(Y-Axis)]"
            cmbListing.DataSource = mdlChart2.PopulateListingCombo
        ElseIf SelectChartCombo = "Chart 3A" OrElse SelectChartCombo = "Chart 3B" Then
            If Year(dtpFrom) = Year(dtpTo) Then
                chkPeriodLevel.Enabled = False
                chkPeriodLevel.Checked = False
            Else
                chkPeriodLevel.Enabled = True
            End If

            If SelectChartCombo = "Chart 3A" Then
                Label8.Text = "(Multi-Series) Plot CATEGORY on [Month(X-Axis) Vs. Amount(Y-Axis)]"
            Else
                Label8.Text = "(Multi-Series) Plot PERIOD on [Category(X-Axis) Vs. Amount(Y-Axis)]"
            End If

            cmbListing.DataSource = mdlChart3.PopulateListingCombo

            If Not (Year(dtpFrom) = Year(dtpTo)) Then
                chkSort.Checked = False
                chkSort.Enabled = False
            End If

        ElseIf SelectChartCombo = "Chart 4" Then
            If Year(dtpFrom) = Year(dtpTo) Then
                chkPeriodLevel.Enabled = False
                chkPeriodLevel.Checked = False
            Else
                chkPeriodLevel.Enabled = True
            End If
            Label8.Text = "Plot CATEGORY on [Year(X-Axis) Vs. Amount(Y-Axis)]"
            cmbListing.DataSource = mdlChart4.PopulateListingCombo
        End If
        cmbListing.DisplayMember = "Value"
        cmbListing.ValueMember = "Key"
        cmbListing.SelectedIndex = -1
    End Sub

    Private Sub cmbChartType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbChartType.SelectedIndexChanged
        ChartTypeCombo = cmbChartType.SelectedItem
        If cmbChartType.SelectedIndex = -1 Then
            btnGo.Enabled = False
        Else
            btnGo.Enabled = True
        End If
    End Sub

    Private Sub cmbListing_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbListing.SelectedIndexChanged
        ListingCombo = cmbListing.SelectedItem

        If cmbListing.Enabled Then
            If cmbListing.SelectedIndex = -1 Then
                btnGo.Enabled = False
            Else
                btnGo.Enabled = True
            End If
        End If
    End Sub

    Private Sub cmbThrMonth_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbThrMonth.SelectionChangeCommitted
        If cmbThrMonth.SelectedIndex = -1 Then
            DataGridThrLimits.DataSource = Nothing
            DataGridThrLimits.Columns.Clear()
        Else
            ThresholdMonth = cmbThrMonth.SelectedItem.Key
            Call PopulateThresholdGrid()
        End If
    End Sub

    Private Sub cmbSelectYear_DropDown(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbSelectYear.DropDown
        If rbCloseYr.Checked Then
            Call PopulateSelectYearCombo(False)
        ElseIf rbOpenYr.Checked Then
            Call PopulateSelectYearCombo(True)
        End If
    End Sub

    Private Sub cmbSummaryType_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbSummaryType.SelectionChangeCommitted
        SummaryType = cmbSummaryType.SelectedItem.key
        cmbCatListRunTot.Enabled = False
        cmbCatListRunTot.SelectedIndex = -1
        chkLBVarComparision.Enabled = False
        btnVarCompare.Enabled = False
        chkLBVarComparision.Items.Clear()

        If SummaryType = En_SummaryType.Monthly Then
            txtHighlight.Text = HighlightSummMonthly
        ElseIf SummaryType = En_SummaryType.Yearly Then
            txtHighlight.Text = HighlightSummYearly
        ElseIf SummaryType = En_SummaryType.RunningTotals Then
            txtHighlight.Text = HighlightSummMonthly
            cmbCatListRunTot.Enabled = True
            DataGridExpSumm.Columns.Clear()
            DataGridExpSumm.DataSource = Nothing
        ElseIf SummaryType = En_SummaryType.Variance Then
            Call BuildYearsListString()
            txtHighlight.Text = HighlightSummYearly
            DataGridExpSumm.Columns.Clear()
            DataGridExpSumm.DataSource = Nothing
            chkLBVarComparision.Enabled = True

            If chkLBVarComparision.Items.Count >= 2 Then btnVarCompare.Enabled = True
        End If

        If Not (SummaryType = En_SummaryType.RunningTotals OrElse SummaryType = En_SummaryType.Variance) Then
            TabControl1.Enabled = False
            Panel5.BringToFront()
            Panel5.Visible = True
            Call PopulateExpenditureSummaryGrid()
            Call FormatSummaryGrid()
            Panel5.SendToBack()
            Panel5.Visible = False
            TabControl1.Enabled = True
        End If
    End Sub

    Private Sub cmbCatListRunTot_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbCatListRunTot.SelectionChangeCommitted
        Call PopulateExpenditureSummaryGrid()
        Call FormatSummaryGrid()
        CmbMonthCurrentIndex = -1
    End Sub

    Private Sub cmbPeriod_SelectionChangeCommitted(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbPeriod.SelectionChangeCommitted
        If cmbPeriod.SelectedIndex = En_Period.CurrentMonth Then
            dtPickerFrom.Value = Now.AddDays((Now.Day - 1) * -1) 'First Day Of Current Month
        ElseIf cmbPeriod.SelectedIndex = En_Period.CurrentYear Then
            dtPickerFrom.Value = Convert.ToDateTime("01/01/" & Year(Now))
        ElseIf cmbPeriod.SelectedIndex = En_Period.BookStart Then
            dtPickerFrom.Value = RecordKeepingStartDate
        End If

        dtpFrom = dtPickerFrom.Value.Date.ToString("yyyy-MM-dd")
        dtPickerTo.Value = DateTime.Today
        dtpTo = dtPickerTo.Value.Date.ToString("yyyy-MM-dd")
    End Sub

#End Region

#Region "DatePicker Events"
    Private Sub dtPickerFrom_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles dtPickerFrom.Validating
        If dtPickerTo.Value.Date < dtPickerFrom.Value.Date Then
            ImpensaAlert("""From"" Date cannot be greater than ""To"" Date", MsgBoxStyle.Information + MsgBoxStyle.OkOnly)
            e.Cancel = True
        Else
            Call SetPeriodComboIndex()
        End If
    End Sub

    Private Sub dtPickerTo_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles dtPickerTo.Validating
        If dtPickerTo.Value.Date < dtPickerFrom.Value.Date Then
            ImpensaAlert("""From"" Date cannot be greater than ""To"" Date", MsgBoxStyle.Information + MsgBoxStyle.OkOnly)
            e.Cancel = True
        Else
            Call SetPeriodComboIndex()
        End If
    End Sub
#End Region

#Region "LinkLabel Events"
    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked

        If ImpensaActionAlert("Changing connection details while application is running will reset the progress made (if any). Continue?", _
                            MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) = Windows.Forms.DialogResult.OK Then
            frmLogin.ShowDialog()
            If IsLoginDetailsChanged Then
                Call ResetFilters()
                IsLoginDetailsChanged = False
            End If
        End If
    End Sub

    Private Sub LinkLabel2_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        If ShowUnsavedDataWarning("You have unsaved data in the grid(s). Navigating away or refreshing grid may cause you to loose unsaved data. Continue?") = Windows.Forms.DialogResult.Cancel Then
            Exit Sub
        End If
        Call ResetFilters()
    End Sub
#End Region

#Region "CheckBox And Radio Button Events"

    Private Sub rbCloseYr_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbCloseYr.CheckedChanged
        If rbCloseYr.Checked Then
            cmbSelectYear.Enabled = True
        End If
    End Sub

    Private Sub rdOpenYr_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rbOpenYr.CheckedChanged
        If rbOpenYr.Checked Then
            cmbSelectYear.Enabled = True
        End If
    End Sub

    Private Sub chkSort_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles chkSort.CheckedChanged
        If chkSort.Checked Then
            cmbSort.Enabled = True
            SortByAmount = True
        Else
            cmbSort.Enabled = False
            SortByAmount = False
        End If
    End Sub

    Private Sub chkShowReminder_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkShowReminder.CheckedChanged
        If chkShowReminder.Checked Then
            txtReminder.Enabled = True
        Else
            txtReminder.Enabled = False
        End If
    End Sub

    Private Sub chkShowAllDet_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkShowAllDet.CheckedChanged
        If chkShowAllDet.Checked Then
            Call SetDefaultPropValue()
            Call PopulateExpenditureDetailGrid()
        End If
    End Sub
#End Region

#Region "Chart_Analysis Events"
    Private Sub Chart_Analysis_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Chart_Analysis.DoubleClick
        If IsDataAvailable Then
            If Application.OpenForms("frmChart") Is Nothing Then
                frmChart.ShowDialog()
            End If
        End If
    End Sub

    Private Sub Chart_Analysis_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles Chart_Analysis.MouseEnter
        tt = New ToolTip
        tt.SetToolTip(sender, "Double click Chart to open it in new window")
    End Sub

    Private Sub Chart_Analysis_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles Chart_Analysis.MouseLeave
        tt.Dispose()
    End Sub
#End Region

#Region "StatusStrip Events"
    Private Sub tslblMTD_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tslblMTD.Click
        SelectChartCombo = "Chart MTD"
        SortByAmount = True

        Call SetChartColorPalette()

        If Application.OpenForms("frmChart") Is Nothing Then
            frmChart.ShowDialog()
        End If
    End Sub

    Private Sub tslblBKSDTD_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tslblBKSDTD.Click
        SelectChartCombo = "Chart BKSDTD"
        SortByAmount = True

        Call SetChartColorPalette()

        If Application.OpenForms("frmChart") Is Nothing Then
            frmChart.ShowDialog()
        End If
    End Sub

    Private Sub tslblYTD_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tslblYTD.Click
        SelectChartCombo = "Chart YTD"
        SortByAmount = True

        Call SetChartColorPalette()

        If Application.OpenForms("frmChart") Is Nothing Then
            frmChart.ShowDialog()
        End If
    End Sub

    Private Sub tslblMTD_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles tslblMTD.MouseEnter
        tslblMTD.Font = New System.Drawing.Font(ImpensaFont.FontFamily, 10, FontStyle.Bold)
        tt = New ToolTip
        tt.SetToolTip(StatusStrip, "Click to view the graphical analysis of 'Month To Date' expenditures")
    End Sub

    Private Sub tslblMTD_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles tslblMTD.MouseLeave
        tslblMTD.Font = New System.Drawing.Font(ImpensaFont.FontFamily, 10, FontStyle.Regular)
        tt.Dispose()
    End Sub

    Private Sub tslblYTD_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles tslblYTD.MouseEnter
        tslblYTD.Font = New System.Drawing.Font(ImpensaFont.FontFamily, 10, FontStyle.Bold)
        tt = New ToolTip
        tt.SetToolTip(StatusStrip, "Click to view the graphical analysis of 'Year To Date' expenditures")
    End Sub

    Private Sub tslblYTD_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles tslblYTD.MouseLeave
        tslblYTD.Font = New System.Drawing.Font(ImpensaFont.FontFamily, 10, FontStyle.Regular)
        tt.Dispose()
    End Sub

    Private Sub tslblBKSDTD_MouseEnter(ByVal sender As Object, ByVal e As System.EventArgs) Handles tslblBKSDTD.MouseEnter
        tslblBKSDTD.Font = New System.Drawing.Font(ImpensaFont.FontFamily, 10, FontStyle.Bold)
        tt = New ToolTip
        tt.SetToolTip(StatusStrip, "Click to view the graphical analysis of all the expenditures till date")
    End Sub

    Private Sub tslblBKSDTD_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles tslblBKSDTD.MouseLeave
        tslblBKSDTD.Font = New System.Drawing.Font(ImpensaFont.FontFamily, 10, FontStyle.Regular)
        tt.Dispose()
    End Sub

    Private Sub tslblRecdCnt_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tslblRecdCnt.Click
        Call PopulateThresholdGrid()
    End Sub

    Private Sub tsMenuAPCats_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsMenuAPCats.Click
        Call PopulateThresholdGrid(En_BudgetBuckets.AtParCats)
    End Sub

    Private Sub tsMenuOBCats_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsMenuOBCats.Click
        Call PopulateThresholdGrid(En_BudgetBuckets.OverBudgetCats)
    End Sub

    Private Sub tsMenuUBCats_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tsMenuUBCats.Click
        Call PopulateThresholdGrid(En_BudgetBuckets.UnderBudgetCats)
    End Sub
#End Region

#Region "Other Control Events"
    Private Sub TabControl1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles TabControl1.SelectedIndexChanged
        Dim HighlightAmt As Long

        Try
            LastTabIndex = TabControl1.SelectedIndex

            btnSave.Enabled = True
            btnSave.Text = "Save"
            btnGo.Enabled = False
            Panel1.Enabled = False
            tslblGridTotal.Visible = False
            tslblSeperator1.Visible = False
            tslblRecdCnt.Visible = False
            tslblRecdCnt.IsLink = False
            tslblSeperator2.Visible = False
            chkShowAllDet.Visible = False
            pnlHighlight.Visible = False
            tsCmbBudgetBuckets.Visible = False

            If TabControl1.SelectedTab.Name = "TabExpDet" Then
                HighlightAmt = HighlightDetail
                txtHighlight.Text = HighlightAmt
                pnlHighlight.Visible = True
                If chkShowAllDet.Enabled AndAlso CallSearchFunction Then chkShowAllDet.Visible = True
                tslblGridTotal.Visible = True
                tslblSeperator1.Visible = True
                tslblRecdCnt.Text = "Total Records Displayed: #" & dtDetailGrid.Select("hKey IS NOT NULL").ToArray.Count
                tslblRecdCnt.Visible = True
                tslblSeperator2.Visible = True
                Panel1.Enabled = True
            End If

            If TabControl1.SelectedTab.Name = "TabCatList" Then
                For Each dr As DataGridViewRow In DataGridCatList.Rows
                    dr.HeaderCell.Value = ""
                    If dr.Cells("CanDelete").Value = 0 And Not dr.Cells("hKey").Value Is Nothing Then
                        dr.Cells("bDelete").ReadOnly = True
                    End If
                Next
                tslblRecdCnt.Text = "Total Records Displayed: #" & (DataGridCatList.Rows.Count - 1)
                tslblRecdCnt.Visible = True
                tslblSeperator2.Visible = True
            End If

            If TabControl1.SelectedTab.Name = "TabExpSumm" Then
                TabControl1.Enabled = False
                Panel5.BringToFront()
                Panel5.Visible = True
                HighlightAmt = IIf(SummaryType = En_SummaryType.Monthly, HighlightSummMonthly, HighlightSummYearly)
                pnlHighlight.Visible = True
                btnSave.Enabled = False
                btnSave.Text = "Export Grid To PDF"
                tslblRecdCnt.Text = "Total Records Displayed: #" & IIf((DataGridExpSumm.Rows.Count - 1) < 0, 0, (DataGridExpSumm.Rows.Count - 1))
                tslblRecdCnt.Visible = True
                tslblSeperator2.Visible = True
                txtHighlight.Text = HighlightAmt
                Call FormatSummaryGrid()
            End If

            If TabControl1.SelectedTab.Name = "TabAnalysis" Then
                If InStr("Chart MTD, Chart YTD, Chart BKSDTD", SelectChartCombo) > 0 Then SelectChartCombo = Nothing

                Call BuildYearsListString()
                If cmbSelectChart.SelectedIndex <> -1 AndAlso cmbListing.SelectedIndex <> -1 Then
                    btnGo.Enabled = True
                End If

                chkPeriodLevel.Text = "Period: " & CDate(dtpFrom).Day.ToString.PadLeft(2, "0") & "/" & MonthName(Month(dtpFrom)).ToString.Substring(0, 3) & " - " & _
                                                   CDate(dtpTo).Day.ToString.PadLeft(2, "0") & "/" & MonthName(Month(dtpTo)).ToString.Substring(0, 3)
                btnSave.Enabled = False

                If CallSearchFunction = True Then
                    Call SetDefaultPropValue(True)
                End If
            End If

            If TabControl1.SelectedTab.Name = "TabFunctions" Then
                rbOpenYr.Checked = False
                rbCloseYr.Checked = False
                cmbSelectYear.Enabled = False
            End If

            If TabControl1.SelectedTab.Name = "TabThreshold" Then
                Call FormatDataGridThrLimits()
                tslblRecdCnt.Text = "Total Records Displayed: #" & (DataGridThrLimits.Rows.Count - 1)
                tslblRecdCnt.IsLink = True
                tslblRecdCnt.LinkBehavior = LinkBehavior.SystemDefault
                tslblRecdCnt.Visible = True
                tslblSeperator2.Visible = True

                If ThresholdCurrentMonthIndex <> -1 Then
                    tsCmbBudgetBuckets.Visible = True
                End If
            End If
        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        Finally
            Panel5.SendToBack()
            Panel5.Visible = False
            TabControl1.Enabled = True
        End Try
    End Sub

    Private Sub TabControl1_Selecting(ByVal sender As Object, ByVal e As System.Windows.Forms.TabControlCancelEventArgs) Handles TabControl1.Selecting
        If ShowUnsavedDataWarning("You have unsaved data in the grid(s). Navigating away or refreshing grid may cause you to loose unsaved data. Continue?") = Windows.Forms.DialogResult.Cancel Then
            e.Cancel = True
        End If
    End Sub

    Private Sub tmrAlert_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tmrAlert.Tick
        lblAlertText.Left -= 5
        If lblAlertText.Right <= Me.Left Then
            lblAlertText.Left = Me.Width
        End If
    End Sub

    Private Sub tmrTicker_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tmrTicker.Tick
        RchTB_MTDTicker.Left -= 5
        RchTB_YTDTicker.Left -= 5
        If RchTB_MTDTicker.Right <= Label24.Right Then
            RchTB_MTDTicker.Left = Panel9.Width
        End If
        If RchTB_YTDTicker.Right <= Label23.Right Then
            RchTB_YTDTicker.Left = Me.Width
        End If
    End Sub

    Private Sub tmrRefresh_Tick(ByVal sender As Object, ByVal e As System.EventArgs) Handles tmrRefresh.Tick
        Call ShowErrorsAndExceptions()
    End Sub

    Private Sub chkLBYears_ItemCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles chkLBYears.ItemCheck
        If e.NewValue = CheckState.Unchecked AndAlso chkLBYears.CheckedItems.Count = 1 Then
            e.NewValue = CheckState.Checked
        End If
    End Sub

    Private Sub chkLBVarComparision_ItemCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.ItemCheckEventArgs) Handles chkLBVarComparision.ItemCheck
        btnVarCompare.Enabled = True

        If e.NewValue = CheckState.Checked AndAlso chkLBVarComparision.CheckedItems.Count > 2 Then
            e.NewValue = CheckState.Unchecked
        End If

        If e.NewValue = CheckState.Unchecked AndAlso chkLBVarComparision.CheckedItems.Count < 2 Then
            btnVarCompare.Enabled = False
        End If

        If chkLBVarComparision.CheckedItems.Count = 2 Then
            If e.NewValue = CheckState.Checked Then
                e.NewValue = CheckState.Unchecked
            Else
                btnVarCompare.Enabled = False
            End If
        End If
    End Sub

    Private Sub NotifyIcon_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon.DoubleClick
        If Not FirstTimeLoadInBkg Then
            If CDate(dtpTo) < DateTime.Today.Date Then
                Dim UnsavedDataWarning_Result As Windows.Forms.DialogResult = Windows.Forms.DialogResult.None
                UnsavedDataWarning_Result = ShowUnsavedDataWarning("You have unsaved data in the grid(s). Are you sure you want to refresh data?")
                If UnsavedDataWarning_Result = Windows.Forms.DialogResult.OK OrElse UnsavedDataWarning_Result = Windows.Forms.DialogResult.None Then
                    Call frmMain_Load(sender, e)
                End If
            End If
        End If
        If Me.Visible = False Then Me.Visible = True
    End Sub

    Private Sub ContextMenu_ItemClicked(ByVal sender As Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ContextMenu.ItemClicked
        Try
            Select Case e.ClickedItem.ToString()
                Case "Exit"
                    Dim UnsavedDataWarning_Result As Windows.Forms.DialogResult = Windows.Forms.DialogResult.None

                    If Not FirstTimeLoadInBkg Then
                        UnsavedDataWarning_Result = ShowUnsavedDataWarning("You have unsaved data in the grid(s). Are you sure you want to exit without saving?")
                    End If
                    If UnsavedDataWarning_Result = Windows.Forms.DialogResult.None Then
                        If ImpensaActionAlert("Are you sure you want to exit?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                            ExitApplication()
                        End If
                    ElseIf UnsavedDataWarning_Result = Windows.Forms.DialogResult.OK Then
                        ExitApplication()
                    ElseIf UnsavedDataWarning_Result = Windows.Forms.DialogResult.Cancel Then
                        Me.Visible = True
                    End If

            End Select
        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub BGWorker_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles BGWorker.DoWork
        Do While Not File.GetLastWriteTime(CSVBackupPath + "\Impensa.xlsm").ToString = ImportFileTimeStamp
            Call DataImport()
        Loop
    End Sub

    Private Sub BGWorker_RunWorkerCompleted(ByVal sender As Object, ByVal e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles BGWorker.RunWorkerCompleted
        If EnableImport Then
            BGWorker.RunWorkerAsync()
        End If
    End Sub

#End Region

#Region "Custom Event Handlers"
    Private Sub Dgv_ValidateNumericCell(ByVal sender As Object, ByVal e As KeyPressEventArgs) Handles DgvTextBox.KeyPress
        Dim ColIndex As Integer = 0

        If Dgv.Name = DataGridExpDet.Name Then
            ColIndex = Dgv.Columns("Amount").Index
        ElseIf Dgv.Name = DataGridThrLimits.Name Then
            ColIndex = Dgv.Columns("TAmount").Index
        End If

        If Dgv.CurrentCell.ColumnIndex = ColIndex Then
            If Not (Char.IsDigit(CChar(CStr(e.KeyChar))) = True _
                    Or (e.KeyChar = "." And Not (DgvTextBox.Text.Contains(".") Or DgvTextBox.Text.Length = 0)) _
                    Or (Dgv.Name = "DataGridExpDet" And e.KeyChar = "-" And Not (DgvTextBox.Text.Contains("-") Or DgvTextBox.SelectionStart > 0)) _
                    ) Then
                e.Handled = True
            Else
                e.Handled = False
            End If
        End If
    End Sub

    Private Sub DgvComboBox_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DgvComboBox.SelectedIndexChanged
        If Dgv.Name = "DataGridExpDet" Then
            If InStr(DgvComboBox.Text, "*") > 0 Then
                ImpensaAlert("Selected Category is obsolete. Cannot use this Category for new entry.", MsgBoxStyle.Information + MsgBoxStyle.OkOnly)
                DgvComboBox.SelectedIndex = -1
            End If
        End If
    End Sub

    Private Sub DataGridCatList_MonthlyOccurrences_TextBox_keyPress(ByVal sender As Object, ByVal e As KeyPressEventArgs)

        If Char.IsDigit(CChar(CStr(e.KeyChar))) = False Then e.Handled = True

        If Char.IsDigit(CChar(CStr(e.KeyChar))) Then
            If CInt(CStr(e.KeyChar)) = 0 Then
                e.Handled = True
            End If
        End If
    End Sub
#End Region
#End Region
End Class
