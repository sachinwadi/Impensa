<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Dim Legend2 As System.Windows.Forms.DataVisualization.Charting.Legend = New System.Windows.Forms.DataVisualization.Charting.Legend()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.cmbPeriod = New System.Windows.Forms.ComboBox()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.LinkLabel2 = New System.Windows.Forms.LinkLabel()
        Me.txtSearch = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.btnSubmit = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.LstCategory = New System.Windows.Forms.CheckedListBox()
        Me.dtPickerTo = New System.Windows.Forms.DateTimePicker()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.dtPickerFrom = New System.Windows.Forms.DateTimePicker()
        Me.TabCategories = New System.Windows.Forms.TabPage()
        Me.DataGridCatList = New System.Windows.Forms.DataGridView()
        Me.TabCharts = New System.Windows.Forms.TabPage()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.Chart_Analysis = New System.Windows.Forms.DataVisualization.Charting.Chart()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.chkShowLabel = New System.Windows.Forms.CheckBox()
        Me.btnGo = New System.Windows.Forms.Button()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.chkLBYears = New System.Windows.Forms.CheckedListBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.chkPeriodLevel = New System.Windows.Forms.CheckBox()
        Me.cmbSort = New System.Windows.Forms.ComboBox()
        Me.chkSort = New System.Windows.Forms.CheckBox()
        Me.cmbChartType = New System.Windows.Forms.ComboBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.cmbListing = New System.Windows.Forms.ComboBox()
        Me.cmbSelectChart = New System.Windows.Forms.ComboBox()
        Me.TabSummary = New System.Windows.Forms.TabPage()
        Me.Panel8 = New System.Windows.Forms.Panel()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.cmbSummaryType = New System.Windows.Forms.ComboBox()
        Me.btnVarCompare = New System.Windows.Forms.Button()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.chkLBVarComparision = New System.Windows.Forms.CheckedListBox()
        Me.cmbCatListRunTot = New System.Windows.Forms.ComboBox()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.DataGridExpSumm = New System.Windows.Forms.DataGridView()
        Me.TabDetails = New System.Windows.Forms.TabPage()
        Me.LstUnpaidBillsPrevMonth = New System.Windows.Forms.ListBox()
        Me.Label29 = New System.Windows.Forms.Label()
        Me.LstUnpaidBillsCurrentMonth = New System.Windows.Forms.ListBox()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.DataGridExpDet = New System.Windows.Forms.DataGridView()
        Me.Panel5 = New System.Windows.Forms.Panel()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.ImpensaTabControl = New System.Windows.Forms.TabControl()
        Me.TabBudget = New System.Windows.Forms.TabPage()
        Me.Panel7 = New System.Windows.Forms.Panel()
        Me.DataGridThrLimits = New System.Windows.Forms.DataGridView()
        Me.Panel6 = New System.Windows.Forms.Panel()
        Me.cmbThrMonth = New System.Windows.Forms.ComboBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.TabSettings = New System.Windows.Forms.TabPage()
        Me.Label36 = New System.Windows.Forms.Label()
        Me.GroupBox7 = New System.Windows.Forms.GroupBox()
        Me.chkExcelDelRows = New System.Windows.Forms.CheckBox()
        Me.chkStartImport = New System.Windows.Forms.CheckBox()
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel()
        Me.gbEmailConfig = New System.Windows.Forms.GroupBox()
        Me.grpEmailSettings = New System.Windows.Forms.GroupBox()
        Me.chkIncludeExpSummary = New System.Windows.Forms.CheckBox()
        Me.txtEmailPassword = New System.Windows.Forms.TextBox()
        Me.txtEmailTo = New System.Windows.Forms.TextBox()
        Me.txtSmtpPort = New System.Windows.Forms.TextBox()
        Me.txtSmtpHost = New System.Windows.Forms.TextBox()
        Me.txtEmailFrom = New System.Windows.Forms.TextBox()
        Me.Label34 = New System.Windows.Forms.Label()
        Me.Label33 = New System.Windows.Forms.Label()
        Me.Label32 = New System.Windows.Forms.Label()
        Me.Label31 = New System.Windows.Forms.Label()
        Me.Label35 = New System.Windows.Forms.Label()
        Me.chkSendEmails = New System.Windows.Forms.CheckBox()
        Me.Label30 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.txtReminder = New System.Windows.Forms.TextBox()
        Me.chkShowReminder = New System.Windows.Forms.CheckBox()
        Me.GroupBox6 = New System.Windows.Forms.GroupBox()
        Me.Label27 = New System.Windows.Forms.Label()
        Me.btnBrowse = New System.Windows.Forms.Button()
        Me.txtCSVBackupPath = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.GroupBox5 = New System.Windows.Forms.GroupBox()
        Me.dtpRecdKeeping = New System.Windows.Forms.DateTimePicker()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.btnOpenYrAll = New System.Windows.Forms.Button()
        Me.btnCloseYr = New System.Windows.Forms.Button()
        Me.btnOpenYr = New System.Windows.Forms.Button()
        Me.btnCloseYrAll = New System.Windows.Forms.Button()
        Me.lstboxClosedYears = New System.Windows.Forms.ListBox()
        Me.lstboxOpenYears = New System.Windows.Forms.ListBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.txtHighlightSummYr = New System.Windows.Forms.TextBox()
        Me.txtHighlightSummMth = New System.Windows.Forms.TextBox()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.txtHighlightDet = New System.Windows.Forms.TextBox()
        Me.Panel3 = New System.Windows.Forms.Panel()
        Me.tmrAlert = New System.Windows.Forms.Timer(Me.components)
        Me.lblAlertText = New System.Windows.Forms.Label()
        Me.chkShowAllDet = New System.Windows.Forms.CheckBox()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.txtHighlight = New System.Windows.Forms.TextBox()
        Me.pnlHighlight = New System.Windows.Forms.Panel()
        Me.btnHighlight = New System.Windows.Forms.Button()
        Me.StatusStrip = New System.Windows.Forms.StatusStrip()
        Me.tslblGridTotal = New System.Windows.Forms.ToolStripStatusLabel()
        Me.tslblSeperator1 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.tslblRecdCnt = New System.Windows.Forms.ToolStripStatusLabel()
        Me.tsCmbBudgetBuckets = New System.Windows.Forms.ToolStripDropDownButton()
        Me.tsMenuOBCats = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsMenuAPCats = New System.Windows.Forms.ToolStripMenuItem()
        Me.tsMenuUBCats = New System.Windows.Forms.ToolStripMenuItem()
        Me.tslblSeperator2 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.tslblMTD = New System.Windows.Forms.ToolStripStatusLabel()
        Me.tslblSeperator6 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.tslblYTD = New System.Windows.Forms.ToolStripStatusLabel()
        Me.tslblSeperator7 = New System.Windows.Forms.ToolStripStatusLabel()
        Me.tslblBKSDTD = New System.Windows.Forms.ToolStripStatusLabel()
        Me.RchTB_MTDTicker = New System.Windows.Forms.RichTextBox()
        Me.RchTB_YTDTicker = New System.Windows.Forms.RichTextBox()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.tmrTicker = New System.Windows.Forms.Timer(Me.components)
        Me.Panel9 = New System.Windows.Forms.Panel()
        Me.Panel10 = New System.Windows.Forms.Panel()
        Me.BGWorker = New System.ComponentModel.BackgroundWorker()
        Me.NotifyIcon = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.ctxContextMenu = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.tmrRefresh = New System.Windows.Forms.Timer(Me.components)
        Me.btnExport = New System.Windows.Forms.Button()
        Me.BgWorker_Email = New System.ComponentModel.BackgroundWorker()
        Me.Timer_Import = New System.Windows.Forms.Timer(Me.components)
        Me.Panel1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabCategories.SuspendLayout()
        CType(Me.DataGridCatList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabCharts.SuspendLayout()
        Me.Panel4.SuspendLayout()
        CType(Me.Chart_Analysis, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel2.SuspendLayout()
        Me.TabSummary.SuspendLayout()
        Me.Panel8.SuspendLayout()
        CType(Me.DataGridExpSumm, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabDetails.SuspendLayout()
        CType(Me.DataGridExpDet, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel5.SuspendLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ImpensaTabControl.SuspendLayout()
        Me.TabBudget.SuspendLayout()
        Me.Panel7.SuspendLayout()
        CType(Me.DataGridThrLimits, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel6.SuspendLayout()
        Me.TabSettings.SuspendLayout()
        Me.GroupBox7.SuspendLayout()
        Me.gbEmailConfig.SuspendLayout()
        Me.grpEmailSettings.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.pnlHighlight.SuspendLayout()
        Me.StatusStrip.SuspendLayout()
        Me.Panel9.SuspendLayout()
        Me.Panel10.SuspendLayout()
        Me.ctxContextMenu.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.Controls.Add(Me.cmbPeriod)
        Me.Panel1.Controls.Add(Me.Label18)
        Me.Panel1.Controls.Add(Me.LinkLabel2)
        Me.Panel1.Controls.Add(Me.txtSearch)
        Me.Panel1.Controls.Add(Me.Label9)
        Me.Panel1.Controls.Add(Me.PictureBox1)
        Me.Panel1.Controls.Add(Me.btnSubmit)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.LstCategory)
        Me.Panel1.Controls.Add(Me.dtPickerTo)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.dtPickerFrom)
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Margin = New System.Windows.Forms.Padding(0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(308, 697)
        Me.Panel1.TabIndex = 1
        '
        'cmbPeriod
        '
        Me.cmbPeriod.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbPeriod.FormattingEnabled = True
        Me.cmbPeriod.Location = New System.Drawing.Point(24, 171)
        Me.cmbPeriod.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cmbPeriod.Name = "cmbPeriod"
        Me.cmbPeriod.Size = New System.Drawing.Size(260, 28)
        Me.cmbPeriod.TabIndex = 19
        '
        'Label18
        '
        Me.Label18.AutoSize = True
        Me.Label18.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label18.Location = New System.Drawing.Point(57, 145)
        Me.Label18.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(197, 20)
        Me.Label18.TabIndex = 18
        Me.Label18.Text = "Select Period (Preset)"
        '
        'LinkLabel2
        '
        Me.LinkLabel2.AutoSize = True
        Me.LinkLabel2.Location = New System.Drawing.Point(110, 669)
        Me.LinkLabel2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LinkLabel2.Name = "LinkLabel2"
        Me.LinkLabel2.Size = New System.Drawing.Size(99, 20)
        Me.LinkLabel2.TabIndex = 17
        Me.LinkLabel2.TabStop = True
        Me.LinkLabel2.Text = "Reset Filters"
        '
        'txtSearch
        '
        Me.txtSearch.ForeColor = System.Drawing.Color.DarkGray
        Me.txtSearch.Location = New System.Drawing.Point(20, 526)
        Me.txtSearch.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtSearch.MaxLength = 500
        Me.txtSearch.Multiline = True
        Me.txtSearch.Name = "txtSearch"
        Me.txtSearch.Size = New System.Drawing.Size(265, 59)
        Me.txtSearch.TabIndex = 15
        Me.txtSearch.Text = "Multiple Keywords Seperated By Comma"
        '
        'Label9
        '
        Me.Label9.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(15, 491)
        Me.Label9.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(267, 35)
        Me.Label9.TabIndex = 14
        Me.Label9.Text = "Search Item(s) In Notes"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'PictureBox1
        '
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(0, -2)
        Me.PictureBox1.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(308, 137)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 12
        Me.PictureBox1.TabStop = False
        '
        'btnSubmit
        '
        Me.btnSubmit.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSubmit.Location = New System.Drawing.Point(20, 605)
        Me.btnSubmit.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.btnSubmit.Name = "btnSubmit"
        Me.btnSubmit.Size = New System.Drawing.Size(267, 51)
        Me.btnSubmit.TabIndex = 3
        Me.btnSubmit.Text = "Submit"
        Me.btnSubmit.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(20, 306)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(267, 35)
        Me.Label4.TabIndex = 7
        Me.Label4.Text = "Select Categories"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'LstCategory
        '
        Me.LstCategory.FormattingEnabled = True
        Me.LstCategory.Location = New System.Drawing.Point(20, 346)
        Me.LstCategory.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.LstCategory.Name = "LstCategory"
        Me.LstCategory.Size = New System.Drawing.Size(265, 142)
        Me.LstCategory.Sorted = True
        Me.LstCategory.TabIndex = 2
        Me.LstCategory.ThreeDCheckBoxes = True
        '
        'dtPickerTo
        '
        Me.dtPickerTo.CustomFormat = "dd/MM/yyyy"
        Me.dtPickerTo.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtPickerTo.Location = New System.Drawing.Point(136, 268)
        Me.dtPickerTo.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.dtPickerTo.MaxDate = New Date(2100, 1, 1, 0, 0, 0, 0)
        Me.dtPickerTo.MinDate = New Date(1900, 1, 1, 0, 0, 0, 0)
        Me.dtPickerTo.Name = "dtPickerTo"
        Me.dtPickerTo.Size = New System.Drawing.Size(148, 26)
        Me.dtPickerTo.TabIndex = 1
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(15, 272)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(112, 22)
        Me.Label3.TabIndex = 3
        Me.Label3.Text = "To Date"
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(15, 225)
        Me.Label1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(112, 22)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "From Date"
        '
        'dtPickerFrom
        '
        Me.dtPickerFrom.CustomFormat = "dd/MM/yyyy"
        Me.dtPickerFrom.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtPickerFrom.Location = New System.Drawing.Point(136, 218)
        Me.dtPickerFrom.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.dtPickerFrom.MaxDate = New Date(2100, 1, 1, 0, 0, 0, 0)
        Me.dtPickerFrom.MinDate = New Date(1900, 1, 1, 0, 0, 0, 0)
        Me.dtPickerFrom.Name = "dtPickerFrom"
        Me.dtPickerFrom.Size = New System.Drawing.Size(148, 26)
        Me.dtPickerFrom.TabIndex = 0
        '
        'TabCategories
        '
        Me.TabCategories.BackColor = System.Drawing.Color.Transparent
        Me.TabCategories.Controls.Add(Me.DataGridCatList)
        Me.TabCategories.Location = New System.Drawing.Point(4, 29)
        Me.TabCategories.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TabCategories.Name = "TabCategories"
        Me.TabCategories.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TabCategories.Size = New System.Drawing.Size(1550, 662)
        Me.TabCategories.TabIndex = 5
        Me.TabCategories.Text = "Categories"
        '
        'DataGridCatList
        '
        Me.DataGridCatList.AllowUserToResizeColumns = False
        Me.DataGridCatList.AllowUserToResizeRows = False
        Me.DataGridCatList.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.DataGridCatList.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridCatList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridCatList.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.DataGridCatList.Location = New System.Drawing.Point(4, 5)
        Me.DataGridCatList.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.DataGridCatList.Name = "DataGridCatList"
        Me.DataGridCatList.RowHeadersWidth = 62
        Me.DataGridCatList.Size = New System.Drawing.Size(1542, 652)
        Me.DataGridCatList.TabIndex = 4
        Me.DataGridCatList.TabStop = False
        '
        'TabCharts
        '
        Me.TabCharts.Controls.Add(Me.Panel4)
        Me.TabCharts.Controls.Add(Me.Panel2)
        Me.TabCharts.Location = New System.Drawing.Point(4, 29)
        Me.TabCharts.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TabCharts.Name = "TabCharts"
        Me.TabCharts.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TabCharts.Size = New System.Drawing.Size(1550, 662)
        Me.TabCharts.TabIndex = 2
        Me.TabCharts.Text = "Graphical Analysis"
        Me.TabCharts.UseVisualStyleBackColor = True
        '
        'Panel4
        '
        Me.Panel4.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel4.Controls.Add(Me.Chart_Analysis)
        Me.Panel4.Location = New System.Drawing.Point(4, 260)
        Me.Panel4.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(1538, 391)
        Me.Panel4.TabIndex = 2
        '
        'Chart_Analysis
        '
        Me.Chart_Analysis.BorderlineColor = System.Drawing.Color.Black
        Me.Chart_Analysis.BorderlineDashStyle = System.Windows.Forms.DataVisualization.Charting.ChartDashStyle.Solid
        Me.Chart_Analysis.Dock = System.Windows.Forms.DockStyle.Fill
        Legend2.Name = "Legend1"
        Legend2.TextWrapThreshold = 0
        Me.Chart_Analysis.Legends.Add(Legend2)
        Me.Chart_Analysis.Location = New System.Drawing.Point(0, 0)
        Me.Chart_Analysis.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Chart_Analysis.Name = "Chart_Analysis"
        Me.Chart_Analysis.Size = New System.Drawing.Size(1538, 391)
        Me.Chart_Analysis.TabIndex = 0
        Me.Chart_Analysis.Text = "Chart1"
        '
        'Panel2
        '
        Me.Panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel2.Controls.Add(Me.chkShowLabel)
        Me.Panel2.Controls.Add(Me.btnGo)
        Me.Panel2.Controls.Add(Me.Label25)
        Me.Panel2.Controls.Add(Me.chkLBYears)
        Me.Panel2.Controls.Add(Me.Label8)
        Me.Panel2.Controls.Add(Me.chkPeriodLevel)
        Me.Panel2.Controls.Add(Me.cmbSort)
        Me.Panel2.Controls.Add(Me.chkSort)
        Me.Panel2.Controls.Add(Me.cmbChartType)
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Controls.Add(Me.Label6)
        Me.Panel2.Controls.Add(Me.Label5)
        Me.Panel2.Controls.Add(Me.cmbListing)
        Me.Panel2.Controls.Add(Me.cmbSelectChart)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(4, 5)
        Me.Panel2.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(1542, 245)
        Me.Panel2.TabIndex = 1
        '
        'chkShowLabel
        '
        Me.chkShowLabel.AutoSize = True
        Me.chkShowLabel.Checked = True
        Me.chkShowLabel.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowLabel.Location = New System.Drawing.Point(231, 198)
        Me.chkShowLabel.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.chkShowLabel.Name = "chkShowLabel"
        Me.chkShowLabel.Size = New System.Drawing.Size(157, 24)
        Me.chkShowLabel.TabIndex = 16
        Me.chkShowLabel.Text = "Show Data Label"
        Me.chkShowLabel.UseVisualStyleBackColor = True
        '
        'btnGo
        '
        Me.btnGo.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnGo.Location = New System.Drawing.Point(392, 194)
        Me.btnGo.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.btnGo.Name = "btnGo"
        Me.btnGo.Size = New System.Drawing.Size(93, 35)
        Me.btnGo.TabIndex = 2
        Me.btnGo.Text = "GO"
        Me.btnGo.UseVisualStyleBackColor = True
        '
        'Label25
        '
        Me.Label25.AutoSize = True
        Me.Label25.Location = New System.Drawing.Point(52, 131)
        Me.Label25.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(61, 20)
        Me.Label25.TabIndex = 15
        Me.Label25.Text = "Year(s)"
        '
        'chkLBYears
        '
        Me.chkLBYears.FormattingEnabled = True
        Me.chkLBYears.Location = New System.Drawing.Point(118, 131)
        Me.chkLBYears.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.chkLBYears.Name = "chkLBYears"
        Me.chkLBYears.Size = New System.Drawing.Size(102, 96)
        Me.chkLBYears.TabIndex = 14
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.ForeColor = System.Drawing.Color.Navy
        Me.Label8.Location = New System.Drawing.Point(492, 17)
        Me.Label8.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(57, 20)
        Me.Label8.TabIndex = 13
        Me.Label8.Text = "Label8"
        Me.Label8.Visible = False
        '
        'chkPeriodLevel
        '
        Me.chkPeriodLevel.AutoSize = True
        Me.chkPeriodLevel.Enabled = False
        Me.chkPeriodLevel.Location = New System.Drawing.Point(231, 131)
        Me.chkPeriodLevel.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.chkPeriodLevel.Name = "chkPeriodLevel"
        Me.chkPeriodLevel.Size = New System.Drawing.Size(84, 24)
        Me.chkPeriodLevel.TabIndex = 12
        Me.chkPeriodLevel.Text = "Period:"
        Me.chkPeriodLevel.UseVisualStyleBackColor = True
        '
        'cmbSort
        '
        Me.cmbSort.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbSort.Enabled = False
        Me.cmbSort.FormattingEnabled = True
        Me.cmbSort.Location = New System.Drawing.Point(392, 158)
        Me.cmbSort.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cmbSort.Name = "cmbSort"
        Me.cmbSort.Size = New System.Drawing.Size(91, 28)
        Me.cmbSort.TabIndex = 10
        '
        'chkSort
        '
        Me.chkSort.AutoSize = True
        Me.chkSort.Location = New System.Drawing.Point(231, 162)
        Me.chkSort.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.chkSort.Name = "chkSort"
        Me.chkSort.Size = New System.Drawing.Size(147, 24)
        Me.chkSort.TabIndex = 9
        Me.chkSort.Text = "Sort By Amount"
        Me.chkSort.UseVisualStyleBackColor = True
        '
        'cmbChartType
        '
        Me.cmbChartType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbChartType.FormattingEnabled = True
        Me.cmbChartType.Location = New System.Drawing.Point(118, 89)
        Me.cmbChartType.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cmbChartType.Name = "cmbChartType"
        Me.cmbChartType.Size = New System.Drawing.Size(364, 28)
        Me.cmbChartType.TabIndex = 7
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(24, 95)
        Me.Label2.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(86, 20)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Chart Type"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(57, 55)
        Me.Label6.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(55, 20)
        Me.Label6.TabIndex = 4
        Me.Label6.Text = "Listing"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(64, 17)
        Me.Label5.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(48, 20)
        Me.Label5.TabIndex = 3
        Me.Label5.Text = "Chart"
        '
        'cmbListing
        '
        Me.cmbListing.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbListing.FormattingEnabled = True
        Me.cmbListing.Location = New System.Drawing.Point(118, 49)
        Me.cmbListing.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cmbListing.Name = "cmbListing"
        Me.cmbListing.Size = New System.Drawing.Size(364, 28)
        Me.cmbListing.TabIndex = 1
        '
        'cmbSelectChart
        '
        Me.cmbSelectChart.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbSelectChart.FormattingEnabled = True
        Me.cmbSelectChart.Location = New System.Drawing.Point(118, 11)
        Me.cmbSelectChart.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cmbSelectChart.Name = "cmbSelectChart"
        Me.cmbSelectChart.Size = New System.Drawing.Size(364, 28)
        Me.cmbSelectChart.TabIndex = 0
        '
        'TabSummary
        '
        Me.TabSummary.AutoScroll = True
        Me.TabSummary.Controls.Add(Me.Panel8)
        Me.TabSummary.Controls.Add(Me.DataGridExpSumm)
        Me.TabSummary.Location = New System.Drawing.Point(4, 29)
        Me.TabSummary.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TabSummary.Name = "TabSummary"
        Me.TabSummary.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TabSummary.Size = New System.Drawing.Size(1550, 662)
        Me.TabSummary.TabIndex = 1
        Me.TabSummary.Text = "Expenditure Summary"
        Me.TabSummary.UseVisualStyleBackColor = True
        '
        'Panel8
        '
        Me.Panel8.Controls.Add(Me.Label19)
        Me.Panel8.Controls.Add(Me.cmbSummaryType)
        Me.Panel8.Controls.Add(Me.btnVarCompare)
        Me.Panel8.Controls.Add(Me.Label26)
        Me.Panel8.Controls.Add(Me.chkLBVarComparision)
        Me.Panel8.Controls.Add(Me.cmbCatListRunTot)
        Me.Panel8.Controls.Add(Me.Label22)
        Me.Panel8.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel8.Location = New System.Drawing.Point(4, 5)
        Me.Panel8.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Panel8.Name = "Panel8"
        Me.Panel8.Size = New System.Drawing.Size(1542, 78)
        Me.Panel8.TabIndex = 1
        '
        'Label19
        '
        Me.Label19.AutoSize = True
        Me.Label19.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label19.Location = New System.Drawing.Point(4, 26)
        Me.Label19.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(126, 20)
        Me.Label19.TabIndex = 0
        Me.Label19.Text = "Summary Type:"
        '
        'cmbSummaryType
        '
        Me.cmbSummaryType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbSummaryType.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbSummaryType.FormattingEnabled = True
        Me.cmbSummaryType.Location = New System.Drawing.Point(134, 20)
        Me.cmbSummaryType.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cmbSummaryType.Name = "cmbSummaryType"
        Me.cmbSummaryType.Size = New System.Drawing.Size(242, 28)
        Me.cmbSummaryType.TabIndex = 1
        '
        'btnVarCompare
        '
        Me.btnVarCompare.AutoSize = True
        Me.btnVarCompare.Enabled = False
        Me.btnVarCompare.Location = New System.Drawing.Point(1104, 23)
        Me.btnVarCompare.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.btnVarCompare.Name = "btnVarCompare"
        Me.btnVarCompare.Size = New System.Drawing.Size(62, 46)
        Me.btnVarCompare.TabIndex = 5
        Me.btnVarCompare.Text = "Go"
        Me.btnVarCompare.UseVisualStyleBackColor = True
        '
        'Label26
        '
        Me.Label26.AutoSize = True
        Me.Label26.Location = New System.Drawing.Point(810, 15)
        Me.Label26.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(174, 40)
        Me.Label26.TabIndex = 4
        Me.Label26.Text = "Select Years (Any Two)" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "[Variance Comparision]"
        Me.Label26.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'chkLBVarComparision
        '
        Me.chkLBVarComparision.Enabled = False
        Me.chkLBVarComparision.FormattingEnabled = True
        Me.chkLBVarComparision.Location = New System.Drawing.Point(994, 14)
        Me.chkLBVarComparision.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.chkLBVarComparision.Name = "chkLBVarComparision"
        Me.chkLBVarComparision.Size = New System.Drawing.Size(98, 50)
        Me.chkLBVarComparision.TabIndex = 3
        '
        'cmbCatListRunTot
        '
        Me.cmbCatListRunTot.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbCatListRunTot.Enabled = False
        Me.cmbCatListRunTot.FormattingEnabled = True
        Me.cmbCatListRunTot.Location = New System.Drawing.Point(486, 20)
        Me.cmbCatListRunTot.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cmbCatListRunTot.Name = "cmbCatListRunTot"
        Me.cmbCatListRunTot.Size = New System.Drawing.Size(298, 28)
        Me.cmbCatListRunTot.TabIndex = 2
        '
        'Label22
        '
        Me.Label22.AutoSize = True
        Me.Label22.Location = New System.Drawing.Point(402, 26)
        Me.Label22.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(77, 20)
        Me.Label22.TabIndex = 0
        Me.Label22.Text = "Category:"
        '
        'DataGridExpSumm
        '
        Me.DataGridExpSumm.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DataGridExpSumm.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridExpSumm.Location = New System.Drawing.Point(4, 82)
        Me.DataGridExpSumm.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.DataGridExpSumm.Name = "DataGridExpSumm"
        Me.DataGridExpSumm.RowHeadersWidth = 62
        Me.DataGridExpSumm.Size = New System.Drawing.Size(1538, 566)
        Me.DataGridExpSumm.TabIndex = 0
        Me.DataGridExpSumm.VirtualMode = True
        '
        'TabDetails
        '
        Me.TabDetails.AutoScroll = True
        Me.TabDetails.Controls.Add(Me.LstUnpaidBillsPrevMonth)
        Me.TabDetails.Controls.Add(Me.Label29)
        Me.TabDetails.Controls.Add(Me.LstUnpaidBillsCurrentMonth)
        Me.TabDetails.Controls.Add(Me.Label28)
        Me.TabDetails.Controls.Add(Me.DataGridExpDet)
        Me.TabDetails.Controls.Add(Me.Panel5)
        Me.TabDetails.Location = New System.Drawing.Point(4, 29)
        Me.TabDetails.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TabDetails.Name = "TabDetails"
        Me.TabDetails.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TabDetails.Size = New System.Drawing.Size(1550, 662)
        Me.TabDetails.TabIndex = 0
        Me.TabDetails.Text = "Expenditure Details"
        Me.TabDetails.UseVisualStyleBackColor = True
        '
        'LstUnpaidBillsPrevMonth
        '
        Me.LstUnpaidBillsPrevMonth.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.LstUnpaidBillsPrevMonth.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.LstUnpaidBillsPrevMonth.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LstUnpaidBillsPrevMonth.ForeColor = System.Drawing.Color.Red
        Me.LstUnpaidBillsPrevMonth.HorizontalScrollbar = True
        Me.LstUnpaidBillsPrevMonth.ItemHeight = 25
        Me.LstUnpaidBillsPrevMonth.Location = New System.Drawing.Point(1254, 317)
        Me.LstUnpaidBillsPrevMonth.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.LstUnpaidBillsPrevMonth.Name = "LstUnpaidBillsPrevMonth"
        Me.LstUnpaidBillsPrevMonth.SelectionMode = System.Windows.Forms.SelectionMode.None
        Me.LstUnpaidBillsPrevMonth.Size = New System.Drawing.Size(244, 202)
        Me.LstUnpaidBillsPrevMonth.TabIndex = 6
        Me.LstUnpaidBillsPrevMonth.TabStop = False
        '
        'Label29
        '
        Me.Label29.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label29.AutoSize = True
        Me.Label29.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label29.Location = New System.Drawing.Point(1246, 283)
        Me.Label29.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(264, 25)
        Me.Label29.TabIndex = 5
        Me.Label29.Text = "Prev. Month's Unpaid Bills"
        '
        'LstUnpaidBillsCurrentMonth
        '
        Me.LstUnpaidBillsCurrentMonth.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.LstUnpaidBillsCurrentMonth.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.LstUnpaidBillsCurrentMonth.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.LstUnpaidBillsCurrentMonth.ForeColor = System.Drawing.Color.Red
        Me.LstUnpaidBillsCurrentMonth.HorizontalScrollbar = True
        Me.LstUnpaidBillsCurrentMonth.ItemHeight = 25
        Me.LstUnpaidBillsCurrentMonth.Location = New System.Drawing.Point(1254, 35)
        Me.LstUnpaidBillsCurrentMonth.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.LstUnpaidBillsCurrentMonth.Name = "LstUnpaidBillsCurrentMonth"
        Me.LstUnpaidBillsCurrentMonth.SelectionMode = System.Windows.Forms.SelectionMode.None
        Me.LstUnpaidBillsCurrentMonth.Size = New System.Drawing.Size(244, 202)
        Me.LstUnpaidBillsCurrentMonth.TabIndex = 4
        Me.LstUnpaidBillsCurrentMonth.TabStop = False
        '
        'Label28
        '
        Me.Label28.Anchor = System.Windows.Forms.AnchorStyles.Top
        Me.Label28.AutoSize = True
        Me.Label28.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label28.Location = New System.Drawing.Point(1246, 5)
        Me.Label28.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(286, 25)
        Me.Label28.TabIndex = 3
        Me.Label28.Text = "Current Month's Unpaid Bills"
        '
        'DataGridExpDet
        '
        Me.DataGridExpDet.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridExpDet.Dock = System.Windows.Forms.DockStyle.Left
        Me.DataGridExpDet.Location = New System.Drawing.Point(4, 5)
        Me.DataGridExpDet.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.DataGridExpDet.Name = "DataGridExpDet"
        Me.DataGridExpDet.RowHeadersWidth = 62
        Me.DataGridExpDet.Size = New System.Drawing.Size(1233, 652)
        Me.DataGridExpDet.TabIndex = 2
        Me.DataGridExpDet.VirtualMode = True
        '
        'Panel5
        '
        Me.Panel5.Controls.Add(Me.Label15)
        Me.Panel5.Controls.Add(Me.Label14)
        Me.Panel5.Controls.Add(Me.PictureBox2)
        Me.Panel5.Location = New System.Drawing.Point(622, 317)
        Me.Panel5.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Panel5.Name = "Panel5"
        Me.Panel5.Size = New System.Drawing.Size(366, 102)
        Me.Panel5.TabIndex = 1
        Me.Panel5.Visible = False
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(92, 60)
        Me.Label15.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(66, 20)
        Me.Label15.TabIndex = 2
        Me.Label15.Text = "Label15"
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.Location = New System.Drawing.Point(86, 25)
        Me.Label14.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(208, 33)
        Me.Label14.TabIndex = 1
        Me.Label14.Text = "Please Wait..."
        '
        'PictureBox2
        '
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(6, 14)
        Me.PictureBox2.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(69, 77)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PictureBox2.TabIndex = 0
        Me.PictureBox2.TabStop = False
        '
        'btnSave
        '
        Me.btnSave.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnSave.Enabled = False
        Me.btnSave.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSave.Location = New System.Drawing.Point(314, 711)
        Me.btnSave.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(248, 35)
        Me.btnSave.TabIndex = 5
        Me.btnSave.Text = "Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'ImpensaTabControl
        '
        Me.ImpensaTabControl.Controls.Add(Me.TabDetails)
        Me.ImpensaTabControl.Controls.Add(Me.TabSummary)
        Me.ImpensaTabControl.Controls.Add(Me.TabCharts)
        Me.ImpensaTabControl.Controls.Add(Me.TabBudget)
        Me.ImpensaTabControl.Controls.Add(Me.TabCategories)
        Me.ImpensaTabControl.Controls.Add(Me.TabSettings)
        Me.ImpensaTabControl.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ImpensaTabControl.Location = New System.Drawing.Point(0, 0)
        Me.ImpensaTabControl.Margin = New System.Windows.Forms.Padding(0)
        Me.ImpensaTabControl.Name = "ImpensaTabControl"
        Me.ImpensaTabControl.SelectedIndex = 0
        Me.ImpensaTabControl.Size = New System.Drawing.Size(1558, 695)
        Me.ImpensaTabControl.TabIndex = 4
        '
        'TabBudget
        '
        Me.TabBudget.Controls.Add(Me.Panel7)
        Me.TabBudget.Controls.Add(Me.Panel6)
        Me.TabBudget.Location = New System.Drawing.Point(4, 29)
        Me.TabBudget.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TabBudget.Name = "TabBudget"
        Me.TabBudget.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TabBudget.Size = New System.Drawing.Size(1550, 662)
        Me.TabBudget.TabIndex = 7
        Me.TabBudget.Text = "Monthly Forecast"
        Me.TabBudget.UseVisualStyleBackColor = True
        '
        'Panel7
        '
        Me.Panel7.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel7.Controls.Add(Me.DataGridThrLimits)
        Me.Panel7.Location = New System.Drawing.Point(4, 75)
        Me.Panel7.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Panel7.Name = "Panel7"
        Me.Panel7.Size = New System.Drawing.Size(1538, 575)
        Me.Panel7.TabIndex = 1
        '
        'DataGridThrLimits
        '
        Me.DataGridThrLimits.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridThrLimits.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridThrLimits.Location = New System.Drawing.Point(0, 0)
        Me.DataGridThrLimits.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.DataGridThrLimits.Name = "DataGridThrLimits"
        Me.DataGridThrLimits.RowHeadersWidth = 62
        Me.DataGridThrLimits.Size = New System.Drawing.Size(1538, 575)
        Me.DataGridThrLimits.TabIndex = 0
        '
        'Panel6
        '
        Me.Panel6.Controls.Add(Me.cmbThrMonth)
        Me.Panel6.Controls.Add(Me.Label17)
        Me.Panel6.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel6.Location = New System.Drawing.Point(4, 5)
        Me.Panel6.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Panel6.Name = "Panel6"
        Me.Panel6.Size = New System.Drawing.Size(1542, 62)
        Me.Panel6.TabIndex = 0
        '
        'cmbThrMonth
        '
        Me.cmbThrMonth.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbThrMonth.FormattingEnabled = True
        Me.cmbThrMonth.Location = New System.Drawing.Point(70, 15)
        Me.cmbThrMonth.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.cmbThrMonth.Name = "cmbThrMonth"
        Me.cmbThrMonth.Size = New System.Drawing.Size(298, 28)
        Me.cmbThrMonth.TabIndex = 2
        '
        'Label17
        '
        Me.Label17.AutoSize = True
        Me.Label17.Location = New System.Drawing.Point(6, 22)
        Me.Label17.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(54, 20)
        Me.Label17.TabIndex = 0
        Me.Label17.Text = "Month"
        '
        'TabSettings
        '
        Me.TabSettings.Controls.Add(Me.Label36)
        Me.TabSettings.Controls.Add(Me.GroupBox7)
        Me.TabSettings.Controls.Add(Me.LinkLabel1)
        Me.TabSettings.Controls.Add(Me.gbEmailConfig)
        Me.TabSettings.Controls.Add(Me.Label30)
        Me.TabSettings.Controls.Add(Me.GroupBox1)
        Me.TabSettings.Controls.Add(Me.GroupBox6)
        Me.TabSettings.Controls.Add(Me.GroupBox5)
        Me.TabSettings.Controls.Add(Me.GroupBox3)
        Me.TabSettings.Controls.Add(Me.GroupBox2)
        Me.TabSettings.Location = New System.Drawing.Point(4, 29)
        Me.TabSettings.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TabSettings.Name = "TabSettings"
        Me.TabSettings.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.TabSettings.Size = New System.Drawing.Size(1550, 662)
        Me.TabSettings.TabIndex = 6
        Me.TabSettings.Text = "Settings"
        Me.TabSettings.UseVisualStyleBackColor = True
        '
        'Label36
        '
        Me.Label36.AutoSize = True
        Me.Label36.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label36.Location = New System.Drawing.Point(258, 18)
        Me.Label36.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label36.Name = "Label36"
        Me.Label36.Size = New System.Drawing.Size(87, 20)
        Me.Label36.TabIndex = 11
        Me.Label36.Text = "Db Name"
        '
        'GroupBox7
        '
        Me.GroupBox7.Controls.Add(Me.chkExcelDelRows)
        Me.GroupBox7.Controls.Add(Me.chkStartImport)
        Me.GroupBox7.Location = New System.Drawing.Point(9, 49)
        Me.GroupBox7.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox7.Name = "GroupBox7"
        Me.GroupBox7.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox7.Size = New System.Drawing.Size(496, 94)
        Me.GroupBox7.TabIndex = 10
        Me.GroupBox7.TabStop = False
        Me.GroupBox7.Text = "Import Service Options"
        '
        'chkExcelDelRows
        '
        Me.chkExcelDelRows.AutoSize = True
        Me.chkExcelDelRows.Location = New System.Drawing.Point(6, 55)
        Me.chkExcelDelRows.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.chkExcelDelRows.Name = "chkExcelDelRows"
        Me.chkExcelDelRows.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkExcelDelRows.Size = New System.Drawing.Size(294, 24)
        Me.chkExcelDelRows.TabIndex = 2
        Me.chkExcelDelRows.Text = "Delete Old Records From Import File"
        Me.chkExcelDelRows.UseVisualStyleBackColor = True
        '
        'chkStartImport
        '
        Me.chkStartImport.AutoSize = True
        Me.chkStartImport.Location = New System.Drawing.Point(4, 29)
        Me.chkStartImport.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.chkStartImport.Name = "chkStartImport"
        Me.chkStartImport.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkStartImport.Size = New System.Drawing.Size(296, 24)
        Me.chkStartImport.TabIndex = 1
        Me.chkStartImport.Text = "Enable Impensa Data Import Service"
        Me.chkStartImport.UseVisualStyleBackColor = True
        '
        'LinkLabel1
        '
        Me.LinkLabel1.AutoSize = True
        Me.LinkLabel1.Location = New System.Drawing.Point(14, 18)
        Me.LinkLabel1.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(235, 20)
        Me.LinkLabel1.TabIndex = 0
        Me.LinkLabel1.TabStop = True
        Me.LinkLabel1.Text = "Change Database Login Details"
        '
        'gbEmailConfig
        '
        Me.gbEmailConfig.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.gbEmailConfig.Controls.Add(Me.grpEmailSettings)
        Me.gbEmailConfig.Controls.Add(Me.chkSendEmails)
        Me.gbEmailConfig.Location = New System.Drawing.Point(514, 368)
        Me.gbEmailConfig.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.gbEmailConfig.Name = "gbEmailConfig"
        Me.gbEmailConfig.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.gbEmailConfig.Size = New System.Drawing.Size(634, 277)
        Me.gbEmailConfig.TabIndex = 9
        Me.gbEmailConfig.TabStop = False
        '
        'grpEmailSettings
        '
        Me.grpEmailSettings.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpEmailSettings.Controls.Add(Me.chkIncludeExpSummary)
        Me.grpEmailSettings.Controls.Add(Me.txtEmailPassword)
        Me.grpEmailSettings.Controls.Add(Me.txtEmailTo)
        Me.grpEmailSettings.Controls.Add(Me.txtSmtpPort)
        Me.grpEmailSettings.Controls.Add(Me.txtSmtpHost)
        Me.grpEmailSettings.Controls.Add(Me.txtEmailFrom)
        Me.grpEmailSettings.Controls.Add(Me.Label34)
        Me.grpEmailSettings.Controls.Add(Me.Label33)
        Me.grpEmailSettings.Controls.Add(Me.Label32)
        Me.grpEmailSettings.Controls.Add(Me.Label31)
        Me.grpEmailSettings.Controls.Add(Me.Label35)
        Me.grpEmailSettings.Location = New System.Drawing.Point(20, 38)
        Me.grpEmailSettings.Margin = New System.Windows.Forms.Padding(0)
        Me.grpEmailSettings.Name = "grpEmailSettings"
        Me.grpEmailSettings.Padding = New System.Windows.Forms.Padding(0)
        Me.grpEmailSettings.Size = New System.Drawing.Size(606, 229)
        Me.grpEmailSettings.TabIndex = 1
        Me.grpEmailSettings.TabStop = False
        '
        'chkIncludeExpSummary
        '
        Me.chkIncludeExpSummary.AutoSize = True
        Me.chkIncludeExpSummary.Location = New System.Drawing.Point(117, 15)
        Me.chkIncludeExpSummary.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.chkIncludeExpSummary.Name = "chkIncludeExpSummary"
        Me.chkIncludeExpSummary.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkIncludeExpSummary.Size = New System.Drawing.Size(224, 24)
        Me.chkIncludeExpSummary.TabIndex = 2
        Me.chkIncludeExpSummary.Text = "Include Expense Summary"
        Me.chkIncludeExpSummary.UseVisualStyleBackColor = True
        '
        'txtEmailPassword
        '
        Me.txtEmailPassword.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtEmailPassword.Location = New System.Drawing.Point(124, 80)
        Me.txtEmailPassword.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtEmailPassword.Name = "txtEmailPassword"
        Me.txtEmailPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtEmailPassword.Size = New System.Drawing.Size(475, 26)
        Me.txtEmailPassword.TabIndex = 9
        '
        'txtEmailTo
        '
        Me.txtEmailTo.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtEmailTo.Location = New System.Drawing.Point(124, 186)
        Me.txtEmailTo.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtEmailTo.Name = "txtEmailTo"
        Me.txtEmailTo.Size = New System.Drawing.Size(475, 26)
        Me.txtEmailTo.TabIndex = 8
        '
        'txtSmtpPort
        '
        Me.txtSmtpPort.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSmtpPort.Location = New System.Drawing.Point(124, 151)
        Me.txtSmtpPort.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtSmtpPort.Name = "txtSmtpPort"
        Me.txtSmtpPort.Size = New System.Drawing.Size(475, 26)
        Me.txtSmtpPort.TabIndex = 7
        '
        'txtSmtpHost
        '
        Me.txtSmtpHost.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSmtpHost.Location = New System.Drawing.Point(124, 115)
        Me.txtSmtpHost.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtSmtpHost.Name = "txtSmtpHost"
        Me.txtSmtpHost.Size = New System.Drawing.Size(475, 26)
        Me.txtSmtpHost.TabIndex = 6
        '
        'txtEmailFrom
        '
        Me.txtEmailFrom.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtEmailFrom.Location = New System.Drawing.Point(124, 45)
        Me.txtEmailFrom.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtEmailFrom.Name = "txtEmailFrom"
        Me.txtEmailFrom.Size = New System.Drawing.Size(475, 26)
        Me.txtEmailFrom.TabIndex = 5
        '
        'Label34
        '
        Me.Label34.AutoSize = True
        Me.Label34.Location = New System.Drawing.Point(22, 157)
        Me.Label34.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(89, 20)
        Me.Label34.TabIndex = 4
        Me.Label34.Text = "SMTP Port:"
        '
        'Label33
        '
        Me.Label33.AutoSize = True
        Me.Label33.Location = New System.Drawing.Point(18, 122)
        Me.Label33.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(94, 20)
        Me.Label33.TabIndex = 3
        Me.Label33.Text = "SMTP Host:"
        '
        'Label32
        '
        Me.Label32.AutoSize = True
        Me.Label32.Location = New System.Drawing.Point(4, 192)
        Me.Label32.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(109, 20)
        Me.Label32.TabIndex = 2
        Me.Label32.Text = "Recipient List:"
        '
        'Label31
        '
        Me.Label31.AutoSize = True
        Me.Label31.Location = New System.Drawing.Point(32, 86)
        Me.Label31.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(82, 20)
        Me.Label31.TabIndex = 1
        Me.Label31.Text = "Password:"
        '
        'Label35
        '
        Me.Label35.AutoSize = True
        Me.Label35.Location = New System.Drawing.Point(28, 51)
        Me.Label35.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label35.Name = "Label35"
        Me.Label35.Size = New System.Drawing.Size(87, 20)
        Me.Label35.TabIndex = 0
        Me.Label35.Text = "Username:"
        '
        'chkSendEmails
        '
        Me.chkSendEmails.AutoSize = True
        Me.chkSendEmails.Location = New System.Drawing.Point(14, 15)
        Me.chkSendEmails.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.chkSendEmails.Name = "chkSendEmails"
        Me.chkSendEmails.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.chkSendEmails.Size = New System.Drawing.Size(207, 24)
        Me.chkSendEmails.TabIndex = 0
        Me.chkSendEmails.Text = "Send Notification Emails"
        Me.chkSendEmails.UseVisualStyleBackColor = True
        '
        'Label30
        '
        Me.Label30.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label30.AutoSize = True
        Me.Label30.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label30.Location = New System.Drawing.Point(1287, 625)
        Me.Label30.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(242, 20)
        Me.Label30.TabIndex = 7
        Me.Label30.Text = "Developed By: Sachin Wadi"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.txtReminder)
        Me.GroupBox1.Controls.Add(Me.chkShowReminder)
        Me.GroupBox1.Location = New System.Drawing.Point(514, 131)
        Me.GroupBox1.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox1.Size = New System.Drawing.Size(634, 231)
        Me.GroupBox1.TabIndex = 6
        Me.GroupBox1.TabStop = False
        '
        'txtReminder
        '
        Me.txtReminder.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtReminder.Location = New System.Drawing.Point(16, 48)
        Me.txtReminder.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtReminder.MaxLength = 1000
        Me.txtReminder.Multiline = True
        Me.txtReminder.Name = "txtReminder"
        Me.txtReminder.Size = New System.Drawing.Size(607, 169)
        Me.txtReminder.TabIndex = 1
        '
        'chkShowReminder
        '
        Me.chkShowReminder.AutoSize = True
        Me.chkShowReminder.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkShowReminder.Location = New System.Drawing.Point(9, 15)
        Me.chkShowReminder.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.chkShowReminder.Name = "chkShowReminder"
        Me.chkShowReminder.Size = New System.Drawing.Size(112, 24)
        Me.chkShowReminder.TabIndex = 0
        Me.chkShowReminder.Text = "Show Alert"
        Me.chkShowReminder.UseVisualStyleBackColor = True
        '
        'GroupBox6
        '
        Me.GroupBox6.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox6.Controls.Add(Me.Label27)
        Me.GroupBox6.Controls.Add(Me.btnBrowse)
        Me.GroupBox6.Controls.Add(Me.txtCSVBackupPath)
        Me.GroupBox6.Controls.Add(Me.Label13)
        Me.GroupBox6.Location = New System.Drawing.Point(514, 0)
        Me.GroupBox6.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox6.Size = New System.Drawing.Size(634, 131)
        Me.GroupBox6.TabIndex = 4
        Me.GroupBox6.TabStop = False
        '
        'Label27
        '
        Me.Label27.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label27.Location = New System.Drawing.Point(12, 62)
        Me.Label27.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(615, 69)
        Me.Label27.TabIndex = 3
        Me.Label27.Text = "Mention the location from where import service should pull the changes.This locat" &
    "ion usually should be the location shared with the cloud service providers such " &
    "as Google Drive, Dropbox, OneDrive etc."
        '
        'btnBrowse
        '
        Me.btnBrowse.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnBrowse.Location = New System.Drawing.Point(514, 17)
        Me.btnBrowse.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.Size = New System.Drawing.Size(112, 31)
        Me.btnBrowse.TabIndex = 2
        Me.btnBrowse.Text = "Browse"
        Me.btnBrowse.UseVisualStyleBackColor = True
        '
        'txtCSVBackupPath
        '
        Me.txtCSVBackupPath.BackColor = System.Drawing.SystemColors.Window
        Me.txtCSVBackupPath.Location = New System.Drawing.Point(57, 18)
        Me.txtCSVBackupPath.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtCSVBackupPath.Name = "txtCSVBackupPath"
        Me.txtCSVBackupPath.ReadOnly = True
        Me.txtCSVBackupPath.Size = New System.Drawing.Size(446, 26)
        Me.txtCSVBackupPath.TabIndex = 1
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.Location = New System.Drawing.Point(9, 25)
        Me.Label13.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(46, 20)
        Me.Label13.TabIndex = 0
        Me.Label13.Text = "Path:"
        '
        'GroupBox5
        '
        Me.GroupBox5.Controls.Add(Me.dtpRecdKeeping)
        Me.GroupBox5.Controls.Add(Me.Label12)
        Me.GroupBox5.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.GroupBox5.Location = New System.Drawing.Point(9, 145)
        Me.GroupBox5.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox5.Size = New System.Drawing.Size(496, 68)
        Me.GroupBox5.TabIndex = 3
        Me.GroupBox5.TabStop = False
        '
        'dtpRecdKeeping
        '
        Me.dtpRecdKeeping.CustomFormat = "dd/MM/yyyy"
        Me.dtpRecdKeeping.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpRecdKeeping.Location = New System.Drawing.Point(220, 20)
        Me.dtpRecdKeeping.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.dtpRecdKeeping.Name = "dtpRecdKeeping"
        Me.dtpRecdKeeping.Size = New System.Drawing.Size(194, 26)
        Me.dtpRecdKeeping.TabIndex = 3
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(4, 26)
        Me.Label12.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(199, 20)
        Me.Label12.TabIndex = 2
        Me.Label12.Text = "Book Keeping Start Date:"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Label11)
        Me.GroupBox3.Controls.Add(Me.Label10)
        Me.GroupBox3.Controls.Add(Me.btnOpenYrAll)
        Me.GroupBox3.Controls.Add(Me.btnCloseYr)
        Me.GroupBox3.Controls.Add(Me.btnOpenYr)
        Me.GroupBox3.Controls.Add(Me.btnCloseYrAll)
        Me.GroupBox3.Controls.Add(Me.lstboxClosedYears)
        Me.GroupBox3.Controls.Add(Me.lstboxOpenYears)
        Me.GroupBox3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.Location = New System.Drawing.Point(9, 389)
        Me.GroupBox3.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox3.Size = New System.Drawing.Size(496, 256)
        Me.GroupBox3.TabIndex = 1
        Me.GroupBox3.TabStop = False
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label11.Location = New System.Drawing.Point(258, 13)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(130, 22)
        Me.Label11.TabIndex = 10
        Me.Label11.Text = "Closed Years"
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(87, 13)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(116, 22)
        Me.Label10.TabIndex = 9
        Me.Label10.Text = "Open Years"
        '
        'btnOpenYrAll
        '
        Me.btnOpenYrAll.Location = New System.Drawing.Point(209, 197)
        Me.btnOpenYrAll.Name = "btnOpenYrAll"
        Me.btnOpenYrAll.Size = New System.Drawing.Size(47, 37)
        Me.btnOpenYrAll.TabIndex = 8
        Me.btnOpenYrAll.Text = "<<"
        Me.btnOpenYrAll.UseVisualStyleBackColor = True
        Me.btnOpenYrAll.Visible = False
        '
        'btnCloseYr
        '
        Me.btnCloseYr.Location = New System.Drawing.Point(209, 103)
        Me.btnCloseYr.Name = "btnCloseYr"
        Me.btnCloseYr.Size = New System.Drawing.Size(47, 37)
        Me.btnCloseYr.TabIndex = 7
        Me.btnCloseYr.Text = ">"
        Me.btnCloseYr.UseVisualStyleBackColor = True
        '
        'btnOpenYr
        '
        Me.btnOpenYr.Location = New System.Drawing.Point(209, 150)
        Me.btnOpenYr.Name = "btnOpenYr"
        Me.btnOpenYr.Size = New System.Drawing.Size(47, 37)
        Me.btnOpenYr.TabIndex = 6
        Me.btnOpenYr.Text = "<"
        Me.btnOpenYr.UseVisualStyleBackColor = True
        '
        'btnCloseYrAll
        '
        Me.btnCloseYrAll.Location = New System.Drawing.Point(209, 56)
        Me.btnCloseYrAll.Name = "btnCloseYrAll"
        Me.btnCloseYrAll.Size = New System.Drawing.Size(47, 37)
        Me.btnCloseYrAll.TabIndex = 5
        Me.btnCloseYrAll.Text = ">>"
        Me.btnCloseYrAll.UseVisualStyleBackColor = True
        Me.btnCloseYrAll.Visible = False
        '
        'lstboxClosedYears
        '
        Me.lstboxClosedYears.FormattingEnabled = True
        Me.lstboxClosedYears.ItemHeight = 20
        Me.lstboxClosedYears.Location = New System.Drawing.Point(262, 37)
        Me.lstboxClosedYears.Name = "lstboxClosedYears"
        Me.lstboxClosedYears.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.lstboxClosedYears.Size = New System.Drawing.Size(120, 204)
        Me.lstboxClosedYears.TabIndex = 4
        '
        'lstboxOpenYears
        '
        Me.lstboxOpenYears.FormattingEnabled = True
        Me.lstboxOpenYears.ItemHeight = 20
        Me.lstboxOpenYears.Location = New System.Drawing.Point(83, 37)
        Me.lstboxOpenYears.Name = "lstboxOpenYears"
        Me.lstboxOpenYears.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended
        Me.lstboxOpenYears.Size = New System.Drawing.Size(120, 204)
        Me.lstboxOpenYears.TabIndex = 3
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.txtHighlightSummYr)
        Me.GroupBox2.Controls.Add(Me.txtHighlightSummMth)
        Me.GroupBox2.Controls.Add(Me.Label21)
        Me.GroupBox2.Controls.Add(Me.Label16)
        Me.GroupBox2.Controls.Add(Me.Label7)
        Me.GroupBox2.Controls.Add(Me.txtHighlightDet)
        Me.GroupBox2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.GroupBox2.Location = New System.Drawing.Point(9, 215)
        Me.GroupBox2.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Padding = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.GroupBox2.Size = New System.Drawing.Size(496, 171)
        Me.GroupBox2.TabIndex = 0
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Highlight Amount >= "
        '
        'txtHighlightSummYr
        '
        Me.txtHighlightSummYr.Location = New System.Drawing.Point(153, 125)
        Me.txtHighlightSummYr.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtHighlightSummYr.MaxLength = 6
        Me.txtHighlightSummYr.Name = "txtHighlightSummYr"
        Me.txtHighlightSummYr.Size = New System.Drawing.Size(148, 26)
        Me.txtHighlightSummYr.TabIndex = 7
        '
        'txtHighlightSummMth
        '
        Me.txtHighlightSummMth.Location = New System.Drawing.Point(153, 78)
        Me.txtHighlightSummMth.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtHighlightSummMth.MaxLength = 6
        Me.txtHighlightSummMth.Name = "txtHighlightSummMth"
        Me.txtHighlightSummMth.Size = New System.Drawing.Size(148, 26)
        Me.txtHighlightSummMth.TabIndex = 6
        '
        'Label21
        '
        Me.Label21.AutoSize = True
        Me.Label21.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label21.Location = New System.Drawing.Point(90, 43)
        Me.Label21.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(67, 20)
        Me.Label21.TabIndex = 5
        Me.Label21.Text = "Details:"
        '
        'Label16
        '
        Me.Label16.AutoSize = True
        Me.Label16.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label16.Location = New System.Drawing.Point(4, 85)
        Me.Label16.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(160, 20)
        Me.Label16.TabIndex = 4
        Me.Label16.Text = "Summary (Monthly):"
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(16, 131)
        Me.Label7.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(148, 20)
        Me.Label7.TabIndex = 3
        Me.Label7.Text = "Summary (Yearly):"
        '
        'txtHighlightDet
        '
        Me.txtHighlightDet.Location = New System.Drawing.Point(153, 37)
        Me.txtHighlightDet.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtHighlightDet.MaxLength = 6
        Me.txtHighlightDet.Name = "txtHighlightDet"
        Me.txtHighlightDet.Size = New System.Drawing.Size(148, 26)
        Me.txtHighlightDet.TabIndex = 0
        '
        'Panel3
        '
        Me.Panel3.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel3.Controls.Add(Me.ImpensaTabControl)
        Me.Panel3.Location = New System.Drawing.Point(314, 0)
        Me.Panel3.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(1560, 697)
        Me.Panel3.TabIndex = 5
        '
        'tmrAlert
        '
        '
        'lblAlertText
        '
        Me.lblAlertText.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblAlertText.AutoSize = True
        Me.lblAlertText.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAlertText.ForeColor = System.Drawing.Color.White
        Me.lblAlertText.Location = New System.Drawing.Point(0, 6)
        Me.lblAlertText.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblAlertText.Name = "lblAlertText"
        Me.lblAlertText.Size = New System.Drawing.Size(75, 20)
        Me.lblAlertText.TabIndex = 1
        Me.lblAlertText.Text = "Label20"
        '
        'chkShowAllDet
        '
        Me.chkShowAllDet.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.chkShowAllDet.AutoSize = True
        Me.chkShowAllDet.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkShowAllDet.Location = New System.Drawing.Point(1196, 719)
        Me.chkShowAllDet.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.chkShowAllDet.Name = "chkShowAllDet"
        Me.chkShowAllDet.Size = New System.Drawing.Size(184, 24)
        Me.chkShowAllDet.TabIndex = 13
        Me.chkShowAllDet.Text = "Show All Records"
        Me.chkShowAllDet.UseVisualStyleBackColor = True
        Me.chkShowAllDet.Visible = False
        '
        'Label20
        '
        Me.Label20.AutoSize = True
        Me.Label20.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label20.Location = New System.Drawing.Point(-2, 18)
        Me.Label20.Margin = New System.Windows.Forms.Padding(0)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(181, 20)
        Me.Label20.TabIndex = 2
        Me.Label20.Text = "Highlight Amount >="
        '
        'txtHighlight
        '
        Me.txtHighlight.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtHighlight.Location = New System.Drawing.Point(184, 12)
        Me.txtHighlight.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.txtHighlight.MaxLength = 6
        Me.txtHighlight.Name = "txtHighlight"
        Me.txtHighlight.Size = New System.Drawing.Size(100, 26)
        Me.txtHighlight.TabIndex = 3
        '
        'pnlHighlight
        '
        Me.pnlHighlight.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.pnlHighlight.Controls.Add(Me.btnHighlight)
        Me.pnlHighlight.Controls.Add(Me.Label20)
        Me.pnlHighlight.Controls.Add(Me.txtHighlight)
        Me.pnlHighlight.Location = New System.Drawing.Point(826, 700)
        Me.pnlHighlight.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.pnlHighlight.Name = "pnlHighlight"
        Me.pnlHighlight.Size = New System.Drawing.Size(360, 60)
        Me.pnlHighlight.TabIndex = 14
        '
        'btnHighlight
        '
        Me.btnHighlight.AutoSize = True
        Me.btnHighlight.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnHighlight.Location = New System.Drawing.Point(297, 11)
        Me.btnHighlight.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.btnHighlight.Name = "btnHighlight"
        Me.btnHighlight.Size = New System.Drawing.Size(64, 46)
        Me.btnHighlight.TabIndex = 4
        Me.btnHighlight.Text = "Go"
        Me.btnHighlight.UseVisualStyleBackColor = True
        '
        'StatusStrip
        '
        Me.StatusStrip.Font = New System.Drawing.Font("Segoe UI", 8.25!)
        Me.StatusStrip.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.StatusStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tslblGridTotal, Me.tslblSeperator1, Me.tslblRecdCnt, Me.tsCmbBudgetBuckets, Me.tslblSeperator2, Me.tslblMTD, Me.tslblSeperator6, Me.tslblYTD, Me.tslblSeperator7, Me.tslblBKSDTD})
        Me.StatusStrip.Location = New System.Drawing.Point(0, 853)
        Me.StatusStrip.Name = "StatusStrip"
        Me.StatusStrip.Padding = New System.Windows.Forms.Padding(2, 0, 21, 0)
        Me.StatusStrip.Size = New System.Drawing.Size(1875, 35)
        Me.StatusStrip.SizingGrip = False
        Me.StatusStrip.TabIndex = 15
        Me.StatusStrip.Text = "StatusStrip1"
        '
        'tslblGridTotal
        '
        Me.tslblGridTotal.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
        Me.tslblGridTotal.ForeColor = System.Drawing.Color.Green
        Me.tslblGridTotal.LinkColor = System.Drawing.Color.Green
        Me.tslblGridTotal.Name = "tslblGridTotal"
        Me.tslblGridTotal.Size = New System.Drawing.Size(0, 28)
        Me.tslblGridTotal.VisitedLinkColor = System.Drawing.Color.Green
        '
        'tslblSeperator1
        '
        Me.tslblSeperator1.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Bold)
        Me.tslblSeperator1.ForeColor = System.Drawing.Color.Blue
        Me.tslblSeperator1.Name = "tslblSeperator1"
        Me.tslblSeperator1.Size = New System.Drawing.Size(26, 28)
        Me.tslblSeperator1.Text = "||"
        '
        'tslblRecdCnt
        '
        Me.tslblRecdCnt.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
        Me.tslblRecdCnt.ForeColor = System.Drawing.Color.Green
        Me.tslblRecdCnt.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline
        Me.tslblRecdCnt.LinkColor = System.Drawing.Color.Green
        Me.tslblRecdCnt.Name = "tslblRecdCnt"
        Me.tslblRecdCnt.Size = New System.Drawing.Size(0, 28)
        Me.tslblRecdCnt.VisitedLinkColor = System.Drawing.Color.Green
        '
        'tsCmbBudgetBuckets
        '
        Me.tsCmbBudgetBuckets.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.tsCmbBudgetBuckets.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.tsMenuOBCats, Me.tsMenuAPCats, Me.tsMenuUBCats})
        Me.tsCmbBudgetBuckets.ImageTransparentColor = System.Drawing.Color.Magenta
        Me.tsCmbBudgetBuckets.Name = "tsCmbBudgetBuckets"
        Me.tsCmbBudgetBuckets.Size = New System.Drawing.Size(18, 32)
        Me.tsCmbBudgetBuckets.Visible = False
        '
        'tsMenuOBCats
        '
        Me.tsMenuOBCats.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tsMenuOBCats.ForeColor = System.Drawing.Color.Red
        Me.tsMenuOBCats.Name = "tsMenuOBCats"
        Me.tsMenuOBCats.Size = New System.Drawing.Size(254, 34)
        Me.tsMenuOBCats.Text = "Over Budget: #"
        '
        'tsMenuAPCats
        '
        Me.tsMenuAPCats.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
        Me.tsMenuAPCats.ForeColor = System.Drawing.Color.Orange
        Me.tsMenuAPCats.Name = "tsMenuAPCats"
        Me.tsMenuAPCats.Size = New System.Drawing.Size(254, 34)
        Me.tsMenuAPCats.Text = "At Par: #"
        '
        'tsMenuUBCats
        '
        Me.tsMenuUBCats.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
        Me.tsMenuUBCats.ForeColor = System.Drawing.Color.Green
        Me.tsMenuUBCats.Name = "tsMenuUBCats"
        Me.tsMenuUBCats.Size = New System.Drawing.Size(254, 34)
        Me.tsMenuUBCats.Text = "Under Budget: #"
        '
        'tslblSeperator2
        '
        Me.tslblSeperator2.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Bold)
        Me.tslblSeperator2.ForeColor = System.Drawing.Color.Blue
        Me.tslblSeperator2.Name = "tslblSeperator2"
        Me.tslblSeperator2.Size = New System.Drawing.Size(26, 28)
        Me.tslblSeperator2.Text = "||"
        '
        'tslblMTD
        '
        Me.tslblMTD.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
        Me.tslblMTD.ForeColor = System.Drawing.Color.Green
        Me.tslblMTD.IsLink = True
        Me.tslblMTD.LinkColor = System.Drawing.Color.Green
        Me.tslblMTD.Name = "tslblMTD"
        Me.tslblMTD.Size = New System.Drawing.Size(0, 28)
        Me.tslblMTD.VisitedLinkColor = System.Drawing.Color.Green
        '
        'tslblSeperator6
        '
        Me.tslblSeperator6.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Bold)
        Me.tslblSeperator6.ForeColor = System.Drawing.Color.Blue
        Me.tslblSeperator6.Name = "tslblSeperator6"
        Me.tslblSeperator6.Size = New System.Drawing.Size(26, 28)
        Me.tslblSeperator6.Text = "||"
        '
        'tslblYTD
        '
        Me.tslblYTD.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
        Me.tslblYTD.ForeColor = System.Drawing.Color.Green
        Me.tslblYTD.IsLink = True
        Me.tslblYTD.LinkColor = System.Drawing.Color.Green
        Me.tslblYTD.Name = "tslblYTD"
        Me.tslblYTD.Size = New System.Drawing.Size(0, 28)
        Me.tslblYTD.VisitedLinkColor = System.Drawing.Color.Green
        '
        'tslblSeperator7
        '
        Me.tslblSeperator7.Font = New System.Drawing.Font("Segoe UI", 10.0!, System.Drawing.FontStyle.Bold)
        Me.tslblSeperator7.ForeColor = System.Drawing.Color.Blue
        Me.tslblSeperator7.Name = "tslblSeperator7"
        Me.tslblSeperator7.Size = New System.Drawing.Size(26, 28)
        Me.tslblSeperator7.Text = "||"
        '
        'tslblBKSDTD
        '
        Me.tslblBKSDTD.Font = New System.Drawing.Font("Microsoft Sans Serif", 10.0!)
        Me.tslblBKSDTD.ForeColor = System.Drawing.Color.Green
        Me.tslblBKSDTD.IsLink = True
        Me.tslblBKSDTD.LinkColor = System.Drawing.Color.Green
        Me.tslblBKSDTD.Name = "tslblBKSDTD"
        Me.tslblBKSDTD.Size = New System.Drawing.Size(0, 28)
        Me.tslblBKSDTD.VisitedLinkColor = System.Drawing.Color.Green
        '
        'RchTB_MTDTicker
        '
        Me.RchTB_MTDTicker.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.RchTB_MTDTicker.BackColor = System.Drawing.SystemColors.Control
        Me.RchTB_MTDTicker.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.RchTB_MTDTicker.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RchTB_MTDTicker.Location = New System.Drawing.Point(152, 12)
        Me.RchTB_MTDTicker.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.RchTB_MTDTicker.Multiline = False
        Me.RchTB_MTDTicker.Name = "RchTB_MTDTicker"
        Me.RchTB_MTDTicker.ReadOnly = True
        Me.RchTB_MTDTicker.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.None
        Me.RchTB_MTDTicker.Size = New System.Drawing.Size(1074, 20)
        Me.RchTB_MTDTicker.TabIndex = 16
        Me.RchTB_MTDTicker.TabStop = False
        Me.RchTB_MTDTicker.Text = ""
        Me.RchTB_MTDTicker.Visible = False
        Me.RchTB_MTDTicker.WordWrap = False
        '
        'RchTB_YTDTicker
        '
        Me.RchTB_YTDTicker.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.RchTB_YTDTicker.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.RchTB_YTDTicker.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.RchTB_YTDTicker.Location = New System.Drawing.Point(152, 43)
        Me.RchTB_YTDTicker.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.RchTB_YTDTicker.Multiline = False
        Me.RchTB_YTDTicker.Name = "RchTB_YTDTicker"
        Me.RchTB_YTDTicker.ReadOnly = True
        Me.RchTB_YTDTicker.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.None
        Me.RchTB_YTDTicker.Size = New System.Drawing.Size(596, 20)
        Me.RchTB_YTDTicker.TabIndex = 17
        Me.RchTB_YTDTicker.TabStop = False
        Me.RchTB_YTDTicker.Text = ""
        Me.RchTB_YTDTicker.Visible = False
        Me.RchTB_YTDTicker.WordWrap = False
        '
        'Label23
        '
        Me.Label23.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label23.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label23.ForeColor = System.Drawing.Color.Blue
        Me.Label23.Location = New System.Drawing.Point(0, 40)
        Me.Label23.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(308, 23)
        Me.Label23.TabIndex = 18
        Me.Label23.Text = "YTD Summary  (Actual / Forecast):"
        '
        'Label24
        '
        Me.Label24.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label24.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label24.ForeColor = System.Drawing.Color.Blue
        Me.Label24.Location = New System.Drawing.Point(0, 9)
        Me.Label24.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(308, 23)
        Me.Label24.TabIndex = 19
        Me.Label24.Text = "MTD Summary (Actual / Forecast):"
        '
        'tmrTicker
        '
        '
        'Panel9
        '
        Me.Panel9.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Panel9.Controls.Add(Me.Label24)
        Me.Panel9.Controls.Add(Me.Label23)
        Me.Panel9.Controls.Add(Me.RchTB_MTDTicker)
        Me.Panel9.Controls.Add(Me.RchTB_YTDTicker)
        Me.Panel9.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Panel9.Location = New System.Drawing.Point(0, 777)
        Me.Panel9.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Panel9.Name = "Panel9"
        Me.Panel9.Size = New System.Drawing.Size(1695, 74)
        Me.Panel9.TabIndex = 20
        '
        'Panel10
        '
        Me.Panel10.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Panel10.BackColor = System.Drawing.Color.GreenYellow
        Me.Panel10.Controls.Add(Me.lblAlertText)
        Me.Panel10.Location = New System.Drawing.Point(0, 752)
        Me.Panel10.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Panel10.Name = "Panel10"
        Me.Panel10.Size = New System.Drawing.Size(300, 32)
        Me.Panel10.TabIndex = 21
        Me.Panel10.Visible = False
        '
        'BGWorker
        '
        '
        'NotifyIcon
        '
        Me.NotifyIcon.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info
        Me.NotifyIcon.BalloonTipTitle = "Impensa Notifications"
        Me.NotifyIcon.ContextMenuStrip = Me.ctxContextMenu
        Me.NotifyIcon.Icon = CType(resources.GetObject("NotifyIcon.Icon"), System.Drawing.Icon)
        Me.NotifyIcon.Visible = True
        '
        'ctxContextMenu
        '
        Me.ctxContextMenu.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.ctxContextMenu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ExitToolStripMenuItem})
        Me.ctxContextMenu.Name = "ContextMenu"
        Me.ctxContextMenu.Size = New System.Drawing.Size(112, 36)
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(111, 32)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'tmrRefresh
        '
        Me.tmrRefresh.Enabled = True
        Me.tmrRefresh.Interval = 5000
        '
        'btnExport
        '
        Me.btnExport.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btnExport.Enabled = False
        Me.btnExport.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnExport.Location = New System.Drawing.Point(570, 711)
        Me.btnExport.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.btnExport.Name = "btnExport"
        Me.btnExport.Size = New System.Drawing.Size(248, 35)
        Me.btnExport.TabIndex = 22
        Me.btnExport.Text = "Export To PDF"
        Me.btnExport.UseVisualStyleBackColor = True
        '
        'BgWorker_Email
        '
        '
        'Timer_Import
        '
        Me.Timer_Import.Enabled = True
        Me.Timer_Import.Interval = 300000
        '
        'frmMain
        '
        Me.AcceptButton = Me.btnSubmit
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1875, 888)
        Me.Controls.Add(Me.btnExport)
        Me.Controls.Add(Me.chkShowAllDet)
        Me.Controls.Add(Me.Panel10)
        Me.Controls.Add(Me.StatusStrip)
        Me.Controls.Add(Me.btnSave)
        Me.Controls.Add(Me.pnlHighlight)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.Panel9)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Margin = New System.Windows.Forms.Padding(4, 5, 4, 5)
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Impensa"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabCategories.ResumeLayout(False)
        CType(Me.DataGridCatList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabCharts.ResumeLayout(False)
        Me.Panel4.ResumeLayout(False)
        CType(Me.Chart_Analysis, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.TabSummary.ResumeLayout(False)
        Me.Panel8.ResumeLayout(False)
        Me.Panel8.PerformLayout()
        CType(Me.DataGridExpSumm, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabDetails.ResumeLayout(False)
        Me.TabDetails.PerformLayout()
        CType(Me.DataGridExpDet, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel5.ResumeLayout(False)
        Me.Panel5.PerformLayout()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ImpensaTabControl.ResumeLayout(False)
        Me.TabBudget.ResumeLayout(False)
        Me.Panel7.ResumeLayout(False)
        CType(Me.DataGridThrLimits, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel6.ResumeLayout(False)
        Me.Panel6.PerformLayout()
        Me.TabSettings.ResumeLayout(False)
        Me.TabSettings.PerformLayout()
        Me.GroupBox7.ResumeLayout(False)
        Me.GroupBox7.PerformLayout()
        Me.gbEmailConfig.ResumeLayout(False)
        Me.gbEmailConfig.PerformLayout()
        Me.grpEmailSettings.ResumeLayout(False)
        Me.grpEmailSettings.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox6.PerformLayout()
        Me.GroupBox5.ResumeLayout(False)
        Me.GroupBox5.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.pnlHighlight.ResumeLayout(False)
        Me.pnlHighlight.PerformLayout()
        Me.StatusStrip.ResumeLayout(False)
        Me.StatusStrip.PerformLayout()
        Me.Panel9.ResumeLayout(False)
        Me.Panel10.ResumeLayout(False)
        Me.Panel10.PerformLayout()
        Me.ctxContextMenu.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents dtPickerFrom As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents dtPickerTo As System.Windows.Forms.DateTimePicker
    Friend WithEvents LstCategory As System.Windows.Forms.CheckedListBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents btnSubmit As System.Windows.Forms.Button
    'Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents TabCategories As System.Windows.Forms.TabPage
    Friend WithEvents DataGridCatList As System.Windows.Forms.DataGridView
    Friend WithEvents TabCharts As System.Windows.Forms.TabPage
    Friend WithEvents TabSummary As System.Windows.Forms.TabPage
    Friend WithEvents TabDetails As System.Windows.Forms.TabPage
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents ImpensaTabControl As System.Windows.Forms.TabControl
    Friend WithEvents DataGridExpSumm As System.Windows.Forms.DataGridView
    Friend WithEvents Chart_Analysis As System.Windows.Forms.DataVisualization.Charting.Chart
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents btnGo As System.Windows.Forms.Button
    Friend WithEvents cmbListing As System.Windows.Forms.ComboBox
    Friend WithEvents cmbSelectChart As System.Windows.Forms.ComboBox
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents cmbChartType As System.Windows.Forms.ComboBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtSearch As System.Windows.Forms.TextBox
    Friend WithEvents TabSettings As System.Windows.Forms.TabPage
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents txtHighlightDet As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents dtpRecdKeeping As System.Windows.Forms.DateTimePicker
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents btnBrowse As System.Windows.Forms.Button
    Friend WithEvents txtCSVBackupPath As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents LinkLabel2 As System.Windows.Forms.LinkLabel
    Friend WithEvents Panel5 As System.Windows.Forms.Panel
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents cmbSort As System.Windows.Forms.ComboBox
    Friend WithEvents chkSort As System.Windows.Forms.CheckBox
    Friend WithEvents TabBudget As System.Windows.Forms.TabPage
    Friend WithEvents Panel6 As System.Windows.Forms.Panel
    Friend WithEvents cmbThrMonth As System.Windows.Forms.ComboBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Panel7 As System.Windows.Forms.Panel
    Friend WithEvents DataGridThrLimits As System.Windows.Forms.DataGridView
    Friend WithEvents DataGridExpDet As System.Windows.Forms.DataGridView
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents txtReminder As System.Windows.Forms.TextBox
    Friend WithEvents chkShowReminder As System.Windows.Forms.CheckBox
    Friend WithEvents tmrAlert As System.Windows.Forms.Timer
    Friend WithEvents lblAlertText As System.Windows.Forms.Label
    Friend WithEvents chkShowAllDet As System.Windows.Forms.CheckBox
    Friend WithEvents cmbSummaryType As System.Windows.Forms.ComboBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents txtHighlight As System.Windows.Forms.TextBox
    Friend WithEvents pnlHighlight As System.Windows.Forms.Panel
    Friend WithEvents btnHighlight As System.Windows.Forms.Button
    Friend WithEvents StatusStrip As System.Windows.Forms.StatusStrip
    Friend WithEvents tslblGridTotal As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslblRecdCnt As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslblMTD As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslblYTD As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslblSeperator1 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslblSeperator2 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslblSeperator6 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tslblSeperator7 As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents txtHighlightSummYr As System.Windows.Forms.TextBox
    Friend WithEvents txtHighlightSummMth As System.Windows.Forms.TextBox
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Panel8 As System.Windows.Forms.Panel
    Friend WithEvents cmbCatListRunTot As System.Windows.Forms.ComboBox
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents chkPeriodLevel As System.Windows.Forms.CheckBox
    Friend WithEvents RchTB_MTDTicker As System.Windows.Forms.RichTextBox
    Friend WithEvents RchTB_YTDTicker As System.Windows.Forms.RichTextBox
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents tmrTicker As System.Windows.Forms.Timer
    Friend WithEvents Panel9 As System.Windows.Forms.Panel
    Friend WithEvents Panel10 As System.Windows.Forms.Panel
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents cmbPeriod As System.Windows.Forms.ComboBox
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents chkLBYears As System.Windows.Forms.CheckedListBox
    Friend WithEvents tslblBKSDTD As System.Windows.Forms.ToolStripStatusLabel
    Friend WithEvents tsCmbBudgetBuckets As System.Windows.Forms.ToolStripDropDownButton
    Friend WithEvents tsMenuAPCats As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsMenuUBCats As System.Windows.Forms.ToolStripMenuItem
    'Friend WithEvents ToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    'Friend WithEvents ToolStripMenuItem2 As System.Windows.Forms.ToolStripMenuItem
    'Friend WithEvents ToolStripMenuItem3 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents tsMenuOBCats As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents chkLBVarComparision As System.Windows.Forms.CheckedListBox
    Friend WithEvents btnVarCompare As System.Windows.Forms.Button
    Friend WithEvents BGWorker As System.ComponentModel.BackgroundWorker
    Friend WithEvents NotifyIcon As System.Windows.Forms.NotifyIcon
    Friend WithEvents ctxContextMenu As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents chkStartImport As System.Windows.Forms.CheckBox
    Friend WithEvents chkShowLabel As System.Windows.Forms.CheckBox
    Friend WithEvents tmrRefresh As System.Windows.Forms.Timer
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents LstUnpaidBillsCurrentMonth As System.Windows.Forms.ListBox
    Friend WithEvents LstUnpaidBillsPrevMonth As System.Windows.Forms.ListBox
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents btnExport As System.Windows.Forms.Button
    Friend WithEvents Label30 As System.Windows.Forms.Label
    Friend WithEvents gbEmailConfig As System.Windows.Forms.GroupBox
    Friend WithEvents grpEmailSettings As System.Windows.Forms.GroupBox
    Friend WithEvents txtEmailPassword As System.Windows.Forms.TextBox
    Friend WithEvents txtEmailTo As System.Windows.Forms.TextBox
    Friend WithEvents txtSmtpPort As System.Windows.Forms.TextBox
    Friend WithEvents txtSmtpHost As System.Windows.Forms.TextBox
    Friend WithEvents txtEmailFrom As System.Windows.Forms.TextBox
    Friend WithEvents Label34 As System.Windows.Forms.Label
    Friend WithEvents Label33 As System.Windows.Forms.Label
    Friend WithEvents Label32 As System.Windows.Forms.Label
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents Label35 As System.Windows.Forms.Label
    Friend WithEvents chkSendEmails As System.Windows.Forms.CheckBox
    Friend WithEvents chkIncludeExpSummary As System.Windows.Forms.CheckBox
    Friend WithEvents BgWorker_Email As System.ComponentModel.BackgroundWorker
    Friend WithEvents GroupBox7 As System.Windows.Forms.GroupBox
    Friend WithEvents chkExcelDelRows As System.Windows.Forms.CheckBox
    Friend WithEvents Label36 As System.Windows.Forms.Label
    Friend WithEvents Timer_Import As Timer
    Friend WithEvents btnOpenYrAll As Button
    Friend WithEvents btnCloseYr As Button
    Friend WithEvents btnOpenYr As Button
    Friend WithEvents btnCloseYrAll As Button
    Friend WithEvents lstboxClosedYears As ListBox
    Friend WithEvents lstboxOpenYears As ListBox
    Friend WithEvents Label11 As Label
    Friend WithEvents Label10 As Label
End Class
