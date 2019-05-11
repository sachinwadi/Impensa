<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmChart
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
        Dim Legend1 As System.Windows.Forms.DataVisualization.Charting.Legend = New System.Windows.Forms.DataVisualization.Charting.Legend()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmChart))
        Me.Chart_Analysis = New System.Windows.Forms.DataVisualization.Charting.Chart()
        Me.ChartContextMenu = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.PrintToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PrintDocument = New System.Drawing.Printing.PrintDocument()
        Me.PrintDialog = New System.Windows.Forms.PrintDialog()
        CType(Me.Chart_Analysis, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ChartContextMenu.SuspendLayout()
        Me.SuspendLayout()
        '
        'Chart_Analysis
        '
        Me.Chart_Analysis.Dock = System.Windows.Forms.DockStyle.Fill
        Legend1.Name = "Legend1"
        Me.Chart_Analysis.Legends.Add(Legend1)
        Me.Chart_Analysis.Location = New System.Drawing.Point(0, 0)
        Me.Chart_Analysis.Name = "Chart_Analysis"
        Me.Chart_Analysis.Size = New System.Drawing.Size(724, 413)
        Me.Chart_Analysis.TabIndex = 0
        Me.Chart_Analysis.Text = "Chart1"
        '
        'ChartContextMenu
        '
        Me.ChartContextMenu.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.PrintToolStripMenuItem})
        Me.ChartContextMenu.Name = "ContextMenuStrip1"
        Me.ChartContextMenu.Size = New System.Drawing.Size(100, 26)
        '
        'PrintToolStripMenuItem
        '
        Me.PrintToolStripMenuItem.Name = "PrintToolStripMenuItem"
        Me.PrintToolStripMenuItem.Size = New System.Drawing.Size(99, 22)
        Me.PrintToolStripMenuItem.Text = "Print"
        '
        'PrintDocument
        '
        '
        'PrintDialog
        '
        Me.PrintDialog.UseEXDialog = True
        '
        'frmChart
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(724, 413)
        Me.ContextMenuStrip = Me.ChartContextMenu
        Me.Controls.Add(Me.Chart_Analysis)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmChart"
        Me.Text = "Impensa"
        Me.TopMost = True
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        CType(Me.Chart_Analysis, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ChartContextMenu.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Chart_Analysis As System.Windows.Forms.DataVisualization.Charting.Chart
    Friend WithEvents ChartContextMenu As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents PrintToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PrintDocument As System.Drawing.Printing.PrintDocument
    Friend WithEvents PrintDialog As System.Windows.Forms.PrintDialog
End Class
