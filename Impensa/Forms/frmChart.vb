Imports Impensa.clsLib
Public Class frmChart
    Private Sub frmChart_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call frmMain.DisplayGraph(Chart_Analysis)
    End Sub
End Class