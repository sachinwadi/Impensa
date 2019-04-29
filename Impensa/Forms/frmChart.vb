Imports Impensa.clsLib
Imports System.Drawing.Printing

Public Class frmChart
    Dim bitmap As Bitmap

    Private Sub frmChart_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call frmMain.DisplayGraph(Chart_Analysis)
    End Sub

    Private Sub PrintToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintToolStripMenuItem.Click
        Me.SendToBack()

        Dim paperSizes As IEnumerable(Of PaperSize) = PrintDocument.PrinterSettings.PaperSizes.Cast(Of PaperSize)()
        Dim sizeA4 As PaperSize = paperSizes.First(Function(size) size.Kind = PaperKind.A4)
        PrintDocument.DefaultPageSettings.PaperSize = sizeA4
        PrintDocument.DefaultPageSettings.Landscape = True

        PrintDialog.Document = PrintDocument
        If (PrintDialog.ShowDialog = Windows.Forms.DialogResult.OK) Then PrintDocument.Print()
    End Sub

    Private Sub PrintDocument1_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument.PrintPage
        Dim myImage As New Bitmap(Me.Chart_Analysis.Width, Me.Chart_Analysis.Height)
        Dim PrintSize As Size = e.MarginBounds.Size
        Dim scale As Double = 1
        Me.Chart_Analysis.DrawToBitmap(myImage, New Rectangle(Point.Empty, Me.Chart_Analysis.Size))

        If myImage.Width > PrintSize.Width Then
            'Form is to big. Scale it down.
            scale = PrintSize.Width / myImage.Width
            e.Graphics.ScaleTransform(scale, scale)
        End If

        If (myImage.Height * scale) > PrintSize.Height Then
            'The form is still to big. Scale it again.
            scale = PrintSize.Height / (myImage.Height * scale)
            e.Graphics.ScaleTransform(scale, scale)
        End If

        e.Graphics.SmoothingMode = Drawing2D.SmoothingMode.HighQuality
        e.Graphics.DrawImage(myImage, e.MarginBounds.Location)
        myImage.Dispose()
    End Sub
End Class