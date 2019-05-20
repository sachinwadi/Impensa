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