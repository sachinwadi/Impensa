Public Class clsRowsToDeleteRange
    Private _StartIndex As Integer
    Private _EndIndex As Integer
    Public Property StartIndex() As Integer
        Get
            Return _StartIndex
        End Get
        Set(ByVal value As Integer)
            _StartIndex = value
        End Set

    End Property

    Public Property EndIndex() As Integer
        Get
            Return _EndIndex
        End Get
        Set(ByVal value As Integer)
            _EndIndex = value
        End Set
    End Property
End Class
