Imports System.Data.SqlClient
Imports Impensa.clsLibrary

' Yearly Comparision
Module mdlChart3
    Public Function PopulateListingCombo() As BindingSource
        Dim dt As New DataTable
        Dim dir As New Dictionary(Of Integer, String)
        Dim dr As DataRow
        Dim BindingSrc As New BindingSource
        Try
            If SelectChartCombo = "Chart 3A" Then
                BindingSrc = mdlChart1.PopulateListingCombo()
            ElseIf SelectChartCombo = "Chart 3B" Then
                'BindingSrc = New BindingSource
                Using Conn = GetConnection()
                    Dim da As New SqlDataAdapter("SELECT Y.iMonth [Key], Datename(M, CONVERT(DATE, '1900-' + CONVERT(VARCHAR, Y.iMonth) + '-01')) [VALUE] FROM (SELECT TOP(12) ROW_NUMBER() OVER(ORDER BY NAME) iMonth FROM sys.objects)Y", Conn)
                    da.Fill(dt)

                    dr = dt.NewRow()
                    dr.Item(0) = 0
                    dr.Item(1) = "Entire Year"
                    dt.Rows.Add(dr)

                    For Each dr1 As DataRow In dt.Rows
                        If InStr(Categories, dr1("Value")) > 0 OrElse Categories = String.Empty Then
                            dir.Add(dr1("Key"), dr1("Value"))
                        End If
                    Next

                    BindingSrc.DataSource = dir
                    BindingSrc.DataMember = Nothing
                End Using
            End If
        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try
        Return BindingSrc
    End Function

    Public Function GetChartData(ByVal P_FromDate As Date, ByVal P_ToDate As Date, ByVal P_iCategory As Integer, ByVal P_SearchStr As String, Optional ByVal P_PeriodLimit As Boolean = False) As DataTable
        Dim dt As New DataTable
        Dim Params As SqlParameter
        Dim Comm As SqlCommand
        Dim Da As SqlDataAdapter

        Try
            Using Connection = GetConnection()
                Comm = New SqlCommand("sp_GetChartData_3", Connection)
                Comm.CommandType = CommandType.StoredProcedure

                Params = New SqlParameter("@P_FromDate", P_FromDate)
                Params.Direction = ParameterDirection.Input
                Params.DbType = DbType.Date
                Comm.Parameters.Add(Params)

                Params = New SqlParameter("@P_ToDate", P_ToDate)
                Params.Direction = ParameterDirection.Input
                Params.DbType = DbType.Date
                Comm.Parameters.Add(Params)

                Params = New SqlParameter("@P_iCategory", P_iCategory)
                Params.Direction = ParameterDirection.Input
                Params.DbType = DbType.Int32
                Comm.Parameters.Add(Params)

                Params = New SqlParameter("@P_SearchStr", IIf(P_SearchStr Is Nothing, String.Empty, P_SearchStr))
                Params.Direction = ParameterDirection.Input
                Params.DbType = DbType.String
                Comm.Parameters.Add(Params)

                Params = New SqlParameter("@P_Years", ChkLBYearsItemsList)
                Params.Direction = ParameterDirection.Input
                Params.DbType = DbType.String
                Comm.Parameters.Add(Params)

                Params = New SqlParameter("@P_PeriodLimit", P_PeriodLimit)
                Params.Direction = ParameterDirection.Input
                Params.DbType = DbType.Boolean
                Comm.Parameters.Add(Params)

                Da = New SqlDataAdapter(Comm)
                Da.Fill(dt)
            End Using
        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try

        Return dt
    End Function

    Public Function GetChartData_3B(ByVal P_FromDate As Date, ByVal P_ToDate As Date, ByVal P_iMonth As Integer, ByVal P_SearchStr As String, Optional ByVal P_PeriodLimit As Boolean = False) As DataTable
        Dim dt As New DataTable
        Dim Params As SqlParameter
        Dim Comm As SqlCommand
        Dim Da As SqlDataAdapter

        Try
            Using Connection = GetConnection()
                Comm = New SqlCommand("sp_GetChartData_3B", Connection)
                Comm.CommandType = CommandType.StoredProcedure

                Params = New SqlParameter("@P_FromDate", P_FromDate)
                Params.Direction = ParameterDirection.Input
                Params.DbType = DbType.Date
                Comm.Parameters.Add(Params)

                Params = New SqlParameter("@P_ToDate", P_ToDate)
                Params.Direction = ParameterDirection.Input
                Params.DbType = DbType.Date
                Comm.Parameters.Add(Params)

                Params = New SqlParameter("@P_iMonth", P_iMonth)
                Params.Direction = ParameterDirection.Input
                Params.DbType = DbType.Int32
                Comm.Parameters.Add(Params)

                Params = New SqlParameter("@P_SearchStr", IIf(P_SearchStr Is Nothing, String.Empty, P_SearchStr))
                Params.Direction = ParameterDirection.Input
                Params.DbType = DbType.String
                Comm.Parameters.Add(Params)

                Params = New SqlParameter("@P_Years", ChkLBYearsItemsList)
                Params.Direction = ParameterDirection.Input
                Params.DbType = DbType.String
                Comm.Parameters.Add(Params)

                Params = New SqlParameter("@P_PeriodLimit", P_PeriodLimit)
                Params.Direction = ParameterDirection.Input
                Params.DbType = DbType.Boolean
                Comm.Parameters.Add(Params)

                Da = New SqlDataAdapter(Comm)
                Da.Fill(dt)
            End Using
        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try

        Return dt
    End Function
End Module
