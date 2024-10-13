Imports System.Data.SqlClient
Imports Impensa.clsLibrary

' Yearly Comparision
Module mdlChart4
    Public Function PopulateListingCombo() As BindingSource
        Dim BindingSrc As BindingSource
        BindingSrc = mdlChart1.PopulateListingCombo()
        Return BindingSrc
    End Function

    Public Function GetChartData(ByVal P_FromDate As Date, ByVal P_ToDate As Date, ByVal P_iCategory As Integer, ByVal P_SearchStr As String, Optional ByVal P_PeriodLimit As Boolean = False) As DataTable
        Dim dt As New DataTable
        Dim Params As SqlParameter
        Dim Comm As SqlCommand
        Dim Da As SqlDataAdapter

        Try
            Using Connection = GetConnection()
                Comm = New SqlCommand("sp_GetChartData_4", Connection)
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

                If ChartTypeCombo.value = "Pie" Then
                    Params = New SqlParameter("@P_ExcludeZeroEntry", True)
                    Params.Direction = ParameterDirection.Input
                    Params.DbType = DbType.Boolean
                    Comm.Parameters.Add(Params)
                End If

                Da = New SqlDataAdapter(Comm)
                Da.Fill(dt)
            End Using
        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try

        Return dt

    End Function
End Module
