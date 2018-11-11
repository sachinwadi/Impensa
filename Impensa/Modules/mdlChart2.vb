' Categorywise monthly distribution
Imports System.Collections.Generic
Imports System.Data.SqlClient
Imports System.Text
Imports Impensa.clsLib

Module mdlChart2
    Public Function PopulateListingCombo() As BindingSource
        Dim dt As New DataTable
        Dim SQL As New StringBuilder
        Dim dir As New Dictionary(Of Date, String)
        Dim BindingSrc As New BindingSource
        Dim ItemCounter As Integer = -1

        Try
            ThresholdCurrentMonthIndex = -1
            SQL.AppendFormat("Select sItemsList [Value],  dtFirstDay [Key] from fn_GetAllMonthsList('{0}', '{1}')", dtpFrom, dtpTo)
            Using Conn = GetConnection()
                Dim da As New SqlDataAdapter(SQL.ToString, Conn)
                da.Fill(dt)

                For Each dr1 As DataRow In dt.Rows
                    dir.Add(dr1("Key"), dr1("Value"))
                    ItemCounter += 1
                    If dr1("Key") = DateSerial(Today.Year, Today.Month, 1) Then
                        ThresholdCurrentMonthIndex = ItemCounter
                    End If
                Next

                BindingSrc.DataSource = dir
                BindingSrc.DataMember = Nothing
            End Using
        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try

        Return BindingSrc
    End Function

    Public Function GetChartData(ByVal P_FromDate As Date, ByVal P_ToDate As Date, ByVal P_SearchStr As String) As DataTable
        Dim dt As New DataTable
        Dim Params As SqlParameter
        Dim Comm As SqlCommand
        Dim Da As SqlDataAdapter

        Try
            Using Connection = GetConnection()
                Comm = New SqlCommand("sp_GetChartData_2", Connection)
                Comm.CommandType = CommandType.StoredProcedure

                Params = New SqlParameter("@P_FromDate", P_FromDate)
                Params.Direction = ParameterDirection.Input
                Params.DbType = DbType.Date
                Comm.Parameters.Add(Params)

                Params = New SqlParameter("@P_ToDate", P_ToDate)
                Params.Direction = ParameterDirection.Input
                Params.DbType = DbType.Date
                Comm.Parameters.Add(Params)


                Params = New SqlParameter("@P_SearchStr", IIf(P_SearchStr Is Nothing, String.Empty, P_SearchStr))
                Params.Direction = ParameterDirection.Input
                Params.DbType = DbType.String
                Comm.Parameters.Add(Params)

                Params = New SqlParameter("@P_Years", ChkLBYearsItemsList)
                Params.Direction = ParameterDirection.Input
                Params.DbType = DbType.String
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
