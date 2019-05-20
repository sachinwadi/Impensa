Imports System.Collections.Generic
Imports System.Data.SqlClient
Imports Impensa.clsLibrary

' Monthly Comparision Per Category
Module mdlChart1
    Public Function PopulateListingCombo(Optional ByVal P_Function As String = "Chart") As BindingSource
        Dim dt As New DataTable
        Dim dr As DataRow
        Dim dir As New Dictionary(Of Integer, String)
        Dim BindingSrc As New BindingSource

        Try
            Using Conn = GetConnection()
                Dim da As New SqlDataAdapter("SELECT hKey [Key], sCategory [Value] FROM tbl_CategoryList WHERE IsObsolete = 0 ORDER BY sCategory", Conn)
                da.Fill(dt)
                If P_Function = "Chart" Then
                    dr = dt.NewRow()
                    dr.Item(0) = 0
                    dr.Item(1) = "All Categories"
                    dt.Rows.Add(dr)
                End If

                For Each dr1 As DataRow In dt.Rows
                    If InStr(Categories, dr1("Value")) > 0 OrElse Categories = String.Empty Then
                        dir.Add(dr1("Key"), dr1("Value"))
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

    Public Function GetChartData(ByVal P_FromDate As Date, ByVal P_ToDate As Date, ByVal P_iCategory As Integer, ByVal P_SearchStr As String) As DataTable
        Dim dt As New DataTable
        Dim Params As SqlParameter
        Dim Comm As SqlCommand
        Dim Da As SqlDataAdapter

        Try
            Using Connection = GetConnection()
                Comm = New SqlCommand("sp_GetChartData_1", Connection)
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

                Da = New SqlDataAdapter(Comm)
                Da.Fill(dt)
            End Using
        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try

        Return dt

    End Function
End Module
