Imports System.Data.SqlClient
Imports Impensa.clsLib

Public Class frmLogin

    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click
        Dim TestConnStr As String = ""

        Try
            TestConnStr = "Data Source=" & txtServer.Text & ";Initial Catalog=" & txtDatabase.Text & ";User ID=" & txtUserName.Text & ";Password=" & txtPassword.Text
            Dim TestConn As New SqlConnection(TestConnStr)
            TestConn.Open()
            If TestConn.State = ConnectionState.Open Then
                TestConn.Close()
                My.Computer.Registry.CurrentUser.CreateSubKey("Impensa")
                My.Computer.Registry.SetValue("HKEY_CURRENT_USER\Software\Impensa", "ConnStr", TestConnStr)

                ImpensaAlert("Database connection established.", MsgBoxStyle.Information)

                IsLoginDetailsChanged = True

                If frmMain.TabControl1.SelectedIndex <> 0 Then frmMain.TabControl1.SelectedIndex = 0
            End If

            Me.Close()

        Catch sqlEx As SqlException
            ImpensaAlert("Database connection could not be established", MsgBoxStyle.Critical)
            IsLoginDetailsChanged = False
            txtUserName.Focus()
        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub frmLogin_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim RegConnStrVal = My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\Impensa", "ConnStr", Nothing)
        Dim UNIndex As Integer = 0
        Dim PWDIndex As Integer = 0
        Dim SerIndex As Integer = 0
        Dim DBIndex As Integer = 0

        Try
            If Not RegConnStrVal Is Nothing Then
                UNIndex = InStr(RegConnStrVal.ToString, "User ID=")
                PWDIndex = InStr(RegConnStrVal.ToString, "Password=")
                SerIndex = InStr(RegConnStrVal.ToString, "Data Source=")
                DBIndex = InStr(RegConnStrVal.ToString, "Initial Catalog=")

                txtUserName.Text = RegConnStrVal.ToString.Substring((UNIndex + 7), ((PWDIndex - 2) - (UNIndex + 7)))
                txtPassword.Text = RegConnStrVal.ToString.Substring((PWDIndex + 8))
                txtServer.Text = RegConnStrVal.ToString.Substring((SerIndex + 11), ((DBIndex - 2) - (SerIndex + 11)))
                txtDatabase.Text = RegConnStrVal.ToString.Substring((DBIndex + 15), ((UNIndex - 2) - (DBIndex + 15)))
            End If
        Catch ex As Exception
            ImpensaAlert(ex.Message, MsgBoxStyle.Critical)
        End Try


    End Sub
End Class