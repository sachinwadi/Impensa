Public NotInheritable Class Splash
    Private Sub Splash_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If My.Computer.Registry.CurrentUser.OpenSubKey("Software\Impensa") Is Nothing Then
            Me.Close()
        End If
    End Sub
End Class
