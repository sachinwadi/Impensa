Imports System.Runtime.CompilerServices
Imports System.Security.Cryptography
Imports System.Text

Module mdlCryptoExtensions
    <Extension()>
    Public Function Encrypt(ByVal text As String) As String
        Return Convert.ToBase64String(ProtectedData.Protect(Encoding.Unicode.GetBytes(text), Nothing, DataProtectionScope.CurrentUser))
    End Function

    <Extension()>
    Public Function Decrypt(ByVal text As String) As String
        Return Encoding.Unicode.GetString(ProtectedData.Unprotect(Convert.FromBase64String(text), Nothing, DataProtectionScope.CurrentUser))
    End Function
End Module
