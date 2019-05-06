Imports System.Text
Imports Impensa.clsLib
Imports System.Net
Imports System.Net.Mail
Imports System.Globalization

Public Class EmailGenerator

    Private _strBuilder As StringBuilder = New StringBuilder()
    Private _collection As Dictionary(Of Integer, String) = New Dictionary(Of Integer, String)
    Private _Changes As DataTable
    Private _ChangeSummary As DataTable

    Public Property Changes() As DataTable
        Get
            Return _Changes
        End Get
        Set(ByVal value As DataTable)
            _Changes = value
        End Set
    End Property

    Public Property ChangeSummary() As DataTable
        Get
            Return _ChangeSummary
        End Get
        Set(ByVal value As DataTable)
            _ChangeSummary = value
        End Set
    End Property

    Private Sub BuildDetailHtmlTable()
        _strBuilder.Append("<table>")
        _strBuilder.Append("<tr>")
        _strBuilder.Append("<td>")
        _strBuilder.Append("<table style='border:1px solid black;border-collapse:collapse; font-family: Arial; font-size:14;'>")
        _strBuilder.Append("<tr>")
        _strBuilder.Append("<th style='border: 1px solid black; padding: 3px;' align='left'>Date</th>")
        _strBuilder.Append("<th style='border: 1px solid black; padding: 3px;' align='left'>Category</th>")
        _strBuilder.Append("<th style='border: 1px solid black; padding: 3px;' align='right'>Amount</th>")
        _strBuilder.Append("<th style='border: 1px solid black; padding: 3px;' align='left'>Notes</th>")
        _strBuilder.Append("</tr>")

        For Each dr As DataRow In _Changes.Rows
            Dim color As String = "black"

            If dr("Action") = "Add" Then
                color = "blue"
            ElseIf dr("Action") = "Update" Then
                color = "orange"
            ElseIf dr("Action") = "Delete" Then
                color = "red"
            End If

            If Not _collection.ContainsKey(Convert.ToInt32(dr("iCategory"))) Then _collection.Add(Convert.ToInt32(dr("iCategory")), dr("Action"))

            _strBuilder.Append("<tr>")
            _strBuilder.AppendFormat("<td style='color: {1}; border: 1px solid black; padding: 3px;' align='left'>{0}</td>", dr("Date"), color)
            _strBuilder.AppendFormat("<td style='color: {1}; border: 1px solid black; padding: 3px;' align='left'>{0}</td>", dr("CategoryName"), color)
            _strBuilder.AppendFormat("<td style='color: {1}; border: 1px solid black; padding: 3px;' align='right'>{0}</td>", Format(dr("Amount"), "#,##0.00"), color)
            _strBuilder.AppendFormat("<td style='color: {1}; border: 1px solid black; padding: 3px;' align='left'>{0}</td>", dr("Notes"), color)
            _strBuilder.Append("<tr>")
        Next

        _strBuilder.Append("</table>")
        _strBuilder.Append("</td>")
        _strBuilder.Append("</tr>")
    End Sub

    Private Sub BuildSummaryHtmlTable()

        _strBuilder.Append("<table style='border:1px solid black;border-collapse:collapse;font-family: Arial;font-size:14;'>")
        _strBuilder.Append("<tr>")
        _strBuilder.Append("<th style='border: 1px solid black; padding: 3px;' align='left'>Category</th>")
        _strBuilder.Append("<th style='border: 1px solid black; padding: 3px;' align='right'>MTD</th>")
        _strBuilder.Append("<th style='border: 1px solid black; padding: 3px;' align='right'>YTD</th>")
        _strBuilder.Append("<th style='border: 1px solid black; padding: 3px;' align='right'>ITD</th>")
        _strBuilder.Append("</tr>")

        For Each dr As DataRow In _ChangeSummary.Rows
            Dim color As String = "black"

            If _collection.ContainsKey(Convert.ToInt32(dr("iCategory"))) Then
                If _collection.Item(Convert.ToInt32(dr("iCategory"))) = "Add" Then
                    color = "blue"
                ElseIf _collection.Item(Convert.ToInt32(dr("iCategory"))) = "Update" Then
                    color = "orange"
                ElseIf _collection.Item(Convert.ToInt32(dr("iCategory"))) = "Delete" Then
                    color = "red"
                End If
            End If

            Dim fontWeight As String = "normal"
            Dim fontSize As String = "14"

            If (dr("Category") = "TOTAL") Then
                fontWeight = "bold"
                fontSize = "18"
            End If

            _strBuilder.Append("<tr>")
            _strBuilder.AppendFormat("<td style='color: {1}; border: 1px solid black; padding: 3px; font-weight: {2}; font-size: {3}' align='left'>{0}</td>", dr("Category"), color, fontWeight, fontSize)
            _strBuilder.AppendFormat("<td style='color: {1}; border: 1px solid black; padding: 3px; font-weight: {2}; font-size: {3}' align='right'>{0}</td>", Format(dr("MTD"), "#,##0.00"), color, fontWeight, fontSize)
            _strBuilder.AppendFormat("<td style='color: {1}; border: 1px solid black; padding: 3px; font-weight: {2}; font-size: {3}' align='right'>{0}</td>", Format(dr("YTD"), "#,##0.00"), color, fontWeight, fontSize)
            _strBuilder.AppendFormat("<td style='color: {1}; border: 1px solid black; padding: 3px; font-weight: {2}; font-size: {3}' align='right'>{0}</td>", Format(dr("ITD"), "#,##0.00"), color, fontWeight, fontSize)
            _strBuilder.Append("</tr>")
        Next

        _strBuilder.Append("</table>")
    End Sub

    Private Sub BuildLegendHtmlTable()
        _strBuilder.Append("<tr>")
        _strBuilder.Append("<td>")
        _strBuilder.Append("<table style='margin-top: 5px;'>")
        _strBuilder.Append("<tr>")
        _strBuilder.Append("<td style='border: 1px solid black; padding: px; width: 15px; background-color:blue'></td>")
        _strBuilder.Append("<td>Added</td>")
        _strBuilder.Append("<td style='border: 1px solid black; padding: 1px; width: 15px; background-color:orange'></td>")
        _strBuilder.Append("<td>Updated</td>")
        _strBuilder.Append("<td style='border: 1px solid black; padding: 1px; width: 15px; background-color:red'></td>")
        _strBuilder.Append("<td>Deleted</td>")
        _strBuilder.Append("</tr>")
        _strBuilder.Append("<table>")
        _strBuilder.Append("</td>")
        _strBuilder.Append("</tr>")
        _strBuilder.Append("</table>")
    End Sub

    Private Sub BuildEmailBody()
        _strBuilder.Append("<p>Hello,<br /><br />This is a notification e-mail from <strong>Impensa Expense Manager</strong></p>")
        _strBuilder.Append("<h4>Here are the Changes:</h4>")
        Call BuildDetailHtmlTable()
        Call BuildLegendHtmlTable()
        If (IncludeExpenseSummary) Then
            _strBuilder.Append("<h4>Summary:</h4>")
            Call BuildSummaryHtmlTable()
        End If
        _strBuilder.Append("<p>Thanks,<br />Team Impensa</p>")
    End Sub

    Public Sub SendEmail()
        Dim fromAddress = New MailAddress(FromEmail, "Impensa Expense Manager")

        Call BuildEmailBody()

        Dim smtp = New SmtpClient() With
        {.Host = SmtpHost,
        .Port = SmtpPort,
        .EnableSsl = True,
        .DeliveryMethod = SmtpDeliveryMethod.Network,
        .UseDefaultCredentials = False,
        .Credentials = New NetworkCredential(fromAddress.Address, FromPassword)
        }

        Dim message = New MailMessage() With
        {.IsBodyHtml = True,
         .Subject = "Impensa Notification - " + DateTime.Now.Date.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture),
         .Body = _strBuilder.ToString,
         .From = fromAddress
        }

        For Each addr As String In ToEmails.Split(";")
            message.Bcc.Add(addr)
        Next

        smtp.Send(message)
    End Sub

End Class
