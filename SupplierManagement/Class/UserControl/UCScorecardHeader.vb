Public Class UCScorecardHeader
    Inherits UCTOHeader

    Public Property QualityDate As Date
    Public Property LogisticDate As Date
    Public Property ProjectDate As Date
    Public Property QLatestPeriod As String

    Public Overloads Sub DisplayLatestUpdate()
        Label6.Text = String.Format("Latest Import date of Quality data : {0:dd-MMM-yyyy} (YTD {1:yyyy MMM})", QualityDate, CDate(String.Format("{0}-{1}-1", QLatestPeriod.Substring(0, 4), QLatestPeriod.Substring(4, 2))))
        Label8.Text = String.Format("Latest Import date of Logistics data : {0:dd-MMM-yyyy}", LogisticDate)
        Label10.Text = String.Format("Latest Import date of Project Data: {0:dd-MMM-yyyy}", ProjectDate)
    End Sub

End Class
