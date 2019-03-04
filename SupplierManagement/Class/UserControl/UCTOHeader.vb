Public Class UCTOHeader
    Private _currentDate As Date = Today.Date
    Public Property ToDate As Date
    Public Property SebPlatformDate As Date
    Public Property CurrentDate() As Date
        Get
            Return _currentDate
        End Get
        Set(ByVal value As Date)
            _currentDate = value
            initLabel()
        End Set
    End Property

    Public WriteOnly Property ProductType As String
        Set(ByVal value As String)
            TextBox1.Text = value
        End Set
    End Property


    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        initLabel()
        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Public Sub DisplayLatestUpdate()
        Label6.Text = String.Format("Latest Import date of turnover data : {0:dd-MMM-yyyy}", ToDate)
        Label8.Text = String.Format("Latest Import date of SEB Asia Platform CMMF List: {0:dd-MMM-yyyy}", SebPlatformDate)
    End Sub

    Public Sub initLabel()

       


        Label3.Text = Year(_currentDate)
        Label5.Text = Year(_currentDate) - 1
        Label7.Text = Year(_currentDate) - 2
        Label9.Text = Year(_currentDate) - 3
        Label11.Text = Year(_currentDate) - 4
        Label4.Text = String.Format("{0:MMM}", _currentDate)

        Label14.Text = String.Format("{0} vs", Year(_currentDate))
        Label13.Text = String.Format("{0}", Year(_currentDate) - 1)
        Label17.Text = String.Format("{0} vs", Year(_currentDate) - 1)
        Label16.Text = String.Format("{0}", Year(_currentDate) - 2)
        Label19.Text = String.Format("{0} vs", Year(_currentDate) - 2)
        Label18.Text = String.Format("{0}", Year(_currentDate) - 3)
        Label21.Text = String.Format("{0} vs", Year(_currentDate) - 3)
        Label20.Text = String.Format("{0}", Year(_currentDate) - 4)
    End Sub

    Private Sub UCTOHeader_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
