Public Class SC_RPT_MAIN
    Inherits System.Web.UI.Page
    Protected WithEvents txtModuleDir As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents txtReportName As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents txtParams As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents txtOpt As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents txtDBParams As System.Web.UI.HtmlControls.HtmlInputHidden

#Region " Web Form �����̳ʿ��� ������ �ڵ� "

    '�� ȣ���� Web Form �����̳ʿ� �ʿ��մϴ�.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: �� �޼��� ȣ���� Web Form �����̳ʿ� �ʿ��մϴ�.
        '�ڵ� �����⸦ ����Ͽ� �������� ���ʽÿ�.
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '���⿡ ����� �ڵ带 ��ġ�Ͽ� �������� �ʱ�ȭ�մϴ�.

        Dim strConStr As String
        Dim iLen As Integer
        Dim vntDBParam As Object

        '-----------------------------
        ' �����ͺ��̽� ������ ���� ����
        '-----------------------------
        Dim DSN As String = ConfigurationSettings.AppSettings("DATA_SOURCE")
        
        If DSN = Nothing Then
            DSN = ConfigurationSettings.AppSettings("DSN")
        Else
            DSN = ConfigurationSettings.AppSettings(DSN)
        End If
        '----------------------------
        '�������� ����
        '----------------------------
        vntDBParam = Split(DSN, ";")

        For i As Integer = 0 To UBound(vntDBParam, 1)
            iLen = Len(vntDBParam(i))
            strConStr &= Mid(vntDBParam(i), 5, iLen - 4) & ";"
        Next

        '---------------------------
        '������ �Ķ���� ���� ����
        '---------------------------
        'txtDSN.Value = strConStr

        ''---------------------------
        ''URL ���� ����
        ''---------------------------
        'txtURLParams.Value = Request.Url.Query
    End Sub
End Class
