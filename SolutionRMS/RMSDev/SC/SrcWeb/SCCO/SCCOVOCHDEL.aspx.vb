Public Class SCCOVOCHDEL
    Inherits System.Web.UI.Page

#Region " Web Form �����̳ʿ��� ������ �ڵ� "

    '�� ȣ���� Web Form �����̳ʿ� �ʿ��մϴ�.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim ZmsfI_RMS_PARK_RETURN1 As SC.ZMSFI_RMS_PARK_RETURN = New SC.ZMSFI_RMS_PARK_RETURN
        Dim configurationAppSettings As System.Configuration.AppSettingsReader = New System.Configuration.AppSettingsReader
        Me.SapProxy31 = New SC.SAPProxy3
        Me.Destination1 = New SAP.Connector.Destination(Me.components)
        '
        'SapProxy31
        '
        ZmsfI_RMS_PARK_RETURN1.Belnr = Nothing
        ZmsfI_RMS_PARK_RETURN1.Gjahr = Nothing
        ZmsfI_RMS_PARK_RETURN1.Subrc = Nothing
        ZmsfI_RMS_PARK_RETURN1.Text = Nothing
        Me.SapProxy31.VOCHNO = ""
        Me.SapProxy31.YEARMON = ""
        '
        'Destination1
        '
        Me.Destination1.AppServerHost = CType(configurationAppSettings.GetValue("Destination1.AppServerHost", GetType(System.String)), String)
        Me.Destination1.Client = CType(configurationAppSettings.GetValue("Destination1.Client", GetType(System.Int16)), Short)
        Me.Destination1.Password = "MARKET"
        Me.Destination1.SystemNumber = CType(configurationAppSettings.GetValue("Destination1.SystemNumber", GetType(System.Int16)), Short)
        Me.Destination1.Username = "SKRMS01"
        Me.SapProxy31.Connection = SAP.Connector.Connection.GetConnection(Me.Destination1)

    End Sub
    Protected WithEvents ImageButton1 As System.Web.UI.WebControls.ImageButton
    Protected WithEvents TextBox1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents TextBox2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents SapProxy31 As SC.SAPProxy3
    Private components As System.ComponentModel.IContainer
    Protected WithEvents Destination1 As SAP.Connector.Destination

    '����: ������ �ڸ� ǥ���� ������ Web Form �����̳��� �ʼ� �����Դϴ�.
    '�����ϰų� �ű��� ���ʽÿ�.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: �� �޼��� ȣ���� Web Form �����̳ʿ� �ʿ��մϴ�.
        '�ڵ� �����⸦ ����Ͽ� �������� ���ʽÿ�.
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '���⿡ ����� �ڵ带 ��ġ�Ͽ� �������� �ʱ�ȭ�մϴ�.
    End Sub

    Private Sub ImageButton1_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImageButton1.Click
        SapProxy31.YEARMON = Me.TextBox1.Text
        SapProxy31.VOCHNO = Me.TextBox2.Text
        SapProxy31.Z_Cfi_Rms_Park_Del()
        Me.DataBind()
    End Sub

End Class
