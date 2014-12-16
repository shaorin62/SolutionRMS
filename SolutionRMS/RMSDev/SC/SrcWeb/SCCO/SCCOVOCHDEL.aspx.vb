Public Class SCCOVOCHDEL
    Inherits System.Web.UI.Page

#Region " Web Form 디자이너에서 생성한 코드 "

    '이 호출은 Web Form 디자이너에 필요합니다.
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

    '참고: 다음의 자리 표시자 선언은 Web Form 디자이너의 필수 선언입니다.
    '삭제하거나 옮기지 마십시오.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 이 메서드 호출은 Web Form 디자이너에 필요합니다.
        '코드 편집기를 사용하여 수정하지 마십시오.
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '여기에 사용자 코드를 배치하여 페이지를 초기화합니다.
    End Sub

    Private Sub ImageButton1_Click(ByVal sender As System.Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImageButton1.Click
        SapProxy31.YEARMON = Me.TextBox1.Text
        SapProxy31.VOCHNO = Me.TextBox2.Text
        SapProxy31.Z_Cfi_Rms_Park_Del()
        Me.DataBind()
    End Sub

End Class
