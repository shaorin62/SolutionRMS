Public Class SCCOSAPBUSINO
    Inherits System.Web.UI.Page

#Region " Web Form 디자이너에서 생성한 코드 "

    '이 호출은 Web Form 디자이너에 필요합니다.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim configurationAppSettings As System.Configuration.AppSettingsReader = New System.Configuration.AppSettingsReader
        Dim ZgsfI0141 As SC.ZGSFI014 = New SC.ZGSFI014
        Me.Destination1 = New SAP.Connector.Destination(Me.components)
        Me.SapProxy51 = New SC.SAPProxy5
        '
        'Destination1
        '
        Me.Destination1.AppServerHost = CType(configurationAppSettings.GetValue("Destination1.AppServerHost", GetType(System.String)), String)
        Me.Destination1.Client = CType(configurationAppSettings.GetValue("Destination1.Client", GetType(System.Int16)), Short)
        Me.Destination1.HttpPort = 3300
        Me.Destination1.Password = "SKPRFCRMS1"
        Me.Destination1.SystemNumber = CType(configurationAppSettings.GetValue("Destination1.SystemNumber", GetType(System.Int16)), Short)
        Me.Destination1.Username = "RFCRMS"
        '
        'SapProxy51
        '
        Me.SapProxy51.BUSINO = ""
        ZgsfI0141.Ceo = Nothing
        ZgsfI0141.Cname = Nothing
        ZgsfI0141.Crtday = Nothing
        ZgsfI0141.Crtwho = Nothing
        ZgsfI0141.J_1kftbus = Nothing
        ZgsfI0141.J_1kftind = Nothing
        ZgsfI0141.Mcod1 = Nothing
        ZgsfI0141.Name1 = Nothing
        ZgsfI0141.Name2 = Nothing
        ZgsfI0141.Ort01 = Nothing
        ZgsfI0141.Pstlz = Nothing
        ZgsfI0141.Regio = Nothing
        ZgsfI0141.Saupno = Nothing
        ZgsfI0141.Stras = Nothing
        ZgsfI0141.Telf1 = Nothing
        Me.SapProxy51.Connection = SAP.Connector.Connection.GetConnection(Me.Destination1)

    End Sub

    Protected WithEvents txtSAPBUSINO As System.Web.UI.WebControls.TextBox
    Private components As System.ComponentModel.IContainer
    Protected WithEvents Destination1 As SAP.Connector.Destination
    Protected WithEvents txtCNT As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtRETURNLIST As System.Web.UI.WebControls.TextBox
    Protected WithEvents txtRETURNBANK As System.Web.UI.WebControls.TextBox
    Protected WithEvents strRETURNLIST As System.Web.UI.WebControls.TextBox
    Protected WithEvents strRETURNBANK As System.Web.UI.WebControls.TextBox
    Protected WithEvents RT_RETURNLIST As System.Web.UI.WebControls.TextBox
    Protected WithEvents RT_RETURNBANK As System.Web.UI.WebControls.TextBox
    Protected WithEvents SapProxy51 As SC.SAPProxy5


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

    Private Sub txtSAPBUSINO_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtSAPBUSINO.TextChanged
        '가져온 사업자 번호와 카운터를 담는다.
        Dim strBUSINO As String = Replace(Me.txtSAPBUSINO.Text(), "-", "")
        Dim intCNT As Integer = Me.txtCNT.Text
        Dim i, j


        SapProxy51.BUSINO = strBUSINO
        SapProxy51.CNT = intCNT

        SapProxy51.Zg_Mfi_Rms_Int()


        Dim strRETURNLIST = ""
        Dim strRETURNBANK = ""

        For i = 0 To SapProxy51.tblBODY.Count - 1
            If i = 0 Then
                strRETURNLIST = SapProxy51.tblBODY.Item(i).Saupno + "|" + SapProxy51.tblBODY.Item(i).Name1 + "|" + SapProxy51.tblBODY.Item(i).Cname
                strRETURNLIST = strRETURNLIST & "|" + SapProxy51.tblBODY.Item(i).Ort01 + "|" + SapProxy51.tblBODY.Item(i).Stras + "|" + SapProxy51.tblBODY.Item(i).Pstlz
                strRETURNLIST = strRETURNLIST & "|" + SapProxy51.tblBODY.Item(i).Telf1 + "|" + SapProxy51.tblBODY.Item(i).Ceo + "|" + SapProxy51.tblBODY.Item(i).J_1kftbus
                strRETURNLIST = strRETURNLIST & "|" + SapProxy51.tblBODY.Item(i).J_1kftind + "|" + SapProxy51.tblBODY.Item(i).Regio + "|" + SapProxy51.tblBODY.Item(i).Mcod1
                strRETURNLIST = strRETURNLIST & "|" + SapProxy51.tblBODY.Item(i).Crtday + "|" + SapProxy51.tblBODY.Item(i).Crtwho + "|" + SapProxy51.tblBODY.Item(i).Name2

            Else
                strRETURNLIST = strRETURNLIST + ":" + SapProxy51.tblBODY.Item(i).Saupno + "|" + SapProxy51.tblBODY.Item(i).Name1 + "|" + SapProxy51.tblBODY.Item(i).Cname
                strRETURNLIST = strRETURNLIST & "|" + SapProxy51.tblBODY.Item(i).Ort01 + "|" + SapProxy51.tblBODY.Item(i).Stras + "|" + SapProxy51.tblBODY.Item(i).Pstlz
                strRETURNLIST = strRETURNLIST & "|" + SapProxy51.tblBODY.Item(i).Telf1 + "|" + SapProxy51.tblBODY.Item(i).Ceo + "|" + SapProxy51.tblBODY.Item(i).J_1kftbus
                strRETURNLIST = strRETURNLIST & "|" + SapProxy51.tblBODY.Item(i).J_1kftind + "|" + SapProxy51.tblBODY.Item(i).Regio + "|" + SapProxy51.tblBODY.Item(i).Mcod1
                strRETURNLIST = strRETURNLIST & "|" + SapProxy51.tblBODY.Item(i).Crtday + "|" + SapProxy51.tblBODY.Item(i).Crtwho + "|" + SapProxy51.tblBODY.Item(i).Name2
            End If
        Next

        For j = 0 To SapProxy51.tblBANK.Count - 1
            If j = 0 Then
                strRETURNBANK = SapProxy51.tblBANK.Item(j).Saupno + "|" + SapProxy51.tblBANK.Item(j).Bvtyp + "|" + SapProxy51.tblBANK.Item(j).Bankl + "|" + SapProxy51.tblBANK.Item(j).Bankn + "|" + SapProxy51.tblBANK.Item(j).Koinh
            Else
                strRETURNBANK = strRETURNBANK & ":" + SapProxy51.tblBANK.Item(j).Saupno + "|" + SapProxy51.tblBANK.Item(j).Bvtyp + "|" + SapProxy51.tblBANK.Item(j).Bankl + "|" + SapProxy51.tblBANK.Item(j).Bankn + "|" + SapProxy51.tblBANK.Item(j).Koinh
            End If
        Next

        Me.RT_RETURNLIST.Text = strRETURNLIST
        Me.RT_RETURNBANK.Text = strRETURNBANK
        Me.DataBind()
        RegisterStartupScript("ButtonClickTest", "<script>Set_CustInfo</script>")
        txtSAPBUSINO.Text = ""
        txtCNT.Text = ""
    End Sub
End Class
