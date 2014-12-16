Public Class SC_RPT_MAIN
    Inherits System.Web.UI.Page
    Protected WithEvents txtModuleDir As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents txtReportName As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents txtParams As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents txtOpt As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents txtDBParams As System.Web.UI.HtmlControls.HtmlInputHidden

#Region " Web Form 디자이너에서 생성한 코드 "

    '이 호출은 Web Form 디자이너에 필요합니다.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 이 메서드 호출은 Web Form 디자이너에 필요합니다.
        '코드 편집기를 사용하여 수정하지 마십시오.
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '여기에 사용자 코드를 배치하여 페이지를 초기화합니다.

        Dim strConStr As String
        Dim iLen As Integer
        Dim vntDBParam As Object

        '-----------------------------
        ' 데이터베이스 연결점 정보 설정
        '-----------------------------
        Dim DSN As String = ConfigurationSettings.AppSettings("DATA_SOURCE")
        
        If DSN = Nothing Then
            DSN = ConfigurationSettings.AppSettings("DSN")
        Else
            DSN = ConfigurationSettings.AppSettings(DSN)
        End If
        '----------------------------
        '연결정보 조작
        '----------------------------
        vntDBParam = Split(DSN, ";")

        For i As Integer = 0 To UBound(vntDBParam, 1)
            iLen = Len(vntDBParam(i))
            strConStr &= Mid(vntDBParam(i), 5, iLen - 4) & ";"
        Next

        '---------------------------
        '연결점 파라미터 생성 셋팅
        '---------------------------
        'txtDSN.Value = strConStr

        ''---------------------------
        ''URL 정보 셋팅
        ''---------------------------
        'txtURLParams.Value = Request.Url.Query
    End Sub
End Class
