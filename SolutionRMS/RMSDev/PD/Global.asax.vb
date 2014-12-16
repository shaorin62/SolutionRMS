Imports System.Web
Imports System.Web.SessionState

Public Class Global
    Inherits System.Web.HttpApplication

#Region " 구성 요소 디자이너에서 생성한 코드 "

    Public Sub New()
        MyBase.New()

        '이 호출은 구성 요소 디자이너에 필요합니다.
        InitializeComponent()

        'InitializeComponent()를 호출한 다음에 초기화 작업을 추가하십시오.

    End Sub

    '구성 요소 디자이너에 필요합니다.
    Private components As System.ComponentModel.IContainer

    '참고: 다음 프로시저는 구성 요소 디자이너에 필요합니다.
    '구성 요소 디자이너를 사용하여 수정할 수 있습니다.
    '코드 편집기를 사용하여 수정하지 마십시오.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        components = New System.ComponentModel.Container()
    End Sub

#End Region

    Sub Application_Start(ByVal sender As Object, ByVal e As EventArgs)
        ' 응용 프로그램이 시작될 때 발생합니다.
    End Sub

    Sub Session_Start(ByVal sender As Object, ByVal e As EventArgs)
        ' 세션이 시작될 때 발생합니다.
    End Sub

    Sub Application_BeginRequest(ByVal sender As Object, ByVal e As EventArgs)
        ' 각 요청을 시작할 때 발생합니다.
    End Sub

    Sub Application_AuthenticateRequest(ByVal sender As Object, ByVal e As EventArgs)
        ' 사용을 인증하려고 할 때 발생합니다.
    End Sub

    Sub Application_Error(ByVal sender As Object, ByVal e As EventArgs)
        ' 오류가 일어날 때 발생합니다.
    End Sub

    Sub Session_End(ByVal sender As Object, ByVal e As EventArgs)
        ' 세션이 끝날 때 발생합니다.
    End Sub

    Sub Application_End(ByVal sender As Object, ByVal e As EventArgs)
        ' 응용 프로그램이 끝날 때 발생합니다.
    End Sub

End Class
