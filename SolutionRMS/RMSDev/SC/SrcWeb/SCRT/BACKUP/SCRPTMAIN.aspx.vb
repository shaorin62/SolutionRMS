Public Class SCRPTMAIN
    Inherits System.Web.UI.Page

#Region " Web Form 디자이너에서 생성한 코드 "

    '이 호출은 Web Form 디자이너에 필요합니다.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

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
        Dim ConnString
        Dim vntParam
        Dim strUrl
        Dim iRow, iLen, Path
        Dim CheckLen
        Dim ReportName
        Dim Parameter
        Dim OptGubun
        Dim vntDBParam, param


        ReportName = Request.QueryString("rpt")            'rpt 파일명
        Parameter = Request.QueryString("Param")           '파라미터 
        OptGubun = Request.QueryString("Opt")              '옵션(출력방법:A,B)

        Path = Request.ServerVariables("PATH_TRANSLATED")  '상대Path 가져오기 

        While (Right(Path, 1) <> "\" And Len(Path) <> 0)
            iLen = Len(Path) - 1
            Path = Left(Path, iLen)
        End While

        CheckLen = 4

        ConnString = System.Configuration.ConfigurationSettings.AppSettings("DSN") 'DB Connect정보 가져오기
        vntParam = Split(ConnString, ";")


        For iRow = 0 To UBound(vntParam, 1)
            iLen = Len(vntParam(iRow))

            vntParam(iRow) = Mid(vntParam(iRow), 5, iLen - CheckLen)

        Next

        For iRow = 0 To UBound(vntParam, 1) 'DB Connect 실지 값을 String정보로 변환
            Select Case iRow
                Case 0
                    vntDBParam = vntDBParam & vntParam(iRow)
                Case Else
                    vntDBParam = vntDBParam & ":" & vntParam(iRow)

            End Select
        Next

        Path = Replace(Path, "\", "/")      '디렉토리 모양 변환('\'..> '/')
        Path = Mid(Path, 3, Len(Path))      '드라이브명 삭제(C: Or D: ..)

        strUrl = Path & "SC_RPT_MAIN.asp?" + "Con=" + vntDBParam + "&" & _
                                             "rpt=" + ReportName + "&" & _
                                             "Param=" + Parameter + "&" & _
                                             "Opt=" + OptGubun


        Response.Write("<script LANGUAGE='vbscript'> " & vbCrLf)
        Response.Write(" Sub window_onload " & vbCrLf)
        Response.Write("     window.open """ & strUrl & """" & ", ""main""" & vbCrLf)
        Response.Write(" End Sub " & vbCrLf)
        Response.Write(" </script> " & vbCrLf)
    End Sub


End Class
