Public Class SCRPTMAIN
    Inherits System.Web.UI.Page

#Region " Web Form �����̳ʿ��� ������ �ڵ� "

    '�� ȣ���� Web Form �����̳ʿ� �ʿ��մϴ�.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

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
        Dim ConnString
        Dim vntParam
        Dim strUrl
        Dim iRow, iLen, Path
        Dim CheckLen
        Dim ReportName
        Dim Parameter
        Dim OptGubun
        Dim vntDBParam, param


        ReportName = Request.QueryString("rpt")            'rpt ���ϸ�
        Parameter = Request.QueryString("Param")           '�Ķ���� 
        OptGubun = Request.QueryString("Opt")              '�ɼ�(��¹��:A,B)

        Path = Request.ServerVariables("PATH_TRANSLATED")  '���Path �������� 

        While (Right(Path, 1) <> "\" And Len(Path) <> 0)
            iLen = Len(Path) - 1
            Path = Left(Path, iLen)
        End While

        CheckLen = 4

        ConnString = System.Configuration.ConfigurationSettings.AppSettings("DSN") 'DB Connect���� ��������
        vntParam = Split(ConnString, ";")


        For iRow = 0 To UBound(vntParam, 1)
            iLen = Len(vntParam(iRow))

            vntParam(iRow) = Mid(vntParam(iRow), 5, iLen - CheckLen)

        Next

        For iRow = 0 To UBound(vntParam, 1) 'DB Connect ���� ���� String������ ��ȯ
            Select Case iRow
                Case 0
                    vntDBParam = vntDBParam & vntParam(iRow)
                Case Else
                    vntDBParam = vntDBParam & ":" & vntParam(iRow)

            End Select
        Next

        Path = Replace(Path, "\", "/")      '���丮 ��� ��ȯ('\'..> '/')
        Path = Mid(Path, 3, Len(Path))      '����̺�� ����(C: Or D: ..)

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
