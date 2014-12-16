
'****************************************************************************************
'�ý��۱��� : ǥ�ػ���/Server�� Control Component
'����  ȯ�� : COM+ Service Server Package
'���α׷��� : ccPDCMGET.vb (�����ڵ� ��ȸ Control Class)
'��      �� : �����ڵ� ��ȸ�� ���� Ŭ����
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008/08/14 By Kim Tae Ho
'           :2) 2008/11/21 By Kim Tae Ho
'****************************************************************************************
Imports SCGLControl                 'Control Class�� Base Class
Imports SCGLUtil.cbSCGLConfig       'Configuration Ŭ����
Imports SCGLUtil.cbSCGLErr          '����ó�� Ŭ����
Imports SCGLUtil.cbSCGLXml          'XMLó�� Ŭ����
Imports SCGLUtil.cbSCGLUtil         '��Ÿ ��ƿ��Ƽ Ŭ����

Public Class ccPDCOGET
    Inherits ccControl

#Region "GROUP BLOCk : ���� �Ǵ� ��ⷹ���� ����/��� ����"
    Private Const CLASS_NAME = "ccPDCOGET"    '�ڽ��� Ŭ������
    'Private Const .DBConnStr = "Provider=SQLOLEDB;Data Source=10.110.10.86;Initial Catalog=MCDEV;DSN=MCDEV;UID=devadmin;Pwd=password"
#End Region

#Region "GROUP BLOCk : Property ����"
#End Region

#Region "GROUP BLOCk : Event ����"
#End Region

#Region "GROUP BLOCk : �ܺο� ���� Method"
    '*****************************************************************
    '�Է� : strInfoXML = ����⺻���� XML
    '       intRowCnt,intColCnt = ��ȸ �Ǽ�,�ʵ� ��
    '       strCLASS_CODE = ��ȸ�ϰ����ϴ� CLASS_CODE
    '       strCODE_NAME = ��ȸ�ϰ����ϴ� �ڵ� �Ǵ� ��
    '       blnUseOnly = ���� ������� �͸� �Ǵ� ��ü
    '       strAddFields = �ڵ�/�� �̿��� ��ȸ �߰� �ʵ�
    '       blnLikeCode  = �ڵ带 Like�� ��ȸ���� ����
    '��ȯ : ó�����
    '��� : SC_CODE�� ��ȸ�ϱ����� �Լ� (Ŭ���� �ڵ带 ��ȸ)
    '*****************************************************************
#Region "1. SC_CODE: Ŭ���� �ڵ� ��ȸ"

    Public Function GetSC_CODE(ByVal strInfoXML As String, _
                               ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                               ByVal strCLASS_CODE As String, _
                               Optional ByVal strCODE_NAME As String = "", _
                               Optional ByVal blnUseOnly As Boolean = True, _
                               Optional ByVal strAddFields As String = "", _
                               Optional ByVal blnLikeCode As Boolean = True) As Object
        Dim strSQL As String
        Dim strFields, strCondition As String
        Dim strChkDate As String = ""
        Dim vntData As Object



        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                If strCODE_NAME <> "" Then
                    '��ȸ ������ �ڵ����� �ڵ������ �����Ͽ� ���� �ʵ� ����
                    If IsNumeric(strCODE_NAME) Then     '������ ���
                        If Not blnLikeCode Then
                            strCondition = String.Format("AND CODE='{0}'", strCODE_NAME)
                        Else
                            strCondition = String.Format("AND CODE LIKE '%{0}%'", strCODE_NAME)
                        End If
                    ElseIf IsSBCS(strCODE_NAME) Then    '������ ���
                        strCondition = String.Format("AND (CODE LIKE '%{0}%' OR CODE_NAME LIKE '%{0}%')", strCODE_NAME)
                    Else                                '�ѱ��� ���
                        strCondition = String.Format("AND CODE_NAME LIKE '%{0}%'", strCODE_NAME)
                    End If
                End If

                '������� ���� �˻�
                If blnUseOnly Then
                    strChkDate = "AND USE_YN='Y'"
                Else    '(����)����ȵ�
                    'strChkDate = String.Format("AND (B.USE_YN='Y' OR B.EDATE>={0})", cbSCGLUtil.BuildToDate(strUseDate))
                End If

                '�߰� ��ȸ �ʵ� ���� �˻�
                strFields = "CODE, CODE_NAME"
                If strAddFields <> "" Then strFields &= "," & strAddFields

                strSQL = String.Format("SELECT {0} FROM SC_CODE WHERE CLASS_CODE='{1}' {2} AND SC_BU_CODE='{3}' {4} ORDER BY SORT_SEQ", _
                                        strFields, strCLASS_CODE, strCondition, .USRCompany, strChkDate)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetSC_CODE")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#End Region

    '*****************************************************************
    '�Է� : strInfoXML = ����⺻���� XML
    '       intRowCnt,intColCnt = ��ȸ �Ǽ�,�ʵ� ��
    '       strLOC_CODE = �����ڵ�
    '       strMU_CODE = MU �ڵ�
    '       strCC_CODE = CC �ڵ�
    '       strCODE_NAME = ��ȸ�ϰ����ϴ� �ڵ� �Ǵ� ��
    '       blnUseOnly = ���� ������� �͸� �Ǵ� ��ü
    '       strUseDate = blnUseOnly�� False�϶� EDATE>=strUseDate�� �˻�
    '       strAddFields = �ڵ�/�� �̿��� ��ȸ �߰� �ʵ�
    '       blnLikeCode  = �ڵ带 Like�� ��ȸ���� ����
    '��ȯ : ó�����
    '��� : PU CODE�� ��ȸ�ϱ����� �Լ�
    '*****************************************************************
#Region "2. PU_CODE: PU �ڵ���ȸ"

    Public Function GetPU(ByVal strInfoXML As String, _
                          ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                          Optional ByVal strLOC_CODE As String = "", _
                          Optional ByVal strMU_CODE As String = "", _
                          Optional ByVal strCC_CODE As String = "", _
                          Optional ByVal strCODE_NAME As String = "", _
                          Optional ByVal blnUseOnly As Boolean = True, _
                          Optional ByVal strUseDate As String = "", _
                          Optional ByVal strAddFields As String = "", _
                          Optional ByVal blnLikeCode As Boolean = True) As Object
        Dim strSQL, strFormat, strSelFields As String
        Dim strCondition As String
        Dim strChkDate As String = ""
        Dim vntData As Object

        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig
            If strCODE_NAME <> "" Then
                '��ȸ ������ �ڵ����� �ڵ������ �����Ͽ� ���� �ʵ� ����
                If IsNumeric(strCODE_NAME) Then '������ ���
                    If Not blnLikeCode Then
                        strCondition = String.Format("AND B.PU_CODE='{0}'", strCODE_NAME)
                    Else
                        strCondition = String.Format("AND B.PU_CODE LIKE '%{0}%'", strCODE_NAME)
                    End If
                ElseIf IsSBCS(strCODE_NAME) Then    '������ ���
                    strCondition = String.Format("AND (B.PU_CODE LIKE '%{0}%' OR B.PU_NAME LIKE '%{0}%')", strCODE_NAME)
                Else                                '�ѱ��� ���
                    strCondition = String.Format("AND B.PU_NAME LIKE '%{0}%'", strCODE_NAME)
                End If
            End If

            '������� ���� �˻�
            If blnUseOnly Then
                strChkDate = "AND B.USE_YN='Y'"
            Else
                strChkDate = String.Format("AND (B.USE_YN='Y' OR B.EDATE>={0})", BuildToDate(strUseDate))
            End If

            '��ȸ �ʵ� ����
            If strAddFields <> "" Then strAddFields = "," & AddAlias(strAddFields, "B")
            strSelFields = "B.PU_CODE,B.PU_NAME" & strAddFields

            'SQL�� ����
            If strCC_CODE <> "" Then
                'CC�� �ִ� ���
                strFormat = "SELECT {0} FROM SC_CC_PU_V A,SC_PU_V B " & _
                             "WHERE A.SC_BU_CODE='{1}' AND A.CC_CODE='{2}' AND A.SC_BU_CODE=B.SC_BU_CODE AND A.PU_CODE=B.PU_CODE {3} {4} " & _
                             "ORDER BY B.PU_CODE"
                strSQL = String.Format(strFormat, _
                                       strSelFields, .USRCompany, strCC_CODE, strChkDate, strCondition)
            ElseIf strMU_CODE <> "" Then
                'MU�� �ִ� ���
                strFormat = "SELECT {0} FROM SC_MU_PU_V A,SC_PU_V B " & _
                             "WHERE A.SC_BU_CODE='{1}' AND A.SC_MU_CODE='{2}' AND A.SC_BU_CODE=B.SC_BU_CODE AND A.PU_CODE=B.PU_CODE {3} {4} " & _
                             "ORDER BY B.PU_CODE"
                strSQL = String.Format(strFormat, _
                                       strSelFields, .USRCompany, strMU_CODE, strChkDate, strCondition)
            ElseIf strLOC_CODE <> "" Then
                'LOC�� �ִ� ���
                strFormat = "SELECT {0} FROM SC_PLANT_PU_V A, SC_PU_V B " & _
                             "WHERE A.SC_BU_CODE='{1}' AND A.LOC_CODE='{2}' AND A.SC_BU_CODE=B.SC_BU_CODE AND A.PU_CODE=B.PU_CODE {3} {4} " & _
                             "ORDER BY B.PU_CODE"
                strSQL = String.Format(strFormat, _
                                       strSelFields, .USRCompany, strLOC_CODE, strChkDate, strCondition)
            Else
                '������ ���� ���
                strFormat = "SELECT {0} FROM SC_PU_V B " & _
                             "WHERE B.SC_BU_CODE='{1}' {2} {3} " & _
                             "ORDER BY B.PU_CODE"
                strSQL = String.Format(strFormat, _
                                       strSelFields, .USRCompany, strChkDate, strCondition)
            End If

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetPU")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#End Region

    '*****************************************************************
    '�Է� : strInfoXML = ����⺻���� XML
    '       intRowCnt,intColCnt = ��ȸ �Ǽ�,�ʵ� ��
    '       strPU_CODE = PU �ڵ�
    '       strCC_CODE = CC �ڵ�
    '       strCODE_NAME = ��ȸ�ϰ����ϴ� MU �ڵ� �Ǵ� ��
    '       blnUseOnly = ���� ������� �͸� �Ǵ� ��ü
    '       strUseDate = ���⼭�� �ǹ� ����
    '       strAddFields = �ڵ�/�� �̿��� ��ȸ �߰� �ʵ�
    '       blnLikeCode  = �ڵ带 Like�� ��ȸ���� ����
    '��ȯ : ó�����
    '��� : MU CODE�� ��ȸ�ϱ����� �Լ�
    '*****************************************************************
#Region "3. MU_CODE: MU �ڵ���ȸ"

    Public Function GetMU(ByVal strInfoXML As String, _
                          ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                          Optional ByVal strPU_CODE As String = "", _
                          Optional ByVal strCC_CODE As String = "", _
                          Optional ByVal strCODE_NAME As String = "", _
                          Optional ByVal blnUseOnly As Boolean = True, _
                          Optional ByVal strUseDate As String = "", _
                          Optional ByVal strAddFields As String = "", _
                          Optional ByVal blnLikeCode As Boolean = True) As Object
        Dim strSQL, strFormat, strSelFields As String
        Dim strCondition As String
        Dim strChkDate As String = ""
        Dim vntData As Object

        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig
            If strCODE_NAME <> "" Then
                '��ȸ ������ �ڵ����� �ڵ������ �����Ͽ� ���� �ʵ� ����
                If IsNumeric(strCODE_NAME) Then '������ ���
                    If Not blnLikeCode Then
                        strCondition = String.Format("AND B.CODE='{0}'", strCODE_NAME)
                    Else
                        strCondition = String.Format("AND B.CODE LIKE '%{0}%'", strCODE_NAME)
                    End If
                ElseIf IsSBCS(strCODE_NAME) Then    '������ ���
                    strCondition = String.Format("AND (B.CODE LIKE '%{0}%' OR B.CODE_NAME LIKE '%{0}%')", strCODE_NAME)
                Else                                    '�ѱ��� ���
                    strCondition = String.Format("AND B.CODE_NAME LIKE '%{0}%'", strCODE_NAME)
                End If
            End If

            '������� ���� �˻�
            If blnUseOnly Then
                strChkDate = "AND B.USE_YN='Y'"
            End If

            '��ȸ �ʵ� ����
            If strAddFields <> "" Then strAddFields = "," & AddAlias(strAddFields, "B")
            strSelFields = "B.CODE,B.CODE_NAME" & strAddFields

            'SQL�� ����
            If strPU_CODE <> "" Then      'PU�� �ִ� ���
                strFormat = "SELECT {0} FROM SC_MU_PU_V A,SC_MU_V B " & _
                             "WHERE A.SC_BU_CODE='{1}' AND A.PU_CODE='{2}' AND A.SC_BU_CODE=B.SC_BU_CODE AND A.SC_MU_CODE=B.CODE {3} {4} " & _
                             "ORDER BY B.SORT_SEQ"
                strSQL = String.Format(strFormat, _
                                       strSelFields, .USRCompany, strPU_CODE, strChkDate, strCondition)
            ElseIf strCC_CODE <> "" Then  'CC�� �ִ� ���
                strFormat = "SELECT {0} FROM SC_MU_PU_V A,SC_MU_V B,SC_CC_PU_V C " & _
                             "WHERE C.SC_BU_CODE='{1}' AND C.CC_CODE='{2}' " & _
                               "AND A.PU_CODE=C.PU_CODE AND A.SC_BU_CODE=C.SC_BU_CODE " & _
                               "AND A.SC_MU_CODE=B.CODE AND B.SC_BU_CODE=C.SC_BU_CODE {3} {4} " & _
                             "ORDER BY B.SORT_SEQ"
                strSQL = String.Format(strFormat, _
                                       strSelFields, .USRCompany, strCC_CODE, strChkDate, strCondition)
            Else                          '��� MU
                strFormat = "SELECT {0} FROM SC_MU_V B " & _
                             "WHERE B.SC_BU_CODE='{1}' {2} {3} " & _
                             "ORDER BY B.SORT_SEQ"
                strSQL = String.Format(strFormat, _
                                       strSelFields, .USRCompany, strChkDate, strCondition)
            End If

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetMU")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#End Region

    '*****************************************************************
    '�Է� : strInfoXML = ����⺻���� XML
    '       intRowCnt,intColCnt = ��ȸ �Ǽ�,�ʵ� ��
    '       strLOC_CODE = �����ڵ�
    '       strMU_CODE = MU �ڵ�
    '       strCC_CODE = CC �ڵ�
    '       strCODE_NAME = ��ȸ�ϰ����ϴ� �ڵ� �Ǵ� ��
    '       blnUseOnly = ���� ������� �͸� �Ǵ� ��ü
    '       strUseDate = blnUseOnly�� False�϶� EDATE>=strUseDate�� �˻�
    '       strAddFields = �ڵ�/�� �̿��� ��ȸ �߰� �ʵ�
    '       blnLikeCode  = �ڵ带 Like�� ��ȸ���� ����
    '��ȯ : ó�����
    '��� : CC CODE�� ��ȸ�ϱ����� �Լ�
    '*****************************************************************
#Region "4. CC_CODE: CC �ڵ���ȸ(Only CC)"
    Public Function GetCC(ByVal strInfoXML As String, _
                         ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                         ByVal strCODE_NAME As String) As Object
        Dim strSQL, strFormat, strSelFields, strKeys As String
        Dim strCondition As String
        Dim strChkDate As String = ""
        Dim vntData As Object

        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            '�ѱ��� ���
            strCondition = String.Format("AND CC_NAME LIKE '%{0}%'", strCODE_NAME)
            '��ȸ �ʵ� ����

            strSelFields = "CC_CODE,CC_NAME"

            strFormat = "select"
            strFormat = strFormat & " {0}"
            strFormat = strFormat & " FROM SC_CC A WHERE 1=1"
            strFormat = strFormat & " AND PC='Y' AND USE_YN = 'Y'  {1}"
            strSQL = String.Format(strFormat, _
                                       strSelFields, strCondition)

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetCC")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

    '*****************************************************************
    '�Է� : strInfoXML = ����⺻���� XML
    '       intRowCnt,intColCnt = ��ȸ �Ǽ�,�ʵ� ��
    '       strCC_CODE = CC �ڵ�
    '       blnUseOnly = ���� ������� �͸� �Ǵ� ��ü
    '       strUseDate = blnUseOnly�� False�϶� EDATE>=strUseDate�� �˻�
    '       strAddFields = �ڵ�/�� �̿��� ��ȸ �߰� �ʵ�
    '��ȯ : ó�����
    '       CC_CODE,CC_NAME,LOC_CODE,LOC_NAME,OC_CODE,OC_NAME,SC_MU_CODE,CODE_NAME,PU_CODE,PU_NAME
    '��� : CC�� �������� ����Ǿ� �ִ� LOC,MU,OC,PU�� ��ȸ�ϱ����� �Լ�
    '*****************************************************************
#Region "5. CC_CODE: CC �ڵ���ȸ(Defalut, CC�� �����Ǿ��� LOC, MU, OC, PU ��ȸ)"

    Public Function GetCCDefault(ByVal strInfoXML As String, _
                                 ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                 ByVal strCC_CODE As String, _
                                 Optional ByVal blnUseOnly As Boolean = True, _
                                 Optional ByVal strUseDate As String = "", _
                                 Optional ByVal strAddFields As String = "") As Object
        Dim strSQL, strFormat, strSelFields, strKeys As String
        Dim strCondition As String
        Dim strChkDate As String = ""
        Dim vntData As Object

        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig
            '������� ���� �˻�
            If blnUseOnly Then
                strChkDate = "AND A.USE_YN='Y'"
            Else
                strChkDate = String.Format("AND (A.USE_YN='Y' OR A.EDATE>={0})", BuildToDate(strUseDate))
            End If

            '��ȸ �ʵ� ����
            If strAddFields <> "" Then strAddFields = "," & AddAlias(strAddFields, "A")
            strSelFields = "A.CC_CODE,A.CC_NAME,A.LOC_CODE,A.LOC_NAME,A.OC_CODE,A.OC_NAME,A.SC_MU_CODE,A.MU_NAME,A.PU_CODE,A.PU_NAME " & strAddFields

            'SQL�� ����
            strFormat = "SELECT {0} " & _
                          "FROM SC_CC_V A " & _
                         "WHERE A.CC_CODE='{1}' AND A.PC='Y' AND A.SC_BU_CODE='{2}' {3} "

            strSQL = String.Format(strFormat, _
                                              strSelFields, strCC_CODE, .USRCompany, strChkDate)
            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetCCDefault")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#End Region

    '*****************************************************************
    '�Է� : strInfoXML = ����⺻���� XML
    '       intRowCnt,intColCnt = ��ȸ �Ǽ�,�ʵ� ��
    '       strCC_CODE = CC �ڵ�
    '       strMU_CODE = MU �ڵ�
    '       strCODE_NAME = ��ȸ�ϰ����ϴ� �ڵ� �Ǵ� ��
    '       blnUseOnly = ���� ������� �͸� �Ǵ� ��ü
    '       strUseDate = blnUseOnly�� False�϶� EDATE>=strUseDate�� �˻�
    '       strAddFields = �ڵ�/�� �̿��� ��ȸ �߰� �ʵ�
    '       blnLikeCode  = �ڵ带 Like�� ��ȸ���� ����
    '��ȯ : ó�����
    '��� : OC CODE�� ��ȸ�ϱ����� �Լ�
    '*****************************************************************
#Region "6. OC_CODE: OC_CODE ��ȸ"

    Public Function GetOC(ByVal strInfoXML As String, _
                          ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                          Optional ByVal strCC_CODE As String = "", _
                          Optional ByVal strMU_CODE As String = "", _
                          Optional ByVal strCODE_NAME As String = "", _
                          Optional ByVal blnUseOnly As Boolean = True, _
                          Optional ByVal strUseDate As String = "", _
                          Optional ByVal strAddFields As String = "", _
                          Optional ByVal blnLikeCode As Boolean = True) As Object
        Dim strSQL, strFormat, strSelFields, strKeys As String
        Dim strCondition As String
        Dim strChkDate As String = ""
        Dim vntData As Object

        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            If strCODE_NAME <> "" Then
                '��ȸ ������ �ڵ����� �ڵ������ �����Ͽ� ���� �ʵ� ����
                If IsNumeric(strCODE_NAME) Then '������ ���
                    If Not blnLikeCode Then
                        strCondition = String.Format("AND A.OC_CODE='{0}'", strCODE_NAME)
                    Else
                        strCondition = String.Format("AND A.OC_CODE LIKE '%{0}%'", strCODE_NAME)
                    End If
                ElseIf IsSBCS(strCODE_NAME) Then    '������ ���
                    strCondition = String.Format("AND (A.OC_CODE LIKE '%{0}%' OR A.OC_NAME LIKE '%{0}%')", strCODE_NAME)
                Else                                    '�ѱ��� ���
                    strCondition = String.Format("AND A.OC_NAME LIKE '%{0}%'", strCODE_NAME)
                End If
            End If

            '������� ���� �˻�
            If blnUseOnly Then
                strChkDate = "AND A.USE_YN='Y'"
            Else
                strChkDate = String.Format("AND (A.USE_YN='Y' OR A.EDATE>={0})", BuildToDate(strUseDate))
            End If

            '��ȸ �ʵ� ����
            If strAddFields <> "" Then strAddFields = "," & AddAlias(strAddFields, "A")
            strSelFields = "A.OC_CODE,A.OC_NAME" & strAddFields

            'SQL�� ����
            If strCC_CODE <> "" Then   'CC�� �ִ� ���
                strFormat = "SELECT {0} FROM SC_OC_V A,SC_OC_CC_V B " & _
                             "WHERE B.SC_BU_CODE='{1}' AND B.CC_CODE='{2}' " & _
                               "AND A.SC_BU_CODE=B.SC_BU_CODE AND A.OC_CODE=B.OC_CODE {3} {4} " & _
                             "ORDER BY A.OC_CODE"
                strSQL = String.Format(strFormat, _
                                       strSelFields, .USRCompany, strCC_CODE, strChkDate, strCondition)
            ElseIf strMU_CODE <> "" Then   'MU�� �ִ� ���
                strFormat = "SELECT {0} FROM SC_OC_V A,SC_OC_CC_V B,SC_MU_CC_V C " & _
                             "WHERE C.SC_BU_CODE='{1}' AND C.SC_MU_CODE='{2}' " & _
                               "AND B.CC_CODE=C.CC_CODE AND B.SC_BU_CODE=C.SC_BU_CODE " & _
                               "AND A.OC_CODE=B.OC_CODE AND A.SC_BU_CODE=B.SC_BU_CODE {3} {4} " & _
                             "ORDER BY A.OC_CODE"
                strSQL = String.Format(strFormat, _
                                       strSelFields, .USRCompany, strMU_CODE, strChkDate, strCondition)
            Else                       '�׿� ��ü OC
                strFormat = "SELECT {0} FROM SC_OC_V A " & _
                             "WHERE A.SC_BU_CODE='{1}' {2} {3} " & _
                             "ORDER BY A.OC_CODE"
                strSQL = String.Format(strFormat, _
                                       strSelFields, .USRCompany, strChkDate, strCondition)
            End If

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetOC")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#End Region

    '*****************************************************************
    '�Է� : strInfoXML = ����⺻���� XML
    '       intRowCnt,intColCnt = ��ȸ �Ǽ�,�ʵ� ��
    '       strLOC_Type = ��������(����/����/�븮��) �ڵ�
    '       strPU_CODE  = PU �ڵ�
    '       strCODE_NAME = ��ȸ�ϰ����ϴ� �ڵ� �Ǵ� ��
    '       blnUseOnly = ���� ������� �͸� �Ǵ� ��ü
    '       strUseDate = blnUseOnly�� False�϶� EDATE>=strUseDate�� �˻�
    '       strAddFields = �ڵ�/�� �̿��� ��ȸ �߰� �ʵ�
    '       blnLikeCode  = �ڵ带 Like�� ��ȸ���� ����
    '��ȯ : ó�����
    '��� : LOC CODE�� ��ȸ�ϱ����� �Լ�
    '*****************************************************************
#Region "7. LOC_CODE: LOC_CODE ��ȸ"

    Public Function GetLOC(ByVal strInfoXML As String, _
                           ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                           Optional ByVal strLOC_Type As String = "", _
                           Optional ByVal strPU_CODE As String = "", _
                           Optional ByVal strCODE_NAME As String = "", _
                           Optional ByVal blnUseOnly As Boolean = True, _
                           Optional ByVal strUseDate As String = "", _
                           Optional ByVal strAddFields As String = "", _
                           Optional ByVal blnLikeCode As Boolean = True) As Object
        Dim strSQL, strFormat, strSelFields, strKeys As String
        Dim strCondition As String
        Dim strChkDate As String = ""
        Dim vntData As Object

        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig
            If strLOC_Type = "" Then strLOC_Type = OPTIONAL_STR
            If strPU_CODE = "" Then strPU_CODE = OPTIONAL_STR

            If strCODE_NAME <> "" Then
                '��ȸ ������ �ڵ����� �ڵ������ �����Ͽ� ���� �ʵ� ����
                If IsNumeric(strCODE_NAME) Then '������ ���
                    If Not blnLikeCode Then
                        strCondition = String.Format("AND A.LOC_CODE='{0}'", strCODE_NAME)
                    Else
                        strCondition = String.Format("AND A.LOC_CODE LIKE '%{0}%'", strCODE_NAME)
                    End If
                ElseIf IsSBCS(strCODE_NAME) Then    '������ ���
                    strCondition = String.Format("AND (A.LOC_CODE LIKE '%{0}%' OR A.LOC_NAME LIKE '%{0}%')", strCODE_NAME)
                Else                                    '�ѱ��� ���
                    strCondition = String.Format("AND A.LOC_NAME LIKE '%{0}%'", strCODE_NAME)
                End If
            End If

            '������� ���� �˻�
            If blnUseOnly Then
                strChkDate = "AND A.USE_YN='Y'"
            Else
                'LOC�� �ǹ� ����
                'strChkDate = String.Format("AND (A.USE_YN='Y' OR A.EDATE>={0})", cbSCGLUtil.BuildToDate(strUseDate))
            End If

            '��ȸ �ʵ� ����
            If strAddFields <> "" Then strAddFields = "," & AddAlias(strAddFields, "A")
            strSelFields = "A.LOC_ID, A.LOC_CODE, A.LOC_NAME, A.SC_LOC_TYPE " & strAddFields

            'SQL�� ����
            strKeys = BuildFields("AND", GetFieldNameValue("A.SC_LOC_TYPE", strLOC_Type))
            If strKeys <> "" Then strKeys = "AND " & strKeys
            If strPU_CODE = OPTIONAL_STR Then 'PU�� ���� ���
                strFormat = "SELECT {0} FROM SC_ORGANIZATION_V A " & _
                             "WHERE A.SC_BU_CODE='{1}' {2} {3} {4} " & _
                             "ORDER BY A.LOC_CODE"
                strSQL = String.Format(strFormat, _
                                       strSelFields, .USRCompany, strKeys, strChkDate, strCondition)
            Else                            'PU�� �ִ� ���
                strFormat = "SELECT {0} FROM SC_ORGANIZATION_V A,SC_PLANT_PU_V B " & _
                             "WHERE A.SC_BU_CODE='{1}' AND A.SC_BU_CODE=B.SC_BU_CODE AND A.LOC_CODE=B.LOC_CODE {2} {3} AND B.PU_CODE='{4}' {5} " & _
                             "ORDER BY A.LOC_CODE"
                strSQL = String.Format(strFormat, _
                                       strSelFields, .USRCompany, strKeys, strChkDate, strPU_CODE, strCondition)
            End If

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetLOC")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

    '*****************************************************************
    '�Է� : strInfoXML = ����⺻���� XML
    '       intRowCnt,intColCnt = ��ȸ �Ǽ�,�ʵ� ��
    '       strCODE_NAME = ��ȸ�ϰ����ϴ� �ڵ� �Ǵ� ��
    '       strAddFields = �ڵ�/�� �̿��� ��ȸ �߰� �ʵ�
    '       blnLikeCode  = �ڵ带 Like�� ��ȸ���� ����
    '��ȯ : ó�����
    '��� : ��������� ��ȸ�ϱ����� �Լ�
    '*****************************************************************
#Region "8. EMP: EMP(���) ��ȸ"
    Public Function GetEMP(ByVal strInfoXML As String, _
                           ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                           Optional ByVal strCODE_NAME As String = "", _
                           Optional ByVal strCC_CODE As String = "", _
                           Optional ByVal strAddFields As String = "", _
                           Optional ByVal blnLikeCode As Boolean = True) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strCondition As String      '������
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strChkDate As String = ""   '��뿩�� �� ��볯¥
        Dim vntData As Object


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            '1.��ȸ�ʵ� ����
            If strAddFields <> "" Then strAddFields = "," & AddAlias(strAddFields, "A")
            strSelFields = "A.EMPNO, A.EMP_NAME, A.SC_JOB_GRADE_CODE, A.SC_JOB_GRADE_NAME, A.CC_CODE, A.PU_CODE, A.SC_MU_CODE " & strAddFields

            '2.������ ����
            If strCC_CODE <> "" Then
                strCondition = String.Format(" AND A.CC_CODE ='{0}' ", strCC_CODE)
            End If

            If strCODE_NAME <> "" Then
                '��ȸ ������ �ڵ����� �ڵ������ �����Ͽ� ���� �ʵ� ����
                If IsNumeric(strCODE_NAME) Then '������ ���
                    If Not blnLikeCode Then
                        strCondition &= String.Format("AND A.EMPNO={0}", strCODE_NAME)
                    Else
                        strCondition &= String.Format("AND A.EMPNO LIKE '%{0}%'", strCODE_NAME)
                    End If
                ElseIf IsSBCS(strCODE_NAME) Then    '������ ���
                    strCondition &= String.Format("AND (A.EMPNO LIKE '%{0}%' OR A.EMP_NAME LIKE '%{0}%')", strCODE_NAME)
                Else                                 '�ѱ��� ���
                    strCondition &= String.Format("AND A.EMP_NAME LIKE '%{0}%'", strCODE_NAME)
                End If
            End If


            ''3.������� ���� �˻� (EDATE ������� ����??)
            'If blnUseOnly Then
            '    strChkDate = "AND A.USE_YN='Y'"
            'Else
            '    strChkDate = String.Format("AND (A.USE_YN='Y' OR A.EDATE>={0})", BuildToDate(strUseDate))
            'End If

            '�ӽ� SQL�� ����
            strFormat = "SELECT {0} FROM SC_EMPLOYEE_MST_V A " & _
                        "WHERE A.SC_BU_CODE='{1}' {2} {3} " & _
                        "ORDER BY A.EMPNO"

            'SQL�� ����
            strSQL = String.Format(strFormat, strSelFields, .USRCompany, strCondition, strChkDate)

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetEMP")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region
    '*****************************************************************
    '�Է� : strInfoXML = ����⺻���� XML
    '       intRowCnt,intColCnt = ��ȸ �Ǽ�,�ʵ� ��
    '       strCODE_NAME = ��ȸ�ϰ����ϴ� �ڵ� �Ǵ� ��
    '       strAddFields = �ڵ�/�� �̿��� ��ȸ �߰� �ʵ�
    '       blnLikeCode  = �ڵ带 Like�� ��ȸ���� ����
    '��ȯ : ó�����
    '��� : ��������� ��ȸ�ϱ����� �Լ�
    '*****************************************************************

#Region "8. USER: �α��λ���� USER ��ȸ"
    Public Function GetUser(ByVal strInfoXML As String, _
                           ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                           Optional ByVal strCODE_NAME As String = "", _
                           Optional ByVal strCC_CODE As String = "", _
                           Optional ByVal strAddFields As String = "", _
                           Optional ByVal blnLikeCode As Boolean = True) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strCondition As String      '������
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strChkDate As String = ""   '��뿩�� �� ��볯¥
        Dim vntData As Object


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            '1.��ȸ�ʵ� ����
            If strAddFields <> "" Then strAddFields = "," & AddAlias(strAddFields, "A")
            strSelFields = "A.EMPNO, A.EMP_NAME, A.SC_JOB_GRADE_CODE, A.SC_JOB_GRADE_NAME, A.CC_CODE, A.PU_CODE, A.SC_MU_CODE " & strAddFields

            '2.������ ����
            If strCC_CODE <> "" Then
                strCondition = String.Format(" AND A.CC_CODE ='{0}' ", strCC_CODE)
            End If

            If strCODE_NAME <> "" Then
                '��ȸ ������ �ڵ����� �ڵ������ �����Ͽ� ���� �ʵ� ����
                If IsNumeric(strCODE_NAME) Then '������ ���
                    If Not blnLikeCode Then
                        strCondition &= String.Format("AND A.EMPNO={0}", strCODE_NAME)
                    Else
                        strCondition &= String.Format("AND A.EMPNO LIKE '%{0}%'", strCODE_NAME)
                    End If
                ElseIf IsSBCS(strCODE_NAME) Then    '������ ���
                    strCondition &= String.Format("AND (A.EMPNO LIKE '%{0}%' OR A.EMP_NAME LIKE '%{0}%')", strCODE_NAME)
                Else                                 '�ѱ��� ���
                    strCondition &= String.Format("AND A.EMP_NAME LIKE '%{0}%'", strCODE_NAME)
                End If
            End If


            ''3.������� ���� �˻� (EDATE ������� ����??)
            'If blnUseOnly Then
            '    strChkDate = "AND A.USE_YN='Y'"
            'Else
            '    strChkDate = String.Format("AND (A.USE_YN='Y' OR A.EDATE>={0})", BuildToDate(strUseDate))
            'End If

            '�ӽ� SQL�� ����
            strFormat = "SELECT {0} FROM SC_USER_INFO_V A " & _
                        "WHERE A.SC_BU_CODE='{1}' {2} {3} " & _
                        "ORDER BY A.EMPNO"

            'SQL�� ����
            strSQL = String.Format(strFormat, strSelFields, .USRCompany, strCondition, strChkDate)

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetEMP")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

    '*****************************************************************
    '�Է� : strInfoXML = ����⺻���� XML
    '       intRowCnt,intColCnt = ��ȸ �Ǽ�,�ʵ� ��
    '       strCODE_NAME = ��ȸ�ϰ����ϴ� �ּ�
    '       strAddFields = �ּ� �̿��� ��ȸ �߰� �ʵ�
    '       blnLikeCode  = �ڵ带 Like�� ��ȸ���� ����
    '��ȯ : ó�����
    '��� : �����ȣ�� ��ȸ�ϱ����� �Լ�
    '*****************************************************************
#Region "9. POST(�����ȣ) ��ȸ"

    Public Function GetPOST(ByVal strInfoXML As String, _
                           ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                           ByVal strSEARCH_ADDR As String, _
                           Optional ByVal strAddFields As String = "", _
                           Optional ByVal blnLikeCode As Boolean = True) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strKeys As String           '������
        Dim strSelFields As String      '��ȸ�ʵ�

        Dim vntData As Object

        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            '��ȸ�ʵ� ����
            If strAddFields <> "" Then strAddFields = "," & AddAlias(strAddFields, "A")
            strSelFields = "A.POST_CODE, A.MAIN_ADDR, A.SEPOINT" & strAddFields

            '������ ����
            strKeys = String.Format("A.SEARCH_ADDR LIKE '%{0}%'", strSEARCH_ADDR)

            '�ӽ� SQL�� ����
            strFormat = "SELECT {0} FROM SC_POST A WHERE {1} ORDER BY A.POST_CODE"

            'SQL�� ����
            strSQL = String.Format(strFormat, strSelFields, strKeys)

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetPOST")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#End Region

    '*****************************************************************
    '�Է� : strInfoXML = ����⺻���� XML
    '       intRowCnt,intColCnt = ��ȸ �Ǽ�,�ʵ� ��
    '       strCC_CODE = CURR_TYPE_CODE OR CURR_TYPE_NAME
    '       blnUseOnly = ���� ������� �͸� �Ǵ� ��ü
    '       strUseDate = blnUseOnly�� False�϶� EDATE>=strUseDate�� �˻�
    '       strAddFields = �ڵ�/�� �̿��� ��ȸ �߰� �ʵ�
    '��ȯ : ó�����
    '��� : CurrencyType (A.CURR_TYPE_CODE, A.CURR_TYPE_NAME) �� ��ȸ
    '*****************************************************************
#Region "10. CurrencyType ��ȸ"

    Public Function GetCurrencyType(ByVal strInfoXML As String, _
                                    ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                    Optional ByVal strCODE_NAME As String = "", _
                                    Optional ByVal strAddFields As String = "", _
                                    Optional ByVal blnUseOnly As Boolean = True, _
                                    Optional ByVal strUseDate As String = "", _
                                    Optional ByVal blnLikeCode As Boolean = True) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strCondition As String      '������
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strChkDate As String        '��뿩�� �� ��볯¥
        Dim vntData As Object


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            '1.��ȸ�ʵ� ����
            If strAddFields <> "" Then strAddFields = "," & AddAlias(strAddFields, "A")
            strSelFields = "A.CURR_TYPE_CODE, A.CURR_TYPE_NAME " & strAddFields

            '2.������ ����
            If strCODE_NAME <> "" Then
                '��ȸ ������ �ڵ����� �ڵ������ �����Ͽ� ���� �ʵ� ����
                If IsNumeric(strCODE_NAME) Then '������ ���
                    If Not blnLikeCode Then
                        strCondition = String.Format("AND A.CURR_TYPE_CODE='{0}'", strCODE_NAME)
                    Else
                        strCondition = String.Format("AND A.CURR_TYPE_CODE LIKE '%{0}%'", strCODE_NAME)
                    End If
                ElseIf IsSBCS(strCODE_NAME) Then    '������ ���
                    strCondition = String.Format("AND (A.CURR_TYPE_CODE LIKE '%{0}%' OR A.CURR_TYPE_NAME LIKE '%{0}%')", strCODE_NAME)
                Else                                    '�ѱ��� ���
                    strCondition = String.Format("AND A.CURR_TYPE_NAME LIKE '%{0}%'", strCODE_NAME)
                End If
            End If

            '3.������� ���� �˻�
            If blnUseOnly Then
                strChkDate = "AND A.USE_YN='Y'"
            Else
                strChkDate = String.Format("AND (A.USE_YN='Y' OR A.EDATE>={0})", BuildToDate(strUseDate))
            End If

            '�ӽ� SQL�� ����
            strFormat = "SELECT {0} FROM SC_CURRENCY_TYPE_V A " & _
                        "WHERE A.SC_BU_CODE='{1}' {2} {3} " & _
                        "ORDER BY A.CURR_TYPE_CODE"

            'SQL�� ����
            strSQL = String.Format(strFormat, strSelFields, .USRCompany, strCondition, strChkDate)

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetCurrencyType")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#End Region

    '*****************************************************************
    '�Է� : strInfoXML = ����⺻���� XML
    '       intRowCnt,intColCnt = ��ȸ �Ǽ�,�ʵ� ��
    '       strCC_CODE = CURR_CODE OR CURRNAME
    '       blnUseOnly = ���� ������� �͸� �Ǵ� ��ü
    '       strUseDate = blnUseOnly�� False�϶� EDATE>=strUseDate�� �˻�
    '       strAddFields = �ڵ�/�� �̿��� ��ȸ �߰� �ʵ�
    '       blnLikeCode = like �� ����� ���ΰ�? (Default True)
    '��ȯ : ó�����
    '��� : Currency (A.CURR_CODE, A.CURRNAME) �� ��ȸ
    '*****************************************************************
#Region "11. Currency ��ȸ"
    Public Function GetCurrency(ByVal strInfoXML As String, _
                                ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                Optional ByVal strCODE_NAME As String = "", _
                                Optional ByVal strAddFields As String = "", _
                                Optional ByVal blnUseOnly As Boolean = True, _
                                Optional ByVal strUseDate As String = "", _
                                Optional ByVal blnLikeCode As Boolean = True) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strCondition As String      '������
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strChkDate As String        '��뿩�� �� ��볯¥
        Dim vntData As Object


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            '1.��ȸ�ʵ� ����
            If strAddFields <> "" Then strAddFields = "," & AddAlias(strAddFields, "A")
            strSelFields = "A.CURR_CODE, A.CURRNAME " & strAddFields

            '2.������ ����
            If strCODE_NAME <> "" Then
                '��ȸ ������ �ڵ����� �ڵ������ �����Ͽ� ���� �ʵ� ����
                If IsNumeric(strCODE_NAME) Then '������ ���
                    If Not blnLikeCode Then
                        strCondition = String.Format("AND A.CURR_CODE='{0}'", strCODE_NAME)
                    Else
                        strCondition = String.Format("AND A.CURR_CODE LIKE '%{0}%'", strCODE_NAME)
                    End If
                ElseIf IsSBCS(strCODE_NAME) Then    '������ ���
                    strCondition = String.Format("AND (A.CURR_CODE LIKE '%{0}%' OR A.CURRNAME LIKE '%{0}%')", strCODE_NAME)
                Else                                 '�ѱ��� ���
                    strCondition = String.Format("AND A.CURRNAME LIKE '%{0}%'", strCODE_NAME)
                End If
            End If

            '3.������� ���� �˻�
            If blnUseOnly Then
                strChkDate = "AND A.USE_YN='Y'"
            Else
                strChkDate = String.Format("AND (A.USE_YN='Y' OR A.EDATE>={0})", BuildToDate(strUseDate))
            End If

            '�ӽ� SQL�� ����
            strFormat = "SELECT {0} FROM SC_CURRENCY_V A " & _
                        "WHERE A.SC_BU_CODE='{1}' {2} {3} " & _
                        "ORDER BY A.CURR_CODE"

            'SQL�� ����
            strSQL = String.Format(strFormat, strSelFields, .USRCompany, strCondition, strChkDate)

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetCurrency")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#End Region

    '*****************************************************************
    '�Է� : strInfoXML = ����⺻���� XML
    '       intRowCnt,intColCnt = ��ȸ �Ǽ�,�ʵ� ��
    '       strCODE = CURR_CODE OR CURRNAME
    '       blnUseOnly = ���� ������� �͸� �Ǵ� ��ü
    '       strUseDate = blnUseOnly�� False�϶� EDATE>=strUseDate�� �˻�
    '       strAddFields = �ڵ�/�� �̿��� ��ȸ �߰� �ʵ�
    '       blnLikeCode = like �� ����� ���ΰ�? (Default True)
    '��ȯ : ó�����
    '��� : CURRENCY_DAILY (A.FROM_CURR_CODE, A.TO_CURR_CODE, A.CURR_RATE) �� ��ȸ 
    '*****************************************************************
#Region "12. CurrencyDaily ��ȸ"

    Public Function GetCurrencyDaily(ByVal strInfoXML As String, _
                                    ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                    Optional ByVal strSTD_DD As String = "", _
                                    Optional ByVal strCURR_TYPE_CODE As String = "", _
                                    Optional ByVal strFROM_CURR_CODE As String = "", _
                                    Optional ByVal strTO_CURR_CODE As String = "", _
                                    Optional ByVal strAddFields As String = "", _
                                    Optional ByVal blnUseOnly As Boolean = True, _
                                    Optional ByVal strUseDate As String = "", _
                                    Optional ByVal blnLikeCode As Boolean = True) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strCondition As String      '������
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strChkDate As String = ""      '��뿩�� �� ��볯¥
        Dim vntData As Object


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            '1.��ȸ�ʵ� ����
            If strAddFields <> "" Then strAddFields = "," & AddAlias(strAddFields, "A")
            strSelFields = "A.STD_DD, A.FROM_CURR_CODE, A.TO_CURR_CODE, A.CURR_RATE, A.CURR_TYPE_CODE " & strAddFields

            '2.������ ����
            '''''''''''''' 1) STD_DD, strCURR_TYPE_CODE, FROM_CURR_CODE, TO_CURR_CODE ����
            If strSTD_DD = "" Then strSTD_DD = OPTIONAL_STR Else strSTD_DD = "%" & strSTD_DD & "%"
            If strCURR_TYPE_CODE = "" Then strCURR_TYPE_CODE = OPTIONAL_STR Else strCURR_TYPE_CODE = "%" & strCURR_TYPE_CODE & "%"
            If strFROM_CURR_CODE = "" Then strFROM_CURR_CODE = OPTIONAL_STR Else strFROM_CURR_CODE = "%" & strFROM_CURR_CODE & "%"
            If strTO_CURR_CODE = "" Then strTO_CURR_CODE = OPTIONAL_STR Else strTO_CURR_CODE = "%" & strTO_CURR_CODE & "%"

            strCondition = BuildFields("AND", _
                    GetFieldNameValue("A.STD_DD", strSTD_DD, "like"), _
                    GetFieldNameValue("A.CURR_TYPE_CODE", strCURR_TYPE_CODE, "like"), _
                    GetFieldNameValue("A.FROM_CURR_CODE", strFROM_CURR_CODE, "like"), _
                    GetFieldNameValue("A.TO_CURR_CODE", strTO_CURR_CODE, "like"))

            If strCondition <> "" Then strCondition = "AND " & strCondition

            '3.������� ���� �˻�(����)
            'If blnUseOnly Then
            '    strChkDate = "AND USE_YN='Y'"
            'Else
            '    strChkDate = String.Format("AND (A.USE_YN='Y' OR A.EDATE>={0})", BuildToDate(strUseDate))
            'End If

            '�ӽ� SQL�� ����
            strFormat = "SELECT {0} FROM SC_CURRENCY_DAILY_V A " & _
                        "WHERE A.SC_BU_CODE='{1}' {2} {3} " & _
                        "ORDER BY A.TO_CURR_CODE"

            'SQL�� ����
            strSQL = String.Format(strFormat, strSelFields, .USRCompany, strCondition, strChkDate)

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetCurrencyDaily")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#End Region

    '*****************************************************************
    '�Է� : strInfoXML = ����⺻���� XML
    '       intRowCnt,intColCnt = ��ȸ �Ǽ�,�ʵ� ��
    '       strCODE = A.MOD_CATEGORY_CODE, A.MOD_CATEGORY_NAME
    '       blnUseOnly = ���� ������� �͸� �Ǵ� ��ü
    '       strUseDate = blnUseOnly�� False�϶� EDATE>=strUseDate�� �˻�
    '       strAddFields = �ڵ�/�� �̿��� ��ȸ �߰� �ʵ�
    '       blnLikeCode = like �� ����� ���ΰ�? (Default True)
    '��ȯ : ó�����
    '��� : MOD_CATEGORY �������� (A.MOD_CATEGORY_CODE, A.MOD_CATEGORY_NAME) �� ��ȸ
    '*****************************************************************
#Region "13. ModCategory ��ȸ"

    Public Function GetModCategory(ByVal strInfoXML As String, _
                                    ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                    Optional ByVal strCODE_NAME As String = "", _
                                    Optional ByVal strAddFields As String = "", _
                                    Optional ByVal blnUseOnly As Boolean = True, _
                                    Optional ByVal strUseDate As String = "", _
                                    Optional ByVal blnLikeCode As Boolean = True) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strCondition As String      '������
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strChkDate As String        '��뿩�� �� ��볯¥
        Dim vntData As Object


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            '1.��ȸ�ʵ� ����
            If strAddFields <> "" Then strAddFields = "," & AddAlias(strAddFields, "A")
            strSelFields = "A.MOD_CATEGORY_CODE, A.MOD_CATEGORY_NAME " & strAddFields

            '2.������ ����
            If strCODE_NAME <> "" Then
                '��ȸ ������ �ڵ����� �ڵ������ �����Ͽ� ���� �ʵ� ����
                If IsNumeric(strCODE_NAME) Then '������ ���
                    If Not blnLikeCode Then
                        strCondition = String.Format("AND A.MOD_CATEGORY_CODE='{0}'", strCODE_NAME)
                    Else
                        strCondition = String.Format("AND A.MOD_CATEGORY_CODE LIKE '%{0}%'", strCODE_NAME)
                    End If
                ElseIf IsSBCS(strCODE_NAME) Then    '������ ���
                    strCondition = String.Format("AND (A.MOD_CATEGORY_CODE LIKE '%{0}%' OR A.MOD_CATEGORY_NAME LIKE '%{0}%')", strCODE_NAME)
                Else                                 '�ѱ��� ���
                    strCondition = String.Format("AND A.MOD_CATEGORY_NAME LIKE '%{0}%'", strCODE_NAME)
                End If
            End If

            '3.������� ���� �˻�
            If blnUseOnly Then
                strChkDate = "AND A.USE_YN='Y'"
            Else
                strChkDate = String.Format("AND (A.USE_YN='Y' OR A.EDATE>={0})", BuildToDate(strUseDate))
            End If

            '�ӽ� SQL�� ����
            strFormat = "SELECT {0} FROM SC_MOD_CATEGORY_V A " & _
                        "WHERE A.SC_BU_CODE='{1}' {2} {3} " & _
                        "ORDER BY A.MOD_CATEGORY_CODE"

            'SQL�� ����
            strSQL = String.Format(strFormat, strSelFields, .USRCompany, strCondition, strChkDate)

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetModCategory")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#End Region

    '*****************************************************************
    '�Է� : strInfoXML = ����⺻���� XML
    '       intRowCnt,intColCnt = ��ȸ �Ǽ�,�ʵ� ��
    '       strCODE = A.VOU_CODE, A.VOU_NAME
    '       blnUseOnly = ���� ������� �͸� �Ǵ� ��ü
    '       strUseDate = blnUseOnly�� False�϶� EDATE>=strUseDate�� �˻�
    '       strAddFields = �ڵ�/�� �̿��� ��ȸ �߰� �ʵ�
    '       blnLikeCode = like �� ����� ���ΰ�? (Default True)
    '��ȯ : ó�����
    '��� : VOUCHER ��ǥ��ȣ���� (A.VOU_CODE, A.SRL, A.VOU_NAME) �� ��ȸ
    '*****************************************************************
#Region "14. Voucher ��ǥ��ȣ���� ��ȸ"

    Public Function GetVoucher(ByVal strInfoXML As String, _
                                    ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                    Optional ByVal strCODE_NAME As String = "", _
                                    Optional ByVal strAddFields As String = "", _
                                    Optional ByVal blnUseOnly As Boolean = True, _
                                    Optional ByVal strUseDate As String = "", _
                                    Optional ByVal blnLikeCode As Boolean = True) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strCondition As String      '������
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strChkDate As String        '��뿩�� �� ��볯¥
        Dim vntData As Object


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            '1.��ȸ�ʵ� ����
            If strAddFields <> "" Then strAddFields = "," & AddAlias(strAddFields, "A")
            strSelFields = "A.VOU_CODE, A.VOU_NAME " & strAddFields

            '2.������ ����
            If strCODE_NAME <> "" Then
                '��ȸ ������ �ڵ����� �ڵ������ �����Ͽ� ���� �ʵ� ����
                If IsNumeric(strCODE_NAME) Then '������ ���
                    If Not blnLikeCode Then
                        strCondition = String.Format("AND A.VOU_CODE='{0}'", strCODE_NAME)
                    Else
                        strCondition = String.Format("AND A.VOU_CODE LIKE '{0}%'", strCODE_NAME)
                    End If
                ElseIf IsSBCS(strCODE_NAME) Then    '������ ���
                    strCondition = String.Format("AND (A.VOU_CODE LIKE '{0}%' OR A.VOU_NAME LIKE '{0}%')", strCODE_NAME)
                Else                                 '�ѱ��� ���
                    strCondition = String.Format("AND A.VOU_NAME LIKE '{0}%'", strCODE_NAME)
                End If
            End If

            '3.������� ���� �˻�
            If blnUseOnly Then
                strChkDate = "AND A.USE_YN='Y'"
            Else
                strChkDate = String.Format("AND (A.USE_YN='Y' OR A.EDATE>={0})", BuildToDate(strUseDate))
            End If

            '�ӽ� SQL�� ����
            strFormat = "SELECT {0} FROM SC_VOUCHER_V A " & _
                        "WHERE A.SC_BU_CODE='{1}' {2} {3} " & _
                        "ORDER BY A.VOU_CODE"

            'SQL�� ����
            strSQL = String.Format(strFormat, strSelFields, .USRCompany, strCondition, strChkDate)

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetVoucher")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#End Region

#Region "14-2. Voucher ��ǥ��ȣ���� ��ȸ"

    Public Function GetVoucher2(ByVal strInfoXML As String, _
                                    ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                    Optional ByVal strCODE_NAME As String = "", _
                                    Optional ByVal strModule As String = "", _
                                    Optional ByVal strAddFields As String = "", _
                                    Optional ByVal blnUseOnly As Boolean = True, _
                                    Optional ByVal strUseDate As String = "", _
                                    Optional ByVal blnLikeCode As Boolean = True) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strCondition As String      '������
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strChkDate As String        '��뿩�� �� ��볯¥
        Dim vntData As Object


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            '1.��ȸ�ʵ� ����
            If strAddFields <> "" Then strAddFields = "," & AddAlias(strAddFields, "A")
            strSelFields = "A.VOUCHER_CODE, B.VOU_NAME " & strAddFields

            '2.������ ����
            If strModule <> "" Then strCondition = String.Format("AND A.MOD_CATEGORY_CODE='{0}' ", strModule)

            If strCODE_NAME <> "" Then
                '��ȸ ������ �ڵ����� �ڵ������ �����Ͽ� ���� �ʵ� ����
                If IsNumeric(strCODE_NAME) Then '������ ���
                    If Not blnLikeCode Then
                        strCondition &= String.Format("AND A.VOUCHER_CODE='{0}'", strCODE_NAME)
                    Else
                        strCondition &= String.Format("AND A.VOUCHER_CODE LIKE '{0}%'", strCODE_NAME)
                    End If
                ElseIf IsSBCS(strCODE_NAME) Then    '������ ���
                    strCondition &= String.Format("AND (A.VOUCHER_CODE LIKE '{0}%' OR B.VOU_NAME LIKE '{0}%')", strCODE_NAME)
                Else                                 '�ѱ��� ���
                    strCondition &= String.Format("AND B.VOU_NAME LIKE '{0}%'", strCODE_NAME)
                End If
            End If

            '�ӽ� SQL�� ����
            strFormat = "SELECT {0} FROM SC_VOU_ASSIGN A, SC_VOUCHER B " & _
                        "WHERE A.VOUCHER_CODE = B.VOU_CODE AND B.SC_BU_CODE='{1}' {2} " & _
                        "ORDER BY A.VOUCHER_CODE"

            'SQL�� ����
            strSQL = String.Format(strFormat, strSelFields, .USRCompany, strCondition)

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetVoucher2")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#End Region
    '*****************************************************************
    '�Է� : strInfoXML = ����⺻���� XML
    '       intRowCnt,intColCnt = ��ȸ �Ǽ�,�ʵ� ��
    '       strCODE = A.BANK_CODE, A.BANK_BRANCH_NAME
    '       blnUseOnly = ���� ������� �͸� �Ǵ� ��ü
    '       strUseDate = blnUseOnly�� False�϶� EDATE>=strUseDate�� �˻�
    '       strAddFields = �ڵ�/�� �̿��� ��ȸ �߰� �ʵ�
    '       blnLikeCode = like �� ����� ���ΰ�? (Default True)
    '��ȯ : ó�����
    '��� : BANK_BRANCH_REG �������� (A.BANK_BRANCH_ID, A.BANK_BRANCH_NAME) �� ��ȸ
    '*****************************************************************
#Region "15. BANK_BRANCH_REG ��������"

    Public Function GetBankBranchReg(ByVal strInfoXML As String, _
                                    ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                    Optional ByVal strCODE_NAME As String = "", _
                                    Optional ByVal strAddFields As String = "", _
                                    Optional ByVal blnUseOnly As Boolean = True, _
                                    Optional ByVal strUseDate As String = "", _
                                    Optional ByVal blnLikeCode As Boolean = True) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strCondition As String      '������
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strChkDate As String = ""   '��뿩�� �� ��볯¥
        Dim vntData As Object


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            '1.��ȸ�ʵ� ����
            If strAddFields <> "" Then strAddFields = "," & AddAlias(strAddFields, "A")
            strSelFields = "A.BANK_CODE, A.BANK_NAME, A.BANK_BRANCH_ID, A.BANK_BRANCH_NAME " & strAddFields

            '2.������ ����
            If strCODE_NAME <> "" Then
                '��ȸ ������ �ڵ����� �ڵ������ �����Ͽ� ���� �ʵ� ����
                If IsNumeric(strCODE_NAME) Then '������ ���
                    If Not blnLikeCode Then
                        strCondition = String.Format("AND A.BANK_CODE='{0}'", strCODE_NAME)
                    Else
                        strCondition = String.Format("AND A.BANK_CODE LIKE '{0}%'", strCODE_NAME)
                    End If
                ElseIf IsSBCS(strCODE_NAME) Then    '������ ���
                    strCondition = String.Format("AND (A.BANK_CODE LIKE '{0}%' OR A.BANK_BRANCH_NAME LIKE '{0}%')", strCODE_NAME)
                Else                                 '�ѱ��� ���
                    strCondition = String.Format("AND A.BANK_BRANCH_NAME LIKE '{0}%'", strCODE_NAME)
                End If
            End If

            '3.������� ���� �˻� (������� ����)
            'If blnUseOnly Then
            '    strChkDate = "AND A.USE_YN='Y'"
            'Else
            '    strChkDate = String.Format("AND (A.USE_YN='Y' OR A.EDATE>={0})", BuildToDate(strUseDate))
            'End If

            '�ӽ� SQL�� ����
            strFormat = "SELECT {0} FROM SC_BANK_V A " & _
                        "WHERE A.SC_BU_CODE='{1}' {2} {3} " & _
                        "GROUP BY A.BANK_CODE, A.BANK_NAME, A.BANK_BRANCH_ID,A.BANK_BRANCH_NAME ORDER BY A.BANK_NAME"

            'SQL�� ����
            strSQL = String.Format(strFormat, strSelFields, .USRCompany, strCondition, strChkDate)

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetBankBranchReg")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

    '*****************************************************************
    '�Է� : strInfoXML = ����⺻���� XML
    '       intRowCnt,intColCnt = ��ȸ �Ǽ�,�ʵ� ��
    '       strCODE_NAME = A.BANK_ACCOUNT_NAME, A.BANK_ACCOUNT_NUM
    '       blnUseOnly = ���� ������� �͸� �Ǵ� ��ü
    '       strUseDate = blnUseOnly�� False�϶� EDATE>=strUseDate�� �˻�
    '       strAddFields = �ڵ�/�� �̿��� ��ȸ �߰� �ʵ�
    '       blnLikeCode = like �� ����� ���ΰ�? (Default True)
    '��ȯ : ó�����
    '��� : BANK_ACCOUNT ������� (A.BANK_ACCOUNT_ID, A.BANK_ACCOUNT_NAME, A.BANK_ACCOUNT_NUM) �� ��ȸ
    '*****************************************************************
#Region "16. BANK_ACCOUNT ������� ��ȸ"

    Public Function GetBankAccount(ByVal strInfoXML As String, _
                                   ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                   Optional ByVal strCODE_NAME As String = "", _
                                   Optional ByVal strBANK_BRANCH_ID As String = "", _
                                   Optional ByVal strBANK_ACCOUNT_TYPE As String = "", _
                                   Optional ByVal strAddFields As String = "", _
                                   Optional ByVal blnUseOnly As Boolean = True, _
                                   Optional ByVal strUseDate As String = "", _
                                   Optional ByVal blnLikeCode As Boolean = True) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strCondition As String      '������
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strChkDate As String = ""   '��뿩�� �� ��볯¥

        Dim vntData As Object
        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            '1.��ȸ�ʵ� ����
            If strAddFields <> "" Then strAddFields = "," & AddAlias(strAddFields, "A")
            strSelFields = "A.BANK_ACCOUNT_ID, A.BANK_ACCOUNT_NAME, A.BANK_ACCOUNT_NUM " & strAddFields

            '2.������ ����
            If strBANK_BRANCH_ID <> "" Then
                strCondition = String.Format(" AND A.BANK_BRANCH_ID = {0} ", strBANK_BRANCH_ID)
            End If

            If strBANK_ACCOUNT_TYPE <> "" Then
                strCondition &= String.Format(" AND A.BANK_ACCOUNT_TYPE = '{0}' ", strBANK_ACCOUNT_TYPE)
            End If

            If strCODE_NAME <> "" Then
                '��ȸ ������ �ڵ����� �ڵ������ �����Ͽ� ���� �ʵ� ����
                If IsNumeric(strCODE_NAME) Then '������ ���
                    If Not blnLikeCode Then
                        strCondition &= String.Format("AND A.BANK_ACCOUNT_NUM='{0}'", strCODE_NAME)
                    Else
                        strCondition &= String.Format("AND A.BANK_ACCOUNT_NUM LIKE '{0}%'", strCODE_NAME)
                    End If
                ElseIf IsSBCS(strCODE_NAME) Then    '������ ���
                    strCondition &= String.Format("AND (A.BANK_ACCOUNT_NUM LIKE '{0}%' OR A.BANK_ACCOUNT_NAME LIKE '{0}%')", strCODE_NAME)
                Else                                 '�ѱ��� ���
                    strCondition &= String.Format("AND A.BANK_ACCOUNT_NAME LIKE '{0}%'", strCODE_NAME)
                End If
            End If


            '3.������� ���� �˻� (������� ����)
            'If blnUseOnly Then
            '    strChkDate = "AND A.USE_YN='Y'"
            'Else
            '    strChkDate = String.Format("AND (A.USE_YN='Y' OR A.EDATE>={0})", BuildToDate(strUseDate))
            'End If


            '�ӽ� SQL�� ����
            strFormat = "SELECT {0} FROM SC_BANK_V A " & _
                        "WHERE A.SC_BU_CODE='{1}' AND A.BANK_ACCOUNT_ID IS NOT NULL  {2} {3} " & _
                        "ORDER BY A.BANK_ACCOUNT_ID"

            'SQL�� ����
            strSQL = String.Format(strFormat, strSelFields, .USRCompany, strCondition, strChkDate)

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetBankAccount")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#End Region
    '''''''''''''''''''''''''''''''''''''
    ''''''���¹�ȣ�� ���� ������ �����´�.
    '''''''''''''''''''''''''''''''''''''
#Region "GetBankInfoByAccID"

    Public Function GetBankInfoByAccID(ByVal strInfoXML As String, _
                                       ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                       Optional ByVal strBANK_ACCOUNT_ID As String = "") As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strCondition As String      '������
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strChkDate As String = ""   '��뿩�� �� ��볯¥
        Dim vntData As Object


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            '1.��ȸ�ʵ� ����
            strSelFields = "A.BANK_CODE, A.BANK_NAME, A.BANK_BRANCH_ID, A.BANK_BRANCH_NAME, A.BANK_ACCOUNT_NAME, A.BANK_ACCOUNT_NUM, A.LOC_CODE, A.LOC_NAME, A.CURR_CODE, A.BANK_ACCOUNT_CLASS, A.CASH_NA_CODE "
            '2.������ ����
            If strBANK_ACCOUNT_ID <> "" Then
                strCondition = String.Format(" AND A.BANK_ACCOUNT_ID = {0} ", strBANK_ACCOUNT_ID)
            End If

            '�ӽ� SQL�� ����
            strFormat = "SELECT {0} FROM SC_BANK_V A " & _
                        "WHERE A.SC_BU_CODE='{1}' {2} " & _
                        "ORDER BY A.BANK_BRANCH_ID"

            'SQL�� ����
            strSQL = String.Format(strFormat, strSelFields, .USRCompany, strCondition)

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetBankInfoByAccID")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#End Region
    '*****************************************************************
    '�Է� : strInfoXML = ����⺻���� XML
    '       intRowCnt,intColCnt = ��ȸ �Ǽ�,�ʵ� ��
    '       strCODE_NAME = A.CUST_NAME, A.REG_NUM
    '       strCUST_TYPE 
    '       blnUseOnly = ���� ������� �͸� �Ǵ� ��ü
    '       strUseDate = blnUseOnly�� False�϶� EDATE>=strUseDate�� �˻�
    '       strAddFields = �ڵ�/�� �̿��� ��ȸ �߰� �ʵ�
    '       blnLikeCode = like �� ����� ���ΰ�? (Default True)
    '��ȯ : ó�����
    '��� : CUSTOMER �ŷ�ó (A.CUST_ID, A.CUST_NAME, A.REG_NUM) �� ��ȸ
    '*****************************************************************
#Region "17. CUSTOMER �ŷ�ó(Header & Detail)"

    Public Function GetCustomer(ByVal strInfoXML As String, _
                                    ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                    Optional ByVal strCODE_NAME As String = "", _
                                    Optional ByVal strCUST_TYPE As String = "", _
                                    Optional ByVal strAddFields As String = "", _
                                    Optional ByVal blnUseOnly As Boolean = True, _
                                    Optional ByVal strUseDate As String = "", _
                                    Optional ByVal blnLikeCode As Boolean = True) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strCondition As String      '������
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strChkDate As String = ""   '��뿩�� �� ��볯¥
        Dim vntData As Object
        Dim connStr As String

        SetConfig(strInfoXML)           '�⺻���� ����
        With mobjSCGLConfig

            '1.��ȸ�ʵ� ����
            If strAddFields <> "" Then strAddFields = "," & AddAlias(strAddFields, "A")
            strSelFields = "A.CUST_ID, A.CUST_NAME, A.REG_NUM, A.CEO, SC_EMP_NAME_FUN(A.SALES_BY,'" & mobjSCGLConfig.USRCompany & "') AS SALESMAN, A.CUST_DESC, A.OLD_CUST_CODE, A.BIZ_DATE, DECODE(A.HANA_CUST_YN,'Y','*',A.HANA_CUST_YN) As HANA_CUST_YN  " & strAddFields

            '2.������ ����
            If strCUST_TYPE <> "" Then
                strCondition = String.Format(" AND A.CUST_TYPE IN ('A','{0}') ", strCUST_TYPE)
            End If

            If strCODE_NAME <> "" Then
                '��ȸ ������ �ڵ����� �ڵ������ �����Ͽ� ���� �ʵ� ����
                If IsNumeric(strCODE_NAME) Then     '������ ���
                    If Not blnLikeCode Then
                        strCondition &= String.Format("AND A.REG_NUM='{0}'", strCODE_NAME)
                    Else
                        strCondition &= String.Format("AND A.REG_NUM LIKE '{0}%'", strCODE_NAME)
                    End If
                ElseIf IsSBCS(strCODE_NAME) Then    '������ ���
                    strCondition &= String.Format("AND (A.REG_NUM LIKE '{0}%' OR UPPER(A.CUST_NAME) LIKE UPPER('{0}')||'%') ", strCODE_NAME)
                Else                                '�ѱ��� ���
                    strCondition &= String.Format("AND A.CUST_NAME LIKE '{0}%'", strCODE_NAME)
                End If
            End If

            '3.������� ���� �˻� (EDATE ������� ����??)
            If blnUseOnly Then
                strChkDate = "AND A.USE_YN='Y'"
            Else
                '    strChkDate = String.Format("AND (A.USE_YN='Y' OR A.EDATE>={0})", BuildToDate(strUseDate))
            End If

            '�ӽ� SQL�� ����
            strFormat = "SELECT {0} FROM SC_CUSTOMER_V A " & _
                        "WHERE A.SC_BU_CODE='{1}' {2} {3} " & _
                        "ORDER BY A.CUST_NAME"

            'SQL�� ����
            strSQL = String.Format(strFormat, strSelFields, .USRCompany, strCondition, strChkDate)

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetCustomer")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "17-1. CUSTOMER �ŷ�ó-Header"

    Public Function GetCustomerHdr(ByVal strInfoXML As String, _
                                    ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                    Optional ByVal strCODE_NAME As String = "", _
                                    Optional ByVal strCUST_TYPE As String = "", _
                                    Optional ByVal strAddFields As String = "", _
                                    Optional ByVal blnUseOnly As Boolean = True, _
                                    Optional ByVal strUseDate As String = "", _
                                    Optional ByVal blnLikeCode As Boolean = True) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strCondition As String      '������
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strChkDate As String = ""   '��뿩�� �� ��볯¥
        Dim vntData As Object


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            '1.��ȸ�ʵ� ����
            If strAddFields <> "" Then strAddFields = "," & AddAlias(strAddFields, "A")
            strSelFields = "A.CUST_HDR_ID, A.CUST_HDR_NAME, A.REG_NUM " & strAddFields

            '2.������ ����
            If strCUST_TYPE <> "" Then
                strCondition = String.Format(" AND A.CUST_TYPE IN ('A','{0}') ", strCUST_TYPE)
            End If

            If strCODE_NAME <> "" Then
                '��ȸ ������ �ڵ����� �ڵ������ �����Ͽ� ���� �ʵ� ����
                If IsNumeric(strCODE_NAME) Then '������ ���
                    If Not blnLikeCode Then
                        strCondition &= String.Format("AND A.REG_NUM='{0}'", strCODE_NAME)
                    Else
                        strCondition &= String.Format("AND A.REG_NUM LIKE '{0}%'", strCODE_NAME)
                    End If
                ElseIf IsSBCS(strCODE_NAME) Then    '������ ���
                    strCondition &= String.Format("AND (A.REG_NUM LIKE '{0}%' OR A.CUST_HDR_NAME LIKE '{0}%')", strCODE_NAME)
                Else                                 '�ѱ��� ���
                    strCondition &= String.Format("AND A.CUST_HDR_NAME LIKE '{0}%'", strCODE_NAME)
                End If
            End If

            '2003.12.03- �������� �䱸����-�ŷ�ó ����� ������� �ʴ� �͵� ���δ�.
            '3.������� ���� �˻� (EDATE ������� ����??)
            ' If blnUseOnly Then
            'strChkDate = "AND A.USE_YN='Y'"
            'Else
            '    strChkDate = String.Format("AND (A.USE_YN='Y' OR A.EDATE>={0})", BuildToDate(strUseDate))
            'End If


            '�ӽ� SQL�� ����
            strFormat = "SELECT {0} FROM SC_CUST_HDR A " & _
                        "WHERE A.SC_BU_CODE='{1}' {2} {3} "

            'SQL�� ����
            strSQL = String.Format(strFormat, strSelFields, .USRCompany, strCondition, strChkDate)

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetCustomerHdr")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "������� ���� 17-2. CUSTOMER �ŷ�ó Detail - ���� CUSTOMER �� ����"

    '    Public Function GetCustomerDtl(ByVal strInfoXML As String, _
    '                                    ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
    '                                    Optional ByVal strCODE_NAME As String = "", _
    '                                    Optional ByVal strAddFields As String = "", _
    '                                    Optional ByVal blnUseOnly As Boolean = True, _
    '                                    Optional ByVal strUseDate As String = "", _
    '                                    Optional ByVal blnLikeCode As Boolean = True) As Object

    '        Dim strSQL As String            'SQL��
    '        Dim strFormat As String         '�ӽ� SQL��
    '        Dim strCondition As String      '������
    '        Dim strSelFields As String      '��ȸ�ʵ�
    '        Dim strChkDate As String = ""   '��뿩�� �� ��볯¥
    '        Dim vntData As Object


    '        SetConfig(strInfoXML)   '�⺻���� ����
    '        With mobjSCGLConfig

    '            '1.��ȸ�ʵ� ����
    '            If strAddFields <> "" Then strAddFields = "," & AddAlias(strAddFields, "A")
    '            strSelFields = "A.CUST_ID, A.CUST_NAME, A.REG_NUM " & strAddFields

    '            '2.������ ����
    '            ' If strCUST_TYPE <> "" Then
    '            '   strCondition = String.Format(" AND A.CUST_TYPE IN ('A','{0}') ", strCUST_TYPE)
    '            ' End If

    '            If strCODE_NAME <> "" Then
    '                '��ȸ ������ �ڵ����� �ڵ������ �����Ͽ� ���� �ʵ� ����
    '                If IsNumeric(strCODE_NAME) Then '������ ���
    '                    If Not blnLikeCode Then
    '                        strCondition &= String.Format("AND A.REG_NUM='{0}'", strCODE_NAME)
    '                    Else
    '                        strCondition &= String.Format("AND A.REG_NUM LIKE '{0}%'", strCODE_NAME)
    '                    End If
    '                ElseIf IsSBCS(strCODE_NAME) Then    '������ ���
    '                    strCondition &= String.Format("AND (A.REG_NUM LIKE '{0}%' OR A.CUST_NAME LIKE '{0}%')", strCODE_NAME)
    '                Else                                 '�ѱ��� ���
    '                    strCondition &= String.Format("AND A.CUST_NAME LIKE '{0}%'", strCODE_NAME)
    '                End If
    '            End If


    '            '3.������� ���� �˻� (EDATE ������� ����??)
    '            If blnUseOnly Then
    '                strChkDate = "AND A.USE_YN='Y'"
    '            Else
    '                '    strChkDate = String.Format("AND (A.USE_YN='Y' OR A.EDATE>={0})", BuildToDate(strUseDate))
    '            End If


    '            '�ӽ� SQL�� ����
    '            strFormat = "SELECT {0} FROM SC_CUSTOMER_V A " & _
    '                        "WHERE A.SC_BU_CODE='{1}' {2} {3} " & _
    '                        "ORDER BY A.CUST_NAME"

    '            'SQL�� ����
    '            strSQL = String.Format(strFormat, strSelFields, .USRCompany, strCondition, strChkDate)

    '            '������ ��ȸ
    '            Try
    '                .mobjSCGLSql.SQLConnect(.DBConnStr)
    '                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
    '                Return vntData
    '            Catch err As Exception
    '                Throw RaiseSysErr(err, CLASS_NAME & ".GetCustomer")
    '            Finally
    '                .mobjSCGLSql.SQLDisconnect()
    '            End Try
    '        End With
    '    End Function
#End Region

    '*****************************************************************
    '�Է� : strInfoXML = ����⺻���� XML
    '       intRowCnt,intColCnt = ��ȸ �Ǽ�,�ʵ� ��
    '       strCODE_NAME = A.CUST_NAME, A.REG_NUM
    '       strCUST_TYPE 
    '       blnUseOnly = ���� ������� �͸� �Ǵ� ��ü
    '       strUseDate = blnUseOnly�� False�϶� EDATE>=strUseDate�� �˻�
    '       strAddFields = �ڵ�/�� �̿��� ��ȸ �߰� �ʵ�
    '       blnLikeCode = like �� ����� ���ΰ�? (Default True)
    '��ȯ : ó�����
    '��� : CUST_CONTACTS �ŷ�ó ����� (A.CUST_ID, A.CUST_NAME, A.REG_NUM) �� ��ȸ
    '*****************************************************************
#Region "18. CUST_CONTACTS �ŷ�ó �����"

    Public Function GetCustContacts(ByVal strInfoXML As String, _
                                    ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                    Optional ByVal strCODE_NAME As String = "", _
                                    Optional ByVal strCUST_ID As String = "", _
                                    Optional ByVal strAddFields As String = "", _
                                    Optional ByVal blnUseOnly As Boolean = True, _
                                    Optional ByVal strUseDate As String = "", _
                                    Optional ByVal blnLikeCode As Boolean = True) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strCondition As String      '������
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strChkDate As String = ""   '��뿩�� �� ��볯¥
        Dim vntData As Object


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            '1.��ȸ�ʵ� ����
            If strAddFields <> "" Then strAddFields = "," & AddAlias(strAddFields, "A")
            strSelFields = "A.CUST_CONTACT_ID, A.CONTACT_NAME " & strAddFields

            '2.������ ����
            If strCUST_ID <> "" Then
                strCondition = String.Format(" AND A.CUST_ID ={0} ", strCUST_ID)
            End If

            If strCODE_NAME <> "" Then
                '��ȸ ������ �ڵ����� �ڵ������ �����Ͽ� ���� �ʵ� ����
                If IsNumeric(strCODE_NAME) Then '������ ���
                    If Not blnLikeCode Then
                        strCondition &= String.Format("AND A.CUST_CONTACT_ID={0}", strCODE_NAME)
                    Else
                        strCondition &= String.Format("AND A.CUST_CONTACT_ID LIKE '{0}%'", strCODE_NAME)
                    End If
                ElseIf IsSBCS(strCODE_NAME) Then    '������ ���
                    strCondition &= String.Format("AND (A.CUST_CONTACT_ID LIKE '{0}%' OR A.CONTACT_NAME LIKE '{0}%')", strCODE_NAME)
                Else                                 '�ѱ��� ���
                    strCondition &= String.Format("AND A.CONTACT_NAME LIKE '{0}%'", strCODE_NAME)
                End If
            End If


            '3.������� ���� �˻� (EDATE ������� ����??)
            If blnUseOnly Then
                strChkDate = "AND A.USE_YN='Y'"
            Else
                '    strChkDate = String.Format("AND (A.USE_YN='Y' OR A.EDATE>={0})", BuildToDate(strUseDate))
            End If


            '�ӽ� SQL�� ����
            strFormat = "SELECT {0} FROM SC_CUST_CONTACTS_V A " & _
                        "WHERE A.SC_BU_CODE='{1}' {2} {3} " & _
                        "ORDER BY A.CONTACT_NAME"

            'SQL�� ����
            strSQL = String.Format(strFormat, strSelFields, .USRCompany, strCondition, strChkDate)

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetCustContacts")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#End Region

    '*****************************************************************
    '�Է� : strInfoXML = ����⺻���� XML
    '       intRowCnt,intColCnt = ��ȸ �Ǽ�,�ʵ� ��
    '       strCODE_NAME = A.INV_NAME, A.INV_NUM
    '       strLOC_ID 
    '       blnUseOnly = ���� ������� �͸� �Ǵ� ��ü
    '       strUseDate = blnUseOnly�� False�϶� EDATE>=strUseDate�� �˻�
    '       strAddFields = �ڵ�/�� �̿��� ��ȸ �߰� �ʵ�
    '       blnLikeCode = like �� ����� ���ΰ�? (Default True)
    '��ȯ : ó�����
    '��� : INVENTORY â�� (A.INV_ID, A.INV_CODE, A.INV_NAME) �� ��ȸ
    '*****************************************************************
#Region "19. INVENTORY â��"

    Public Function GetInventory(ByVal strInfoXML As String, _
                                    ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                    Optional ByVal strCODE_NAME As String = "", _
                                    Optional ByVal strLOC_ID As String = "", _
                                    Optional ByVal strAddFields As String = "", _
                                    Optional ByVal blnUseOnly As Boolean = True, _
                                    Optional ByVal strUseDate As String = "", _
                                    Optional ByVal blnLikeCode As Boolean = True) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strCondition As String      '������
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strChkDate As String = ""   '��뿩�� �� ��볯¥
        Dim vntData As Object


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            '1.��ȸ�ʵ� ����
            If strAddFields <> "" Then strAddFields = "," & AddAlias(strAddFields, "A")
            strSelFields = "A.INV_ID, A.INV_CODE, A.INV_NAME, A.CC, A.CC_NAME, A.LOC_ID, A.LOC_CODE, A.LOC_NAME " & strAddFields
            '2.������ ����
            If strLOC_ID <> "" Then
                strCondition = String.Format(" AND A.LOC_ID ={0} ", strLOC_ID)
            End If
            If strCODE_NAME <> "" Then
                '��ȸ ������ �ڵ����� �ڵ������ �����Ͽ� ���� �ʵ� ����
                If IsNumeric(strCODE_NAME) Then '������ ���
                    If Not blnLikeCode Then
                        strCondition &= String.Format("AND A.INV_CODE={0}", strCODE_NAME)
                    Else
                        strCondition &= String.Format("AND A.INV_CODE LIKE '{0}%'", strCODE_NAME)
                    End If
                ElseIf IsSBCS(strCODE_NAME) Then    '������ ���
                    strCondition &= String.Format("AND (A.INV_CODE LIKE '{0}%' OR A.INV_NAME LIKE '{0}%')", strCODE_NAME)
                Else                                 '�ѱ��� ���
                    strCondition &= String.Format("AND A.INV_NAME LIKE '{0}%'", strCODE_NAME)
                End If
            End If
            ''3.������� ���� �˻� (EDATE ������� ����??)
            If blnUseOnly Then
                strChkDate = "AND A.USE_YN='Y'"
            Else
                '    '    strChkDate = String.Format("AND (A.USE_YN='Y' OR A.EDATE>={0})", BuildToDate(strUseDate))
            End If
            '�ӽ� SQL�� ����
            strFormat = "SELECT {0} FROM SC_INVENTORY_V A " & _
                        "WHERE A.SC_BU_CODE='{1}' {2} {3} " & _
                        "ORDER BY A.LOC_CODE, A.INV_CODE "
            'SQL�� ����
            strSQL = String.Format(strFormat, strSelFields, .USRCompany, strCondition, strChkDate)
            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetInventory")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#End Region

    '*****************************************************************
    '�Է� : strInfoXML = ����⺻���� XML
    '       intRowCnt,intColCnt = ��ȸ �Ǽ�,�ʵ� ��
    '       strCODE_NAME = A.INV_NAME, A.INV_NUM
    '       strSC_CATEGORY_GROUP 
    '       blnUseOnly = ���� ������� �͸� �Ǵ� ��ü
    '       strUseDate = blnUseOnly�� False�϶� EDATE>=strUseDate�� �˻�
    '       strAddFields = �ڵ�/�� �̿��� ��ȸ �߰� �ʵ�
    '       blnLikeCode = like �� ����� ���ΰ�? (Default True)
    '��ȯ : ó�����
    '��� : CATEGORY ITEM CATEGORY (A.CATEGORY_ID, A.CATEGORY_NAME) �� ��ȸ
    '*****************************************************************
#Region "20. CATEGORY : ITEM CATEGORY"
    Public Function GetCategory(ByVal strInfoXML As String, _
                                 ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                 Optional ByVal strCODE_NAME As String = "", _
                                 Optional ByVal strSC_CATEGORY_GROUP As String = "", _
                                 Optional ByVal strAddFields As String = "", _
                                 Optional ByVal blnUseOnly As Boolean = True, _
                                 Optional ByVal strUseDate As String = "", _
                                 Optional ByVal blnLikeCode As Boolean = True) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strCondition As String      '������
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strChkDate As String = ""   '��뿩�� �� ��볯¥
        Dim vntData As Object

        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            '1.��ȸ�ʵ� ����
            If strAddFields <> "" Then strAddFields = "," & AddAlias(strAddFields, "A")
            strSelFields = "A.SC_CATEGORY_GROUP, A.SC_CATEGORY_GROUP_NAME, A.CATEGORY_ID, A.CATEGORY_NAME, A.CATEGORY_DESC " & strAddFields

            '2.������ ����
            If strSC_CATEGORY_GROUP <> "" Then
                strCondition = String.Format(" AND A.SC_CATEGORY_GROUP LIKE '{0}%' ", strSC_CATEGORY_GROUP)
            End If

            If strCODE_NAME <> "" Then
                '��ȸ ������ �ڵ����� �ڵ������ �����Ͽ� ���� �ʵ� ����
                If IsNumeric(strCODE_NAME) Then '������ ���
                    If Not blnLikeCode Then
                        strCondition &= String.Format("AND A.CATEGORY_ID={0}", strCODE_NAME)
                    Else
                        strCondition &= String.Format("AND A.CATEGORY_ID LIKE '{0}%'", strCODE_NAME)
                    End If
                ElseIf IsSBCS(strCODE_NAME) Then    '������ ���
                    strCondition &= String.Format("AND (A.CATEGORY_ID LIKE '{0}%' OR A.CATEGORY_NAME LIKE '{0}%')", strCODE_NAME)
                Else                                 '�ѱ��� ���
                    strCondition &= String.Format("AND A.CATEGORY_NAME LIKE '{0}%'", strCODE_NAME)
                End If
            End If
            '3.������� ���� �˻� (EDATE ������� ����??)
            If blnUseOnly Then
                strChkDate = "AND A.USE_YN='Y'"
            Else
                ' strChkDate = String.Format("AND (A.USE_YN='Y' OR A.EDATE>={0})", BuildToDate(strUseDate))
            End If
            '�ӽ� SQL�� ����
            strFormat = "SELECT {0} FROM SC_CATEGORY_V A " & _
                        "WHERE A.SC_BU_CODE='{1}' {2} {3} " & _
                        "ORDER BY A.CATEGORY_ID"

            'SQL�� ����
            strSQL = String.Format(strFormat, strSelFields, .USRCompany, strCondition, strChkDate)

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetCategory")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#End Region

    '*****************************************************************
    '�Է� : strInfoXML = ����⺻���� XML
    '       intRowCnt,intColCnt = ��ȸ �Ǽ�,�ʵ� ��
    '       strCODE_NAME = A.INV_NAME, A.INV_NUM
    '       strCATEGORY
    '       blnUseOnly = ���� ������� �͸� �Ǵ� ��ü
    '       strUseDate = blnUseOnly�� False�϶� EDATE>=strUseDate�� �˻�
    '       strAddFields = �ڵ�/�� �̿��� ��ȸ �߰� �ʵ�
    '       blnLikeCode = like �� ����� ���ΰ�? (Default True)
    '��ȯ : ó�����
    '��� : ITEM CATEGORY ������� (A.SC_CATALOG_CODE, A.CAT_ELEMENT_NAME, A.CAT_ELEMENT_SEQ) �� ��ȸ
    '*****************************************************************
#Region "21. CAT_ELEMENT: ITEM CATEGORY �������"

    Public Function GetCatElement(ByVal strInfoXML As String, _
                                ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                Optional ByVal strCODE_NAME As String = "", _
                                Optional ByVal strSC_CATALOG As String = "", _
                                Optional ByVal strAddFields As String = "", _
                                Optional ByVal blnUseOnly As Boolean = True, _
                                Optional ByVal strUseDate As String = "", _
                                Optional ByVal blnLikeCode As Boolean = True) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strCondition As String      '������
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strChkDate As String = ""   '��뿩�� �� ��볯¥
        Dim vntData As Object


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            '1.��ȸ�ʵ� ����
            If strAddFields <> "" Then strAddFields = "," & AddAlias(strAddFields, "A")
            strSelFields = "A.SC_CATALOG, A.CAT_ELEMENT_NAME, A.CAT_ELEMENT_SEQ " & strAddFields

            '2.������ ����
            If strSC_CATALOG <> "" Then
                strCondition = String.Format(" AND A.SC_CATALOG ='{0}' ", strSC_CATALOG)
            End If

            If strCODE_NAME <> "" Then
                strCondition &= String.Format("AND A.CAT_ELEMENT_NAME LIKE '{0}%'", strCODE_NAME)
                ''��ȸ ������ �ڵ����� �ڵ������ �����Ͽ� ���� �ʵ� ����
                'If IsNumeric(strCODE_NAME) Then '������ ���
                '    If Not blnLikeCode Then
                '        strCondition = String.Format("AND A.CAT_ELEMENT_NAME={0}", strCODE_NAME)
                '    Else
                '        strCondition = String.Format("AND A.CAT_ELEMENT_NAME LIKE '{0}%'", strCODE_NAME)
                '    End If
                'ElseIf IsSBCS(strCODE_NAME) Then    '������ ���
                '    strCondition = String.Format("AND A.CAT_ELEMENT_NAME LIKE '{0}%'", strCODE_NAME)
                'Else                                 '�ѱ��� ���
                '    strCondition = String.Format("AND AND A.CAT_ELEMENT_NAME LIKE '{0}%'", strCODE_NAME)
                'End If
            End If


            '3.������� ���� �˻� (EDATE ������� ����??)
            If blnUseOnly Then
                strChkDate = "AND A.USE_YN='Y'"
            Else
                '    strChkDate = String.Format("AND (A.USE_YN='Y' OR A.EDATE>={0})", BuildToDate(strUseDate))
            End If


            '�ӽ� SQL�� ����
            strFormat = "SELECT {0} FROM SC_CAT_ELEMENT_V A " & _
                        "WHERE A.SC_BU_CODE='{1}' {2} {3} " & _
                        "ORDER BY A.CAT_ELEMENT_SEQ"

            'SQL�� ����
            strSQL = String.Format(strFormat, strSelFields, .USRCompany, strCondition, strChkDate)

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetCatElement")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#End Region

    '*****************************************************************
    '�Է� : strInfoXML = ����⺻���� XML
    '       intRowCnt,intColCnt = ��ȸ �Ǽ�,�ʵ� ��
    '       strCODE_NAME = A.ITEM_CODE, A.ITEM_NAME
    '       blnUseOnly = ���� ������� �͸� �Ǵ� ��ü
    '       strUseDate = blnUseOnly�� False�϶� EDATE>=strUseDate�� �˻�
    '       strAddFields = �ڵ�/�� �̿��� ��ȸ �߰� �ʵ�
    '       blnLikeCode = like �� ����� ���ΰ�? (Default True)
    '��ȯ : ó�����
    '��� : ITEM  (A.ITEM_ID, A.ITEM_CODE, A.ITEM_NAME, A.UOM_CODE) �� ��ȸ
    '*****************************************************************
#Region "22. ITEM_MST : ITEM"

    Public Function GetItemMst(ByVal strInfoXML As String, _
                                ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                Optional ByVal strCODE_NAME As String = "", _
                                Optional ByVal strAddFields As String = "", _
                                Optional ByVal blnUseOnly As Boolean = True, _
                                Optional ByVal strUseDate As String = "", _
                                Optional ByVal blnLikeCode As Boolean = True, _
                                Optional ByVal strAddWhere As String = "") As Object
        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strCondition As String      '������
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strChkDate As String = ""   '��뿩�� �� ��볯¥
        Dim vntData As Object

        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            '1.��ȸ�ʵ� ����
            If strAddFields <> "" Then strAddFields = "," & AddAlias(strAddFields, "A")
            strSelFields = "A.ITEM_ID, A.ITEM_CODE, A.ITEM_NAME, A.UOM_CODE " & strAddFields

            '2.������ ����

            If strCODE_NAME <> "" Then
                '��ȸ ������ �ڵ����� �ڵ������ �����Ͽ� ���� �ʵ� ����
                If IsNumeric(strCODE_NAME) Then '������ ���
                    If Not blnLikeCode Then
                        strCondition = String.Format("AND A.ITEM_CODE={0}", strCODE_NAME)
                    Else
                        strCondition = String.Format("AND A.ITEM_CODE LIKE '{0}%'", strCODE_NAME)
                    End If
                ElseIf IsSBCS(strCODE_NAME) Then    '������ ���
                    strCondition = String.Format("AND (A.ITEM_CODE LIKE '{0}%' OR A.ITEM_NAME LIKE '{0}%')", strCODE_NAME)
                Else                                 '�ѱ��� ���
                    strCondition = String.Format("AND A.ITEM_NAME LIKE '{0}%'", strCODE_NAME)
                End If
            End If


            '3.������� ���� �˻� (EDATE ������� ����??)
            If blnUseOnly Then
                strChkDate = "AND A.USE_YN='Y'"
            Else
                '    '    strChkDate = String.Format("AND (A.USE_YN='Y' OR A.EDATE>={0})", BuildToDate(strUseDate))
            End If

            '�ӽ� SQL�� ����
            strFormat = "SELECT {0} FROM SC_ITEM_MST A " & _
                        "WHERE A.SC_BU_CODE='{1}' {2} {3} {4}" & _
                        "ORDER BY A.ITEM_CODE"

            'SQL�� ����
            strSQL = String.Format(strFormat, strSelFields, .USRCompany, strCondition, strChkDate, strAddWhere)

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetItemMst")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

    '*****************************************************************
    '�Է� : strInfoXML = ����⺻���� XML
    '       intRowCnt,intColCnt = ��ȸ �Ǽ�,�ʵ� ��
    '       strCODE_NAME = A.UOM_CODE, A.UOM_NAME
    '       strSC_UOM_CLASS
    '       blnUseOnly = ���� ������� �͸� �Ǵ� ��ü
    '       strUseDate = blnUseOnly�� False�϶� EDATE>=strUseDate�� �˻�
    '       strAddFields = �ڵ�/�� �̿��� ��ȸ �߰� �ʵ�
    '       blnLikeCode = like �� ����� ���ΰ�? (Default True)
    '��ȯ : ó�����
    '��� : UOM  (A.UOM_ID, A.UOM_CODE, A.UOM_NAME, A.SC_UOM_CLASS) �� ��ȸ
    '*****************************************************************
#Region "23. UOM"

    Public Function GetUOM(ByVal strInfoXML As String, _
                            ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                            Optional ByVal strCODE_NAME As String = "", _
                            Optional ByVal strSC_UOM_CLASS As String = "", _
                            Optional ByVal strAddFields As String = "", _
                            Optional ByVal blnUseOnly As Boolean = True, _
                            Optional ByVal strUseDate As String = "", _
                            Optional ByVal blnLikeCode As Boolean = True) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strCondition As String      '������
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strChkDate As String = ""   '��뿩�� �� ��볯¥
        Dim vntData As Object


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            '1.��ȸ�ʵ� ����
            If strAddFields <> "" Then strAddFields = "," & AddAlias(strAddFields, "A")
            strSelFields = "A.UOM_ID, A.UOM_CODE, A.UOM_NAME, A.SC_UOM_CLASS " & strAddFields

            '2.������ ����
            If strSC_UOM_CLASS <> "" Then
                strCondition = String.Format("AND A.SC_UOM_CLASS = '{0}'", strSC_UOM_CLASS)
            End If

            If strCODE_NAME <> "" Then
                '��ȸ ������ �ڵ����� �ڵ������ �����Ͽ� ���� �ʵ� ����
                If IsNumeric(strCODE_NAME) Then '������ ���
                    If Not blnLikeCode Then
                        strCondition &= String.Format("AND A.UOM_CODE={0}", strCODE_NAME)
                    Else
                        strCondition &= String.Format("AND A.UOM_CODE LIKE '{0}%'", strCODE_NAME)
                    End If
                ElseIf IsSBCS(strCODE_NAME) Then    '������ ���
                    If Not blnLikeCode Then
                        strCondition &= String.Format("AND A.UOM_CODE='{0}'", strCODE_NAME)
                    Else
                        strCondition &= String.Format("AND (A.UOM_CODE LIKE '{0}%' OR A.UOM_NAME LIKE '{0}%')", strCODE_NAME)
                    End If
                Else                                 '�ѱ��� ���
                    strCondition &= String.Format("AND A.UOM_NAME LIKE '{0}%'", strCODE_NAME)
                End If
            End If


            '3.������� ���� �˻� (EDATE ������� ����??)
            If blnUseOnly Then
                strChkDate = "AND A.USE_YN='Y'"
            Else
                '    '    strChkDate = String.Format("AND (A.USE_YN='Y' OR A.EDATE>={0})", BuildToDate(strUseDate))
            End If


            '�ӽ� SQL�� ����
            strFormat = "SELECT {0} FROM SC_UOM_V A " & _
                        "WHERE A.SC_BU_CODE='{1}' {2} {3} " & _
                        "ORDER BY A.UOM_CODE"

            'SQL�� ����
            strSQL = String.Format(strFormat, strSelFields, .USRCompany, strCondition, strChkDate)

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetUOM")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#End Region

    '*****************************************************************
    '�Է� : strInfoXML = ����⺻���� XML
    '       intRowCnt,intColCnt = ��ȸ �Ǽ�,�ʵ� ��
    '       strCODE_NAME = A.SFX_CODE, A.SFX_NAME
    '       strACC1
    '       blnUseOnly = ���� ������� �͸� �Ǵ� ��ü
    '       strUseDate = blnUseOnly�� False�϶� EDATE>=strUseDate�� �˻�
    '       strAddFields = �ڵ�/�� �̿��� ��ȸ �߰� �ʵ�
    '       blnLikeCode = like �� ����� ���ΰ�? (Default True)
    '��ȯ : ó�����
    '��� : SUFFIX_CODE ����ڵ� (A.SFX_CODE, A.SFX_NAME) �� ��ȸ
    '*****************************************************************
#Region "24. SUFFIX_CODE: ����ڵ�"

    Public Function GetSuffixCode(ByVal strInfoXML As String, _
                                ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                Optional ByVal strCODE_NAME As String = "", _
                                Optional ByVal strACC1 As String = "", _
                                Optional ByVal strAddFields As String = "", _
                                Optional ByVal blnUseOnly As Boolean = True, _
                                Optional ByVal strUseDate As String = "", _
                                Optional ByVal blnLikeCode As Boolean = True) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strCondition As String      '������
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strChkDate As String = ""   '��뿩�� �� ��볯¥
        Dim vntData As Object


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            '1.��ȸ�ʵ� ����
            If strAddFields <> "" Then strAddFields = "," & AddAlias(strAddFields, "A")
            strSelFields = "A.SFX_CODE, A.SFX_NAME " & strAddFields

            '2.������ ����
            If strACC1 <> "" Then
                strCondition = String.Format("AND A.ACC1= '{0}'", strACC1)
            End If

            If strCODE_NAME <> "" Then
                '��ȸ ������ �ڵ����� �ڵ������ �����Ͽ� ���� �ʵ� ����
                If IsNumeric(strCODE_NAME) Then '������ ���
                    If Not blnLikeCode Then
                        strCondition &= String.Format("AND A.SFX_CODE={0}", strCODE_NAME)
                    Else
                        strCondition &= String.Format("AND A.SFX_CODE LIKE '{0}%'", strCODE_NAME)
                    End If
                ElseIf IsSBCS(strCODE_NAME) Then    '������ ���
                    strCondition &= String.Format("AND (A.SFX_CODE LIKE '{0}%' OR A.SFX_NAME LIKE '{0}%')", strCODE_NAME)
                Else                                 '�ѱ��� ���
                    strCondition &= String.Format("AND A.SFX_NAME LIKE '{0}%'", strCODE_NAME)
                End If
            End If


            '3.������� ���� �˻� (EDATE ������� ����??)
            If blnUseOnly Then
                strChkDate = "AND A.USE_YN='Y'"
            Else
                strChkDate = String.Format("AND (A.USE_YN='Y' OR A.EDATE>={0})", BuildToDate(strUseDate))
            End If


            '�ӽ� SQL�� ����
            strFormat = "SELECT {0} FROM SC_SFX_CODE A " & _
                        "WHERE A.SC_BU_CODE='{1}' {2} {3} " & _
                        "ORDER BY A.SFX_CODE"

            'SQL�� ����
            strSQL = String.Format(strFormat, strSelFields, .USRCompany, strCondition, strChkDate)

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetSuffixCode")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#End Region

    '*****************************************************************
    '�Է� : strInfoXML = ����⺻���� XML
    '       intRowCnt,intColCnt = ��ȸ �Ǽ�,�ʵ� ��
    '       strCODE_NAME = A.PAY_CODE, A.PAY_NAME
    '       blnUseOnly = ���� ������� �͸� �Ǵ� ��ü
    '       strUseDate = blnUseOnly�� False�϶� EDATE>=strUseDate�� �˻�
    '       strAddFields = �ڵ�/�� �̿��� ��ȸ �߰� �ʵ�
    '       blnLikeCode = like �� ����� ���ΰ�? (Default True)
    '��ȯ : ó�����
    '��� : Pay_Condition �������� (A.PAY_CODE, A.PAY_NAME) �� ��ȸ
    '*****************************************************************
#Region "25. Pay_Cond ��������"

    Public Function GetPayCondition(ByVal strInfoXML As String, _
                                ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                Optional ByVal strCODE_NAME As String = "", _
                                Optional ByVal strAddFields As String = "", _
                                Optional ByVal blnUseOnly As Boolean = True, _
                                Optional ByVal strUseDate As String = "", _
                                Optional ByVal blnLikeCode As Boolean = True) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strCondition As String      '������
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strChkDate As String = ""   '��뿩�� �� ��볯¥
        Dim vntData As Object


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            '1.��ȸ�ʵ� ����
            If strAddFields <> "" Then strAddFields = "," & AddAlias(strAddFields, "A")
            strSelFields = "A.PAY_CODE, A.PAY_NAME, A.REGULAR_PAY_DD, A.PAY_TERM, A.NOTES_DAYS, A.PAY_NA_CODE, A.CASH_NOTES_CLASS " & strAddFields

            '2.������ ����
            If strCODE_NAME <> "" Then
                '��ȸ ������ �ڵ����� �ڵ������ �����Ͽ� ���� �ʵ� ����
                If IsNumeric(strCODE_NAME) Then '������ ���
                    If Not blnLikeCode Then
                        strCondition = String.Format("AND A.PAY_CODE={0}", strCODE_NAME)
                    Else
                        strCondition = String.Format("AND A.PAY_CODE LIKE '{0}%'", strCODE_NAME)
                    End If
                ElseIf IsSBCS(strCODE_NAME) Then    '������ ���
                    strCondition = String.Format("AND (A.PAY_CODE LIKE '{0}%' OR A.PAY_NAME LIKE '{0}%')", strCODE_NAME)
                Else                                 '�ѱ��� ���
                    strCondition = String.Format("AND A.PAY_NAME LIKE '{0}%'", strCODE_NAME)
                End If
            End If


            '3.������� ���� �˻� (EDATE ������� ����??)
            If blnUseOnly Then
                strChkDate = "AND A.USE_YN='Y'"
            Else
                '' strChkDate = String.Format("AND (A.USE_YN='Y' OR A.EDATE>={0})", BuildToDate(strUseDate))
            End If


            '�ӽ� SQL�� ����
            strFormat = "SELECT {0} FROM SC_PAY_CONDITION_V A " & _
                        "WHERE A.SC_BU_CODE='{1}' {2} {3} " & _
                        "ORDER BY A.PAY_CODE"

            'SQL�� ����
            strSQL = String.Format(strFormat, strSelFields, .USRCompany, strCondition, strChkDate)

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetPayCondition")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#End Region
    '*****************************************************************
    '�Է� : strInfoXML = ����⺻���� XML
    '       intRowCnt,intColCnt = ��ȸ �Ǽ�,�ʵ� ��
    '       strCODE_NAME = A.PAY_GRP_CODE, A.PAY_GRP_NAME
    '       blnUseOnly = ���� ������� �͸� �Ǵ� ��ü
    '       strUseDate = blnUseOnly�� False�϶� EDATE>=strUseDate�� �˻�
    '       strAddFields = �ڵ�/�� �̿��� ��ȸ �߰� �ʵ�
    '       blnLikeCode = like �� ����� ���ΰ�? (Default True)
    '��ȯ : ó�����
    '��� : PayGroup ���ұ׷� (A.PAY_GRP_CODE, A.PAY_GRP_NAME) �� ��ȸ
    '*****************************************************************
#Region "25-1. PayGroup ���ұ׷�"

    Public Function GetPayGroup(ByVal strInfoXML As String, _
                                ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                Optional ByVal strCODE_NAME As String = "", _
                                Optional ByVal strAddFields As String = "", _
                                Optional ByVal blnUseOnly As Boolean = True, _
                                Optional ByVal strUseDate As String = "", _
                                Optional ByVal blnLikeCode As Boolean = True) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strCondition As String      '������
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strChkDate As String = ""   '��뿩�� �� ��볯¥
        Dim vntData As Object


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            '1.��ȸ�ʵ� ����
            If strAddFields <> "" Then strAddFields = "," & AddAlias(strAddFields, "A")
            strSelFields = "A.PAY_GRP_CODE,A.PAY_GRP_NAME,A.NEW_PAY_CODE,A.CURR_CODE " & strAddFields

            '2.������ ����
            If strCODE_NAME <> "" Then
                '��ȸ ������ �ڵ����� �ڵ������ �����Ͽ� ���� �ʵ� ����
                If IsNumeric(strCODE_NAME) Then '������ ���
                    If Not blnLikeCode Then
                        strCondition = String.Format("AND A.PAY_GRP_CODE={0}", strCODE_NAME)
                    Else
                        strCondition = String.Format("AND A.PAY_GRP_CODE LIKE '{0}%'", strCODE_NAME)
                    End If
                ElseIf IsSBCS(strCODE_NAME) Then    '������ ���
                    strCondition = String.Format("AND (A.PAY_GRP_CODE LIKE '{0}%' OR A.PAY_GRP_NAME LIKE '{0}%')", strCODE_NAME)
                Else                                 '�ѱ��� ���
                    strCondition = String.Format("AND A.PAY_GRP_NAME LIKE '{0}%'", strCODE_NAME)
                End If
            End If


            '3.������� ���� �˻� (EDATE ������� ����??)
            If blnUseOnly Then
                strChkDate = "AND A.USE_YN='Y'"
            Else
                strChkDate = String.Format("AND (A.USE_YN='Y' OR A.EDATE>={0})", BuildToDate(strUseDate))
            End If


            '�ӽ� SQL�� ����
            strFormat = "SELECT {0} FROM SC_PAY_GROUP A " & _
                        "WHERE A.SC_BU_CODE='{1}' {2} {3} " & _
                        "ORDER BY A.PAY_GRP_CODE"

            'SQL�� ����
            strSQL = String.Format(strFormat, strSelFields, .USRCompany, strCondition, strChkDate)

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetPayGroup")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#End Region

    '*****************************************************************
    '�Է� : strInfoXML = ����⺻���� XML
    '       intRowCnt,intColCnt = ��ȸ �Ǽ�,�ʵ� ��
    '       strCODE_NAME = A.NA_TYPE_CODE, A.NA_TYPE_NAME
    '       blnUseOnly = ���� ������� �͸� �Ǵ� ��ü
    '       strUseDate = blnUseOnly�� False�϶� EDATE>=strUseDate�� �˻�
    '       strAddFields = �ڵ�/�� �̿��� ��ȸ �߰� �ʵ�
    '       blnLikeCode = like �� ����� ���ΰ�? (Default True)
    '��ȯ : ó�����
    '��� : NA_TYPE  (A.NA_TYPE_CODE, A.NA_TYPE_NAME) �� ��ȸ
    '*****************************************************************
#Region "26. NA TYPE ��ȸ"

    Public Function GetNaType(ByVal strInfoXML As String, _
                                ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                Optional ByVal strCODE_NAME As String = "", _
                                Optional ByVal strAddFields As String = "", _
                                Optional ByVal blnUseOnly As Boolean = True, _
                                Optional ByVal strUseDate As String = "", _
                                Optional ByVal blnLikeCode As Boolean = True) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strCondition As String      '������
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strChkDate As String = ""   '��뿩�� �� ��볯¥
        Dim vntData As Object


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            '1.��ȸ�ʵ� ����
            If strAddFields <> "" Then strAddFields = "," & AddAlias(strAddFields, "A")
            strSelFields = "A.NA_TYPE_CODE, A.NA_TYPE_NAME " & strAddFields

            '2.������ ����
            If strCODE_NAME <> "" Then
                '��ȸ ������ �ڵ����� �ڵ������ �����Ͽ� ���� �ʵ� ����
                If IsNumeric(strCODE_NAME) Then '������ ���
                    If Not blnLikeCode Then
                        strCondition = String.Format("AND A.NA_TYPE_CODE={0}", strCODE_NAME)
                    Else
                        strCondition = String.Format("AND A.NA_TYPE_CODE LIKE '{0}%'", strCODE_NAME)
                    End If
                ElseIf IsSBCS(strCODE_NAME) Then    '������ ���
                    strCondition = String.Format("AND (A.NA_TYPE_CODE LIKE '{0}%' OR A.NA_TYPE_NAME LIKE '{0}%')", strCODE_NAME)
                Else                                 '�ѱ��� ���
                    strCondition = String.Format("AND A.NA_TYPE_NAME LIKE '{0}%'", strCODE_NAME)
                End If
            End If


            '3.������� ���� �˻� (EDATE ������� ����??)
            '    If blnUseOnly Then
            '         strChkDate = "AND A.USE_YN='Y'"
            '    Else
            '        strChkDate = String.Format("AND (A.USE_YN='Y' OR A.EDATE>={0})", BuildToDate(strUseDate))
            '   End If


            '�ӽ� SQL�� ����
            strFormat = "SELECT {0} FROM SC_NA_TYPE_V A " & _
                        "WHERE A.SC_BU_CODE='{1}' {2} {3} " & _
                        "ORDER BY A.NA_TYPE_CODE"

            'SQL�� ����
            strSQL = String.Format(strFormat, strSelFields, .USRCompany, strCondition, strChkDate)

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetNaType")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#End Region

    '*****************************************************************
    '�Է� : strInfoXML = ����⺻���� XML
    '       intRowCnt,intColCnt = ��ȸ �Ǽ�,�ʵ� ��
    '       strCODE_NAME = A.NA_CODE, A.NA_NAME :: A.PC = 'C' ����
    '       strNA_TYPE_CODE = NA_TYPE 
    '       blnUseOnly = ���� ������� �͸� �Ǵ� ��ü
    '       strUseDate = blnUseOnly�� False�϶� EDATE>=strUseDate�� �˻�
    '       strAddFields = �ڵ�/�� �̿��� ��ȸ �߰� �ʵ�
    '       blnLikeCode = like �� ����� ���ΰ�? (Default True)
    '��ȯ : ó�����
    '��� : NA ��ȸ  (A.NA_CODE, A.NA_NAME) �� ��ȸ
    '*****************************************************************
#Region "27. NA ��ȸ A.PC='Y'"

    Public Function GetNa(ByVal strInfoXML As String, _
                                ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                Optional ByVal strCODE_NAME As String = "", _
                                Optional ByVal strNA_TYPE_CODE As String = "", _
                                Optional ByVal strAddFields As String = "", _
                                Optional ByVal blnUseOnly As Boolean = True, _
                                Optional ByVal strUseDate As String = "", _
                                Optional ByVal blnLikeCode As Boolean = True) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strCondition As String      '������
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strChkDate As String = ""   '��뿩�� �� ��볯¥
        Dim vntData As Object


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            '1.��ȸ�ʵ� ����
            If strAddFields <> "" Then strAddFields = "," & AddAlias(strAddFields, "A")
            strSelFields = "A.NA_CODE, A.NA_NAME, GA_SEG4_NAME_FUN(A.NA_CODE,'" & .USRCompany & "') AS NA_SEG4_NAME " & strAddFields

            '2.������ ����
            If strNA_TYPE_CODE <> "" Then strCondition = String.Format("AND A.NA_TYPE_CODE='{0}' ", strNA_TYPE_CODE)

            If strCODE_NAME <> "" Then
                '��ȸ ������ �ڵ����� �ڵ������ �����Ͽ� ���� �ʵ� ����
                If IsNumeric(strCODE_NAME) Then '������ ���
                    If Not blnLikeCode Then
                        strCondition &= String.Format("AND A.NA_CODE={0}", strCODE_NAME)
                    Else
                        strCondition &= String.Format("AND A.NA_CODE LIKE '{0}%'", strCODE_NAME)
                    End If
                ElseIf IsSBCS(strCODE_NAME) Then    '������ ���
                    strCondition &= String.Format("AND (A.NA_CODE LIKE '{0}%' OR A.NA_NAME LIKE '{0}%')", strCODE_NAME)
                Else                                 '�ѱ��� ���
                    strCondition &= String.Format("AND A.NA_NAME LIKE '{0}%'", strCODE_NAME)
                End If
            End If

            '3.������� ���� �˻� (EDATE ������� ����??)
            If blnUseOnly Then
                strChkDate = "AND A.USE_YN='Y'"
                'Else
                '    strChkDate = String.Format("AND (A.USE_YN='Y' OR A.EDATE>={0})", BuildToDate(strUseDate))
            End If


            '�ӽ� SQL�� ����
            strFormat = "SELECT {0} FROM SC_NA_V A " & _
                        "WHERE A.SC_BU_CODE='{1}' AND A.PC = 'Y' {2} {3} " & _
                        "ORDER BY A.NA_CODE"

            'SQL�� ����
            strSQL = String.Format(strFormat, strSelFields, .USRCompany, strCondition, strChkDate)

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetNa")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#End Region

    '*****************************************************************
    '�Է� : strInfoXML = ����⺻���� XML
    '       intRowCnt,intColCnt = ��ȸ �Ǽ�,�ʵ� ��
    '       strCODE_NAME = A.NA_CODE, A.NA_NAME :: A.PC = 'P' ����
    '       strNA_LEVEL = NA_LEVEL 
    '       strNA_TYPE_CODE = NA_TYPE 
    '       blnUseOnly = ���� ������� �͸� �Ǵ� ��ü
    '       strUseDate = blnUseOnly�� False�϶� EDATE>=strUseDate�� �˻�
    '       strAddFields = �ڵ�/�� �̿��� ��ȸ �߰� �ʵ�
    '       blnLikeCode = like �� ����� ���ΰ�? (Default True)
    '��ȯ : ó�����
    '��� : NA ��ȸ  (A.NA_CODE, A.NA_NAME) �� ��ȸ
    '*****************************************************************
#Region "27-1. NA1 ��ȸ( A.PC = 'N' )"

    Public Function GetNa1(ByVal strInfoXML As String, _
                                ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                Optional ByVal strCODE_NAME As String = "", _
                                Optional ByVal strNA_LEVEL As String = "", _
                                Optional ByVal strNA_TYPE_CODE As String = "", _
                                Optional ByVal strAddFields As String = "", _
                                Optional ByVal blnUseOnly As Boolean = True, _
                                Optional ByVal strUseDate As String = "", _
                                Optional ByVal blnLikeCode As Boolean = True) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strCondition As String      '������
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strChkDate As String = ""   '��뿩�� �� ��볯¥
        Dim vntData As Object


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            '1.��ȸ�ʵ� ����
            If strAddFields <> "" Then strAddFields = "," & AddAlias(strAddFields, "A")
            strSelFields = "A.NA_CODE, A.NA_NAME " & strAddFields

            '2.������ ����
            If strNA_LEVEL <> "" Then strCondition = String.Format("AND A.NA_LEVEL ={0} ", strNA_LEVEL)
            If strNA_TYPE_CODE <> "" Then strCondition &= String.Format("AND A.NA_TYPE_CODE='{0}' ", strNA_TYPE_CODE)

            If strCODE_NAME <> "" Then
                '��ȸ ������ �ڵ����� �ڵ������ �����Ͽ� ���� �ʵ� ����
                If IsNumeric(strCODE_NAME) Then '������ ���
                    If Not blnLikeCode Then
                        strCondition &= String.Format("AND A.NA_CODE={0}", strCODE_NAME)
                    Else
                        strCondition &= String.Format("AND A.NA_CODE LIKE '{0}%'", strCODE_NAME)
                    End If
                ElseIf IsSBCS(strCODE_NAME) Then    '������ ���
                    strCondition &= String.Format("AND (A.NA_CODE LIKE '{0}%' OR A.NA_NAME LIKE '{0}%')", strCODE_NAME)
                Else                                 '�ѱ��� ���
                    strCondition &= String.Format("AND A.NA_NAME LIKE '{0}%'", strCODE_NAME)
                End If
            End If

            '3.������� ���� �˻� (EDATE ������� ����??)
            If blnUseOnly Then
                strChkDate = "AND A.USE_YN='Y'"
                'Else
                '    strChkDate = String.Format("AND (A.USE_YN='Y' OR A.EDATE>={0})", BuildToDate(strUseDate))
            End If


            '�ӽ� SQL�� ����
            strFormat = "SELECT {0} FROM SC_NA_V A " & _
                        "WHERE A.SC_BU_CODE='{1}' AND A.PC = 'N' {2} {3} " & _
                        "ORDER BY A.NA_CODE"

            'SQL�� ����
            strSQL = String.Format(strFormat, strSelFields, .USRCompany, strCondition, strChkDate)

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetNa1")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#End Region

    '*****************************************************************
    '�Է� : strInfoXML = ����⺻���� XML
    '       intRowCnt,intColCnt = ��ȸ �Ǽ�,�ʵ� ��
    '       strCODE_NAME = A.NA_CODE, A.NA_NAME :: A.PC = 'P' ����
    '       strNA_LEVEL = NA_LEVEL 
    '       strNA_TYPE_CODE = NA_TYPE 
    '       blnUseOnly = ���� ������� �͸� �Ǵ� ��ü
    '       strUseDate = blnUseOnly�� False�϶� EDATE>=strUseDate�� �˻�
    '       strAddFields = �ڵ�/�� �̿��� ��ȸ �߰� �ʵ�
    '       blnLikeCode = like �� ����� ���ΰ�? (Default True)
    '��ȯ : ó�����
    '��� : NA ALL ��ȸ  (A.NA_CODE, A.NA_NAME) �� ��ȸ
    '*****************************************************************
#Region "27-2. NA ALL ��ȸ( ALL )"

    Public Function GetNaAll(ByVal strInfoXML As String, _
                             ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                             Optional ByVal strCODE_NAME As String = "", _
                             Optional ByVal strNA_LEVEL As String = "", _
                             Optional ByVal strNA_TYPE_CODE As String = "", _
                             Optional ByVal strAddFields As String = "", _
                             Optional ByVal blnUseOnly As Boolean = True, _
                             Optional ByVal strUseDate As String = "", _
                             Optional ByVal blnLikeCode As Boolean = True) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strCondition As String      '������
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strChkDate As String = ""   '��뿩�� �� ��볯¥
        Dim vntData As Object


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            '1.��ȸ�ʵ� ����
            If strAddFields <> "" Then strAddFields = "," & AddAlias(strAddFields, "A")
            strSelFields = "A.NA_CODE, A.NA_NAME " & strAddFields

            '2.������ ����
            If strNA_LEVEL <> "" Then strCondition = String.Format("AND A.NA_LEVEL ={0} ", strNA_LEVEL)
            If strNA_TYPE_CODE <> "" Then strCondition &= String.Format("AND A.NA_TYPE_CODE='{0}' ", strNA_TYPE_CODE)

            If strCODE_NAME <> "" Then
                '��ȸ ������ �ڵ����� �ڵ������ �����Ͽ� ���� �ʵ� ����
                If IsNumeric(strCODE_NAME) Then '������ ���
                    If Not blnLikeCode Then
                        strCondition &= String.Format("AND A.NA_CODE={0}", strCODE_NAME)
                    Else
                        strCondition &= String.Format("AND A.NA_CODE LIKE '{0}%'", strCODE_NAME)
                    End If
                ElseIf IsSBCS(strCODE_NAME) Then    '������ ���
                    strCondition &= String.Format("AND (A.NA_CODE LIKE '{0}%' OR A.NA_NAME LIKE '{0}%')", strCODE_NAME)
                Else                                 '�ѱ��� ���
                    strCondition &= String.Format("AND A.NA_NAME LIKE '{0}%'", strCODE_NAME)
                End If
            End If

            '3.������� ���� �˻� (EDATE ������� ����??)
            If blnUseOnly Then
                strChkDate = "AND A.USE_YN='Y'"
                'Else
                '    strChkDate = String.Format("AND (A.USE_YN='Y' OR A.EDATE>={0})", BuildToDate(strUseDate))
            End If


            '�ӽ� SQL�� ����
            strFormat = "SELECT {0} FROM SC_NA_V A " & _
                        "WHERE A.SC_BU_CODE='{1}' {2} {3} " & _
                        "ORDER BY A.NA_CODE"

            'SQL�� ����
            strSQL = String.Format(strFormat, strSelFields, .USRCompany, strCondition, strChkDate)

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetNaAll")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#End Region

    '*****************************************************************
    '�Է� : strInfoXML = ����⺻���� XML
    '       intRowCnt,intColCnt = ��ȸ �Ǽ�,�ʵ� ��
    '       strNA_LEVEL = NA_LEVEL 
    '       blnUseOnly = ���� ������� �͸� �Ǵ� ��ü
    '       strUseDate = blnUseOnly�� False�϶� EDATE>=strUseDate�� �˻�
    '       strAddFields = �ڵ�/�� �̿��� ��ȸ �߰� �ʵ�
    '       blnLikeCode = like �� ����� ���ΰ�? (Default True)
    '��ȯ : ó�����
    '��� : NALEVEL �� ��ȸ
    '*****************************************************************
#Region "27-2. NA_LEVEL ��ȸ"

    Public Function GetNaLevel(ByVal strInfoXML As String, _
                                ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                Optional ByVal strAddFields As String = "", _
                                Optional ByVal blnUseOnly As Boolean = True, _
                                Optional ByVal strUseDate As String = "", _
                                Optional ByVal blnLikeCode As Boolean = True) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strCondition As String      '������
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strChkDate As String = ""   '��뿩�� �� ��볯¥
        Dim vntData As Object


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            '1.SQL 
            strFormat = "SELECT DISTINCT NA_LEVEL As NA_LEVEL FROM SC_NA_V WHERE SC_BU_CODE = '{0}'"

            strSQL = String.Format(strFormat, .USRCompany)

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetNaLevel")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#End Region

    '*****************************************************************
    '�Է� : strInfoXML = ����⺻���� XML
    '       intRowCnt,intColCnt = ��ȸ �Ǽ�,�ʵ� ��
    '       strCODE_NAME = A.NA_CODE, A.NA_NAME :: A.PC = 'P' ����
    '       strNA_LEVEL = NA_LEVEL 
    '       strNA_TYPE_CODE = NA_TYPE 
    '       blnUseOnly = ���� ������� �͸� �Ǵ� ��ü
    '       strUseDate = blnUseOnly�� False�϶� EDATE>=strUseDate�� �˻�
    '       strAddFields = �ڵ�/�� �̿��� ��ȸ �߰� �ʵ�
    '       blnLikeCode = like �� ����� ���ΰ�? (Default True)
    '��ȯ : ó�����
    '��� : NA ��ȸ  (A.NA_CODE, A.NA_NAME) �� ��ȸ
    '*****************************************************************
#Region "27-2. ���� �����׸� ���� MGMT_CODE"

    Public Function GetMGMT(ByVal strInfoXML As String, _
                                ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                Optional ByVal strCODE_NAME As String = "", _
                                Optional ByVal strMGMT_TYPE As String = "", _
                                Optional ByVal strAddFields As String = "", _
                                Optional ByVal blnUseOnly As Boolean = True, _
                                Optional ByVal strUseDate As String = "", _
                                Optional ByVal blnLikeCode As Boolean = True) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strCondition As String      '������
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strChkDate As String = ""   '��뿩�� �� ��볯¥
        Dim vntData As Object


        'SELECT A.MGMT_CODE,A.MGMT_NAME
        '  FROM  SC_NA_MGMT       A
        ' WHERE  A.SC_BU_CODE      = 'H-PHARM' <PARAMETER>
        '   AND  A.MGMT_TYPE       = 'C'       <PARAMETER>
        '   AND  A.USE_YB          = 'Y'
        ' ORDER BY A.MGMT_CODE

        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            '1.��ȸ�ʵ� ����
            If strAddFields <> "" Then strAddFields = "," & AddAlias(strAddFields, "A")
            strSelFields = "A.MGMT_CODE, A.MGMT_NAME " & strAddFields

            '2.������ ����
            If strMGMT_TYPE <> "" Then strCondition = String.Format("AND A.MGMT_TYPE='{0}' ", strMGMT_TYPE)

            If strCODE_NAME <> "" Then
                '��ȸ ������ �ڵ����� �ڵ������ �����Ͽ� ���� �ʵ� ����
                If IsNumeric(strCODE_NAME) Then '������ ���
                    If Not blnLikeCode Then
                        strCondition &= String.Format("AND A.MGMT_CODE={0}", strCODE_NAME)
                    Else
                        strCondition &= String.Format("AND A.MGMT_CODE LIKE '{0}%'", strCODE_NAME)
                    End If
                ElseIf IsSBCS(strCODE_NAME) Then    '������ ���
                    strCondition &= String.Format("AND (A.MGMT_CODE LIKE '{0}%' OR A.MGMT_NAME LIKE '{0}%')", strCODE_NAME)
                Else                                 '�ѱ��� ���
                    strCondition &= String.Format("AND A.MGMT_NAME LIKE '{0}%'", strCODE_NAME)
                End If
            End If


            '3.������� ���� �˻� (EDATE ������� ����??)
            If blnUseOnly Then
                strChkDate = "AND A.USE_YN='Y'"
                'Else
                '    strChkDate = String.Format("AND (A.USE_YN='Y' OR A.EDATE>={0})", BuildToDate(strUseDate))
            End If


            '�ӽ� SQL�� ����
            strFormat = "SELECT {0} FROM SC_NA_MGMT_V A " & _
                        "WHERE A.SC_BU_CODE='{1}' {2} {3} " & _
                        "ORDER BY A.MGMT_CODE"

            'SQL�� ����
            strSQL = String.Format(strFormat, strSelFields, .USRCompany, strCondition, strChkDate)

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetMGMT")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#End Region
    '*****************************************************************
    '�Է� : strInfoXML = ����⺻���� XML
    '       intRowCnt,intColCnt = ��ȸ �Ǽ�,�ʵ� ��
    '       strCODE_NAME = A.COUNTRY_CODE, A.COUNTRY_NAME_KOR
    '       blnUseOnly = ���� ������� �͸� �Ǵ� ��ü
    '       strUseDate = blnUseOnly�� False�϶� EDATE>=strUseDate�� �˻�
    '       strAddFields = �ڵ�/�� �̿��� ��ȸ �߰� �ʵ�
    '       blnLikeCode = like �� ����� ���ΰ�? (Default True)
    '��ȯ : ó�����
    '��� : COUNTRY  (A.COUNTRY_CODE, A.COUNTRY_NAME_KOR) �� ��ȸ
    '*****************************************************************
#Region "28. COUNTRY ��ȸ"

    Public Function GetCountry(ByVal strInfoXML As String, _
                                ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                Optional ByVal strCODE_NAME As String = "", _
                                Optional ByVal strAddFields As String = "", _
                                Optional ByVal blnUseOnly As Boolean = True, _
                                Optional ByVal strUseDate As String = "", _
                                Optional ByVal blnLikeCode As Boolean = True) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strCondition As String      '������
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strChkDate As String = ""   '��뿩�� �� ��볯¥
        Dim vntData As Object


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            '1.��ȸ�ʵ� ����
            If strAddFields <> "" Then strAddFields = "," & AddAlias(strAddFields, "A")
            strSelFields = "A.COUNTRY_CODE, A.COUNTRY_NAME_KOR " & strAddFields

            '2.������ ����
            If strCODE_NAME <> "" Then
                '��ȸ ������ �ڵ����� �ڵ������ �����Ͽ� ���� �ʵ� ����
                If IsNumeric(strCODE_NAME) Then '������ ���
                    If Not blnLikeCode Then
                        strCondition = String.Format("WHERE A.COUNTRY_CODE={0}", strCODE_NAME)
                    Else
                        strCondition = String.Format("WHERE A.COUNTRY_CODE LIKE '{0}%'", strCODE_NAME)
                    End If
                ElseIf IsSBCS(strCODE_NAME) Then    '������ ���
                    strCondition = String.Format("WHERE (A.COUNTRY_CODE LIKE '{0}%' OR A.COUNTRY_NAME_KOR LIKE '{0}%')", strCODE_NAME)
                Else                                 '�ѱ��� ���
                    strCondition = String.Format("WHERE A.COUNTRY_NAME_KOR LIKE '{0}%'", strCODE_NAME)
                End If
            End If


            '3.������� ���� �˻� (EDATE ������� ����??)
            '    If blnUseOnly Then
            '         strChkDate = "AND A.USE_YN='Y'"
            '    Else
            '        strChkDate = String.Format("AND (A.USE_YN='Y' OR A.EDATE>={0})", BuildToDate(strUseDate))
            '   End If


            '�ӽ� SQL�� ����
            strFormat = "SELECT {0} FROM SC_COUNTRY_V A " & _
                        "{1} {2} " & _
                        "ORDER BY A.COUNTRY_CODE"

            'SQL�� ����
            strSQL = String.Format(strFormat, strSelFields, strCondition, strChkDate)

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetCountry")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#End Region

    '*****************************************************************
    '�Է� : strInfoXML = ����⺻���� XML
    '       intRowCnt,intColCnt = ��ȸ �Ǽ�,�ʵ� ��
    '       strCODE_NAME = A.ITEM_TYPE_CODE, A.ITEM_TYPE_NAME
    '       blnUseOnly = ���� ������� �͸� �Ǵ� ��ü
    '       strUseDate = blnUseOnly�� False�϶� EDATE>=strUseDate�� �˻�
    '       strAddFields = �ڵ�/�� �̿��� ��ȸ �߰� �ʵ�
    '       blnLikeCode = like �� ����� ���ΰ�? (Default True)
    '��ȯ : ó�����
    '��� : ITEM_TYPE  (A.ITEM_TYPE_CODE, A.ITEM_TYPE_NAME) �� ��ȸ
    '*****************************************************************
#Region "29. ITEM TYPE ��ȸ"

    Public Function GetItemType(ByVal strInfoXML As String, _
                                ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                Optional ByVal strCODE_NAME As String = "", _
                                Optional ByVal strAddFields As String = "", _
                                Optional ByVal blnUseOnly As Boolean = True, _
                                Optional ByVal strUseDate As String = "", _
                                Optional ByVal blnLikeCode As Boolean = True) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strCondition As String      '������
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strChkDate As String = ""   '��뿩�� �� ��볯¥
        Dim vntData As Object


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            '1.��ȸ�ʵ� ����
            If strAddFields <> "" Then strAddFields = "," & AddAlias(strAddFields, "A")
            strSelFields = "A.ITEM_TYPE_CODE, A.ITEM_TYPE_NAME " & strAddFields

            '2.������ ����
            If strCODE_NAME <> "" Then
                '��ȸ ������ �ڵ����� �ڵ������ �����Ͽ� ���� �ʵ� ����
                If IsNumeric(strCODE_NAME) Then '������ ���
                    If Not blnLikeCode Then
                        strCondition = String.Format("AND A.ITEM_TYPE_CODE={0}", strCODE_NAME)
                    Else
                        strCondition = String.Format("AND A.ITEM_TYPE_CODE LIKE '{0}%'", strCODE_NAME)
                    End If
                ElseIf IsSBCS(strCODE_NAME) Then    '������ ���
                    strCondition = String.Format("AND (A.ITEM_TYPE_CODE LIKE '{0}%' OR A.ITEM_TYPE_NAME LIKE '{0}%')", strCODE_NAME)
                Else                                 '�ѱ��� ���
                    strCondition = String.Format("AND A.ITEM_TYPE_NAME LIKE '{0}%'", strCODE_NAME)
                End If
            End If


            '3.������� ���� �˻� (EDATE ������� ����??)
            If blnUseOnly Then
                strChkDate = "AND A.USE_YN='Y'"
            Else
                '        strChkDate = String.Format("AND (A.USE_YN='Y' OR A.EDATE>={0})", BuildToDate(strUseDate))
            End If


            '�ӽ� SQL�� ����
            strFormat = "SELECT {0} FROM SC_ITEM_TYPE_V A " & _
                        "WHERE A.SC_BU_CODE='{1}' {2} {3} " & _
                        "ORDER BY A.ITEM_TYPE_CODE"

            'SQL�� ����
            strSQL = String.Format(strFormat, strSelFields, .USRCompany, strCondition, strChkDate)

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetItemType")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#End Region

    '*****************************************************************
    '�Է� : strInfoXML = ����⺻���� XML
    '       intRowCnt,intColCnt = ��ȸ �Ǽ�,�ʵ� ��
    '       strCODE_NAME = A.TAX_CODE, A.TAX_NAME
    '       blnUseOnly = ���� ������� �͸� �Ǵ� ��ü
    '       strUseDate = blnUseOnly�� False�϶� EDATE>=strUseDate�� �˻�
    '       strAddFields = �ڵ�/�� �̿��� ��ȸ �߰� �ʵ�
    '       blnLikeCode = like �� ����� ���ΰ�? (Default True)
    '��ȯ : ó�����
    '��� : TAXCODE  (A.TAX_CODE, A.TAX_NAME, A.TAX_RATE, A.WON_PROC) �� ��ȸ
    '*****************************************************************
#Region "30. TAX CODE ��ȸ"

    Public Function GetTaxCode(ByVal strInfoXML As String, _
                                ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                Optional ByVal strCODE_NAME As String = "", _
                                Optional ByVal strTAX_CLASS As String = "", _
                                Optional ByVal strAddFields As String = "", _
                                Optional ByVal blnUseOnly As Boolean = True, _
                                Optional ByVal strUseDate As String = "", _
                                Optional ByVal blnLikeCode As Boolean = True) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strCondition As String      '������
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strChkDate As String = ""   '��뿩�� �� ��볯¥
        Dim vntData As Object


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            '1.��ȸ�ʵ� ����
            If strAddFields <> "" Then strAddFields = "," & AddAlias(strAddFields, "A")
            strSelFields = "A.TAX_CODE, A.TAX_NAME, A.TAX_RATE, A.WON_PROC " & strAddFields

            '2.������ ����

            If strTAX_CLASS <> "" Then strCondition &= String.Format("AND A.TAX_CLASS='{0}' ", strTAX_CLASS)

            If strCODE_NAME <> "" Then
                '��ȸ ������ �ڵ����� �ڵ������ �����Ͽ� ���� �ʵ� ����
                If IsNumeric(strCODE_NAME) Then '������ ���
                    If Not blnLikeCode Then
                        strCondition &= String.Format("AND A.TAX_CODE={0}", strCODE_NAME)
                    Else
                        strCondition &= String.Format("AND A.TAX_CODE LIKE '{0}%'", strCODE_NAME)
                    End If
                ElseIf IsSBCS(strCODE_NAME) Then    '������ ���
                    strCondition &= String.Format("AND (A.TAX_CODE LIKE '{0}%' OR A.TAX_NAME LIKE '{0}%')", strCODE_NAME)
                Else                                 '�ѱ��� ���
                    strCondition &= String.Format("AND (A.TAX_CODE LIKE '{0}%' OR A.TAX_NAME LIKE '{0}%')", strCODE_NAME)
                End If
            End If


            '3.������� ���� �˻� (EDATE ������� ����??)
            'If blnUseOnly Then
            '    strChkDate = "AND A.USE_YN='Y'"
            'Else
            '        strChkDate = String.Format("AND (A.USE_YN='Y' OR A.EDATE>={0})", BuildToDate(strUseDate))
            'End If


            '�ӽ� SQL�� ����
            strFormat = "SELECT {0} FROM SC_TAX_CODE_V A " & _
                        "WHERE A.SC_BU_CODE='{1}' {2} {3} " & _
                        "ORDER BY A.SEQ "

            'SQL�� ����
            strSQL = String.Format(strFormat, strSelFields, .USRCompany, strCondition, strChkDate)

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetTaxCode")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#End Region

    '*****************************************************************
    '�Է� : strInfoXML = ����⺻���� XML
    '       intRowCnt,intColCnt = ��ȸ �Ǽ�,�ʵ� ��
    '       strSYSID = ���� ������ �ѱ�� �ý��� ������
    '       strLINEKEY = ���� ������ �ѱ�� ����Ű
    '��ȯ : ó�����
    '��� : ���� ������ ��ȸ
    '*****************************************************************
#Region "31. ���� ������ ��ȸ"

    Public Function GetApprovalList(ByVal strInfoXML As String, _
                                ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                ByVal strSYSID As String, _
                                ByVal strLINEKEY As String) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strCondition As String      '������
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strChkDate As String = ""   '��뿩�� �� ��볯¥
        Dim vntData As Object


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            strFormat = "SELECT TITLE, DRAFTEMPNO, DRAFTNAME, EMPNO, EMP_NAME, STATE, APPDATE FROM APPRO_HISTORY_V WHERE  SYSID = '{0}' AND LINEKEY ='{1}' ORDER BY SORT"
            strSQL = String.Format(strFormat, strSYSID, strLINEKEY)
            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetApprovalList")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "32. ���ڵ�"
    Public Function GetVessel(ByVal strInfoXML As String, _
                                ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                Optional ByVal strCODE_NAME As String = "", _
                                Optional ByVal blnLikeCode As Boolean = True) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strCondition As String      '������
        Dim vntData As Object

        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            If strCODE_NAME <> "" Then
                strCondition &= String.Format("AND (A.VESSELNO LIKE '{0}%' OR A.VESL_NM LIKE '{0}%')", strCODE_NAME)
            End If

            '�ӽ� SQL�� ����
            strFormat = "SELECT A.VESSELNO, A.VESL_NM, A.ETA FROM AP_VESSEL_V A " & _
                        "WHERE  1=1 {0} " & _
                        "ORDER BY A.VESSELNO "

            'SQL�� ����
            strSQL = String.Format(strFormat, strCondition)

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetVessel")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "33. �����׸� ��ȸ"
    Public Function GetSC_ITEMCODE_MST(ByVal strInfoXML As String, _
                                       ByRef intRowCnt As Integer, _
                                       ByRef intColCnt As Integer, _
                                       ByVal strCUSTCODE As String, _
                                       ByVal strCUSTNAME As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2 As String
        Dim vntData As Object

        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""

                If strCUSTCODE <> "" Then Con1 = String.Format(" AND (ITEMCODE LIKE '%{0}%')", strCUSTCODE)
                If strCUSTNAME <> "" Then Con2 = String.Format(" AND (ITEMNAME LIKE '%{0}%')", strCUSTNAME)
                strWhere = BuildFields(" ", Con1, Con2)

                strSelFields = "itemcode,PD_ITEMCODE_NAME_FUN(substrING(itemcode,1,3)) as classcode, itemname"
                strFormet = "select {0} from sc_itemcode_mst where 1=1 {1} "


                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetSC_ITEMCODE_MST")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "34. ����SEQ�˾���ȸ"
    Public Function GetJOBSEQList(ByVal strInfoXML As String, _
                                ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                ByVal strCODE As String, _
                                ByVal strCODENAME As String) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strCondition As String      '������
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strChkDate As String = ""   '��뿩�� �� ��볯¥
        Dim strWhere As String
        Dim vntData As Object
        Dim Con1 As String
        Dim Con2 As String


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            Con1 = ""
            Con2 = ""
            If strCODE <> "" Then Con1 = String.Format(" AND (SEQNO = '{0}')", strCODE)
            If strCODENAME <> "" Then Con2 = String.Format(" AND (SEQNAME like '%{0}%')", strCODENAME)

            strWhere = BuildFields(" ", Con1, Con2)

            strFormat = "SELECT SEQNO,SEQNAME FROM SC_JOBCUST WHERE  1=1 {0} ORDER BY SEQNO"
            strSQL = String.Format(strFormat, strWhere)
            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetJOBSEQList")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "34. �������� ���μ� �˾���ȸ"
    Public Function GetDEPT_CDBYCUSTSEQList(ByVal strInfoXML As String, _
                                            ByRef intRowCnt As Integer, _
                                            ByRef intColCnt As Integer, _
                                            ByVal strCODE As String, _
                                            ByVal strCODENAME As String, _
                                            ByVal strCUSTCODE As String, _
                                            ByVal strCUSTNAME As String) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strCondition As String      '������
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strChkDate As String = ""   '��뿩�� �� ��볯¥
        Dim strWhere As String
        Dim vntData As Object
        Dim Con1 As String
        Dim Con2 As String
        Dim Con3 As String
        Dim Con4 As String

        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            Con1 = ""
            Con2 = ""
            Con3 = ""
            Con4 = ""
            If strCODE <> "" Then Con1 = String.Format(" AND (SEQNO = '{0}')", strCODE)
            If strCODENAME <> "" Then Con2 = String.Format(" AND (SEQNAME like '%{0}%')", strCODENAME)
            If strCUSTCODE <> "" Then Con3 = String.Format(" AND (CUSTCODE = '{0}')", strCUSTCODE)
            If strCUSTNAME <> "" Then Con4 = String.Format(" AND (DBO.MD_GET_CUSTNAME_FUN(CUSTCODE) like '%{0}%')", strCUSTNAME)

            strWhere = BuildFields(" ", Con1, Con2, Con3, Con4)
            strFormat = " SELECT CUSTCODE+'-'+SEQNO CODE,SEQNO,SEQNAME,CUSTCODE,DBO.MD_GET_CUSTNAME_FUN(CUSTCODE) CUSTNAME,DEPTCD DEPT_CD,DBO.SC_DEPT_NAME_FUN(DEPTCD) DEPTNAME, CLIENTSUBCODE, DBO.MD_GET_CUSTNAME_FUN(CLIENTSUBCODE) CLIENTSUBNAME "
            strFormat = strFormat & " FROM SC_JOBCUST  "
            strFormat = strFormat & " WHERE  1=1 {0} "
            strFormat = strFormat & " ORDER BY  "
            strFormat = strFormat & " CASE SUBSTRING(LTRIM(SEQNAME),1,3) WHEN '(��)' THEN LTRIM(SUBSTRING(SEQNAME,4,LEN(SEQNAME))) "
            strFormat = strFormat & " WHEN '(��)' THEN LTRIM(SUBSTRING(SEQNAME,4,LEN(SEQNAME))) "
            strFormat = strFormat & " WHEN '(��)' THEN LTRIM(SUBSTRING(SEQNAME,4,LEN(SEQNAME))) "
            strFormat = strFormat & " WHEN '(��)' THEN LTRIM(SUBSTRING(SEQNAME,4,LEN(SEQNAME))) "
            strFormat = strFormat & " WHEN '(���' THEN LTRIM(SUBSTRING(SEQNAME,5,LEN(SEQNAME))) "
            strFormat = strFormat & " WHEN '(��)' THEN LTRIM(SUBSTRING(SEQNAME,4,LEN(SEQNAME))) "
            strFormat = strFormat & " ELSE LTRIM(SEQNAME) END "
            strSQL = String.Format(strFormat, strWhere)
            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetJOBSEQList")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "35. �μ��ü ����Ź �ŷ��� ��ȸ"
    Public Function Get_PRINTTRANS_HDR(ByVal strInfoXML As String, _
                                       ByRef intRowCnt As Integer, _
                                       ByRef intColCnt As Integer, _
                                       ByVal strYEARMON As String, _
                                       ByVal strTRANSNO As String, _
                                       ByVal strCLIENTCODE As String) As String

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strWhere As String
        Dim strXMLData As String
        Dim Con1 As String
        Dim Con2 As String
        Dim Con3 As String

        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            Con1 = ""
            Con2 = ""
            Con3 = ""
            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
            If strTRANSNO <> "" Then Con2 = String.Format(" AND (TRANSNO = '{0}')", strTRANSNO)
            If strCLIENTCODE <> "" Then Con3 = String.Format(" AND (CLIENTCODE like '%{0}%')", strCLIENTCODE)

            strWhere = BuildFields(" ", Con1, Con2, Con3)
            strFormat = "SELECT TRANSYEARMON, TRANSNO, DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, DBO.SC_DEPT_NAME_FUN(DEPT_CD) DEPT_NAME, DEMANDDAY, PRINTDAY, AMT, VAT , (AMT + VAT) SUMAMTVAT FROM MD_PRINTTRANS_HDR WHERE 1=1 {0} ORDER BY MED_FLAG "
            strSQL = String.Format(strFormat, strWhere)
            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                strXMLData = .mobjSCGLSql.SQLSelectXml(strSQL, intRowCnt, intColCnt)
                Return strXMLData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".Get_PRINTTRANS_HDR")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    Public Function Get_PRINTTRANS_LIST(ByVal strInfoXML As String, _
                                        ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                        ByVal strYEARMON As String, _
                                        ByVal strTRANSNO As String, _
                                        ByVal strCLIENTCODE As String) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strWhere As String
        Dim vntData As Object
        Dim Con1 As String
        Dim Con2 As String
        Dim Con3 As String


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            Con1 = ""
            Con2 = ""
            Con3 = ""
            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
            If strTRANSNO <> "" Then Con2 = String.Format(" AND (TRANSNO = '{0}')", strTRANSNO)
            If strCLIENTCODE <> "" Then Con3 = String.Format(" AND (CLIENTCODE like '%{0}%')", strCLIENTCODE)

            strWhere = BuildFields(" ", Con1, Con2, Con3)
            strFormat = "SELECT TRANSYEARMON,  TRANSNO, SEQ, CLIENTCODE, DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, TRUST_SEQ, MEDCODE, DBO.MD_GET_CUSTNAME_FUN(MEDCODE) MEDNAME,REAL_MED_CODE,DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) REAL_MED_NAME,DBO.PD_JOBCUST_NAME_FUN(SUBSEQ) SUBSEQNAME, DBO.MD_GET_CUSTNAME_FUN(CLIENTSUBCODE) CLIENTSUBNAME,DEPT_CD, dbo.SC_DEPT_NAME_FUN(DEPT_CD) DEPT_NAME, DEMANDDAY, PRINTDAY, PROGRAM_NAME, STD_STEP, STD_CM, COL_DEG, PUB_DATE, AMT, TRU_TAX_FLAG, VAT,SUMAMTVAT, MED_FLAG, DBO.MD_GET_MEDNAME_FUN(MED_FLAG) MED_FLAGNAME, MEMO, SPONSOR, TAXYEARMON, TAXNO, CONFIRMFLAG FROM MD_PRINTTRANS_DTL WHERE 1=1 {0} ORDER BY MED_FLAG "
            strSQL = String.Format(strFormat, strWhere)
            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".Get_PRINTTRANS_LIST")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "35. �μ��ü ������ �ŷ��� ��ȸ"
    Public Function Get_PRINTCOMMI_HDR(ByVal strInfoXML As String, _
                                       ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                       ByVal strYEARMON As String, _
                                       ByVal strTRANSNO As String, _
                                       ByVal strREAL_MED_CODE As String) As String

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strWhere As String
        Dim strXMLData As String
        Dim Con1 As String
        Dim Con2 As String
        Dim Con3 As String


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            Con1 = ""
            Con2 = ""
            Con3 = ""
            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
            If strTRANSNO <> "" Then Con2 = String.Format(" AND (TRANSNO = '{0}')", strTRANSNO)
            If strREAL_MED_CODE <> "" Then Con3 = String.Format(" AND (REAL_MED_CODE like '%{0}%')", strREAL_MED_CODE)

            strWhere = BuildFields(" ", Con1, Con2, Con3)
            strFormat = "SELECT TRANSYEARMON, TRANSNO, DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, DBO.SC_DEPT_NAME_FUN(DEPT_CD) DEPT_NAME, DEMANDDAY, PRINTDAY, AMT, VAT , (AMT + VAT) SUMAMTVAT  FROM MD_PRINTCOMMI_HDR WHERE 1=1 {0} ORDER BY MED_FLAG "
            strSQL = String.Format(strFormat, strWhere)
            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                strXMLData = .mobjSCGLSql.SQLSelectXml(strSQL, intRowCnt, intColCnt)
                Return strXMLData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetJOBSEQList")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    Public Function Get_PRINTCOMMI_LIST(ByVal strInfoXML As String, _
                                       ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                       ByVal strYEARMON As String, _
                                       ByVal strTRANSNO As String, _
                                       ByVal strREAL_MED_CODE As String) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strWhere As String
        Dim vntData As Object
        Dim Con1 As String
        Dim Con2 As String
        Dim Con3 As String


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            Con1 = ""
            Con2 = ""
            Con3 = ""
            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
            If strTRANSNO <> "" Then Con2 = String.Format(" AND (TRANSNO = '{0}')", strTRANSNO)
            If strREAL_MED_CODE <> "" Then Con3 = String.Format(" AND (REAL_MED_CODE like '%{0}%')", strREAL_MED_CODE)

            strWhere = BuildFields(" ", Con1, Con2, Con3)
            strFormat = "SELECT TRANSYEARMON,  TRANSNO, SEQ, CLIENTCODE, DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, MEDCODE, DBO.MD_GET_CUSTNAME_FUN(MEDCODE) MEDNAME,REAL_MED_CODE,DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) REAL_MED_NAME,DEPT_CD, dbo.SC_DEPT_NAME_FUN(DEPT_CD) DEPT_NAME, DEMANDDAY, PRINTDAY, AMT, SUSURATE, SUSU, VAT, DBO.MD_GET_MEDNAME_FUN(MED_FLAG) MED_FLAG, TAXYEARMON, TAXNO, TRUST_SEQ FROM MD_PRINTCOMMI_DTL WHERE 1=1 {0} ORDER BY MED_FLAG "
            strSQL = String.Format(strFormat, strWhere)
            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetJOBSEQList")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "35. CATV ����Ź �ŷ��� ��ȸ"
    Public Function Get_CATVTRANS_HDR(ByVal strInfoXML As String, _
                                       ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                       ByVal strYEARMON As String, _
                                       ByVal strTRANSNO As String, _
                                       ByVal strCLIENTCODE As String) As String

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strWhere As String
        Dim strXMLData As String
        Dim Con1 As String
        Dim Con2 As String
        Dim Con3 As String

        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            Con1 = ""
            Con2 = ""
            Con3 = ""
            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
            If strTRANSNO <> "" Then Con2 = String.Format(" AND (TRANSNO = '{0}')", strTRANSNO)
            If strCLIENTCODE <> "" Then Con3 = String.Format(" AND (CLIENTCODE like '%{0}%')", strCLIENTCODE)

            strWhere = BuildFields(" ", Con1, Con2, Con3)
            strFormat = "SELECT TRANSYEARMON, TRANSNO, dbo.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, dbo.SC_DEPT_NAME_FUN(DEPT_CD) DEPT_NAME, DEMANDDAY, PRINTDAY, AMT, VAT , (AMT + VAT) SUMAMTVAT FROM MD_CATVTRANS_HDR WHERE 1=1 {0} ORDER BY DEMANDDAY "
            strSQL = String.Format(strFormat, strWhere)

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                strXMLData = .mobjSCGLSql.SQLSelectXml(strSQL, intRowCnt, intColCnt)
                Return strXMLData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetJOBSEQList")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    Public Function Get_CATVTRANS_LIST(ByVal strInfoXML As String, _
                                       ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                       ByVal strYEARMON As String, _
                                       ByVal strTRANSNO As String, _
                                       ByVal strCLIENTCODE As String) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strWhere As String
        Dim vntData As Object
        Dim Con1 As String
        Dim Con2 As String
        Dim Con3 As String


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            Con1 = ""
            Con2 = ""
            Con3 = ""
            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
            If strTRANSNO <> "" Then Con2 = String.Format(" AND (TRANSNO = '{0}')", strTRANSNO)
            If strCLIENTCODE <> "" Then Con3 = String.Format(" AND (CLIENTCODE like '%{0}%')", strCLIENTCODE)

            strWhere = BuildFields(" ", Con1, Con2, Con3)
            strFormat = "SELECT TRANSYEARMON,  TRANSNO, SEQ, CLIENTCODE, dbo.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, MEDCODE, dbo.MD_GET_CUSTNAME_FUN(MEDCODE) MEDNAME,REAL_MED_CODE,dbo.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) REAL_MED_NAME,DBO.MD_GET_CUSTNAME_FUN(MPP) MPP_NAME, DEPT_CD, DEMANDDAY, PRINTDAY, PROGRAM, PROGNAME,TBRDSTDATE, TBRDEDDATE, CNT, AMT,TRU_TAX_FLAG,VAT,AMT+VAT SUMAMTVAT, TRUST_SEQ, MEMO,MED_FLAG,SPONSOR, TAXYEARMON, TAXNO, CONFIRMFLAG FROM MD_CATVTRANS_DTL WHERE 1=1 {0} ORDER BY DEMANDDAY "
            strSQL = String.Format(strFormat, strWhere)

            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetJOBSEQList")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "35. CATV ������ �ŷ��� ��ȸ"
    Public Function Get_CATVCOMMI_HDR(ByVal strInfoXML As String, _
                                       ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                       ByVal strYEARMON As String, _
                                       ByVal strTRANSNO As String, _
                                       ByVal strREAL_MED_CODE As String) As String

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strWhere As String
        Dim strXMLData As String
        Dim Con1 As String
        Dim Con2 As String
        Dim Con3 As String


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            Con1 = ""
            Con2 = ""
            Con3 = ""
            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
            If strTRANSNO <> "" Then Con2 = String.Format(" AND (TRANSNO = '{0}')", strTRANSNO)
            If strREAL_MED_CODE <> "" Then Con3 = String.Format(" AND (REAL_MED_CODE like '%{0}%')", strREAL_MED_CODE)

            strWhere = BuildFields(" ", Con1, Con2, Con3)
            strFormat = "SELECT TRANSYEARMON, TRANSNO, dbo.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, dbo.SC_DEPT_NAME_FUN(DEPT_CD) DEPT_NAME, DEMANDDAY, PRINTDAY, AMT, VAT, (AMT + VAT) SUMAMTVAT  FROM MD_CATVCOMMI_HDR WHERE 1=1 {0} ORDER BY MED_FLAG "
            strSQL = String.Format(strFormat, strWhere)
            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                strXMLData = .mobjSCGLSql.SQLSelectXml(strSQL, intRowCnt, intColCnt)
                Return strXMLData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetJOBSEQList")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    Public Function Get_CATVCOMMI_LIST(ByVal strInfoXML As String, _
                                       ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                       ByVal strYEARMON As String, _
                                       ByVal strTRANSNO As String, _
                                       ByVal strREAL_MED_CODE As String) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strWhere As String
        Dim vntData As Object
        Dim Con1 As String
        Dim Con2 As String
        Dim Con3 As String


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            Con1 = ""
            Con2 = ""
            Con3 = ""
            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
            If strTRANSNO <> "" Then Con2 = String.Format(" AND (TRANSNO = '{0}')", strTRANSNO)
            If strREAL_MED_CODE <> "" Then Con3 = String.Format(" AND (REAL_MED_CODE like '%{0}%')", strREAL_MED_CODE)

            strWhere = BuildFields(" ", Con1, Con2, Con3)
            strFormat = "SELECT TRANSYEARMON,  TRANSNO, SEQ, CLIENTCODE, dbo.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, MEDCODE, dbo.MD_GET_CUSTNAME_FUN(MEDCODE) MEDNAME,REAL_MED_CODE,dbo.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) REAL_MED_NAME,DEPT_CD, DEMANDDAY, PRINTDAY, AMT, SUSURATE, SUSU, VAT, MEMO, MED_FLAG, SPONSOR, TAXYEARMON, TAXNO, TRUST_SEQ FROM MD_CATVCOMMI_DTL WHERE 1=1 {0} ORDER BY REAL_MED_CODE "
            strSQL = String.Format(strFormat, strWhere)
            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetJOBSEQList")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "35. ���ͳ� ����Ź �ŷ��� ��ȸ"
    Public Function Get_INTERNETTRANS_HDR(ByVal strInfoXML As String, _
                                          ByRef intRowCnt As Integer, _
                                          ByRef intColCnt As Integer, _
                                          ByRef strYEARMON As String, _
                                          ByRef strTRANSNO As String, _
                                          ByRef strCLIENTCODE As String) As String

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strWhere As String
        Dim strXMLData As String
        Dim Con1 As String
        Dim Con2 As String
        Dim Con3 As String

        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            Con1 = ""
            Con2 = ""
            Con3 = ""
            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
            If strTRANSNO <> "" Then Con2 = String.Format(" AND (TRANSNO = '{0}')", strTRANSNO)
            If strCLIENTCODE <> "" Then Con3 = String.Format(" AND (CLIENTCODE like '%{0}%')", strCLIENTCODE)

            strWhere = BuildFields(" ", Con1, Con2, Con3)
            strFormat = "SELECT TRANSYEARMON, TRANSNO, dbo.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, dbo.SC_DEPT_NAME_FUN(DEPT_CD) DEPT_NAME, DEMANDDAY, PRINTDAY, AMT, VAT , (AMT + VAT) SUMAMTVAT FROM MD_INTERNETTRANS_HDR WHERE 1=1 {0} ORDER BY MED_FLAG "
            strSQL = String.Format(strFormat, strWhere)
            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                strXMLData = .mobjSCGLSql.SQLSelectXml(strSQL, intRowCnt, intColCnt)
                Return strXMLData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetJOBSEQList")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    Public Function Get_INTERNETTRANS_LIST(ByVal strInfoXML As String, _
                                       ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                       ByRef strYEARMON As String, _
                                       ByRef strTRANSNO As String, _
                                       ByRef strCLIENTCODE As String) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strWhere As String
        Dim vntData As Object
        Dim Con1 As String
        Dim Con2 As String
        Dim Con3 As String


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            Con1 = ""
            Con2 = ""
            Con3 = ""
            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
            If strTRANSNO <> "" Then Con2 = String.Format(" AND (TRANSNO = '{0}')", strTRANSNO)
            If strCLIENTCODE <> "" Then Con3 = String.Format(" AND (CLIENTCODE like '%{0}%')", strCLIENTCODE)

            strWhere = BuildFields(" ", Con1, Con2, Con3)
            strFormat = "SELECT TRANSYEARMON,  TRANSNO,SEQ,  CLIENTCODE, dbo.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, TRUST_SEQ, MEDCODE, dbo.MD_GET_CUSTNAME_FUN(MEDCODE) MEDNAME,dbo.MD_GET_REALMEDCODE_FUN(REAL_MED_LOWCODE) REAL_MED_LOWNAME, REAL_MED_CODE,dbo.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) REAL_MED_NAME,DEPT_CD, DEMANDDAY, PRINTDAY, PROGRAM, TBRDSTDATE, TBRDEDDATE, AMT, TRU_TAX_FLAG, VAT, SUMAMTVAT, MEMO , MED_FLAG, SPONSOR, TAXYEARMON, TAXNO, CONFIRMFLAG FROM MD_INTERNETTRANS_DTL WHERE 1=1 {0} ORDER BY DEMANDDAY "
            strSQL = String.Format(strFormat, strWhere)
            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetJOBSEQList")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "35. ���ͳ� ������ �ŷ��� ��ȸ"
    Public Function Get_INTERNETCOMMI_HDR(ByVal strInfoXML As String, _
                                       ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                       ByVal strYEARMON As String, _
                                       ByVal strTRANSNO As String, _
                                       ByVal strREAL_MED_CODE As String) As String

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strWhere As String
        Dim strXMLData As String
        Dim Con1 As String
        Dim Con2 As String
        Dim Con3 As String


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            Con1 = ""
            Con2 = ""
            Con3 = ""
            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
            If strTRANSNO <> "" Then Con2 = String.Format(" AND (TRANSNO = '{0}')", strTRANSNO)
            If strREAL_MED_CODE <> "" Then Con3 = String.Format(" AND (REAL_MED_CODE like '%{0}%')", strREAL_MED_CODE)

            strWhere = BuildFields(" ", Con1, Con2, Con3)
            strFormat = "SELECT TRANSYEARMON, TRANSNO, dbo.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, dbo.SC_DEPT_NAME_FUN(DEPT_CD) DEPT_NAME, DEMANDDAY, PRINTDAY, AMT, VAT, (AMT + VAT) SUMAMTVAT  FROM MD_INTERNETCOMMI_HDR WHERE 1=1 {0} ORDER BY MED_FLAG "
            strSQL = String.Format(strFormat, strWhere)
            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                strXMLData = .mobjSCGLSql.SQLSelectXml(strSQL, intRowCnt, intColCnt)
                Return strXMLData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetJOBSEQList")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    Public Function Get_INTERNETCOMMI_LIST(ByVal strInfoXML As String, _
                                       ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                       ByVal strYEARMON As String, _
                                       ByVal strTRANSNO As String, _
                                       ByVal strREAL_MED_CODE As String) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strWhere As String
        Dim vntData As Object
        Dim Con1 As String
        Dim Con2 As String
        Dim Con3 As String


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            Con1 = ""
            Con2 = ""
            Con3 = ""
            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
            If strTRANSNO <> "" Then Con2 = String.Format(" AND (TRANSNO = '{0}')", strTRANSNO)
            If strREAL_MED_CODE <> "" Then Con3 = String.Format(" AND (REAL_MED_CODE like '%{0}%')", strREAL_MED_CODE)

            strWhere = BuildFields(" ", Con1, Con2, Con3)
            strFormat = "SELECT TRANSYEARMON,  TRANSNO, CLIENTCODE, dbo.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, MEDCODE, dbo.MD_GET_CUSTNAME_FUN(MEDCODE) MEDNAME,REAL_MED_CODE, dbo.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) REAL_MED_NAME,DEPT_CD, DEMANDDAY, PRINTDAY, AMT, SUSURATE, SUSU, VAT, MEMO, MED_FLAG, SPONSOR FROM MD_INTERNETCOMMI_DTL WHERE 1=1 {0} ORDER BY REAL_MED_CODE "
            strSQL = String.Format(strFormat, strWhere)
            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetJOBSEQList")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "36. ���� ����Ź �ŷ��� ��ȸ"
    Public Function Get_OUTDOORTRANS_HDR(ByVal strInfoXML As String, _
                                         ByRef intRowCnt As Integer, _
                                         ByRef intColCnt As Integer, _
                                         ByRef strYEARMON As String, _
                                         ByRef strTRANSNO As String, _
                                         ByRef strCLIENTCODE As String) As String

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strWhere As String
        Dim strXMLData As String
        Dim Con1 As String
        Dim Con2 As String
        Dim Con3 As String

        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            Con1 = ""
            Con2 = ""
            Con3 = ""
            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
            If strTRANSNO <> "" Then Con2 = String.Format(" AND (TRANSNO = '{0}')", strTRANSNO)
            If strCLIENTCODE <> "" Then Con3 = String.Format(" AND (CLIENTCODE like '%{0}%')", strCLIENTCODE)

            strWhere = BuildFields(" ", Con1, Con2, Con3)
            strFormat = "SELECT TRANSYEARMON, TRANSNO, dbo.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, dbo.SC_DEPT_NAME_FUN(DEPT_CD) DEPT_NAME, DEMANDDAY, PRINTDAY, AMT, VAT , (AMT + VAT) SUMAMTVAT FROM MD_OUTDOORTRANS_HDR WHERE 1=1 {0} ORDER BY MED_FLAG "
            strSQL = String.Format(strFormat, strWhere)
            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                strXMLData = .mobjSCGLSql.SQLSelectXml(strSQL, intRowCnt, intColCnt)
                Return strXMLData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".Get_OUTDOORTRANS_HDR")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    Public Function Get_OUTDOORTRANS_LIST(ByVal strInfoXML As String, _
                                          ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                          ByRef strYEARMON As String, _
                                          ByRef strTRANSNO As String, _
                                          ByRef strCLIENTCODE As String) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strWhere As String
        Dim vntData As Object
        Dim Con1 As String
        Dim Con2 As String
        Dim Con3 As String


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            Con1 = ""
            Con2 = ""
            Con3 = ""
            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
            If strTRANSNO <> "" Then Con2 = String.Format(" AND (TRANSNO = '{0}')", strTRANSNO)
            If strCLIENTCODE <> "" Then Con3 = String.Format(" AND (CLIENTCODE like '%{0}%')", strCLIENTCODE)

            strWhere = BuildFields(" ", Con1, Con2, Con3)
            strFormat = "SELECT TRANSYEARMON,  TRANSNO,SEQ,  CLIENTCODE, "
            strFormat = strFormat & "  dbo.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, "
            strFormat = strFormat & "  TRUST_SEQ, MEDCODE, dbo.MD_GET_CUSTNAME_FUN(MEDCODE) MEDNAME,"
            strFormat = strFormat & "  REAL_MED_CODE,"
            strFormat = strFormat & "  dbo.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) REAL_MED_NAME,"
            strFormat = strFormat & "  DBO.PD_JOBCUST_NAME_FUN(SUBSEQ) SUBSEQNAME,  "
            strFormat = strFormat & "  DBO.MD_GET_CUSTNAME_FUN(CLIENTSUBCODE) CLIENTSUBNAME,  "
            strFormat = strFormat & "  DEPT_CD, DEMANDDAY, PRINTDAY, PROGRAM, "
            strFormat = strFormat & "  TBRDSTDATE, TBRDEDDATE, AMT, TRU_TAX_FLAG, VAT, SUMAMTVAT, "
            strFormat = strFormat & "  MEMO , MED_FLAG, SPONSOR, TAXYEARMON, TAXNO "
            strFormat = strFormat & "  FROM MD_OUTDOORTRANS_DTL "
            strFormat = strFormat & "  WHERE 1=1 {0} ORDER BY DEMANDDAY "
            strSQL = String.Format(strFormat, strWhere)
            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".Get_OUTDOORTRANS_LIST")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "35. ���� ������ �ŷ��� ��ȸ"
    Public Function Get_OUTDOORCOMMI_HDR(ByVal strInfoXML As String, _
                                         ByRef intRowCnt As Integer, _
                                         ByRef intColCnt As Integer, _
                                         ByVal strYEARMON As String, _
                                         ByVal strTRANSNO As String, _
                                         ByVal strREAL_MED_CODE As String) As String

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strWhere As String
        Dim strXMLData As String
        Dim Con1 As String
        Dim Con2 As String
        Dim Con3 As String


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            Con1 = ""
            Con2 = ""
            Con3 = ""
            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
            If strTRANSNO <> "" Then Con2 = String.Format(" AND (TRANSNO = '{0}')", strTRANSNO)
            If strREAL_MED_CODE <> "" Then Con3 = String.Format(" AND (REAL_MED_CODE like '%{0}%')", strREAL_MED_CODE)

            strWhere = BuildFields(" ", Con1, Con2, Con3)
            strFormat = "SELECT TRANSYEARMON, TRANSNO, dbo.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, dbo.SC_DEPT_NAME_FUN(DEPT_CD) DEPT_NAME, DEMANDDAY, PRINTDAY, AMT, VAT, (AMT + VAT) SUMAMTVAT  FROM MD_OUTDOORCOMMI_HDR WHERE 1=1 {0} ORDER BY MED_FLAG "
            strSQL = String.Format(strFormat, strWhere)
            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                strXMLData = .mobjSCGLSql.SQLSelectXml(strSQL, intRowCnt, intColCnt)
                Return strXMLData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".Get_OUTDOORCOMMI_HDR")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    Public Function Get_OUTDOORCOMMI_LIST(ByVal strInfoXML As String, _
                                          ByRef intRowCnt As Integer, _
                                          ByRef intColCnt As Integer, _
                                          ByVal strYEARMON As String, _
                                          ByVal strTRANSNO As String, _
                                          ByVal strREAL_MED_CODE As String) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strWhere As String
        Dim vntData As Object
        Dim Con1 As String
        Dim Con2 As String
        Dim Con3 As String


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            Con1 = ""
            Con2 = ""
            Con3 = ""
            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
            If strTRANSNO <> "" Then Con2 = String.Format(" AND (TRANSNO = '{0}')", strTRANSNO)
            If strREAL_MED_CODE <> "" Then Con3 = String.Format(" AND (REAL_MED_CODE like '%{0}%')", strREAL_MED_CODE)

            strWhere = BuildFields(" ", Con1, Con2, Con3)
            strFormat = "SELECT TRANSYEARMON,  TRANSNO, CLIENTCODE, dbo.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, MEDCODE, dbo.MD_GET_CUSTNAME_FUN(MEDCODE) MEDNAME,REAL_MED_CODE, dbo.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) REAL_MED_NAME,DEPT_CD, DEMANDDAY, PRINTDAY, AMT, SUSURATE, SUSU, VAT, MEMO, MED_FLAG, SPONSOR FROM MD_OUTDOORCOMMI_DTL WHERE 1=1 {0} ORDER BY REAL_MED_CODE "
            strSQL = String.Format(strFormat, strWhere)
            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".Get_OUTDOORCOMMI_LIST")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "35. ���θ�� ����Ź �ŷ��� ��ȸ"
    Public Function Get_PROMOTIONTRANS_HDR(ByVal strInfoXML As String, _
                                           ByRef intRowCnt As Integer, _
                                           ByRef intColCnt As Integer, _
                                           ByRef strYEARMON As String, _
                                           ByRef strTRANSNO As String, _
                                           ByRef strCLIENTCODE As String) As String

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strWhere As String
        Dim strXMLData As String
        Dim Con1 As String
        Dim Con2 As String
        Dim Con3 As String

        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            Con1 = ""
            Con2 = ""
            Con3 = ""
            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
            If strTRANSNO <> "" Then Con2 = String.Format(" AND (TRANSNO = '{0}')", strTRANSNO)
            If strCLIENTCODE <> "" Then Con3 = String.Format(" AND (CLIENTCODE like '%{0}%')", strCLIENTCODE)

            strWhere = BuildFields(" ", Con1, Con2, Con3)
            strFormat = "SELECT TRANSYEARMON, TRANSNO, dbo.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, dbo.SC_DEPT_NAME_FUN(DEPT_CD) DEPT_NAME, DEMANDDAY, PRINTDAY, AMT, VAT , (AMT + VAT) SUMAMTVAT FROM MD_PROMOTIONTRANS_HDR WHERE 1=1 {0} ORDER BY MED_FLAG "
            strSQL = String.Format(strFormat, strWhere)
            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                strXMLData = .mobjSCGLSql.SQLSelectXml(strSQL, intRowCnt, intColCnt)
                Return strXMLData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetJOBSEQList")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    Public Function Get_PROMOTIONTRANS_LIST(ByVal strInfoXML As String, _
                                            ByRef intRowCnt As Integer, _
                                            ByRef intColCnt As Integer, _
                                            ByRef strYEARMON As String, _
                                            ByRef strTRANSNO As String, _
                                            ByRef strCLIENTCODE As String) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strWhere As String
        Dim vntData As Object
        Dim Con1 As String
        Dim Con2 As String
        Dim Con3 As String


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            Con1 = ""
            Con2 = ""
            Con3 = ""
            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
            If strTRANSNO <> "" Then Con2 = String.Format(" AND (TRANSNO = '{0}')", strTRANSNO)
            If strCLIENTCODE <> "" Then Con3 = String.Format(" AND (CLIENTCODE like '%{0}%')", strCLIENTCODE)

            strWhere = BuildFields(" ", Con1, Con2, Con3)
            strFormat = "SELECT TRANSYEARMON,  TRANSNO,SEQ,  CLIENTCODE, dbo.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, TRUST_SEQ, MEDCODE, dbo.MD_GET_CUSTNAME_FUN(MEDCODE) MEDNAME,REAL_MED_CODE,dbo.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) REAL_MED_NAME,DEPT_CD, DEMANDDAY, PRINTDAY, PROGRAM, TBRDSTDATE, TBRDEDDATE, AMT, TRU_TAX_FLAG, VAT, SUMAMTVAT, MEMO , MED_FLAG, SPONSOR, TAXYEARMON, TAXNO FROM MD_PROMOTIONTRANS_DTL WHERE 1=1 {0} ORDER BY DEMANDDAY "
            strSQL = String.Format(strFormat, strWhere)
            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetJOBSEQList")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "35. ���θ�� ������ �ŷ��� ��ȸ"
    Public Function Get_PROMOTIONCOMMI_HDR(ByVal strInfoXML As String, _
                                           ByRef intRowCnt As Integer, _
                                           ByRef intColCnt As Integer, _
                                           ByVal strYEARMON As String, _
                                           ByVal strTRANSNO As String, _
                                           ByVal strREAL_MED_CODE As String) As String

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strWhere As String
        Dim strXMLData As String
        Dim Con1 As String
        Dim Con2 As String
        Dim Con3 As String


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            Con1 = ""
            Con2 = ""
            Con3 = ""
            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
            If strTRANSNO <> "" Then Con2 = String.Format(" AND (TRANSNO = '{0}')", strTRANSNO)
            If strREAL_MED_CODE <> "" Then Con3 = String.Format(" AND (REAL_MED_CODE like '%{0}%')", strREAL_MED_CODE)

            strWhere = BuildFields(" ", Con1, Con2, Con3)
            strFormat = "SELECT TRANSYEARMON, TRANSNO, dbo.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, dbo.SC_DEPT_NAME_FUN(DEPT_CD) DEPT_NAME, DEMANDDAY, PRINTDAY, AMT, VAT, (AMT + VAT) SUMAMTVAT  FROM MD_PROMOTIONCOMMI_HDR WHERE 1=1 {0} ORDER BY MED_FLAG "
            strSQL = String.Format(strFormat, strWhere)
            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                strXMLData = .mobjSCGLSql.SQLSelectXml(strSQL, intRowCnt, intColCnt)
                Return strXMLData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetJOBSEQList")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    Public Function Get_PROMOTIONCOMMI_LIST(ByVal strInfoXML As String, _
                                           ByRef intRowCnt As Integer, _
                                           ByRef intColCnt As Integer, _
                                           ByVal strYEARMON As String, _
                                           ByVal strTRANSNO As String, _
                                           ByVal strREAL_MED_CODE As String) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strWhere As String
        Dim vntData As Object
        Dim Con1 As String
        Dim Con2 As String
        Dim Con3 As String


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            Con1 = ""
            Con2 = ""
            Con3 = ""
            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
            If strTRANSNO <> "" Then Con2 = String.Format(" AND (TRANSNO = '{0}')", strTRANSNO)
            If strREAL_MED_CODE <> "" Then Con3 = String.Format(" AND (REAL_MED_CODE like '%{0}%')", strREAL_MED_CODE)

            strWhere = BuildFields(" ", Con1, Con2, Con3)
            strFormat = "SELECT TRANSYEARMON,  TRANSNO, CLIENTCODE, dbo.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, MEDCODE, dbo.MD_GET_CUSTNAME_FUN(MEDCODE) MEDNAME,REAL_MED_CODE, dbo.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) REAL_MED_NAME,DEPT_CD, DEMANDDAY, PRINTDAY, AMT, SUSURATE, SUSU, VAT, MEMO, MED_FLAG, SPONSOR FROM MD_PROMOTIONCOMMI_DTL WHERE 1=1 {0} ORDER BY REAL_MED_CODE "
            strSQL = String.Format(strFormat, strWhere)
            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetJOBSEQList")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "36. ������ ��ȸ"
    ' =============== SelectRtnSample Code
    Public Function GetCUSTLISTNO(ByVal strInfoXML As String, _
                                  ByRef intRowCnt As Integer, _
                                  ByRef intColCnt As Integer, _
                                  ByVal strCUSTCODE As String, _
                                  ByVal strCUSTNAME As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""

                If strCUSTCODE <> "" Then Con1 = String.Format(" AND (CUSTCODE LIKE '%{0}%')", strCUSTCODE)
                If strCUSTNAME <> "" Then Con2 = String.Format(" AND (COMPANYNAME LIKE '%{0}%')", strCUSTNAME)
                strWhere = BuildFields(" ", Con1, Con2)

                strSelFields = "CUSTCODE , CUSTNAME , COMPANYNAME"

                strFormet = "select {0} from SC_CUST_TEMP where 1=1 {1} "


                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetCUSTLISTNO")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "37. ����ó ��ȸ"
    ' =============== ����ó ��ȸ
    Public Function GetOUTCUSTNO(ByVal strInfoXML As String, _
                                 ByRef intRowCnt As Integer, _
                                 ByRef intColCnt As Integer, _
                                 ByVal strCUSTCODE As String, _
                                 ByVal strCUSTNAME As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""

                If strCUSTCODE <> "" Then Con1 = String.Format(" AND (CUSTCODE LIKE '%{0}%')", strCUSTCODE)
                If strCUSTNAME <> "" Then Con2 = String.Format(" AND (CUSTNAME LIKE '%{0}%')", strCUSTNAME)
                strWhere = BuildFields(" ", Con1, Con2)

                strSelFields = "CUSTCODE , CUSTNAME , COMPANYNAME"

                strFormet = "select {0} from SC_CUST_TEMP where CUSTCODE LIKE 'M%'{1} "


                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetOUTCUSTNO")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "38. ������ ��ȸ"

    ' =============== SelectRtnSample Code
    Public Function GetCUSTNO(ByVal strInfoXML As String, _
                              ByRef intRowCnt As Integer, _
                              ByRef intColCnt As Integer, _
                              ByVal strCUSTCODE As String, _
                              ByVal strCUSTNAME As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""

                If strCUSTCODE <> "" Then Con1 = String.Format(" AND (CUSTCODE LIKE '%{0}%')", strCUSTCODE)
                If strCUSTNAME <> "" Then Con2 = String.Format(" AND (CUSTNAME LIKE '%{0}%')", strCUSTNAME)
                strWhere = BuildFields(" ", Con1, Con2)

                strSelFields = "CUSTCODE, CUSTNAME, BUSINO, COMPANYNAME,ACCUSTCODE"

                strFormet = "select {0} from SC_CUST_TEMP where isnull(DEMANDFLAG,'') = '1' AND CUSTCODE LIKE 'A%' AND ATTR10 =1 {1}  ORDER BY "
                strFormet = strFormet & " CASE SUBSTRING(LTRIM(CUSTNAME),1,3) WHEN '(��)' THEN LTRIM(SUBSTRING(CUSTNAME,4,LEN(CUSTNAME))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(CUSTNAME,4,LEN(CUSTNAME))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(CUSTNAME,4,LEN(CUSTNAME))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(CUSTNAME,4,LEN(CUSTNAME))) "
                strFormet = strFormet & " WHEN '(���' THEN LTRIM(SUBSTRING(CUSTNAME,5,LEN(CUSTNAME))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(CUSTNAME,4,LEN(CUSTNAME))) "
                strFormet = strFormet & " ELSE LTRIM(CUSTNAME) END "


                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetCUSTNO")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "06. �ŷ����� POP��ȸ"
    Public Function GetTRANSNO(ByVal strInfoXML As String, _
                             ByRef intRowCnt As Integer, _
                             ByRef intColCnt As Integer, _
                             ByVal strTRANSYEARMON As String, _
                             ByVal strTRANSNO As String, _
                             ByVal strCLIENTCODE As String, _
                             ByVal strCLIENTNAME As String, _
                             ByVal strFlag As String, _
                             ByVal strTBL_Flag As String, _
                             ByVal strEndFlag As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2, Con3, Con4, Con5 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""
                Con3 = ""
                Con4 = ""
                Con5 = ""


                If strTRANSYEARMON <> "" Then Con1 = String.Format(" AND (A.TRANSYEARMON = '{0}')", strTRANSYEARMON)
                If strTRANSNO <> "" Then Con2 = String.Format(" AND (A.TRANSNO = '{0}')", strTRANSNO)
                If strTBL_Flag = "ETC" Then
                    If strFlag = "trans" Then
                        If strCLIENTCODE <> "" Then Con3 = String.Format(" AND (A.CLIENTCODE = '{0}')", strCLIENTCODE)
                        If strCLIENTNAME <> "" Then Con4 = String.Format(" AND (DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE) LIKE '%{0}%')", strCLIENTNAME)
                    ElseIf strFlag = "commi" Then
                        If strCLIENTCODE <> "" Then Con3 = String.Format(" AND (A.CLIENTCODE = '{0}')", strCLIENTCODE)
                        If strCLIENTNAME <> "" Then Con4 = String.Format(" AND (DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE) LIKE '%{0}%')", strCLIENTNAME)
                    End If
                Else
                    If strFlag = "trans" Then
                        If strCLIENTCODE <> "" Then Con3 = String.Format(" AND (A.CLIENTCODE = '{0}')", strCLIENTCODE)
                        If strCLIENTNAME <> "" Then Con4 = String.Format(" AND (DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE) LIKE '%{0}%')", strCLIENTNAME)
                    ElseIf strFlag = "commi" Then
                        If strCLIENTCODE <> "" Then Con3 = String.Format(" AND (A.REAL_MED_CODE = '{0}')", strCLIENTCODE)
                        If strCLIENTNAME <> "" Then Con4 = String.Format(" AND (DBO.MD_GET_REALMEDCODE_FUN(A.REAL_MED_CODE) LIKE '%{0}%')", strCLIENTNAME)
                    End If
                End If


                If strTBL_Flag = "ELEC" Then
                    If strEndFlag = "0" Then
                        Con5 = String.Format(" AND GBN = '{0}'", "�̿Ϸ�")
                    Else
                        Con5 = String.Format(" AND GBN = '{0}'", "�Ϸ�")
                    End If
                ElseIf strTBL_Flag = "PRINT" Then
                    If strEndFlag = "0" Then
                        Con5 = String.Format(" AND GBN = '{0}'", "�̿Ϸ�")
                    Else
                        Con5 = String.Format(" AND GBN = '{0}'", "�Ϸ�")
                    End If
                ElseIf strTBL_Flag = "CATV" Then
                    If strEndFlag = "0" Then
                        Con5 = String.Format(" AND GBN = '{0}'", "�̿Ϸ�")
                    Else
                        Con5 = String.Format(" AND GBN = '{0}'", "�Ϸ�")
                    End If
                ElseIf strTBL_Flag = "INTERNET" Then
                    If strEndFlag = "0" Then
                        Con5 = String.Format(" AND GBN = '{0}'", "�̿Ϸ�")
                    Else
                        Con5 = String.Format(" AND GBN = '{0}'", "�Ϸ�")
                    End If
                ElseIf strTBL_Flag = "OUTDOOR" Then
                    If strEndFlag = "0" Then
                        Con5 = String.Format(" AND GBN = '{0}'", "�̿Ϸ�")
                    Else
                        Con5 = String.Format(" AND GBN = '{0}'", "�Ϸ�")
                    End If
                ElseIf strTBL_Flag = "PROMOTION" Then
                    If strEndFlag = "0" Then
                        Con5 = String.Format(" AND GBN = '{0}'", "�̿Ϸ�")
                    Else
                        Con5 = String.Format(" AND GBN = '{0}'", "�Ϸ�")
                    End If
                ElseIf strTBL_Flag = "ETC" Then
                    If strEndFlag = "0" Then
                        Con5 = String.Format(" AND GBN = '{0}'", "�̿Ϸ�")
                    Else
                        Con5 = String.Format(" AND GBN = '{0}'", "�Ϸ�")
                    End If
                ElseIf strTBL_Flag = "ELECSPON" Then
                    If strEndFlag = "0" Then
                        Con5 = String.Format(" AND GBN = '{0}'", "�̿Ϸ�")
                    Else
                        Con5 = String.Format(" AND GBN = '{0}'", "�Ϸ�")
                    End If
                End If
                strWhere = BuildFields(" ", Con1, Con2, Con3, Con4, Con5)


                If strTBL_Flag = "ELEC" Then
                    If strFlag = "trans" Then
                        strSelFields = "A.TRANSYEARMON , A.TRANSNO , A.CLIENTCODE, DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE) CLIENTNAME,B.GBN"
                        strFormet = " select {0} "
                        strFormet = strFormet & " FROM MD_ELEC_TRANS_HDR A LEFT JOIN "
                        strFormet = strFormet & " ("
                        strFormet = strFormet & "  SELECT TRANSYEARMON,TRANSNO,CASE SUM(CASE ISNULL(TAXNO,0) WHEN 0 THEN 1 ELSE 0 END) WHEN 0 THEN '�Ϸ�' ELSE '�̿Ϸ�' END AS GBN"
                        strFormet = strFormet & "  FROM MD_ELEC_TRANS_DTL"
                        strFormet = strFormet & "  WHERE AMT >0 AND AMT IS NOT NULL "
                        strFormet = strFormet & "  GROUP BY TRANSYEARMON,TRANSNO"
                        strFormet = strFormet & "  )"
                        strFormet = strFormet & " B ON A.TRANSYEARMON = B.TRANSYEARMON"
                        strFormet = strFormet & " AND A.TRANSNO = B.TRANSNO WHERE 1=1 and (A.ATTR03 = 'N' OR ISNULL(A.ATTR03,'') = '') {1} ORDER BY "
                        strFormet = strFormet & " CASE SUBSTRING(LTRIM(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)),1,3) WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)))) "
                        strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)))) "
                        strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)))) "
                        strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)))) "
                        strFormet = strFormet & " WHEN '(���' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE),5,LEN(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)))) "
                        strFormet = strFormet & " WHEN '(�纹' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE),5,LEN(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)))) "
                        strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)))) "
                        strFormet = strFormet & " ELSE LTRIM(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)) END "
                    ElseIf strFlag = "commi" Then
                        strSelFields = "A.TRANSYEARMON , A.TRANSNO , A.REAL_MED_CODE, DBO.MD_GET_REALMEDCODE_FUN(A.REAL_MED_CODE) REAL_MED_NAME, B.GBN"
                        strFormet = " select {0} "
                        strFormet = strFormet & " FROM MD_ELECCOMMI_HDR A LEFT JOIN "
                        strFormet = strFormet & " ("
                        strFormet = strFormet & "  SELECT TRANSYEARMON,TRANSNO,CASE SUM(CASE ATTR02 WHEN '' THEN 1 ELSE 0 END) WHEN 0 THEN '�Ϸ�' ELSE '�̿Ϸ�' END AS GBN"
                        strFormet = strFormet & "  FROM MD_ELECCOMMI_DTL"
                        strFormet = strFormet & "  GROUP BY TRANSYEARMON,TRANSNO"
                        strFormet = strFormet & "  )"
                        strFormet = strFormet & " B ON A.TRANSYEARMON = B.TRANSYEARMON"
                        strFormet = strFormet & " AND A.TRANSNO = B.TRANSNO WHERE 1=1 {1}"
                    End If
                ElseIf strTBL_Flag = "PRINT" Then
                    If strFlag = "trans" Then
                        strSelFields = "A.TRANSYEARMON , A.TRANSNO , A.CLIENTCODE, DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE) CLIENTNAME,B.GBN"
                        strFormet = " select {0} "
                        strFormet = strFormet & " FROM MD_PRINTTRANS_HDR A LEFT JOIN "
                        strFormet = strFormet & " ("
                        strFormet = strFormet & "  SELECT TRANSYEARMON,TRANSNO,CASE SUM(CASE ISNULL(TAXNO,0) WHEN 0 THEN 1 ELSE 0 END) WHEN 0 THEN '�Ϸ�' ELSE '�̿Ϸ�' END AS GBN"
                        strFormet = strFormet & "  FROM MD_PRINTTRANS_DTL"
                        strFormet = strFormet & "  WHERE AMT >0 AND AMT IS NOT NULL "
                        'strFormet = strFormet & "  AND ISNULL(CONFIRMFLAG,0) = 1 "
                        strFormet = strFormet & "  GROUP BY TRANSYEARMON,TRANSNO"
                        strFormet = strFormet & "  )"
                        strFormet = strFormet & " B ON A.TRANSYEARMON = B.TRANSYEARMON"
                        strFormet = strFormet & " AND A.TRANSNO = B.TRANSNO WHERE 1=1 {1} ORDER BY  "
                        strFormet = strFormet & " CASE SUBSTRING(LTRIM(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)),1,3) WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)))) "
                        strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)))) "
                        strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)))) "
                        strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)))) "
                        strFormet = strFormet & " WHEN '(���' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE),5,LEN(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)))) "
                        strFormet = strFormet & " WHEN '(�纹' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE),5,LEN(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)))) "
                        strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)))) "
                        strFormet = strFormet & " ELSE LTRIM(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)) END "
                    ElseIf strFlag = "commi" Then
                        strSelFields = "A.TRANSYEARMON , A.TRANSNO , A.REAL_MED_CODE, DBO.MD_GET_REALMEDCODE_FUN(A.REAL_MED_CODE) REAL_MED_NAME, B.GBN"
                        strFormet = " select {0} "
                        strFormet = strFormet & " FROM MD_PRINTCOMMI_HDR A LEFT JOIN "
                        strFormet = strFormet & " ("
                        strFormet = strFormet & "  SELECT TRANSYEARMON,TRANSNO,CASE SUM(CASE ISNULL(TAXNO,0) WHEN 0 THEN 1 ELSE 0 END) WHEN 0 THEN '�Ϸ�' ELSE '�̿Ϸ�' END AS GBN"
                        strFormet = strFormet & "  FROM MD_PRINTCOMMI_DTL"
                        strFormet = strFormet & "  WHERE AMT >0 AND AMT IS NOT NULL "
                        strFormet = strFormet & "  GROUP BY TRANSYEARMON,TRANSNO"
                        strFormet = strFormet & "  )"
                        strFormet = strFormet & " B ON A.TRANSYEARMON = B.TRANSYEARMON"
                        strFormet = strFormet & " AND A.TRANSNO = B.TRANSNO WHERE 1=1 {1}"
                    End If
                ElseIf strTBL_Flag = "CATV" Then
                    If strFlag = "trans" Then
                        strSelFields = "A.TRANSYEARMON , A.TRANSNO , A.CLIENTCODE, DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE) CLIENTNAME,B.GBN"
                        strFormet = " select {0} "
                        strFormet = strFormet & " FROM MD_CATVTRANS_HDR A LEFT JOIN "
                        strFormet = strFormet & " ("
                        strFormet = strFormet & "  SELECT TRANSYEARMON,TRANSNO,CASE SUM(CASE ISNULL(TAXNO,0) WHEN 0 THEN 1 ELSE 0 END) WHEN 0 THEN '�Ϸ�' ELSE '�̿Ϸ�' END AS GBN"
                        strFormet = strFormet & "  FROM MD_CATVTRANS_DTL"
                        strFormet = strFormet & "  WHERE AMT >0 AND AMT IS NOT NULL "
                        'strFormet = strFormet & "  AND ISNULL(CONFIRMFLAG,0) = 1 "
                        strFormet = strFormet & "  GROUP BY TRANSYEARMON,TRANSNO"
                        strFormet = strFormet & "  )"
                        strFormet = strFormet & " B ON A.TRANSYEARMON = B.TRANSYEARMON"
                        strFormet = strFormet & " AND A.TRANSNO = B.TRANSNO WHERE 1=1 {1} ORDER BY  "
                        strFormet = strFormet & " CASE SUBSTRING(LTRIM(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)),1,3) WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)))) "
                        strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)))) "
                        strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)))) "
                        strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)))) "
                        strFormet = strFormet & " WHEN '(���' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE),5,LEN(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)))) "
                        strFormet = strFormet & " WHEN '(�纹' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE),5,LEN(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)))) "
                        strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)))) "
                        strFormet = strFormet & " ELSE LTRIM(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)) END "
                    ElseIf strFlag = "commi" Then
                        strSelFields = "A.TRANSYEARMON , A.TRANSNO , A.REAL_MED_CODE, DBO.MD_GET_REALMEDCODE_FUN(A.REAL_MED_CODE) REAL_MED_NAME,B.GBN"
                        strFormet = "select {0} "
                        strFormet = strFormet & " FROM MD_CATVCOMMI_HDR A LEFT JOIN "
                        strFormet = strFormet & " ("
                        strFormet = strFormet & " SELECT TRANSYEARMON,TRANSNO,CASE SUM(CASE ISNULL(TAXNO,0) WHEN 0 THEN 1 ELSE 0 END) WHEN 0 THEN '�Ϸ�' ELSE '�̿Ϸ�' END AS GBN"
                        strFormet = strFormet & " FROM MD_CATVCOMMI_DTL"
                        strFormet = strFormet & " GROUP BY TRANSYEARMON,TRANSNO"
                        strFormet = strFormet & " )"
                        strFormet = strFormet & " B ON A.TRANSYEARMON = B.TRANSYEARMON"
                        strFormet = strFormet & " AND A.TRANSNO = B.TRANSNO WHERE 1=1 {1}"
                    End If
                ElseIf strTBL_Flag = "INTERNET" Then
                    If strFlag = "trans" Then
                        strSelFields = "A.TRANSYEARMON , A.TRANSNO , A.CLIENTCODE, DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE) CLIENTNAME,B.GBN"
                        strFormet = " select {0} "
                        strFormet = strFormet & " FROM MD_INTERNETTRANS_HDR A LEFT JOIN "
                        strFormet = strFormet & " ("
                        strFormet = strFormet & "  SELECT TRANSYEARMON,TRANSNO,CASE SUM(CASE ISNULL(TAXNO,0) WHEN 0 THEN 1 ELSE 0 END) WHEN 0 THEN '�Ϸ�' ELSE '�̿Ϸ�' END AS GBN"
                        strFormet = strFormet & "  FROM MD_INTERNETTRANS_DTL"
                        strFormet = strFormet & "  WHERE AMT >0 AND AMT IS NOT NULL "
                        'strFormet = strFormet & "  AND ISNULL(CONFIRMFLAG,0) = 1 "
                        strFormet = strFormet & "  GROUP BY TRANSYEARMON,TRANSNO"
                        strFormet = strFormet & "  )"
                        strFormet = strFormet & " B ON A.TRANSYEARMON = B.TRANSYEARMON"
                        strFormet = strFormet & " AND A.TRANSNO = B.TRANSNO WHERE 1=1 {1} ORDER BY"
                        strFormet = strFormet & " CASE SUBSTRING(LTRIM(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)),1,3) WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)))) "
                        strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)))) "
                        strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)))) "
                        strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)))) "
                        strFormet = strFormet & " WHEN '(���' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE),5,LEN(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)))) "
                        strFormet = strFormet & " WHEN '(�纹' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE),5,LEN(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)))) "
                        strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)))) "
                        strFormet = strFormet & " ELSE LTRIM(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)) END "
                    ElseIf strFlag = "commi" Then
                        strSelFields = "A.TRANSYEARMON , A.TRANSNO , A.REAL_MED_CODE, DBO.MD_GET_REALMEDCODE_FUN(A.REAL_MED_CODE) REAL_MED_NAME, B.GBN"
                        strFormet = " select {0} "
                        strFormet = strFormet & " FROM MD_INTERNETCOMMI_HDR A LEFT JOIN "
                        strFormet = strFormet & " ("
                        strFormet = strFormet & "  SELECT TRANSYEARMON,TRANSNO,CASE SUM(CASE ISNULL(TAXNO,0) WHEN 0 THEN 1 ELSE 0 END) WHEN 0 THEN '�Ϸ�' ELSE '�̿Ϸ�' END AS GBN"
                        strFormet = strFormet & "  FROM MD_INTERNETCOMMI_DTL"
                        strFormet = strFormet & "  GROUP BY TRANSYEARMON,TRANSNO"
                        strFormet = strFormet & "  )"
                        strFormet = strFormet & " B ON A.TRANSYEARMON = B.TRANSYEARMON"
                        strFormet = strFormet & " AND A.TRANSNO = B.TRANSNO WHERE 1=1 {1}"
                    End If
                ElseIf strTBL_Flag = "OUTDOOR" Then
                    If strFlag = "trans" Then
                        strSelFields = "A.TRANSYEARMON , A.TRANSNO , A.CLIENTCODE, DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE) CLIENTNAME,B.GBN"
                        strFormet = " select {0} "
                        strFormet = strFormet & " FROM MD_OUTDOORTRANS_HDR A LEFT JOIN "
                        strFormet = strFormet & " ("
                        strFormet = strFormet & "  SELECT TRANSYEARMON,TRANSNO,CASE SUM(CASE ISNULL(TAXNO,0) WHEN 0 THEN 1 ELSE 0 END) WHEN 0 THEN '�Ϸ�' ELSE '�̿Ϸ�' END AS GBN"
                        strFormet = strFormet & "  FROM MD_OUTDOORTRANS_DTL"
                        strFormet = strFormet & "  WHERE AMT >0 AND AMT IS NOT NULL "
                        'strFormet = strFormet & "  AND ISNULL(CONFIRMFLAG,0) = 1 "
                        strFormet = strFormet & "  GROUP BY TRANSYEARMON,TRANSNO"
                        strFormet = strFormet & "  )"
                        strFormet = strFormet & " B ON A.TRANSYEARMON = B.TRANSYEARMON"
                        strFormet = strFormet & " AND A.TRANSNO = B.TRANSNO WHERE 1=1 {1} ORDER BY"
                        strFormet = strFormet & " CASE SUBSTRING(LTRIM(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)),1,3) WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)))) "
                        strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)))) "
                        strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)))) "
                        strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)))) "
                        strFormet = strFormet & " WHEN '(���' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE),5,LEN(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)))) "
                        strFormet = strFormet & " WHEN '(�纹' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE),5,LEN(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)))) "
                        strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)))) "
                        strFormet = strFormet & " ELSE LTRIM(DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE)) END "
                    ElseIf strFlag = "commi" Then
                        strSelFields = "A.TRANSYEARMON , A.TRANSNO , A.REAL_MED_CODE, DBO.MD_GET_REALMEDCODE_FUN(A.REAL_MED_CODE) REAL_MED_NAME, B.GBN"
                        strFormet = " select {0} "
                        strFormet = strFormet & " FROM MD_OUTDOORCOMMI_HDR A LEFT JOIN "
                        strFormet = strFormet & " ("
                        strFormet = strFormet & "  SELECT TRANSYEARMON,TRANSNO,CASE SUM(CASE ISNULL(TAXNO,0) WHEN 0 THEN 1 ELSE 0 END) WHEN 0 THEN '�Ϸ�' ELSE '�̿Ϸ�' END AS GBN"
                        strFormet = strFormet & "  FROM MD_OUTDOORCOMMI_DTL"
                        strFormet = strFormet & "  GROUP BY TRANSYEARMON,TRANSNO"
                        strFormet = strFormet & "  )"
                        strFormet = strFormet & " B ON A.TRANSYEARMON = B.TRANSYEARMON"
                        strFormet = strFormet & " AND A.TRANSNO = B.TRANSNO WHERE 1=1 {1}"
                    End If
                ElseIf strTBL_Flag = "ELECSPON" Then
                    If strFlag = "trans" Then
                        strSelFields = "A.TRANSYEARMON , A.TRANSNO , A.CLIENTCODE, DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE) CLIENTNAME,B.GBN"
                        strFormet = " select {0} "
                        strFormet = strFormet & " FROM MD_ELEC_TRANS_HDR A LEFT JOIN "
                        strFormet = strFormet & " ("
                        strFormet = strFormet & "  SELECT TRANSYEARMON,TRANSNO,CASE SUM(CASE ISNULL(TAXNO,0) WHEN 0 THEN 1 ELSE 0 END) WHEN 0 THEN '�Ϸ�' ELSE '�̿Ϸ�' END AS GBN"
                        strFormet = strFormet & "  FROM MD_ELEC_TRANS_DTL"
                        strFormet = strFormet & "  GROUP BY TRANSYEARMON,TRANSNO"
                        strFormet = strFormet & "  )"
                        strFormet = strFormet & " B ON A.TRANSYEARMON = B.TRANSYEARMON"
                        strFormet = strFormet & " AND A.TRANSNO = B.TRANSNO WHERE 1=1 AND (A.ATTR03 = 'Y' AND A.ATTR03 IS NOT NULL) {1} "
                    ElseIf strFlag = "commi" Then
                        strSelFields = "A.TRANSYEARMON , A.TRANSNO , A.REAL_MED_CODE, DBO.MD_GET_REALMEDCODE_FUN(A.REAL_MED_CODE) REAL_MED_NAME, B.GBN"
                        strFormet = " select {0} "
                        strFormet = strFormet & " FROM MD_ELECCOMMI_HDR A LEFT JOIN "
                        strFormet = strFormet & " ("
                        strFormet = strFormet & "  SELECT TRANSYEARMON,TRANSNO,CASE SUM(CASE ATTR02 WHEN '' THEN 1 ELSE 0 END) WHEN 0 THEN '�Ϸ�' ELSE '�̿Ϸ�' END AS GBN"
                        strFormet = strFormet & "  FROM MD_ELECCOMMI_DTL"
                        strFormet = strFormet & "  GROUP BY TRANSYEARMON,TRANSNO"
                        strFormet = strFormet & "  )"
                        strFormet = strFormet & " B ON A.TRANSYEARMON = B.TRANSYEARMON"
                        strFormet = strFormet & " AND A.TRANSNO = B.TRANSNO WHERE 1=1 AND A.ATTR03 = 'Y' {1}"
                    End If

                End If

                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetTRANSNO")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function


#End Region

#Region "06. �ŷ����������� POP��ȸ"
    Public Function GetTRANSCUSTNO(ByVal strInfoXML As String, _
                                   ByRef intRowCnt As Integer, _
                                   ByRef intColCnt As Integer, _
                                   ByVal strYEARMON As String, _
                                   ByVal strCLIENTCODE As String, _
                                   ByVal strCLIENTNAME As String, _
                                   ByVal strCOMMITCHECK As String, _
                                   ByVal strFlag As String, _
                                   ByVal strTBL_Flag As String, _
                                   ByVal strSPONSOR As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2, Con3, Con4, Con5 As String
        Dim vntData As Object


        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""
                Con3 = ""
                Con4 = ""
                Con5 = ""

                If strTBL_Flag = "PRINT" Or strTBL_Flag = "CATV" Or strTBL_Flag = "INTERNET" Or strTBL_Flag = "OUTDOOR" Then
                    If strFlag = "trans" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
                            If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (CLIENTCODE LIKE '%{0}%')", strCLIENTCODE)
                            If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCLIENTNAME)
                            If strSPONSOR <> "" Then Con4 = String.Format(" AND (SPONSOR = '{0}')", strSPONSOR)
                        Else
                            If strYEARMON <> "" Then Con1 = String.Format(" AND (YEARMON = '{0}')", strYEARMON)
                            If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (CLIENTCODE LIKE '%{0}%')", strCLIENTCODE)
                            If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCLIENTNAME)
                            If strSPONSOR <> "" Then Con4 = String.Format(" AND (SPONSOR = '{0}')", strSPONSOR)
                        End If
                    ElseIf strFlag = "commi" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
                            If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (REAL_MED_CODE LIKE '%{0}%')", strCLIENTCODE)
                            If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) LIKE '%{0}%')", strCLIENTNAME)
                            If strSPONSOR <> "" Then Con4 = String.Format(" AND (SPONSOR = '{0}')", strSPONSOR)
                        Else
                            If strYEARMON <> "" Then Con1 = String.Format(" AND (YEARMON = '{0}')", strYEARMON)
                            If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (REAL_MED_CODE LIKE '%{0}%')", strCLIENTCODE)
                            If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) LIKE '%{0}%')", strCLIENTNAME)
                            If strSPONSOR <> "" Then Con4 = String.Format(" AND (SPONSOR = '{0}')", strSPONSOR)
                        End If
                    End If

                ElseIf strTBL_Flag = "ELECSPON" Then
                    If strFlag = "trans" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
                            If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (CLIENTCODE LIKE '%{0}%')", strCLIENTCODE)
                            If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCLIENTNAME)
                            If strSPONSOR <> "" Then Con4 = String.Format(" AND (SPONSOR = '{0}')", strSPONSOR)
                        Else
                            If strYEARMON <> "" Then Con1 = String.Format(" AND (YEARMON = '{0}')", strYEARMON)
                            If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (CLIENTCODE LIKE '%{0}%')", strCLIENTCODE)
                            If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCLIENTNAME)
                            If strSPONSOR <> "" Then Con4 = String.Format(" AND (SPONSOR = '{0}')", strSPONSOR)
                        End If
                    ElseIf strFlag = "commi" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
                            If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (REAL_MED_CODE LIKE '%{0}%')", strCLIENTCODE)
                            If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) LIKE '%{0}%')", strCLIENTNAME)
                            If strSPONSOR <> "" Then Con4 = String.Format(" AND (ATTR03 = '{0}')", strSPONSOR)
                        Else
                            If strYEARMON <> "" Then Con1 = String.Format(" AND (YEARMON = '{0}')", strYEARMON)
                            If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (REAL_MED_CODE LIKE '%{0}%')", strCLIENTCODE)
                            If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) LIKE '%{0}%')", strCLIENTNAME)
                            If strSPONSOR <> "" Then Con4 = String.Format(" AND (SPONSOR = '{0}')", strSPONSOR)
                        End If
                    End If

                ElseIf strTBL_Flag = "ELEC" Then
                    If strFlag = "trans" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
                            If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (CLIENTCODE LIKE '%{0}%')", strCLIENTCODE)
                            If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCLIENTNAME)
                            If strSPONSOR <> "" Then Con4 = String.Format(" AND (SPONSOR = '{0}')", strSPONSOR)
                        Else
                            If strYEARMON <> "" Then Con1 = String.Format(" AND (YEARMON = '{0}')", strYEARMON)
                            If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (CLIENTCODE LIKE '%{0}%')", strCLIENTCODE)
                            If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCLIENTNAME)
                            If strSPONSOR <> "" Then Con4 = String.Format(" AND (SPONSOR = '{0}')", strSPONSOR)
                        End If
                    ElseIf strFlag = "commi" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
                            If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (REAL_MED_CODE LIKE '%{0}%')", strCLIENTCODE)
                            If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) LIKE '%{0}%')", strCLIENTNAME)
                            If strSPONSOR <> "" Then Con4 = String.Format(" AND (SPONSOR = '{0}')", strSPONSOR)
                        Else
                            If strYEARMON <> "" Then Con1 = String.Format(" AND (YEARMON = '{0}')", strYEARMON)
                            If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (REAL_MED_CODE LIKE '%{0}%')", strCLIENTCODE)
                            If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) LIKE '%{0}%')", strCLIENTNAME)
                            If strSPONSOR <> "" Then Con4 = String.Format(" AND (SPONSOR = '{0}')", strSPONSOR)
                        End If
                    End If
                End If
                strWhere = BuildFields(" ", Con1, Con2, Con3, Con4)


                If strTBL_Flag = "PRINT" Then
                    If strFlag = "trans" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            strSelFields = "TRANSYEARMON , TRANSNO, DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, '�Ϸ�' GBN, CLIENTCODE"
                            strFormet = "select {0}  "
                            strFormet = strFormet & " FROM MD_PRINTTRANS_DTL  WHERE 1=1 {1} "
                            strFormet = strFormet & " GROUP BY CLIENTCODE , TRANSYEARMON, TRANSNO "
                            strFormet = strFormet & " ORDER BY DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE)"
                        Else
                            strSelFields = "YEARMON , CLIENTCODE, DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, '�̿Ϸ�' GBN"
                            strFormet = "select {0}    "
                            strFormet = strFormet & " FROM MD_BOOKING_MEDIUM  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND gflag in('B') "
                            strFormet = strFormet & " AND isnull(tru_trans_no, '') = ''  "
                            strFormet = strFormet & " GROUP BY CLIENTCODE , YEARMON "
                            strFormet = strFormet & " ORDER BY "
                            strFormet = strFormet & " DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE)"
                        End If
                    ElseIf strFlag = "commi" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            strSelFields = "TRANSYEARMON , TRANSNO, DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) REAL_MED_NAME, '�Ϸ�' GBN, REAL_MED_CODE"
                            strFormet = "select {0}  "
                            strFormet = strFormet & " FROM MD_PRINTCOMMI_DTL  WHERE 1=1 {1} "
                            strFormet = strFormet & " GROUP BY REAL_MED_CODE , TRANSYEARMON, TRANSNO "
                            strFormet = strFormet & " ORDER BY DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) "

                        Else
                            strSelFields = "YEARMON , REAL_MED_CODE, DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) REAL_MED_NAME, '�̿Ϸ�' GBN"
                            strFormet = "select {0}    "
                            strFormet = strFormet & " FROM MD_BOOKING_MEDIUM  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND gflag in('B','J', 'S') "
                            strFormet = strFormet & " AND isnull(commi_trans_no, '') = ''  "
                            strFormet = strFormet & " GROUP BY REAL_MED_CODE , YEARMON "
                            strFormet = strFormet & " ORDER BY "
                            strFormet = strFormet & " DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE)"
                        End If
                    End If
                ElseIf strTBL_Flag = "CATV" Then
                    If strFlag = "trans" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            strSelFields = "TRANSYEARMON , TRANSNO, DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, '�Ϸ�' GBN, CLIENTCODE"
                            strFormet = "select {0}  "
                            strFormet = strFormet & " FROM MD_CATVTRANS_DTL  WHERE 1=1 {1} "
                            strFormet = strFormet & " GROUP BY CLIENTCODE , TRANSYEARMON, TRANSNO "
                            strFormet = strFormet & " ORDER BY DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE)"

                        Else
                            strSelFields = "YEARMON , CLIENTCODE, DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, '�̿Ϸ�' GBN"
                            strFormet = "select {0}    "
                            strFormet = strFormet & " FROM MD_CATV_MEDIUM  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND GFLAG = '1' "
                            strFormet = strFormet & " AND isnull(tru_trans_no, '') = ''  "
                            strFormet = strFormet & " GROUP BY CLIENTCODE , YEARMON "
                            strFormet = strFormet & " ORDER BY "
                            strFormet = strFormet & " DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE)"
                        End If
                    ElseIf strFlag = "commi" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            strSelFields = "TRANSYEARMON , TRANSNO, DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) REAL_MED_NAME, '�Ϸ�' GBN, REAL_MED_CODE"
                            strFormet = "select {0}  "
                            strFormet = strFormet & " FROM MD_CATVCOMMI_DTL  WHERE 1=1 {1} "
                            strFormet = strFormet & " GROUP BY REAL_MED_CODE , TRANSYEARMON, TRANSNO "
                            strFormet = strFormet & " ORDER BY DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) "
                        Else
                            strSelFields = "YEARMON , REAL_MED_CODE, DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) REAL_MED_NAME, '�̿Ϸ�' GBN"
                            strFormet = "select {0}    "
                            strFormet = strFormet & " FROM MD_CATV_MEDIUM  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND isnull(commi_trans_no, '') = ''  "
                            strFormet = strFormet & " GROUP BY REAL_MED_CODE , YEARMON "
                            strFormet = strFormet & " ORDER BY "
                            strFormet = strFormet & " DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE)"
                        End If
                    End If

                ElseIf strTBL_Flag = "INTERNET" Then
                    If strFlag = "trans" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            strSelFields = "TRANSYEARMON , TRANSNO, DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, '�Ϸ�' GBN, CLIENTCODE"
                            strFormet = "select {0}  "
                            strFormet = strFormet & " FROM MD_INTERNETTRANS_DTL  WHERE 1=1 {1} "
                            strFormet = strFormet & " GROUP BY CLIENTCODE , TRANSYEARMON, TRANSNO "
                            strFormet = strFormet & " ORDER BY DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE)"
                        Else
                            strSelFields = "YEARMON , CLIENTCODE, DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, '�̿Ϸ�' GBN"
                            strFormet = "select {0}    "
                            strFormet = strFormet & " FROM MD_INTERNET_MEDIUM  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND GFLAG = '1' "
                            strFormet = strFormet & " AND isnull(tru_trans_no, '') = ''  "
                            strFormet = strFormet & " GROUP BY CLIENTCODE , YEARMON "
                            strFormet = strFormet & " ORDER BY "
                            strFormet = strFormet & " DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE)"
                        End If
                    ElseIf strFlag = "commi" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            strSelFields = "TRANSYEARMON , TRANSNO, DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) REAL_MED_NAME, '�Ϸ�' GBN, REAL_MED_CODE"
                            strFormet = "select {0}  "
                            strFormet = strFormet & " FROM MD_INTERNETCOMMI_DTL  WHERE 1=1 {1} "
                            strFormet = strFormet & " GROUP BY REAL_MED_CODE , TRANSYEARMON, TRANSNO "
                            strFormet = strFormet & " ORDER BY DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) "
                        Else
                            strSelFields = "YEARMON , REAL_MED_CODE, DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) REAL_MED_NAME, '�̿Ϸ�' GBN"
                            strFormet = "select {0}    "
                            strFormet = strFormet & " FROM MD_INTERNET_MEDIUM  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND isnull(commi_trans_no, '') = ''  "
                            strFormet = strFormet & " GROUP BY REAL_MED_CODE , YEARMON "
                            strFormet = strFormet & " ORDER BY "
                            strFormet = strFormet & " DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE)"
                        End If
                    End If
                ElseIf strTBL_Flag = "OUTDOOR" Then
                    If strFlag = "trans" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            strSelFields = "TRANSYEARMON , TRANSNO, DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, '�Ϸ�' GBN, CLIENTCODE"
                            strFormet = "select {0}  "
                            strFormet = strFormet & " FROM MD_OUTDOORTRANS_DTL  WHERE 1=1 {1} "
                            strFormet = strFormet & " GROUP BY CLIENTCODE , TRANSYEARMON, TRANSNO "
                            strFormet = strFormet & " ORDER BY DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE)"
                        Else
                            strSelFields = "YEARMON , CLIENTCODE, DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, '�̿Ϸ�' GBN"
                            strFormet = "select {0}    "
                            strFormet = strFormet & " FROM MD_OUTDOOR_MEDIUM  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND GFLAG = '1' "
                            strFormet = strFormet & " AND GBN_FLAG='0' "
                            strFormet = strFormet & " AND isnull(tru_trans_no, '') = ''  "
                            strFormet = strFormet & " GROUP BY CLIENTCODE , YEARMON "
                            strFormet = strFormet & " ORDER BY "
                            strFormet = strFormet & " DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE)"
                        End If
                    ElseIf strFlag = "commi" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            strSelFields = "TRANSYEARMON , TRANSNO, DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) REAL_MED_NAME, '�Ϸ�' GBN, REAL_MED_CODE"
                            strFormet = "select {0}  "
                            strFormet = strFormet & " FROM MD_OUTDOORCOMMI_DTL  WHERE 1=1 {1} "
                            strFormet = strFormet & " GROUP BY REAL_MED_CODE , TRANSYEARMON, TRANSNO "
                            strFormet = strFormet & " ORDER BY DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) "
                        Else
                            strSelFields = "YEARMON , REAL_MED_CODE, DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) REAL_MED_NAME, '�̿Ϸ�' GBN"
                            strFormet = "select {0}    "
                            strFormet = strFormet & " FROM MD_OUTDOOR_MEDIUM  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND isnull(commi_trans_no, '') = ''  "
                            strFormet = strFormet & " GROUP BY REAL_MED_CODE , YEARMON "
                            strFormet = strFormet & " ORDER BY "
                            strFormet = strFormet & " DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE)"
                        End If
                    End If
                ElseIf strTBL_Flag = "ELECSPON" Then
                    If strFlag = "trans" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            strSelFields = "TRANSYEARMON , TRANSNO, DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, '�Ϸ�' GBN, CLIENTCODE"
                            strFormet = "select {0}  "
                            strFormet = strFormet & " FROM MD_ELEC_TRANS_DTL  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND SPONSOR ='Y' "
                            strFormet = strFormet & " GROUP BY CLIENTCODE , TRANSYEARMON, TRANSNO "
                            strFormet = strFormet & " ORDER BY DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE)"
                        Else
                            strSelFields = "YEARMON , CLIENTCODE, DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, '�̿Ϸ�' GBN"
                            strFormet = "select {0}    "
                            strFormet = strFormet & " FROM MD_ELECTRIC_MEDIUM  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND SPONSOR ='Y' "
                            strFormet = strFormet & " AND isnull(tru_trans_no, '') = ''  "
                            strFormet = strFormet & " GROUP BY CLIENTCODE , YEARMON "
                            strFormet = strFormet & " ORDER BY "
                            strFormet = strFormet & " DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE)"
                        End If
                    ElseIf strFlag = "commi" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            strSelFields = "TRANSYEARMON , TRANSNO, DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) REAL_MED_NAME, '�Ϸ�' GBN, REAL_MED_CODE"
                            strFormet = "select {0}  "
                            strFormet = strFormet & " FROM MD_ELECCOMMI_DTL  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND ATTR03 ='Y' "
                            strFormet = strFormet & " GROUP BY REAL_MED_CODE , TRANSYEARMON, TRANSNO "
                            strFormet = strFormet & " ORDER BY DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) "
                        Else
                            strSelFields = "YEARMON , REAL_MED_CODE, DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) REAL_MED_NAME, '�̿Ϸ�' GBN"
                            strFormet = "select {0}    "
                            strFormet = strFormet & " FROM MD_ELECTRIC_MEDIUM  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND ATTR03 ='Y' "
                            strFormet = strFormet & " AND isnull(commi_trans_no, '') = ''  "
                            strFormet = strFormet & " GROUP BY REAL_MED_CODE , YEARMON "
                            strFormet = strFormet & " ORDER BY "
                            strFormet = strFormet & " DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE)"
                        End If
                    End If
                ElseIf strTBL_Flag = "ELEC" Then
                    'If strFlag = "trans" Then
                    '    If strCOMMITCHECK = "COMMIT" Then
                    '        strSelFields = "TRANSYEARMON , TRANSNO, DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, '�Ϸ�' GBN, CLIENTCODE"
                    '        strFormet = "select {0}  "
                    '        strFormet = strFormet & " FROM MD_ELEC_TRANS_DTL  WHERE 1=1 {1} "
                    '        strFormet = strFormet & " AND SPONSOR ='N' "
                    '        strFormet = strFormet & " AND isnull(tru_trans_no, '') = ''  "
                    '        strFormet = strFormet & " GROUP BY CLIENTCODE , TRANSYEARMON, TRANSNO "
                    '        strFormet = strFormet & " ORDER BY DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE)"
                    '    Else
                    '        strSelFields = "YEARMON , CLIENTCODE, DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, '�̿Ϸ�' GBN"
                    '        strFormet = "select {0}    "
                    '        strFormet = strFormet & " FROM MD_ELECTRIC_MEDIUM  WHERE 1=1 {1} "
                    '        strFormet = strFormet & " GROUP BY CLIENTCODE , YEARMON "
                    '        strFormet = strFormet & " ORDER BY "
                    '        strFormet = strFormet & " DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE)"
                    '    End If
                    'ElseIf strFlag = "commi" Then
                    '    If strCOMMITCHECK = "COMMIT" Then
                    '        strSelFields = "TRANSYEARMON , TRANSNO, DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) REAL_MED_NAME, '�Ϸ�' GBN, REAL_MED_CODE"
                    '        strFormet = "select {0}  "
                    '        strFormet = strFormet & " FROM MD_ELECCOMMI_DTL  WHERE 1=1 {1} "
                    '        strFormet = strFormet & " GROUP BY REAL_MED_CODE , TRANSYEARMON, TRANSNO "
                    '        strFormet = strFormet & " ORDER BY DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) "
                    '    Else
                    '        strSelFields = "YEARMON , REAL_MED_CODE, DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) REAL_MED_NAME, '�̿Ϸ�' GBN"
                    '        strFormet = "select {0}    "
                    '        strFormet = strFormet & " FROM MD_ELECTRIC_MEDIUM  WHERE 1=1 {1} "
                    '        strFormet = strFormet & " GROUP BY REAL_MED_CODE , YEARMON "
                    '        strFormet = strFormet & " ORDER BY "
                    '        strFormet = strFormet & " DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE)"
                    '    End If
                    'End If
                    If strFlag = "trans" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            strSelFields = "TRANSYEARMON , TRANSNO, DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, '�Ϸ�' GBN, CLIENTCODE"
                            strFormet = "select {0}  "
                            strFormet = strFormet & " FROM MD_ELEC_TRANS_DTL  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND CASE ISNULL(SPONSOR,'') WHEN '' THEN 'N' WHEN 'N' THEN 'N' ELSE 'Y' END = 'N' "

                            strFormet = strFormet & " GROUP BY CLIENTCODE , TRANSYEARMON, TRANSNO "
                            strFormet = strFormet & " ORDER BY DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE)"
                        Else
                            strSelFields = "YEARMON , CLIENTCODE, DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, '�̿Ϸ�' GBN"
                            strFormet = "select {0}    "
                            strFormet = strFormet & " FROM MD_ELECTRIC_MEDIUM  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND CASE ISNULL(SPONSOR,'') WHEN '' THEN 'N' WHEN 'N' THEN 'N' ELSE 'Y' END = 'N'"
                            strFormet = strFormet & " AND isnull(tru_trans_no, '') = ''  "
                            strFormet = strFormet & " GROUP BY CLIENTCODE , YEARMON "
                            strFormet = strFormet & " ORDER BY "
                            strFormet = strFormet & " DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE)"
                        End If
                    ElseIf strFlag = "commi" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            strSelFields = "TRANSYEARMON , TRANSNO, DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) REAL_MED_NAME, '�Ϸ�' GBN, REAL_MED_CODE"
                            strFormet = "select {0}  "
                            strFormet = strFormet & " FROM MD_ELEC_TRANS_DTL  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND CASE ISNULL(SPONSOR,'') WHEN '' THEN 'N' WHEN 'N' THEN 'N' ELSE 'Y' END = 'N'"
                            strFormet = strFormet & " GROUP BY REAL_MED_CODE , TRANSYEARMON, TRANSNO "
                            strFormet = strFormet & " ORDER BY DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) "
                        Else
                            strSelFields = "YEARMON , REAL_MED_CODE, DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) REAL_MED_NAME, '�̿Ϸ�' GBN"
                            strFormet = "select {0}    "
                            strFormet = strFormet & " FROM MD_ELECTRIC_MEDIUM  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND CASE ISNULL(SPONSOR,'') WHEN '' THEN 'N' WHEN 'N' THEN 'N' ELSE 'Y' END = 'N'"
                            strFormet = strFormet & " AND isnull(commi_trans_no, '') = ''  "
                            strFormet = strFormet & " GROUP BY REAL_MED_CODE , YEARMON "
                            strFormet = strFormet & " ORDER BY "
                            strFormet = strFormet & " DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE)"
                        End If
                    End If
                End If
                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetTRANSCUSTNO")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function


#End Region

#Region "38. CATV ���������� ������ ��ȸ"

    ' =============== SelectRtnSample Code
    Public Function GetCATVSUSCUSTNO(ByVal strInfoXML As String, _
                              ByRef intRowCnt As Integer, _
                              ByRef intColCnt As Integer, _
                              ByVal strYEARMON As String, _
                              ByVal strCLIENTCODE As String, _
                              ByVal strCLIENTNAME As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim Con1, Con2, Con3 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""
                Con3 = ""

                If strYEARMON <> "" Then Con1 = String.Format(" AND (YEARMON = '{0}')", strYEARMON)
                If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (CLIENTCODE LIKE '%{0}%')", strCLIENTCODE)
                If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (MD_GET_CUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCLIENTNAME)

                strWhere = BuildFields(" ", Con1, Con2, Con3)

                strSelFields = "YEARMON , CLIENTCODE, MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME"

                strFormet = "select {0} FROM MD_CATV_MEDIUM  WHERE (1=1) {1}  GROUP BY CLIENTCODE , YEARMON"
                'select {0} FROM MD_BOOKING_MEDIUM A, SC_CUST_TEMP B  WHERE (A.CLIENTCODE = B.CUSTCODE)


                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetCATVSUSCUSTNO")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "06. ���ݰ�꼭 POP��ȸ"
    Public Function GetTAXNO(ByVal strInfoXML As String, _
                             ByRef intRowCnt As Integer, _
                             ByRef intColCnt As Integer, _
                             ByVal strTRANSYEARMON As String, _
                             ByVal strTRANSNO As String, _
                             ByVal strCLIENTCODE As String, _
                             ByVal strFlag As String, _
                             ByVal strTBL_Flag As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2, Con3, Con4 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""
                Con3 = ""
                Con4 = ""

                If strTRANSYEARMON <> "" Then Con1 = String.Format(" AND (TAXYEARMON = '{0}')", strTRANSYEARMON)
                If strTRANSNO <> "" Then Con2 = String.Format(" AND (TAXNO = '{0}')", strTRANSNO)
                If strFlag = "trans" Then
                    If strCLIENTCODE <> "" Then Con3 = String.Format(" AND (CLIENTCODE = '{0}')", strCLIENTCODE)
                ElseIf strFlag = "commi" Then
                    If strCLIENTCODE <> "" Then Con3 = String.Format(" AND (REAL_MED_CODE = '{0}')", strCLIENTCODE)
                End If
                If strTBL_Flag <> "" Then Con4 = String.Format(" AND (MEDFLAG = '{0}')", strTBL_Flag)

                strWhere = BuildFields(" ", Con1, Con2, Con3, Con4)

                If strFlag = "trans" Then
                    strSelFields = "TAXYEARMON , TAXNO , CLIENTCODE, MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME"
                    strFormet = "select {0} from MD_TRUTAX_HDR where 1=1 {1}"
                ElseIf strFlag = "commi" Then
                    strSelFields = "TAXYEARMON , TAXNO , REAL_MED_CODE, MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) REAL_MED_NAME"
                    strFormet = "select {0} from MD_COMMITAX_HDR where 1=1 {1}"
                End If


                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetTAXNO")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    Public Function GetTAXGUNNO(ByVal strInfoXML As String, _
                                ByRef intRowCnt As Integer, _
                                ByRef intColCnt As Integer, _
                                ByVal strTAXYEARMON As String, _
                                ByVal strTAXNO As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2, Con3, Con4 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""


                If strTAXYEARMON <> "" Then Con1 = String.Format(" AND (TAXYEARMON = '{0}')", strTAXYEARMON)
                If strTAXNO <> "" Then Con2 = String.Format(" AND (TAXNO = '{0}')", strTAXNO)


                strWhere = BuildFields(" ", Con1, Con2)


                strSelFields = "TAXYEARMON , TAXNO , CLIENTCODE, DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, SUMM"
                strFormet = "select {0} from PD_TAX_HDR where 1=1 {1}"



                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetTAXGUNNO")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "39. ��ü�� ��ȸ"

    ' =============== SelectRtnSample Code
    Public Function GetMEDNO(ByVal strInfoXML As String, _
                             ByRef intRowCnt As Integer, _
                             ByRef intColCnt As Integer, _
                             ByVal strCUSTCODE As String, _
                             ByVal strCUSTNAME As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""

                If strCUSTCODE <> "" Then Con1 = String.Format(" AND (CUSTCODE LIKE '%{0}%')", strCUSTCODE)
                If strCUSTNAME <> "" Then Con2 = String.Format(" AND (CUSTNAME LIKE '%{0}%')", strCUSTNAME)
                strWhere = BuildFields(" ", Con1, Con2)

                strSelFields = "CUSTCODE , CUSTNAME , BUSINO, COMPANYNAME, MPP, DBO.MD_GET_CUSTNAME_FUN(MPP) MPP_NAME"

                strFormet = "select {0} from SC_CUST_TEMP where 1=1 AND CUSTCODE LIKE 'B%'  AND ATTR10 = 1 {1} ORDER BY  CASE SUBSTRING(LTRIM(CUSTNAME),1,3) WHEN '(��)' THEN LTRIM(SUBSTRING(CUSTNAME,4,LEN(CUSTNAME))) ELSE LTRIM(CUSTNAME) END"


                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetMEDNO")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "40. û���� ��ȸ"

    ' =============== SelectRtnSample Code
    Public Function GetREALMEDNO(ByVal strInfoXML As String, _
                                 ByRef intRowCnt As Integer, _
                                 ByRef intColCnt As Integer, _
                                 ByVal strCUSTCODE As String, _
                                 ByVal strCUSTNAME As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""

                If strCUSTCODE <> "" Then Con1 = String.Format(" AND (HIGHCUSTCODE LIKE '%{0}%')", strCUSTCODE)
                If strCUSTNAME <> "" Then Con2 = String.Format(" AND (COMPANYNAME LIKE '%{0}%')", strCUSTNAME)
                strWhere = BuildFields(" ", Con1, Con2)

                strSelFields = "DISTINCT CASE ISNULL(HIGHCUSTCODE,'') WHEN '' THEN CUSTCODE ELSE HIGHCUSTCODE END HIGHCUSTCODE, CASE ISNULL(HIGHCUSTCODE,'') WHEN '' THEN CUSTNAME ELSE COMPANYNAME END HIGHCUSTNAME,BUSINO"

                strFormet = "select {0} from SC_CUST_TEMP where 1=1 AND MEDFLAG = 'B'  AND ATTR10=1  {1} "

                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetREALMEDNO")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "41. û���� ��ȸ"

    ' =============== SelectRtnSample Code
    Public Function GetREALMEDNO1(ByVal strInfoXML As String, _
                                  ByRef intRowCnt As Integer, _
                                  ByRef intColCnt As Integer, _
                                  ByVal strCUSTCODE As String, _
                                  ByVal strCUSTNAME As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""

                If strCUSTCODE <> "" Then Con1 = String.Format(" AND (CUSTCODE LIKE '%{0}%')", strCUSTCODE)
                If strCUSTNAME <> "" Then Con2 = String.Format(" AND (CUSTNAME LIKE '%{0}%')", strCUSTNAME)
                strWhere = BuildFields(" ", Con1, Con2)

                strSelFields = "DISTINCT CASE ISNULL(HIGHCUSTCODE,'') WHEN '' THEN CUSTCODE ELSE HIGHCUSTCODE END HIGHCUSTCODE, CASE ISNULL(HIGHCUSTCODE,'') WHEN '' THEN CUSTNAME ELSE COMPANYNAME END HIGHCUSTNAME,BUSINO"

                strFormet = "select {0} from SC_CUST_TEMP where 1=1 {1}  "


                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetREALMEDNO1")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "42. �귣�� ��ȸ"
    ' =============== SelectRtnSample Code
    Public Function GetBRANDNO(ByVal strInfoXML As String, _
                             ByRef intRowCnt As Integer, _
                             ByRef intColCnt As Integer, _
                             ByVal strBRANDCODE As String, _
                             ByVal strBRANDNAME As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2, Con3 As String
        Dim vntData As Object

        Dim strCUSTCODE, strSEQ
        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""
                Con3 = ""

                If strBRANDCODE <> "" Then
                    If strBRANDCODE.Length = 10 Then
                        strCUSTCODE = Mid(strBRANDCODE, 1, 6)
                        strSEQ = Mid(strBRANDCODE, 7, 4)
                        Con1 = String.Format(" AND (CUSTCODE = '{0}')", strCUSTCODE)
                        Con2 = String.Format(" AND (SEQ = '{0}')", strSEQ)
                    Else
                        strCUSTCODE = strBRANDCODE
                        strSEQ = ""
                        Con1 = String.Format(" AND (CUSTCODE LIKE '%{0}%')", strCUSTCODE)
                        Con2 = ""
                    End If
                End If

                If strBRANDNAME <> "" Then Con3 = String.Format(" AND (BRANDNAME LIKE '%{0}%')", strBRANDNAME)
                strWhere = BuildFields(" ", Con1, Con2, Con3)

                strSelFields = "CUSTCODE || SEQ , BRANDNAME"

                strFormet = "select {0} from SC_BRANDCODE where 1=1 {1} "


                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetBRANDNO")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "43. �귣�� ��ȸ"
    ' =============== SelectRtnSample Code
    Public Function GetBRANDNO1(ByVal strInfoXML As String, _
                             ByRef intRowCnt As Integer, _
                             ByRef intColCnt As Integer, _
                             ByVal strBRANDCODE As String, _
                             ByVal strBRANDNAME As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2, Con3 As String
        Dim vntData As Object

        Dim strCUSTCODE, strSEQ
        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""
                Con3 = ""

                If strBRANDCODE <> "" Then
                    If strBRANDCODE.Length = 10 Then
                        strCUSTCODE = Mid(strBRANDCODE, 1, 6)
                        strSEQ = Mid(strBRANDCODE, 7, 4)
                        Con1 = String.Format(" AND (CUSTCODE = '{0}')", strCUSTCODE)
                        Con2 = String.Format(" AND (SEQ = '{0}')", strSEQ)
                    Else
                        strCUSTCODE = strBRANDCODE
                        strSEQ = ""
                        Con1 = String.Format(" AND (CUSTCODE LIKE '%{0}%')", strCUSTCODE)
                        Con2 = ""
                    End If
                End If

                If strBRANDNAME <> "" Then Con3 = String.Format(" AND (BRANDNAME LIKE '%{0}%')", strBRANDNAME)
                strWhere = BuildFields(" ", Con1, Con2, Con3)

                strSelFields = "DEPT_CD, SC_DEPT_NAME_FUN(DEPT_CD) DEPT_NAME"

                strFormet = "select {0} from SC_BRANDCODE where 1=1 {1} "


                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetBRANDNO")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "41. �귣�庰 ���μ� ��ȸ"

    ' =============== SelectRtnSample Code
    Public Function GetBrandAndDept(ByVal strInfoXML As String, _
                                    ByRef intRowCnt As Integer, _
                                    ByRef intColCnt As Integer, _
                                    ByRef strCUSTCODE As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""

                If strCUSTCODE = "" Then Exit Function

                'strSelFields = "DECODE(SC_CUSTCODE_SEQCODE_FUN('" & strCUSTCODE & "'),'ERROR', '',SC_CUSTCODE_SEQCODE_FUN('" & strCUSTCODE & "')) , "
                'strSelFields = strSelFields & "PD_JOBCUST_NAME_FUN (DECODE(SC_CUSTCODE_SEQCODE_FUN('" & strCUSTCODE & "'),'ERROR', '',SC_CUSTCODE_SEQCODE_FUN('" & strCUSTCODE & "'))) , "
                'strSelFields = strSelFields & "SC_SEQCODE_DEPTCD_FUN(DECODE(SC_CUSTCODE_SEQCODE_FUN('" & strCUSTCODE & "'),'ERROR','',SC_CUSTCODE_SEQCODE_FUN('" & strCUSTCODE & "'))), "
                'strSelFields = strSelFields & "SC_DEPT_NAME_FUN (SC_SEQCODE_DEPTCD_FUN(DECODE(SC_CUSTCODE_SEQCODE_FUN('" & strCUSTCODE & "'),'ERROR','',SC_CUSTCODE_SEQCODE_FUN('" & strCUSTCODE & "')))), "
                'strSelFields = strSelFields & "MD_GET_BRANDNAME_FUN(DECODE(SC_CUSTCODE_SEQCODE_FUN('" & strCUSTCODE & "'),'ERROR', '',SC_CUSTCODE_SEQCODE_FUN('" & strCUSTCODE & "'))) "
                strSelFields = " CASE DBO.PD_JOBCUST_NAME_FUN(DBO.SC_CUSTCODE_SEQCODE_FUN('" & strCUSTCODE & "')) WHEN 'error' THEN ''"
                strSelFields = strSelFields & "ELSE DBO.PD_JOBCUST_NAME_FUN(DBO.SC_CUSTCODE_SEQCODE_FUN('" & strCUSTCODE & "')) END,"
                strSelFields = strSelFields & "CASE DBO.PD_JOBCUST_NAME_FUN(DBO.SC_CUSTCODE_SEQCODE_FUN('" & strCUSTCODE & "')) WHEN 'error' THEN ''		    "
                strSelFields = strSelFields & "ELSE DBO.PD_JOBCUST_NAME_FUN(DBO.SC_CUSTCODE_SEQCODE_FUN('" & strCUSTCODE & "')) END,"
                strSelFields = strSelFields & "CASE DBO.SC_DEPT_NAME_FUN (DBO.SC_SEQCODE_DEPTCD_FUN(DBO.SC_CUSTCODE_SEQCODE_FUN('" & strCUSTCODE & "'))) WHEN 'error' THEN ''                                        "
                strSelFields = strSelFields & "ELSE DBO.SC_DEPT_NAME_FUN (DBO.SC_SEQCODE_DEPTCD_FUN(DBO.SC_CUSTCODE_SEQCODE_FUN('" & strCUSTCODE & "'))) END,"
                strSelFields = strSelFields & "CASE DBO.MD_GET_BRANDNAME_FUN(DBO.SC_CUSTCODE_SEQCODE_FUN('" & strCUSTCODE & "')) WHEN 'error' THEN '' "
                strSelFields = strSelFields & "ELSE DBO.MD_GET_BRANDNAME_FUN(DBO.SC_CUSTCODE_SEQCODE_FUN('" & strCUSTCODE & "')) END "
                strFormet = "select {0} where 1=1 {1}  "


                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetBrandAndDept")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "45. �׷챤�� ��ȸ"

    ' =============== SelectRtnSample Code
    Public Function GetGROUPCUST(ByVal strInfoXML As String, _
                                 ByRef intRowCnt As Integer, _
                                 ByRef intColCnt As Integer, _
                                 ByVal strYEARMON As String, _
                                 ByVal strTBLFLAG As String) As Object

        Dim strSQL As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                If strTBLFLAG = "PRINT" Then
                    strSQL = " select CLIENTCODE, MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, SUM(AMOUNT) AMT"
                    strSQL = strSQL & " from MD_BOOKING_MEDIUM where 1=1 "
                    strSQL = strSQL & " AND GFLAG = 'J'"
                    strSQL = strSQL & " AND CLIENTCODE <> 'A00000' "
                    strSQL = strSQL & " AND ATTR01 = 'G' "
                    strSQL = strSQL & " AND YEARMON = '" & strYEARMON & "'"
                    strSQL = strSQL & " GROUP BY CLIENTCODE"
                    strSQL = strSQL & " ORDER BY CLIENTCODE"
                ElseIf strTBLFLAG = "CATV" Then
                    strSQL = " select CLIENTCODE, MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, SUM(AMT) AMT"
                    strSQL = strSQL & " from MD_CATV_MEDIUM where 1=1 "
                    strSQL = strSQL & " AND CLIENTCODE <> 'A00000' "
                    strSQL = strSQL & " AND ATTR01 = 'G' "
                    strSQL = strSQL & " AND YEARMON = '" & strYEARMON & "'"
                    strSQL = strSQL & " GROUP BY CLIENTCODE"
                    strSQL = strSQL & " ORDER BY CLIENTCODE"
                ElseIf strTBLFLAG = "INTERNET" Then
                    strSQL = " select CLIENTCODE, MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, SUM(AMT) AMT"
                    strSQL = strSQL & " from MD_INTERNET_MEDIUM where 1=1 "
                    strSQL = strSQL & " AND GFLAG = 'J'"
                    strSQL = strSQL & " AND CLIENTCODE <> 'A00000' "
                    strSQL = strSQL & " AND ATTR01 = 'G' "
                    strSQL = strSQL & " AND YEARMON = '" & strYEARMON & "'"
                    strSQL = strSQL & " GROUP BY CLIENTCODE"
                    strSQL = strSQL & " ORDER BY CLIENTCODE"
                ElseIf strTBLFLAG = "ELEC" Then
                    strSQL = " select CLIENTCODE, MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, SUM(AMT) AMT"
                    strSQL = strSQL & " from MD_ELECTRIC_MEDIUM where 1=1 "
                    strSQL = strSQL & " AND CLIENTCODE <> 'A00000' "
                    strSQL = strSQL & " AND ATTR01 = 'G' "
                    strSQL = strSQL & " AND YEARMON = '" & strYEARMON & "'"
                    strSQL = strSQL & " GROUP BY CLIENTCODE"
                    strSQL = strSQL & " ORDER BY CLIENTCODE"
                End If

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetGROUPCUST")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "46. �����ڿ� ���� ������ ��ȸ"
    ' =============== SelectRtnSample Code
    Public Function GetCUSTLISTNO_GBN(ByVal strInfoXML As String, _
                                  ByRef intRowCnt As Integer, _
                                  ByRef intColCnt As Integer, _
                                  ByVal strCUSTCODE As String, _
                                  ByVal strCUSTNAME As String, _
                                  ByVal strGUBUN As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2, Con3 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""
                Con3 = ""
                If strCUSTCODE <> "" Then Con1 = String.Format(" AND (CUSTCODE LIKE '%{0}%')", strCUSTCODE)
                If strCUSTNAME <> "" Then Con2 = String.Format(" AND (CUSTNAME LIKE '%{0}%')", strCUSTNAME)
                If strGUBUN <> "" Then Con3 = String.Format(" AND (MEDFLAG  = '{0}')", strGUBUN)
                strWhere = BuildFields(" ", Con1, Con2, Con3)

                strSelFields = "CUSTCODE , CUSTNAME , COMPANYNAME"

                strFormet = "select {0} from SC_CUST_TEMP where 1=1 {1} "


                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetCUSTLISTNO_GBN")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "47. ����ڵ�Ϲ�ȣ ȸ��ŷ�ó ��ȸ"
    ' =============== SelectRtnSample Code
    Public Function GetREGNO(ByVal strInfoXML As String, _
                             ByRef intRowCnt As Integer, _
                             ByRef intColCnt As Integer, _
                             ByVal strACCCODE As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����
                strSelFields = "CUST_ID,CUST_NAME,REG_NUM, CEO, BIZ_CAT, BIZ_TYPE, POST_CODE, ADDR1, ADDR2, TEL_NUM, FAX_NUM"

                strFormet = "select {0} from SC_CUST_DTL where REG_NUM = '" & strACCCODE & "' "


                strSQL = String.Format(strFormet, strSelFields)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetREGNO")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "48. ����ڵ�Ϲ�ȣ �ŷ�ó ��ȸ"
    ' =============== SelectRtnSample Code
    Public Function GetCUSTREGNO(ByVal strInfoXML As String, _
                             ByRef intRowCnt As Integer, _
                             ByRef intColCnt As Integer, _
                             ByVal strACCCODE As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����


                strSQL = " SELECT"
                strSQL = strSQL & " CUSTCODE,CUSTNAME, "
                strSQL = strSQL & " CASE MEDFLAG WHEN 'A' THEN '������'"
                strSQL = strSQL & "              WHEN 'B' THEN '��ü��'"
                strSQL = strSQL & "              WHEN 'C' THEN '��Ƽ��ȭ��'"
                strSQL = strSQL & "	             WHEN 'D' THEN '����ũ��'"
                strSQL = strSQL & "              WHEN 'E' THEN 'ȿ����'"
                strSQL = strSQL & "              WHEN 'F' THEN '�Կ���'"
                strSQL = strSQL & "              WHEN 'H' THEN '��'"
                strSQL = strSQL & "              WHEN 'I' THEN '�Ϸ���Ʈ'"
                strSQL = strSQL & "              WHEN 'L' THEN '�����'"
                strSQL = strSQL & "              WHEN 'M' THEN '����ó'"
                strSQL = strSQL & "              WHEN 'P' THEN '�ǽ�Ʈ'"
                strSQL = strSQL & "              WHEN 'S' THEN '��Ʈ��'"
                strSQL = strSQL & "              WHEN 'V' THEN '�����'"
                strSQL = strSQL & "              WHEN 'G' THEN '������'"
                strSQL = strSQL & "              WHEN 'Z' THEN '��Ÿ' END AS MEDFLAG"
                strSQL = strSQL & " FROM SC_CUST_TEMP"
                strSQL = strSQL & " WHERE BUSINO = '" & strACCCODE & "'"




                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetCUSTREGNO")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "50. �ŷ����������� POP��ȸ"
    Public Function GetTAXCUSTNO(ByVal strInfoXML As String, _
                                  ByRef intRowCnt As Integer, _
                                  ByRef intColCnt As Integer, _
                                  ByVal strYEARMON As String, _
                                  ByVal strCLIENTCODE As String, _
                                  ByVal strCLIENTNAME As String, _
                                  ByVal strFlag As String, _
                                  ByVal strTBL_Flag As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2, Con3, Con4 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""
                Con3 = ""
                Con4 = ""

                If strTBL_Flag = "PRINT" Then
                    If strFlag = "trans" Then
                        If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
                        If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (CLIENTCODE LIKE '%{0}%')", strCLIENTCODE)
                        If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCLIENTNAME)

                        'Con4 = " HAVING DECODE(SUM(DECODE(TRU_TRANS_NO,NULL,1,0)),0,'�Ϸ�','�̿Ϸ�') = '�Ϸ�' "
                        Con4 = "HAVING (CASE WHEN SUM(CASE TRU_TRANS_NO WHEN '' THEN 1 ELSE 0 END) WHEN 0 THEN '�Ϸ�' ELSE '�̿Ϸ�' END) = '�Ϸ�'"

                    ElseIf strFlag = "commi" Then
                        If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
                        If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (REAL_MED_CODE LIKE '%{0}%')", strCLIENTCODE)
                        If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) LIKE '%{0}%')", strCLIENTNAME)

                        Con4 = "HAVING (CASE WHEN SUM(CASE COMMI_TRANS_NO WHEN '' THEN 1 ELSE 0 END) WHEN 0 THEN '�Ϸ�' ELSE '�̿Ϸ�' END) = '�Ϸ�'"

                    End If
                ElseIf strTBL_Flag = "CATV" Then
                    If strFlag = "trans" Then
                        If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
                        If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (CLIENTCODE LIKE '%{0}%')", strCLIENTCODE)
                        If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCLIENTNAME)

                        Con4 = "HAVING (CASE WHEN SUM(CASE TRU_TRANS_NO WHEN '' THEN 1 ELSE 0 END) WHEN 0 THEN '�Ϸ�' ELSE '�̿Ϸ�' END) = '�Ϸ�'"

                    ElseIf strFlag = "commi" Then
                        If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
                        If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (REAL_MED_CODE LIKE '%{0}%')", strCLIENTCODE)
                        If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) LIKE '%{0}%')", strCLIENTNAME)

                        Con4 = "HAVING (CASE WHEN SUM(CASE COMMI_TRANS_NO WHEN '' THEN 1 ELSE 0 END) WHEN 0 THEN '�Ϸ�' ELSE '�̿Ϸ�' END) = '�Ϸ�'"

                    End If

                ElseIf strTBL_Flag = "INTERNET" Then
                    If strFlag = "trans" Then
                        If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
                        If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (CLIENTCODE LIKE '%{0}%')", strCLIENTCODE)
                        If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCLIENTNAME)

                        Con4 = "HAVING (CASE WHEN SUM(CASE TRU_TRANS_NO WHEN '' THEN 1 ELSE 0 END) WHEN 0 THEN '�Ϸ�' ELSE '�̿Ϸ�' END) = '�Ϸ�'"

                    ElseIf strFlag = "commi" Then
                        If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
                        If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (REAL_MED_CODE LIKE '%{0}%')", strCLIENTCODE)
                        If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) LIKE '%{0}%')", strCLIENTNAME)

                        Con4 = "HAVING (CASE WHEN SUM(CASE COMMI_TRANS_NO WHEN '' THEN 1 ELSE 0 END) WHEN 0 THEN '�Ϸ�' ELSE '�̿Ϸ�' END) = '�Ϸ�'"

                    End If
                ElseIf strTBL_Flag = "OUTDOOR" Then
                    If strFlag = "trans" Then
                        If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
                        If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (CLIENTCODE LIKE '%{0}%')", strCLIENTCODE)
                        If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCLIENTNAME)

                        Con4 = "HAVING (CASE WHEN SUM(CASE TRU_TRANS_NO WHEN '' THEN 1 ELSE 0 END) WHEN 0 THEN '�Ϸ�' ELSE '�̿Ϸ�' END) = '�Ϸ�'"

                    ElseIf strFlag = "commi" Then
                        If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
                        If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (REAL_MED_CODE LIKE '%{0}%')", strCLIENTCODE)
                        If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) LIKE '%{0}%')", strCLIENTNAME)

                        Con4 = "HAVING (CASE WHEN SUM(CASE COMMI_TRANS_NO WHEN '' THEN 1 ELSE 0 END) WHEN 0 THEN '�Ϸ�' ELSE '�̿Ϸ�' END) = '�Ϸ�'"

                    End If
                ElseIf strTBL_Flag = "PROMOTION" Then
                    If strFlag = "trans" Then
                        If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
                        If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (CLIENTCODE LIKE '%{0}%')", strCLIENTCODE)
                        If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCLIENTNAME)

                        Con4 = "HAVING (CASE WHEN SUM(CASE TRU_TRANS_NO WHEN '' THEN 1 ELSE 0 END) WHEN 0 THEN '�Ϸ�' ELSE '�̿Ϸ�' END) = '�Ϸ�'"

                    ElseIf strFlag = "commi" Then
                        If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
                        If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (REAL_MED_CODE LIKE '%{0}%')", strCLIENTCODE)
                        If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) LIKE '%{0}%')", strCLIENTNAME)

                        Con4 = "HAVING (CASE WHEN SUM(CASE COMMI_TRANS_NO WHEN '' THEN 1 ELSE 0 END) WHEN 0 THEN '�Ϸ�' ELSE '�̿Ϸ�' END) = '�Ϸ�'"

                    End If
                ElseIf strTBL_Flag = "ELEC" Then
                    If strFlag = "trans" Then
                        If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
                        If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (CLIENTCODE LIKE '%{0}%')", strCLIENTCODE)
                        If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCLIENTNAME)

                        Con4 = "HAVING (CASE WHEN SUM(CASE TRU_TRANS_NO WHEN '' THEN 1 ELSE 0 END) WHEN 0 THEN '�Ϸ�' ELSE '�̿Ϸ�' END) = '�Ϸ�'"

                    ElseIf strFlag = "commi" Then
                        If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
                        If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (REAL_MED_CODE LIKE '%{0}%')", strCLIENTCODE)
                        If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) LIKE '%{0}%')", strCLIENTNAME)

                        Con4 = "HAVING (CASE WHEN SUM(CASE COMMI_TRANS_NO WHEN '' THEN 1 ELSE 0 END) WHEN 0 THEN '�Ϸ�' ELSE '�̿Ϸ�' END) = '�Ϸ�'"

                    End If
                ElseIf strTBL_Flag = "ELECSPON" Then
                    If strFlag = "trans" Then
                        If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
                        If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (CLIENTCODE LIKE '%{0}%')", strCLIENTCODE)
                        If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.MD_GET_CUSTNAME_FUN(REAL_MED_CODE) LIKE '%{0}%')", strCLIENTNAME)

                        Con4 = "HAVING (CASE WHEN SUM(CASE TRU_TRANS_NO WHEN '' THEN 1 ELSE 0 END) WHEN 0 THEN '�Ϸ�' ELSE '�̿Ϸ�' END) = '�Ϸ�'"

                    ElseIf strFlag = "commi" Then
                        If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
                        If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (REAL_MED_CODE LIKE '%{0}%')", strCLIENTCODE)
                        If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) LIKE '%{0}%')", strCLIENTNAME)

                        Con4 = "HAVING (CASE WHEN SUM(CASE COMMI_TRANS_NO WHEN '' THEN 1 ELSE 0 END) WHEN 0 THEN '�Ϸ�' ELSE '�̿Ϸ�' END) = '�Ϸ�'"

                    End If
                ElseIf strTBL_Flag = "ETC" Then
                    If strFlag = "trans" Then
                        If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
                        If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (CLIENTCODE LIKE '%{0}%')", strCLIENTCODE)
                        If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.MD_GET_CUSTNAME_FUN(REAL_MED_CODE) LIKE '%{0}%')", strCLIENTNAME)

                        Con4 = "HAVING (CASE WHEN SUM(CASE TRU_TRANS_NO WHEN '' THEN 1 ELSE 0 END) WHEN 0 THEN '�Ϸ�' ELSE '�̿Ϸ�' END) = '�Ϸ�'"

                    ElseIf strFlag = "commi" Then
                        If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
                        If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (REAL_MED_CODE LIKE '%{0}%')", strCLIENTCODE)
                        If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) LIKE '%{0}%')", strCLIENTNAME)

                        Con4 = "HAVING (CASE WHEN SUM(CASE COMMI_TRANS_NO WHEN '' THEN 1 ELSE 0 END) WHEN 0 THEN '�Ϸ�' ELSE '�̿Ϸ�' END) = '�Ϸ�'"

                    End If
                End If
                strWhere = BuildFields(" ", Con1, Con2, Con3)


                If strTBL_Flag = "PRINT" Then
                    If strFlag = "trans" Then
                        strSelFields = "TRANSYEARMON , CLIENTCODE, DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, DBO.SC_CUSTCODE_BISINO_FUN(CLIENTCODE) GBN"
                        strFormet = "select {0} FROM MD_PRINTTRANS_HDR  WHERE 1=1  {1} GROUP BY CLIENTCODE , TRANSYEARMON  ORDER BY  DBO.SC_CUSTCODE_BISINO_FUN(CLIENTCODE), DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE)"
                    ElseIf strFlag = "commi" Then
                        strSelFields = "TRANSYEARMON, REAL_MED_CODE, DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) REAL_MED_NAME,  DBO.SC_CUSTCODE_BISINO_FUN(REAL_MED_CODE) GBN"
                        strFormet = "select {0} FROM MD_PRINTCOMMI_HDR  WHERE 1=1  {1} GROUP BY REAL_MED_CODE ,TRANSYEARMON  "
                    End If
                ElseIf strTBL_Flag = "CATV" Then
                    If strFlag = "trans" Then
                        strSelFields = "TRANSYEARMON , CLIENTCODE, DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME,  DBO.SC_CUSTCODE_BISINO_FUN(CLIENTCODE) GBN"
                        strFormet = "select {0} FROM MD_CATVTRANS_HDR  WHERE (1=1)  {1} GROUP BY CLIENTCODE ,TRANSYEARMON  ORDER BY  DBO.SC_CUSTCODE_BISINO_FUN(CLIENTCODE), DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE)"
                    ElseIf strFlag = "commi" Then
                        strSelFields = "TRANSYEARMON, REAL_MED_CODE CLIENTCODE, DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) CLIENTNAME,  DBO.SC_CUSTCODE_BISINO_FUN(REAL_MED_CODE) GBN"
                        strFormet = "select {0} FROM MD_CATVCOMMI_HDR  WHERE 1=1  {1} GROUP BY REAL_MED_CODE ,TRANSYEARMON  "
                    End If
                ElseIf strTBL_Flag = "INTERNET" Then
                    If strFlag = "trans" Then
                        strSelFields = "TRANSYEARMON , CLIENTCODE, DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME,  DBO.SC_CUSTCODE_BISINO_FUN(CLIENTCODE) GBN"
                        strFormet = "select {0} FROM MD_INTERNETTRANS_HDR  WHERE (1=1)  {1} GROUP BY CLIENTCODE ,TRANSYEARMON  ORDER BY  DBO.SC_CUSTCODE_BISINO_FUN(CLIENTCODE), DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE)"
                    ElseIf strFlag = "commi" Then
                        strSelFields = "TRANSYEARMON, REAL_MED_CODE, DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) REAL_MED_NAME, DBO.SC_CUSTCODE_BISINO_FUN(REAL_MED_CODE) GBN"
                        strFormet = "select {0} FROM MD_INTERNETCOMMI_HDR  WHERE (1=1)  {1} GROUP BY REAL_MED_CODE ,TRANSYEARMON  "
                    End If
                ElseIf strTBL_Flag = "PROMOTION" Then
                    If strFlag = "trans" Then
                        strSelFields = "TRANSYEARMON , CLIENTCODE, DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME,  DBO.SC_CUSTCODE_BISINO_FUN(CLIENTCODE) GBN"
                        strFormet = "select {0} FROM MD_PROMOTIONTRANS_HDR  WHERE (1=1)  {1} GROUP BY CLIENTCODE ,TRANSYEARMON  ORDER BY  DBO.SC_CUSTCODE_BISINO_FUN(CLIENTCODE), DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE)"
                    ElseIf strFlag = "commi" Then
                        strSelFields = "TRANSYEARMON, REAL_MED_CODE, DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) REAL_MED_NAME, DBO.SC_CUSTCODE_BISINO_FUN(REAL_MED_CODE) GBN"
                        strFormet = "select {0} FROM MD_PROMOTIONCOMMI_HDR  WHERE (1=1)  {1} GROUP BY REAL_MED_CODE ,TRANSYEARMON  "
                    End If
                ElseIf strTBL_Flag = "OUTDOOR" Then
                    If strFlag = "trans" Then
                        strSelFields = "TRANSYEARMON , CLIENTCODE, DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME,  DBO.SC_CUSTCODE_BISINO_FUN(CLIENTCODE) GBN"
                        strFormet = "select {0} FROM MD_OUTDOORTRANS_HDR  WHERE (1=1)  {1} GROUP BY CLIENTCODE ,TRANSYEARMON  ORDER BY  DBO.SC_CUSTCODE_BISINO_FUN(CLIENTCODE), DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE)"
                    ElseIf strFlag = "commi" Then
                        strSelFields = "TRANSYEARMON, REAL_MED_CODE, DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) REAL_MED_NAME, DBO.SC_CUSTCODE_BISINO_FUN(REAL_MED_CODE) GBN"
                        strFormet = "select {0} FROM MD_OUTDOORCOMMI_HDR  WHERE (1=1)  {1} GROUP BY REAL_MED_CODE ,TRANSYEARMON  "
                    End If
                ElseIf strTBL_Flag = "ELEC" Then
                    If strFlag = "trans" Then
                        strSelFields = "TRANSYEARMON , CLIENTCODE, DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, DBO.SC_CUSTCODE_BISINO_FUN(CLIENTCODE) GBN"
                        strFormet = "select {0} FROM MD_ELEC_TRANS_HDR  WHERE 1=1  AND (ATTR03 <> 'Y' or ATTR03 IS NULL) {1} GROUP BY CLIENTCODE , TRANSYEARMON   ORDER BY  DBO.SC_CUSTCODE_BISINO_FUN(CLIENTCODE), DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE)"
                    ElseIf strFlag = "commi" Then
                        strSelFields = "TRANSYEARMON, REAL_MED_CODE, DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) REAL_MED_NAME,  DBO.SC_CUSTCODE_BISINO_FUN(REAL_MED_CODE) GBN"
                        strFormet = "select {0} FROM MD_ELECCOMMI_HDR  WHERE 1=1  AND (ATTR03 <> 'Y' or ATTR03 IS NULL)  {1} GROUP BY REAL_MED_CODE ,TRANSYEARMON  "
                    End If
                ElseIf strTBL_Flag = "ELECSPON" Then
                    If strFlag = "trans" Then
                        strSelFields = "TRANSYEARMON , CLIENTCODE, DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME,  DBO.SC_CUSTCODE_BISINO_FUN(CLIENTCODE) GBN"
                        strFormet = "select {0} FROM MD_ELEC_TRANS_HDR  WHERE (1=1)  AND ATTR03 ='Y' {1} GROUP BY CLIENTCODE ,TRANSYEARMON  ORDER BY  DBO.SC_CUSTCODE_BISINO_FUN(CLIENTCODE), DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE)"
                    ElseIf strFlag = "commi" Then
                        strSelFields = "TRANSYEARMON, REAL_MED_CODE, DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) REAL_MED_NAME,  DBO.SC_CUSTCODE_BISINO_FUN(REAL_MED_CODE) GBN"
                        strFormet = "select {0} FROM MD_ELECCOMMI_HDR  WHERE (1=1)   AND ATTR03 ='Y' {1} GROUP BY REAL_MED_CODE ,TRANSYEARMON  "
                    End If
                ElseIf strTBL_Flag = "ETC" Then
                    If strFlag = "trans" Then
                        strSelFields = "TRANSYEARMON , CLIENTCODE, DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME,  DBO.SC_CUSTCODE_BISINO_FUN(CLIENTCODE) GBN"
                        strFormet = "select {0} FROM MD_ETCTRANS_HDR  WHERE (1=1)   {1} GROUP BY CLIENTCODE ,TRANSYEARMON  ORDER BY  DBO.SC_CUSTCODE_BISINO_FUN(CLIENTCODE), DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE)"
                    ElseIf strFlag = "commi" Then
                        strSelFields = "TRANSYEARMON, REAL_MED_CODE, DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) REAL_MED_NAME,  DBO.SC_CUSTCODE_BISINO_FUN(REAL_MED_CODE) GBN"
                        strFormet = "select {0} FROM MD_ETCCOMMI_HDR  WHERE (1=1)    {1} GROUP BY REAL_MED_CODE ,TRANSYEARMON  "
                    End If
                End If

                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetTRNASCUSTNO")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#End Region

#Region "getcustexelist"
    Public Function GetCUSTEXELIST(ByVal strInfoXML As String, _
                                     ByRef intRowCnt As Integer, _
                                     ByRef intColCnt As Integer, _
                                     ByVal strYEAR As String, _
                                     ByVal strCUSTNAME As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""

                If strYEAR <> "" Then Con1 = String.Format(" AND (SUBSTRING(YEARMON,1,4) = '{0}')", strYEAR)
                If strCUSTNAME <> "" Then Con2 = String.Format(" AND (DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) LIKE '%{0}%')", strCUSTNAME)
                strWhere = BuildFields(" ", Con1, Con2)

                strFormet = " SELECT"
                strFormet = strFormet & " 0 CHK,"
                strFormet = strFormet & " REAL_MED_CODE CUSTNO,"
                strFormet = strFormet & " DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) CUSTNAME"
                strFormet = strFormet & " FROM "
                strFormet = strFormet & " ("
                strFormet = strFormet & " SELECT "
                strFormet = strFormet & "  0 CHK,"
                strFormet = strFormet & "  REAL_MED_CODE,"
                strFormet = strFormet & "  DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) CUSTNAME"
                strFormet = strFormet & "  FROM MD_BOOKING_MEDIUM"
                strFormet = strFormet & "  WHERE GFLAG IN('J','S') {0}"
                strFormet = strFormet & "  GROUP BY REAL_MED_CODE"
                'strFormet = strFormet & "  UNION ALL"
                'strFormet = strFormet & "  SELECT "
                'strFormet = strFormet & "  0 CHK,"
                'strFormet = strFormet & "  CLIENTCODE,"
                'strFormet = strFormet & "  MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME"
                'strFormet = strFormet & "  FROM MD_INTERNET_MEDIUM"
                'strFormet = strFormet & "  WHERE GFLAG IN('J','S')   {0}"
                'strFormet = strFormet & "  GROUP BY CLIENTCODE"
                strFormet = strFormet & " ) A "
                strFormet = strFormet & " GROUP BY REAL_MED_CODE"
                strFormet = strFormet & " ORDER BY CASE SUBSTRING(LTRIM(DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE)),1,3) WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE),4,LEN(DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE)))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE),4,LEN(DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE)))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE),4,LEN(DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE)))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE),4,LEN(DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE)))) "
                strFormet = strFormet & " WHEN '(���' THEN LTRIM(SUBSTRING(DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE),5,LEN(DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE)))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE),4,LEN(DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE)))) "
                strFormet = strFormet & " ELSE LTRIM(DBO.MD_GET_REALMEDCODE_FUN(REAL_MED_CODE)) END "

                strSQL = String.Format(strFormet, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetCUSTEXELIST")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    ' =============== SelectRtnSample Code
    Public Function GetCUSTDBLLIST(ByVal strInfoXML As String, _
                                   ByRef intRowCnt As Integer, _
                                   ByRef intColCnt As Integer, _
                                   ByVal strYEAR As String, _
                                   ByVal strCUSTNAME As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""

                If strYEAR <> "" Then Con1 = String.Format(" AND (SUBSTRING(YEARMON,1,4) = '{0}')", strYEAR)
                If strCUSTNAME <> "" Then Con2 = String.Format(" AND (dbo.MD_GET_CUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCUSTNAME)
                strWhere = BuildFields(" ", Con1, Con2)

                strFormet = " SELECT "
                strFormet = strFormet & " 0 CHK,"
                strFormet = strFormet & " CLIENTCODE,"
                strFormet = strFormet & " dbo.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME"
                strFormet = strFormet & " FROM MD_BOOKING_MEDIUM"
                strFormet = strFormet & " WHERE GFLAG IN('J','S') {0}"
                strFormet = strFormet & " GROUP BY CLIENTCODE"
                strFormet = strFormet & " ORDER BY CASE SUBSTRING(LTRIM(DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE)),1,3) WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE)))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE)))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE)))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE)))) "
                strFormet = strFormet & " WHEN '(���' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE),5,LEN(DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE)))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE)))) "
                strFormet = strFormet & " ELSE LTRIM(DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE)) END "

                strSQL = String.Format(strFormet, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetCUSTDBLLIST")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    ' =============== SelectRtnSample Code
    Public Function GetMED_DBLLIST(ByVal strInfoXML As String, _
                                   ByRef intRowCnt As Integer, _
                                   ByRef intColCnt As Integer, _
                                   ByVal strYEAR As String, _
                                   ByVal strMEDNAME As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""

                If strYEAR <> "" Then Con1 = String.Format(" AND (SUBSTRING(YEARMON,1,4) = '{0}')", strYEAR)
                If strMEDNAME <> "" Then Con2 = String.Format(" AND (dbo.MD_GET_CUSTNAME_FUN(MEDCODE) LIKE '%{0}%')", strMEDNAME)
                strWhere = BuildFields(" ", Con1, Con2)

                strFormet = " SELECT "
                strFormet = strFormet & "  0 CHK,"
                strFormet = strFormet & "  MEDCODE,"
                strFormet = strFormet & "  dbo.MD_GET_CUSTNAME_FUN(MEDCODE) MEDNAME"
                strFormet = strFormet & "  FROM MD_BOOKING_MEDIUM"
                strFormet = strFormet & "  WHERE GFLAG IN('J','S') {0}"
                strFormet = strFormet & "  GROUP BY MEDCODE"
                strFormet = strFormet & " ORDER BY CASE SUBSTRING(LTRIM(DBO.MD_GET_CUSTNAME_FUN(MEDCODE)),1,3) WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(MEDCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(MEDCODE)))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(MEDCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(MEDCODE)))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(MEDCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(MEDCODE)))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(MEDCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(MEDCODE)))) "
                strFormet = strFormet & " WHEN '(���' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(MEDCODE),5,LEN(DBO.MD_GET_CUSTNAME_FUN(MEDCODE)))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(MEDCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(MEDCODE)))) "
                strFormet = strFormet & " ELSE LTRIM(DBO.MD_GET_CUSTNAME_FUN(MEDCODE)) END "

                strSQL = String.Format(strFormet, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetMED_DBLLIST")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "52. ��� ��ü�� ������ ���� ��ȸ"

    ' =============== SelectRtnSample Code
    Public Function GetCUSTMEDLIST(ByVal strInfoXML As String, _
                                   ByRef intRowCnt As Integer, _
                                   ByRef intColCnt As Integer, _
                                   ByVal strYEARMON As String, _
                                   ByVal strCUSTNAME As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""

                If strYEARMON <> "" Then Con1 = String.Format(" AND (YEARMON = '{0}')", strYEARMON)
                If strCUSTNAME <> "" Then Con2 = String.Format(" AND (MD_GET_CUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCUSTNAME)
                strWhere = BuildFields(" ", Con1, Con2)

                strFormet = " SELECT"
                strFormet = strFormet & "  0 CHK,"
                strFormet = strFormet & "  CLIENTCODE,"
                strFormet = strFormet & "  MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME"
                strFormet = strFormet & "  FROM "
                strFormet = strFormet & "  ("
                strFormet = strFormet & "   SELECT "
                strFormet = strFormet & "   0 CHK,"
                strFormet = strFormet & "   CLIENTCODE,"
                strFormet = strFormet & "   MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME"
                strFormet = strFormet & "   FROM MD_BOOKING_MEDIUM"
                strFormet = strFormet & "   WHERE GFLAG IN('J','S')  {0}"
                strFormet = strFormet & "   GROUP BY CLIENTCODE"
                strFormet = strFormet & "   UNION ALL"
                strFormet = strFormet & "   SELECT "
                strFormet = strFormet & "   0 CHK,"
                strFormet = strFormet & "   CLIENTCODE,"
                strFormet = strFormet & "   MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME"
                strFormet = strFormet & "   FROM MD_INTERNET_MEDIUM"
                strFormet = strFormet & "   WHERE GFLAG IN('J','S')   {0}"
                strFormet = strFormet & "   GROUP BY CLIENTCODE"
                strFormet = strFormet & "   UNION ALL"
                strFormet = strFormet & "   SELECT "
                strFormet = strFormet & "   0 CHK,"
                strFormet = strFormet & "   CLIENTCODE,"
                strFormet = strFormet & "   MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME"
                strFormet = strFormet & "   FROM MD_ELECTRIC_MEDIUM"
                strFormet = strFormet & "   WHERE 1=1  {0}"
                strFormet = strFormet & "   GROUP BY CLIENTCODE"
                strFormet = strFormet & "   UNION ALL"
                strFormet = strFormet & "   SELECT "
                strFormet = strFormet & "   0 CHK,"
                strFormet = strFormet & "   CLIENTCODE,"
                strFormet = strFormet & "   MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME"
                strFormet = strFormet & "   FROM MD_CATV_MEDIUM"
                strFormet = strFormet & "   WHERE 1=1  {0}"
                strFormet = strFormet & "   GROUP BY CLIENTCODE"
                strFormet = strFormet & "  ) A "
                strFormet = strFormet & "  GROUP BY CLIENTCODE"
                strFormet = strFormet & "  ORDER BY MD_GET_CUSTNAME_FUN(CLIENTCODE)"

                strSQL = String.Format(strFormet, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetCUSTMEDLIST")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "53. ��� ��ü�� �ϳ������� ���� ��ȸ"

    ' =============== SelectRtnSample Code
    Public Function GetCUSTMEDONELIST(ByVal strInfoXML As String, _
                                   ByRef intRowCnt As Integer, _
                                   ByRef intColCnt As Integer, _
                                   ByVal strYEARMON As String, _
                                   ByVal strCUSTNAME As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""

                If strYEARMON <> "" Then Con1 = String.Format(" AND (YEARMON = '{0}')", strYEARMON)
                If strCUSTNAME <> "" Then Con2 = String.Format(" AND (MD_GET_CUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCUSTNAME)
                strWhere = BuildFields(" ", Con1, Con2)

                strFormet = " SELECT"
                strFormet = strFormet & "  CLIENTCODE,"
                strFormet = strFormet & "  MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME"
                strFormet = strFormet & "  FROM "
                strFormet = strFormet & "  ("
                strFormet = strFormet & "   SELECT "
                strFormet = strFormet & "   CLIENTCODE,"
                strFormet = strFormet & "   MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME"
                strFormet = strFormet & "   FROM MD_BOOKING_MEDIUM"
                strFormet = strFormet & "   WHERE GFLAG IN('J','S')  {0}"
                strFormet = strFormet & "   GROUP BY CLIENTCODE"
                strFormet = strFormet & "   UNION ALL"
                strFormet = strFormet & "   SELECT "
                strFormet = strFormet & "   CLIENTCODE,"
                strFormet = strFormet & "   MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME"
                strFormet = strFormet & "   FROM MD_INTERNET_MEDIUM"
                strFormet = strFormet & "   WHERE GFLAG IN('J','S')   {0}"
                strFormet = strFormet & "   GROUP BY CLIENTCODE"
                strFormet = strFormet & "   UNION ALL"
                strFormet = strFormet & "   SELECT "
                strFormet = strFormet & "   CLIENTCODE,"
                strFormet = strFormet & "   MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME"
                strFormet = strFormet & "   FROM MD_ELECTRIC_MEDIUM"
                strFormet = strFormet & "   WHERE 1=1  {0}"
                strFormet = strFormet & "   GROUP BY CLIENTCODE"
                strFormet = strFormet & "   UNION ALL"
                strFormet = strFormet & "   SELECT "
                strFormet = strFormet & "   CLIENTCODE,"
                strFormet = strFormet & "   MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME"
                strFormet = strFormet & "   FROM MD_CATV_MEDIUM"
                strFormet = strFormet & "   WHERE 1=1  {0}"
                strFormet = strFormet & "   GROUP BY CLIENTCODE"
                strFormet = strFormet & "  ) A "
                strFormet = strFormet & "  GROUP BY CLIENTCODE"
                strFormet = strFormet & "  ORDER BY MD_GET_CUSTNAME_FUN(CLIENTCODE)"

                strSQL = String.Format(strFormet, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetCUSTMEDONELIST")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "54. ��� ��ü�� �ϳ���ü�� ���� ��ȸ"

    ' =============== SelectRtnSample Code
    Public Function GetCOMPANYMEDONELIST(ByVal strInfoXML As String, _
                                   ByRef intRowCnt As Integer, _
                                   ByRef intColCnt As Integer, _
                                   ByVal strYEARMON As String, _
                                   ByVal strMEDNAME As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""

                If strYEARMON <> "" Then Con1 = String.Format(" AND (YEARMON = '{0}')", strYEARMON)
                If strMEDNAME <> "" Then Con2 = String.Format(" AND (MD_GET_CUSTNAME_FUN(MEDCODE) LIKE '%{0}%')", strMEDNAME)
                strWhere = BuildFields(" ", Con1, Con2)

                strFormet = " SELECT"
                strFormet = strFormet & "  MEDCODE,"
                strFormet = strFormet & "  MD_GET_CUSTNAME_FUN(MEDCODE) MEDNAME"
                strFormet = strFormet & "  FROM "
                strFormet = strFormet & "  ("
                strFormet = strFormet & "   SELECT "
                strFormet = strFormet & "   MEDCODE,"
                strFormet = strFormet & "   MD_GET_CUSTNAME_FUN(MEDCODE) MEDNAME"
                strFormet = strFormet & "   FROM MD_ELECTRIC_MEDIUM"
                strFormet = strFormet & "   WHERE 1=1  {0}"
                strFormet = strFormet & "   GROUP BY MEDCODE"
                strFormet = strFormet & "   UNION ALL"
                strFormet = strFormet & "   SELECT "
                strFormet = strFormet & "   MEDCODE,"
                strFormet = strFormet & "   MD_GET_CUSTNAME_FUN(MEDCODE) MEDNAME"
                strFormet = strFormet & "   FROM MD_CATV_MEDIUM"
                strFormet = strFormet & "   WHERE 1=1  {0}"
                strFormet = strFormet & "   GROUP BY MEDCODE"
                strFormet = strFormet & "  ) A "
                strFormet = strFormet & "  GROUP BY MEDCODE"
                strFormet = strFormet & "  ORDER BY MD_GET_CUSTNAME_FUN(MEDCODE)"

                strSQL = String.Format(strFormet, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetCOMPANYMEDONELIST")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "55. CATV ��ü�� ä�θ� ���� ��ȸ"

    ' =============== SelectRtnSample Code
    Public Function GetCATVMED_LIST(ByVal strInfoXML As String, _
                                   ByRef intRowCnt As Integer, _
                                   ByRef intColCnt As Integer, _
                                   ByVal strYEAR As String, _
                                   ByVal strMEDNAME As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""

                If strYEAR <> "" Then Con1 = String.Format(" AND (SUBSTRING(YEARMON,1,4) = '{0}')", strYEAR)
                If strMEDNAME <> "" Then Con2 = String.Format(" AND (MD_GET_CUSTNAME_FUN(MEDCODE) LIKE '%{0}%')", strMEDNAME)
                strWhere = BuildFields(" ", Con1, Con2)

                strFormet = "   SELECT "
                strFormet = strFormet & "   MEDCODE,"
                strFormet = strFormet & "   MD_GET_CUSTNAME_FUN(MEDCODE) MEDNAME"
                strFormet = strFormet & "   FROM MD_CATV_MEDIUM"
                strFormet = strFormet & "   WHERE 1=1  {0}"
                strFormet = strFormet & "   GROUP BY MEDCODE"
                strFormet = strFormet & "  ORDER BY MD_GET_CUSTNAME_FUN(MEDCODE)"

                strSQL = String.Format(strFormet, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetCATVMED_LIST")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "56. CATV ��ü�� ������ ���� ��ȸ"

    ' =============== SelectRtnSample Code
    Public Function GetCATVCLIENT_LIST(ByVal strInfoXML As String, _
                                   ByRef intRowCnt As Integer, _
                                   ByRef intColCnt As Integer, _
                                   ByVal strYEAR As String, _
                                   ByVal strCUSTNAME As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""

                If strYEAR <> "" Then Con1 = String.Format(" AND (SUBSTRING(YEARMON,1,4) = '{0}')", strYEAR)
                If strCUSTNAME <> "" Then Con2 = String.Format(" AND (MD_GET_CUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCUSTNAME)
                strWhere = BuildFields(" ", Con1, Con2)

                strFormet = "   SELECT "
                strFormet = strFormet & "   CLIENTCODE,"
                strFormet = strFormet & "   MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME"
                strFormet = strFormet & "   FROM MD_CATV_MEDIUM"
                strFormet = strFormet & "   WHERE 1=1  {0}"
                strFormet = strFormet & "   GROUP BY CLIENTCODE"
                strFormet = strFormet & "  ORDER BY MD_GET_CUSTNAME_FUN(CLIENTCODE)"

                strSQL = String.Format(strFormet, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetCATVCLIENT_LIST")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "57. CATV ��ü�� ���� ��ȸ"

    ' =============== SelectRtnSample Code
    Public Function GetCATVREAL_MED_LIST(ByVal strInfoXML As String, _
                                   ByRef intRowCnt As Integer, _
                                   ByRef intColCnt As Integer, _
                                   ByVal strFROMYEAR As String, _
                                   ByVal strTOYEAR As String, _
                                   ByVal strREAL_MED_NAME As String, _
                                   ByVal strSPONSOR As String, _
                                   ByVal strCOMMIT As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2, Con3, Con4 As String
        Dim vntData As Object
        Dim strYEAR

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""
                Con3 = ""
                Con4 = ""

                If strFROMYEAR <> "" Then
                    If strTOYEAR <> "" Then
                        Con1 = String.Format(" AND (SUBSTRING(YEARMON,1,4) BETWEEN '{0}' AND '{1}')", strFROMYEAR, strTOYEAR)
                    Else
                        Con1 = String.Format(" AND (SUBSTRING(YEARMON,1,4) = '{0}')", strFROMYEAR)
                    End If
                Else
                    If strTOYEAR <> "" Then
                        Con1 = String.Format(" AND (SUBSTRING(YEARMON,1,4) = '{0}')", strTOYEAR)

                    Else
                        Con1 = " AND (SUBSTRING(YEARMON,1,4) BETWEEN '0000' AND '9999')"
                    End If
                End If

                If strSPONSOR = "Y" Then
                    Con2 = " AND (SPONSOR = 'Y')"
                ElseIf strSPONSOR = "N" Then
                    Con2 = " AND (SPONSOR IS NULL OR SPONSOR <> 'Y')"
                Else
                    Con2 = ""
                End If

                If strCOMMIT = "Y" Then
                    Con3 = " AND (TRU_TRANS_NO IS NOT NULL)"
                ElseIf strCOMMIT = "N" Then
                    Con3 = " AND (TRU_TRANS_NO IS NULL)"
                Else
                    Con3 = ""
                End If

                If strREAL_MED_NAME <> "" Then Con4 = String.Format(" AND (MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) LIKE '%{0}%')", strREAL_MED_NAME)
                strWhere = BuildFields(" ", Con1, Con2, Con3, Con4)

                strFormet = " SELECT "
                strFormet = strFormet & "    REAL_MED_CODE,"
                strFormet = strFormet & "    MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) REAL_MED_NAME"
                strFormet = strFormet & "    FROM MD_CATV_MEDIUM"
                strFormet = strFormet & "    WHERE 1=1 {0}"
                strFormet = strFormet & "    GROUP BY REAL_MED_CODE"
                strFormet = strFormet & "   ORDER BY MD_GET_REALMEDCODE_FUN(REAL_MED_CODE)"

                strSQL = String.Format(strFormet, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetCATVREAL_MED_LIST")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "56. �μ� ��ü�� ������ ���� ��ȸ"

    ' =============== SelectRtnSample Code
    Public Function GetPRINTCLIENT_LIST(ByVal strInfoXML As String, _
                                        ByRef intRowCnt As Integer, _
                                        ByRef intColCnt As Integer, _
                                        ByVal strFPUB_DATE As String, _
                                        ByVal strTPUB_DATE As String, _
                                        ByVal strCUSTNAME As String, _
                                        ByVal strMEDCODE As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2, Con3 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""
                Con3 = ""

                strFPUB_DATE = Replace(strFPUB_DATE, "-", "")
                strTPUB_DATE = Replace(strTPUB_DATE, "-", "")

                If strFPUB_DATE <> "" Then
                    If strTPUB_DATE <> "" Then
                        Con1 = String.Format(" AND (PUB_DATE BETWEEN '{0}' AND '{1}')", strFPUB_DATE, strTPUB_DATE)
                    Else
                        Con1 = String.Format(" AND (PUB_DATE BETWEEN '{0}' AND '99999999')", strFPUB_DATE)
                    End If
                Else
                    If strTPUB_DATE <> "" Then
                        Con1 = String.Format(" AND (PUB_DATE BETWEEN '00000000' AND '{0}')", strTPUB_DATE)
                    Else
                        Con1 = " AND (PUB_DATE BETWEEN '00000000' AND '99999999')"
                    End If
                End If

                If strCUSTNAME <> "" Then Con2 = String.Format(" AND (DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCUSTNAME)
                If strMEDCODE <> "" Then Con3 = String.Format(" AND (MEDCODE = '{0}')", strMEDCODE)
                strWhere = BuildFields(" ", Con1, Con2, Con3)

                strFormet = "   SELECT "
                strFormet = strFormet & "   CLIENTCODE,"
                strFormet = strFormet & "   DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME"
                strFormet = strFormet & "   FROM MD_BOOKING_MEDIUM"
                strFormet = strFormet & "   WHERE 1=1 AND GFLAG IN('J','S') AND TRU_TRANS_NO IS NOT NULL {0}"
                strFormet = strFormet & "   GROUP BY CLIENTCODE"
                strFormet = strFormet & " ORDER BY CASE SUBSTRING(LTRIM(DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE)),1,3) WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE)))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE)))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE)))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE)))) "
                strFormet = strFormet & " WHEN '(���' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE),5,LEN(DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE)))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE)))) "
                strFormet = strFormet & " ELSE LTRIM(DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE)) END "

                strSQL = String.Format(strFormet, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetPRINTCLIENT_LIST")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "57. �μ� ��ü�� ��ü�� ���� ��ȸ"

    ' =============== SelectRtnSample Code
    Public Function GetPRINTMED_LIST(ByVal strInfoXML As String, _
                                        ByRef intRowCnt As Integer, _
                                        ByRef intColCnt As Integer, _
                                        ByVal strFPUB_DATE As String, _
                                        ByVal strTPUB_DATE As String, _
                                        ByVal strMEDNAME As String, _
                                        ByVal strCLIENTCODE As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2, Con3 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""
                Con3 = ""

                strFPUB_DATE = Replace(strFPUB_DATE, "-", "")
                strTPUB_DATE = Replace(strTPUB_DATE, "-", "")

                If strFPUB_DATE <> "" Then
                    If strTPUB_DATE <> "" Then
                        Con1 = String.Format(" AND (PUB_DATE BETWEEN '{0}' AND '{1}')", strFPUB_DATE, strTPUB_DATE)
                    Else
                        Con1 = String.Format(" AND (PUB_DATE BETWEEN '{0}' AND '99999999')", strFPUB_DATE)
                    End If
                Else
                    If strTPUB_DATE <> "" Then
                        Con1 = String.Format(" AND (PUB_DATE BETWEEN '00000000' AND '{0}')", strTPUB_DATE)
                    Else
                        Con1 = " AND (PUB_DATE BETWEEN '00000000' AND '99999999')"
                    End If
                End If

                If strMEDNAME <> "" Then Con2 = String.Format(" AND (DBO.MD_GET_CUSTNAME_FUN(MEDCODE) LIKE '%{0}%')", strMEDNAME)
                If strCLIENTCODE <> "" Then Con3 = String.Format(" AND (CLIENTCODE = '{0}')", strCLIENTCODE)
                strWhere = BuildFields(" ", Con1, Con2, Con3)

                strFormet = "    SELECT "
                strFormet = strFormet & "    MEDCODE,"
                strFormet = strFormet & "    DBO.MD_GET_CUSTNAME_FUN(MEDCODE) MEDNAME"
                strFormet = strFormet & "    FROM MD_BOOKING_MEDIUM"
                strFormet = strFormet & "    WHERE 1=1 AND TRU_TRANS_NO IS NOT NULL   {0}"
                strFormet = strFormet & "    GROUP BY MEDCODE"
                strFormet = strFormet & " ORDER BY CASE SUBSTRING(LTRIM(DBO.MD_GET_CUSTNAME_FUN(MEDCODE)),1,3) WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(MEDCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(MEDCODE)))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(MEDCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(MEDCODE)))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(MEDCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(MEDCODE)))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(MEDCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(MEDCODE)))) "
                strFormet = strFormet & " WHEN '(���' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(MEDCODE),5,LEN(DBO.MD_GET_CUSTNAME_FUN(MEDCODE)))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(MEDCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(MEDCODE)))) "
                strFormet = strFormet & " ELSE LTRIM(DBO.MD_GET_CUSTNAME_FUN(MEDCODE)) END "

                strSQL = String.Format(strFormet, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetPRINTMED_LIST")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "58. ������ �˾� ������ ���� ��ȸ"

    ' =============== SelectRtnSample Code
    Public Function GetCOMPANYCLIENTONELIST(ByVal strInfoXML As String, _
                                            ByRef intRowCnt As Integer, _
                                            ByRef intColCnt As Integer, _
                                            ByVal strYEARMON As String, _
                                            ByVal strCUSTNAME As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""

                If strYEARMON <> "" Then Con1 = String.Format(" AND (YEARMON = '{0}')", strYEARMON)
                If strCUSTNAME <> "" Then Con2 = String.Format(" AND (MD_GET_CUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCUSTNAME)
                strWhere = BuildFields(" ", Con1, Con2)


                strFormet = " SELECT"
                strFormet = strFormet & "  CLIENTCODE,"
                strFormet = strFormet & "  MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME"
                strFormet = strFormet & "  FROM "
                strFormet = strFormet & "  ("
                strFormet = strFormet & "   SELECT "
                strFormet = strFormet & "   CLIENTCODE,"
                strFormet = strFormet & "   MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME"
                strFormet = strFormet & "   FROM MD_ELECTRIC_MEDIUM"
                strFormet = strFormet & "   WHERE 1=1  {0}"
                strFormet = strFormet & "   GROUP BY CLIENTCODE"
                strFormet = strFormet & "   UNION ALL"
                strFormet = strFormet & "   SELECT "
                strFormet = strFormet & "   CLIENTCODE,"
                strFormet = strFormet & "   MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME"
                strFormet = strFormet & "   FROM MD_CATV_MEDIUM"
                strFormet = strFormet & "   WHERE 1=1  {0}"
                strFormet = strFormet & "   GROUP BY CLIENTCODE"
                strFormet = strFormet & "  ) A "
                strFormet = strFormet & "  GROUP BY CLIENTCODE"
                strFormet = strFormet & "  ORDER BY MD_GET_CUSTNAME_FUN(CLIENTCODE)"

                strSQL = String.Format(strFormet, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetCOMPANYCLIENTONELIST")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "59. ������ ������ �ŷ��� ��ȸ"
    Public Function Get_ELECCOMMI_HDR(ByVal strInfoXML As String, _
                                       ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                       ByVal strYEARMON As String, _
                                       ByVal strTRANSNO As String, _
                                       ByVal strREAL_MED_CODE As String) As String

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strWhere As String
        Dim strXMLData As String
        Dim Con1 As String
        Dim Con2 As String
        Dim Con3 As String


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            Con1 = ""
            Con2 = ""
            Con3 = ""
            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
            If strTRANSNO <> "" Then Con2 = String.Format(" AND (TRANSNO = '{0}')", strTRANSNO)
            If strREAL_MED_CODE <> "" Then Con3 = String.Format(" AND (REAL_MED_CODE like '%{0}%')", strREAL_MED_CODE)

            strWhere = BuildFields(" ", Con1, Con2, Con3)
            strFormat = "SELECT TRANSYEARMON, TRANSNO, MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) REAL_MED_NAME,  SC_DEPT_NAME_FUN(DEPT_CD) DEPT_NAME, DEMANDDAY, PRINTDAY, AMT, VAT, (AMT+VAT) SUMAMTVAT  FROM MD_ELECCOMMI_HDR WHERE 1=1 {0} ORDER BY MED_FLAG "
            strSQL = String.Format(strFormat, strWhere)
            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                strXMLData = .mobjSCGLSql.SQLSelectXml(strSQL, intRowCnt, intColCnt)
                Return strXMLData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetJOBSEQList")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    Public Function Get_ELECCOMMI_ALLLIST(ByVal strInfoXML As String, _
                                      ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                      ByVal strYEARMON As String, _
                                      ByVal strTRANSNO As String, _
                                      ByVal strREAL_MED_CODE As String) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strWhere As String
        Dim vntData As Object
        Dim Con1 As String
        Dim Con2 As String
        Dim Con3 As String


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            Con1 = ""
            Con2 = ""
            Con3 = ""
            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
            If strTRANSNO <> "" Then Con2 = String.Format(" AND (TRANSNO = '{0}')", strTRANSNO)
            If strREAL_MED_CODE <> "" Then Con3 = String.Format(" AND (CLIENTCODE like '%{0}%')", strREAL_MED_CODE)

            strWhere = BuildFields(" ", Con1, Con2, Con3)
            strFormat = "SELECT 0 CHK ,TRANSYEARMON,  TRANSNO, CLIENTCODE, MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, MEDCODE, MD_GET_CUSTNAME_FUN(MEDCODE) MEDNAME,REAL_MED_CODE,MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) REAL_MED_NAME,AMT, VAT, (AMT+VAT) SUMAMTVAT FROM MD_ELECCOMMI_HDR WHERE 1=1 and attr03 <> 'Y' {0} ORDER BY MED_FLAG "
            strSQL = String.Format(strFormat, strWhere)
            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetJOBSEQList")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    Public Function Get_ELECCOMMI_LIST(ByVal strInfoXML As String, _
                                       ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                       ByVal strYEARMON As String, _
                                       ByVal strTRANSNO As String, _
                                       ByVal strREAL_MED_CODE As String) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strWhere As String
        Dim vntData As Object
        Dim Con1 As String
        Dim Con2 As String
        Dim Con3 As String


        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            Con1 = ""
            Con2 = ""
            Con3 = ""
            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
            If strTRANSNO <> "" Then Con2 = String.Format(" AND (TRANSNO = '{0}')", strTRANSNO)
            If strREAL_MED_CODE <> "" Then Con3 = String.Format(" AND (REAL_MED_CODE like '%{0}%')", strREAL_MED_CODE)

            strWhere = BuildFields(" ", Con1, Con2, Con3)
            strFormat = "SELECT TRANSYEARMON,  TRANSNO, MEDCODE, MD_GET_CUSTNAME_FUN(MEDCODE) MEDNAME,CLIENTCODE, MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME, REAL_MED_CODE,MD_GET_REALMEDCODE_FUN(REAL_MED_CODE) REAL_MED_NAME,DEPT_CD, DEMANDDAY, PRINTDAY, AMT, SUSURATE, SUSU, VAT, MD_GET_ELECMEDNAME_FUN(MED_FLAG) MED_FLAG FROM MD_ELECCOMMI_DTL WHERE 1=1 {0} ORDER BY MED_FLAG "
            strSQL = String.Format(strFormat, strWhere)
            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetJOBSEQList")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "60. ������ ��ȸ"

    ' =============== SelectRtnSample Code
    Public Function GetETCCUSTNO(ByVal strInfoXML As String, _
                                 ByRef intRowCnt As Integer, _
                                 ByRef intColCnt As Integer, _
                                 ByVal strCUSTCODE As String, _
                                 ByVal strCUSTNAME As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""

                If strCUSTCODE <> "" Then Con1 = String.Format(" AND (CUSTCODE LIKE '%{0}%')", strCUSTCODE)
                If strCUSTNAME <> "" Then Con2 = String.Format(" AND (CUSTNAME LIKE '%{0}%')", strCUSTNAME)
                strWhere = BuildFields(" ", Con1, Con2)

                strSelFields = "CUSTCODE , CUSTNAME , BUSINO, COMPANYNAME"

                strFormet = "select {0} from SC_CUST_TEMP where 1=1 AND ATTR10 =1 AND CUSTCODE LIKE 'A%' AND  instr(CUSTNAME,'Ʋ��') =0 AND instr(CUSTNAME,'���') =0 {1}  ORDER BY DECODE(SUBSTRING(CUSTNAME,1,3),'(��)', TRIM(SUBSTRING(CUSTNAME,4)),'(��)',TRIM(SUBSTRING(CUSTNAME,4)),'(�̵�',TRIM(SUBSTRING(CUSTNAME,6)),'(��)',TRIM(SUBSTRING(CUSTNAME,4)), '(��)',TRIM(SUBSTRING(CUSTNAME,4)), '(��)',TRIM(SUBSTRING(CUSTNAME,4)), '(��)',TRIM(SUBSTRING(CUSTNAME,4)), '(��)',TRIM(SUBSTRING(CUSTNAME,4)), '(��)',TRIM(SUBSTRING(CUSTNAME,4)),'(���', TRIM(SUBSTRING(CUSTNAME,5)), TRIM(CUSTNAME))"


                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetETCCUSTNO")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "60. ������ ��ȸ"

    ' =============== SelectRtnSample Code
    Public Function GetETCMEDNO(ByVal strInfoXML As String, _
                                ByRef intRowCnt As Integer, _
                                ByRef intColCnt As Integer, _
                                ByVal strCUSTCODE As String, _
                                ByVal strCUSTNAME As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""

                If strCUSTCODE <> "" Then Con1 = String.Format(" AND (CUSTCODE LIKE '%{0}%')", strCUSTCODE)
                If strCUSTNAME <> "" Then Con2 = String.Format(" AND (CUSTNAME LIKE '%{0}%')", strCUSTNAME)
                strWhere = BuildFields(" ", Con1, Con2)

                strSelFields = "CUSTCODE , CUSTNAME , BUSINO, COMPANYNAME"

                strFormet = "select {0} from SC_CUST_TEMP where 1=1 AND ATTR10 =1 AND CUSTCODE LIKE 'B%' AND  instr(CUSTNAME,'Ʋ��') =0 AND instr(CUSTNAME,'���') =0 {1}  ORDER BY DECODE(SUBSTRING(CUSTNAME,1,3),'(��)', TRIM(SUBSTRING(CUSTNAME,4)),'(��)',TRIM(SUBSTRING(CUSTNAME,4)),'(�̵�',TRIM(SUBSTRING(CUSTNAME,6)),'(��)',TRIM(SUBSTRING(CUSTNAME,4)), '(��)',TRIM(SUBSTRING(CUSTNAME,4)), '(��)',TRIM(SUBSTRING(CUSTNAME,4)), '(��)',TRIM(SUBSTRING(CUSTNAME,4)), '(��)',TRIM(SUBSTRING(CUSTNAME,4)), '(��)',TRIM(SUBSTRING(CUSTNAME,4)),'(���', TRIM(SUBSTRING(CUSTNAME,5)), TRIM(CUSTNAME))"


                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetETCMEDNO")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "61. �д�� ��ȸ"
    ' =============== SelectRtnSample Code
    Public Function GetSUMDIVAMT(ByVal strInfoXML As String, _
                                 ByRef intRowCnt As Double, _
                                 ByRef intColCnt As Double, _
                                 ByVal strYEARMON As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����


                strSQL = "SELECT MEDNAME,AMT FROM MD_CATV_MEDIUMDIVAMT_V WHERE YEARMON = '" & strYEARMON & "' "



                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetSUMDIVAMT")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "62. ������ ��ȸ"

    ' =============== SelectRtnSample Code
    Public Function GetEXCUSTNO(ByVal strInfoXML As String, _
                              ByRef intRowCnt As Integer, _
                              ByRef intColCnt As Integer, _
                              ByVal strCUSTCODE As String, _
                              ByVal strCUSTNAME As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""

                If strCUSTCODE <> "" Then Con1 = String.Format(" AND (CUSTCODE LIKE '%{0}%')", strCUSTCODE)
                If strCUSTNAME <> "" Then Con2 = String.Format(" AND (CUSTNAME LIKE '%{0}%')", strCUSTNAME)
                strWhere = BuildFields(" ", Con1, Con2)

                strSelFields = "CUSTCODE, CUSTNAME, BUSINO, COMPANYNAME"

                strFormet = "select {0} from SC_CUST_TEMP where 1=1 AND CUSTCODE LIKE 'G%' AND ATTR10 =1 {1}  ORDER BY "
                strFormet = strFormet & " CASE SUBSTRING(LTRIM(CUSTNAME),1,3) WHEN '(��)' THEN LTRIM(SUBSTRING(CUSTNAME,4,LEN(CUSTNAME))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(CUSTNAME,4,LEN(CUSTNAME))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(CUSTNAME,4,LEN(CUSTNAME))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(CUSTNAME,4,LEN(CUSTNAME))) "
                strFormet = strFormet & " WHEN '(���' THEN LTRIM(SUBSTRING(CUSTNAME,5,LEN(CUSTNAME))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(CUSTNAME,4,LEN(CUSTNAME))) "
                strFormet = strFormet & " ELSE LTRIM(CUSTNAME) END "


                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetEXCUSTNO")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "63. ����� ��ȸ"

    ' =============== SelectRtnSample Code
    Public Function GetCUSTNO_HIGHCUSTCODE(ByVal strInfoXML As String, _
                                           ByRef intRowCnt As Integer, _
                                           ByRef intColCnt As Integer, _
                                           ByVal strCLIENTSUBCODE As String, _
                                           ByVal strCLIENTSUBNAME As String, _
                                           ByVal strCLIENTCODE As String, _
                                           ByVal strCLIENTNAME As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2, Con3, Con4 As String
        Dim vntData As Object
        Dim strHIGH As String
        Dim strHIGHNAME As String

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""
                Con3 = ""
                Con4 = ""
                .mobjSCGLSql.SQLConnect(.DBConnStr)


                If strCLIENTCODE <> "" Then
                    strHIGH = SelectRtn_HIGH(strCLIENTCODE)
                End If

                If strCLIENTNAME <> "" Then
                    strHIGHNAME = SelectRtn_HIGHNAME(strCLIENTNAME)
                End If

                If strHIGH <> "" Then Con1 = String.Format(" AND (HIGHCUSTCODE LIKE '%{0}%')", strHIGH)
                If strHIGHNAME <> "" Then Con2 = String.Format(" AND (COMPANYNAME LIKE '%{0}%')", Trim(strHIGHNAME))
                If strCLIENTSUBCODE <> "" Then Con3 = String.Format(" AND (CUSTCODE LIKE '%{0}%')", strCLIENTSUBCODE)
                If strCLIENTSUBNAME <> "" Then Con4 = String.Format(" AND (CUSTNAME LIKE '%{0}%')", strCLIENTSUBNAME)
                strWhere = BuildFields(" ", Con1, Con2, Con3, Con4)

                strSelFields = "CUSTCODE, CUSTNAME, BUSINO, COMPANYNAME,ACCUSTCODE, HIGHCUSTCODE, DBO.MD_GET_CUSTNAME_FUN(HIGHCUSTCODE) HIGHCUSTNAME"

                strFormet = "select {0} from SC_CUST_TEMP where 1=1 AND CUSTCODE LIKE 'A%' AND ATTR10 =1 {1}  ORDER BY "
                strFormet = strFormet & " CASE SUBSTRING(LTRIM(CUSTNAME),1,3) WHEN '(��)' THEN LTRIM(SUBSTRING(CUSTNAME,4,LEN(CUSTNAME))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(CUSTNAME,4,LEN(CUSTNAME))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(CUSTNAME,4,LEN(CUSTNAME))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(CUSTNAME,4,LEN(CUSTNAME))) "
                strFormet = strFormet & " WHEN '(���' THEN LTRIM(SUBSTRING(CUSTNAME,5,LEN(CUSTNAME))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(CUSTNAME,4,LEN(CUSTNAME))) "
                strFormet = strFormet & " ELSE LTRIM(CUSTNAME) END "


                strSQL = String.Format(strFormet, strSelFields, strWhere)


                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetCUSTNO")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
    Public Function SelectRtn_HIGH(ByVal strCUSTCODE As String) As String
        '������� �ܼ���ȸ
        Dim strSQL, strFormat, strRtn As String

        'SetConfig(strInfoXML) '�⺻���� Setting

        With mobjSCGLConfig '�⺻���� Config ��ü

            Try
                strSQL = String.Format("SELECT HIGHCUSTCODE FROM SC_CUST_TEMP WHERE CUSTCODE = '" & strCUSTCODE & "'")
                strRtn = .mobjSCGLSql.SQLSelectOneScalar(strSQL)
                Return strRtn
            Catch err As Exception
                ' ���� ����
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_SEQNO")
            Finally
            End Try
        End With
        '������� �ܼ���ȸ
    End Function
    Public Function SelectRtn_HIGHNAME(ByVal strHIGH As String) As String
        '������� �ܼ���ȸ
        Dim strSQL, strFormat, strRtn As String

        'SetConfig(strInfoXML) '�⺻���� Setting

        With mobjSCGLConfig '�⺻���� Config ��ü

            Try
                strSQL = String.Format("SELECT COMPANYNAME FROM SC_CUST_TEMP WHERE CUSTNAME LIKE '%" & strHIGH & "%'")
                strRtn = .mobjSCGLSql.SQLSelectOneScalar(strSQL)
                Return strRtn
            Catch err As Exception
                ' ���� ����
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_SEQNO")
            Finally
            End Try
        End With
        '������� �ܼ���ȸ
    End Function
#End Region

#Region "64. ����귣�� ��ȸ"
    Public Function GetCUSTNO_COMBRAND(ByVal strInfoXML As String, _
                                       ByRef intRowCnt As Integer, _
                                       ByRef intColCnt As Integer, _
                                       ByVal strCUSTCODE As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2 As String
        Dim vntData As Object
        Dim strHIGH As String
        Dim strHIGHNAME As String

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""
                .mobjSCGLSql.SQLConnect(.DBConnStr)


                strHIGH = SelectRtn_HIGH(strCUSTCODE)
                'strHIGHNAME = SelectRtn_HIGHNAME(strHIGH)




                If strCUSTCODE <> "" Then Con1 = String.Format(" AND (CUSTCODE LIKE '%{0}%')", strCUSTCODE)
                'If strCUSTNAME <> "" Then Con2 = String.Format(" AND (DBO.MD_GET_CUSTNAME_FUN(HIGHCUSTCODE) LIKE '%{0}%')", strCUSTNAME)
                strWhere = BuildFields(" ", Con1)

                strSelFields = "SEQNO,SEQNAME,DEPTCD,DBO.SC_DEPT_NAME_FUN(DEPTCD) DEPTNAME"

                strFormet = "select {0} from SC_JOBCUST where 1=1  AND SEQNAME = '����' {1}  "


                strSQL = String.Format(strFormet, strSelFields, strWhere)


                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetCUSTNO_COMBRAND")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function


    Public Function Get_HIGHCUST(ByVal strInfoXML As String, _
                                 ByRef intRowCnt As Integer, _
                                 ByRef intColCnt As Integer, _
                                 ByVal strCUSTCODE As String, _
                                 ByVal strCUSTNAME As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2 As String
        Dim vntData As Object
        Dim strHIGH As String
        Dim strHIGHNAME As String

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""
                .mobjSCGLSql.SQLConnect(.DBConnStr)

                If strCUSTCODE <> "" Then Con1 = String.Format(" AND (HIGHCUSTCODE LIKE '%{0}%')", strHIGH)
                If strCUSTNAME <> "" Then Con2 = String.Format(" AND (DBO.MD_GET_CUSTNAME_FUN(HIGHCUSTCODE) LIKE '%{0}%')", strCUSTNAME)
                strWhere = BuildFields(" ", Con1)

                strSelFields = "HIGHCUSTCODE, DBO.MD_GET_CUSTNAME_FUN(HIGHCUSTCODE) HIGHCUSTNAME, BUSINO"

                strFormet = "select {0} from SC_CUST_TEMP where 1=1 AND CUSTCODE LIKE 'A%' AND ATTR10 =1 {1} GROUP BY HIGHCUSTCODE, BUSINO  ORDER BY "
                strFormet = strFormet & " CASE SUBSTRING(LTRIM(DBO.MD_GET_CUSTNAME_FUN(HIGHCUSTCODE)),1,3) WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(HIGHCUSTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(HIGHCUSTCODE)))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(HIGHCUSTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(HIGHCUSTCODE)))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(HIGHCUSTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(HIGHCUSTCODE)))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(HIGHCUSTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(HIGHCUSTCODE)))) "
                strFormet = strFormet & " WHEN '(���' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(HIGHCUSTCODE),5,LEN(DBO.MD_GET_CUSTNAME_FUN(HIGHCUSTCODE)))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(DBO.MD_GET_CUSTNAME_FUN(HIGHCUSTCODE),4,LEN(DBO.MD_GET_CUSTNAME_FUN(HIGHCUSTCODE)))) "
                strFormet = strFormet & " ELSE LTRIM(DBO.MD_GET_CUSTNAME_FUN(HIGHCUSTCODE)) END "


                strSQL = String.Format(strFormet, strSelFields, strWhere)


                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetCUSTNO")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "65. MPP ��ȸ"

    ' =============== SelectRtnSample Code
    Public Function GetMPP(ByVal strInfoXML As String, _
                             ByRef intRowCnt As Integer, _
                             ByRef intColCnt As Integer, _
                             ByVal strCUSTCODE As String, _
                             ByVal strCUSTNAME As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""

                If strCUSTCODE <> "" Then Con1 = String.Format(" AND (CUSTCODE LIKE '%{0}%')", strCUSTCODE)
                If strCUSTNAME <> "" Then Con2 = String.Format(" AND (CUSTNAME LIKE '%{0}%')", strCUSTNAME)
                strWhere = BuildFields(" ", Con1, Con2)

                strSelFields = "CUSTCODE , CUSTNAME , BUSINO, COMPANYNAME"

                strFormet = "select {0} from SC_CUST_TEMP where 1=1 AND CUSTCODE LIKE 'P%'  AND ATTR10 = 1 {1} ORDER BY  CASE SUBSTRING(LTRIM(CUSTNAME),1,3) WHEN '(��)' THEN LTRIM(SUBSTRING(CUSTNAME,4,LEN(CUSTNAME))) ELSE LTRIM(CUSTNAME) END"


                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetMPP")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "66. PROJECT NUMBER ��ȸ"

    ' =============== SelectRtnSample Code
    Public Function GetPONO(ByVal strInfoXML As String, _
                            ByRef intRowCnt As Integer, _
                            ByRef intColCnt As Integer, _
                            ByVal strCODE As String, _
                            ByVal strNAME As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""

                If strCODE <> "" Then Con1 = String.Format(" AND (PROJECTNO LIKE '%{0}%')", strCODE)
                If strNAME <> "" Then Con2 = String.Format(" AND (PROJECTNM LIKE '%{0}%')", strNAME)
                strWhere = BuildFields(" ", Con1, Con2)

                strSelFields = " PROJECTNO,"
                strSelFields = strSelFields & " PROJECTNM,"
                strSelFields = strSelFields & " DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME,"
                strSelFields = strSelFields & " DBO.SC_GET_CUSTNAME_FUN(TIMCODE) TIMNAME,"
                strSelFields = strSelFields & " DBO.SC_GET_SUBSEQNAME_FUN(SUBSEQ) SUBSEQNAME,"
                strSelFields = strSelFields & " CASE ISNULL(GROUPGBN,'') WHEN '2' THEN '�׷�' ELSE '��׷�' END AS GROUPGBN,"
                strSelFields = strSelFields & " SUBSTRING(CREDAY,1,4)+'-'+SUBSTRING(CREDAY,5,2)+'-'+SUBSTRING(CREDAY,7,2) CREDAY,"
                strSelFields = strSelFields & " DBO.SC_DEPT_NAME_FUN(CPDEPTCD) CPDEPTNAME,"
                strSelFields = strSelFields & " DBO.SC_EMPNAME_FUN(CPEMPNO) CPEMPNAME,"
                strSelFields = strSelFields & " MEMO"
                strFormet = "select {0} FROM PD_PONO where 1=1 {1} ORDER BY 1"


                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetPONO")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region


#Region "67. EMPNO ��ȸ"
    Public Function GetPDEMP(ByVal strInfoXML As String, _
                             ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                             ByVal strCODE As String, _
                             ByVal strNAME As String, _
                             ByVal strGUBUN As String, _
                             ByVal strDEPTCD As String, _
                             ByVal strDEPTNAME As String) As Object

        Dim strSQL, strFormat, strSelFields, strKeys As String
        Dim strCondition As String
        Dim strCondition2 As String
        Dim strChkDate As String = ""
        Dim vntData As Object
        Dim Con1, Con2, Con3, Con4, Con5 As String

        Con1 = ""
        Con2 = ""
        Con3 = ""
        Con4 = ""
        Con5 = ""

        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig
            If Len(strCODE) = 5 Then
                strCODE = "000" & strCODE
            End If

            '�ѱ��� ���
            If strCODE <> "" Then Con1 = String.Format(" AND (EMPNO = '{0}')", strCODE)
            If strNAME <> "" Then Con2 = String.Format(" AND EMP_NAME LIKE '%{0}%'", strNAME)
            If strGUBUN <> "A" Then Con3 = String.Format(" AND SC_EMP_STATUS = '{0}'", strGUBUN)
            If strDEPTCD <> "" Then Con4 = String.Format(" AND (CC_CODE = '{0}')", strDEPTCD)
            If strDEPTNAME <> "" Then Con5 = String.Format(" AND DBO.SC_DEPT_NAME_FUN(CC_CODE) LIKE '%{0}%'", strDEPTNAME)

            '��ȸ �ʵ� ����

            strSelFields = "EMPNO,EMP_NAME,CC_CODE,DBO.SC_DEPT_NAME_FUN(CC_CODE) CC_NAME,CASE SC_EMP_STATUS WHEN '0' THEN '����' WHEN '1' THEN '����' WHEN '3' THEN '����' END SC_EMP_STATUS,CASE ISNULL(E_MAIL,'') WHEN 'NULL' THEN '' ELSE ISNULL(E_MAIL,'') END E_MAIL,TEL,CELLPHONE,PASSWORD"
            strFormat = "SELECT {0} FROM SC_EMPLOYEE_MST A " & _
                                     "WHERE USE_YN = 'Y'  {1} {2} {3} {4} {5} " & _
                                     "ORDER BY CC_CODE"
            strSQL = String.Format(strFormat, _
                                   strSelFields, Con1, Con2, Con3, Con4, Con5)



            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetPDEMP")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "67-1. EMPNO_MANAGER ��ȸ"
    Public Function GetPDEMP_MANAGER(ByVal strInfoXML As String, _
                                     ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                     ByVal strCODE As String, _
                                     ByVal strNAME As String, _
                                     ByVal strGUBUN As String, _
                                     ByVal strDEPTCD As String, _
                                     ByVal strDEPTNAME As String) As Object

        Dim strSQL, strFormat, strSelFields, strKeys As String
        Dim strCondition As String
        Dim strCondition2 As String
        Dim strChkDate As String = ""
        Dim vntData As Object
        Dim Con1, Con2, Con3, Con4, Con5 As String

        Con1 = ""
        Con2 = ""
        Con3 = ""
        Con4 = ""
        Con5 = ""

        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig
            If Len(strCODE) = 5 Then
                strCODE = "000" & strCODE
            End If

            '�ѱ��� ���
            If strCODE <> "" Then Con1 = String.Format(" AND (EMPNO = '{0}')", strCODE)
            If strNAME <> "" Then Con2 = String.Format(" AND EMP_NAME LIKE '%{0}%'", strNAME)
            If strGUBUN <> "A" Then Con3 = String.Format(" AND SC_EMP_STATUS = '{0}'", strGUBUN)
            If strDEPTCD <> "" Then Con4 = String.Format(" AND (CC_CODE = '{0}')", strDEPTCD)
            If strDEPTNAME <> "" Then Con5 = String.Format(" AND DBO.SC_DEPT_NAME_FUN(CC_CODE) LIKE '%{0}%'", strDEPTNAME)

            '��ȸ �ʵ� ����

            strSelFields = " EMPNO,EMP_NAME,CC_CODE,DBO.SC_DEPT_NAME_FUN(CC_CODE) CC_NAME,CASE SC_EMP_STATUS WHEN '0' THEN '����' WHEN '1' THEN '����' WHEN '3' THEN '����' END SC_EMP_STATUS,CASE ISNULL(E_MAIL,'') WHEN 'NULL' THEN '' ELSE ISNULL(E_MAIL,'') END E_MAIL,TEL,CELLPHONE,PASSWORD "
            strFormat = " SELECT {0} FROM SC_EMPLOYEE_MST A " & _
                                     " WHERE USE_YN = 'Y' AND ISNULL(MANAGER,'N') ='Y' {1} {2} {3} {4} {5} " & _
                                     " ORDER BY CC_CODE "

            strSQL = String.Format(strFormat, _
                                   strSelFields, Con1, Con2, Con3, Con4, Con5)


            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetPDEMP_MANAGER")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "68. JOB NUMBER ��ȸ"

    ' =============== SelectRtnSample Code
    Public Function GetJOBNO(ByVal strInfoXML As String, _
                            ByRef intRowCnt As Integer, _
                            ByRef intColCnt As Integer, _
                            ByVal strCODE As String, _
                            ByVal strNAME As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""

                If strCODE <> "" Then Con1 = String.Format(" AND (A.JOBNO LIKE '%{0}%')", strCODE)
                If strNAME <> "" Then Con2 = String.Format(" AND (A.JOBNAME LIKE '%{0}%')", strNAME)
                strWhere = BuildFields(" ", Con1, Con2)

                strSelFields = " A.JOBNO JOBNO,A.JOBNAME JOBNAME,B.TIMCODE TIMCODE,DBO.PD_COMMITION_FUN(B.CLIENTCODE) COMMITION,B.CLIENTCODE,B.SUBSEQ,A.REQDAY,DBO.MD_GETPUBNAME_FUN(A.CREPART) CREPART"

                strFormet = "select {0} FROM PD_JOBNO A LEFT JOIN PD_PONO B ON A.PROJECTNO = B.PROJECTNO where 1=1 {1} ORDER BY 1"


                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetJOBNO")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "69. ITEMCODE ��ȸ"

    ' =============== SelectRtnSample Code
    Public Function GetITEMCODE(ByVal strInfoXML As String, _
                                ByRef intRowCnt As Integer, _
                                ByRef intColCnt As Integer, _
                                ByVal strDIV As String, _
                                ByVal strCLASSNAME As String, _
                                ByVal strITEMNAME As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2, Con3 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""

                If strDIV <> "0" Then Con1 = String.Format(" AND (DIV = '{0}')", strDIV)
                If strCLASSNAME <> "" Then Con2 = String.Format(" AND (CLASSNAME LIKE '%{0}%')", strCLASSNAME)
                If strITEMNAME <> "" Then Con3 = String.Format(" AND (ITEMNAME LIKE '%{0}%')", strITEMNAME)
                strWhere = BuildFields(" ", Con1, Con2, Con3)

                strSelFields = "ITEMCODE,DIVNAME,CLASSNAME,ITEMNAME, CASE DIV WHEN '2' THEN '1' ELSE '0' END AS COMMIFLAG,DIV,CLASS,DETAIL_YN,PRODUCTIONCOMMISSION"

                strFormet = "select {0} FROM PD_ITEMCODE where 1=1 {1} ORDER BY 1"


                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetITEMCODE")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "70. ITEMCODE �Һз� Combo Setting"
    '�޺�Ÿ�԰�������
    Public Function GetDIVITEM(ByVal strInfoXML As String, _
                                ByRef intRowCnt As Integer, _
                                ByRef intColCnt As Integer) As Object



        Dim strSQL, strFormat, strSelFields As String
        Dim vntData As Object
        Dim strWhere

        SetConfig(strInfoXML)   '�⺻���� ����

        '��ȸ �ʵ� ����
        strSelFields = "DIV,DIVNAME "

        'SQL�� ����

        strFormat = "SELECT '0' as DIV,'��ü' DIVNAME UNION ALL SELECT {0} " & _
                    "FROM PD_ITEMCODE " & _
                    "WHERE 1=1 " & _
                    "GROUP BY {0} ORDER BY DIV"

        With mobjSCGLConfig
            strSQL = String.Format(strFormat, strSelFields)

            ''������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetDIVITEM")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#Region "71. INSJOBNO POP��ȸ"
    ' =============== SelectRtnSample Code GetINSJOBNO
    Public Function GetINSJOBNO(ByVal strInfoXML As String, ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                ByVal strINSJOBNO As String, _
                                ByVal strSEQ As String, _
                                ByVal strJOBNAME As String, _
                                ByVal strCUSTCODE As String) As Object
        Dim strCols As String         '�÷�����
        Dim strWhere As String       'Where���� ����
        Dim strFormat As String      'SQL Format ����
        Dim strSQL As String          'SQL ����
        Dim vntData As Object        'Array Return ����(Array ������� �� ����)
        Dim strXMLData As String    'XML  Return ����(XML  �� ����� �� ����)
        Dim intCnt, strCREPART
        Dim strVAL
        Dim strUSER
        Dim Con1
        SetConfig(strInfoXML) '�⺻���� Setting
        strUSER = mobjSCGLConfig.WRKUSR
        strVAL = strSEQ & strINSJOBNO
        strCols = "SEQ,JOBNO,JOBNAME"
        Con1 = ""

        strJOBNAME = Replace(strJOBNAME, "'", "''")
        If strJOBNAME <> "" Then Con1 = String.Format(" AND (JOBNAME LIKE '%{0}%')", strJOBNAME)
        strWhere = BuildFields(" ", Con1)
        strFormat = "SELECT {0} FROM  PD_TRANS_INSJOBTEMP WHERE cast(SEQ as varchar(20))+JOBNO NOT IN  ('" & strVAL & "') AND CUSER = '" & strUSER & "' AND ATTR01 = '" & strCUSTCODE & "' {1}"

        With mobjSCGLConfig '�⺻���� Config ��ü
            strSQL = String.Format(strFormat, strCols, strWhere)
            Try
                ' DB ����
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                ' ------ Array ������ ��ȸ (True �϶� ������� ���� ��ȸ(Sheet Data Binding �� ��� ���), False �϶� �����͸� ��ȸ)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
                ' ------ XML ������ ��ȸ
                'strXMLData = .mobjSCGLSql.SQLSelectXml(strSQL, intRowCnt, intColCnt)
                'Return strXMLData
            Catch err As Exception
                ' ���� ����
                Throw RaiseSysErr(err, CLASS_NAME & ".GetINSJOBNO")
            Finally
                ' DB ���� ����
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
    Public Function GetsINSJOBNO(ByVal strInfoXML As String, ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                 ByVal strINSJOBNO As String, _
                                 ByVal strSEQ As String, _
                                 ByVal strJOBNO As String, _
                                 ByVal strCUSTCODE As String) As Object
        Dim strCols As String         '�÷�����
        Dim strWhere As String       'Where���� ����
        Dim strFormat As String      'SQL Format ����
        Dim strSQL As String          'SQL ����
        Dim vntData As Object        'Array Return ����(Array ������� �� ����)
        Dim strXMLData As String    'XML  Return ����(XML  �� ����� �� ����)
        Dim intCnt, strCREPART
        Dim strVAL
        Dim strUSER
        Dim Con1
        SetConfig(strInfoXML) '�⺻���� Setting
        strUSER = mobjSCGLConfig.WRKUSR
        strVAL = strSEQ & strINSJOBNO
        strCols = "SEQ,JOBNO,JOBNAME"
        Con1 = ""
        If strJOBNO <> "" Then Con1 = String.Format(" AND (JOBNO LIKE '%{0}%')", strJOBNO)
        strWhere = BuildFields(" ", Con1)
        strFormat = "SELECT {0} FROM  PD_TRANS_INSJOBTEMP WHERE cast(SEQ as varchar(20))+JOBNO NOT IN  ('" & strVAL & "') AND CUSER = '" & strUSER & "' AND ATTR01 = '" & strCUSTCODE & "' {1}"

        With mobjSCGLConfig '�⺻���� Config ��ü
            strSQL = String.Format(strFormat, strCols, strWhere)
            Try
                ' DB ����
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                ' ------ Array ������ ��ȸ (True �϶� ������� ���� ��ȸ(Sheet Data Binding �� ��� ���), False �϶� �����͸� ��ȸ)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
                ' ------ XML ������ ��ȸ
                'strXMLData = .mobjSCGLSql.SQLSelectXml(strSQL, intRowCnt, intColCnt)
                'Return strXMLData
            Catch err As Exception
                ' ���� ����
                Throw RaiseSysErr(err, CLASS_NAME & ".GetsINSJOBNO")
            Finally
                ' DB ���� ����
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "72. �ŷ����� POP��ȸ"
    Public Function GetPDTRANSNO(ByVal strInfoXML As String, _
                                 ByRef intRowCnt As Integer, _
                                 ByRef intColCnt As Integer, _
                                 ByVal strTRANSYEARMON As String, _
                                 ByVal strTRANSNO As String, _
                                 ByVal strCHK As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2, Con3, Con4 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""
                Con3 = ""

                '���ݰ�꼭 �Ϸ� ���� �� ���� ��ȸ
                If strTRANSYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strTRANSYEARMON)
                If strTRANSNO <> "" Then Con2 = String.Format(" AND (TRANSNO = '{0}')", strTRANSNO)
                If strCHK = "-1" Then
                    Con3 = String.Format(" AND (CASE SUM(DIV) WHEN 0 THEN '�Ϸ�' ELSE '�̿Ϸ�' END = '{0}')", "�Ϸ�")
                Else
                    Con3 = String.Format(" AND (CASE SUM(DIV) WHEN 0 THEN '�Ϸ�' ELSE '�̿Ϸ�' END = '{0}')", "�̿Ϸ�")
                End If

                strWhere = BuildFields(" ", Con1, Con2, Con3)

                'strSelFields = "TRANSYEARMON , TRANSNO , MEMO"

                'strFormet = "select {0} from PD_TRANS_HDR where 1=1 {1}"
                strFormet = " SELECT"
                strFormet = strFormet & " TRANSYEARMON,TRANSNO,DBO.MD_GET_MEDBUSINO_FUN(CLIENTCODE) BUSINO,"
                strFormet = strFormet & " CASE SUM(DIV) WHEN 0 THEN '�Ϸ�' ELSE '�̿Ϸ�' END MEMO,"
                strFormet = strFormet & " CLIENTCODE CUSTCODE,DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) CUSTNAME"
                strFormet = strFormet & " FROM"
                strFormet = strFormet & " ("
                strFormet = strFormet & " SELECT A.TRANSYEARMON,A.TRANSNO,A.TAXSEQ,"
                strFormet = strFormet & " CASE ISNULL(A.TAXSEQ,0) WHEN 0 THEN 1 ELSE 0 END AS DIV,   "
                strFormet = strFormet & " CLIENTCODE FROM PD_TRANS_DTL A LEFT JOIN PD_TRANS_HDR B ON A.TRANSYEARMON = B.TRANSYEARMON AND A.TRANSNO = B.TRANSNO"
                strFormet = strFormet & " GROUP BY A.TRANSYEARMON,A.TRANSNO,A.TAXSEQ,B.CLIENTCODE"
                strFormet = strFormet & " ) DATA"
                strFormet = strFormet & " GROUP BY TRANSYEARMON,TRANSNO,CLIENTCODE"
                strFormet = strFormet & " HAVING 1=1 {0}"
                strSQL = String.Format(strFormet, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True) '
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetPDTRANSNO")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#End Region

#Region "73. ����ó ��ȸ"

    ' =============== SelectRtnSample Code
    Public Function GetEXECUSTNO(ByVal strInfoXML As String, _
                                 ByRef intRowCnt As Integer, _
                                 ByRef intColCnt As Integer, _
                                 ByVal strCUSTCODE As String, _
                                 ByVal strCUSTNAME As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""

                If strCUSTCODE <> "" Then Con1 = String.Format(" AND (HIGHCUSTCODE LIKE '%{0}%')", strCUSTCODE)
                If strCUSTNAME <> "" Then Con2 = String.Format(" AND (CUSTNAME LIKE '%{0}%')", strCUSTNAME)
                strWhere = BuildFields(" ", Con1, Con2)

                strSelFields = "HIGHCUSTCODE CUSTCODE, CUSTNAME, BUSINO, COMPANYNAME,ACCUSTCODE"

                strFormet = "select {0} from SC_CUST_HDR where 1=1 AND MEDFLAG = 'M' AND USE_FLAG =1 {1}  ORDER BY "
                strFormet = strFormet & " CASE SUBSTRING(LTRIM(CUSTNAME),1,3) WHEN '(��)' THEN LTRIM(SUBSTRING(CUSTNAME,4,LEN(CUSTNAME))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(CUSTNAME,4,LEN(CUSTNAME))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(CUSTNAME,4,LEN(CUSTNAME))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(CUSTNAME,4,LEN(CUSTNAME))) "
                strFormet = strFormet & " WHEN '(���' THEN LTRIM(SUBSTRING(CUSTNAME,5,LEN(CUSTNAME))) "
                strFormet = strFormet & " WHEN '(��)' THEN LTRIM(SUBSTRING(CUSTNAME,4,LEN(CUSTNAME))) "
                strFormet = strFormet & " ELSE LTRIM(CUSTNAME) END "


                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetCUSTNO")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "74. ��ǥ�������� ��ȸ"
    Public Function COMMIVOCHNO_CHECKED(ByVal strInfoXML As String, _
                                        ByRef intRowCnt As Integer, _
                                        ByRef intColCnt As Integer, _
                                        ByVal strTAXYEARMON As String, _
                                        ByVal strTAXNO As String) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strWhere As String
        Dim vntData As Object
        Dim Con1 As String
        Dim Con2 As String
        Dim Con3 As String

        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            Con1 = ""
            Con2 = ""
            Con3 = ""
            If strTAXYEARMON <> "" Then Con1 = String.Format(" AND (TAXYEARMON = '{0}')", strTAXYEARMON)
            If strTAXNO <> "" Then Con2 = String.Format(" AND (TAXNO = '{0}')", strTAXNO)

            strWhere = BuildFields(" ", Con1, Con2, Con3)
            strFormat = "SELECT TAXYEARMON, TAXNO, RMSNO FROM PD_VOCH_MST WHERE 1=1 and semu in('BR','B5') {0} "
            strSQL = String.Format(strFormat, strWhere)
            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".COMMIVOCHNO_CHECKED")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    Public Function COMMIVOCHNO_CHECKED_MERGE(ByVal strInfoXML As String, _
                                              ByRef intRowCnt As Integer, _
                                              ByRef intColCnt As Integer, _
                                              ByVal strTAXYEARMON As String, _
                                              ByVal strTAXNO As String, _
                                              ByVal strMEDFLAG As String) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strWhere As String
        Dim vntData As Object
        Dim Con1 As String
        Dim Con2 As String
        Dim Con3 As String

        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            Con1 = ""
            Con2 = ""
            Con3 = ""

            If strTAXYEARMON <> "" Then Con1 = String.Format(" AND (TAXYEARMON = '{0}')", strTAXYEARMON)
            If strTAXNO <> "" Then Con2 = String.Format(" AND (TAXNO = '{0}')", strTAXNO)
            If strMEDFLAG <> "" Then Con3 = String.Format(" AND (MEDFLAG = '{0}')", strMEDFLAG)

            strWhere = BuildFields(" ", Con1, Con2, Con3)
            strFormat = "SELECT MTAXYEARMON, MTAXNO FROM PD_MERGETAX_DTL WHERE 1=1 {0} AND ISNULL(ATTR10,0) <> 999999"
            strSQL = String.Format(strFormat, strWhere)
            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".VOCHNO_CHECKED")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#End Region


#Region "74-2. ��ǥ�������� ��ȸ"
    Public Function VOCHNO_CHECKED(ByVal strInfoXML As String, _
                                        ByRef intRowCnt As Integer, _
                                        ByRef intColCnt As Integer, _
                                        ByVal strTAXYEARMON As String, _
                                        ByVal strTAXNO As String) As Object

        Dim strSQL As String            'SQL��
        Dim strFormat As String         '�ӽ� SQL��
        Dim strSelFields As String      '��ȸ�ʵ�
        Dim strWhere As String
        Dim vntData As Object
        Dim Con1 As String
        Dim Con2 As String
        Dim Con3 As String

        SetConfig(strInfoXML)   '�⺻���� ����
        With mobjSCGLConfig

            Con1 = ""
            Con2 = ""
            Con3 = ""
            If strTAXYEARMON <> "" Then Con1 = String.Format(" AND (TAXYEARMON = '{0}')", strTAXYEARMON)
            If strTAXNO <> "" Then Con2 = String.Format(" AND (TAXNO = '{0}')", strTAXNO)
            Con3 = String.Format(" AND (GBN = '{0}')", "P")

            strWhere = BuildFields(" ", Con1, Con2, Con3)
            strFormat = "SELECT TAXYEARMON, TAXNO, RMSNO FROM PD_VOCH_MST WHERE 1=1 {0} "
            strSQL = String.Format(strFormat, strWhere)
            '������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".VOCHNO_CHECKED")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#End Region

#Region "75-1. ITEMCODE �з��׸� ��ȸ"

    ' =============== SelectRtnSample Code
    Public Function GetDIVCLASS(ByVal strInfoXML As String, _
                                ByRef intRowCnt As Integer, _
                                ByRef intColCnt As Integer, _
                                ByVal strCODE As String, _
                                ByVal strNAME As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""

                If strCODE <> "" Then Con1 = String.Format(" AND (CLASSCD = '{0}')", strCODE)
                If strNAME <> "" Then Con2 = String.Format(" AND (CLASSNM LIKE '%{0}%')", strNAME)

                strWhere = BuildFields(" ", Con1, Con2)

                strSelFields = "CLASSCD,CLASSNM,DIVCD,DIVNM"

                strFormet = "select {0} FROM PD_ITEMCLASSCODE where 1=1 {1} ORDER BY 3,1"


                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetDIVCLASS")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "76. ITEMCODE ��з� PopUp"

    ' =============== SelectRtnSample Code
    Public Function GetDIVCODE(ByVal strInfoXML As String, _
                               ByRef intRowCnt As Integer, _
                               ByRef intColCnt As Integer, _
                               ByVal strCODE As String, _
                               ByVal strNAME As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""

                If strCODE <> "" Then Con1 = String.Format(" AND (DIVCD = '{0}')", strCODE)
                If strNAME <> "" Then Con2 = String.Format(" AND (DIVNM LIKE '%{0}%')", strNAME)

                strWhere = BuildFields(" ", Con1, Con2)

                strSelFields = "DIVCD,DIVNM"

                strFormet = "select {0} FROM PD_ITEMCLASSCODE where 1=1 {1} GROUP BY DIVCD,DIVNM ORDER BY 1"


                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetDIVCODE")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region


#Region "77. ITEMCODE �ߺз� PopUp"
    ' =============== SelectRtnSample Code
    Public Function GetDIVCLASSCODE(ByVal strInfoXML As String, _
                                    ByRef intRowCnt As Integer, _
                                    ByRef intColCnt As Integer, _
                                    ByVal strDIVCODE As String, _
                                    ByVal strDIVNAME As String, _
                                    ByVal strCLASSCODE As String, _
                                    ByVal strCLASSNAME As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2, Con3, Con4 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""
                Con3 = ""
                Con4 = ""

                If strDIVCODE <> "" Then Con1 = String.Format(" AND (DIVCD = '{0}')", strDIVCODE)
                If strDIVNAME <> "" Then Con2 = String.Format(" AND (DIVNM LIKE '%{0}%')", strDIVNAME)
                If strCLASSCODE <> "" Then Con3 = String.Format(" AND (CLASSCD = '{0}')", strCLASSCODE)
                If strCLASSNAME <> "" Then Con4 = String.Format(" AND (CLASSNM LIKE '%{0}%')", strCLASSNAME)

                strWhere = BuildFields(" ", Con1, Con2, Con3, Con4)

                strSelFields = "DIVCD,DIVNM,CLASSCD,CLASSNM"

                strFormet = "select {0} FROM PD_ITEMCLASSCODE where 1=1 {1} GROUP BY DIVCD,DIVNM,CLASSCD,CLASSNM ORDER BY 1,3"


                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetDIVCLASSCODE")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region


#End Region

#Region "71. ����ȸ"
    ' =============== SelectRtnSample Code
    Public Function GetTIMCODE(ByVal strInfoXML As String, _
                               ByRef intRowCnt As Integer, _
                               ByRef intColCnt As Integer, _
                               ByVal strCLIENTCODE As String, _
                               ByVal strCLIENTNAME As String, _
                               ByVal strCLIENTSUBCODE As String, _
                               ByVal strCLIENTSUBNAME As String) As Object

        Dim strCols As String        '�÷�����
        Dim strWhere As String       'Where���� ����
        Dim strFormat As String      'SQL Format ����
        Dim strSQL As String         'SQL ����
        Dim vntData As Object        'Array Return ����(Array ������� �� ����)
        Dim strSelFields
        Dim intCnt
        Dim intRtn
        Dim Con1, Con2, Con3, Con4
        SetConfig(strInfoXML)

        Con1 = "" : Con2 = "" : Con3 = "" : Con4 = ""

        If strCLIENTCODE <> "" Then Con1 = String.Format(" AND (HIGHCUSTCODE = '{0}')", strCLIENTCODE)
        If strCLIENTNAME <> "" Then Con2 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(HIGHCUSTCODE) LIKE '%{0}%')", strCLIENTNAME)
        If strCLIENTSUBCODE <> "" Then Con3 = String.Format(" AND (CUSTCODE = '{0}')", strCLIENTSUBCODE)
        If strCLIENTSUBNAME <> "" Then Con4 = String.Format(" AND (CUSTNAME LIKE '%{0}%')", strCLIENTSUBNAME)


        strWhere = BuildFields(" ", Con1, Con2, Con3, Con4)

        strFormat = "  SELECT "
        strFormat = strFormat & "  CUSTCODE, "
        strFormat = strFormat & "  CUSTNAME, "
        strFormat = strFormat & "  CLIENTSUBCODE, DBO.SC_GET_CUSTNAME_FUN(CLIENTSUBCODE) CLIENTSUBNAME, "
        strFormat = strFormat & "  HIGHCUSTCODE,"
        strFormat = strFormat & "  DBO.SC_GET_HIGHCUSTNAME_FUN(HIGHCUSTCODE) COMPANYNAME "
        strFormat = strFormat & "  FROM SC_CUST_DTL"
        strFormat = strFormat & "  WHERE 1=1 AND MEDFLAG = 'A' AND GBNFLAG = '0' AND USE_FLAG = '1' {0} ORDER BY CUSTNAME"

        strSQL = String.Format(strFormat, strWhere)

        '�⺻���� Setting
        With mobjSCGLConfig '�⺻���� Config ��ü
            Try
                ' DB ����
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                ' ------ Array ������ ��ȸ (True �϶� ������� ���� ��ȸ(Sheet Data Binding �� ��� ���), False �϶� �����͸� ��ȸ)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)

                Return vntData
            Catch err As Exception
                ' ���� ����
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn")
            Finally
                ' DB ���� ����
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "72. �޺�Ÿ�� ��ȸ"
    Public Function Get_COMBO_VALUE(ByVal strInfoXML As String, _
                                    ByRef intRowCnt As Integer, _
                                    ByRef intColCnt As Integer, _
                                    ByVal strCODE As String) As Object

        Dim strSQL, strFormat, strSelFields As String
        Dim vntData As Object

        SetConfig(strInfoXML)   '�⺻���� ����					

        '��ȸ �ʵ� ����					
        strSelFields = "CODE, CODE_NAME"

        'SQL�� ����

        strFormat = "SELECT {0} " & _
                    "FROM SC_CODE " & _
                    "WHERE CLASS_CODE = '" & strCODE & "'" & _
                    "ORDER BY SORT_SEQ "

        With mobjSCGLConfig
            strSQL = String.Format(strFormat, strSelFields)

            ''������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".Get_COMBO_VALUE")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "73. ��з��� �׷���ȸ"

    '********************************************************
    ' GetDataType_All()  ��ü���� ��ü���� �˻� �޺� select ó��
    '********************************************************
    Public Function GetDataType_DIVNM(ByVal strInfoXML As String, _
                                      ByRef intRowCnt As Integer, _
                                      ByRef intColCnt As Integer) As Object

        Dim strSQL, strFormat, strSelFields As String
        Dim vntData As Object

        SetConfig(strInfoXML)   '�⺻���� ����

        '��ȸ �ʵ� ����
        strSelFields = "DIVCD, DIVNM"


        strFormat = "SELECT {0} " & _
                           "FROM PD_ITEMCLASSCODE " & _
                           "GROUP BY DIVCD, DIVNM " & _
                           "UNION ALL  " & _
                           "SELECT ''DIVCD,'��ü' DIVNM " & _
                           "ORDER BY DIVCD "

        With mobjSCGLConfig
            strSQL = String.Format(strFormat, strSelFields)

            ''������ ��ȸ
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetDataType_DIVNM")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#End Region

#Region "74 JOB NUMBER, JOB ��Ͽ� ��ȸ"

    Public Function GetJOBNOSALE(ByVal strInfoXML As String, _
                                    ByRef intRowCnt As Integer, _
                                    ByRef intColCnt As Integer, _
                                    ByVal strCODE As String, _
                                    ByVal strNAME As String, _
                                    ByVal strREGDAY As String, _
                                    ByVal strCREPART As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2, Con3, Con4 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""
                Con3 = ""
                Con4 = ""

                If strCODE <> "" Then Con1 = String.Format(" AND (A.JOBNO LIKE '%{0}%')", strCODE)
                If strNAME <> "" Then Con2 = String.Format(" AND (A.JOBNAME LIKE '%{0}%')", strNAME)
                If strREGDAY <> "" Then Con2 = String.Format(" AND (A.REQDAY LIKE '%{0}%')", strREGDAY)
                If strCREPART <> "" Then Con2 = String.Format(" AND (A.CREPART LIKE '%{0}%')", strCREPART)

                strWhere = BuildFields(" ", Con1, Con2, Con3, Con4)

                strSelFields = " A.REQDAY,A.JOBNO JOBNO,A.JOBNAME JOBNAME,B.TIMCODE TIMCODE,DBO.PD_COMMITION_FUN(B.CLIENTCODE) COMMITION,B.CLIENTCODE,B.SUBSEQ,DBO.MD_GETPUBNAME_FUN(A.CREPART) CREPART"

                strFormet = "select {0} FROM PD_JOBNO A LEFT JOIN PD_PONO B ON A.PROJECTNO = B.PROJECTNO where 1=1 {1} ORDER BY 1"


                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetJOBNO")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#End Region

#Region "GROUP BLOCk : �ܺο� ����� Method"
    Private Function AddAlias(ByVal strFields As String, ByVal strAlias As String) As String
        Dim vntData() As String
        Dim i As Integer
        Dim strResult As New System.Text.StringBuilder

        vntData = Split(strFields, ",")
        For i = 0 To UBound(vntData)
            If strResult.Length = 0 Then
                strResult.Append(strAlias & "." & vntData(i))
            Else
                strResult.Append("," & strAlias & "." & vntData(i))
            End If

        Next
        Return strResult.ToString
    End Function
#End Region

End Class
