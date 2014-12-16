'****************************************************************************************
'Generated By: MakeSFAR V.2.0.0 - ��Ʈ�� Ŭ���� ����Ŀ - ��ȭ S&C
'�ý��۱���    : �ַ�Ǹ� /�ý��۸�/Server Control Class
'����   ȯ��    : COM+ Service Server Package
'���α׷���    : ccMDCMPRINTREG.vb
'��         ��    : - ����� ���� �մϴ�.
'Ư��  ����     : - Ư�̻��׿� ���� ǥ��
'                     -
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2004-03-30 ���� 10:32:13 By MakeSFARV.2.0.0
'            2) 2004-03-30 ���� 10:32:13 By �ۼ��ڸ��� ���ϴ�.
'****************************************************************************************

Imports System.Xml                  ' XMLó��
Imports SCGLControl                 ' ControlClass�� Base Class
Imports SCGLUtil.cbSCGLConfig       ' ConfigurationClass
Imports SCGLUtil.cbSCGLErr          '����ó�� Ŭ����
Imports SCGLUtil.cbSCGLXml          'XMLó�� Ŭ����
Imports SCGLUtil.cbSCGLUtil         '��Ÿ��ƿ��Ƽ Ŭ����
Imports eMDCO                       '����Ƽ �߰�

' ��ƼƼ Ŭ���� ���� �ش� ��ƼƼ Ŭ������ ������Ʈ�� ������ �� Imports �Ͻʽÿ�. 
' Imports ��ƼƼ������Ʈ

Public Class ccMDOTCLOUDCONTRACT
    Inherits ccControl

#Region "GROUP BLOCK : ���� �Ǵ� ��ⷹ���� ����/��� ����"
    Private CLASS_NAME = "ccMDOTCLOUDCONTRACT"                      ' �ڽ��� Ŭ������
    Private mobjceMD_CLOUD_CONTRACT As eMDCO.ceMD_CLOUD_CONTRACT    ' ����� Entity ���� ����
  #End Region

#Region "GROUP BLOCK : Property ����"
#End Region

#Region "GROUP BLOCK : Event ����"
    '********************************************************
    ' GetDataType()
    '********************************************************
    Public Function GetDataType_SEARCH(ByVal strInfoXML As String, _
                                       ByRef intRowCnt As Integer, _
                                       ByRef intColCnt As Integer, _
                                       ByRef strCLASS_CODE As String) As Object

        Dim strSQL As String
        Dim vntData As Object

        SetConfig(strInfoXML)   '�⺻���� ����

        strSQL = "SELECT '' CODE,'��ü' CODE_NAME UNION ALL "
        strSQL = strSQL & " SELECT "
        strSQL = strSQL & " CODE, CODE_NAME "
        strSQL = strSQL & " FROM SC_CODE "
        strSQL = strSQL & " WHERE CLASS_CODE = '" & strCLASS_CODE & "' "
        strSQL = strSQL & " ORDER BY CODE "

        With mobjSCGLConfig
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetDataType")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    Public Function GetDataType(ByVal strInfoXML As String, _
                                ByRef intRowCnt As Integer, _
                                ByRef intColCnt As Integer, _
                                ByRef strCLASS_CODE As String) As Object

        Dim strSQL As String
        Dim vntData As Object

        SetConfig(strInfoXML)   '�⺻���� ����

        strSQL = " SELECT "
        strSQL = strSQL & " CODE, CODE_NAME "
        strSQL = strSQL & " FROM SC_CODE "
        strSQL = strSQL & " WHERE CLASS_CODE = '" & strCLASS_CODE & "' "
        strSQL = strSQL & " ORDER BY CODE "

        With mobjSCGLConfig
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetDataType")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "GROUP BLOCK : �ܺο� ���� Method"
    Public Function SelectRtn(ByVal strInfoXML As String, _
                              ByRef intRowCnt As Integer, _
                              ByRef intColCnt As Integer, _
                              ByVal strYEARMON As String, _
                              ByVal strCLIENTCODE As String, _
                              ByVal strCONT_NAME As String, _
                              ByVal strCONT_TYPE As String) As Object     'XML  ������ ��ȸ��

        Dim strSQL As String
        Dim strFormet, strWhere As String
        Dim Con1, Con2, Con3, Con4 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = "" : Con2 = "" : Con3 = "" : Con4 = ""

                If strYEARMON <> "" Then Con1 = String.Format(" AND ('{0}' BETWEEN substring(TBRDSTDATE,1,6) AND substring(TBRDEDDATE,1,6))", strYEARMON)
                If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (CLIENTCODE = '{0}')", strCLIENTCODE)
                If strCONT_NAME <> "" Then Con3 = String.Format(" AND (CONT_NAME LIKE '%{0}%')", strCONT_NAME)
                If strCONT_TYPE <> "" Then Con4 = String.Format(" AND (CONT_TYPE = '{0}')", strCONT_TYPE)

                strWhere = BuildFields(" ", Con1, Con2, Con3, Con4)

                strFormet = " SELECT"
                strFormet = strFormet & " 0 CHK,"
                strFormet = strFormet & " GUBUN,"
                strFormet = strFormet & " CONT_TYPE,"
                strFormet = strFormet & " CONT_CODE,"
                strFormet = strFormet & " CONT_NAME,"
                strFormet = strFormet & " TBRDSTDATE,"
                strFormet = strFormet & " TBRDEDDATE,"
                strFormet = strFormet & " TOTAL_AMT,"
                strFormet = strFormet & " CLIENTCODE,"
                strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME,"
                strFormet = strFormet & " TIMCODE,"
                strFormet = strFormet & " DBO.SC_GET_CUSTNAME_FUN(TIMCODE) TIMNAME,"
                strFormet = strFormet & " EXCLIENTCODE,"
                strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(EXCLIENTCODE) EXCLIENTNAME,"
                strFormet = strFormet & " TIM_RATE,"
                strFormet = strFormet & " EX_RATE,"
                strFormet = strFormet & " CGV_RATE,"
                strFormet = strFormet & " MEMO"
                strFormet = strFormet & " FROM MD_CLOUD_CONTRACT"
                strFormet = strFormet & " WHERE 1=1 {0} "
                strFormet = strFormet & " ORDER BY CONT_CODE DESC, CLIENTNAME"

                strSQL = String.Format(strFormet, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    ' ============== ProcessRtn (Master & Detail) Sample Code 
    Public Function ProcessRtn(ByVal strInfoXML As String, _
                               ByVal vntData As Object, _
                               ByVal strYEAR As String) As String '������ INSERT/UPDATE �μ��� ���������� ByVal �̸� �߰� ������ ByRef


        Dim intRtn As Integer '����� ����  
        Dim i, intColCnt, intRows As Integer '����, �÷�Cnt, �ο�Cnt ����
        Dim strCONT_CODE
        Dim strTBRDSTDATE, strTBRDEDDATE

        SetConfig(strInfoXML) '�⺻���� Setting

        With mobjSCGLConfig '�⺻������ ������ �ִ� Config ��ü
            Try
                'DB���� �� Ʈ����� ����
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()

                'DetailOne ������ ó��
                If IsArray(vntData) Then
                    '''����� Entity ��ü����(Config ������ �Ѱܻ���)
                    mobjceMD_CLOUD_CONTRACT = New ceMD_CLOUD_CONTRACT(mobjSCGLConfig)
                    '''vntData�� �÷���, �ο���� �����Է�
                    intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)
                    '''�ش��ϴ�Row ��ŭ Loop

                    For i = 1 To intRows
                        If GetElement(vntData, "CONT_CODE", intColCnt, i, OPTIONAL_STR) = "" Then
                            If GetElement(vntData, "TBRDSTDATE", intColCnt, i, OPTIONAL_STR) <> "" Then strTBRDSTDATE = GetElement(vntData, "TBRDSTDATE", intColCnt, i, OPTIONAL_STR).Substring(0, 4) & GetElement(vntData, "TBRDSTDATE", intColCnt, i, OPTIONAL_STR).Substring(5, 2) & GetElement(vntData, "TBRDSTDATE", intColCnt, i, OPTIONAL_STR).Substring(8, 2)
                            If GetElement(vntData, "TBRDEDDATE", intColCnt, i, OPTIONAL_STR) <> "" Then strTBRDEDDATE = GetElement(vntData, "TBRDEDDATE", intColCnt, i, OPTIONAL_STR).Substring(0, 4) & GetElement(vntData, "TBRDEDDATE", intColCnt, i, OPTIONAL_STR).Substring(5, 2) & GetElement(vntData, "TBRDEDDATE", intColCnt, i, OPTIONAL_STR).Substring(8, 2)

                            strCONT_CODE = SelectRtn_SEQ(strYEAR)


                            intRtn = InsertRtn_MD_CLOUD_CONTRACT(vntData, intColCnt, i, strCONT_CODE, strTBRDSTDATE, strTBRDEDDATE)

                        Else
                            If GetElement(vntData, "TBRDSTDATE", intColCnt, i, OPTIONAL_STR) <> "" Then strTBRDSTDATE = GetElement(vntData, "TBRDSTDATE", intColCnt, i, OPTIONAL_STR).Substring(0, 4) & GetElement(vntData, "TBRDSTDATE", intColCnt, i, OPTIONAL_STR).Substring(5, 2) & GetElement(vntData, "TBRDSTDATE", intColCnt, i, OPTIONAL_STR).Substring(8, 2)
                            If GetElement(vntData, "TBRDEDDATE", intColCnt, i, OPTIONAL_STR) <> "" Then strTBRDEDDATE = GetElement(vntData, "TBRDEDDATE", intColCnt, i, OPTIONAL_STR).Substring(0, 4) & GetElement(vntData, "TBRDEDDATE", intColCnt, i, OPTIONAL_STR).Substring(5, 2) & GetElement(vntData, "TBRDEDDATE", intColCnt, i, OPTIONAL_STR).Substring(8, 2)

                            strCONT_CODE = GetElement(vntData, "CONT_CODE", intColCnt, i, OPTIONAL_STR)

                            intRtn = UpdateRtn_MD_CLOUD_CONTRACT(vntData, intColCnt, i, strCONT_CODE, strTBRDSTDATE, strTBRDEDDATE)
                        End If
                    Next
                End If

                'Ʈ�����Commit
                .mobjSCGLSql.SQLCommitTrans()
                Return intRtn
            Catch err As Exception
                'Ʈ�����RollBack �� ���� ����
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".ProcessRtn")
            Finally
                'Resource����
                .mobjSCGLSql.SQLDisconnect()
                mobjceMD_CLOUD_CONTRACT.Dispose()
            End Try
        End With
    End Function

    Public Function SelectRtn_SEQ(ByVal strYEAR As String) As String
        '������� �ܼ���ȸ
        Dim strSQL, strRtn As String

        With mobjSCGLConfig '�⺻���� Config ��ü
            Try
                strSQL = "SELECT '" & strYEAR & "'+DBO.LPAD(ISNULL(MAX(CAST(SUBSTRING(CONT_CODE,5,4) AS NUMERIC(4,0))),0)+1,4,'0') FROM MD_CLOUD_CONTRACT WHERE SUBSTRING(CONT_CODE,1,4) = '" & strYEAR & "'"
                strRtn = .mobjSCGLSql.SQLSelectOneScalar(strSQL)
                Return strRtn

            Catch err As Exception
                ' ���� ����
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_SEQ")
            Finally
            End Try
        End With
        '������� �ܼ���ȸ
    End Function

    Public Function SelectRtn_CountCheck(ByVal strInfoXML As String, _
                                         ByRef intRowCnt As Integer, _
                                         ByRef intColCnt As Integer, _
                                         ByVal strCONT_CODE As String) As Object ' xml ������ ��Ʈ�� ��������� object 

        Dim strSelFields As String         '�÷�����
        Dim strWhere As String       'Where���� ����
        Dim strFormat As String      'SQL Format ����
        Dim strSQL As String          'SQL ����
        Dim Con1 As String
        Dim vntData As Object        'Array Return ����(Array ������� �� ����)
        Dim strXMLData As String    'XML  Return ����(XML  �� ����� �� ����)

        Con1 = ""

        If strCONT_CODE <> "" Then Con1 = String.Format(" AND (CONT_CODE = '{0}')", strCONT_CODE)

        strWhere = BuildFields(" ", Con1)

        strSelFields = "SEQ "

        strFormat = "SELECT {0} from MD_CLOUD_AMT where 1=1 {1} "

        SetConfig(strInfoXML)
        With mobjSCGLConfig
            strSQL = String.Format(strFormat, strSelFields, strWhere)
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                ' ���� ����
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_CountCheck")
            Finally
                ' DB ���� ����
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With

    End Function

    ' =============== DeleteRtn Sample Code
    Public Function DeleteRtn(ByVal strInfoXML As String, _
                              ByVal strCONT_CODE As String) As Integer   '������ DELETE

        Dim intRtn As Integer
        Dim intRtn2 As Integer

        SetConfig(strInfoXML)    '�⺻���� Setting
        With mobjSCGLConfig    '�⺻���� Config ��ü
            Try
                ' �����Entity ��ü����(Config ������ �Ѱܻ���)
                mobjceMD_CLOUD_CONTRACT = New ceMD_CLOUD_CONTRACT(mobjSCGLConfig)
                ' DB ���� �� Ʈ����� ����
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                ' ��ƼƼ ������Ʈ�� Delete �޼ҵ� ȣ��
                intRtn = mobjceMD_CLOUD_CONTRACT.DeleteDo(strCONT_CODE)

                ' Ʈ����� Commit
                .mobjSCGLSql.SQLCommitTrans()
                Return intRtn
            Catch err As Exception
                'Ʈ����� RollBack �� ���� ����
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & "DeleteRtn")
            Finally
                'DB���� ����
                .mobjSCGLSql.SQLDisconnect()
                mobjceMD_CLOUD_CONTRACT.Dispose()
            End Try
        End With
    End Function
#End Region

#Region "GROUP BLOCK : �ܺο� ����� Method"
    Private Function InsertRtn_MD_CLOUD_CONTRACT(ByVal vntData As Object, _
                                                 ByVal intColCnt As Integer, _
                                                 ByVal intRow As Integer, _
                                                 ByRef strCONT_CODE As String, _
                                                 ByRef strTBRDSTDATE As Integer, _
                                                 ByRef strTBRDEDDATE As String) As Integer

        Dim intRtn As Integer
        intRtn = mobjceMD_CLOUD_CONTRACT.InsertDo( _
                                       strCONT_CODE, _
                                       GetElement(vntData, "GUBUN", intColCnt, intRow), _
                                       GetElement(vntData, "CLIENTCODE", intColCnt, intRow), _
                                       GetElement(vntData, "TIMCODE", intColCnt, intRow), _
                                       GetElement(vntData, "EXCLIENTCODE", intColCnt, intRow), _
                                       strTBRDSTDATE, _
                                       strTBRDEDDATE, _
                                       GetElement(vntData, "CONT_NAME", intColCnt, intRow), _
                                       GetElement(vntData, "CONT_TYPE", intColCnt, intRow), _
                                       GetElement(vntData, "TOTAL_AMT", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "TIM_RATE", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "EX_RATE", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "CGV_RATE", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "MEMO", intColCnt, intRow), _
                                       "N", _
                                       GetElement(vntData, "ATTR01", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR02", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR03", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR04", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR05", intColCnt, intRow))
        Return intRtn
    End Function

    Private Function UpdateRtn_MD_CLOUD_CONTRACT(ByVal vntData As Object, _
                                                 ByVal intColCnt As Integer, _
                                                 ByVal intRow As Integer, _
                                                 ByRef strCONT_CODE As String, _
                                                 ByRef strTBRDSTDATE As Integer, _
                                                 ByRef strTBRDEDDATE As String) As Integer
        Dim intRtn As Integer
        intRtn = mobjceMD_CLOUD_CONTRACT.UpdateDo( _
                                                strCONT_CODE, _
                                                GetElement(vntData, "GUBUN", intColCnt, intRow), _
                                                GetElement(vntData, "CLIENTCODE", intColCnt, intRow), _
                                                GetElement(vntData, "TIMCODE", intColCnt, intRow), _
                                                GetElement(vntData, "EXCLIENTCODE", intColCnt, intRow), _
                                                strTBRDSTDATE, _
                                                strTBRDEDDATE, _
                                                GetElement(vntData, "CONT_NAME", intColCnt, intRow), _
                                                GetElement(vntData, "CONT_TYPE", intColCnt, intRow), _
                                                GetElement(vntData, "TOTAL_AMT", intColCnt, intRow, NULL_NUM, True), _
                                                GetElement(vntData, "TIM_RATE", intColCnt, intRow, NULL_NUM, True), _
                                                GetElement(vntData, "EX_RATE", intColCnt, intRow, NULL_NUM, True), _
                                                GetElement(vntData, "CGV_RATE", intColCnt, intRow, NULL_NUM, True), _
                                                GetElement(vntData, "MEMO", intColCnt, intRow), _
                                                GetElement(vntData, "ATTR01", intColCnt, intRow), _
                                                GetElement(vntData, "ATTR02", intColCnt, intRow), _
                                                GetElement(vntData, "ATTR03", intColCnt, intRow), _
                                                GetElement(vntData, "ATTR04", intColCnt, intRow), _
                                                GetElement(vntData, "ATTR05", intColCnt, intRow))
        Return intRtn
    End Function

#End Region
End Class