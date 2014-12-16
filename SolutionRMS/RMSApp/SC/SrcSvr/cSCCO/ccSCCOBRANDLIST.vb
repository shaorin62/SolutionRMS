'****************************************************************************************
'�ý��۱���    : �ַ�Ǹ� /�ý��۸�/Server Control Class
'����   ȯ��    : COM+ Service Server Package
'���α׷���    : ccMDCMCUST_TRAN.vb
'��         ��    : - ����� ��� �մϴ�.
'Ư��  ����     : - Ư�̻��׿� ���� ǥ��
'                     -
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009-07-03 ���� 10:32:13 By KTY
'****************************************************************************************

Imports System.Xml                  ' XMLó��
Imports SCGLControl                 ' ControlClass�� Base Class 
Imports SCGLUtil.cbSCGLConfig       ' ConfigurationClass
Imports SCGLUtil.cbSCGLErr          '����ó�� Ŭ����
Imports SCGLUtil.cbSCGLXml          'XMLó�� Ŭ����
Imports SCGLUtil.cbSCGLUtil         '��Ÿ��ƿ��Ƽ Ŭ����
Imports eSCCO '����Ƽ �߰�

' ��ƼƼ Ŭ���� ���� �ش� ��ƼƼ Ŭ������ ������Ʈ�� ������ �� Imports �Ͻʽÿ�. 
' Imports ��ƼƼ������Ʈ

Public Class ccSCCOBRANDLIST
    Inherits ccControl

#Region "GROUP BLOCK : ���� �Ǵ� ��ⷹ���� ����/��� ����"
    Private CLASS_NAME = "ccSCCOBRANDLIST"                  '�ڽ��� Ŭ������
    Private mobjceSC_SUBSEQ_HDR As eSCCO.ceSC_SUBSEQ_HDR            '����� Entity ���� ����
    Private mobjceSC_SUBSEQ_DTL As eSCCO.ceSC_SUBSEQ_DTL              '����� Entity ���� ����
#End Region

#Region "GROUP BLOCK : Property ����"
#End Region

#Region "GROUP BLOCK : Event ����"
    Public Function HIGHSEQNAME_Check(ByVal strInfoXML As String, _
                                      ByRef intRowCnt As Integer, _
                                      ByRef intColCnt As Integer, _
                                      ByVal strHIGHSEQNAME As String, _
                                      ByVal strHIGHCUSTCODE As String) As Object                                      'XML  ������ ��ȸ��

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = "" : Con2 = ""

                If strHIGHSEQNAME <> "" Then Con1 = String.Format(" AND (Ltrim(Rtrim(HIGHSEQNAME)) = '{0}')", strHIGHSEQNAME)
                If strHIGHCUSTCODE <> "" Then Con2 = String.Format(" AND (CUSTCODE = '{0}')", strHIGHCUSTCODE)

                strWhere = BuildFields(" ", Con1, Con2)

                strFormet = "SELECT HIGHSEQNAME FROM SC_SUBSEQ_HDR WHERE 1=1 {0}"

                strSQL = String.Format(strFormet, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)

                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    Public Function GET_HighSeq_COMBO(ByVal strInfoXML As String, _
                                      ByRef intRowCnt As Integer, _
                                      ByRef intColCnt As Integer) As Object                                      'XML  ������ ��ȸ��

        Dim strSQL As String
        Dim strFormet, strWhere As String
        Dim Con1 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""

                strWhere = BuildFields(" ", Con1)

                strFormet = "SELECT HIGHSEQNO, HIGHSEQNAME  FROM SC_SUBSEQ_HDR WHERE 1=1 {0} ORDER BY HIGHSEQNAME"

                strSQL = String.Format(strFormet, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)

                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GET_HighSeq_COMBO")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    Public Function GET_HighSeq_COMBO_ROW(ByVal strInfoXML As String, _
                                         ByRef intRowCnt As Integer, _
                                         ByRef intColCnt As Integer, _
                                         ByRef strCLIENTCODE As String) As Object                                      'XML  ������ ��ȸ��

        Dim strSQL As String
        Dim strFormet, strWhere As String
        Dim Con1 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""

                If strCLIENTCODE <> "" Then Con1 = String.Format(" and CUSTCODE = '{0}'", strCLIENTCODE) '���

                strWhere = BuildFields(" ", Con1)

                strFormet = "SELECT HIGHSEQNO, HIGHSEQNAME  FROM SC_SUBSEQ_HDR WHERE 1=1 {0} ORDER BY HIGHSEQNAME"

                strSQL = String.Format(strFormet, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)

                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GET_HighSeq_COMBO")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    Public Function SelectRtn_CountCheck(ByVal strInfoXML As String, _
                                         ByRef intRowCnt As Integer, _
                                         ByRef intColCnt As Integer, _
                                         ByVal strSUBSEQ As String, _
                                         ByVal strMEDFLAG As String) As Object     'XML  ������ ��ȸ��

        Dim strSQL As String
        Dim strFormat, strSelFields, strWhere As String
        Dim Con1, Con2 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = "" : Con2 = ""


                If strSUBSEQ <> "" Then
                    Con1 = String.Format(" AND (SUBSEQ = '{0}')", strSUBSEQ)
                    Con2 = String.Format(" AND (HIGHSUBSEQ = '{0}')", strSUBSEQ)
                End If
                strSQL = strSQL & "  SELECT MEDFLAG, COUNT(*) FROM ("
                strSQL = strSQL & "  	SELECT 'B' MEDFLAG, SUBSEQ FROM MD_BOOKING_MEDIUM"
                strSQL = strSQL & "  	WHERE 1=1 " & Con1
                strSQL = strSQL & "  	UNION ALL"
                strSQL = strSQL & "  	SELECT 'A2' MEDFLAG, SUBSEQ FROM MD_CATV_MEDIUM"
                strSQL = strSQL & "  	WHERE 1=1 " & Con1
                strSQL = strSQL & "  	UNION ALL"
                strSQL = strSQL & "  	SELECT 'A' MEDFLAG, SUBSEQ FROM MD_ELECTRIC_MEDIUM"
                strSQL = strSQL & "  	WHERE 1=1 " & Con1
                strSQL = strSQL & "  	UNION ALL"
                strSQL = strSQL & "  	SELECT 'O' MEDFLAG, SUBSEQ FROM MD_INTERNET_MEDIUM"
                strSQL = strSQL & "  	WHERE 1=1 " & Con1
                strSQL = strSQL & "  	UNION ALL"
                strSQL = strSQL & "  	SELECT 'D' MEDFLAG, HIGHSUBSEQ SUBSEQ FROM MD_OUTDOOR_MEDIUM"
                strSQL = strSQL & "  	WHERE 1=1 " & Con2
                strSQL = strSQL & "  ) AAA"
                strSQL = strSQL & "  GROUP BY MEDFLAG"

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    Public Function Get_SESSION_DEPT_CD(ByVal strInfoXML As String, _
                                        ByRef intRowCnt As Integer, _
                                        ByRef intColCnt As Integer, _
                                        ByRef strUSERID As String) As String

        Dim strFormat As String      'SQL Format ����
        Dim strSQL As String         'SQL ����
        Dim vntData As Object        'Array Return ����(Array ������� �� ����)
        Dim strDEPT_CD

        SetConfig(strInfoXML)

        strSQL = "  SELECT  "
        strSQL = strSQL & "  CC_CODE "
        strSQL = strSQL & "  From SC_EMPLOYEE_MST"
        strSQL = strSQL & "  WHERE EMPNO = '" & strUSERID & "' and use_yn = 'Y'"

        '�⺻���� Setting
        With mobjSCGLConfig '�⺻���� Config ��ü
            Try
                ' DB ����
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                strDEPT_CD = .mobjSCGLSql.SQLSelectOneScalar(strSQL)

                Return strDEPT_CD
            Catch err As Exception
                ' ���� ����
                Throw RaiseSysErr(err, CLASS_NAME & ".Get_SESSION_DEPT_CD")
            Finally
                ' DB ���� ����
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#End Region

#Region "GROUP BLOCK : �ܺο� ���� Method"
    ' =============== SelectRtn_HIGHSUBSEQ ��ǥ�귣�� ���
    Public Function SelectRtn_HIGHSUBSEQ(ByVal strInfoXML As String, _
                                         ByRef intRowCnt As Integer, _
                                         ByRef intColCnt As Integer, _
                                         ByVal strCUSTNAME As String, _
                                         ByVal strHIGHSEQNAME As String) As Object     'XML  ������ ��ȸ��

        Dim strSQL As String
        Dim strFormat, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = ""
                Con2 = ""

                If strCUSTNAME <> "" Then Con1 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(CUSTCODE) LIKE '%{0}%')", strCUSTNAME)
                If strHIGHSEQNAME <> "" Then Con2 = String.Format(" AND (HIGHSEQNAME LIKE '%{0}%')", strHIGHSEQNAME)

                strWhere = BuildFields(" ", Con1, Con2)

                strFormat = "  SELECT "
                strFormat = strFormat & "  0 CHK, "
                strFormat = strFormat & "  HIGHSEQNO, HIGHSEQNAME, "
                strFormat = strFormat & "  CUSTCODE, '' BTN, DBO.SC_GET_HIGHCUSTNAME_FUN(CUSTCODE) CUSTNAME,  "
                strFormat = strFormat & "  DBO.SC_GET_SUMBRAND_FUN(HIGHSEQNO) SEQNAMES"
                strFormat = strFormat & "  FROM SC_SUBSEQ_HDR"
                strFormat = strFormat & "  WHERE 1=1 {0} ORDER BY HIGHSEQNO, DBO.SC_GET_HIGHCUSTNAME_FUN(CUSTCODE) , DBO.SC_GET_SUMBRAND_FUN(HIGHSEQNO)"


                strSQL = String.Format(strFormat, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)

                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_HIGHSUBSEQ")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    ' =============== SelectRtn_SUBSEQ �귣��
    Public Function SelectRtn_SUBSEQ(ByVal strInfoXML As String, _
                                     ByRef intRowCnt As Integer, _
                                     ByRef intColCnt As Integer, _
                                     ByVal strCUSTNAME As String, _
                                     ByVal strSEQNAME As String, _
                                     ByVal strHIGHSEQNAME As String, _
                                     ByVal strUSE_YN As String) As Object

        Dim strSQL As String
        Dim strFormat, strWhere As String
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

                If strCUSTNAME <> "" Then Con1 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(CUSTCODE) LIKE '%{0}%')", strCUSTNAME)
                If strSEQNAME <> "" Then Con2 = String.Format(" AND (SEQNAME LIKE '%{0}%')", strSEQNAME)
                If strHIGHSEQNAME <> "" Then Con3 = String.Format(" AND (DBO.SC_HIGHSEQNAME_FUN(HIGHSEQNO) LIKE '%{0}%')", strHIGHSEQNAME)
                If strUSE_YN <> "" Then Con4 = String.Format(" AND (ISNULL(ATTR01,'') = '{0}')", strUSE_YN)

                strWhere = BuildFields(" ", Con1, Con2, Con3, Con4)

                'strFormat = "  SELECT "
                'strFormat = strFormat & "  0 CHK, SEQNO, "
                'strFormat = strFormat & "  SEQNAME, HIGHSEQNO, "
                'strFormat = strFormat & "  TIMCODE, '' BTNTIM, DBO.SC_GET_CUSTNAME_FUN(TIMCODE) TIMNAME, "
                'strFormat = strFormat & "  CLIENTSUBCODE, '' BTNSUB, DBO.SC_GET_CUSTNAME_FUN(CLIENTSUBCODE) CLIENTSUBNAME, "
                'strFormat = strFormat & "  CUSTCODE, '' BTN, DBO.SC_GET_HIGHCUSTNAME_FUN(CUSTCODE) CUSTNAME, "
                'strFormat = strFormat & "  DEPT_CD, '' BTNDEPT,dbo.SC_DEPT_NAME_FUN(DEPT_CD)  DEPT_NAME, "
                'strFormat = strFormat & "  MEMO, case isnull(ATTR01,'N') when 'N' then '�̻��' else '���' end as Attr01 "
                'strFormat = strFormat & "  FROM SC_SUBSEQ_DTL"
                'strFormat = strFormat & "  WHERE 1=1 {0} "

                strFormat = " SELECT "
                strFormat = strFormat & " 0 CHK, SEQNO, "
                strFormat = strFormat & " SEQNAME, HIGHSEQNO, "
                strFormat = strFormat & " TIMCODE, '' BTNTIM, DBO.SC_GET_CUSTNAME_FUN(TIMCODE) TIMNAME, "
                strFormat = strFormat & " CLIENTSUBCODE, '' BTNSUB, DBO.SC_GET_CUSTNAME_FUN(CLIENTSUBCODE) CLIENTSUBNAME, "
                strFormat = strFormat & " CUSTCODE, '' BTN, DBO.SC_GET_HIGHCUSTNAME_FUN(CUSTCODE) CUSTNAME, "
                strFormat = strFormat & " DEPT_CD, '' BTNDEPT,dbo.SC_DEPT_NAME_FUN(DEPT_CD)  DEPT_NAME, DBO.SC_EMPNAME_FUN(A.CUSER) CUSER, CDATE,"
                strFormat = strFormat & " MEMO, case isnull(ATTR01,'N') when 'N' then '�̻��' WHEN 'Y' THEN '���' WHEN 'S' THEN '���ο�û' ELSE '���' end as Attr01, "
                strFormat = strFormat & " B.YEARMON MAXYEARMON"
                strFormat = strFormat & " FROM SC_SUBSEQ_DTL A"
                strFormat = strFormat & " LEFT JOIN ("
                strFormat = strFormat & "	SELECT SUBSEQ, MAX(YEARMON) YEARMON FROM ("
                strFormat = strFormat & "		SELECT 'B' MEDFLAG, SUBSEQ, MAX(YEARMON) YEARMON FROM MD_BOOKING_MEDIUM"
                strFormat = strFormat & "		WHERE 1=1 AND ISNULL(SUBSEQ,'') <> '' AND YEARMON >= CONVERT(CHAR(6), DATEADD(MM,-12,GETDATE()),112)"
                strFormat = strFormat & "		group by SUBSEQ"
                strFormat = strFormat & "		UNION ALL"
                strFormat = strFormat & "		SELECT 'A2' MEDFLAG, SUBSEQ, MAX(YEARMON) YEARMON FROM MD_CATV_MEDIUM"
                strFormat = strFormat & "		WHERE 1=1  AND YEARMON >= CONVERT(CHAR(6), DATEADD(MM,-12,GETDATE()),112)"
                strFormat = strFormat & "		group by SUBSEQ"
                strFormat = strFormat & "		UNION ALL"
                strFormat = strFormat & "		SELECT 'A' MEDFLAG, SUBSEQ, MAX(YEARMON) YEARMON FROM MD_ELECTRIC_MEDIUM"
                strFormat = strFormat & "		WHERE 1=1  AND YEARMON >= CONVERT(CHAR(6), DATEADD(MM,-12,GETDATE()),112)"
                strFormat = strFormat & "		group by SUBSEQ"
                strFormat = strFormat & "		UNION ALL"
                strFormat = strFormat & "		SELECT 'O' MEDFLAG, SUBSEQ, MAX(YEARMON) YEARMON FROM MD_INTERNET_MEDIUM"
                strFormat = strFormat & "		WHERE 1=1  AND YEARMON >= CONVERT(CHAR(6), DATEADD(MM,-12,GETDATE()),112)"
                strFormat = strFormat & "		group by SUBSEQ"
                strFormat = strFormat & "		UNION ALL"
                strFormat = strFormat & "		SELECT 'D' MEDFLAG, HIGHSUBSEQ SUBSEQ, MAX(YEARMON) YEARMON FROM MD_OUTDOOR_MEDIUM"
                strFormat = strFormat & "		WHERE 1=1  AND YEARMON >= CONVERT(CHAR(6), DATEADD(MM,-12,GETDATE()),112)"
                strFormat = strFormat & "		group by HIGHSUBSEQ"
                strFormat = strFormat & "		UNION ALL"
                strFormat = strFormat & "		SELECT 'P' MEDFLAG, SUBSEQ, MAX(YEARMON) YEARMON FROM PD_DIVAMT"
                strFormat = strFormat & "		WHERE 1=1  AND YEARMON >= CONVERT(CHAR(6), DATEADD(MM,-12,GETDATE()),112)"
                strFormat = strFormat & "		group by SUBSEQ"
                strFormat = strFormat & "		UNION ALL"
                strFormat = strFormat & "		SELECT 'P2' MEDFLAG, SUBSEQ, max(SUBSTRING(CREDAY,1,6)) YEARMON FROM PD_PONO"
                strFormat = strFormat & "		WHERE 1=1  AND SUBSTRING(CREDAY,1,6) >= CONVERT(CHAR(6), DATEADD(MM,-12,GETDATE()),112)"
                strFormat = strFormat & "		group by SUBSEQ"
                strFormat = strFormat & "	) AAA GROUP BY SUBSEQ"
                strFormat = strFormat & " ) B ON A.SEQNO = B.SUBSEQ"
                strFormat = strFormat & " WHERE 1=1 {0} "


                strSQL = String.Format(strFormat, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)

                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_HIGHSUBSEQ")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    ' =============== SelectRtn_SUBSEQ �귣��
    Public Function SelectRtn_SUBSEQ_SRC(ByVal strInfoXML As String, _
                                         ByRef intRowCnt As Integer, _
                                         ByRef intColCnt As Integer, _
                                         ByVal strCUSTNAME As String, _
                                         ByVal strSEQNAME As String, _
                                         ByVal strHIGHSEQNAME As String, _
                                         ByVal strUSE_YN As String) As Object

        Dim strSQL As String
        Dim strFormat, strWhere As String
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

                If strCUSTNAME <> "" Then Con1 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(CUSTCODE) LIKE '%{0}%')", strCUSTNAME)
                If strSEQNAME <> "" Then Con2 = String.Format(" AND (SEQNAME LIKE '%{0}%')", strSEQNAME)
                If strHIGHSEQNAME <> "" Then Con3 = String.Format(" AND (DBO.SC_HIGHSEQNAME_FUN(HIGHSEQNO) LIKE '%{0}%')", strHIGHSEQNAME)
                If strUSE_YN <> "" Then Con4 = String.Format(" AND (ISNULL(ATTR01,'') = '{0}')", strUSE_YN)

                strWhere = BuildFields(" ", Con1, Con2, Con3, Con4)

                strFormat = "  SELECT "
                strFormat = strFormat & "  0 CHK, SEQNO, "
                strFormat = strFormat & "  SEQNAME, HIGHSEQNO, "
                strFormat = strFormat & "  TIMCODE, '' BTNTIM, DBO.SC_GET_CUSTNAME_FUN(TIMCODE) TIMNAME, "
                strFormat = strFormat & "  CLIENTSUBCODE, '' BTNSUB, DBO.SC_GET_CUSTNAME_FUN(CLIENTSUBCODE) CLIENTSUBNAME, "
                strFormat = strFormat & "  CUSTCODE, '' BTN, DBO.SC_GET_HIGHCUSTNAME_FUN(CUSTCODE) CUSTNAME, "
                strFormat = strFormat & "  DEPT_CD, '' BTNDEPT,dbo.SC_DEPT_NAME_FUN(DEPT_CD)  DEPT_NAME, DBO.SC_EMPNAME_FUN(CUSER) CUSER, CDATE,"
                strFormat = strFormat & "  MEMO, case isnull(ATTR01,'N') when 'N' then '�̻��' WHEN 'Y' THEN '���' WHEN 'S' THEN '���ο�û' ELSE '���' end as Attr01 "
                strFormat = strFormat & "  FROM SC_SUBSEQ_DTL"
                strFormat = strFormat & "  WHERE 1=1 {0} "

                strSQL = String.Format(strFormat, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)

                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_HIGHSUBSEQ")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function


    ' =============== SelectRtn_CUSTDTL �����ֵ�����
    Public Function SelectRtn_CUSTDTL(ByVal strInfoXML As String, _
                                      ByRef intRowCnt As Integer, _
                                      ByRef intColCnt As Integer, _
                                      ByRef strHIGHCUSTCODE As String) As Object     'XML  ������ ��ȸ��

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

                If strHIGHCUSTCODE <> "" Then Con1 = String.Format(" AND (HIGHCUSTCODE = '{0}')", strHIGHCUSTCODE)

                strWhere = BuildFields(" ", Con1)

                strSelFields = " CASE GBNFLAG WHEN '0' THEN '��' WHEN '1' THEN 'CIC/�����' ELSE '' END GBNFLAG, "
                strSelFields = strSelFields & " CLIENTSUBCODE, '' BTN, DBO.SC_GET_CUSTNAME_FUN(CLIENTSUBCODE) AS CLIENTSUBNAME, "
                strSelFields = strSelFields & " CUSTNAME, CUSTCODE, "
                strSelFields = strSelFields & " HIGHCUSTCODE, '' BTNHIGH, DBO.SC_GET_HIGHCUSTNAME_FUN(HIGHCUSTCODE) COMPANYNAME,"
                strSelFields = strSelFields & " USE_FLAG "

                strFormet = "select {0} from SC_CUST_DTL where 1=1 AND MEDFLAG = 'A' {1} "

                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)

                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_CUSTDTL")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    Public Function DeleteRtn_HDR(ByVal strInfoXML As String, _
                                   ByVal strHIGHSEQNO As String) As Integer   '������ DELETE

        Dim intRtn_desc As Integer      'Return����( ó���Ǽ� �Ǵ� 0 )
        Dim intRtn As Integer      'Return����( ó���Ǽ� �Ǵ� 0 )

        SetConfig(strInfoXML)    '�⺻���� Setting
        With mobjSCGLConfig    '�⺻���� Config ��ü
            Try
                ' �����Entity ��ü����(Config ������ �Ѱܻ���)
                mobjceSC_SUBSEQ_HDR = New ceSC_SUBSEQ_HDR(mobjSCGLConfig)
                ' DB ���� �� Ʈ����� ����
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                ' ��ƼƼ ������Ʈ�� Delete �޼ҵ� ȣ��
                intRtn = mobjceSC_SUBSEQ_HDR.DeleteDo(strHIGHSEQNO)
                ' Ʈ����� Commit
                .mobjSCGLSql.SQLCommitTrans()
                Return intRtn
            Catch err As Exception
                'Ʈ����� RollBack �� ���� ����
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & "DeleteRtn_HDR")
            Finally
                'DB���� ����
                .mobjSCGLSql.SQLDisconnect()
                '����� Entity(��üDispose)
                mobjceSC_SUBSEQ_HDR.Dispose()
            End Try
        End With
    End Function

    Public Function DeleteRtn_DTL(ByVal strInfoXML As String, _
                                  ByVal strSEQNO As String) As Integer   '������ DELETE

        Dim intRtn_desc As Integer      'Return����( ó���Ǽ� �Ǵ� 0 )
        Dim intRtn As Integer      'Return����( ó���Ǽ� �Ǵ� 0 )

        SetConfig(strInfoXML)    '�⺻���� Setting
        With mobjSCGLConfig    '�⺻���� Config ��ü
            Try
                ' �����Entity ��ü����(Config ������ �Ѱܻ���)
                mobjceSC_SUBSEQ_DTL = New ceSC_SUBSEQ_DTL(mobjSCGLConfig)
                ' DB ���� �� Ʈ����� ����
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                ' ��ƼƼ ������Ʈ�� Delete �޼ҵ� ȣ��
                intRtn = mobjceSC_SUBSEQ_DTL.DeleteDo(strSEQNO)
                ' Ʈ����� Commit
                .mobjSCGLSql.SQLCommitTrans()
                Return intRtn
            Catch err As Exception
                'Ʈ����� RollBack �� ���� ����
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & "DeleteRtn_DTL")
            Finally
                'DB���� ����
                .mobjSCGLSql.SQLDisconnect()
                '����� Entity(��üDispose)
                mobjceSC_SUBSEQ_DTL.Dispose()
            End Try
        End With
    End Function

    Public Function ProcessRtn_CONF(ByVal strInfoXML As String, _
                                    ByVal strSEQNO As String) As Integer   '������ DELETE

        Dim intRtn_desc As Integer      'Return����( ó���Ǽ� �Ǵ� 0 )
        Dim intRtn As Integer      'Return����( ó���Ǽ� �Ǵ� 0 )

        SetConfig(strInfoXML)    '�⺻���� Setting
        With mobjSCGLConfig    '�⺻���� Config ��ü
            Try
                ' �����Entity ��ü����(Config ������ �Ѱܻ���)
                mobjceSC_SUBSEQ_DTL = New ceSC_SUBSEQ_DTL(mobjSCGLConfig)
                ' DB ���� �� Ʈ����� ����
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                ' ��ƼƼ ������Ʈ�� Delete �޼ҵ� ȣ��
                intRtn = mobjceSC_SUBSEQ_DTL.Update_Conf(strSEQNO)
                ' Ʈ����� Commit
                .mobjSCGLSql.SQLCommitTrans()
                Return intRtn
            Catch err As Exception
                'Ʈ����� RollBack �� ���� ����
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & "DeleteRtn_DTL")
            Finally
                'DB���� ����
                .mobjSCGLSql.SQLDisconnect()
                '����� Entity(��üDispose)
                mobjceSC_SUBSEQ_DTL.Dispose()
            End Try
        End With
    End Function

    Public Function ProcessRtn_CONFOK(ByVal strInfoXML As String, _
                                      ByVal strSEQNO As String) As Integer   '������ DELETE

        Dim intRtn_desc As Integer      'Return����( ó���Ǽ� �Ǵ� 0 )
        Dim intRtn As Integer      'Return����( ó���Ǽ� �Ǵ� 0 )

        SetConfig(strInfoXML)    '�⺻���� Setting
        With mobjSCGLConfig    '�⺻���� Config ��ü
            Try
                ' �����Entity ��ü����(Config ������ �Ѱܻ���)
                mobjceSC_SUBSEQ_DTL = New ceSC_SUBSEQ_DTL(mobjSCGLConfig)
                ' DB ���� �� Ʈ����� ����
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                ' ��ƼƼ ������Ʈ�� Delete �޼ҵ� ȣ��
                intRtn = mobjceSC_SUBSEQ_DTL.Update_ConfOK(strSEQNO)
                ' Ʈ����� Commit
                .mobjSCGLSql.SQLCommitTrans()
                Return intRtn
            Catch err As Exception
                'Ʈ����� RollBack �� ���� ����
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & "DeleteRtn_DTL")
            Finally
                'DB���� ����
                .mobjSCGLSql.SQLDisconnect()
                '����� Entity(��üDispose)
                mobjceSC_SUBSEQ_DTL.Dispose()
            End Try
        End With
    End Function



    ' =============== ProcessRtn_HIGHSUBSEQ    ��ǥ�귣�� �ش� ����
    Public Function ProcessRtn_HIGHSUBSEQ(ByVal strInfoXML As String, _
                                          ByVal vntData As Object, _
                                          ByVal strYEAR As String) As Object

        Dim intRtn As Integer
        Dim intRtn2 As Integer
        Dim i, intColCnt, intRows As Integer
        Dim strHIGHSEQNO

        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                If IsArray(vntData) Then
                    '''����� Entity ��ü����(Config ������ �Ѱܻ���)
                    mobjceSC_SUBSEQ_HDR = New ceSC_SUBSEQ_HDR(mobjSCGLConfig)
                    '''vntData�� �÷���, �ο���� �����Է�
                    intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)
                    For i = 1 To intRows
                        strHIGHSEQNO = ""

                        If GetElement(vntData, "HIGHSEQNO", intColCnt, i, OPTIONAL_STR) = "" Then
                            strHIGHSEQNO = Get_NewHighSeqNo(strYEAR)
                            intRtn = InsertRtn_SC_SUBSEQ_HDR(vntData, intColCnt, i, strHIGHSEQNO)
                        Else
                            strHIGHSEQNO = GetElement(vntData, "HIGHSEQNO", intColCnt, i, OPTIONAL_STR)
                            intRtn = UpdateRtn_SC_SUBSEQ_HDR(vntData, intColCnt, i, strHIGHSEQNO)
                        End If
                    Next
                End If
                .mobjSCGLSql.SQLCommitTrans()
                Return intRows
            Catch err As Exception
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".ProcessRtn_HIGHSUBSEQ")
            Finally
                .mobjSCGLSql.SQLDisconnect()
                mobjceSC_SUBSEQ_HDR.Dispose()
            End Try
        End With
    End Function

    ' =============== ProcessRtn_SUBSEQ    �귣�� �ش� ����
    Public Function ProcessRtn_SUBSEQ(ByVal strInfoXML As String, _
                                      ByVal vntData As Object, _
                                      ByVal strYEAR As String) As Object

        Dim intRtn As Integer
        Dim intRtn2 As Integer
        Dim i, intColCnt, intRows As Integer
        Dim strSEQNO
        Dim strRETURNVALUE
        Dim strATTR01

        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                If IsArray(vntData) Then
                    '''����� Entity ��ü����(Config ������ �Ѱܻ���)
                    mobjceSC_SUBSEQ_DTL = New ceSC_SUBSEQ_DTL(mobjSCGLConfig)
                    '''vntData�� �÷���, �ο���� �����Է�
                    intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)
                    For i = 1 To intRows
                        strSEQNO = ""
                        strATTR01 = ""

                        If GetElement(vntData, "SEQNO", intColCnt, i, OPTIONAL_STR) = "" Then
                            strSEQNO = Get_NewSeqNo(strYEAR)
                            strATTR01 = GetElement(vntData, "ATTR01", intColCnt, i, OPTIONAL_STR)
                            If strATTR01 = "���" Then
                                strATTR01 = "Y"
                            ElseIf strATTR01 = "�̻��" Then
                                strATTR01 = "N"
                            ElseIf strATTR01 = "�̻��" Then
                                strATTR01 = "N"
                            ElseIf strATTR01 = "���ο�û" Then
                                strATTR01 = "S"
                            Else
                                strATTR01 = "R"
                            End If
                            intRtn = InsertRtn_SC_SUBSEQ_DTL(vntData, intColCnt, i, strSEQNO, strATTR01)
                            strRETURNVALUE = intRtn & "-" & strSEQNO
                        Else
                            strSEQNO = GetElement(vntData, "SEQNO", intColCnt, i, OPTIONAL_STR)
                            strATTR01 = GetElement(vntData, "ATTR01", intColCnt, i, OPTIONAL_STR)
                            If strATTR01 = "���" Then
                                strATTR01 = "Y"
                            ElseIf strATTR01 = "�̻��" Then
                                strATTR01 = "N"
                            ElseIf strATTR01 = "���ο�û" Then
                                strATTR01 = "S"
                            Else
                                strATTR01 = "R"
                            End If

                            intRtn = UpdateRtn_SC_SUBSEQ_DTL(vntData, intColCnt, i, strSEQNO, strATTR01)
                            strRETURNVALUE = intRtn & "-" & strSEQNO
                        End If
                    Next
                End If
                .mobjSCGLSql.SQLCommitTrans()
                Return strRETURNVALUE
            Catch err As Exception
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".ProcessRtn_SUBSEQ")
            Finally
                .mobjSCGLSql.SQLDisconnect()
                mobjceSC_SUBSEQ_DTL.Dispose()
            End Try
        End With
    End Function

    '==============SC_SUBSEQ_HDR ���̺��� �ű� HIGHSEQNO ��������
    Public Function Get_NewHighSeqNo(ByVal strYEAR As String) As String
        Dim strSQL, strFormat, strRtn As String

        With mobjSCGLConfig '�⺻���� Config ��ü
            Try
                strSQL = "select 'S' +'" & strYEAR & "' + DBO.LPAD(ISNULL(MAX(CAST(SUBSTRING(HIGHSEQNO,4,5) AS NUMERIC(5,0))),0)+1,5,'0') From SC_SUBSEQ_HDR "
                strRtn = .mobjSCGLSql.SQLSelectOneScalar(strSQL)
                Return strRtn
            Catch err As Exception
                ' ���� ����
                Throw RaiseSysErr(err, CLASS_NAME & ".Get_NewHighSeqNo")
            End Try
        End With
    End Function

    '==============SC_SUBSEQ_DTL ���̺��� �ű� SEQNO ��������
    Public Function Get_NewSeqNo(ByVal strYEAR As String) As String
        Dim strSQL, strFormat, strRtn As String

        With mobjSCGLConfig '�⺻���� Config ��ü
            Try
                strSQL = "select 'S' +'" & strYEAR & "' + DBO.LPAD(ISNULL(MAX(CAST(SUBSTRING(SEQNO,4,5) AS NUMERIC(5,0))),0)+1,5,'0') From SC_SUBSEQ_DTL "
                strRtn = .mobjSCGLSql.SQLSelectOneScalar(strSQL)
                Return strRtn
            Catch err As Exception
                ' ���� ����
                Throw RaiseSysErr(err, CLASS_NAME & ".Get_NewSeqNo")
            End Try
        End With
    End Function

    '�ű� CUSTCODE ����
    Public Function SelectRtn_CUSTCODE(ByVal strMEDFLAG As String) As String

        Dim strSQL As String
        Dim strFormat As String
        Dim strRtn As String

        With mobjSCGLConfig '�⺻���� Config ��ü

            Try
                strSQL = String.Format("select '{0}' + dbo.lpad(isnull(Max(substring(custcode,2,6)),0)+1,5,0) From SC_CUST_DTL WHERE MEDFLAG =  '{1}'", strMEDFLAG, strMEDFLAG)
                strRtn = .mobjSCGLSql.SQLSelectOneScalar(strSQL)
                Return strRtn
            Catch err As Exception
                ' ���� ����
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_CUSTCODE")
            Finally
            End Try
        End With
    End Function
#End Region

#Region "GROUP BLOCK : �ܺο� ����� Method"
    Private Function InsertRtn_SC_SUBSEQ_HDR(ByVal vntData As Object, _
                                             ByVal intColCnt As Integer, _
                                             ByVal intRow As Integer, _
                                             ByVal strHIGHSEQNO As String) As Integer

        Dim intRtn As Integer
        intRtn = mobjceSC_SUBSEQ_HDR.InsertDo( _
                                       strHIGHSEQNO, _
                                       GetElement(vntData, "HIGHSEQNAME", intColCnt, intRow), _
                                       GetElement(vntData, "CUSTCODE", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR01", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR02", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR03", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR04", intColCnt, intRow, NULL_NUM, True))
        Return intRtn
    End Function

    Private Function UpdateRtn_SC_SUBSEQ_HDR(ByVal vntData As Object, _
                                             ByVal intColCnt As Integer, _
                                             ByVal intRow As Integer, _
                                             ByVal strHIGHSEQNO As String) As Integer
        Dim intRtn As Integer

        intRtn = mobjceSC_SUBSEQ_HDR.UpdateDo( _
                                       strHIGHSEQNO, _
                                       GetElement(vntData, "HIGHSEQNAME", intColCnt, intRow), _
                                       GetElement(vntData, "CUSTCODE", intColCnt, intRow))

        Return intRtn
    End Function

    Private Function InsertRtn_SC_SUBSEQ_DTL(ByVal vntData As Object, _
                                             ByVal intColCnt As Integer, _
                                             ByVal intRow As Integer, _
                                             ByVal strSEQNO As String, _
                                             ByVal strATTR01 As String) As Integer
        Dim intRtn As Integer
        intRtn = mobjceSC_SUBSEQ_DTL.InsertDo( _
                                       strSEQNO, _
                                       GetElement(vntData, "SEQNAME", intColCnt, intRow), _
                                       GetElement(vntData, "HIGHSEQNO", intColCnt, intRow), _
                                       GetElement(vntData, "CUSTCODE", intColCnt, intRow), _
                                       GetElement(vntData, "CLIENTSUBCODE", intColCnt, intRow), _
                                       GetElement(vntData, "TIMCODE", intColCnt, intRow), _
                                       GetElement(vntData, "DEPT_CD", intColCnt, intRow), _
                                       GetElement(vntData, "MEMO", intColCnt, intRow), _
                                       strATTR01, _
                                       GetElement(vntData, "ATTR02", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR03", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR04", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR05", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR06", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR07", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR08", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR09", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR10", intColCnt, intRow, NULL_NUM, True))

        Return intRtn
    End Function

    Private Function UpdateRtn_SC_SUBSEQ_DTL(ByVal vntData As Object, _
                                             ByVal intColCnt As Integer, _
                                             ByVal intRow As Integer, _
                                             ByVal strSEQNO As String, _
                                             ByVal strATTR01 As String) As Integer
        Dim intRtn As Integer

        intRtn = mobjceSC_SUBSEQ_DTL.UpdateDo( _
                                       strSEQNO, _
                                       GetElement(vntData, "SEQNAME", intColCnt, intRow), _
                                       GetElement(vntData, "HIGHSEQNO", intColCnt, intRow), _
                                       GetElement(vntData, "CUSTCODE", intColCnt, intRow), _
                                       GetElement(vntData, "CLIENTSUBCODE", intColCnt, intRow), _
                                       GetElement(vntData, "TIMCODE", intColCnt, intRow), _
                                       GetElement(vntData, "DEPT_CD", intColCnt, intRow), _
                                       GetElement(vntData, "MEMO", intColCnt, intRow), _
                                       strATTR01)

        Return intRtn
    End Function
#End Region
End Class



