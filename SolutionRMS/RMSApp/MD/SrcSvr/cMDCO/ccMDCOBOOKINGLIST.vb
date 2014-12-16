'****************************************************************************************
'����   ȯ��    : COM+ Service Server Package
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009-07-28 ���� 5:04:13 By KTY
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

Public Class ccMDCOBOOKINGLIST
    Inherits ccControl

#Region "GROUP BLOCK : ���� �Ǵ� ��ⷹ���� ����/��� ����"
    Private CLASS_NAME = "ccMDCOBOOKINGLIST"                        '�ڽ��� Ŭ������
    Private mobjceMD_BOOKING_MEDIUM As eMDCO.ceMD_BOOKING_MEDIUM    '����� Entity ���� ����
    Private mobjceMD_CATV_MEDIUM As eMDCO.ceMD_CATV_MEDIUM          '����� Entity ���� ����
    Private mobjceMD_INTERNET_MEDIUM As eMDCO.ceMD_INTERNET_MEDIUM  '����� Entity ���� ����
    Private mobjceMD_OUTDOOR_MEDIUM As eMDCO.ceMD_OUTDOOR_MEDIUM    '����� Entity ���� ����
    Private mobjceMD_TOTAL_MEDIUM As eMDCO.ceMD_TOTAL_MEDIUM        '����� Entity ���� ����
#End Region

#Region "GROUP BLOCK : Event ����"

    '********************************************************
    ' GetDataType()  ��ü���� �޺� select ó��
    '********************************************************
    Public Function GetDataType(ByVal strInfoXML As String, _
                                ByRef intRowCnt As Integer, _
                                ByRef intColCnt As Integer) As Object

        Dim strSQL, strFormat, strSelFields As String
        Dim vntData As Object

        SetConfig(strInfoXML)   '�⺻���� ����

        '��ȸ �ʵ� ����
        strSelFields = "CODE, CODE_NAME"

        'SQL�� ����

        strFormat = "SELECT {0} " & _
                    "FROM SC_CODE " & _
                    "WHERE CLASS_CODE = 'MP_KIND' " & _
                    "ORDER BY SORT_SEQ "

        With mobjSCGLConfig
            strSQL = String.Format(strFormat, strSelFields)

            ''������ ��ȸ
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

#Region "����û�೻�� ��ȸ/����"
    ' =============== ��ŷ û�೻����ȸ
    Public Function SelectRtn_PRINT(ByVal strInfoXML As String, _
                                    ByRef intRowCnt As Integer, _
                                    ByRef intColCnt As Integer, _
                                    ByVal strYEARMON As String, _
                                    ByVal strCLIENTCODE As String, _
                                    ByVal strCLIENTNAME As String, _
                                    ByVal strREAL_MED_CODE As String, _
                                    ByVal strREAL_MED_NAME As String, _
                                    ByVal strTIMCODE As String, _
                                    ByVal strTIMNAME As String, _
                                    ByVal strMEDCODE As String, _
                                    ByVal strMEDNAME As String, _
                                    ByVal strSUBSEQ As String, _
                                    ByVal strSUBSEQNAME As String, _
                                    ByVal strMEDFLAG As String, _
                                    ByVal strGFLAG As String, _
                                    ByVal strFPUB_DATE As String, _
                                    ByVal strTPUB_DATE As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim Con1, Con2, Con3, Con4, Con5, Con6, Con7, Con8, Con9, Con10, Con11, Con12, Con13, Con14, Con15 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = "" : Con2 = "" : Con3 = "" : Con4 = "" : Con5 = ""
                Con6 = "" : Con7 = "" : Con8 = "" : Con9 = "" : Con10 = ""
                Con11 = "" : Con12 = "" : Con13 = "" : Con14 = "" : Con15 = ""

                If strYEARMON <> "" Then Con1 = String.Format(" AND (SUBSTRING(DEMANDDAY,1,6) = '{0}')", strYEARMON) '���
                If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (CLIENTCODE = '{0}')", strCLIENTCODE) '�������ڵ�
                If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCLIENTNAME) '�����ָ�
                If strREAL_MED_CODE <> "" Then Con4 = String.Format(" AND (REAL_MED_CODE = '{0}')", strREAL_MED_CODE) '��ü���ڵ�
                If strREAL_MED_NAME <> "" Then Con5 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) LIKE '%{0}%')", strREAL_MED_NAME) '��ü���
                If strTIMCODE <> "" Then Con6 = String.Format(" AND (TIMCODE = '{0}')", strTIMCODE) '���ڵ�
                If strTIMNAME <> "" Then Con7 = String.Format(" AND (DBO.SC_GET_CUSTNAME_FUN(TIMCODE) LIKE '%{0}%')", strTIMNAME) '����
                If strMEDCODE <> "" Then Con8 = String.Format(" AND (MEDCODE = '{0}')", strMEDCODE) '��ü�ڵ�
                If strMEDNAME <> "" Then Con9 = String.Format(" AND (DBO.SC_GET_CUSTNAME_FUN(MEDCODE) LIKE '%{0}%')", strMEDNAME) '��ü��
                If strSUBSEQ <> "" Then Con10 = String.Format(" AND (SUBSEQ = '{0}')", strSUBSEQ) '�귣���ڵ�
                If strSUBSEQNAME <> "" Then Con11 = String.Format(" AND (DBO.SC_GET_SUBSEQNAME_FUN(SUBSEQ) LIKE '%{0}%')", strSUBSEQNAME) '�귣���
                If strMEDFLAG <> "" Then Con12 = String.Format(" AND (MED_FLAG = '{0}')", strMEDFLAG) '��ü����
                If strGFLAG <> "" Then Con13 = String.Format(" AND (GFLAG = '{0}')", strGFLAG) '���౸��

                If strFPUB_DATE <> "" Then
                    strFPUB_DATE = Replace(strFPUB_DATE, "-", "")
                    Con14 = String.Format(" AND (PUB_DATE >= '{0}')", strFPUB_DATE) '���౸��
                Else
                    Con14 = String.Format(" AND (PUB_DATE >= '{0}')", "00000000") '���౸��
                End If

                If strTPUB_DATE <> "" Then
                    strTPUB_DATE = Replace(strTPUB_DATE, "-", "")
                    Con15 = String.Format(" AND (PUB_DATE <= '{0}')", strTPUB_DATE) '���౸��
                Else
                    Con15 = String.Format(" AND (PUB_DATE <= '{0}')", "99999999") '���౸��
                End If

                strWhere = BuildFields(" ", Con1, Con2, Con3, Con4, Con5, Con6, Con7, Con8, Con9, Con10, Con11, Con12, Con13, Con14, Con15)

                If strMEDFLAG = "MP01" Then
                    strFormet = " SELECT "
                    strFormet = strFormet & " 0 CHK, "
                    strFormet = strFormet & " CASE ISNULL(GFLAG,'') WHEN 'M' THEN '�̽���'  WHEN 'B' THEN '����' ELSE  '���Ϸ�' END AS GFLAGNAME, "
                    strFormet = strFormet & " DBO.MD_TRANS_YN_FUN(A.YEARMON,A.SEQ, 'B') CONFIRMFLAG, "
                    strFormet = strFormet & " A.YEARMON, A.SEQ, "
                    strFormet = strFormet & " DISPPUB_DATE, MEDNAME, CLIENTNAME, TIMNAME, MATTERNAME, STD, COL_DEG, PRICE, AMT, "
                    strFormet = strFormet & " COMMI_RATE, PUB_FACENAME, EXECUTE_FACE, DELIVER_NAME, CONTACT_FLAGNAME, TRU_TRANS_NO, MEMO, REAL_MED_NAME,"
                    strFormet = strFormet & " EXCLIENTCODE, EXCLIENTNAME, VOCH_TYPE"
                    strFormet = strFormet & " FROM ( "
                    strFormet = strFormet & "   SELECT "
                    strFormet = strFormet & "   YEARMON, SEQ,  GFLAG, "
                    strFormet = strFormet & "   (SUBSTRING(PUB_DATE,5,2) + '-' + SUBSTRING(PUB_DATE,7,2)) DISPPUB_DATE,  "
                    strFormet = strFormet & "   DBO.SC_GET_CUSTNAME_FUN(MEDCODE) MEDNAME,"
                    strFormet = strFormet & "   DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME,  "
                    strFormet = strFormet & "   DBO.SC_GET_CUSTNAME_FUN(TIMCODE) TIMNAME,  "
                    strFormet = strFormet & "   DBO.MD_GET_MATTERNAME_FUN(MATTERCODE) MATTERNAME, "
                    strFormet = strFormet & "   CAST(STD_STEP AS NVARCHAR(10)) + '��' + CAST(STD_CM AS NVARCHAR(10)) + 'CM' + ISNULL(CAST(STD_FACE AS NVARCHAR(10)),'')  STD, "
                    strFormet = strFormet & "   COL_DEG,  PRICE, AMT, COMMI_RATE, PUB_FACE PUB_FACENAME, "
                    strFormet = strFormet & "   EXECUTE_FACE, TRU_TRANS_NO, MEMO, CLIENTCODE, PUB_DATE, DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) REAL_MED_NAME,"
                    strFormet = strFormet & "   EXCLIENTCODE,"
                    strFormet = strFormet & "   DBO.SC_GET_HIGHCUSTNAME_FUN(EXCLIENTCODE) EXCLIENTNAME, "
                    strFormet = strFormet & "   VOCH_TYPE"
                    strFormet = strFormet & "   FROM MD_BOOKING_MEDIUM "
                    strFormet = strFormet & "   WHERE 1=1  "
                    strFormet = strFormet & "   {0} "
                    strFormet = strFormet & " )  "
                    strFormet = strFormet & " A LEFT JOIN   "
                    strFormet = strFormet & " (  "
                    strFormet = strFormet & "   SELECT YEARMON, SEQ, DELIVER_NAME,   "
                    strFormet = strFormet & "   CASE CONTACT_FLAG WHEN 'Y' THEN '��' WHEN 'N' THEN '��' ELSE '' END CONTACT_FLAGNAME  "
                    strFormet = strFormet & "   FROM MD_WONGO_MST  "
                    strFormet = strFormet & " ) B  "
                    strFormet = strFormet & " ON A.YEARMON = B.YEARMON "
                    strFormet = strFormet & " AND A.SEQ = B.SEQ  "
                    strFormet = strFormet & " ORDER BY CASE WHEN substring(dbo.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE),1,3) = '(��)' THEN "
                    strFormet = strFormet & " substring(dbo.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE),4,100) ELSE dbo.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) END, PUB_DATE"

                Else
                    strFormet = " select "
                    strFormet = strFormet & " 0 CHK,"
                    strFormet = strFormet & " CASE ISNULL(GFLAG,'') WHEN 'M' THEN '�̽���'  WHEN 'B' THEN '����' ELSE  '���Ϸ�' END AS GFLAGNAME, "
                    strFormet = strFormet & " DBO.MD_TRANS_YN_FUN(A.YEARMON,A.SEQ, 'B') CONFIRMFLAG,"
                    strFormet = strFormet & " A.YEARMON, A.SEQ,"
                    strFormet = strFormet & " CLIENTNAME, MATTERNAME, MEDNAME, STD, DISPPUB_DATE, DISPPUB_DATE1, "
                    strFormet = strFormet & " STD_PAGE, AMT, REAL_MED_NAME, COMMI_RATE, BOOKING, GUBUN_NAME, DELIVER_NAME, "
                    strFormet = strFormet & " OUTFLAG, PUB_FACENAME,TRU_TRANS_NO,CONTACT_FLAGNAME, MEMO, EXCLIENTCODE, EXCLIENTNAME, VOCH_TYPE"
                    strFormet = strFormet & " FROM ("
                    strFormet = strFormet & "   select"
                    strFormet = strFormet & "   YEARMON, SEQ, GFLAG,"
                    strFormet = strFormet & "   dbo.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME, "
                    strFormet = strFormet & "   DBO.MD_GET_MATTERNAME_FUN(MATTERCODE) MATTERNAME, "
                    strFormet = strFormet & "   dbo.SC_GET_CUSTNAME_FUN(MEDCODE)  MEDNAME, "
                    strFormet = strFormet & "   STD, "
                    strFormet = strFormet & "   (substring(PUB_DATE,5,2) + '-' + substring(PUB_DATE,7,2)) DISPPUB_DATE, "
                    strFormet = strFormet & "   (substring(PUB_DATE,5,2) + '-' + substring(PUB_DATE,7,2)) DISPPUB_DATE1, "
                    strFormet = strFormet & "   STD_PAGE, "
                    strFormet = strFormet & "   AMT, "
                    strFormet = strFormet & "   DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) REAL_MED_NAME, "
                    strFormet = strFormet & "   COMMI_RATE, "
                    strFormet = strFormet & "   GFLAG BOOKING, "
                    strFormet = strFormet & "   PUB_FACE PUB_FACENAME,"
                    strFormet = strFormet & "   TRU_TRANS_NO,  "
                    strFormet = strFormet & "   MEMO, CLIENTCODE, PUB_DATE, EXCLIENTCODE,"
                    strFormet = strFormet & "   DBO.SC_GET_HIGHCUSTNAME_FUN(EXCLIENTCODE) EXCLIENTNAME, "
                    strFormet = strFormet & "   VOCH_TYPE "
                    strFormet = strFormet & "   FROM MD_BOOKING_MEDIUM where 1=1 "
                    strFormet = strFormet & "   {0} "
                    strFormet = strFormet & " ) A LEFT JOIN ( "
                    strFormet = strFormet & "   SELECT "
                    strFormet = strFormet & "   YEARMON, SEQ, DELIVER_NAME,"
                    strFormet = strFormet & "   CASE GUBUN WHEN 'N' THEN '��' WHEN 'O' THEN '��' ELSE '' END GUBUN_NAME,"
                    strFormet = strFormet & "   CASE CONTACT_FLAG WHEN 'Y' THEN '��' WHEN 'N' THEN '��' ELSE '' END CONTACT_FLAGNAME,  "
                    strFormet = strFormet & "   OUTFLAG "
                    strFormet = strFormet & "   FROM MD_WONGO_MST "
                    strFormet = strFormet & " ) B  "
                    strFormet = strFormet & " ON A.YEARMON = B.YEARMON "
                    strFormet = strFormet & " AND A.SEQ = B.SEQ "
                    strFormet = strFormet & " ORDER BY CASE WHEN substring(dbo.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE),1,3) = '(��)' THEN "
                    strFormet = strFormet & " substring(dbo.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE),4,100) ELSE dbo.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) END, PUB_DATE"

                End If


                strSQL = String.Format(strFormet, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)

                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)

                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_PRINT")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    '============== ���ι� ��� ó�� 1
    Public Function ProcessRtn_ConfirmBooking_OK(ByVal strInfoXML As String, _
                                                 ByVal vntData As Object, _
                                                 ByVal strFLAG As String) As Integer
        Dim intRtn As Integer '����� ����
        Dim i, intColCnt, intRows As Integer '����, �÷�Cnt, �ο�Cnt ����

        SetConfig(strInfoXML) '�⺻���� Setting

        With mobjSCGLConfig '�⺻������ ������ �ִ� Config ��ü
            Try
                'XML Element ���� ���� (strMasterXML�� ��ȯ)
                Dim xmlRoot As XmlElement
                'xmlRoot = XMLGetRoot(strMasterXML) 'XML ������

                'DB���� �� Ʈ����� ����
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()

                If IsArray(vntData) Then
                    intRtn = strDETAIL_DIVAMTBOOKING(strInfoXML, vntData, strFLAG)
                End If

                'Ʈ�����Commit
                .mobjSCGLSql.SQLCommitTrans()
                Return intRtn
            Catch err As Exception
                'Ʈ�����RollBack �� ���� ����
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".ProcessRtn_ConfirmBooking_OK")
            Finally
                'Resource����
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    '============== ���� �� ��� ó��2
    Public Function strDETAIL_DIVAMTBOOKING(ByVal strInfoXML As String, _
                                            ByVal vntData As Object, _
                                            ByVal strFLAG As String) As Integer '������ INSERT/UPDATE

        Dim intRtn, intRtn2 As Integer
        Dim i, intColCnt, intRows, intSEQ As Integer
        Dim dblID As Double '�ڵ� ID ������� ���� ���
        Dim strGFLAG

        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try
                If strFLAG = "CONFIRM" Then
                    strGFLAG = "B"
                Else
                    strGFLAG = "M"
                End If
                If IsArray(vntData) Then
                    '''����� Entity ��ü����(Config ������ �Ѱܻ���)
                    'ceMD_BOOKING_MEDIUM
                    mobjceMD_BOOKING_MEDIUM = New ceMD_BOOKING_MEDIUM(mobjSCGLConfig)
                    '''vntData�� �÷���, �ο���� �����Է�
                    intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)
                    '''�ش��ϴ�Row ��ŭ Loop

                    For i = 1 To intRows
                        If GetElement(vntData, "CHK", intColCnt, i, OPTIONAL_STR) = "1" And GetElement(vntData, "CONFIRMFLAG", intColCnt, i, OPTIONAL_STR) = "N" Then
                            intRtn = UpdateRtn_GFLAGBOOKING(vntData, intColCnt, i, strGFLAG)

                        End If
                    Next
                End If

                Return intRtn
            Catch err As Exception

                Throw RaiseSysErr(err, CLASS_NAME & ".strDETAIL_DIVAMTBOOKING")
            Finally
                mobjceMD_BOOKING_MEDIUM.Dispose()
            End Try
        End With
    End Function
#End Region

#Region "���̺�û�೻�� ��ȸ/����"
    ' =============== ���̺� û�೻����ȸ
    Public Function SelectRtn_CATV(ByVal strInfoXML As String, ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                    ByVal strYEARMON As String, _
                                    ByVal strCLIENTCODE As String, ByVal strCLIENTNAME As String, _
                                    ByVal strREAL_MED_CODE As String, ByVal strREAL_MED_NAME As String, _
                                    ByVal strTIMCODE As String, ByVal strTIMNAME As String, _
                                    ByVal strMEDCODE As String, ByVal strMEDNAME As String, _
                                    ByVal strSUBSEQ As String, ByVal strSUBSEQNAME As String) As Object

        Dim strSQL As String
        Dim strFormet, strWhere As String
        Dim Con1, Con2, Con3, Con4, Con5, Con6, Con7, Con8, Con9, Con10, Con11, Con12, Con13 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = "" : Con2 = "" : Con3 = "" : Con4 = "" : Con5 = "" : Con6 = ""
                Con7 = "" : Con8 = "" : Con9 = "" : Con10 = "" : Con11 = "" : Con12 = "" : Con13 = ""

                If strYEARMON <> "" Then Con1 = String.Format(" AND (DEMANDDAY LIKE '{0}%')", strYEARMON) '���
                If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (CLIENTCODE = '{0}')", strCLIENTCODE) '�������ڵ�
                If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCLIENTNAME) '�����ָ�
                If strREAL_MED_CODE <> "" Then Con4 = String.Format(" AND (REAL_MED_CODE = '{0}')", strREAL_MED_CODE) '��ü���ڵ�
                If strREAL_MED_NAME <> "" Then Con5 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) LIKE '%{0}%')", strREAL_MED_NAME) '��ü���
                If strTIMCODE <> "" Then Con6 = String.Format(" AND (TIMCODE = '{0}')", strTIMCODE) '���ڵ�
                If strTIMNAME <> "" Then Con7 = String.Format(" AND (DBO.SC_GET_CUSTNAME_FUN(TIMCODE) LIKE '%{0}%')", strTIMNAME) '����
                If strMEDCODE <> "" Then Con8 = String.Format(" AND (MEDCODE = '{0}')", strMEDCODE) '��ü�ڵ�
                If strMEDNAME <> "" Then Con9 = String.Format(" AND (DBO.SC_GET_CUSTNAME_FUN(MEDCODE) LIKE '%{0}%')", strMEDNAME) '��ü��
                If strSUBSEQ <> "" Then Con10 = String.Format(" AND (SUBSEQ = '{0}')", strSUBSEQ) '�귣���ڵ�
                If strSUBSEQNAME <> "" Then Con11 = String.Format(" AND (DBO.SC_GET_SUBSEQNAME_FUN(SUBSEQ) LIKE '%{0}%')", strSUBSEQNAME) '�귣���

                strWhere = BuildFields(" ", Con1, Con2, Con3, Con4, Con5, Con6, Con7, Con8, Con9, Con10, Con11, Con12, Con13)

                strFormet = strFormet & " SELECT "
                strFormet = strFormet & " 0 CHK, "
                strFormet = strFormet & " YEARMON, SEQ, GFLAG, "
                strFormet = strFormet & " CASE ISNULL(GFLAG,'') WHEN '' THEN '�̽���' WHEN '0' THEN '�̽���' ELSE CASE ISNULL(TRU_TRANS_NO,'') WHEN '' THEN '����' ELSE '���Ϸ�' END  END AS GFLAGNAME, "
                strFormet = strFormet & " DEMANDDAY, "
                strFormet = strFormet & " CLIENTCODE, "
                strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME,"
                strFormet = strFormet & " MEDCODE, "
                strFormet = strFormet & " DBO.SC_GET_CUSTNAME_FUN(MEDCODE) MEDNAME,"
                strFormet = strFormet & " REAL_MED_CODE,"
                strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) REAL_MED_NAME, "
                strFormet = strFormet & " SUBSEQ,"
                strFormet = strFormet & " DBO.SC_GET_SUBSEQNAME_FUN(SUBSEQ) SUBSEQNAME, "
                strFormet = strFormet & " TIMCODE,"
                strFormet = strFormet & " DBO.SC_GET_CUSTNAME_FUN(TIMCODE) TIMNAME, "
                strFormet = strFormet & " MATTERCODE,"
                strFormet = strFormet & " DBO.MD_GET_MATTERNAME_FUN(MATTERCODE) MATTERNAME, "
                strFormet = strFormet & " DEPT_CD,"
                strFormet = strFormet & " DBO.SC_DEPT_NAME_FUN(DEPT_CD) DEPT_NAME,"
                strFormet = strFormet & " EXCLIENTCODE, "
                strFormet = strFormet & " dbo.SC_GET_EXCLIENTALLNAME_FUN(EXCLIENTCODE) EXCLIENTNAME, "
                strFormet = strFormet & " GREATCODE, "
                strFormet = strFormet & " dbo.SC_GET_GREATCUSTNAME_FUN(GREATCODE) GREATNAME,"
                strFormet = strFormet & " MPP AS MPP_CODE, "
                strFormet = strFormet & " dbo.SC_GET_HIGHCUSTNAME_FUN(MPP) MPP_NAME,"
                strFormet = strFormet & " CLIENTSUBCODE, "
                strFormet = strFormet & " DBO.SC_GET_CUSTNAME_FUN(CLIENTSUBCODE) CLIENSUBMNAME,  "
                strFormet = strFormet & " PROGRAM, "
                strFormet = strFormet & " CNT, "
                strFormet = strFormet & " AMT, COMMI_RATE, COMMISSION, "
                strFormet = strFormet & " TBRDSTDATE, TBRDEDDATE, "
                strFormet = strFormet & " VOCH_TYPE, "
                strFormet = strFormet & " TRU_TAX_FLAG,COMMI_TAX_FLAG, "
                strFormet = strFormet & " MEMO AS BIGO, "
                strFormet = strFormet & " CASE WHEN ISNULL(TRU_TRANS_NO,'') <> '' THEN "
                strFormet = strFormet & " SUBSTRING(TRU_TRANS_NO, 0, DBO.INSTR(0, TRU_TRANS_NO,'-')) + '-' + "
                strFormet = strFormet & " dbo.lpad(SUBSTRING(TRU_TRANS_NO, DBO.INSTR(0, TRU_TRANS_NO,'-')+1, DBO.INSTR(0, SUBSTRING(TRU_TRANS_NO,8,LEN(TRU_TRANS_NO)),'-')-1), 4, 0) + '-' + "
                strFormet = strFormet & " dbo.lpad(SUBSTRING(TRU_TRANS_NO, DBO.INSTR(0, TRU_TRANS_NO,'-')+1 + DBO.INSTR(0, SUBSTRING(TRU_TRANS_NO,8,LEN(TRU_TRANS_NO)),'-'), LEN(TRU_TRANS_NO)),4,0) "
                strFormet = strFormet & " ELSE "
                strFormet = strFormet & " TRU_TRANS_NO "
                strFormet = strFormet & " END AS TRU_TRANS_NO , "
                strFormet = strFormet & " CASE WHEN ISNULL(COMMI_TRANS_NO,'') <> '' THEN "
                strFormet = strFormet & " SUBSTRING(COMMI_TRANS_NO, 0, DBO.INSTR(0, COMMI_TRANS_NO,'-')) + '-' + "
                strFormet = strFormet & " dbo.lpad(SUBSTRING(COMMI_TRANS_NO, DBO.INSTR(0, COMMI_TRANS_NO,'-')+1, DBO.INSTR(0, SUBSTRING(COMMI_TRANS_NO,8,LEN(COMMI_TRANS_NO)),'-')-1), 4, 0) + '-' + "
                strFormet = strFormet & " dbo.lpad(SUBSTRING(COMMI_TRANS_NO, DBO.INSTR(0, COMMI_TRANS_NO,'-')+1 + DBO.INSTR(0, SUBSTRING(COMMI_TRANS_NO,8,LEN(COMMI_TRANS_NO)),'-'), LEN(COMMI_TRANS_NO)),4,0) "
                strFormet = strFormet & " ELSE "
                strFormet = strFormet & " COMMI_TRANS_NO "
                strFormet = strFormet & " END AS COMMI_TRANS_NO"
                strFormet = strFormet & " FROM MD_CATV_MEDIUM "
                strFormet = strFormet & " WHERE 1=1  "
                strFormet = strFormet & " {0}"
                strFormet = strFormet & " ORDER BY CASE WHEN SUBSTRING(DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE),1,3) = '(��)' THEN"
                strFormet = strFormet & " SUBSTRING(DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE),4,100)"
                strFormet = strFormet & " ELSE"
                strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) END"

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

    '============== ���̺� ����ó��
    Public Function ProcessRtn_ConfirmCatv_OK(ByVal strInfoXML As String, _
                                              ByVal vntData As Object, _
                                              ByVal strFLAG As String) As Integer '������ INSERT/UPDATE
        Dim intRtn, intRtn2 As Integer
        Dim i, intColCnt, intRows As Integer
        Dim strGFLAG As String

        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()

                If strFLAG = "CONFIRM" Then
                    strGFLAG = "1"
                Else
                    strGFLAG = "0"
                End If

                If IsArray(vntData) Then
                    '''����� Entity ��ü����(Config ������ �Ѱܻ���)
                    'ceMD_BOOKING_MEDIUM
                    mobjceMD_CATV_MEDIUM = New ceMD_CATV_MEDIUM(mobjSCGLConfig)
                    intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)

                    For i = 1 To intRows
                        If GetElement(vntData, "CHK", intColCnt, i, OPTIONAL_STR) = "1" Then
                            intRtn = UpdateRtn_GFLAGCATV(vntData, intColCnt, i, strGFLAG)
                        End If
                    Next
                End If
                .mobjSCGLSql.SQLCommitTrans()
                Return intRows
            Catch err As Exception
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".ProcessRtn_ConfirmCatv_OK")
            Finally
                .mobjSCGLSql.SQLDisconnect()
                mobjceMD_CATV_MEDIUM.Dispose()
            End Try
        End With
    End Function
#End Region

#Region "������ä�� û�೻�� ��ȸ/����"
    ' =============== ������ä����ȸ
    Public Function SelectRtn_TOTAL(ByVal strInfoXML As String, ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                                    ByVal strYEARMON As String, _
                                    ByVal strCLIENTCODE As String, ByVal strCLIENTNAME As String, _
                                    ByVal strREAL_MED_CODE As String, ByVal strREAL_MED_NAME As String, _
                                    ByVal strTIMCODE As String, ByVal strTIMNAME As String, _
                                    ByVal strMEDCODE As String, ByVal strMEDNAME As String, _
                                    ByVal strSUBSEQ As String, ByVal strSUBSEQNAME As String) As Object

        Dim strSQL As String
        Dim strFormet, strWhere As String
        Dim Con1, Con2, Con3, Con4, Con5, Con6, Con7, Con8, Con9, Con10, Con11 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = "" : Con2 = "" : Con3 = "" : Con4 = "" : Con5 = "" : Con6 = ""
                Con7 = "" : Con8 = "" : Con9 = "" : Con10 = "" : Con11 = ""

                If strYEARMON <> "" Then Con1 = String.Format(" AND (DEMANDDAY LIKE '{0}%')", strYEARMON) '���
                If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (CLIENTCODE = '{0}')", strCLIENTCODE) '�������ڵ�
                If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCLIENTNAME) '�����ָ�
                If strREAL_MED_CODE <> "" Then Con4 = String.Format(" AND (REAL_MED_CODE = '{0}')", strREAL_MED_CODE) '��ü���ڵ�
                If strREAL_MED_NAME <> "" Then Con5 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) LIKE '%{0}%')", strREAL_MED_NAME) '��ü���
                If strTIMCODE <> "" Then Con6 = String.Format(" AND (TIMCODE = '{0}')", strTIMCODE) '���ڵ�
                If strTIMNAME <> "" Then Con7 = String.Format(" AND (DBO.SC_GET_CUSTNAME_FUN(TIMCODE) LIKE '%{0}%')", strTIMNAME) '����
                If strMEDCODE <> "" Then Con8 = String.Format(" AND (MEDCODE = '{0}')", strMEDCODE) '��ü�ڵ�
                If strMEDNAME <> "" Then Con9 = String.Format(" AND (DBO.SC_GET_CUSTNAME_FUN(MEDCODE) LIKE '%{0}%')", strMEDNAME) '��ü��
                If strSUBSEQ <> "" Then Con10 = String.Format(" AND (SUBSEQ = '{0}')", strSUBSEQ) '�귣���ڵ�
                If strSUBSEQNAME <> "" Then Con11 = String.Format(" AND (DBO.SC_GET_SUBSEQNAME_FUN(SUBSEQ) LIKE '%{0}%')", strSUBSEQNAME) '�귣���


                strWhere = BuildFields(" ", Con1, Con2, Con3, Con4, Con5, Con6, Con7, Con8, Con9, Con10, Con11)

                strFormet = strFormet & "  select "
                strFormet = strFormet & "  0 CHK, "
                strFormet = strFormet & "  YEARMON, SEQ, GFLAG, "
                strFormet = strFormet & "  CASE ISNULL(GFLAG,'') WHEN '' THEN '�̽���' WHEN '0' THEN '�̽���' ELSE CASE ISNULL(TRU_TRANS_NO,'') WHEN '' THEN '����' ELSE '���Ϸ�' END  END AS GFLAGNAME, "
                strFormet = strFormet & "  DEMANDDAY, "
                strFormet = strFormet & "  CLIENTCODE, "
                strFormet = strFormet & "  DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME,"
                strFormet = strFormet & "  MEDCODE, "
                strFormet = strFormet & "  DBO.SC_GET_CUSTNAME_FUN(MEDCODE) MEDNAME,"
                strFormet = strFormet & "  REAL_MED_CODE,"
                strFormet = strFormet & "  DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) REAL_MED_NAME, "
                strFormet = strFormet & "  SUBSEQ,"
                strFormet = strFormet & "  DBO.SC_GET_SUBSEQNAME_FUN(SUBSEQ) SUBSEQNAME, "
                strFormet = strFormet & "  TIMCODE,"
                strFormet = strFormet & "  DBO.SC_GET_CUSTNAME_FUN(TIMCODE) TIMNAME, "
                strFormet = strFormet & "  MATTERCODE,"
                strFormet = strFormet & "  DBO.MD_GET_MATTERNAME_FUN(MATTERCODE) MATTERNAME, "
                strFormet = strFormet & "  DEPT_CD,"
                strFormet = strFormet & "  DBO.SC_DEPT_NAME_FUN(DEPT_CD) DEPT_NAME,"
                strFormet = strFormet & "  EXCLIENTCODE, "
                strFormet = strFormet & "  dbo.SC_GET_EXCLIENTALLNAME_FUN(EXCLIENTCODE) EXCLIENTNAME, "
                strFormet = strFormet & "  GREATCODE, "
                strFormet = strFormet & "  dbo.SC_GET_GREATCUSTNAME_FUN(GREATCODE) GREATNAME,"
                strFormet = strFormet & "  MPP AS MPP_CODE, "
                strFormet = strFormet & "  dbo.SC_GET_HIGHCUSTNAME_FUN(MPP) MPP_NAME,"
                strFormet = strFormet & "  CLIENTSUBCODE, "
                strFormet = strFormet & "  DBO.SC_GET_CUSTNAME_FUN(CLIENTSUBCODE) CLIENSUBMNAME,  "
                strFormet = strFormet & "  PROGRAM, "
                strFormet = strFormet & "  CNT, "
                strFormet = strFormet & "  AMT, COMMI_RATE, COMMISSION, "
                strFormet = strFormet & "  TBRDSTDATE, TBRDEDDATE, "
                strFormet = strFormet & "  VOCH_TYPE, "
                strFormet = strFormet & "  TRU_TAX_FLAG,COMMI_TAX_FLAG, "
                strFormet = strFormet & "  MEMO AS BIGO, "
                strFormet = strFormet & "  CASE WHEN ISNULL(TRU_TRANS_NO,'') <> '' THEN "
                strFormet = strFormet & "  SUBSTRING(TRU_TRANS_NO, 0, DBO.INSTR(0, TRU_TRANS_NO,'-')) + '-' + "
                strFormet = strFormet & "  dbo.lpad(SUBSTRING(TRU_TRANS_NO, DBO.INSTR(0, TRU_TRANS_NO,'-')+1, DBO.INSTR(0, SUBSTRING(TRU_TRANS_NO,8,LEN(TRU_TRANS_NO)),'-')-1), 4, 0) + '-' + "
                strFormet = strFormet & "  dbo.lpad(SUBSTRING(TRU_TRANS_NO, DBO.INSTR(0, TRU_TRANS_NO,'-')+1 + DBO.INSTR(0, SUBSTRING(TRU_TRANS_NO,8,LEN(TRU_TRANS_NO)),'-'), LEN(TRU_TRANS_NO)),4,0) "
                strFormet = strFormet & "  ELSE "
                strFormet = strFormet & "  TRU_TRANS_NO "
                strFormet = strFormet & "  END AS TRU_TRANS_NO , "
                strFormet = strFormet & "  CASE WHEN ISNULL(COMMI_TRANS_NO,'') <> '' THEN "
                strFormet = strFormet & "  SUBSTRING(COMMI_TRANS_NO, 0, DBO.INSTR(0, COMMI_TRANS_NO,'-')) + '-' + "
                strFormet = strFormet & "  dbo.lpad(SUBSTRING(COMMI_TRANS_NO, DBO.INSTR(0, COMMI_TRANS_NO,'-')+1, DBO.INSTR(0, SUBSTRING(COMMI_TRANS_NO,8,LEN(COMMI_TRANS_NO)),'-')-1), 4, 0) + '-' + "
                strFormet = strFormet & "  dbo.lpad(SUBSTRING(COMMI_TRANS_NO, DBO.INSTR(0, COMMI_TRANS_NO,'-')+1 + DBO.INSTR(0, SUBSTRING(COMMI_TRANS_NO,8,LEN(COMMI_TRANS_NO)),'-'), LEN(COMMI_TRANS_NO)),4,0) "
                strFormet = strFormet & "  ELSE "
                strFormet = strFormet & "  COMMI_TRANS_NO "
                strFormet = strFormet & "  END AS COMMI_TRANS_NO "
                strFormet = strFormet & "  from MD_TOTAL_MEDIUM "
                strFormet = strFormet & "  where 1=1  "
                strFormet = strFormet & "  {0} "
                strFormet = strFormet & "  ORDER BY CASE WHEN SUBSTRING(DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE),1,3) = '(��)' THEN"
                strFormet = strFormet & "  SUBSTRING(DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE),4,100)"
                strFormet = strFormet & "  ELSE"
                strFormet = strFormet & "  dbo.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) END"

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

    '============== ������ä�� ����ó��
    Public Function ProcessRtn_ConfirmTotal_OK(ByVal strInfoXML As String, _
                                               ByVal vntData As Object, _
                                               ByVal strFLAG As String) As Integer '������ INSERT/UPDATE
        'strYEARMON, strSEQ, strSUSU, strAMT
        Dim intRtn As Integer
        Dim i, intColCnt, intRows As Integer

        Dim strGFLAG As String
        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()

                If strFLAG = "CONFIRM" Then
                    strGFLAG = "1"
                Else
                    strGFLAG = "0"
                End If

                If IsArray(vntData) Then
                    '''����� Entity ��ü����(Config ������ �Ѱܻ���)
                    'ceMD_BOOKING_MEDIUM
                    mobjceMD_TOTAL_MEDIUM = New ceMD_TOTAL_MEDIUM(mobjSCGLConfig)
                    intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)

                    For i = 1 To intRows
                        If GetElement(vntData, "CHK", intColCnt, i, OPTIONAL_STR) = "1" Then
                            intRtn = UpdateRtn_GFLAGTOTAL(vntData, intColCnt, i, strGFLAG)
                        End If
                    Next
                End If
                .mobjSCGLSql.SQLCommitTrans()
                Return intRows
            Catch err As Exception
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".ProcessRtn_ConfirmTotal_OK")
            Finally
                .mobjSCGLSql.SQLDisconnect()
                mobjceMD_TOTAL_MEDIUM.Dispose()
            End Try
        End With
    End Function
#End Region

#Region "���ͳ�û�೻�� ��ȸ/����"
    ' =============== ���ͳ� û�೻����ȸ
    Public Function SelectRtn_INTERNET(ByVal strInfoXML As String, _
                                       ByRef intRowCnt As Integer, _
                                       ByRef intColCnt As Integer, _
                                       ByVal strYEARMON As String, _
                                       ByVal strCAMPAIGN_CODE As String, _
                                       ByVal strCAMPAIGN_NAME As String, _
                                       ByVal strCLIENTCODE As String, _
                                       ByVal strCLIENTNAME As String, _
                                       ByVal strREAL_MED_CODE As String, _
                                       ByVal strREAL_MED_NAME As String, _
                                       ByVal strTIMCODE As String, _
                                       ByVal strTIMNAME As String, _
                                       ByVal strMEDCODE As String, _
                                       ByVal strMEDNAME As String, _
                                       ByVal strSUBSEQ As String, _
                                       ByVal strSUBSEQNAME As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim Con1, Con2, Con3, Con4, Con5, Con6, Con7, Con8, Con9, Con10, Con11, Con12, Con13 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '�⺻���� ����

                Con1 = "" : Con2 = "" : Con3 = "" : Con4 = "" : Con5 = ""
                Con6 = "" : Con7 = "" : Con8 = "" : Con9 = "" : Con10 = ""
                Con11 = "" : Con12 = "" : Con13 = ""

                If strYEARMON <> "" Then Con1 = String.Format(" AND (SUBSTRING(DEMANDDAY,1,6) = '{0}')", strYEARMON) '���
                If strCAMPAIGN_CODE <> "" Then Con2 = String.Format(" AND (CAMPAIGN_CODE = '{0}')", strCAMPAIGN_CODE) 'ķ�����ڵ�
                If strCAMPAIGN_NAME <> "" Then Con3 = String.Format(" AND (DBO.MD_GET_CAMPAIGNNAME_FUN(CAMPAIGN_CODE) LIKE '%{0}%')", strCAMPAIGN_NAME) 'ķ���θ�
                If strCLIENTCODE <> "" Then Con4 = String.Format(" AND (CLIENTCODE = '{0}')", strCLIENTCODE) '�������ڵ�
                If strCLIENTNAME <> "" Then Con5 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCLIENTNAME) '�����ָ�
                If strREAL_MED_CODE <> "" Then Con6 = String.Format(" AND (REAL_MED_LOWCODE = '{0}')", strREAL_MED_CODE) '��ü���ڵ�
                If strREAL_MED_NAME <> "" Then Con7 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_LOWCODE) LIKE '%{0}%')", strREAL_MED_NAME) '��ü���
                If strTIMCODE <> "" Then Con8 = String.Format(" AND (TIMCODE = '{0}')", strTIMCODE) '���ڵ�
                If strTIMNAME <> "" Then Con9 = String.Format(" AND (DBO.SC_GET_CUSTNAME_FUN(TIMCODE) LIKE '%{0}%')", strTIMNAME) '����
                If strMEDCODE <> "" Then Con10 = String.Format(" AND (MEDCODE = '{0}')", strMEDCODE) '��ü�ڵ�
                If strMEDNAME <> "" Then Con11 = String.Format(" AND (DBO.SC_GET_CUSTNAME_FUN(MEDCODE) LIKE '%{0}%')", strMEDNAME) '��ü��
                If strSUBSEQ <> "" Then Con12 = String.Format(" AND (SUBSEQ = '{0}')", strSUBSEQ) '�귣���ڵ�
                If strSUBSEQNAME <> "" Then Con13 = String.Format(" AND (DBO.SC_GET_SUBSEQNAME_FUN(SUBSEQ) LIKE '%{0}%')", strSUBSEQNAME) '�귣���


                strWhere = BuildFields(" ", Con1, Con2, Con3, Con4, Con5, Con6, Con7, Con8, Con9, Con10, Con11, Con12, Con13)

                strFormet = " SELECT "
                strFormet = strFormet & " 0 CHK, "
                strFormet = strFormet & " CASE ISNULL(GFLAG,'') WHEN '' THEN '�̽���' WHEN '0' THEN '�̽���' ELSE CASE ISNULL(TRU_TRANS_NO,'') WHEN '' THEN '����' ELSE '���Ϸ�' END  END AS GFLAGNAME, "
                strFormet = strFormet & " DBO.MD_TRANS_YN_FUN(YEARMON,SEQ, 'O') CONFIRMFLAG, "
                strFormet = strFormet & " YEARMON, SEQ, "
                strFormet = strFormet & " DBO.MD_GET_CAMPAIGNNAME_FUN(CAMPAIGN_CODE) CAMPAIGN_NAME,"
                strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME, "
                strFormet = strFormet & " DBO.SC_GET_CUSTNAME_FUN(TIMCODE) TIMNAME, "
                strFormet = strFormet & " DBO.SC_GET_CUSTNAME_FUN(MEDCODE) MEDNAME,"
                strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_LOWCODE) REAL_MED_LOWNAME,"
                strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) REAL_MED_NAME,"
                strFormet = strFormet & " TBRDSTDATE,"
                strFormet = strFormet & " TBRDEDDATE,"
                strFormet = strFormet & " MATTERNAME, "
                strFormet = strFormet & " AMT, "
                strFormet = strFormet & " COMMI_RATE, "
                strFormet = strFormet & " COMMISSION, "
                strFormet = strFormet & " MEMO, "
                strFormet = strFormet & " TRU_TRANS_NO,"
                strFormet = strFormet & " EXCLIENTCODE,"
                strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(EXCLIENTCODE) EXCLIENTNAME "
                strFormet = strFormet & " FROM MD_INTERNET_MEDIUM"
                strFormet = strFormet & " WHERE 1=1  "
                strFormet = strFormet & " {0} "
                strFormet = strFormet & " ORDER BY "
                strFormet = strFormet & " CASE WHEN substring(dbo.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE),1,3) = '(��)' THEN "
                strFormet = strFormet & " substring(dbo.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE),4,100) "
                strFormet = strFormet & " ELSE dbo.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) "
                strFormet = strFormet & " END"

                strSQL = String.Format(strFormet, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)

                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)

                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_INTERNET")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function


    '============== ���ι� ��� ó�� 1
    Public Function ProcessRtn_ConfirmINTERNET_OK(ByVal strInfoXML As String, _
                                                  ByVal vntData As Object, _
                                                  ByVal strFLAG As String) As Integer
        Dim intRtn As Integer '����� ����
        Dim i, intColCnt, intRows As Integer '����, �÷�Cnt, �ο�Cnt ����

        SetConfig(strInfoXML) '�⺻���� Setting

        With mobjSCGLConfig '�⺻������ ������ �ִ� Config ��ü
            Try
                'XML Element ���� ���� (strMasterXML�� ��ȯ)
                Dim xmlRoot As XmlElement
                'xmlRoot = XMLGetRoot(strMasterXML) 'XML ������

                'DB���� �� Ʈ����� ����
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()

                If IsArray(vntData) Then
                    intRtn = strDETAIL_DIVAMTINTERNET(strInfoXML, vntData, strFLAG)
                End If

                'Ʈ�����Commit
                .mobjSCGLSql.SQLCommitTrans()
                Return intRtn
            Catch err As Exception
                'Ʈ�����RollBack �� ���� ����
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".ProcessRtn_ConfirmINTERNET_OK")
            Finally
                'Resource����
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function


    '============== ���� �� ��� ó��2
    Public Function strDETAIL_DIVAMTINTERNET(ByVal strInfoXML As String, _
                                             ByVal vntData As Object, _
                                             ByVal strFLAG As String) As Integer '������ INSERT/UPDATE

        Dim intRtn, intRtn2 As Integer
        Dim i, intColCnt, intRows, intSEQ As Integer
        Dim dblID As Double '�ڵ� ID ������� ���� ���
        Dim strGFLAG


        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try
                If strFLAG = "CONFIRM" Then
                    strGFLAG = "1"
                Else
                    strGFLAG = "0"
                End If
                If IsArray(vntData) Then
                    '''����� Entity ��ü����(Config ������ �Ѱܻ���)
                    'ceMD_BOOKING_MEDIUM
                    mobjceMD_INTERNET_MEDIUM = New ceMD_INTERNET_MEDIUM(mobjSCGLConfig)
                    '''vntData�� �÷���, �ο���� �����Է�
                    intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)
                    '''�ش��ϴ�Row ��ŭ Loop

                    For i = 1 To intRows
                        If GetElement(vntData, "CHK", intColCnt, i, OPTIONAL_STR) = "1" And GetElement(vntData, "CONFIRMFLAG", intColCnt, i, OPTIONAL_STR) = "N" Then
                            intRtn = UpdateRtn_GFLAGINTERNET(vntData, intColCnt, i, strGFLAG)
                        End If
                    Next
                End If

                Return intRtn
            Catch err As Exception

                Throw RaiseSysErr(err, CLASS_NAME & ".strDETAIL_DIVAMTINTERNET")
            Finally
                mobjceMD_INTERNET_MEDIUM.Dispose()
            End Try
        End With
    End Function
#End Region

#End Region

#Region "GROUP BLOCK : �ܺο� ����� Method"

    '����ó�� Entity ó��
    Private Function UpdateRtn_GFLAGCATV(ByVal vntData As Object, _
                                         ByVal intColCnt As Integer, _
                                         ByVal intRow As Integer, _
                                         ByVal strGFLAG As String) As Integer
        Dim intRtn As Integer
        intRtn = mobjceMD_CATV_MEDIUM.GFLAGUpdate( _
                                       strGFLAG, _
                                       GetElement(vntData, "YEARMON", intColCnt, intRow), _
                                       GetElement(vntData, "SEQ", intColCnt, intRow, NULL_NUM, True))
    End Function
    Private Function UpdateRtn_GFLAGTOTAL(ByVal vntData As Object, _
                                          ByVal intColCnt As Integer, _
                                          ByVal intRow As Integer, _
                                          ByVal strGFLAG As String) As Integer
        Dim intRtn As Integer
        intRtn = mobjceMD_TOTAL_MEDIUM.GFLAGUpdate( _
                                       strGFLAG, _
                                       GetElement(vntData, "YEARMON", intColCnt, intRow), _
                                       GetElement(vntData, "SEQ", intColCnt, intRow, NULL_NUM, True))
    End Function

    Private Function UpdateRtn_GFLAGINTERNET(ByVal vntData As Object, _
                                             ByVal intColCnt As Integer, _
                                             ByVal intRow As Integer, _
                                             ByVal strGFLAG As String) As Integer
        Dim intRtn As Integer
        intRtn = mobjceMD_INTERNET_MEDIUM.GFLAGUpdate( _
                                                strGFLAG, _
                                                GetElement(vntData, "YEARMON", intColCnt, intRow), _
                                                GetElement(vntData, "SEQ", intColCnt, intRow, NULL_NUM, True))
    End Function

    Private Function UpdateRtn_GFLAGBOOKING(ByVal vntData As Object, _
                                            ByVal intColCnt As Integer, _
                                            ByVal intRow As Integer, _
                                            ByVal strGFLAG As String) As Integer
        Dim intRtn As Integer
        intRtn = mobjceMD_BOOKING_MEDIUM.GFLAGUpdate( _
                                                strGFLAG, _
                                                GetElement(vntData, "YEARMON", intColCnt, intRow), _
                                                GetElement(vntData, "SEQ", intColCnt, intRow, NULL_NUM, True))
    End Function

#End Region
End Class