'****************************************************************************************
'시스템구분 : SFAR/표준샘플/Server단 Control Component
'실행  환경 : COM+ Service Server Package
'프로그램명 : ccSCCDGet.vb (공통코드 조회 Control Class)
'기      능 : 공통코드 조회를 위한 클래스
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008/08/14 By Kim Tae Ho
'           :2) 
'****************************************************************************************
Imports SCGLControl                 'Control Class의 Base Class
Imports SCGLUtil.cbSCGLConfig       'Configuration 클래스
Imports SCGLUtil.cbSCGLErr          '오류처리 클래스
Imports SCGLUtil.cbSCGLXml          'XML처리 클래스
Imports SCGLUtil.cbSCGLUtil         '기타 유틸리티 클래스
Imports eSCCO                       ' 엔터티 추가

Public Class ccSCCOGET
    Inherits ccControl

#Region "GROUP BLOCk : 전역 또는 모듈레벨의 변수/상수 선언"
    Private Const CLASS_NAME = "ccSCCOGET"    '자신의 클래스명
#End Region

#Region "GROUP BLOCk : Property 선언"
    Private mobjceSC_EMPLOYEE_MST As eSCCO.ceSC_EMPLOYEE_MST        '사용할 Entity 변수 선언
#End Region

#Region "GROUP BLOCk : Event 선언"
#End Region

#Region "GROUP BLOCk : 외부에 공개 Method"

#Region "1. 청구지조회"
    ' =============== SelectRtnSample Code
    Public Function GetHIGHCUSTCODE(ByVal strInfoXML As String, _
                                    ByRef intRowCnt As Integer, _
                                    ByRef intColCnt As Integer, _
                                    ByRef strCUSTCODE As String, _
                                    ByRef strCUSTNAME As String, _
                                    ByRef strMEDFLAG As String) As Object

        Dim strCols As String        '컬럼변수
        Dim strWhere As String       'Where조건 변수
        Dim strFormat As String      'SQL Format 변수
        Dim strSQL As String         'SQL 변수
        Dim vntData As Object        'Array Return 변수(Array 를사용할 때 선언)
        Dim strSelFields
        Dim intCnt
        Dim intRtn
        Dim Con1, Con2, Con3
        SetConfig(strInfoXML)

        Con1 = "" : Con2 = "" : Con3 = ""

        If strCUSTCODE <> "" Then Con1 = String.Format(" AND (HIGHCUSTCODE = '{0}')", strCUSTCODE)
        If strCUSTNAME <> "" Then Con2 = String.Format(" AND (CUSTNAME LIKE '%{0}%')", strCUSTNAME)
        If strMEDFLAG <> "" Then Con3 = String.Format(" AND (MEDFLAG = '{0}')", strMEDFLAG)



        strWhere = BuildFields(" ", Con1, Con2, Con3)

        strFormat = "  SELECT "
        strFormat = strFormat & "  HIGHCUSTCODE,"
        strFormat = strFormat & "  CUSTNAME,"
        strFormat = strFormat & "  BUSINO,"
        strFormat = strFormat & "  COMPANYNAME,DBO.SC_GET_CUSTGROUPGBN_FUN(HIGHCUSTCODE) GROUPGBN "
        strFormat = strFormat & "  FROM SC_CUST_HDR"
        strFormat = strFormat & "  WHERE 1=1 AND USE_FLAG = '1' {0} ORDER BY CUSTNAME"



        strSQL = String.Format(strFormat, strWhere)

        '기본정보 Setting
        With mobjSCGLConfig '기본정보 Config 개체
            Try
                ' DB 접속
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                ' ------ Array 데이터 조회 (True 일때 헤더정보 포함 조회(Sheet Data Binding 할 경우 사용), False 일때 데이터만 조회)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)

                Return vntData
            Catch err As Exception
                ' 오류 전송
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn")
            Finally
                ' DB 접속 종료
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "2. 사업부조회"
    ' =============== SelectRtnSample Code
    Public Function GetCLIENTSUBCODE(ByVal strInfoXML As String, _
                                     ByRef intRowCnt As Integer, _
                                     ByRef intColCnt As Integer, _
                                     ByVal strCLIENTCODE As String, _
                                     ByVal strCLIENTNAME As String, _
                                     ByVal strCLIENTSUBCODE As String, _
                                     ByVal strCLIENTSUBNAME As String) As Object

        Dim strCols As String        '컬럼변수
        Dim strWhere As String       'Where조건 변수
        Dim strFormat As String      'SQL Format 변수
        Dim strSQL As String         'SQL 변수
        Dim vntData As Object        'Array Return 변수(Array 를사용할 때 선언)
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
        strFormat = strFormat & "  CUSTCODE,"
        strFormat = strFormat & "  CUSTNAME,"
        strFormat = strFormat & "  dbo.SC_GET_BUSINO_FUN(HIGHCUSTCODE) BUSINO,"
        strFormat = strFormat & "  HIGHCUSTCODE,"
        strFormat = strFormat & "  DBO.SC_GET_HIGHCUSTNAME_FUN(HIGHCUSTCODE) COMPANYNAME "
        strFormat = strFormat & "  FROM SC_CUST_DTL"
        strFormat = strFormat & "  WHERE 1=1 AND MEDFLAG = 'A' AND GBNFLAG = '1' AND USE_FLAG = '1' {0} ORDER BY CUSTNAME"

        strSQL = String.Format(strFormat, strWhere)

        '기본정보 Setting
        With mobjSCGLConfig '기본정보 Config 개체
            Try
                ' DB 접속
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                ' ------ Array 데이터 조회 (True 일때 헤더정보 포함 조회(Sheet Data Binding 할 경우 사용), False 일때 데이터만 조회)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)

                Return vntData
            Catch err As Exception
                ' 오류 전송
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn")
            Finally
                ' DB 접속 종료
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "3. 대표브랜드조회"
    ' =============== SelectRtnSample Code
    Public Function GetHIGHSUBSEQ(ByVal strInfoXML As String, _
                                  ByRef intRowCnt As Integer, _
                                  ByRef intColCnt As Integer, _
                                  ByVal strHIGHSEQNO As String, _
                                  ByVal strHIGHSEQNAME As String) As Object

        Dim strCols As String        '컬럼변수
        Dim strWhere As String       'Where조건 변수
        Dim strFormat As String      'SQL Format 변수
        Dim strSQL As String         'SQL 변수
        Dim vntData As Object        'Array Return 변수(Array 를사용할 때 선언)
        Dim strSelFields
        Dim intCnt
        Dim intRtn
        Dim Con1, Con2

        SetConfig(strInfoXML)

        Con1 = "" : Con2 = ""

        If strHIGHSEQNO <> "" Then Con1 = String.Format(" AND (HIGHSEQNO = '{0}')", strHIGHSEQNO)
        If strHIGHSEQNAME <> "" Then Con2 = String.Format(" AND (HIGHSEQNAME LIKE '%{0}%')", strHIGHSEQNAME)


        strWhere = BuildFields(" ", Con1, Con2)

        strFormat = "  SELECT "
        strFormat = strFormat & "  HIGHSEQNO,"
        strFormat = strFormat & "  HIGHSEQNAME"
        strFormat = strFormat & "  FROM SC_SUBSEQ_HDR"
        strFormat = strFormat & "  WHERE 1=1  {0} ORDER BY HIGHSEQNAME"

        strSQL = String.Format(strFormat, strWhere)

        '기본정보 Setting
        With mobjSCGLConfig '기본정보 Config 개체
            Try
                ' DB 접속
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                ' ------ Array 데이터 조회 (True 일때 헤더정보 포함 조회(Sheet Data Binding 할 경우 사용), False 일때 데이터만 조회)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)

                Return vntData
            Catch err As Exception
                ' 오류 전송
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn")
            Finally
                ' DB 접속 종료
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "4. CC_CODE: CC 코드조회(Only CC)"
    Public Function GetCC(ByVal strInfoXML As String, _
                          ByRef intRowCnt As Integer, _
                          ByRef intColCnt As Integer, _
                          ByVal strCODE_NAME As String) As Object

        Dim strSQL, strFormat, strSelFields, strKeys As String
        Dim strCondition As String
        Dim strChkDate As String = ""
        Dim vntData As Object

        SetConfig(strInfoXML)   '기본정보 설정
        With mobjSCGLConfig

            '한글인 경우
            strCondition = String.Format("AND CC_NAME LIKE '%{0}%'", strCODE_NAME)

            '조회 필드 설정
            strSelFields = "CC_CODE,CC_NAME"

            strFormat = "select"
            strFormat = strFormat & " {0}"
            strFormat = strFormat & " FROM SC_CC A WHERE 1=1"
            strFormat = strFormat & " AND PC='Y' AND USE_YN = 'Y'  {1}"
            strSQL = String.Format(strFormat, _
                                       strSelFields, strCondition)

            '데이터 조회
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetCC")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "5. 팀조회"
    ' =============== SelectRtnSample Code
    Public Function GetTIMCODE(ByVal strInfoXML As String, _
                                     ByRef intRowCnt As Integer, _
                                     ByRef intColCnt As Integer, _
                                     ByVal strCLIENTCODE As String, _
                                     ByVal strCLIENTNAME As String, _
                                     ByVal strCLIENTSUBCODE As String, _
                                     ByVal strCLIENTSUBNAME As String) As Object

        Dim strCols As String        '컬럼변수
        Dim strWhere As String       'Where조건 변수
        Dim strFormat As String      'SQL Format 변수
        Dim strSQL As String         'SQL 변수
        Dim vntData As Object        'Array Return 변수(Array 를사용할 때 선언)
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
        strFormat = strFormat & "  DBO.SC_GET_HIGHCUSTNAME_FUN(HIGHCUSTCODE) COMPANYNAME, DBO.SC_GET_CUSTGROUPGBN_FUN(HIGHCUSTCODE) GROUPGBN "
        strFormat = strFormat & "  FROM SC_CUST_DTL"
        strFormat = strFormat & "  WHERE 1=1 AND MEDFLAG = 'A' AND GBNFLAG = '0' AND USE_FLAG = '1' {0} ORDER BY CUSTNAME"

        strSQL = String.Format(strFormat, strWhere)

        '기본정보 Setting
        With mobjSCGLConfig '기본정보 Config 개체
            Try
                ' DB 접속
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                ' ------ Array 데이터 조회 (True 일때 헤더정보 포함 조회(Sheet Data Binding 할 경우 사용), False 일때 데이터만 조회)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)

                Return vntData
            Catch err As Exception
                ' 오류 전송
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn")
            Finally
                ' DB 접속 종료
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "6. 브랜드별 (광고주,사업부,팁,담당부서) 조회"
    ' =============== Get_BrandInfo
    Public Function Get_BrandInfo(ByVal strInfoXML As String, _
                                  ByRef intRowCnt As Integer, _
                                  ByRef intColCnt As Integer, _
                                  ByVal strSUBSEQ As String, _
                                  ByVal strSUBSEQNAME As String, _
                                  ByVal strCUSTCODE As String, _
                                  ByVal strCUSTNAME As String) As Object

        Dim strCols As String        '컬럼변수
        Dim strWhere As String       'Where조건 변수
        Dim strFormat As String      'SQL Format 변수
        Dim strSQL As String         'SQL 변수
        Dim vntData As Object        'Array Return 변수(Array 를사용할 때 선언)
        Dim strSelFields
        Dim intCnt
        Dim intRtn
        Dim Con1, Con2, Con3, Con4
        SetConfig(strInfoXML)

        Con1 = "" : Con2 = "" : Con3 = "" : Con4 = ""

        If strSUBSEQ <> "" Then Con1 = String.Format(" AND (SEQNO = '{0}')", strSUBSEQ)
        If strSUBSEQNAME <> "" Then Con2 = String.Format(" AND (SEQNAME LIKE '%{0}%')", strSUBSEQNAME)
        If strCUSTCODE <> "" Then Con3 = String.Format(" AND (CUSTCODE = '{0}')", strCUSTCODE)
        If strCUSTNAME <> "" Then Con4 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(CUSTCODE) LIKE '%{0}%')", strCUSTNAME)


        strWhere = BuildFields(" ", Con1, Con2, Con3, Con4)

        strFormat = "  SELECT "
        strFormat = strFormat & "  SEQNO, "
        strFormat = strFormat & "  SEQNAME, "
        strFormat = strFormat & "  CUSTCODE, "
        strFormat = strFormat & "  DBO.SC_GET_HIGHCUSTNAME_FUN(CUSTCODE) CUSTNAME,"
        strFormat = strFormat & "  TIMCODE,"
        strFormat = strFormat & "  DBO.SC_GET_CUSTNAME_FUN(TIMCODE) TIMNAME,"
        strFormat = strFormat & "  CLIENTSUBCODE,"
        strFormat = strFormat & "  DBO.SC_GET_CUSTNAME_FUN(CLIENTSUBCODE) CLIENTSUBNAME,"
        strFormat = strFormat & "  DEPT_CD,"
        strFormat = strFormat & "  DBO.SC_DEPT_NAME_FUN(DEPT_CD) DEPT_NAME, DBO.SC_GET_CUSTGROUPGBN_FUN(CUSTCODE) GROUPGBN"
        strFormat = strFormat & "  FROM SC_SUBSEQ_DTL"
        strFormat = strFormat & "  WHERE 1=1 AND ISNULL(ATTR01,'N') = 'Y' {0} ORDER BY SEQNAME"

        strSQL = String.Format(strFormat, strWhere)

        '기본정보 Setting
        With mobjSCGLConfig '기본정보 Config 개체
            Try
                ' DB 접속
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                ' ------ Array 데이터 조회 (True 일때 헤더정보 포함 조회(Sheet Data Binding 할 경우 사용), False 일때 데이터만 조회)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)

                Return vntData
            Catch err As Exception
                ' 오류 전송
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn")
            Finally
                ' DB 접속 종료
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "7. 광고처조회"
    ' =============== SelectRtnSample Code
    Public Function GetGREATCUSTCODE(ByVal strInfoXML As String, _
                                     ByRef intRowCnt As Integer, _
                                     ByRef intColCnt As Integer, _
                                     ByRef strCUSTCODE As String, _
                                     ByRef strCUSTNAME As String) As Object

        Dim strCols As String        '컬럼변수
        Dim strWhere As String       'Where조건 변수
        Dim strFormat As String      'SQL Format 변수
        Dim strSQL As String         'SQL 변수
        Dim vntData As Object        'Array Return 변수(Array 를사용할 때 선언)
        Dim strSelFields
        Dim intCnt
        Dim intRtn
        Dim Con1, Con2
        SetConfig(strInfoXML)

        Con1 = "" : Con2 = ""

        If strCUSTCODE <> "" Then Con1 = String.Format(" AND (HIGHCUSTCODE = '{0}')", strCUSTCODE)
        If strCUSTNAME <> "" Then Con2 = String.Format(" AND (GREATNAME LIKE '%{0}%')", strCUSTNAME)

        strWhere = BuildFields(" ", Con1, Con2)

        strFormat = "  SELECT "
        strFormat = strFormat & "  HIGHCUSTCODE,"
        strFormat = strFormat & "  CUSTNAME,"
        strFormat = strFormat & "  GREATNAME "
        strFormat = strFormat & "  FROM SC_CUST_HDR"
        strFormat = strFormat & "  WHERE 1=1 AND USE_FLAG = '1' AND MEDFLAG = 'A' {0} ORDER BY CUSTNAME"



        strSQL = String.Format(strFormat, strWhere)

        '기본정보 Setting
        With mobjSCGLConfig '기본정보 Config 개체
            Try
                ' DB 접속
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                ' ------ Array 데이터 조회 (True 일때 헤더정보 포함 조회(Sheet Data Binding 할 경우 사용), False 일때 데이터만 조회)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)

                Return vntData
            Catch err As Exception
                ' 오류 전송
                Throw RaiseSysErr(err, CLASS_NAME & ".GetGREATCUSTCODE")
            Finally
                ' DB 접속 종료
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "8. EMPNO 조회"
    Public Function GetSCEMP(ByVal strInfoXML As String, _
                             ByRef intRowCnt As Integer, _
                             ByRef intColCnt As Integer, _
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

        SetConfig(strInfoXML)   '기본정보 설정
        With mobjSCGLConfig
            If Len(strCODE) = 5 Then
                strCODE = "000" & strCODE
            End If

            '한글인 경우
            If strCODE <> "" Then Con1 = String.Format(" AND (EMPNO = '{0}')", strCODE)
            If strNAME <> "" Then Con2 = String.Format(" AND EMP_NAME LIKE '%{0}%'", strNAME)
            If strGUBUN <> "A" Then Con3 = String.Format(" AND SC_EMP_STATUS = '{0}'", strGUBUN)
            If strDEPTCD <> "" Then Con4 = String.Format(" AND (CC_CODE = '{0}')", strDEPTCD)
            If strDEPTNAME <> "" Then Con5 = String.Format(" AND DBO.SC_DEPT_NAME_FUN(CC_CODE) LIKE '%{0}%'", strDEPTNAME)

            '조회 필드 설정

            strSelFields = "EMPNO,EMP_NAME,CC_CODE,DBO.SC_DEPT_NAME_FUN(CC_CODE) CC_NAME,CASE SC_EMP_STATUS WHEN '0' THEN '재직' WHEN '1' THEN '휴직' WHEN '3' THEN '퇴직' END SC_EMP_STATUS,CASE ISNULL(E_MAIL,'') WHEN 'NULL' THEN '' ELSE ISNULL(E_MAIL,'') END E_MAIL,TEL,CELLPHONE,PASSWORD"
            strFormat = "SELECT {0} FROM SC_EMPLOYEE_MST A " & _
                                     "WHERE USE_YN = 'Y'  {1} {2} {3} {4} {5} " & _
                                     "AND USE_YN = 'Y' ORDER BY CC_CODE"
            strSQL = String.Format(strFormat, _
                                   strSelFields, Con1, Con2, Con3, Con4, Con5)

            '데이터 조회
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetSCEMP")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "9. 요청받는자 정보 조회"
    Public Function Get_SENDINFO(ByVal strInfoXML As String, _
                                 ByRef intRowCnt As Integer, _
                                 ByRef intColCnt As Integer, _
                                 ByVal strCODE As String, _
                                 ByVal strNAME As String) As Object

        Dim strSQL, strFormat, strSelFields, strKeys As String

        Dim vntData As Object
        Dim Con1, Con2 As String
        Dim strWhere As String

        Con1 = ""
        Con2 = ""


        SetConfig(strInfoXML)   '기본정보 설정
        With mobjSCGLConfig

            '한글인 경우
            If strCODE <> "" Then Con1 = String.Format(" AND (EMPNO = '{0}')", strCODE)
            If strNAME <> "" Then Con2 = String.Format(" AND EMP_NAME LIKE '%{0}%'", strNAME)


            strWhere = BuildFields(" ", Con1, Con2)

            strFormat = "  SELECT "
            strFormat = strFormat & "  EMP_NAME ,"
            strFormat = strFormat & "  E_MAIL ,"
            strFormat = strFormat & "  REPLACE(CELLPHONE,'-','') CELLPHONE"
            strFormat = strFormat & "  FROM SC_EMPLOYEE_MST"
            strFormat = strFormat & "  WHERE 1=1 AND USE_YN = 'Y' {0} "

            strFormat = strFormat & " UNION ALL"

            strFormat = strFormat & " SELECT "
            strFormat = strFormat & " EMP_NAME ,"
            strFormat = strFormat & " E_MAIL ,"
            strFormat = strFormat & " REPLACE(CELLPHONE,'-','') CELLPHONE "
            strFormat = strFormat & " FROM SC_EMPLOYEE_MST"
            strFormat = strFormat & " WHERE 1=1 AND USE_YN = 'Y' AND EMPNO = '{1}' "


            strSQL = String.Format(strFormat, strWhere, .WRKUSR)

            '데이터 조회
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetSCEMP")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    Public Function Get_SENDINFO2(ByVal strInfoXML As String, _
                                 ByRef intRowCnt As Integer, _
                                 ByRef intColCnt As Integer, _
                                 ByVal strTOEMPNO As String) As Object

        Dim strSQL, strFormat, strSelFields, strKeys As String

        Dim vntData As Object
        Dim Con1, Con2 As String
        Dim strWhere As String

        Con1 = ""
        Con2 = ""


        SetConfig(strInfoXML)   '기본정보 설정
        With mobjSCGLConfig

            '한글인 경우
            If strTOEMPNO <> "" Then Con1 = String.Format(" AND (EMPNO = '{0}')", strTOEMPNO)


            strWhere = BuildFields(" ", Con1)

            strFormat = "  SELECT "
            strFormat = strFormat & "  EMP_NAME ,"
            strFormat = strFormat & "  E_MAIL ,"
            strFormat = strFormat & "  REPLACE(CELLPHONE,'-','') CELLPHONE"
            strFormat = strFormat & "  FROM SC_EMPLOYEE_MST"
            strFormat = strFormat & "  WHERE 1=1 AND USE_YN = 'Y' {0} "

            strFormat = strFormat & " UNION ALL"

            strFormat = strFormat & " SELECT "
            strFormat = strFormat & " EMP_NAME ,"
            strFormat = strFormat & " E_MAIL ,"
            strFormat = strFormat & " REPLACE(CELLPHONE,'-','') CELLPHONE "
            strFormat = strFormat & " FROM SC_EMPLOYEE_MST"
            strFormat = strFormat & " WHERE 1=1 AND USE_YN = 'Y' AND EMPNO = '{1}' "


            strSQL = String.Format(strFormat, strWhere, .WRKUSR)

            '데이터 조회
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetSCEMP")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "10. 제작대행사 (대행사, 크리조직, 사원) 조회"
    ' =============== Get_BrandInfo
    Public Function Get_EXCLIENT_ALL(ByVal strInfoXML As String, _
                                     ByRef intRowCnt As Integer, _
                                     ByRef intColCnt As Integer, _
                                     ByVal strEXCLIENTCODE As String, _
                                     ByVal strEXCLIENTNAME As String, _
                                     ByVal strFLAG As String) As Object

        Dim strCols As String        '컬럼변수
        Dim strWhere As String       'Where조건 변수
        Dim strFormat As String      'SQL Format 변수
        Dim strSQL As String         'SQL 변수
        Dim vntData As Object        'Array Return 변수(Array 를사용할 때 선언)
        Dim strSelFields
        Dim intCnt
        Dim intRtn
        Dim Con1, Con2, Con3
        SetConfig(strInfoXML)

        Con1 = "" : Con2 = "" : Con3 = ""

        If strEXCLIENTCODE <> "" Then Con1 = String.Format(" AND (EXCLIENTCODE = '{0}')", strEXCLIENTCODE)
        If strEXCLIENTNAME <> "" Then Con2 = String.Format(" AND (EXCLIENTNAME LIKE '%{0}%')", strEXCLIENTNAME)
        If strFLAG <> "" Then Con3 = String.Format(" AND (MEDFLAG = '{0}')", strFLAG)


        strWhere = BuildFields(" ", Con1, Con2, Con3)



        strFormat = "  SELECT "
        strFormat = strFormat & "  MEDFLAGNAME, EXCLIENTCODE, EXCLIENTNAME"
        strFormat = strFormat & "  FROM ("
        strFormat = strFormat & "  	SELECT "
        strFormat = strFormat & "  	MEDFLAG,  "
        strFormat = strFormat & "  	CASE MEDFLAG "
        strFormat = strFormat & "  	WHEN 'G' THEN '대행사'"
        strFormat = strFormat & "  	WHEN 'K' THEN '크리조직'"
        strFormat = strFormat & "  	END AS	MEDFLAGNAME,"
        strFormat = strFormat & "  	HIGHCUSTCODE EXCLIENTCODE,"
        strFormat = strFormat & "  	CUSTNAME EXCLIENTNAME"
        strFormat = strFormat & "  	FROM SC_CUST_HDR"
        strFormat = strFormat & "  	WHERE MEDFLAG IN('G','K') AND USE_FLAG = '1'"
        strFormat = strFormat & "  	UNION ALL"
        strFormat = strFormat & "  	SELECT"
        strFormat = strFormat & "  	'C' MEDFLAG,  "
        strFormat = strFormat & "  	'CD' MEDFLAGNAME,"
        strFormat = strFormat & "  	EMPNO EXCLIENTCODE,"
        strFormat = strFormat & "  	EMP_NAME EXCLIENTNAME"
        strFormat = strFormat & "  	FROM SC_EMPLOYEE_MST"
        strFormat = strFormat & "  	WHERE USE_YN = 'Y'"
        strFormat = strFormat & "  )AAA"
        strFormat = strFormat & "  WHERE 1=1 {0} ORDER BY 1,2"

        strSQL = String.Format(strFormat, strWhere)

        '기본정보 Setting
        With mobjSCGLConfig '기본정보 Config 개체
            Try
                ' DB 접속
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                ' ------ Array 데이터 조회 (True 일때 헤더정보 포함 조회(Sheet Data Binding 할 경우 사용), False 일때 데이터만 조회)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)

                Return vntData
            Catch err As Exception
                ' 오류 전송
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn")
            Finally
                ' DB 접속 종료
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "11. SMS 조회"
    Public Function GetSMS(ByVal strInfoXML As String, _
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

        SetConfig(strInfoXML)   '기본정보 설정
        With mobjSCGLConfig
            If Len(strCODE) = 5 Then
                strCODE = "000" & strCODE
            End If

            '한글인 경우
            If strCODE <> "" Then Con1 = String.Format(" AND (EMPNO = '{0}')", strCODE)
            If strNAME <> "" Then Con2 = String.Format(" AND EMP_NAME LIKE '%{0}%'", strNAME)
            If strGUBUN <> "A" Then Con3 = String.Format(" AND SC_EMP_STATUS = '{0}'", strGUBUN)
            If strDEPTCD <> "" Then Con4 = String.Format(" AND (CC_CODE = '{0}')", strDEPTCD)
            If strDEPTNAME <> "" Then Con5 = String.Format(" AND DBO.SC_DEPT_NAME_FUN(CC_CODE) LIKE '%{0}%'", strDEPTNAME)

            '조회 필드 설정

            strSelFields = "EMPNO,EMP_NAME,CC_CODE,DBO.SC_DEPT_NAME_FUN(CC_CODE) CC_NAME,CASE SC_EMP_STATUS WHEN '0' THEN '재직' WHEN '1' THEN '휴직' WHEN '3' THEN '퇴직' END SC_EMP_STATUS,CASE ISNULL(E_MAIL,'') WHEN 'NULL' THEN '' ELSE ISNULL(E_MAIL,'') END E_MAIL,TEL,REPLACE(CELLPHONE,'-','') CELLPHONE,PASSWORD"
            strFormat = "SELECT {0} FROM SC_EMPLOYEE_MST A " & _
                                     "WHERE USE_YN = 'Y'  {1} {2} {3} {4} {5} " & _
                                     "ORDER BY CC_CODE"
            strSQL = String.Format(strFormat, _
                                   strSelFields, Con1, Con2, Con3, Con4, Con5)



            '데이터 조회
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetSMS")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "12. SMS 발송리스트 저장"
    Public Function ProcessRtn_SMS(ByVal strInfoXML As String, _
                                   ByVal vntData As Object, _
                                   ByVal mstrFromUserName As String, _
                                   ByVal mstrFromUserPhone As String, _
                                   ByVal strMSG As String, _
                                   ByRef dblID As Double) As Integer '데이터 INSERT/UPDATE
        Dim intRtn As Integer
        Dim i, intColCnt, intRows As Integer
        Dim strCREDAY As String
        Dim strMEMO As String
        Dim strSQL As String
        Dim strUSER As String
        Dim strMANAGER As String
        Dim strTOUSERPHONE As String

        SetConfig(strInfoXML)
        strUSER = mobjSCGLConfig.WRKUSR
        With mobjSCGLConfig
            'mstrFromUserName,mstrFromUserPhone,.txtMSG.value,dblID
            Try

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                If IsArray(vntData) Then
                    '''사용할 Entity 개체생성(Config 정보를 넘겨생성)
                    mobjceSC_EMPLOYEE_MST = New ceSC_EMPLOYEE_MST(mobjSCGLConfig)
                    '''vntData의 컬럼수, 로우수를 변수입력
                    intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)
                    'SMS 번호 생성
                    dblID = SelectRtn_SMSID(strUSER)

                    For i = 1 To intRows
                        strMANAGER = GetElement(vntData, "EMPNO", intColCnt, i)
                        strTOUSERPHONE = GetElement(vntData, "CELLPHONE", intColCnt, i)
                        strSQL = "INSERT INTO SC_SMS(USERNO,SEQ,MANAGER,SENDFLAG,FROMUSERNAME,FROMUSERPHONE,TOUSERPHONE,MSTMSG,SENDLOG) SELECT"
                        strSQL = strSQL & " '" & strUSER & "' USERNO," & dblID & " SEQ,'" & strMANAGER & "' MANAGER,'N' SENDFLAG,'" & mstrFromUserName & "' FROMUSERNAME,'" & mstrFromUserPhone & "' FROMUSERPHONE,'" & strTOUSERPHONE & "' TOUSERPHONE,'" & strMSG & "' MSTMSG,'' SENDLOG "

                        intRtn = mobjceSC_EMPLOYEE_MST.Clipping_Update(strSQL)
                    Next
                End If
                .mobjSCGLSql.SQLCommitTrans()
                Return dblID
            Catch err As Exception
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".ProcessRtn_SMS")
            Finally
                .mobjSCGLSql.SQLDisconnect()
                mobjceSC_EMPLOYEE_MST.Dispose()
            End Try
        End With
    End Function

    '============== SMS 번호생성
    Public Function SelectRtn_SMSID(ByVal strUSER As String) As String
        '여기부터 단순조회
        Dim strSQL, strFormat, strRtn As String
        'SetConfig(strInfoXML) '기본정보 Setting
        Dim strPRECODE As String
        With mobjSCGLConfig '기본정보 Config 개체

            Try

                strSQL = "SELECT ISNULL(MAX(SEQ),0) +1 FROM SC_SMS  "
                strRtn = .mobjSCGLSql.SQLSelectOneScalar(strSQL)
                Return strRtn
            Catch err As Exception
                ' 오류 전송
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_SMSID")
            Finally
            End Try
        End With
        '여기까지 단순조회
    End Function
#End Region

#Region "13. 계약자 조회(모든 청구지를 가져온다. 중복 제거)"
    Public Function GetCONTRACT_EXE(ByVal strInfoXML As String, _
                                    ByRef intRowCnt As Integer, _
                                    ByRef intColCnt As Integer, _
                                    ByRef strCUSTCODE As String, _
                                    ByRef strCUSTNAME As String) As Object

        Dim strCols As String        '컬럼변수
        Dim strWhere As String       'Where조건 변수
        Dim strFormet As String      'SQL Format 변수
        Dim strSQL As String         'SQL 변수
        Dim vntData As Object        'Array Return 변수(Array 를사용할 때 선언)
        Dim strSelFields
        Dim intCnt
        Dim intRtn
        Dim Con1, Con2, Con3
        SetConfig(strInfoXML)

        Con1 = "" : Con2 = ""

        If strCUSTCODE <> "" Then Con1 = String.Format(" AND (BUSINO LIKE '%{0}%')", strCUSTCODE)
        If strCUSTNAME <> "" Then Con2 = String.Format(" AND (CUSTNAME LIKE '%{0}%')", strCUSTNAME)

        strWhere = BuildFields(" ", Con1, Con2)

        strFormet = strFormet & "  SELECT "
        strFormet = strFormet & "  BUSINO,"
        strFormet = strFormet & "  CUSTNAME"
        strFormet = strFormet & "  FROM SC_CUST_HDR"
        strFormet = strFormet & "  WHERE 1=1 {0}"
        strFormet = strFormet & "  AND USE_FLAG = '1' "
        strFormet = strFormet & "  GROUP BY BUSINO, CUSTNAME"

        strSQL = String.Format(strFormet, strWhere)

        '기본정보 Setting
        With mobjSCGLConfig '기본정보 Config 개체
            Try
                ' DB 접속
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                ' ------ Array 데이터 조회 (True 일때 헤더정보 포함 조회(Sheet Data Binding 할 경우 사용), False 일때 데이터만 조회)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)

                Return vntData
            Catch err As Exception
                ' 오류 전송
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn")
            Finally
                ' DB 접속 종료
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region


#Region "14. CC_HIGHCODE: CC 상위부서코드 조회(Only CC)"
    Public Function GetCC_HIGH(ByVal strInfoXML As String, _
                               ByRef intRowCnt As Integer, _
                               ByRef intColCnt As Integer, _
                               ByVal strCODE_NAME As String) As Object

        Dim strSQL, strFormat As String
        Dim strCondition As String
        Dim vntData As Object

        SetConfig(strInfoXML)   '기본정보 설정
        With mobjSCGLConfig

            '한글인 경우
            strCondition = String.Format("AND DBO.SC_DEPT_NAME_FUN(HIGHDEPT_CD) LIKE '%{0}%'", strCODE_NAME)

            '조회 필드 설정
            strFormat = " SELECT "
            strFormat = strFormat & " HIGHDEPT_CD  CC_CODE,  "
            strFormat = strFormat & " DBO.SC_DEPT_NAME_FUN(HIGHDEPT_CD) CC_NAME"
            strFormat = strFormat & " FROM SC_CCTR"
            strFormat = strFormat & " WHERE 1=1"
            strFormat = strFormat & " AND USE_YN = 'Y'"
            strFormat = strFormat & " {0}"

            strSQL = String.Format(strFormat, strCondition)

            '데이터 조회
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetCC_HIGH")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region


#End Region

#Region "GROUP BLOCk : 외부에 비공개 Method"
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
