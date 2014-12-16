'****************************************************************************************
'시스템구분 : SFAR/표준샘플/Server단 Control Component
'실행  환경 : COM+ Service Server Package
'프로그램명 : ccMDCOGET.vb (공통코드 조회 Control Class)
'기      능 : 공통코드 조회를 위한 클래스
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/07/05 By Kim Tae Yub
'           :2) 
'****************************************************************************************
Imports SCGLControl                 'Control Class의 Base Class
Imports SCGLUtil.cbSCGLConfig       'Configuration 클래스
Imports SCGLUtil.cbSCGLErr          '오류처리 클래스
Imports SCGLUtil.cbSCGLXml          'XML처리 클래스
Imports SCGLUtil.cbSCGLUtil         '기타 유틸리티 클래스

Public Class ccMDCOGET
    Inherits ccControl

#Region "GROUP BLOCk : 전역 또는 모듈레벨의 변수/상수 선언"
    Private Const CLASS_NAME = "ccMDCOGET"    '자신의 클래스명
    'Private Const .DBConnStr = "Provider=SQLOLEDB;Data Source=10.110.10.86;Initial Catalog=MCDEV;DSN=MCDEV;UID=devadmin;Pwd=password" '커넥션Setting
#End Region

#Region "GROUP BLOCk : Property 선언"
#End Region

#Region "GROUP BLOCk : Event 선언"
#End Region

#Region "GROUP BLOCk : 외부에 공개 Method"

#Region "1. EMP: EMP(사원) 조회"
    Public Function GetMDEMP(ByVal strInfoXML As String, _
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

        Con1 = "" : Con2 = "" : Con3 = "" : Con4 = "" : Con5 = ""

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

            strFormat = " SELECT {0} FROM SC_EMPLOYEE_MST A " & _
                        " WHERE USE_YN = 'Y'  {1} {2} {3} {4} {5} " & _
                        " ORDER BY CC_CODE "

            strSQL = String.Format(strFormat, strSelFields, Con1, Con2, Con3, Con4, Con5)

            '데이터 조회
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetMDEMP")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "2. USER: 로그인사용자 USER 조회"
    Public Function GetUser(ByVal strInfoXML As String, _
                           ByRef intRowCnt As Integer, ByRef intColCnt As Integer, _
                           Optional ByVal strCODE_NAME As String = "", _
                           Optional ByVal strCC_CODE As String = "", _
                           Optional ByVal strAddFields As String = "", _
                           Optional ByVal blnLikeCode As Boolean = True) As Object

        Dim strSQL As String            'SQL문
        Dim strFormat As String         '임시 SQL문
        Dim strCondition As String      '조건절
        Dim strSelFields As String      '조회필드
        Dim strChkDate As String = ""   '사용여부 및 사용날짜
        Dim vntData As Object


        SetConfig(strInfoXML)   '기본정보 설정
        With mobjSCGLConfig

            '1.조회필드 설정
            If strAddFields <> "" Then strAddFields = "," & AddAlias(strAddFields, "A")
            strSelFields = "A.EMPNO, A.EMP_NAME, A.SC_JOB_GRADE_CODE, A.SC_JOB_GRADE_NAME, A.CC_CODE, A.PU_CODE, A.SC_MU_CODE " & strAddFields

            '2.조건절 설정
            If strCC_CODE <> "" Then
                strCondition = String.Format(" AND A.CC_CODE ='{0}' ", strCC_CODE)
            End If

            If strCODE_NAME <> "" Then
                '조회 조건이 코드인지 코드명인지 구분하여 조건 필드 구성
                If IsNumeric(strCODE_NAME) Then '숫자인 경우
                    If Not blnLikeCode Then
                        strCondition &= String.Format("AND A.EMPNO={0}", strCODE_NAME)
                    Else
                        strCondition &= String.Format("AND A.EMPNO LIKE '%{0}%'", strCODE_NAME)
                    End If
                ElseIf IsSBCS(strCODE_NAME) Then    '영문인 경우
                    strCondition &= String.Format("AND (A.EMPNO LIKE '%{0}%' OR A.EMP_NAME LIKE '%{0}%')", strCODE_NAME)
                Else                                 '한글인 경우
                    strCondition &= String.Format("AND A.EMP_NAME LIKE '%{0}%'", strCODE_NAME)
                End If
            End If


            ''3.사용일자 여부 검사 (EDATE 사용하지 않음??)
            'If blnUseOnly Then
            '    strChkDate = "AND A.USE_YN='Y'"
            'Else
            '    strChkDate = String.Format("AND (A.USE_YN='Y' OR A.EDATE>={0})", BuildToDate(strUseDate))
            'End If

            '임시 SQL문 설정
            strFormat = "SELECT {0} FROM SC_USER_INFO_V A " & _
                        "WHERE A.SC_BU_CODE='{1}' {2} {3} " & _
                        "ORDER BY A.EMPNO"

            'SQL문 설정
            strSQL = String.Format(strFormat, strSelFields, .USRCompany, strCondition, strChkDate)

            '데이터 조회
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

#Region "3. 청구지(상호명)조회"
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
        strFormat = strFormat & "  COMPANYNAME, "
        strFormat = strFormat & "  GREATCODE, "
        strFormat = strFormat & "  GREATNAME "
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

    '2009/10/15 케이블 신탁수정하면서 광고주 팝업 수정. 
#Region "3. 청구지조회(광고처포함)"
    ' =============== SelectRtnSample Code
    Public Function GetHIGHCUSTCODE_ALL(ByVal strInfoXML As String, _
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
        strFormat = strFormat & "  COMPANYNAME, "
        strFormat = strFormat & "  GREATCODE, "
        strFormat = strFormat & "  GREATNAME "
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

    '2009/12/11 매체사 팝업용 조회
#Region "3. 매체사조회(GetHIGHCUSTCODE 에서 분리)"
    ' =============== SelectRtnSample Code
    Public Function GetREAL_MED_CODE(ByVal strInfoXML As String, _
                                     ByRef intRowCnt As Integer, _
                                     ByRef intColCnt As Integer, _
                                     ByRef strCUSTCODE As String, _
                                     ByRef strCUSTNAME As String, _
                                     ByRef strMEDDIV As String) As Object

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

        ''MED_TV, MED_RD, MED_DMB, MED_CATV, MED_PAP, MED_MAG, MED_NET, MED_OUT, MED_ELEC, MED_GEN
        If strMEDDIV <> "" Then
            If strMEDDIV = "MED_PAP" Then
                Con3 = " AND ( MED_PAP = '1' OR MED_MAG = '1') "
            Else
                Con3 = String.Format(" AND ( {0} = '1')", strMEDDIV)
            End If
        End If

        If strMEDDIV <> "" Then

            strSQL = "  SELECT "
            strSQL = strSQL & "  HIGHCUSTCODE, "
            strSQL = strSQL & "  CUSTNAME, "
            strSQL = strSQL & "  BUSINO, "
            strSQL = strSQL & "  COMPANYNAME, "
            strSQL = strSQL & "  GREATCODE,"
            strSQL = strSQL & "  GREATNAME "
            strSQL = strSQL & "  FROM SC_CUST_HDR"
            strSQL = strSQL & "  WHERE 1=1 AND MEDFLAG = 'B' AND USE_FLAG = '1'"
            strSQL = strSQL & "  " & Con1
            strSQL = strSQL & "  " & Con2
            strSQL = strSQL & "  AND HIGHCUSTCODE IN("
            strSQL = strSQL & "       SELECT HIGHCUSTCODE FROM SC_CUST_DTL "
            strSQL = strSQL & "       WHERE MEDFLAG = 'B' "
            strSQL = strSQL & "       " & Con3
            strSQL = strSQL & "       GROUP BY HIGHCUSTCODE )"
        Else

            strSQL = " SELECT "
            strSQL = strSQL & "  HIGHCUSTCODE,"
            strSQL = strSQL & "  CUSTNAME,"
            strSQL = strSQL & "  BUSINO,"
            strSQL = strSQL & "  COMPANYNAME, "
            strSQL = strSQL & "  GREATCODE, "
            strSQL = strSQL & "  GREATNAME "
            strSQL = strSQL & "  FROM SC_CUST_HDR"
            strSQL = strSQL & "  WHERE MEDFLAG = 'B' AND USE_FLAG = '1' "
            strSQL = strSQL & "  " & Con1
            strSQL = strSQL & "  " & Con2
            strSQL = strSQL & " ORDER BY CUSTNAME "

        End If

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
                Throw RaiseSysErr(err, CLASS_NAME & ".GetREAL_MED_CODE")
            Finally
                ' DB 접속 종료
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "3. 매체사별 매체조회"
    Public Function SelectRtn_REAL_MED_LIST(ByVal strInfoXML As String, _
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

        If strCUSTCODE <> "" Then Con1 = String.Format(" AND (A.HIGHCUSTCODE = '{0}')", strCUSTCODE)
        If strCUSTNAME <> "" Then Con2 = String.Format(" AND (A.CUSTNAME LIKE '%{0}%')", strCUSTNAME)


        strWhere = BuildFields(" ", Con1, Con2)

        strFormat = " select"
        strFormat = strFormat & " a.highcustcode REAL_MED_CODE, a.companyname REAL_MED_NAME, a.busino REAL_MED_BUSINO, "
        strFormat = strFormat & " b.custcode MEDCODE, b.custname MEDNAME"
        strFormat = strFormat & " from sc_cust_hdr a"
        strFormat = strFormat & " left join sc_cust_dtl b"
        strFormat = strFormat & " on a.highcustcode = b.highcustcode"
        strFormat = strFormat & " where a.medflag = 'B' {0}"
        strFormat = strFormat & " order by a.companyname, a.busino"



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


#Region "4. 사업부조회"
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

    '2009/10/15 케이블 신탁수정하면서 사업부 팝업 수정. 
#Region "4. 사업부조회 (광고처포함)"
    ' =============== SelectRtnSample Code
    Public Function GetCLIENTSUBCODE_ALL(ByVal strInfoXML As String, _
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
        strFormat = strFormat & "  DBO.SC_GET_HIGHCUSTNAME_FUN(HIGHCUSTCODE) COMPANYNAME, "
        strFormat = strFormat & "  DBO.SC_GET_CLIENTCODE_GREATCODE_FUN(HIGHCUSTCODE) GREATCODE, "
        strFormat = strFormat & "  DBO.SC_GET_CLIENTCODE_GREATNAME_FUN(HIGHCUSTCODE) GREATNAME "
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

#Region "5. 대표브랜드조회"
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

#Region "6. CC_CODE: CC 코드조회(Only CC)"
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


#Region "6_1. CC_CODE: CC 코드조회(CGV 모든 부서)"
    Public Function GetCC_CGV(ByVal strInfoXML As String, _
                              ByRef intRowCnt As Integer, _
                              ByRef intColCnt As Integer, _
                              ByVal strDEPT_CODE As String, _
                              ByVal strDEPT_NAME As String) As Object

        Dim strSQL, strFormat As String
        Dim strWhere, strWhere2 As String
        Dim strChkDate As String = ""
        Dim vntData As Object
        Dim Con1, Con2, Con3, Con4

        SetConfig(strInfoXML)   '기본정보 설정
        With mobjSCGLConfig

            Con1 = "" : Con2 = "" : Con3 = "" : Con4 = ""

            If strDEPT_CODE <> "" Then Con1 = String.Format(" AND (CC_CODE LIKE '%{0}%')", strDEPT_CODE)
            If strDEPT_NAME <> "" Then Con2 = String.Format(" AND (CC_NAME LIKE '%{0}%')", strDEPT_NAME)


            If strDEPT_CODE <> "" Then Con3 = String.Format(" AND (HIGHCUSTCODE LIKE '%{0}%')", strDEPT_CODE)
            If strDEPT_NAME <> "" Then Con4 = String.Format(" AND (CUSTNAME LIKE '%{0}%')", strDEPT_NAME)


            strWhere = BuildFields(" ", Con1, Con2)
            strWhere2 = BuildFields(" ", Con3, Con4)

            '조회 필드 설정

            strFormat = "select"
            strFormat = strFormat & " CC_CODE,CC_NAME"
            strFormat = strFormat & " FROM SC_CC A WHERE 1=1"
            strFormat = strFormat & " AND PC='Y' AND USE_YN = 'Y' {0}"
            strFormat = strFormat & " UNION ALL"
            strFormat = strFormat & " SELECT HIGHCUSTCODE, CUSTNAME "
            strFormat = strFormat & " FROM SC_CUST_HDR A WHERE 1=1 AND MEDFLAG = 'G' AND USE_FLAG='1' {1}"

            strSQL = String.Format(strFormat, strWhere, strWhere2)

            '데이터 조회
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetCC_CGV")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region


#Region "7. 팀조회"
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
        strFormat = strFormat & "  DBO.SC_GET_HIGHCUSTNAME_FUN(HIGHCUSTCODE) COMPANYNAME "
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

    '2009/10/15 케이블 신탁수정하면서 팀 팝업 수정.
#Region "7. 팀조회(광고처포함)"
    ' =============== SelectRtnSample Code
    Public Function GetTIMCODE_ALL(ByVal strInfoXML As String, _
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
        strFormat = strFormat & "  DBO.SC_GET_HIGHCUSTNAME_FUN(HIGHCUSTCODE) COMPANYNAME, "
        strFormat = strFormat & "  DBO.SC_GET_CLIENTCODE_GREATCODE_FUN(HIGHCUSTCODE) GREATCODE, "
        strFormat = strFormat & "  DBO.SC_GET_CLIENTCODE_GREATNAME_FUN(HIGHCUSTCODE) GREATNAME "
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

#Region "8. 브랜드별 (광고주,사업부,팁,담당부서) 조회"
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
        strFormat = strFormat & "  DBO.SC_DEPT_NAME_FUN(DEPT_CD) DEPT_NAME"
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

    '2009/10/15 케이블 신탁수정하면서 브랜드 팝업 수정.
#Region "8. 브랜드별 (광고처포함) 조회"
    ' =============== Get_BrandInfo
    Public Function Get_BrandInfo_ALL(ByVal strInfoXML As String, _
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
        strFormat = strFormat & "  DBO.SC_GET_CLIENTCODE_GREATCODE_FUN(CUSTCODE) GREATCODE, "
        strFormat = strFormat & "  DBO.SC_GET_CLIENTCODE_GREATNAME_FUN(CUSTCODE) GREATNAME, "
        strFormat = strFormat & "  TIMCODE,"
        strFormat = strFormat & "  DBO.SC_GET_CUSTNAME_FUN(TIMCODE) TIMNAME,"
        strFormat = strFormat & "  CLIENTSUBCODE,"
        strFormat = strFormat & "  DBO.SC_GET_CUSTNAME_FUN(CLIENTSUBCODE) CLIENTSUBNAME,"
        strFormat = strFormat & "  DEPT_CD,"
        strFormat = strFormat & "  DBO.SC_DEPT_NAME_FUN(DEPT_CD) DEPT_NAME"
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

    '2009/12/27 신탁용은 Interactive Marketing팀이 담당인 브랜드는 뜨지 않도록 해야 함
#Region "8. 브랜드별 (신탁용) 조회"
    ' =============== Get_BrandInfo
    Public Function Get_BrandInfo_TIMCODE(ByVal strInfoXML As String, _
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
        strFormat = strFormat & "  DBO.SC_GET_CLIENTCODE_GREATCODE_FUN(CUSTCODE) GREATCODE, "
        strFormat = strFormat & "  DBO.SC_GET_CLIENTCODE_GREATNAME_FUN(CUSTCODE) GREATNAME, "
        strFormat = strFormat & "  TIMCODE,"
        strFormat = strFormat & "  DBO.SC_GET_CUSTNAME_FUN(TIMCODE) TIMNAME,"
        strFormat = strFormat & "  CLIENTSUBCODE,"
        strFormat = strFormat & "  DBO.SC_GET_CUSTNAME_FUN(CLIENTSUBCODE) CLIENTSUBNAME,"
        strFormat = strFormat & "  DEPT_CD,"
        strFormat = strFormat & "  DBO.SC_DEPT_NAME_FUN(DEPT_CD) DEPT_NAME"
        strFormat = strFormat & "  FROM SC_SUBSEQ_DTL"
        strFormat = strFormat & "  WHERE ISNULL(DEPT_CD,'') <> '10000041' AND ISNULL(ATTR01,'N') = 'Y' {0} ORDER BY SEQNAME"

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

#Region "9. 크리조직 조회"

    ' =============== SelectRtnSample Code
    Public Function GetCRECUSTNO(ByVal strInfoXML As String, _
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
                SetConfig(strInfoXML)   '기본정보 설정

                Con1 = ""
                Con2 = ""

                If strCUSTCODE <> "" Then Con1 = String.Format(" AND (CUSTCODE LIKE '%{0}%')", strCUSTCODE)
                If strCUSTNAME <> "" Then Con2 = String.Format(" AND (CUSTNAME LIKE '%{0}%')", strCUSTNAME)
                strWhere = BuildFields(" ", Con1, Con2)

                strSelFields = "A.CUSTCODE, A.CUSTNAME, B.BUSINO, B.COMPANYNAME"

                strFormet = "select {0} from SC_CUST_DTL A, SC_CUST_HDR B where 1=1 AND A.HIGHCUSTCODE= B.HIGHCUSTCODE AND A.CUSTCODE LIKE 'K%' AND A.USE_FLAG =1 {1} "
                strFormet = strFormet & " ORDER BY A.CUSTNAME "


                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetCRECUSTNO")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "10. 제작대행사 (대행사) 조회"
    ' =============== Get_BrandInfo
    Public Function Get_EXCLIENT(ByVal strInfoXML As String, _
                                 ByRef intRowCnt As Integer, _
                                 ByRef intColCnt As Integer, _
                                 ByVal strEXCLIENTCODE As String, _
                                 ByVal strEXCLIENTNAME As String) As Object

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

        If strEXCLIENTCODE <> "" Then Con1 = String.Format(" AND (EXCLIENTCODE = '{0}')", strEXCLIENTCODE)
        If strEXCLIENTNAME <> "" Then Con2 = String.Format(" AND (CUSTNAME LIKE '%{0}%')", strEXCLIENTNAME)
        'If strFLAG <> "" Then Con3 = String.Format(" AND (MEDFLAG = '{0}')", strFLAG)


        strWhere = BuildFields(" ", Con1, Con2)

        strFormat = "  SELECT "
        strFormat = strFormat & "  	MEDFLAG,  "
        strFormat = strFormat & "  	CASE MEDFLAG "
        strFormat = strFormat & "  	WHEN 'G' THEN '대행사'"
        strFormat = strFormat & "  	WHEN 'K' THEN '크리조직'"
        strFormat = strFormat & "  	END AS	MEDFLAGNAME,"
        strFormat = strFormat & "  	HIGHCUSTCODE EXCLIENTCODE,"
        strFormat = strFormat & "  	CUSTNAME EXCLIENTNAME "
        strFormat = strFormat & "  	FROM SC_CUST_HDR"
        strFormat = strFormat & "  	WHERE MEDFLAG IN ('G','K') AND USE_FLAG = '1' "
        strFormat = strFormat & "   {0} ORDER BY MEDFLAG , EXCLIENTNAME"

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
                Throw RaiseSysErr(err, CLASS_NAME & ".Get_EXCLIENT")
            Finally
                ' DB 접속 종료
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
        strFormat = strFormat & "  	CUSTNAME EXCLIENTNAME, CASE MEDFLAG WHEN 'G' THEN '2' WHEN 'K' THEN '1' END AS SORTGBN"
        strFormat = strFormat & "  	FROM SC_CUST_HDR"
        strFormat = strFormat & "  	WHERE MEDFLAG IN('G','K') AND USE_FLAG = '1'"
        strFormat = strFormat & "  	UNION ALL"
        strFormat = strFormat & "  	SELECT"
        strFormat = strFormat & "  	'C' MEDFLAG,  "
        strFormat = strFormat & "  	'CD' MEDFLAGNAME,"
        strFormat = strFormat & "  	EMPNO EXCLIENTCODE,"
        strFormat = strFormat & "  	EMP_NAME EXCLIENTNAME, '3' SORTGBN"
        strFormat = strFormat & "  	FROM SC_EMPLOYEE_MST"
        strFormat = strFormat & "  	WHERE USE_YN = 'Y'"
        strFormat = strFormat & "  )AAA"
        strFormat = strFormat & "  WHERE 1=1 {0} ORDER BY SORTGBN, EXCLIENTNAME"

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

#Region "11. 매체조회"
    ' =============== SelectRtnSample Code
    Public Function GetMEDCODE(ByVal strInfoXML As String, _
                               ByRef intRowCnt As Integer, _
                               ByRef intColCnt As Integer, _
                               ByVal strREAL_MED_CODE As String, _
                               ByVal strREAL_MED_NAME As String, _
                               ByVal strMEDCODE As String, _
                               ByVal strMEDNAME As String) As Object

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

        If strREAL_MED_CODE <> "" Then Con1 = String.Format(" AND (HIGHCUSTCODE = '{0}')", strREAL_MED_CODE)
        If strREAL_MED_NAME <> "" Then Con2 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(HIGHCUSTCODE) LIKE '%{0}%')", strREAL_MED_NAME)
        If strMEDCODE <> "" Then Con3 = String.Format(" AND (CUSTCODE = '{0}')", strMEDCODE)
        If strMEDNAME <> "" Then Con4 = String.Format(" AND (CUSTNAME LIKE '%{0}%')", strMEDNAME)


        strWhere = BuildFields(" ", Con1, Con2, Con3, Con4)

        strFormat = "  SELECT "
        strFormat = strFormat & "  CUSTCODE,"
        strFormat = strFormat & "  CUSTNAME,"
        strFormat = strFormat & "  dbo.SC_GET_BUSINO_FUN(HIGHCUSTCODE) BUSINO,"
        strFormat = strFormat & "  HIGHCUSTCODE,"
        strFormat = strFormat & "  DBO.SC_GET_HIGHCUSTNAME_FUN(HIGHCUSTCODE) COMPANYNAME "
        strFormat = strFormat & "  FROM SC_CUST_DTL"
        strFormat = strFormat & "  WHERE 1=1 AND MEDFLAG = 'B' AND USE_FLAG = '1' {0} ORDER BY CUSTNAME"

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

#Region "12. 소재팝업 조회"
    Public Function GetMATTER(ByVal strInfoXML As String, _
                              ByRef intRowCnt As Integer, _
                              ByRef intColCnt As Integer, _
                              ByRef strCLIENTNAME As String, _
                              ByRef strTIMNAME As String, _
                              ByRef strSUBSEQNAME As String, _
                              ByRef strEXCLIENTNAME As String, _
                              ByRef strMATTERNAME As String, _
                              ByRef strDEPT_NAME As String, _
                              ByRef strMEDFLAG As String) As Object


        Dim strCols As String        '컬럼변수
        Dim strWhere As String       'Where조건 변수
        Dim strFormat As String      'SQL Format 변수
        Dim strSQL As String         'SQL 변수
        Dim vntData As Object        'Array Return 변수(Array 를사용할 때 선언)
        Dim strSelFields
        Dim intCnt
        Dim intRtn
        Dim Con1, Con2, Con3, Con4, Con5, Con6, Con7, Con8
        SetConfig(strInfoXML)

        Con1 = "" : Con2 = "" : Con3 = "" : Con4 = "" : Con5 = "" : Con6 = "" : Con7 = "" : Con8 = ""

        If strCLIENTNAME <> "" Then Con1 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(CUSTCODE) LIKE '%{0}%')", strCLIENTNAME)
        If strTIMNAME <> "" Then Con2 = String.Format(" AND (DBO.SC_GET_SUBSEQ_TIMNAME_FUN(SUBSEQ) LIKE '%{0}%')", strTIMNAME)
        If strSUBSEQNAME <> "" Then Con3 = String.Format(" AND (DBO.SC_GET_SUBSEQNAME_FUN(SUBSEQ) LIKE '%{0}%')", strSUBSEQNAME)
        If strEXCLIENTNAME <> "" Then Con4 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(EXCLIENTCODE) LIKE '%{0}%')", strEXCLIENTNAME)
        If strMATTERNAME <> "" Then Con5 = String.Format(" AND (MATTER LIKE '%{0}%')", strMATTERNAME)
        If strDEPT_NAME <> "" Then Con6 = String.Format(" AND (DBO.SC_GET_SUBSEQ_DEPT_NAME_FUN(SUBSEQ) LIKE '%{0}%')", strDEPT_NAME)

        If strMEDFLAG <> "" Then
            If strMEDFLAG = "B" Then
                Con7 = String.Format(" AND (MEDFLAG IN('B'))", strMEDFLAG)
                Con8 = " AND (ISNULL(EXCLIENTCODE,'') <> '') "
            ElseIf strMEDFLAG = "A2" Then
                Con7 = String.Format(" AND (MEDFLAG IN('A2','A','T'))", strMEDFLAG)
                Con8 = " AND (ISNULL(EXCLIENTCODE,'') <> '') "
            ElseIf strMEDFLAG = "A" Then
                Con7 = String.Format(" AND (MEDFLAG IN('A','A2','T'))", strMEDFLAG)
            End If
        End If

        strWhere = BuildFields(" ", Con1, Con2, Con3, Con4, Con5, Con6, Con7, Con8)

        strFormat = strFormat & "  SELECT "
        strFormat = strFormat & "  MATTERCODE, MATTER MATTERNAME, "
        strFormat = strFormat & "  CUSTCODE, DBO.SC_GET_HIGHCUSTNAME_FUN(CUSTCODE) CUSTNAME,"
        strFormat = strFormat & "  DBO.SC_GET_SUBSEQ_TIMCODE_FUN(SUBSEQ) TIMCODE, DBO.SC_GET_SUBSEQ_TIMNAME_FUN(SUBSEQ) TIMNAME,"
        strFormat = strFormat & "  SUBSEQ, DBO.SC_GET_SUBSEQNAME_FUN(SUBSEQ) SUBSEQNAME,"
        strFormat = strFormat & "  EXCLIENTCODE, DBO.SC_GET_HIGHCUSTNAME_FUN(EXCLIENTCODE) EXCLIENTNAME,"
        strFormat = strFormat & "  DBO.SC_GET_SUBSEQ_DEPT_CD_FUN(SUBSEQ) DEPT_CD, DBO.SC_GET_SUBSEQ_DEPT_NAME_FUN(SUBSEQ) DEPT_NAME, "
        strFormat = strFormat & "  DBO.SC_GET_SUBSEQ_CLIENTSUBCODE_FUN(SUBSEQ) CLIENTSUBCODE, DBO.SC_GET_SUBSEQ_CLIENTSUBNAME_FUN(SUBSEQ) CLIENTSUBNAME, "
        strFormat = strFormat & "  DBO.SC_GET_CLIENTCODE_GREATCODE_FUN(CUSTCODE) GREATCODE, DBO.SC_GET_CLIENTCODE_GREATNAME_FUN(CUSTCODE) GREATNAME"
        strFormat = strFormat & "  FROM MD_MATTER_MST"
        strFormat = strFormat & "  WHERE 1=1 AND ISNULL(ATTR02,'N') ='Y' {0} "

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
                Throw RaiseSysErr(err, CLASS_NAME & ".GetMATTER")
            Finally
                ' DB 접속 종료
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

    '2009/10/15 케이블 신탁수정하면서 소재 팝업 수정. (MATTERMST도 수정.)
#Region "12. 소재팝업 조회 (사업부 , 광고처 포함)"
    Public Function GetMATTER_ALL(ByVal strInfoXML As String, _
                                  ByRef intRowCnt As Integer, _
                                  ByRef intColCnt As Integer, _
                                  ByRef strCLIENTNAME As String, _
                                  ByRef strTIMNAME As String, _
                                  ByRef strSUBSEQNAME As String, _
                                  ByRef strEXCLIENTNAME As String, _
                                  ByRef strMATTERNAME As String, _
                                  ByRef strDEPT_NAME As String, _
                                  ByRef strMEDFLAG As String) As Object


        Dim strCols As String        '컬럼변수
        Dim strWhere As String       'Where조건 변수
        Dim strFormat As String      'SQL Format 변수
        Dim strSQL As String         'SQL 변수
        Dim vntData As Object        'Array Return 변수(Array 를사용할 때 선언)
        Dim strSelFields
        Dim intCnt
        Dim intRtn
        Dim Con1, Con2, Con3, Con4, Con5, Con6, Con7, Con8
        SetConfig(strInfoXML)

        Con1 = "" : Con2 = "" : Con3 = "" : Con4 = "" : Con5 = "" : Con6 = "" : Con7 = "" : Con8 = ""

        If strCLIENTNAME <> "" Then Con1 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(CUSTCODE) LIKE '%{0}%')", strCLIENTNAME)
        If strTIMNAME <> "" Then Con2 = String.Format(" AND (DBO.SC_GET_SUBSEQ_TIMNAME_FUN(SUBSEQ) LIKE '%{0}%')", strTIMNAME)
        If strSUBSEQNAME <> "" Then Con3 = String.Format(" AND (DBO.SC_GET_SUBSEQNAME_FUN(SUBSEQ) LIKE '%{0}%')", strSUBSEQNAME)
        If strEXCLIENTNAME <> "" Then Con4 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(EXCLIENTCODE) LIKE '%{0}%')", strEXCLIENTNAME)
        If strMATTERNAME <> "" Then Con5 = String.Format(" AND (MATTER LIKE '%{0}%')", strMATTERNAME)
        If strDEPT_NAME <> "" Then Con6 = String.Format(" AND (DBO.SC_GET_SUBSEQ_DEPT_NAME_FUN(SUBSEQ) LIKE '%{0}%')", strDEPT_NAME)

        If strMEDFLAG <> "" Then
            If strMEDFLAG = "B" Then
                Con7 = String.Format(" AND (MEDFLAG IN('B'))", strMEDFLAG)
                Con8 = " AND (ISNULL(EXCLIENTCODE,'') <> '') "
            ElseIf strMEDFLAG = "A2" Then
                Con7 = String.Format(" AND (MEDFLAG IN('A2', 'A','T'))", strMEDFLAG)
                Con8 = " AND (ISNULL(EXCLIENTCODE,'') <> '') "
            ElseIf strMEDFLAG = "A" Then
                Con7 = String.Format(" AND (MEDFLAG IN('A','A2','T'))", strMEDFLAG)
            End If
        End If

        strWhere = BuildFields(" ", Con1, Con2, Con3, Con4, Con5, Con6, Con7, Con8)

        strFormat = " SELECT "
        strFormat = strFormat & "  MATTERCODE, "
        strFormat = strFormat & "  MATTER MATTERNAME, "
        strFormat = strFormat & "  CUSTCODE, DBO.SC_GET_HIGHCUSTNAME_FUN(CUSTCODE) CUSTNAME,"
        strFormat = strFormat & "  DBO.SC_GET_SUBSEQ_TIMCODE_FUN(SUBSEQ) TIMCODE, "
        strFormat = strFormat & "  DBO.SC_GET_SUBSEQ_TIMNAME_FUN(SUBSEQ) TIMNAME,"
        strFormat = strFormat & "  SUBSEQ, "
        strFormat = strFormat & "  DBO.SC_GET_SUBSEQNAME_FUN(SUBSEQ) SUBSEQNAME,"
        strFormat = strFormat & "  EXCLIENTCODE, "
        strFormat = strFormat & "  DBO.SC_GET_HIGHCUSTNAME_FUN(EXCLIENTCODE) EXCLIENTNAME,"
        strFormat = strFormat & "  DBO.SC_GET_SUBSEQ_DEPT_CD_FUN(SUBSEQ) DEPT_CD, "
        strFormat = strFormat & "  DBO.SC_GET_SUBSEQ_DEPT_NAME_FUN(SUBSEQ) DEPT_NAME, "
        strFormat = strFormat & "  DBO.SC_GET_SUBSEQ_CLIENTSUBCODE_FUN(SUBSEQ) CLIENTSUBCODE, "
        strFormat = strFormat & "  DBO.SC_GET_SUBSEQ_CLIENTSUBNAME_FUN(SUBSEQ) CLIENTSUBNAME, "
        strFormat = strFormat & "  DBO.SC_GET_CLIENTCODE_GREATCODE_FUN(CUSTCODE) GREATCODE, "
        strFormat = strFormat & "  DBO.SC_GET_CLIENTCODE_GREATNAME_FUN(CUSTCODE) GREATNAME"
        strFormat = strFormat & "  FROM MD_MATTER_MST"
        strFormat = strFormat & "  WHERE 1=1 "
        strFormat = strFormat & "  AND ISNULL(ATTR02,'N') ='Y' "
        strFormat = strFormat & "  {0} "

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
                Throw RaiseSysErr(err, CLASS_NAME & ".GetMATTER_ALL")
            Finally
                ' DB 접속 종료
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "13. 거래명세서광고주 POP조회"
    Public Function GetTRANSCUSTNO(ByVal strInfoXML As String, _
                                   ByRef intRowCnt As Integer, _
                                   ByRef intColCnt As Integer, _
                                   ByVal strYEARMON As String, _
                                   ByVal strCLIENTCODE As String, _
                                   ByVal strCLIENTNAME As String, _
                                   ByVal strCOMMITCHECK As String, _
                                   ByVal strFlag As String, _
                                   ByVal strTBL_Flag As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2, Con3, Con4, Con5 As String
        Dim vntData As Object


        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '기본정보 설정

                Con1 = "" : Con2 = "" : Con3 = ""

                If strTBL_Flag = "PRINT" Or strTBL_Flag = "CATV" Or strTBL_Flag = "INTERNET" Or strTBL_Flag = "ELECSPON" Or strTBL_Flag = "TOTAL" Or strTBL_Flag = "AOR" Then
                    If strFlag = "trans" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
                            If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (CLIENTCODE = '{0}')", strCLIENTCODE)
                            If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCLIENTNAME)
                        Else
                            If strYEARMON <> "" Then Con1 = String.Format(" AND (SUBSTRING(DEMANDDAY,1,6) = '{0}')", strYEARMON)
                            If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (CLIENTCODE = '{0}')", strCLIENTCODE)
                            If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCLIENTNAME)
                        End If
                    ElseIf strFlag = "commi" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
                            If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (REAL_MED_CODE = '{0}')", strCLIENTCODE)
                            If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) LIKE '%{0}%')", strCLIENTNAME)
                        Else
                            If strYEARMON <> "" Then Con1 = String.Format(" AND (SUBSTRING(DEMANDDAY,1,6) = '{0}')", strYEARMON)
                            If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (REAL_MED_CODE = '{0}')", strCLIENTCODE)
                            If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) LIKE '%{0}%')", strCLIENTNAME)
                        End If
                    End If
                ElseIf strTBL_Flag = "OUTDOOR" Then
                    If strFlag = "commi" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
                            If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (CLIENTCODE = '{0}')", strCLIENTCODE)
                            If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCLIENTNAME)
                        Else
                            If strYEARMON <> "" Then Con1 = String.Format(" AND (SUBSTRING(DEMANDDAY,1,6) = '{0}')", strYEARMON)
                            If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (CLIENTCODE = '{0}')", strCLIENTCODE)
                            If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCLIENTNAME)
                        End If
                    End If
                ElseIf strTBL_Flag = "ELEC" Then
                    If strFlag = "trans" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
                            If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (CLIENTCODE = '{0}')", strCLIENTCODE)
                            If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCLIENTNAME)
                        Else
                            If strYEARMON <> "" Then Con1 = String.Format(" AND (YEARMON = '{0}')", strYEARMON)
                            If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (CLIENTCODE = '{0}')", strCLIENTCODE)
                            If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCLIENTNAME)
                        End If
                    ElseIf strFlag = "commi" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
                            If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (REAL_MED_CODE = '{0}')", strCLIENTCODE)
                            If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) LIKE '%{0}%')", strCLIENTNAME)
                        Else
                            If strYEARMON <> "" Then Con1 = String.Format(" AND (YEARMON = '{0}')", strYEARMON)
                            If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (REAL_MED_CODE = '{0}')", strCLIENTCODE)
                            If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) LIKE '%{0}%')", strCLIENTNAME)
                        End If
                    End If
                End If
                strWhere = BuildFields(" ", Con1, Con2, Con3)


                If strTBL_Flag = "PRINT" Then
                    If strFlag = "trans" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            strSelFields = "TRANSYEARMON , TRANSNO, DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME, '완료' GBN, CLIENTCODE"
                            strFormet = "select {0}  "
                            strFormet = strFormet & " FROM MD_PRINTTRANS_HDR  WHERE 1=1 {1} "
                            strFormet = strFormet & " GROUP BY CLIENTCODE , TRANSYEARMON, TRANSNO "
                            strFormet = strFormet & " ORDER BY DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE)"
                        Else
                            strSelFields = "SUBSTRING(DEMANDDAY,1,6) YEARMON , CLIENTCODE, DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME, '미완료' GBN"
                            strFormet = "select {0}    "
                            strFormet = strFormet & " FROM MD_BOOKING_MEDIUM  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND gflag in('B') "
                            strFormet = strFormet & " AND isnull(tru_trans_no, '') = ''  "
                            strFormet = strFormet & " GROUP BY CLIENTCODE , SUBSTRING(DEMANDDAY,1,6) "
                            strFormet = strFormet & " ORDER BY "
                            strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE)"
                        End If
                    ElseIf strFlag = "commi" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            strSelFields = "TRANSYEARMON , TRANSNO, DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) REAL_MED_NAME, '완료' GBN, REAL_MED_CODE"
                            strFormet = "select {0}  "
                            strFormet = strFormet & " FROM MD_PRINTCOMMI_DTL  WHERE 1=1 {1} "
                            strFormet = strFormet & " GROUP BY REAL_MED_CODE , TRANSYEARMON, TRANSNO "
                            strFormet = strFormet & " ORDER BY DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) "
                        Else
                            strSelFields = "SUBSTRING(DEMANDDAY,1,6) YEARMON , REAL_MED_CODE, DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) REAL_MED_NAME, '미완료' GBN"
                            strFormet = "select {0}    "
                            strFormet = strFormet & " FROM MD_BOOKING_MEDIUM  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND gflag in('B','J', 'S') "
                            strFormet = strFormet & " AND isnull(commi_trans_no, '') = ''  "
                            strFormet = strFormet & " GROUP BY REAL_MED_CODE , SUBSTRING(DEMANDDAY,1,6) "
                            strFormet = strFormet & " ORDER BY "
                            strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE)"
                        End If
                    End If
                ElseIf strTBL_Flag = "CATV" Then
                    If strFlag = "trans" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            strSelFields = "TRANSYEARMON , TRANSNO, DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME, '완료' GBN, CLIENTCODE"
                            strFormet = "select {0}  "
                            strFormet = strFormet & " FROM MD_CATVTRANS_HDR  WHERE 1=1 {1} "
                            strFormet = strFormet & " GROUP BY CLIENTCODE , TRANSYEARMON, TRANSNO "
                            strFormet = strFormet & " ORDER BY DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE)"
                        Else
                            strSelFields = "SUBSTRING(DEMANDDAY,1,6) YEARMON , CLIENTCODE, DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME, '미완료' GBN"
                            strFormet = "select {0}    "
                            strFormet = strFormet & " FROM MD_CATV_MEDIUM  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND gflag in('1') "
                            strFormet = strFormet & " AND isnull(tru_trans_no, '') = ''  "
                            strFormet = strFormet & " GROUP BY CLIENTCODE , SUBSTRING(DEMANDDAY,1,6) "
                            strFormet = strFormet & " ORDER BY "
                            strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE)"
                        End If
                    ElseIf strFlag = "commi" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            strSelFields = "TRANSYEARMON , TRANSNO, DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) REAL_MED_NAME, '완료' GBN, REAL_MED_CODE"
                            strFormet = "select {0}  "
                            strFormet = strFormet & " FROM MD_CATVCOMMI_DTL  WHERE 1=1 {1} "
                            strFormet = strFormet & " GROUP BY REAL_MED_CODE , TRANSYEARMON, TRANSNO "
                            strFormet = strFormet & " ORDER BY DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) "
                        Else
                            strSelFields = "SUBSTRING(DEMANDDAY,1,6) YEARMON , REAL_MED_CODE, DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) REAL_MED_NAME, '미완료' GBN"
                            strFormet = "select {0}    "
                            strFormet = strFormet & " FROM MD_CATV_MEDIUM  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND gflag in('1') "
                            strFormet = strFormet & " AND isnull(commi_trans_no, '') = ''  "
                            strFormet = strFormet & " GROUP BY REAL_MED_CODE , SUBSTRING(DEMANDDAY,1,6) "
                            strFormet = strFormet & " ORDER BY "
                            strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE)"
                        End If
                    End If
                ElseIf strTBL_Flag = "TOTAL" Then
                    If strFlag = "trans" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            strSelFields = "TRANSYEARMON , TRANSNO, DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME, '완료' GBN, CLIENTCODE"
                            strFormet = "select {0}  "
                            strFormet = strFormet & " FROM MD_TOTALTRANS_HDR  WHERE 1=1 {1} "
                            strFormet = strFormet & " GROUP BY CLIENTCODE , TRANSYEARMON, TRANSNO "
                            strFormet = strFormet & " ORDER BY DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE)"
                        Else
                            strSelFields = "SUBSTRING(DEMANDDAY,1,6) YEARMON , CLIENTCODE, DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME, '미완료' GBN"
                            strFormet = "select {0}    "
                            strFormet = strFormet & " FROM MD_TOTAL_MEDIUM  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND gflag in('1') "
                            strFormet = strFormet & " AND isnull(tru_trans_no, '') = ''  "
                            strFormet = strFormet & " GROUP BY CLIENTCODE , SUBSTRING(DEMANDDAY,1,6) "
                            strFormet = strFormet & " ORDER BY "
                            strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE)"
                        End If
                    ElseIf strFlag = "commi" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            strSelFields = "TRANSYEARMON , TRANSNO, DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) REAL_MED_NAME, '완료' GBN, REAL_MED_CODE"
                            strFormet = "select {0}  "
                            strFormet = strFormet & " FROM MD_TOTALCOMMI_DTL  WHERE 1=1 {1} "
                            strFormet = strFormet & " GROUP BY REAL_MED_CODE , TRANSYEARMON, TRANSNO "
                            strFormet = strFormet & " ORDER BY DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) "
                        Else
                            strSelFields = "SUBSTRING(DEMANDDAY,1,6) YEARMON , REAL_MED_CODE, DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) REAL_MED_NAME, '미완료' GBN"
                            strFormet = "select {0}    "
                            strFormet = strFormet & " FROM MD_TOTAL_MEDIUM  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND gflag in('1') "
                            strFormet = strFormet & " AND isnull(commi_trans_no, '') = ''  "
                            strFormet = strFormet & " GROUP BY REAL_MED_CODE , SUBSTRING(DEMANDDAY,1,6) "
                            strFormet = strFormet & " ORDER BY "
                            strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE)"
                        End If
                    End If
                ElseIf strTBL_Flag = "INTERNET" Then
                    If strFlag = "trans" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            strSelFields = "TRANSYEARMON , TRANSNO, DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME, '완료' GBN, CLIENTCODE"
                            strFormet = "select {0}  "
                            strFormet = strFormet & " FROM MD_INTERNETTRANS_HDR  WHERE 1=1 {1} "
                            strFormet = strFormet & " GROUP BY CLIENTCODE , TRANSYEARMON, TRANSNO "
                            strFormet = strFormet & " ORDER BY DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE)"
                        Else
                            strSelFields = "SUBSTRING(DEMANDDAY,1,6) YEARMON , CLIENTCODE, DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME, '미완료' GBN"
                            strFormet = "select {0}    "
                            strFormet = strFormet & " FROM MD_INTERNET_MEDIUM  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND gflag in('1') "
                            strFormet = strFormet & " AND isnull(tru_trans_no, '') = ''  "
                            strFormet = strFormet & " GROUP BY CLIENTCODE , SUBSTRING(DEMANDDAY,1,6) "
                            strFormet = strFormet & " ORDER BY "
                            strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE)"
                        End If
                    ElseIf strFlag = "commi" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            strSelFields = "TRANSYEARMON , TRANSNO, DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) REAL_MED_NAME, '완료' GBN, REAL_MED_CODE"
                            strFormet = "select {0}  "
                            strFormet = strFormet & " FROM MD_INTERNETCOMMI_DTL  WHERE 1=1 {1} "
                            strFormet = strFormet & " GROUP BY REAL_MED_CODE , TRANSYEARMON, TRANSNO "
                            strFormet = strFormet & " ORDER BY DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) "
                        Else
                            strSelFields = "SUBSTRING(DEMANDDAY,1,6) YEARMON , REAL_MED_CODE, DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) REAL_MED_NAME, '미완료' GBN"
                            strFormet = "select {0}    "
                            strFormet = strFormet & " FROM MD_INTERNET_MEDIUM  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND gflag in('1') "
                            strFormet = strFormet & " AND isnull(commi_trans_no, '') = ''  "
                            strFormet = strFormet & " GROUP BY REAL_MED_CODE , SUBSTRING(DEMANDDAY,1,6) "
                            strFormet = strFormet & " ORDER BY "
                            strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE)"
                        End If
                    End If
                ElseIf strTBL_Flag = "ELEC" Then
                    If strFlag = "trans" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            strSelFields = "TRANSYEARMON , TRANSNO, DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME, '완료' GBN, CLIENTCODE"
                            strFormet = "select {0}  "
                            strFormet = strFormet & " FROM MD_ELEC_TRANS_HDR  WHERE 1=1 {1} "
                            strFormet = strFormet & " GROUP BY CLIENTCODE , TRANSYEARMON, TRANSNO "
                            strFormet = strFormet & " ORDER BY DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE)"
                        Else
                            strSelFields = "YEARMON , CLIENTCODE, DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME, '미완료' GBN"
                            strFormet = "select {0}    "
                            strFormet = strFormet & " FROM MD_ELECTRIC_MEDIUM  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND isnull(tru_trans_no, '') = ''  "
                            strFormet = strFormet & " GROUP BY CLIENTCODE , YEARMON "
                            strFormet = strFormet & " ORDER BY "
                            strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE)"
                        End If
                    ElseIf strFlag = "commi" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            strSelFields = "TRANSYEARMON , TRANSNO, DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) REAL_MED_NAME, '완료' GBN, REAL_MED_CODE"
                            strFormet = "select {0}  "
                            strFormet = strFormet & " FROM MD_ELECCOMMI_DTL  WHERE 1=1 {1} "
                            strFormet = strFormet & " GROUP BY REAL_MED_CODE , TRANSYEARMON, TRANSNO "
                            strFormet = strFormet & " ORDER BY DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) "
                        Else
                            strSelFields = "YEARMON , REAL_MED_CODE, DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) REAL_MED_NAME, '미완료' GBN"
                            strFormet = "select {0}    "
                            strFormet = strFormet & " FROM MD_ELECTRIC_MEDIUM  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND isnull(commi_trans_no, '') = ''  "
                            strFormet = strFormet & " GROUP BY REAL_MED_CODE , YEARMON  "
                            strFormet = strFormet & " ORDER BY "
                            strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE)"
                        End If
                    End If
                ElseIf strTBL_Flag = "OUTDOOR" Then
                    If strFlag = "commi" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            strSelFields = "TRANSYEARMON , TRANSNO, DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) REAL_MED_NAME, '완료' GBN, CLIENTCODE REAL_MED_CODE"
                            strFormet = "select {0}  "
                            strFormet = strFormet & " FROM MD_OUTDOORCOMMI_HDR  WHERE 1=1 {1} "
                            strFormet = strFormet & " GROUP BY CLIENTCODE , TRANSYEARMON, TRANSNO "
                            strFormet = strFormet & " ORDER BY DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) "
                        Else
                            strSelFields = "SUBSTRING(DEMANDDAY,1,6) YEARMON , CLIENTCODE REAL_MED_CODE, DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) REAL_MED_NAME, '미완료' GBN"
                            strFormet = "select {0}    "
                            strFormet = strFormet & " FROM MD_OUTDOOR_MEDIUM  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND isnull(commi_trans_no, '') = ''  "
                            strFormet = strFormet & " GROUP BY CLIENTCODE , SUBSTRING(DEMANDDAY,1,6) "
                            strFormet = strFormet & " ORDER BY "
                            strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE)"
                        End If
                    End If
                ElseIf strTBL_Flag = "ELECSPON" Then
                    If strFlag = "commi" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            strSelFields = "TRANSYEARMON , TRANSNO, DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) REAL_MED_NAME, '완료' GBN, CLIENTCODE REAL_MED_CODE"
                            strFormet = "select {0}  "
                            strFormet = strFormet & " FROM MD_ELECCOMMI_HDR  WHERE 1=1 {1} "
                            strFormet = strFormet & " GROUP BY CLIENTCODE , TRANSYEARMON, TRANSNO "
                            strFormet = strFormet & " ORDER BY DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) "
                        Else
                            strSelFields = "SUBSTRING(DEMANDDAY,1,6) YEARMON , CLIENTCODE REAL_MED_CODE, DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) REAL_MED_NAME, '미완료' GBN"
                            strFormet = "select {0}    "
                            strFormet = strFormet & " FROM MD_ELECTRIC_MEDIUM  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND isnull(commi_trans_no, '') = ''  AND ISNULL(SPONSOR,'N') = 'Y'"
                            strFormet = strFormet & " GROUP BY CLIENTCODE , SUBSTRING(DEMANDDAY,1,6) "
                            strFormet = strFormet & " ORDER BY "
                            strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE)"
                        End If
                    End If
                ElseIf strTBL_Flag = "AOR" Then
                    If strFlag = "commi" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            strSelFields = "TRANSYEARMON , TRANSNO, DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) REAL_MED_NAME, '완료' GBN, REAL_MED_CODE"
                            strFormet = "select {0}  "
                            strFormet = strFormet & " FROM MD_AORCOMMI_HDR  WHERE 1=1 {1} "
                            strFormet = strFormet & " GROUP BY REAL_MED_CODE , TRANSYEARMON, TRANSNO "
                            strFormet = strFormet & " ORDER BY DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) "
                        Else
                            strSelFields = "SUBSTRING(DEMANDDAY,1,6) YEARMON , REAL_MED_CODE, DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) REAL_MED_NAME, '미완료' GBN"
                            strFormet = "select {0}    "
                            strFormet = strFormet & " FROM MD_AOR_MEDIUM  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND isnull(commi_trans_no, '') = '' "
                            strFormet = strFormet & " GROUP BY REAL_MED_CODE , SUBSTRING(DEMANDDAY,1,6) "
                            strFormet = strFormet & " ORDER BY "
                            strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE)"
                        End If
                    End If
                 ElseIf strTBL_Flag = "POINTAD" Then
                    If strFlag = "commi" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            strSelFields = "TRANSYEARMON , TRANSNO, DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) REAL_MED_NAME, '완료' GBN, CLIENTCODE REAL_MED_CODE"
                            strFormet = " SELECT {0} "
                            strFormet = strFormet & " FROM MD_POINTADCOMMI_HDR  WHERE 1=1 {1} "
                            strFormet = strFormet & " GROUP BY CLIENTCODE , TRANSYEARMON, TRANSNO "
                            strFormet = strFormet & " ORDER BY DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) "

                        Else
                            strSelFields = "SUBSTRING(DEMANDDAY,1,6) YEARMON , CLIENTCODE REAL_MED_CODE, DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) REAL_MED_NAME, '미완료' GBN"
                            strFormet = " SELECT {0} "
                            strFormet = strFormet & " FROM MD_POINTAD_AMT  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND ISNULL(COMMI_TRANS_NO, '') = '' "
                            strFormet = strFormet & " GROUP BY CLIENTCODE , SUBSTRING(DEMANDDAY,1,6) "
                            strFormet = strFormet & " ORDER BY DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE)"
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

#Region "14. 거래명세서 팀 POP조회"
    Public Function GetTRANSTIMCODE(ByVal strInfoXML As String, _
                                    ByRef intRowCnt As Integer, _
                                    ByRef intColCnt As Integer, _
                                    ByVal strYEARMON As String, _
                                    ByVal strCLIENTCODE As String, _
                                    ByVal strCLIENTNAME As String, _
                                    ByVal strTIMCODE As String, _
                                    ByVal strTIMNAME As String, _
                                    ByVal strCOMMITCHECK As String, _
                                    ByVal strFlag As String, _
                                    ByVal strTBL_Flag As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2, Con3, Con4, Con5 As String
        Dim vntData As Object


        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '기본정보 설정

                Con1 = "" : Con2 = "" : Con3 = "" : Con4 = "" : Con5 = ""

                If strTBL_Flag = "PRINT" Or strTBL_Flag = "CATV" Or strTBL_Flag = "INTERNET" Or strTBL_Flag = "TOTAL" Then
                    If strFlag = "trans" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
                            If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (CLIENTCODE = '{0}')", strCLIENTCODE)
                            If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCLIENTNAME)
                            If strTIMCODE <> "" Then Con4 = String.Format(" AND (TIMCODE = '{0}')", strTIMCODE)
                            If strTIMNAME <> "" Then Con5 = String.Format(" AND (DBO.SC_GET_CUSTNAME_FUN(TIMCODE) LIKE '%{0}%')", strTIMNAME)
                        Else
                            If strYEARMON <> "" Then Con1 = String.Format(" AND (SUBSTRING(DEMANDDAY,1,6) = '{0}')", strYEARMON)
                            If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (CLIENTCODE = '{0}')", strCLIENTCODE)
                            If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCLIENTNAME)
                            If strTIMCODE <> "" Then Con4 = String.Format(" AND (TIMCODE = '{0}')", strTIMCODE)
                            If strTIMNAME <> "" Then Con5 = String.Format(" AND (DBO.SC_GET_CUSTNAME_FUN(TIMCODE) LIKE '%{0}%')", strTIMNAME)
                        End If
                    ElseIf strFlag = "commi" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
                            If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (REAL_MED_CODE = '{0}')", strCLIENTCODE)
                            If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) LIKE '%{0}%')", strCLIENTNAME)
                        Else
                            If strYEARMON <> "" Then Con1 = String.Format(" AND (SUBSTRING(DEMANDDAY,1,6) = '{0}')", strYEARMON)
                            If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (REAL_MED_CODE = '{0}')", strCLIENTCODE)
                            If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) LIKE '%{0}%')", strCLIENTNAME)
                        End If
                    End If
                ElseIf strTBL_Flag = "OUTDOOR" Then
                    If strFlag = "commi" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
                            If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (CLIENTCODE = '{0}')", strCLIENTCODE)
                            If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCLIENTNAME)
                        Else
                            If strYEARMON <> "" Then Con1 = String.Format(" AND (SUBSTRING(DEMANDDAY,1,6) = '{0}')", strYEARMON)
                            If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (CLIENTCODE = '{0}')", strCLIENTCODE)
                            If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCLIENTNAME)
                        End If
                    End If
                ElseIf strTBL_Flag = "ELEC" Then
                    If strFlag = "trans" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
                            If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (CLIENTCODE = '{0}')", strCLIENTCODE)
                            If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCLIENTNAME)
                            If strTIMCODE <> "" Then Con4 = String.Format(" AND (TIMCODE = '{0}')", strTIMCODE)
                            If strTIMNAME <> "" Then Con5 = String.Format(" AND (DBO.SC_GET_CUSTNAME_FUN(TIMCODE) LIKE '%{0}%')", strTIMNAME)
                        Else
                            If strYEARMON <> "" Then Con1 = String.Format(" AND (YEARMON = '{0}')", strYEARMON)
                            If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (CLIENTCODE = '{0}')", strCLIENTCODE)
                            If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCLIENTNAME)
                            If strTIMCODE <> "" Then Con4 = String.Format(" AND (TIMCODE = '{0}')", strTIMCODE)
                            If strTIMNAME <> "" Then Con5 = String.Format(" AND (DBO.SC_GET_CUSTNAME_FUN(TIMCODE) LIKE '%{0}%')", strTIMNAME)
                        End If
                    ElseIf strFlag = "commi" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
                            If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (REAL_MED_CODE = '{0}')", strCLIENTCODE)
                            If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) LIKE '%{0}%')", strCLIENTNAME)
                        Else
                            If strYEARMON <> "" Then Con1 = String.Format(" AND (YEARMON = '{0}')", strYEARMON)
                            If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (REAL_MED_CODE = '{0}')", strCLIENTCODE)
                            If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) LIKE '%{0}%')", strCLIENTNAME)
                        End If
                    End If
                End If
                strWhere = BuildFields(" ", Con1, Con2, Con3, Con4, Con5)

                If strTBL_Flag = "PRINT" Then
                    If strFlag = "trans" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            strSelFields = "TRANSYEARMON , DBO.SC_GET_CUSTNAME_FUN(TIMCODE) TIMNAME, DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME, '완료' GBN, CLIENTCODE, TIMCODE"
                            strFormet = "select {0}  "
                            strFormet = strFormet & " FROM MD_PRINTTRANS_DTL  WHERE 1=1 {1} "
                            strFormet = strFormet & " GROUP BY CLIENTCODE, TIMCODE, TRANSYEARMON "
                            strFormet = strFormet & " ORDER BY DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE), DBO.SC_GET_CUSTNAME_FUN(TIMCODE)"
                        Else
                            strSelFields = "SUBSTRING(DEMANDDAY,1,6) YEARMON, DBO.SC_GET_CUSTNAME_FUN(TIMCODE) TIMNAME, DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME, '미완료' GBN, CLIENTCODE, TIMCODE"
                            strFormet = "select {0}    "
                            strFormet = strFormet & " FROM MD_BOOKING_MEDIUM  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND gflag in('B') "
                            strFormet = strFormet & " AND isnull(tru_trans_no, '') = ''  "
                            strFormet = strFormet & " GROUP BY CLIENTCODE , TIMCODE, SUBSTRING(DEMANDDAY,1,6) "
                            strFormet = strFormet & " ORDER BY "
                            strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE), DBO.SC_GET_CUSTNAME_FUN(TIMCODE) "
                        End If
                    ElseIf strFlag = "commi" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            strSelFields = "TRANSYEARMON , TRANSNO, DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) REAL_MED_NAME, '완료' GBN, REAL_MED_CODE"
                            strFormet = "select {0}  "
                            strFormet = strFormet & " FROM MD_PRINTCOMMI_DTL  WHERE 1=1 {1} "
                            strFormet = strFormet & " GROUP BY REAL_MED_CODE , TRANSYEARMON, TRANSNO "
                            strFormet = strFormet & " ORDER BY DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) "
                        Else
                            strSelFields = "SUBSTRING(DEMANDDAY,1,6) YEARMON , REAL_MED_CODE, DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) REAL_MED_NAME, '미완료' GBN"
                            strFormet = "select {0}    "
                            strFormet = strFormet & " FROM MD_BOOKING_MEDIUM  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND gflag in('B','J', 'S') "
                            strFormet = strFormet & " AND isnull(commi_trans_no, '') = ''  "
                            strFormet = strFormet & " GROUP BY REAL_MED_CODE , SUBSTRING(DEMANDDAY,1,6) "
                            strFormet = strFormet & " ORDER BY "
                            strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE)"
                        End If
                    End If
                ElseIf strTBL_Flag = "CATV" Then
                    If strFlag = "trans" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            strSelFields = "TRANSYEARMON , DBO.SC_GET_CUSTNAME_FUN(TIMCODE) TIMNAME, DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME, '완료' GBN, CLIENTCODE, TIMCODE"
                            strFormet = "select {0}  "
                            strFormet = strFormet & " FROM MD_CATVTRANS_DTL  WHERE 1=1 {1} "
                            strFormet = strFormet & " GROUP BY CLIENTCODE, TIMCODE, TRANSYEARMON "
                            strFormet = strFormet & " ORDER BY DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE), DBO.SC_GET_CUSTNAME_FUN(TIMCODE)"
                        Else
                            strSelFields = "SUBSTRING(DEMANDDAY,1,6) YEARMON, DBO.SC_GET_CUSTNAME_FUN(TIMCODE) TIMNAME, DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME, '미완료' GBN, CLIENTCODE, TIMCODE"
                            strFormet = "select {0}    "
                            strFormet = strFormet & " FROM MD_CATV_MEDIUM  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND gflag in('1') "
                            strFormet = strFormet & " AND isnull(tru_trans_no, '') = ''  "
                            strFormet = strFormet & " GROUP BY CLIENTCODE , TIMCODE, SUBSTRING(DEMANDDAY,1,6) "
                            strFormet = strFormet & " ORDER BY "
                            strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE), DBO.SC_GET_CUSTNAME_FUN(TIMCODE) "
                        End If
                    ElseIf strFlag = "commi" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            strSelFields = "TRANSYEARMON , TRANSNO, DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) REAL_MED_NAME, '완료' GBN, REAL_MED_CODE"
                            strFormet = "select {0}  "
                            strFormet = strFormet & " FROM MD_CATVCOMMI_DTL  WHERE 1=1 {1} "
                            strFormet = strFormet & " GROUP BY REAL_MED_CODE , TRANSYEARMON, TRANSNO "
                            strFormet = strFormet & " ORDER BY DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) "
                        Else
                            strSelFields = "SUBSTRING(DEMANDDAY,1,6) YEARMON , REAL_MED_CODE, DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) REAL_MED_NAME, '미완료' GBN"
                            strFormet = "select {0}    "
                            strFormet = strFormet & " FROM MD_CATV_MEDIUM  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND gflag in('1') "
                            strFormet = strFormet & " AND isnull(commi_trans_no, '') = ''  "
                            strFormet = strFormet & " GROUP BY REAL_MED_CODE , SUBSTRING(DEMANDDAY,1,6) "
                            strFormet = strFormet & " ORDER BY "
                            strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE)"
                        End If
                    End If
                ElseIf strTBL_Flag = "TOTAL" Then
                    If strFlag = "trans" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            strSelFields = "TRANSYEARMON , DBO.SC_GET_CUSTNAME_FUN(TIMCODE) TIMNAME, DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME, '완료' GBN, CLIENTCODE, TIMCODE"
                            strFormet = "select {0}  "
                            strFormet = strFormet & " FROM MD_TOTALTRANS_DTL  WHERE 1=1 {1} "
                            strFormet = strFormet & " GROUP BY CLIENTCODE, TIMCODE, TRANSYEARMON "
                            strFormet = strFormet & " ORDER BY DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE), DBO.SC_GET_CUSTNAME_FUN(TIMCODE)"
                        Else
                            strSelFields = "SUBSTRING(DEMANDDAY,1,6) YEARMON, DBO.SC_GET_CUSTNAME_FUN(TIMCODE) TIMNAME, DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME, '미완료' GBN, CLIENTCODE, TIMCODE"
                            strFormet = "select {0}    "
                            strFormet = strFormet & " FROM MD_TOTAL_MEDIUM  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND gflag in('1') "
                            strFormet = strFormet & " AND isnull(tru_trans_no, '') = ''  "
                            strFormet = strFormet & " GROUP BY CLIENTCODE , TIMCODE, SUBSTRING(DEMANDDAY,1,6) "
                            strFormet = strFormet & " ORDER BY "
                            strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE), DBO.SC_GET_CUSTNAME_FUN(TIMCODE) "
                        End If
                    ElseIf strFlag = "commi" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            strSelFields = "TRANSYEARMON , TRANSNO, DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) REAL_MED_NAME, '완료' GBN, REAL_MED_CODE"
                            strFormet = "select {0}  "
                            strFormet = strFormet & " FROM MD_TOTALCOMMI_DTL  WHERE 1=1 {1} "
                            strFormet = strFormet & " GROUP BY REAL_MED_CODE , TRANSYEARMON, TRANSNO "
                            strFormet = strFormet & " ORDER BY DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) "
                        Else
                            strSelFields = "SUBSTRING(DEMANDDAY,1,6) YEARMON , REAL_MED_CODE, DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) REAL_MED_NAME, '미완료' GBN"
                            strFormet = "select {0}    "
                            strFormet = strFormet & " FROM MD_TOTAL_MEDIUM  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND gflag in('1') "
                            strFormet = strFormet & " AND isnull(commi_trans_no, '') = ''  "
                            strFormet = strFormet & " GROUP BY REAL_MED_CODE , SUBSTRING(DEMANDDAY,1,6) "
                            strFormet = strFormet & " ORDER BY "
                            strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE)"
                        End If
                    End If
                ElseIf strTBL_Flag = "INTERNET" Then
                    If strFlag = "trans" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            strSelFields = "TRANSYEARMON , DBO.SC_GET_CUSTNAME_FUN(TIMCODE) TIMNAME, DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME, '완료' GBN, CLIENTCODE, TIMCODE"
                            strFormet = "select {0}  "
                            strFormet = strFormet & " FROM MD_INTERNETTRANS_DTL  WHERE 1=1 {1} "
                            strFormet = strFormet & " GROUP BY CLIENTCODE, TIMCODE, TRANSYEARMON "
                            strFormet = strFormet & " ORDER BY DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE), DBO.SC_GET_CUSTNAME_FUN(TIMCODE)"
                        Else
                            strSelFields = "SUBSTRING(DEMANDDAY,1,6) YEARMON, DBO.SC_GET_CUSTNAME_FUN(TIMCODE) TIMNAME, DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME, '미완료' GBN, CLIENTCODE, TIMCODE"
                            strFormet = "select {0}    "
                            strFormet = strFormet & " FROM MD_INTERNET_MEDIUM  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND gflag in('1') "
                            strFormet = strFormet & " AND isnull(tru_trans_no, '') = ''  "
                            strFormet = strFormet & " GROUP BY CLIENTCODE , TIMCODE, SUBSTRING(DEMANDDAY,1,6) "
                            strFormet = strFormet & " ORDER BY "
                            strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE), DBO.SC_GET_CUSTNAME_FUN(TIMCODE) "
                        End If
                    ElseIf strFlag = "commi" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            strSelFields = "TRANSYEARMON , TRANSNO, DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) REAL_MED_NAME, '완료' GBN, REAL_MED_CODE"
                            strFormet = "select {0}  "
                            strFormet = strFormet & " FROM MD_INTERNETCOMMI_DTL  WHERE 1=1 {1} "
                            strFormet = strFormet & " GROUP BY REAL_MED_CODE , TRANSYEARMON, TRANSNO "
                            strFormet = strFormet & " ORDER BY DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) "
                        Else
                            strSelFields = "SUBSTRING(DEMANDDAY,1,6) YEARMON , REAL_MED_CODE, DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) REAL_MED_NAME, '미완료' GBN"
                            strFormet = "select {0}    "
                            strFormet = strFormet & " FROM MD_INTERNET_MEDIUM  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND gflag in('1') "
                            strFormet = strFormet & " AND isnull(commi_trans_no, '') = ''  "
                            strFormet = strFormet & " GROUP BY REAL_MED_CODE , SUBSTRING(DEMANDDAY,1,6) "
                            strFormet = strFormet & " ORDER BY "
                            strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE)"
                        End If
                    End If
                ElseIf strTBL_Flag = "OUTDOOR" Then
                    If strFlag = "commi" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            strSelFields = "TRANSYEARMON , TRANSNO, DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) REAL_MED_NAME, '완료' GBN, CLIENTCODE REAL_MED_CODE"
                            strFormet = "select {0}  "
                            strFormet = strFormet & " FROM MD_OUTDOORCOMMI_DTL  WHERE 1=1 {1} "
                            strFormet = strFormet & " GROUP BY CLIENTCODE , TRANSYEARMON, TRANSNO "
                            strFormet = strFormet & " ORDER BY DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) "
                        Else
                            strSelFields = "SUBSTRING(DEMANDDAY,1,6) YEARMON , CLIENTCODE REAL_MED_CODE, DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) REAL_MED_NAME, '미완료' GBN"
                            strFormet = "select {0}    "
                            strFormet = strFormet & " FROM MD_OUTDOOR_MEDIUM  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND isnull(commi_trans_no, '') = ''  "
                            strFormet = strFormet & " GROUP BY CLIENTCODE , SUBSTRING(DEMANDDAY,1,6) "
                            strFormet = strFormet & " ORDER BY "
                            strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE)"
                        End If
                    End If
                ElseIf strTBL_Flag = "ELEC" Then
                    If strFlag = "trans" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            strSelFields = "TRANSYEARMON , DBO.SC_GET_CUSTNAME_FUN(TIMCODE) TIMNAME, DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME, '완료' GBN, CLIENTCODE, TIMCODE"
                            strFormet = "select {0}  "
                            strFormet = strFormet & " FROM MD_ELEC_TRANS_DTL  WHERE 1=1 {1} "
                            strFormet = strFormet & " GROUP BY CLIENTCODE, TIMCODE, TRANSYEARMON "
                            strFormet = strFormet & " ORDER BY DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE), DBO.SC_GET_CUSTNAME_FUN(TIMCODE)"
                        Else
                            strSelFields = " YEARMON, DBO.SC_GET_CUSTNAME_FUN(TIMCODE) TIMNAME, DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME, '미완료' GBN, CLIENTCODE, TIMCODE"
                            strFormet = "select {0}    "
                            strFormet = strFormet & " FROM MD_ELECTRIC_MEDIUM  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND isnull(tru_trans_no, '') = ''  "
                            strFormet = strFormet & " GROUP BY CLIENTCODE , TIMCODE, YEARMON "
                            strFormet = strFormet & " ORDER BY "
                            strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE), DBO.SC_GET_CUSTNAME_FUN(TIMCODE) "
                        End If
                    ElseIf strFlag = "commi" Then
                        If strCOMMITCHECK = "COMMIT" Then
                            strSelFields = "TRANSYEARMON , TRANSNO, DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) REAL_MED_NAME, '완료' GBN, REAL_MED_CODE"
                            strFormet = "select {0}  "
                            strFormet = strFormet & " FROM MD_ELECCOMMI_DTL  WHERE 1=1 {1} "
                            strFormet = strFormet & " GROUP BY REAL_MED_CODE , TRANSYEARMON, TRANSNO "
                            strFormet = strFormet & " ORDER BY DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) "
                        Else
                            strSelFields = "YEARMON , REAL_MED_CODE, DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) REAL_MED_NAME, '미완료' GBN"
                            strFormet = "select {0}    "
                            strFormet = strFormet & " FROM MD_ELECTRIC_MEDIUM  WHERE 1=1 {1} "
                            strFormet = strFormet & " AND isnull(commi_trans_no, '') = ''  "
                            strFormet = strFormet & " GROUP BY REAL_MED_CODE , YEARMON "
                            strFormet = strFormet & " ORDER BY "
                            strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE)"
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

#Region "15. 세금계산서 생성광고주 POP조회"
    Public Function GetTAXCUSTNO(ByVal strInfoXML As String, _
                                 ByRef intRowCnt As Integer, _
                                 ByRef intColCnt As Integer, _
                                 ByVal strYEARMON As String, _
                                 ByVal strCLIENTCODE As String, _
                                 ByVal strCLIENTNAME As String, _
                                 ByVal strTBL_Flag As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2, Con3 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '기본정보 설정

                Con1 = ""
                Con2 = ""
                Con3 = ""

                If strTBL_Flag = "PRINT" Or strTBL_Flag = "CATV" Or strTBL_Flag = "INTERNET" Or strTBL_Flag = "OUTDOOR" Or strTBL_Flag = "CGV" Or strTBL_Flag = "TOTAL" Then
                    If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
                    If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (CLIENTCODE LIKE '%{0}%')", strCLIENTCODE)
                    If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCLIENTNAME)
                End If
                strWhere = BuildFields(" ", Con1, Con2, Con3)


                If strTBL_Flag = "PRINT" Then
                    strFormet = strFormet & " select "
                    strFormet = strFormet & " TRANSYEARMON,"
                    strFormet = strFormet & " CLIENTCODE,"
                    strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME,"
                    strFormet = strFormet & " DBO.SC_GET_BUSINO_FUN(CLIENTCODE) BUSINO"
                    strFormet = strFormet & " FROM MD_PRINTTRANS_DTL"
                    strFormet = strFormet & " WHERE 1=1 {0} "
                    strFormet = strFormet & " GROUP BY TRANSYEARMON, CLIENTCODE"
                    strFormet = strFormet & " ORDER BY  DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE)"
                ElseIf strTBL_Flag = "CATV" Then
                    strFormet = strFormet & " select "
                    strFormet = strFormet & " TRANSYEARMON,"
                    strFormet = strFormet & " CLIENTCODE,"
                    strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME,"
                    strFormet = strFormet & " DBO.SC_GET_BUSINO_FUN(CLIENTCODE) BUSINO"
                    strFormet = strFormet & " FROM MD_CATVTRANS_DTL"
                    strFormet = strFormet & " WHERE 1=1 {0} "
                    strFormet = strFormet & " GROUP BY TRANSYEARMON, CLIENTCODE"
                    strFormet = strFormet & " ORDER BY  DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE)"
                ElseIf strTBL_Flag = "TOTAL" Then
                    strFormet = strFormet & " select "
                    strFormet = strFormet & " TRANSYEARMON,"
                    strFormet = strFormet & " CLIENTCODE,"
                    strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME,"
                    strFormet = strFormet & " DBO.SC_GET_BUSINO_FUN(CLIENTCODE) BUSINO"
                    strFormet = strFormet & " FROM MD_TOTALTRANS_DTL"
                    strFormet = strFormet & " WHERE 1=1 {0} "
                    strFormet = strFormet & " GROUP BY TRANSYEARMON, CLIENTCODE"
                    strFormet = strFormet & " ORDER BY  DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE)"
                ElseIf strTBL_Flag = "INTERNET" Then
                    strFormet = strFormet & " select "
                    strFormet = strFormet & " TRANSYEARMON,"
                    strFormet = strFormet & " CLIENTCODE,"
                    strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME,"
                    strFormet = strFormet & " DBO.SC_GET_BUSINO_FUN(CLIENTCODE) BUSINO"
                    strFormet = strFormet & " FROM MD_INTERNETTRANS_DTL"
                    strFormet = strFormet & " WHERE 1=1 {0} "
                    strFormet = strFormet & " GROUP BY TRANSYEARMON, CLIENTCODE"
                    strFormet = strFormet & " ORDER BY  DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE)"
                ElseIf strTBL_Flag = "OUTDOOR" Then
                    strFormet = strFormet & " select "
                    strFormet = strFormet & " TRANSYEARMON,"
                    strFormet = strFormet & " CLIENTCODE,"
                    strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME,"
                    strFormet = strFormet & " DBO.SC_GET_BUSINO_FUN(CLIENTCODE) BUSINO"
                    strFormet = strFormet & " FROM MD_OUTDOORCOMMI_DTL"
                    strFormet = strFormet & " WHERE 1=1 {0} "
                    strFormet = strFormet & " GROUP BY TRANSYEARMON, CLIENTCODE"
                    strFormet = strFormet & " ORDER BY  DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE)"
                ElseIf strTBL_Flag = "CGV" Then
                    strFormet = strFormet & " select "
                    strFormet = strFormet & " TRANSYEARMON,"
                    strFormet = strFormet & " CLIENTCODE,"
                    strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME,"
                    strFormet = strFormet & " DBO.SC_GET_BUSINO_FUN(CLIENTCODE) BUSINO"
                    strFormet = strFormet & " FROM MD_CLOUDCOMMI_DTL"
                    strFormet = strFormet & " WHERE 1=1 {0} "
                    strFormet = strFormet & " GROUP BY TRANSYEARMON, CLIENTCODE"
                    strFormet = strFormet & " ORDER BY  DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE)"
                End If

                strSQL = String.Format(strFormet, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetTRNASCUSTNO")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "16. 세금계산서 생성 팀 POP조회"
    Public Function GetTAXTIMNO(ByVal strInfoXML As String, _
                                ByRef intRowCnt As Integer, _
                                ByRef intColCnt As Integer, _
                                ByVal strYEARMON As String, _
                                ByVal strCLIENTCODE As String, _
                                ByVal strCLIENTNAME As String, _
                                ByVal strTIMCODE As String, _
                                ByVal strTIMNAME As String, _
                                ByVal strTBL_Flag As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2, Con3, Con4, Con5 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '기본정보 설정

                Con1 = "" : Con2 = "" : Con3 = "" : Con4 = "" : Con5 = ""

                If strTBL_Flag = "PRINT" Or strTBL_Flag = "CATV" Or strTBL_Flag = "INTERNET" Or strTBL_Flag = "INTERNET" Or strTBL_Flag = "OUTDOOR" Or strTBL_Flag = "TOTAL" Then
                    If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
                    If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (CLIENTCODE LIKE '%{0}%')", strCLIENTCODE)
                    If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCLIENTNAME)
                    If strTIMCODE <> "" Then Con4 = String.Format(" AND (TIMCODE '{0}')", strTIMCODE)
                    If strTIMNAME <> "" Then Con5 = String.Format(" AND (DBO.SC_GET_CUSTNAME_FUN(TIMCODE) LIKE '%{0}%')", strTIMNAME)
                End If
                strWhere = BuildFields(" ", Con1, Con2, Con3, Con4, Con5)


                If strTBL_Flag = "PRINT" Then
                    strFormet = strFormet & " select "
                    strFormet = strFormet & " TRANSYEARMON,"
                    strFormet = strFormet & " DBO.SC_GET_CUSTNAME_FUN(TIMCODE) TIMNAME,"
                    strFormet = strFormet & " TIMCODE,"
                    strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME,"
                    strFormet = strFormet & " CLIENTCODE"
                    strFormet = strFormet & " FROM MD_PRINTTRANS_DTL"
                    strFormet = strFormet & " WHERE 1=1 {0} "
                    strFormet = strFormet & " GROUP BY TRANSYEARMON, CLIENTCODE, TIMCODE"
                    strFormet = strFormet & " ORDER BY DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE),"
                    strFormet = strFormet & " DBO.SC_GET_CUSTNAME_FUN(TIMCODE)"
                ElseIf strTBL_Flag = "CATV" Then
                    strFormet = strFormet & " select "
                    strFormet = strFormet & " TRANSYEARMON,"
                    strFormet = strFormet & " DBO.SC_GET_CUSTNAME_FUN(TIMCODE) TIMNAME,"
                    strFormet = strFormet & " TIMCODE,"
                    strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME,"
                    strFormet = strFormet & " CLIENTCODE"
                    strFormet = strFormet & " FROM MD_CATVTRANS_DTL"
                    strFormet = strFormet & " WHERE 1=1 {0} "
                    strFormet = strFormet & " GROUP BY TRANSYEARMON, CLIENTCODE, TIMCODE"
                    strFormet = strFormet & " ORDER BY DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE),"
                    strFormet = strFormet & " DBO.SC_GET_CUSTNAME_FUN(TIMCODE)"
                ElseIf strTBL_Flag = "TOTAL" Then
                    strFormet = strFormet & " select "
                    strFormet = strFormet & " TRANSYEARMON,"
                    strFormet = strFormet & " DBO.SC_GET_CUSTNAME_FUN(TIMCODE) TIMNAME,"
                    strFormet = strFormet & " TIMCODE,"
                    strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME,"
                    strFormet = strFormet & " CLIENTCODE"
                    strFormet = strFormet & " FROM MD_TOTALTRANS_DTL"
                    strFormet = strFormet & " WHERE 1=1 {0} "
                    strFormet = strFormet & " GROUP BY TRANSYEARMON, CLIENTCODE, TIMCODE"
                    strFormet = strFormet & " ORDER BY DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE),"
                    strFormet = strFormet & " DBO.SC_GET_CUSTNAME_FUN(TIMCODE)"
                ElseIf strTBL_Flag = "INTERNET" Then
                    strFormet = strFormet & " select "
                    strFormet = strFormet & " TRANSYEARMON,"
                    strFormet = strFormet & " DBO.SC_GET_CUSTNAME_FUN(TIMCODE) TIMNAME,"
                    strFormet = strFormet & " TIMCODE,"
                    strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME,"
                    strFormet = strFormet & " CLIENTCODE"
                    strFormet = strFormet & " FROM MD_INTERNETTRANS_DTL"
                    strFormet = strFormet & " WHERE 1=1 {0} "
                    strFormet = strFormet & " GROUP BY TRANSYEARMON, CLIENTCODE, TIMCODE"
                    strFormet = strFormet & " ORDER BY DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE),"
                    strFormet = strFormet & " DBO.SC_GET_CUSTNAME_FUN(TIMCODE)"
                ElseIf strTBL_Flag = "OUTDOOR" Then
                    strFormet = strFormet & " select "
                    strFormet = strFormet & " TRANSYEARMON,"
                    strFormet = strFormet & " DBO.SC_GET_CUSTNAME_FUN(TIMCODE) TIMNAME,"
                    strFormet = strFormet & " TIMCODE,"
                    strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME,"
                    strFormet = strFormet & " CLIENTCODE"
                    strFormet = strFormet & " FROM MD_OUTDOORCOMMI_DTL"
                    strFormet = strFormet & " WHERE 1=1 {0} "
                    strFormet = strFormet & " GROUP BY TRANSYEARMON, CLIENTCODE, TIMCODE"
                    strFormet = strFormet & " ORDER BY DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE),"
                    strFormet = strFormet & " DBO.SC_GET_CUSTNAME_FUN(TIMCODE)"
                End If

                strSQL = String.Format(strFormet, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetTRNASCUSTNO")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "17. 전표처래내역 조회"
    Public Function TRUVOCHNO_CHECKED(ByVal strInfoXML As String, _
                                      ByRef intRowCnt As Integer, _
                                      ByRef intColCnt As Integer, _
                                      ByVal strTAXYEARMON As String, _
                                      ByVal strTAXNO As String) As Object

        Dim strSQL As String            'SQL문
        Dim strFormat As String         '임시 SQL문
        Dim strSelFields As String      '조회필드
        Dim strWhere As String
        Dim vntData As Object
        Dim Con1 As String
        Dim Con2 As String
        Dim Con3 As String

        SetConfig(strInfoXML)   '기본정보 설정
        With mobjSCGLConfig

            Con1 = ""
            Con2 = ""
            Con3 = ""
            If strTAXYEARMON <> "" Then Con1 = String.Format(" AND (TAXYEARMON = '{0}')", strTAXYEARMON)
            If strTAXNO <> "" Then Con2 = String.Format(" AND (TAXNO = '{0}')", strTAXNO)
            Con3 = " AND (RIGHT(RMSNO,1) = 'T')"

            strWhere = BuildFields(" ", Con1, Con2, Con3)
            strFormat = "SELECT TAXYEARMON, TAXNO, RMSNO FROM MD_TRUVOCH_MST WHERE 1=1 {0} "
            strSQL = String.Format(strFormat, strWhere)
            '데이터 조회
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


    Public Function COMMIVOCHNO_CHECKED(ByVal strInfoXML As String, _
                                        ByRef intRowCnt As Integer, _
                                        ByRef intColCnt As Integer, _
                                        ByVal strTAXYEARMON As String, _
                                        ByVal strTAXNO As String) As Object

        Dim strSQL As String            'SQL문
        Dim strFormat As String         '임시 SQL문
        Dim strSelFields As String      '조회필드
        Dim strWhere As String
        Dim vntData As Object
        Dim Con1 As String
        Dim Con2 As String
        Dim Con3 As String


        SetConfig(strInfoXML)   '기본정보 설정
        With mobjSCGLConfig

            Con1 = ""
            Con2 = ""
            Con3 = ""
            If strTAXYEARMON <> "" Then Con1 = String.Format(" AND (TAXYEARMON = '{0}')", strTAXYEARMON)
            If strTAXNO <> "" Then Con2 = String.Format(" AND (TAXNO = '{0}')", strTAXNO)
            Con3 = " AND (RIGHT(RMSNO,1) = 'S')"

            strWhere = BuildFields(" ", Con1, Con2, Con3)
            strFormat = "SELECT TAXYEARMON, TAXNO, RMSNO FROM MD_TRUVOCH_MST WHERE 1=1 {0} "
            strSQL = String.Format(strFormat, strWhere)
            '데이터 조회
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

    Public Function COMMIVOCHNO_CHECKED_MERGE(ByVal strInfoXML As String, _
                                              ByRef intRowCnt As Integer, _
                                              ByRef intColCnt As Integer, _
                                              ByVal strTAXYEARMON As String, _
                                              ByVal strTAXNO As String, _
                                              ByVal strMEDFLAG As String) As Object

        Dim strSQL As String            'SQL문
        Dim strFormat As String         '임시 SQL문
        Dim strSelFields As String      '조회필드
        Dim strWhere As String
        Dim vntData As Object
        Dim Con1 As String
        Dim Con2 As String
        Dim Con3 As String

        SetConfig(strInfoXML)   '기본정보 설정
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
            '데이터 조회
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

#Region "18. 세금계산서 생성매체사 POP조회"
    Public Function GetTAXREAL_MED_NO(ByVal strInfoXML As String, _
                                      ByRef intRowCnt As Integer, _
                                      ByRef intColCnt As Integer, _
                                      ByVal strYEARMON As String, _
                                      ByVal strREAL_MED_CODE As String, _
                                      ByVal strREAL_MED_NAME As String, _
                                      ByVal strTBL_Flag As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2, Con3 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '기본정보 설정

                Con1 = ""
                Con2 = ""
                Con3 = ""

                If strTBL_Flag = "PRINT" Or strTBL_Flag = "CATV" Or strTBL_Flag = "INTERNET" Or strTBL_Flag = "TOTAL" Then
                    If strYEARMON <> "" Then Con1 = String.Format(" AND (TRANSYEARMON = '{0}')", strYEARMON)
                    If strREAL_MED_CODE <> "" Then Con2 = String.Format(" AND (REAL_MED_CODE LIKE '%{0}%')", strREAL_MED_CODE)
                    If strREAL_MED_NAME <> "" Then Con3 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) LIKE '%{0}%')", strREAL_MED_NAME)
                End If

                strWhere = BuildFields(" ", Con1, Con2, Con3)

                If strTBL_Flag = "PRINT" Then
                    strFormet = strFormet & " select "
                    strFormet = strFormet & " TRANSYEARMON,"
                    strFormet = strFormet & " REAL_MED_CODE,"
                    strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) REAL_MED_NAME,"
                    strFormet = strFormet & " DBO.SC_GET_BUSINO_FUN(REAL_MED_CODE) BUSINO"
                    strFormet = strFormet & " FROM MD_PRINTCOMMI_DTL"
                    strFormet = strFormet & " WHERE 1=1 {0} "
                    strFormet = strFormet & " GROUP BY TRANSYEARMON, REAL_MED_CODE"
                    strFormet = strFormet & " ORDER BY  DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE)"
                ElseIf strTBL_Flag = "CATV" Then
                    strFormet = strFormet & " select "
                    strFormet = strFormet & " TRANSYEARMON,"
                    strFormet = strFormet & " REAL_MED_CODE,"
                    strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) REAL_MED_NAME,"
                    strFormet = strFormet & " DBO.SC_GET_BUSINO_FUN(REAL_MED_CODE) BUSINO"
                    strFormet = strFormet & " FROM MD_CATVCOMMI_DTL"
                    strFormet = strFormet & " WHERE 1=1 {0} "
                    strFormet = strFormet & " GROUP BY TRANSYEARMON, REAL_MED_CODE"
                    strFormet = strFormet & " ORDER BY  DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE)"
                ElseIf strTBL_Flag = "TOTAL" Then
                    strFormet = strFormet & " select "
                    strFormet = strFormet & " TRANSYEARMON,"
                    strFormet = strFormet & " REAL_MED_CODE,"
                    strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) REAL_MED_NAME,"
                    strFormet = strFormet & " DBO.SC_GET_BUSINO_FUN(REAL_MED_CODE) BUSINO"
                    strFormet = strFormet & " FROM MD_TOTALCOMMI_DTL"
                    strFormet = strFormet & " WHERE 1=1 {0} "
                    strFormet = strFormet & " GROUP BY TRANSYEARMON, REAL_MED_CODE"
                    strFormet = strFormet & " ORDER BY  DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE)"
                ElseIf strTBL_Flag = "INTERNET" Then
                    strFormet = strFormet & " select "
                    strFormet = strFormet & " TRANSYEARMON,"
                    strFormet = strFormet & " REAL_MED_CODE,"
                    strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) REAL_MED_NAME,"
                    strFormet = strFormet & " DBO.SC_GET_BUSINO_FUN(REAL_MED_CODE) BUSINO"
                    strFormet = strFormet & " FROM MD_INTERNETCOMMI_DTL"
                    strFormet = strFormet & " WHERE 1=1 {0} "
                    strFormet = strFormet & " GROUP BY TRANSYEARMON, REAL_MED_CODE"
                    strFormet = strFormet & " ORDER BY  DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE)"
                End If

                strSQL = String.Format(strFormet, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetTRNASCUSTNO")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "19. 광고처조회"
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

#Region "20. 캠페인 조회"
    Public Function GetCAMPAIGN_INFO(ByVal strInfoXML As String, _
                                     ByRef intRowCnt As Integer, _
                                     ByRef intColCnt As Integer, _
                                     ByVal strYEARMON As String, _
                                     ByVal strCAMPAIGN_CODE As String, _
                                     ByVal strCAMPAIGN_NAME As String) As Object

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2, Con3 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '기본정보 설정

                Con1 = ""
                Con2 = ""
                Con3 = ""

                If strYEARMON <> "" Then Con1 = String.Format(" AND {0} BETWEEN SUBSTRING(TBRDSTDATE,1,6) AND SUBSTRING(TBRDEDDATE,1,6) ", strYEARMON)
                If strCAMPAIGN_CODE <> "" Then Con2 = String.Format(" AND (CAMPAIGN_CODE = '{0}')", strCAMPAIGN_CODE)
                If strCAMPAIGN_NAME <> "" Then
                    If Mid(strCAMPAIGN_NAME, 1, 1) = "[" Then
                        strCAMPAIGN_NAME = Replace(strCAMPAIGN_NAME, "[", "[[]")
                    End If
                    Con3 = String.Format(" AND (CAMPAIGN_NAME like '%{0}%')", strCAMPAIGN_NAME)
                End If
                strWhere = BuildFields(" ", Con1, Con2, Con3)

                strSelFields = "'" & strYEARMON & "' YEARMON, "
                strSelFields = strSelFields & " CAMPAIGN_CODE, "
                strSelFields = strSelFields & " CAMPAIGN_NAME, "
                strSelFields = strSelFields & " CLIENTCODE, "
                strSelFields = strSelFields & " DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME, "
                strSelFields = strSelFields & " SUBSEQ, "
                strSelFields = strSelFields & " DBO.SC_GET_SUBSEQNAME_FUN(SUBSEQ) SUBSEQNAME, "
                strSelFields = strSelFields & " TIMCODE, "
                strSelFields = strSelFields & " DBO.SC_GET_CUSTNAME_FUN(TIMCODE) TIMNAME, "
                strSelFields = strSelFields & " DEPT_CD, "
                strSelFields = strSelFields & " DBO.SC_DEPT_NAME_FUN(DEPT_CD) DEPT_NAME, "
                strSelFields = strSelFields & " TBRDSTDATE, "
                strSelFields = strSelFields & " TBRDEDDATE, "
                strSelFields = strSelFields & " EXCLIENTCODE, "
                strSelFields = strSelFields & " DBO.SC_GET_HIGHCUSTNAME_FUN(EXCLIENTCODE) EXCLIENTNAME, "
                strSelFields = strSelFields & " MEMO "

                strFormet = "select {0} from MD_INTERNET_CAMPAIGN where 1=1 {1} "


                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetCAMPAIGN_INFO")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "21. MPP 조회"

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
                SetConfig(strInfoXML)   '기본정보 설정

                Con1 = ""
                Con2 = ""

                If strCUSTCODE <> "" Then Con1 = String.Format(" AND (CUSTCODE LIKE '%{0}%')", strCUSTCODE)
                If strCUSTNAME <> "" Then Con2 = String.Format(" AND (CUSTNAME LIKE '%{0}%')", strCUSTNAME)
                strWhere = BuildFields(" ", Con1, Con2)

                strSelFields = "HIGHCUSTCODE , CUSTNAME , BUSINO, COMPANYNAME"

                strFormet = "SELECT {0} FROM SC_CUST_HDR WHERE 1=1 AND HIGHCUSTCODE LIKE 'P%'  AND ISNULL(USE_FLAG,'') = '1' {1} ORDER BY  CASE SUBSTRING(LTRIM(CUSTNAME),1,3) WHEN '(주)' THEN LTRIM(SUBSTRING(CUSTNAME,4,LEN(CUSTNAME))) ELSE LTRIM(CUSTNAME) END"


                strSQL = String.Format(strFormet, strSelFields, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetMPP")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "22. NEW_매체명 조회"

    ' =============== SelectRtnSample Code
    Public Function GetMEDNO(ByVal strInfoXML As String, _
                             ByRef intRowCnt As Integer, _
                             ByRef intColCnt As Integer, _
                             ByVal strCUSTCODE As String, _
                             ByVal strCUSTNAME As String) As Object

        Dim strSQL As String
        Dim strFormat, strSelFields, strWhere As String

        Dim strChkDate As String = ""
        Dim Con1, Con2 As String
        Dim vntData As Object


        SetConfig(strInfoXML)   '기본정보 설정

        Con1 = ""
        Con2 = ""

        If strCUSTCODE <> "" Then Con1 = String.Format(" AND (CUSTCODE LIKE '%{0}%')", strCUSTCODE)
        If strCUSTNAME <> "" Then Con2 = String.Format(" AND (CUSTNAME LIKE '%{0}%')", strCUSTNAME)

        strWhere = BuildFields(" ", Con1, Con2)

        strFormat = "  SELECT "
        strFormat = strFormat & "  CUSTCODE,"
        strFormat = strFormat & "  CUSTNAME,"
        strFormat = strFormat & "  DBO.SC_GET_BUSINO_FUN(HIGHCUSTCODE) BUSINO,"
        strFormat = strFormat & "  HIGHCUSTCODE,"
        strFormat = strFormat & "  DBO.SC_GET_HIGHCUSTNAME_FUN(HIGHCUSTCODE) HIGHCUSTNAME, "
        strFormat = strFormat & "  MPP, DBO.SC_GET_HIGHCUSTNAME_FUN(MPP) MPP_NAME "
        strFormat = strFormat & "  FROM SC_CUST_DTL"
        strFormat = strFormat & "  WHERE 1=1 AND MEDFLAG = 'B' AND USE_FLAG = '1' {0} ORDER BY CUSTNAME"

        strSQL = String.Format(strFormat, strWhere)

        With mobjSCGLConfig
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)

                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , False)
                Return vntData

            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".GetMEDNO")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "23. 매체구분조회"
    ' =============== SelectRtnSample Code
    Public Function GetMEDGUBNCODE(ByVal strInfoXML As String, _
                                   ByRef intRowCnt As Integer, _
                                   ByRef intColCnt As Integer, _
                                   ByVal strREAL_MED_CODE As String, _
                                   ByVal strREAL_MED_NAME As String, _
                                   ByVal strMEDCODE As String, _
                                   ByVal strMEDNAME As String, _
                                   ByVal strMEDFLAG As String) As Object

        Dim strCols As String        '컬럼변수
        Dim strWhere As String       'Where조건 변수
        Dim strFormat As String      'SQL Format 변수
        Dim strSQL As String         'SQL 변수
        Dim vntData As Object        'Array Return 변수(Array 를사용할 때 선언)
        Dim strSelFields
        Dim intCnt
        Dim intRtn
        Dim Con1, Con2, Con3, Con4, Con5
        SetConfig(strInfoXML)

        Con1 = "" : Con2 = "" : Con3 = "" : Con4 = "" : Con5 = ""

        If strREAL_MED_CODE <> "" Then Con1 = String.Format(" AND (HIGHCUSTCODE = '{0}')", strREAL_MED_CODE)
        If strREAL_MED_NAME <> "" Then Con2 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(HIGHCUSTCODE) LIKE '%{0}%')", strREAL_MED_NAME)
        If strMEDCODE <> "" Then Con3 = String.Format(" AND (CUSTCODE = '{0}')", strMEDCODE)
        If strMEDNAME <> "" Then Con4 = String.Format(" AND (CUSTNAME LIKE '%{0}%')", strMEDNAME)

        '보내올때 컬럼이름을 플래그로 넘긴다.    A2라고할시에는 IF문으로 다 걸러줘야 하고  어차피 1인것을 가져올것이기 때문에 ...
        'If strMEDFLAG <> "" Then Con5 = String.Format(" AND (" & strMEDFLAG & " = '1')")

        If strMEDFLAG <> "" Then
            If strMEDFLAG = "MED_PRINT" Then
                Con5 = String.Format(" AND (MED_PAP = '1' OR MED_MAG = '1')")
            ElseIf strMEDFLAG = "MED_ELECTRIC" Then
                Con5 = String.Format(" AND (MED_TV = '1' OR MED_RD = '1' OR MED_DMB = '1' )")
            Else
                Con5 = String.Format(" AND (" & strMEDFLAG & " = '1')")
            End If
        End If

        strWhere = BuildFields(" ", Con1, Con2, Con3, Con4, Con5)

        strFormat = "  SELECT "
        strFormat = strFormat & "  CUSTCODE,"
        strFormat = strFormat & "  CUSTNAME,"
        strFormat = strFormat & "  dbo.SC_GET_BUSINO_FUN(HIGHCUSTCODE) BUSINO,"
        strFormat = strFormat & "  HIGHCUSTCODE,"
        strFormat = strFormat & "  DBO.SC_GET_HIGHCUSTNAME_FUN(HIGHCUSTCODE) COMPANYNAME, "
        strFormat = strFormat & "  MPP, DBO.SC_GET_HIGHCUSTNAME_FUN(MPP) MPPNAME "
        strFormat = strFormat & "  FROM SC_CUST_DTL"
        strFormat = strFormat & "  WHERE 1=1 AND MEDFLAG = 'B' AND USE_FLAG = '1' {0} ORDER BY CUSTNAME"

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

#Region "24. 대대행사 조회"

    ' =============== SelectRtnSample Code
    Public Function GetEXCUSTNO(ByVal strInfoXML As String, _
                                ByRef intRowCnt As Integer, _
                                ByRef intColCnt As Integer, _
                                ByVal strCUSTCODE As String, _
                                ByVal strCUSTNAME As String) As Object

        Dim strSQL As String
        Dim strFormet, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '기본정보 설정

                Con1 = ""
                Con2 = ""

                If strCUSTCODE <> "" Then Con1 = String.Format(" AND (HIGHCUSTCODE LIKE '%{0}%')", strCUSTCODE)
                If strCUSTNAME <> "" Then Con2 = String.Format(" AND (CUSTNAME LIKE '%{0}%')", strCUSTNAME)
                strWhere = BuildFields(" ", Con1, Con2)


                strFormet = " SELECT "
                strFormet = strFormet & " HIGHCUSTCODE CUSTCODE, CUSTNAME, BUSINO, COMPANYNAME "
                strFormet = strFormet & " FROM SC_CUST_HDR WHERE 1=1 AND HIGHCUSTCODE LIKE 'G%' AND USE_FLAG =1 "
                strFormet = strFormet & " {0} "
                strFormet = strFormet & " ORDER BY  CASE SUBSTRING(LTRIM(CUSTNAME),1,3) WHEN '(주)' THEN LTRIM(SUBSTRING(CUSTNAME,4,LEN(CUSTNAME))) "
                strFormet = strFormet & " WHEN '(사)' THEN LTRIM(SUBSTRING(CUSTNAME,4,LEN(CUSTNAME))) "
                strFormet = strFormet & " WHEN '(유)' THEN LTRIM(SUBSTRING(CUSTNAME,4,LEN(CUSTNAME))) "
                strFormet = strFormet & " WHEN '(재)' THEN LTRIM(SUBSTRING(CUSTNAME,4,LEN(CUSTNAME))) "
                strFormet = strFormet & " WHEN '(사단' THEN LTRIM(SUBSTRING(CUSTNAME,5,LEN(CUSTNAME))) "
                strFormet = strFormet & " WHEN '(학)' THEN LTRIM(SUBSTRING(CUSTNAME,4,LEN(CUSTNAME))) "
                strFormet = strFormet & " ELSE LTRIM(CUSTNAME) END "


                strSQL = String.Format(strFormet, strWhere)

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

#Region "25. 거래처담당자조회"
    ' =============== SelectRtnSample Code
    Public Function GetSC_CUST_EMP(ByVal strInfoXML As String, _
                                   ByRef intRowCnt As Integer, _
                                   ByRef intColCnt As Integer, _
                                   ByVal strHIGHCUSTCODE As String, _
                                   ByVal strHIGHCUSTNAME As String, _
                                   ByVal strCLIENTCODE As String, _
                                   ByVal strCLIENTNAME As String, _
                                   ByVal strEMPNAME As String) As Object

        Dim strCols As String        '컬럼변수
        Dim strWhere As String       'Where조건 변수
        Dim strFormat As String      'SQL Format 변수
        Dim strSQL As String         'SQL 변수
        Dim vntData As Object        'Array Return 변수(Array 를사용할 때 선언)
        Dim strSelFields
        Dim intCnt
        Dim intRtn
        Dim Con1, Con2, Con3, Con4, Con5
        SetConfig(strInfoXML)

        Con1 = "" : Con2 = "" : Con3 = "" : Con4 = "" : Con5 = ""

        If strHIGHCUSTCODE <> "" Then Con1 = String.Format(" AND (HIGHCUSTCODE = '{0}')", strHIGHCUSTCODE)
        If strHIGHCUSTNAME <> "" Then Con2 = String.Format(" AND (DBO.SC_GET_HIGHCOMPANYNAME_FUN(HIGHCUSTCODE) LIKE '%{0}%')", strHIGHCUSTNAME)
        If strCLIENTCODE <> "" Then Con3 = String.Format(" AND (CUSTCODE = '{0}')", strCLIENTCODE)
        If strCLIENTNAME <> "" Then Con4 = String.Format(" AND (DBO.SC_GET_CUSTNAME_FUN(CUSTCODE) LIKE '%{0}%')", strCLIENTNAME)
        If strEMPNAME <> "" Then Con5 = String.Format(" AND (EMP_NAME LIKE '%{0}%')", strEMPNAME)


        strWhere = BuildFields(" ", Con1, Con2, Con3, Con4, Con5)

        strFormat = "  SELECT "
        strFormat = strFormat & "  CUSTCODE,"
        strFormat = strFormat & "  DBO.SC_GET_CUSTNAME_FUN(CUSTCODE) CUSTNAME,"
        strFormat = strFormat & "  EMP_NAME,"
        strFormat = strFormat & "  EMP_EMAIL,"
        strFormat = strFormat & "  CASE ISNULL(DEF_GBN,'') WHEN '1' THEN 'Y' ELSE '' END DEF_GBN"
        strFormat = strFormat & "  FROM SC_CUST_EMP"
        strFormat = strFormat & "  WHERE 1=1 {0}"

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
                Throw RaiseSysErr(err, CLASS_NAME & ".GetSC_CUST_EMP")
            Finally
                ' DB 접속 종료
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "26. 대행사 조회"
    ' =============== Get_BrandInfo
    Public Function GetCGVCODE(ByVal strInfoXML As String, _
                               ByRef intRowCnt As Integer, _
                               ByRef intColCnt As Integer, _
                               ByVal strMEDCODE As String, _
                               ByVal strMEDNAME As String) As Object


        Dim strWhere As String       'Where조건 변수
        Dim strFormat As String      'SQL Format 변수
        Dim strSQL As String         'SQL 변수
        Dim vntData As Object        'Array Return 변수(Array 를사용할 때 선언)
        Dim Con1, Con2
        SetConfig(strInfoXML)

        Con1 = "" : Con2 = ""

        If strMEDCODE <> "" Then Con1 = String.Format(" AND (MEDCODE = '{0}')", strMEDCODE)
        If strMEDNAME <> "" Then Con2 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(MEDCODE) LIKE '%{0}%')", strMEDNAME)


        strWhere = BuildFields(" ", Con1, Con2)

        strFormat = " SELECT "
        strFormat = strFormat & " MEDCODE CUSTCODE,"
        strFormat = strFormat & " DBO.SC_GET_HIGHCUSTNAME_FUN(MEDCODE) CUSTNAME,"
        strFormat = strFormat & " DBO.SC_GET_BUSINO_FUN(MEDCODE) BUSINO,"
        strFormat = strFormat & " REAL_MED_CODE HIGHCUSTCODE,"
        strFormat = strFormat & " DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) COMPANYNAME"
        strFormat = strFormat & " FROM MD_CLOUD_CUST"
        strFormat = strFormat & " WHERE 1=1 {0} "
        strFormat = strFormat & " ORDER BY DBO.SC_GET_HIGHCUSTNAME_FUN(MEDCODE)"


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
                Throw RaiseSysErr(err, CLASS_NAME & ".GetCGVCODE")
            Finally
                ' DB 접속 종료
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region


#Region "10. 제작대행사 (대행사) 조회"
    ' =============== Get_BrandInfo
    Public Function Get_EXCLIENTCODE(ByVal strInfoXML As String, _
                                 ByRef intRowCnt As Integer, _
                                 ByRef intColCnt As Integer, _
                                 ByVal strEXCLIENTCODE As String, _
                                 ByVal strEXCLIENTNAME As String) As Object

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

        If strEXCLIENTCODE <> "" Then Con1 = String.Format(" AND (HIGHCUSTCODE = '{0}')", strEXCLIENTCODE)
        If strEXCLIENTNAME <> "" Then Con2 = String.Format(" AND (CUSTNAME LIKE '%{0}%')", strEXCLIENTNAME)
        'If strFLAG <> "" Then Con3 = String.Format(" AND (MEDFLAG = '{0}')", strFLAG)


        strWhere = BuildFields(" ", Con1, Con2)

        strFormat = "  SELECT "
        strFormat = strFormat & "  	CASE MEDFLAG "
        strFormat = strFormat & "  	WHEN 'G' THEN '대행사'"
        strFormat = strFormat & "  	WHEN 'K' THEN '크리조직'"
        strFormat = strFormat & "  	END AS	MEDFLAGNAME,"
        strFormat = strFormat & "  	HIGHCUSTCODE EXCLIENTCODE,"
        strFormat = strFormat & "  	CUSTNAME EXCLIENTNAME "
        strFormat = strFormat & "  	FROM SC_CUST_HDR"
        strFormat = strFormat & "  	WHERE MEDFLAG IN ('G') AND USE_FLAG = '1' "
        strFormat = strFormat & "   {0} ORDER BY MEDFLAG , EXCLIENTNAME"

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
                Throw RaiseSysErr(err, CLASS_NAME & ".Get_EXCLIENT")
            Finally
                ' DB 접속 종료
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
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

#End Region

End Class
