'****************************************************************************************
'시스템구분    : 솔루션명 /시스템명/Server Control Class
'실행   환경    : COM+ Service Server Package
'프로그램명    : ccMDCMCUST_TRAN.vb
'기         능    : - 기능을 명시 합니다.
'특이  사항     : - 특이사항에 대해 표현
'                     -
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009-07-03 오전 10:32:13 By KTY
'****************************************************************************************

Imports System.Xml                  ' XML처리
Imports SCGLControl                 ' ControlClass의 Base Class 
Imports SCGLUtil.cbSCGLConfig       ' ConfigurationClass
Imports SCGLUtil.cbSCGLErr          '오류처리 클래스
Imports SCGLUtil.cbSCGLXml          'XML처리 클래스
Imports SCGLUtil.cbSCGLUtil         '기타유틸리티 클래스
Imports eSCCO '엔터티 추가

' 엔티티 클래스 사용시 해당 엔티티 클래스의 프로젝트를 참조한 후 Imports 하십시요. 
' Imports 엔티티프로젝트

Public Class ccSCCOBRANDLIST
    Inherits ccControl

#Region "GROUP BLOCK : 전역 또는 모듈레벨의 변수/상수 선언"
    Private CLASS_NAME = "ccSCCOBRANDLIST"                  '자신의 클래스명
    Private mobjceSC_SUBSEQ_HDR As eSCCO.ceSC_SUBSEQ_HDR            '사용할 Entity 변수 선언
    Private mobjceSC_SUBSEQ_DTL As eSCCO.ceSC_SUBSEQ_DTL              '사용할 Entity 변수 선언
#End Region

#Region "GROUP BLOCK : Property 선언"
#End Region

#Region "GROUP BLOCK : Event 선언"
    Public Function HIGHSEQNAME_Check(ByVal strInfoXML As String, _
                                      ByRef intRowCnt As Integer, _
                                      ByRef intColCnt As Integer, _
                                      ByVal strHIGHSEQNAME As String, _
                                      ByVal strHIGHCUSTCODE As String) As Object                                      'XML  데이터 조회시

        Dim strSQL As String
        Dim strFormet, strSelFields, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '기본정보 설정

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
                                      ByRef intColCnt As Integer) As Object                                      'XML  데이터 조회시

        Dim strSQL As String
        Dim strFormet, strWhere As String
        Dim Con1 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '기본정보 설정

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
                                         ByRef strCLIENTCODE As String) As Object                                      'XML  데이터 조회시

        Dim strSQL As String
        Dim strFormet, strWhere As String
        Dim Con1 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '기본정보 설정

                Con1 = ""

                If strCLIENTCODE <> "" Then Con1 = String.Format(" and CUSTCODE = '{0}'", strCLIENTCODE) '년월

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
                                         ByVal strMEDFLAG As String) As Object     'XML  데이터 조회시

        Dim strSQL As String
        Dim strFormat, strSelFields, strWhere As String
        Dim Con1, Con2 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '기본정보 설정

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

        Dim strFormat As String      'SQL Format 변수
        Dim strSQL As String         'SQL 변수
        Dim vntData As Object        'Array Return 변수(Array 를사용할 때 선언)
        Dim strDEPT_CD

        SetConfig(strInfoXML)

        strSQL = "  SELECT  "
        strSQL = strSQL & "  CC_CODE "
        strSQL = strSQL & "  From SC_EMPLOYEE_MST"
        strSQL = strSQL & "  WHERE EMPNO = '" & strUSERID & "' and use_yn = 'Y'"

        '기본정보 Setting
        With mobjSCGLConfig '기본정보 Config 개체
            Try
                ' DB 접속
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                strDEPT_CD = .mobjSCGLSql.SQLSelectOneScalar(strSQL)

                Return strDEPT_CD
            Catch err As Exception
                ' 오류 전송
                Throw RaiseSysErr(err, CLASS_NAME & ".Get_SESSION_DEPT_CD")
            Finally
                ' DB 접속 종료
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

#End Region

#Region "GROUP BLOCK : 외부에 공개 Method"
    ' =============== SelectRtn_HIGHSUBSEQ 대표브랜드 헤더
    Public Function SelectRtn_HIGHSUBSEQ(ByVal strInfoXML As String, _
                                         ByRef intRowCnt As Integer, _
                                         ByRef intColCnt As Integer, _
                                         ByVal strCUSTNAME As String, _
                                         ByVal strHIGHSEQNAME As String) As Object     'XML  데이터 조회시

        Dim strSQL As String
        Dim strFormat, strWhere As String
        Dim strChkDate As String = ""
        Dim Con1, Con2 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '기본정보 설정

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

    ' =============== SelectRtn_SUBSEQ 브랜드
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
                SetConfig(strInfoXML)   '기본정보 설정

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
                'strFormat = strFormat & "  MEMO, case isnull(ATTR01,'N') when 'N' then '미사용' else '사용' end as Attr01 "
                'strFormat = strFormat & "  FROM SC_SUBSEQ_DTL"
                'strFormat = strFormat & "  WHERE 1=1 {0} "

                strFormat = " SELECT "
                strFormat = strFormat & " 0 CHK, SEQNO, "
                strFormat = strFormat & " SEQNAME, HIGHSEQNO, "
                strFormat = strFormat & " TIMCODE, '' BTNTIM, DBO.SC_GET_CUSTNAME_FUN(TIMCODE) TIMNAME, "
                strFormat = strFormat & " CLIENTSUBCODE, '' BTNSUB, DBO.SC_GET_CUSTNAME_FUN(CLIENTSUBCODE) CLIENTSUBNAME, "
                strFormat = strFormat & " CUSTCODE, '' BTN, DBO.SC_GET_HIGHCUSTNAME_FUN(CUSTCODE) CUSTNAME, "
                strFormat = strFormat & " DEPT_CD, '' BTNDEPT,dbo.SC_DEPT_NAME_FUN(DEPT_CD)  DEPT_NAME, DBO.SC_EMPNAME_FUN(A.CUSER) CUSER, CDATE,"
                strFormat = strFormat & " MEMO, case isnull(ATTR01,'N') when 'N' then '미사용' WHEN 'Y' THEN '사용' WHEN 'S' THEN '승인요청' ELSE '등록' end as Attr01, "
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

    ' =============== SelectRtn_SUBSEQ 브랜드
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
                SetConfig(strInfoXML)   '기본정보 설정

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
                strFormat = strFormat & "  MEMO, case isnull(ATTR01,'N') when 'N' then '미사용' WHEN 'Y' THEN '사용' WHEN 'S' THEN '승인요청' ELSE '등록' end as Attr01 "
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


    ' =============== SelectRtn_CUSTDTL 광고주디테일
    Public Function SelectRtn_CUSTDTL(ByVal strInfoXML As String, _
                                      ByRef intRowCnt As Integer, _
                                      ByRef intColCnt As Integer, _
                                      ByRef strHIGHCUSTCODE As String) As Object     'XML  데이터 조회시

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

                If strHIGHCUSTCODE <> "" Then Con1 = String.Format(" AND (HIGHCUSTCODE = '{0}')", strHIGHCUSTCODE)

                strWhere = BuildFields(" ", Con1)

                strSelFields = " CASE GBNFLAG WHEN '0' THEN '팀' WHEN '1' THEN 'CIC/사업부' ELSE '' END GBNFLAG, "
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
                                   ByVal strHIGHSEQNO As String) As Integer   '데이터 DELETE

        Dim intRtn_desc As Integer      'Return변수( 처리건수 또는 0 )
        Dim intRtn As Integer      'Return변수( 처리건수 또는 0 )

        SetConfig(strInfoXML)    '기본정보 Setting
        With mobjSCGLConfig    '기본정보 Config 개체
            Try
                ' 사용할Entity 개체생성(Config 정보를 넘겨생성)
                mobjceSC_SUBSEQ_HDR = New ceSC_SUBSEQ_HDR(mobjSCGLConfig)
                ' DB 접속 및 트랜잭션 시작
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                ' 엔티티 오브젝트의 Delete 메소드 호출
                intRtn = mobjceSC_SUBSEQ_HDR.DeleteDo(strHIGHSEQNO)
                ' 트랜잭션 Commit
                .mobjSCGLSql.SQLCommitTrans()
                Return intRtn
            Catch err As Exception
                '트랜잭션 RollBack 및 오류 전송
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & "DeleteRtn_HDR")
            Finally
                'DB접속 종료
                .mobjSCGLSql.SQLDisconnect()
                '사용한 Entity(개체Dispose)
                mobjceSC_SUBSEQ_HDR.Dispose()
            End Try
        End With
    End Function

    Public Function DeleteRtn_DTL(ByVal strInfoXML As String, _
                                  ByVal strSEQNO As String) As Integer   '데이터 DELETE

        Dim intRtn_desc As Integer      'Return변수( 처리건수 또는 0 )
        Dim intRtn As Integer      'Return변수( 처리건수 또는 0 )

        SetConfig(strInfoXML)    '기본정보 Setting
        With mobjSCGLConfig    '기본정보 Config 개체
            Try
                ' 사용할Entity 개체생성(Config 정보를 넘겨생성)
                mobjceSC_SUBSEQ_DTL = New ceSC_SUBSEQ_DTL(mobjSCGLConfig)
                ' DB 접속 및 트랜잭션 시작
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                ' 엔티티 오브젝트의 Delete 메소드 호출
                intRtn = mobjceSC_SUBSEQ_DTL.DeleteDo(strSEQNO)
                ' 트랜잭션 Commit
                .mobjSCGLSql.SQLCommitTrans()
                Return intRtn
            Catch err As Exception
                '트랜잭션 RollBack 및 오류 전송
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & "DeleteRtn_DTL")
            Finally
                'DB접속 종료
                .mobjSCGLSql.SQLDisconnect()
                '사용한 Entity(개체Dispose)
                mobjceSC_SUBSEQ_DTL.Dispose()
            End Try
        End With
    End Function

    Public Function ProcessRtn_CONF(ByVal strInfoXML As String, _
                                    ByVal strSEQNO As String) As Integer   '데이터 DELETE

        Dim intRtn_desc As Integer      'Return변수( 처리건수 또는 0 )
        Dim intRtn As Integer      'Return변수( 처리건수 또는 0 )

        SetConfig(strInfoXML)    '기본정보 Setting
        With mobjSCGLConfig    '기본정보 Config 개체
            Try
                ' 사용할Entity 개체생성(Config 정보를 넘겨생성)
                mobjceSC_SUBSEQ_DTL = New ceSC_SUBSEQ_DTL(mobjSCGLConfig)
                ' DB 접속 및 트랜잭션 시작
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                ' 엔티티 오브젝트의 Delete 메소드 호출
                intRtn = mobjceSC_SUBSEQ_DTL.Update_Conf(strSEQNO)
                ' 트랜잭션 Commit
                .mobjSCGLSql.SQLCommitTrans()
                Return intRtn
            Catch err As Exception
                '트랜잭션 RollBack 및 오류 전송
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & "DeleteRtn_DTL")
            Finally
                'DB접속 종료
                .mobjSCGLSql.SQLDisconnect()
                '사용한 Entity(개체Dispose)
                mobjceSC_SUBSEQ_DTL.Dispose()
            End Try
        End With
    End Function

    Public Function ProcessRtn_CONFOK(ByVal strInfoXML As String, _
                                      ByVal strSEQNO As String) As Integer   '데이터 DELETE

        Dim intRtn_desc As Integer      'Return변수( 처리건수 또는 0 )
        Dim intRtn As Integer      'Return변수( 처리건수 또는 0 )

        SetConfig(strInfoXML)    '기본정보 Setting
        With mobjSCGLConfig    '기본정보 Config 개체
            Try
                ' 사용할Entity 개체생성(Config 정보를 넘겨생성)
                mobjceSC_SUBSEQ_DTL = New ceSC_SUBSEQ_DTL(mobjSCGLConfig)
                ' DB 접속 및 트랜잭션 시작
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                ' 엔티티 오브젝트의 Delete 메소드 호출
                intRtn = mobjceSC_SUBSEQ_DTL.Update_ConfOK(strSEQNO)
                ' 트랜잭션 Commit
                .mobjSCGLSql.SQLCommitTrans()
                Return intRtn
            Catch err As Exception
                '트랜잭션 RollBack 및 오류 전송
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & "DeleteRtn_DTL")
            Finally
                'DB접속 종료
                .mobjSCGLSql.SQLDisconnect()
                '사용한 Entity(개체Dispose)
                mobjceSC_SUBSEQ_DTL.Dispose()
            End Try
        End With
    End Function



    ' =============== ProcessRtn_HIGHSUBSEQ    대표브랜드 해더 저장
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
                    '''사용할 Entity 개체생성(Config 정보를 넘겨생성)
                    mobjceSC_SUBSEQ_HDR = New ceSC_SUBSEQ_HDR(mobjSCGLConfig)
                    '''vntData의 컬럼수, 로우수를 변수입력
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

    ' =============== ProcessRtn_SUBSEQ    브랜드 해더 저장
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
                    '''사용할 Entity 개체생성(Config 정보를 넘겨생성)
                    mobjceSC_SUBSEQ_DTL = New ceSC_SUBSEQ_DTL(mobjSCGLConfig)
                    '''vntData의 컬럼수, 로우수를 변수입력
                    intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)
                    For i = 1 To intRows
                        strSEQNO = ""
                        strATTR01 = ""

                        If GetElement(vntData, "SEQNO", intColCnt, i, OPTIONAL_STR) = "" Then
                            strSEQNO = Get_NewSeqNo(strYEAR)
                            strATTR01 = GetElement(vntData, "ATTR01", intColCnt, i, OPTIONAL_STR)
                            If strATTR01 = "사용" Then
                                strATTR01 = "Y"
                            ElseIf strATTR01 = "미사용" Then
                                strATTR01 = "N"
                            ElseIf strATTR01 = "미사용" Then
                                strATTR01 = "N"
                            ElseIf strATTR01 = "승인요청" Then
                                strATTR01 = "S"
                            Else
                                strATTR01 = "R"
                            End If
                            intRtn = InsertRtn_SC_SUBSEQ_DTL(vntData, intColCnt, i, strSEQNO, strATTR01)
                            strRETURNVALUE = intRtn & "-" & strSEQNO
                        Else
                            strSEQNO = GetElement(vntData, "SEQNO", intColCnt, i, OPTIONAL_STR)
                            strATTR01 = GetElement(vntData, "ATTR01", intColCnt, i, OPTIONAL_STR)
                            If strATTR01 = "사용" Then
                                strATTR01 = "Y"
                            ElseIf strATTR01 = "미사용" Then
                                strATTR01 = "N"
                            ElseIf strATTR01 = "승인요청" Then
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

    '==============SC_SUBSEQ_HDR 테이블의 신규 HIGHSEQNO 가져오기
    Public Function Get_NewHighSeqNo(ByVal strYEAR As String) As String
        Dim strSQL, strFormat, strRtn As String

        With mobjSCGLConfig '기본정보 Config 개체
            Try
                strSQL = "select 'S' +'" & strYEAR & "' + DBO.LPAD(ISNULL(MAX(CAST(SUBSTRING(HIGHSEQNO,4,5) AS NUMERIC(5,0))),0)+1,5,'0') From SC_SUBSEQ_HDR "
                strRtn = .mobjSCGLSql.SQLSelectOneScalar(strSQL)
                Return strRtn
            Catch err As Exception
                ' 오류 전송
                Throw RaiseSysErr(err, CLASS_NAME & ".Get_NewHighSeqNo")
            End Try
        End With
    End Function

    '==============SC_SUBSEQ_DTL 테이블의 신규 SEQNO 가져오기
    Public Function Get_NewSeqNo(ByVal strYEAR As String) As String
        Dim strSQL, strFormat, strRtn As String

        With mobjSCGLConfig '기본정보 Config 개체
            Try
                strSQL = "select 'S' +'" & strYEAR & "' + DBO.LPAD(ISNULL(MAX(CAST(SUBSTRING(SEQNO,4,5) AS NUMERIC(5,0))),0)+1,5,'0') From SC_SUBSEQ_DTL "
                strRtn = .mobjSCGLSql.SQLSelectOneScalar(strSQL)
                Return strRtn
            Catch err As Exception
                ' 오류 전송
                Throw RaiseSysErr(err, CLASS_NAME & ".Get_NewSeqNo")
            End Try
        End With
    End Function

    '신규 CUSTCODE 생성
    Public Function SelectRtn_CUSTCODE(ByVal strMEDFLAG As String) As String

        Dim strSQL As String
        Dim strFormat As String
        Dim strRtn As String

        With mobjSCGLConfig '기본정보 Config 개체

            Try
                strSQL = String.Format("select '{0}' + dbo.lpad(isnull(Max(substring(custcode,2,6)),0)+1,5,0) From SC_CUST_DTL WHERE MEDFLAG =  '{1}'", strMEDFLAG, strMEDFLAG)
                strRtn = .mobjSCGLSql.SQLSelectOneScalar(strSQL)
                Return strRtn
            Catch err As Exception
                ' 오류 전송
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_CUSTCODE")
            Finally
            End Try
        End With
    End Function
#End Region

#Region "GROUP BLOCK : 외부에 비공개 Method"
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



