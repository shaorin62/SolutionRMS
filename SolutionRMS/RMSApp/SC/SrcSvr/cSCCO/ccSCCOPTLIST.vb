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

Public Class ccSCCOPTLIST
    Inherits ccControl

#Region "GROUP BLOCK : 전역 또는 모듈레벨의 변수/상수 선언"
    Private CLASS_NAME = "ccSCCOPTLIST"                  '자신의 클래스명
    Private mobjceSC_PTLIST_MST As eSCCO.ceSC_PTLIST_MST '사용할 Entity 변수 선언
#End Region

#Region "GROUP BLOCK : Property 선언"
#End Region

#Region "GROUP BLOCK : Event 팝업"

    ' =============== SelectRtn_CUSTHDR 광고주 헤더
    Public Function SelectRtn_CalEndar(ByVal strInfoXML As String, _
                                       ByRef intRowCnt As Integer, _
                                       ByRef intColCnt As Integer, _
                                       ByVal strYEAR As String, _
                                       ByVal strMON As String) As Object     'XML  데이터 조회시

        Dim strSQL As String
        Dim strFormet As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '기본정보 설정

                strFormet = " DECLARE @TARGETYEAR  VARCHAR(4);  SET @TARGETYEAR = '" & strYEAR & "';"
                strFormet = strFormet & " DECLARE @TARGETMONTH VARCHAR(2);  SET @TARGETMONTH = '" & strMON & "';"
                strFormet = strFormet & " DECLARE @LOOPCOUNT  INT;   SET @LOOPCOUNT = 0;"
                strFormet = strFormet & " DECLARE @STARTDAY  DATETIME;"
                strFormet = strFormet & " DECLARE @DAYS   TABLE("
                strFormet = strFormet & "         DAYOFTHISMONTH  DATETIME, "
                strFormet = strFormet & "         DAYOFTHEWEEK  INT, "
                strFormet = strFormet & "         DAYOFTHEWEEKEX  NVARCHAR(3)"
                strFormet = strFormet & "         );"
                strFormet = strFormet & " IF (LEN(@TARGETMONTH) = 1) SET @TARGETMONTH = '0' + @TARGETMONTH;"
                strFormet = strFormet & " SET @STARTDAY = CONVERT(DATETIME, (@TARGETYEAR + @TARGETMONTH + '01'), 112);"
                strFormet = strFormet & " WHILE( 1 = 1 )"
                strFormet = strFormet & "  BEGIN"
                strFormet = strFormet & "   DECLARE @DAYOFTHEWEEKEX NVARCHAR(3);"
                strFormet = strFormet & "   DECLARE @DAY   DATETIME;"
                strFormet = strFormet & "    SET @DAY = DATEADD(DAY, @LOOPCOUNT, @STARTDAY);"
                strFormet = strFormet & "   IF (MONTH(@DAY) <> CONVERT(INT, @TARGETMONTH)) BREAK;"
                strFormet = strFormet & "   SET @DAYOFTHEWEEKEX = CASE DATEPART(WEEKDAY,@DAY)"
                strFormet = strFormet & "          WHEN 1 THEN '일'"
                strFormet = strFormet & "          WHEN 2 THEN '월'"
                strFormet = strFormet & "          WHEN 3 THEN '화'"
                strFormet = strFormet & "          WHEN 4 THEN '수'"
                strFormet = strFormet & "          WHEN 5 THEN '목'"
                strFormet = strFormet & "          WHEN 6 THEN '금'"
                strFormet = strFormet & "          WHEN 7 THEN '토'"
                strFormet = strFormet & "         END;"
                strFormet = strFormet & "   INSERT INTO @DAYS(DAYOFTHISMONTH,DAYOFTHEWEEK, DAYOFTHEWEEKEX) VALUES(@DAY, DATEPART(WEEKDAY,@DAY), @DAYOFTHEWEEKEX);"
                strFormet = strFormet & "   SET @LOOPCOUNT = @LOOPCOUNT + 1;"
                strFormet = strFormet & "  END"
                strFormet = strFormet & " SELECT "
                strFormet = strFormet & " YEARMON,"
                strFormet = strFormet & " MAX(CASE DAYOFTHEWEEK WHEN 1 THEN REALDAY end ) SUN,"
                strFormet = strFormet & " MAX(CASE DAYOFTHEWEEK WHEN 2 THEN REALDAY end ) MON,"
                strFormet = strFormet & " MAX(CASE DAYOFTHEWEEK WHEN 3 THEN REALDAY end ) TUE,"
                strFormet = strFormet & " MAX(CASE DAYOFTHEWEEK WHEN 4 THEN REALDAY end ) WED,"
                strFormet = strFormet & " MAX(CASE DAYOFTHEWEEK WHEN 5 THEN REALDAY end ) THU,"
                strFormet = strFormet & " MAX(CASE DAYOFTHEWEEK WHEN 6 THEN REALDAY end ) FRI,"
                strFormet = strFormet & " MAX(CASE DAYOFTHEWEEK WHEN 7 THEN REALDAY end ) SAT"
                strFormet = strFormet & " FROM ("
                strFormet = strFormet & " 	SELECT "
                strFormet = strFormet & " 	SUBSTRING(CONVERT(VARCHAR(10),DAYOFTHISMONTH,112),1,6) YEARMON,"
                strFormet = strFormet & " 	SUBSTRING(CONVERT(VARCHAR(10),DAYOFTHISMONTH,112),7,2) REALDAY,"
                strFormet = strFormet & " 	DAYOFTHEWEEK, DAYOFTHEWEEKEX,"
                strFormet = strFormet & " 	CEILING((CAST(SUBSTRING(CONVERT(VARCHAR(10),DAYOFTHISMONTH,112),7,2) AS NUMERIC)+"
                strFormet = strFormet & " 	DATEPART(WEEKDAY,CONVERT(DATETIME,SUBSTRING(CONVERT(VARCHAR(10),DAYOFTHISMONTH,112),1,6)+'01',112))-1)/7) WEEK"
                strFormet = strFormet & " 	FROM @DAYS"
                strFormet = strFormet & " )A  "
                strFormet = strFormet & " GROUP BY YEARMON,WEEK"

                strSQL = String.Format(strFormet)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)

                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_CalEndar")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    ' =============== SelectRtn_date 
    Public Function SelectRtn_date(ByVal strInfoXML As String, _
                                   ByRef intRowCnt As Integer, _
                                   ByRef intColCnt As Integer, _
                                   ByVal strYEARMON As String, _
                                   ByVal dblSEQ As String) As Object     'XML  데이터 조회시

        Dim strSQL As String
        Dim strFormet, strWhere As String
        Dim Con1, Con2 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '기본정보 설정

                Con1 = "" : Con2 = 0

                If strYEARMON <> "" Then Con1 = String.Format(" AND (YEARMON = '{0}')", strYEARMON)
                If dblSEQ <> "" Then Con2 = String.Format(" AND (SEQ = {0})", dblSEQ)
                strWhere = BuildFields(" ", Con1, Con2)

                strFormet = strFormet & " SELECT "
                strFormet = strFormet & " YEARMON, SEQ,"
                strFormet = strFormet & " DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME,"
                strFormet = strFormet & " OT_DATE,"
                strFormet = strFormet & " PT_DATE1,PT_DATE2,PT_DATE3 "
                strFormet = strFormet & " FROM SC_PTLIST_MST"
                strFormet = strFormet & " WHERE 1=1"
                strFormet = strFormet & " {0}"

                strSQL = String.Format(strFormet, strWhere)

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)

                Return vntData
            Catch err As Exception
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_date")
            Finally
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
#End Region

#Region "GROUP BLOCK : 외부에 공개 Method"
    ' =============== SelectRtn PT사용자 조회 입력자와 입력한 사람의 같은부서의 사람들끼리볼수 있다
    Public Function SelectRtn(ByVal strInfoXML As String, _
                              ByRef intRowCnt As Integer, _
                              ByRef intColCnt As Integer, _
                              ByVal strSTYEARMON As String, _
                              ByVal strEDYEARMON As String, _
                              ByVal strCLIENTCODE As String, _
                              ByVal strCLIENTNAME As String) As Object     'XML  데이터 조회시

        Dim strSQL As String
        Dim strFormet, strWhere As String
        Dim Con1, Con2, Con3 As String
        Dim vntData As Object
        Dim strUSER

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '기본정보 설정

                strUSER = .WRKUSR

                Con1 = "" : Con2 = "" : Con3 = ""

                If strSTYEARMON <> "" And strEDYEARMON <> "" Then
                    Con1 = String.Format(" AND (SUBSTRING(YEARMON,1,6) BETWEEN '{0}' AND  '{1}')", strSTYEARMON, strEDYEARMON)
                End If
                If strSTYEARMON <> "" And strEDYEARMON = "" Then
                    Con1 = String.Format(" AND (SUBSTRING(YEARMON,1,6) >= '{0}')", strSTYEARMON)
                End If
                If strSTYEARMON = "" And strEDYEARMON <> "" Then
                    Con1 = String.Format(" AND (SUBSTRING(YEARMON,1,6) <= '{0}')", strEDYEARMON)
                End If


                If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (CLIENTCODE = '{0}')", strCLIENTCODE)
                If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCLIENTNAME)
                strWhere = BuildFields(" ", Con1, Con2, Con3)

                strFormet = strFormet & " SELECT "
                strFormet = strFormet & " 0 CHK,"
                strFormet = strFormet & " YEARMON, SEQ,"
                strFormet = strFormet & " CLIENTCODE,"
                strFormet = strFormet & " CLIENTNAME,"
                strFormet = strFormet & " GREATCODE,"
                strFormet = strFormet & " GREATNAME,"
                strFormet = strFormet & " BUSINO,"
                strFormet = strFormet & " PT_STATUS,"
                strFormet = strFormet & " PT_LIST,"
                strFormet = strFormet & " ISNULL(EX_BILL,0) EX_BILL,"
                strFormet = strFormet & " OLDCLIENTNAME,"
                strFormet = strFormet & " PT_CLASS,"
                strFormet = strFormet & " EX_CONDITION,"
                strFormet = strFormet & " EX_INFO,"
                strFormet = strFormet & " OT_DATE,"
                strFormet = strFormet & " OT_INFO,"
                strFormet = strFormet & " ISNULL(PT_ESTAMT,0) PT_ESTAMT,"
                strFormet = strFormet & " ISNULL(PT_ACTAMT,0) PT_ACTAMT,"
                strFormet = strFormet & " PT_DATE1,PT_DATE2,PT_DATE3,"
                strFormet = strFormet & " PT_CLIENTNAME1,PT_CLIENTNAME2,PT_CLIENTNAME3,"
                strFormet = strFormet & " PT_RESULT,"
                strFormet = strFormet & " DEPT_CD,"
                strFormet = strFormet & " DBO.SC_DEPT_NAME_FUN(DEPT_CD) DEPT_NAME,"
                strFormet = strFormet & " EXCLIENTNAME,"
                strFormet = strFormet & " PT_TEXT"
                strFormet = strFormet & " FROM SC_PTLIST_MST"
                strFormet = strFormet & " WHERE 1=1"
                strFormet = strFormet & " AND DBO.SC_EMPNO_DEPT_CODE_FUN(CUSER) = DBO.SC_EMPNO_DEPT_CODE_FUN('" & strUSER & "') "
                strFormet = strFormet & " {0}"

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

    ' =============== SelectRtn PT리스트 전체 조회  모든 데이터 조회 가능
    Public Function SelectRtn_CON(ByVal strInfoXML As String, _
                                  ByRef intRowCnt As Integer, _
                                  ByRef intColCnt As Integer, _
                                  ByVal strSTYEARMON As String, _
                                  ByVal strEDYEARMON As String, _
                                  ByVal strCLIENTCODE As String, _
                                  ByVal strCLIENTNAME As String) As Object     'XML  데이터 조회시

        Dim strSQL As String
        Dim strFormet, strWhere As String
        Dim Con1, Con2, Con3 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '기본정보 설정

                Con1 = "" : Con2 = "" : Con3 = ""

                If strSTYEARMON <> "" And strEDYEARMON <> "" Then
                    Con1 = String.Format(" AND (SUBSTRING(YEARMON,1,6) BETWEEN '{0}' AND  '{1}')", strSTYEARMON, strEDYEARMON)
                End If
                If strSTYEARMON <> "" And strEDYEARMON = "" Then
                    Con1 = String.Format(" AND (SUBSTRING(YEARMON,1,6) >= '{0}')", strSTYEARMON)
                End If
                If strSTYEARMON = "" And strEDYEARMON <> "" Then
                    Con1 = String.Format(" AND (SUBSTRING(YEARMON,1,6) <= '{0}')", strEDYEARMON)
                End If


                If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (CLIENTCODE = '{0}')", strCLIENTCODE)
                If strCLIENTNAME <> "" Then Con3 = String.Format(" AND (DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) LIKE '%{0}%')", strCLIENTNAME)
                strWhere = BuildFields(" ", Con1, Con2, Con3)

                strFormet = strFormet & " SELECT "
                strFormet = strFormet & " 0 CHK,"
                strFormet = strFormet & " YEARMON, SEQ,"
                strFormet = strFormet & " CLIENTCODE,"
                strFormet = strFormet & " CLIENTNAME,"
                strFormet = strFormet & " GREATCODE,"
                strFormet = strFormet & " GREATNAME,"
                strFormet = strFormet & " BUSINO,"
                strFormet = strFormet & " PT_STATUS,"
                strFormet = strFormet & " PT_LIST,"
                strFormet = strFormet & " ISNULL(EX_BILL,0) EX_BILL,"
                strFormet = strFormet & " OLDCLIENTNAME,"
                strFormet = strFormet & " PT_CLASS,"
                strFormet = strFormet & " EX_CONDITION,"
                strFormet = strFormet & " EX_INFO,"
                strFormet = strFormet & " OT_DATE,"
                strFormet = strFormet & " OT_INFO,"
                strFormet = strFormet & " ISNULL(PT_ESTAMT,0) PT_ESTAMT,"
                strFormet = strFormet & " ISNULL(PT_ACTAMT,0) PT_ACTAMT,"
                strFormet = strFormet & " PT_DATE1,PT_DATE2,PT_DATE3,"
                strFormet = strFormet & " PT_CLIENTNAME1,PT_CLIENTNAME2,PT_CLIENTNAME3,"
                strFormet = strFormet & " PT_RESULT,"
                strFormet = strFormet & " DEPT_CD,"
                strFormet = strFormet & " DBO.SC_DEPT_NAME_FUN(DEPT_CD) DEPT_NAME,"
                strFormet = strFormet & " EXCLIENTNAME,"
                strFormet = strFormet & " PT_TEXT"
                strFormet = strFormet & " FROM SC_PTLIST_MST"
                strFormet = strFormet & " WHERE 1=1"
                strFormet = strFormet & " {0}"

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

    ' =============== ProcessRtn
    Public Function ProcessRtn(ByVal strInfoXML As String, _
                               ByVal vntData As Object) As Object
        Dim intRtn As Integer
        Dim i, intColCnt, intRows As Integer
        Dim dblSEQ
        Dim strYEARMON, strOT_DATE, strPT_DATE1, strPT_DATE2, strPT_DATE3


        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()

                If IsArray(vntData) Then
                    '''사용할 Entity 개체생성(Config 정보를 넘겨생성)
                    mobjceSC_PTLIST_MST = New ceSC_PTLIST_MST(mobjSCGLConfig)
                    '''vntData의 컬럼수, 로우수를 변수입력
                    intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)

                    For i = 1 To intRows
                        strYEARMON = "" : dblSEQ = ""
                        'strYEARMON = GetElement(vntData, "YEARMON", intColCnt, i, OPTIONAL_STR)

                        If GetElement(vntData, "YEARMON", intColCnt, i, OPTIONAL_STR) <> "" Then strYEARMON = GetElement(vntData, "YEARMON", intColCnt, i, OPTIONAL_STR).Substring(0, 4) & GetElement(vntData, "YEARMON", intColCnt, i, OPTIONAL_STR).Substring(5, 2) & GetElement(vntData, "YEARMON", intColCnt, i, OPTIONAL_STR).Substring(8, 2)

                        If GetElement(vntData, "OT_DATE", intColCnt, i, OPTIONAL_STR) <> "" Then strOT_DATE = GetElement(vntData, "OT_DATE", intColCnt, i, OPTIONAL_STR).Substring(0, 4) & GetElement(vntData, "OT_DATE", intColCnt, i, OPTIONAL_STR).Substring(5, 2) & GetElement(vntData, "OT_DATE", intColCnt, i, OPTIONAL_STR).Substring(8, 2)
                        If GetElement(vntData, "PT_DATE1", intColCnt, i, OPTIONAL_STR) <> "" Then strPT_DATE1 = GetElement(vntData, "PT_DATE1", intColCnt, i, OPTIONAL_STR).Substring(0, 4) & GetElement(vntData, "PT_DATE1", intColCnt, i, OPTIONAL_STR).Substring(5, 2) & GetElement(vntData, "PT_DATE1", intColCnt, i, OPTIONAL_STR).Substring(8, 2)
                        If GetElement(vntData, "PT_DATE2", intColCnt, i, OPTIONAL_STR) <> "" Then strPT_DATE2 = GetElement(vntData, "PT_DATE2", intColCnt, i, OPTIONAL_STR).Substring(0, 4) & GetElement(vntData, "PT_DATE2", intColCnt, i, OPTIONAL_STR).Substring(5, 2) & GetElement(vntData, "PT_DATE2", intColCnt, i, OPTIONAL_STR).Substring(8, 2)
                        If GetElement(vntData, "PT_DATE3", intColCnt, i, OPTIONAL_STR) <> "" Then strPT_DATE3 = GetElement(vntData, "PT_DATE3", intColCnt, i, OPTIONAL_STR).Substring(0, 4) & GetElement(vntData, "PT_DATE3", intColCnt, i, OPTIONAL_STR).Substring(5, 2) & GetElement(vntData, "PT_DATE3", intColCnt, i, OPTIONAL_STR).Substring(8, 2)


                        If GetElement(vntData, "SEQ", intColCnt, i, NULL_NUM, True) = -999999 Then
                            intRtn = InsertRtn(vntData, intColCnt, i, strYEARMON, strOT_DATE, strPT_DATE1, strPT_DATE2, strPT_DATE3)
                        Else
                            dblSEQ = GetElement(vntData, "SEQ", intColCnt, i, OPTIONAL_NUM)
                            intRtn = UpdateRtn(vntData, intColCnt, i, strYEARMON, dblSEQ, strOT_DATE, strPT_DATE1, strPT_DATE2, strPT_DATE3)
                        End If
                    Next
                End If

                .mobjSCGLSql.SQLCommitTrans()
                Return intRows
            Catch err As Exception
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".ProcessRtn")
            Finally
                .mobjSCGLSql.SQLDisconnect()
                mobjceSC_PTLIST_MST.Dispose()
            End Try
        End With
    End Function

    ' =============== DeleteRtn 
    Public Function DeleteRtn(ByVal strInfoXML As String, _
                              ByVal strYEARMON As String, _
                              ByVal dblSEQ As Integer) As Integer   '데이터 DELETE

        Dim intRtn As Integer      'Return변수( 처리건수 또는 0 )

        SetConfig(strInfoXML)    '기본정보 Setting
        With mobjSCGLConfig    '기본정보 Config 개체
            Try
                ' DB 접속 및 트랜잭션 시작
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                ' 사용할Entity 개체생성(Config 정보를 넘겨생성)
                mobjceSC_PTLIST_MST = New ceSC_PTLIST_MST(mobjSCGLConfig)

                strYEARMON = strYEARMON.Substring(0, 4) & strYEARMON.Substring(5, 2) & strYEARMON.Substring(8, 2)

                ' 엔티티 오브젝트의 Delete 메소드 호출
                intRtn = mobjceSC_PTLIST_MST.DeleteDO(strYEARMON, dblSEQ)
                ' 트랜잭션 Commit
                .mobjSCGLSql.SQLCommitTrans()
                Return intRtn
            Catch err As Exception
                '트랜잭션 RollBack 및 오류 전송
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & "DeleteRtn")
            Finally
                'DB접속 종료
                .mobjSCGLSql.SQLDisconnect()
                '사용한 Entity(개체Dispose)
                mobjceSC_PTLIST_MST.Dispose()
            End Try
        End With
    End Function

#End Region

#Region "GROUP BLOCK : 외부에 비공개 Method"
    Private Function InsertRtn(ByVal vntData As Object, _
                               ByVal intColCnt As Integer, _
                               ByVal intRow As Integer, _
                               ByVal strYEARMON As String, _
                               ByVal strOT_DATE As String, _
                               ByVal strPT_DATE1 As String, _
                               ByVal strPT_DATE2 As String, _
                               ByVal strPT_DATE3 As String) As Integer

        Dim intRtn As Integer
        intRtn = mobjceSC_PTLIST_MST.InsertDo( _
                                       strYEARMON, _
                                       GetElement(vntData, "CLIENTCODE", intColCnt, intRow), _
                                       GetElement(vntData, "CLIENTNAME", intColCnt, intRow), _
                                       GetElement(vntData, "GREATCODE", intColCnt, intRow), _
                                       GetElement(vntData, "GREATNAME", intColCnt, intRow), _
                                       GetElement(vntData, "BUSINO", intColCnt, intRow), _
                                       GetElement(vntData, "PT_STATUS", intColCnt, intRow), _
                                       GetElement(vntData, "PT_LIST", intColCnt, intRow), _
                                       GetElement(vntData, "EX_BILL", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "OLDCLIENTNAME", intColCnt, intRow), _
                                       GetElement(vntData, "PT_CLASS", intColCnt, intRow), _
                                       GetElement(vntData, "EX_CONDITION", intColCnt, intRow), _
                                       GetElement(vntData, "EX_INFO", intColCnt, intRow), _
                                       strOT_DATE, _
                                       GetElement(vntData, "OT_INFO", intColCnt, intRow), _
                                       GetElement(vntData, "PT_ESTAMT", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "PT_ACTAMT", intColCnt, intRow, NULL_NUM, True), _
                                       strPT_DATE1, _
                                       strPT_DATE2, _
                                       strPT_DATE3, _
                                       GetElement(vntData, "PT_CLIENTNAME1", intColCnt, intRow), _
                                       GetElement(vntData, "PT_CLIENTNAME2", intColCnt, intRow), _
                                       GetElement(vntData, "PT_CLIENTNAME3", intColCnt, intRow), _
                                       GetElement(vntData, "DEPT_CD", intColCnt, intRow), _
                                       GetElement(vntData, "EXCLIENTNAME", intColCnt, intRow), _
                                       GetElement(vntData, "PT_RESULT", intColCnt, intRow), _
                                       GetElement(vntData, "PT_TEXT", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR01", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR02", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR03", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR04", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR05", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR06", intColCnt, intRow, NULL_NUM, True))

        Return intRtn
    End Function

    Private Function UpdateRtn(ByVal vntData As Object, _
                               ByVal intColCnt As Integer, _
                               ByVal intRow As Integer, _
                               ByVal strYEARMON As String, _
                               ByVal dblSEQ As Double, _
                               ByVal strOT_DATE As String, _
                               ByVal strPT_DATE1 As String, _
                               ByVal strPT_DATE2 As String, _
                               ByVal strPT_DATE3 As String) As Integer
        Dim intRtn As Integer

        intRtn = mobjceSC_PTLIST_MST.UpdateDo( _
                                       strYEARMON, _
                                       dblSEQ, _
                                       GetElement(vntData, "CLIENTCODE", intColCnt, intRow), _
                                       GetElement(vntData, "CLIENTNAME", intColCnt, intRow), _
                                       GetElement(vntData, "GREATCODE", intColCnt, intRow), _
                                       GetElement(vntData, "GREATNAME", intColCnt, intRow), _
                                       GetElement(vntData, "BUSINO", intColCnt, intRow), _
                                       GetElement(vntData, "PT_STATUS", intColCnt, intRow), _
                                       GetElement(vntData, "PT_LIST", intColCnt, intRow), _
                                       GetElement(vntData, "EX_BILL", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "OLDCLIENTNAME", intColCnt, intRow), _
                                       GetElement(vntData, "PT_CLASS", intColCnt, intRow), _
                                       GetElement(vntData, "EX_CONDITION", intColCnt, intRow), _
                                       GetElement(vntData, "EX_INFO", intColCnt, intRow), _
                                       strOT_DATE, _
                                       GetElement(vntData, "OT_INFO", intColCnt, intRow), _
                                       GetElement(vntData, "PT_ESTAMT", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "PT_ACTAMT", intColCnt, intRow, NULL_NUM, True), _
                                       strPT_DATE1, _
                                       strPT_DATE2, _
                                       strPT_DATE3, _
                                       GetElement(vntData, "PT_CLIENTNAME1", intColCnt, intRow), _
                                       GetElement(vntData, "PT_CLIENTNAME2", intColCnt, intRow), _
                                       GetElement(vntData, "PT_CLIENTNAME3", intColCnt, intRow), _
                                       GetElement(vntData, "DEPT_CD", intColCnt, intRow), _
                                       GetElement(vntData, "EXCLIENTNAME", intColCnt, intRow), _
                                       GetElement(vntData, "PT_RESULT", intColCnt, intRow), _
                                       GetElement(vntData, "PT_TEXT", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR01", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR02", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR03", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR04", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR05", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "ATTR06", intColCnt, intRow, NULL_NUM, True))
        Return intRtn
    End Function
#End Region
End Class



