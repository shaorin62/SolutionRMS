'****************************************************************************************
'Generated By: MakeSFAR V.2.0.0 - 컨트롤 클래스 메이커 - 한화 S&C
'시스템구분    : 솔루션명 /시스템명/Server Control Class
'실행   환경    : COM+ Service Server Package
'프로그램명    : ccMDOTOUTDOOR_REPORT.vb
'기         능    : - 기능을 명시 합니다.
'특이  사항     : - 특이사항에 대해 표현
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2004-03-30 오전 10:32:13 By MakeSFARV.2.0.0
'            2) 2004-03-30 오전 10:32:13 By 작성자명을 씁니다.
'****************************************************************************************

Imports System.Xml                  ' XML처리
Imports SCGLControl                 ' ControlClass의 Base Class
Imports SCGLUtil.cbSCGLConfig       ' ConfigurationClass
Imports SCGLUtil.cbSCGLErr          '오류처리 클래스
Imports SCGLUtil.cbSCGLXml          'XML처리 클래스
Imports SCGLUtil.cbSCGLUtil         '기타유틸리티 클래스
Imports eMDCO                       '엔터티 추가

' 엔티티 클래스 사용시 해당 엔티티 클래스의 프로젝트를 참조한 후 Imports 하십시요. 
' Imports 엔티티프로젝트

Public Class ccMDOTOUTDOOR_REPORT
    Inherits ccControl

#Region "GROUP BLOCK : 전역 또는 모듈레벨의 변수/상수 선언"
    Private CLASS_NAME = "ccMDOTOUTDOOR_REPORT"                  '자신의 클래스명
    Private mobjceMD_OUTDOOR_REPORT As eMDCO.ceMD_OUTDOOR_REPORT             '사용할 Entity 변수 선언
#End Region

#Region "GROUP BLOCK : Event 선언"
#End Region

#Region "GROUP BLOCK : 외부에 공개 Method"
    Public Function SelectRtn(ByVal strInfoXML As String, _
                              ByRef intRowCnt As Integer, _
                              ByRef intColCnt As Integer, _
                              ByVal strYEARMON As String, _
                              ByVal strCLIENTCODE As String, _
                              ByVal strREAL_MED_CODE As String, _
                              ByVal strTIMCODE As String, _
                              ByVal strTITLE As String) As Object     'XML  데이터 조회시

        Dim strSQL As String
        Dim strFormet, strWhere As String
        Dim Con1, Con2, Con3, Con4, Con5 As String
        Dim vntData As Object

        With mobjSCGLConfig
            Try
                SetConfig(strInfoXML)   '기본정보 설정

                Con1 = "" : Con2 = "" : Con3 = "" : Con4 = "" : Con5 = ""

                If strYEARMON <> "" Then Con1 = String.Format(" AND (YEARMON = '{0}')", strYEARMON)
                If strCLIENTCODE <> "" Then Con2 = String.Format(" AND (CLIENTCODE = '{0}')", strCLIENTCODE)
                If strREAL_MED_CODE <> "" Then Con3 = String.Format(" AND (REAL_MED_CODE = '{0}')", strREAL_MED_CODE)
                If strTIMCODE <> "" Then Con4 = String.Format(" AND (TIMCODE = '{0}')", strTIMCODE)
                If strTITLE <> "" Then Con5 = String.Format(" AND (TITLE LIKE '%{0}%')", strTITLE)

                strWhere = BuildFields(" ", Con1, Con2, Con3, Con4, Con5)

                strFormet = " SELECT  "
                strFormet = strFormet & " 0 CHK, "
                strFormet = strFormet & " YEARMON, SEQ, "
                strFormet = strFormet & " '승인' GUBUN, "
                strFormet = strFormet & " CLIENTCODE, DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) CLIENTNAME, "
                strFormet = strFormet & " TIMCODE, DBO.SC_GET_CUSTNAME_FUN(TIMCODE) TIMNAME, "
                strFormet = strFormet & " REAL_MED_CODE, DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) REAL_MED_NAME, "
                strFormet = strFormet & " DBO.SC_GET_BUSINO_FUN(REAL_MED_CODE) REAL_MED_BISNO, "
                strFormet = strFormet & " MED_FLAG, "
                strFormet = strFormet & " DEMANDDAY, TBRDSTDATE, TBRDEDDATE, "
                strFormet = strFormet & " GBN_FLAG, TITLE, MATTERNAME, TOTALAMT, AMT, OUT_AMT, COMMI_RATE, COMMISSION, "
                strFormet = strFormet & " MED_GBN, LOCATION, MEMO, VOCH_TYPE,  "
                strFormet = strFormet & " COMMI_TRANS_NO, TRU_VOCH_NO, "
                strFormet = strFormet & " CASE ISNULL(ATTR01,'') WHEN '' THEN YEARMON + '-' + CAST(SEQ AS VARCHAR(10)) "
                strFormet = strFormet & " ELSE ATTR01 END AS ATTR01 "
                strFormet = strFormet & " FROM MD_OUTDOOR_REPORT "
                strFormet = strFormet & " WHERE 1=1 {0} "
                strFormet = strFormet & " order by SEQ ASC "

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

    Public Function SelectRtn_SEQ(ByVal strYEARMON As String) As String
        '여기부터 단순조회
        Dim strSQL, strFormat, strRtn As String

        With mobjSCGLConfig '기본정보 Config 개체

            Try
                strSQL = "SELECT "
                strSQL = strSQL & " ISNULL(Max(SEQ),0) +1 "
                strSQL = strSQL & " FROM MD_OUTDOOR_REPORT "
                strSQL = strSQL & " WHERE YEARMON = '" & strYEARMON & "'"

                strRtn = .mobjSCGLSql.SQLSelectOneScalar(strSQL)

                Return strRtn
            Catch err As Exception
                ' 오류 전송
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_SEQ")
            Finally
            End Try
        End With
        '여기까지 단순조회
    End Function

    ' =============== ProcessRtn
    Public Function ProcessRtn(ByVal strInfoXML As String, _
                               ByVal vntData As Object) As Object

        Dim intRtn As Integer
        Dim i, intColCnt, intRows As Integer
        Dim strYEARMON, strOLDYEARMON
        Dim strSEQ, strSEQ1, strOLDSEQ
        Dim strDEMANDDAY
        Dim strTBRDSTDATE, strTBRDEDDATE
        Dim strSQL, strMEMO, strCOMMI_TAX_FLAG, strGBN_FLAG

        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                If IsArray(vntData) Then
                    '''사용할 Entity 개체생성(Config 정보를 넘겨생성)
                    mobjceMD_OUTDOOR_REPORT = New ceMD_OUTDOOR_REPORT(mobjSCGLConfig)
                    '''vntData의 컬럼수, 로우수를 변수입력
                    intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)
                    For i = 1 To intRows
                        strSEQ = ""
                        '신규자료면
                        If GetElement(vntData, "SEQ", intColCnt, i, NULL_NUM, True) = -999999 Then
                            '년월 정리
                            If Len(GetElement(vntData, "YEARMON", intColCnt, i, OPTIONAL_STR)) <> 6 Then
                                strYEARMON = GetElement(vntData, "YEARMON", intColCnt, i, OPTIONAL_STR).Substring(0, 4) & GetElement(vntData, "YEARMON", intColCnt, i, OPTIONAL_STR).Substring(5, 2)
                            Else
                                strYEARMON = GetElement(vntData, "YEARMON", intColCnt, i, OPTIONAL_STR)
                            End If

                            If GetElement(vntData, "DEMANDDAY", intColCnt, i, OPTIONAL_STR) <> "" Then strDEMANDDAY = GetElement(vntData, "DEMANDDAY", intColCnt, i, OPTIONAL_STR).Substring(0, 4) & GetElement(vntData, "DEMANDDAY", intColCnt, i, OPTIONAL_STR).Substring(5, 2) & GetElement(vntData, "DEMANDDAY", intColCnt, i, OPTIONAL_STR).Substring(8, 2)
                            If GetElement(vntData, "TBRDSTDATE", intColCnt, i, OPTIONAL_STR) <> "" Then strTBRDSTDATE = GetElement(vntData, "TBRDSTDATE", intColCnt, i, OPTIONAL_STR).Substring(0, 4) & GetElement(vntData, "TBRDSTDATE", intColCnt, i, OPTIONAL_STR).Substring(5, 2) & GetElement(vntData, "TBRDSTDATE", intColCnt, i, OPTIONAL_STR).Substring(8, 2)
                            If GetElement(vntData, "TBRDEDDATE", intColCnt, i, OPTIONAL_STR) <> "" Then strTBRDEDDATE = GetElement(vntData, "TBRDEDDATE", intColCnt, i, OPTIONAL_STR).Substring(0, 4) & GetElement(vntData, "TBRDEDDATE", intColCnt, i, OPTIONAL_STR).Substring(5, 2) & GetElement(vntData, "TBRDEDDATE", intColCnt, i, OPTIONAL_STR).Substring(8, 2)

                            strSEQ = SelectRtn_SEQ(strYEARMON)

                            intRtn = InsertRtn_MD_OUTDOOR_REPORT(vntData, intColCnt, i, strYEARMON, strSEQ, strDEMANDDAY, strTBRDSTDATE, strTBRDEDDATE)
                        Else
                            '년월 정리
                            If Len(GetElement(vntData, "YEARMON", intColCnt, i, OPTIONAL_STR)) <> 6 Then
                                strYEARMON = GetElement(vntData, "YEARMON", intColCnt, i, OPTIONAL_STR).Substring(0, 4) & GetElement(vntData, "YEARMON", intColCnt, i, OPTIONAL_STR).Substring(5, 2)
                            Else
                                strYEARMON = GetElement(vntData, "YEARMON", intColCnt, i, OPTIONAL_STR)
                            End If

                            If GetElement(vntData, "DEMANDDAY", intColCnt, i, OPTIONAL_STR) <> "" Then strDEMANDDAY = GetElement(vntData, "DEMANDDAY", intColCnt, i, OPTIONAL_STR).Substring(0, 4) & GetElement(vntData, "DEMANDDAY", intColCnt, i, OPTIONAL_STR).Substring(5, 2) & GetElement(vntData, "DEMANDDAY", intColCnt, i, OPTIONAL_STR).Substring(8, 2)
                            If GetElement(vntData, "TBRDSTDATE", intColCnt, i, OPTIONAL_STR) <> "" Then strTBRDSTDATE = GetElement(vntData, "TBRDSTDATE", intColCnt, i, OPTIONAL_STR).Substring(0, 4) & GetElement(vntData, "TBRDSTDATE", intColCnt, i, OPTIONAL_STR).Substring(5, 2) & GetElement(vntData, "TBRDSTDATE", intColCnt, i, OPTIONAL_STR).Substring(8, 2)
                            If GetElement(vntData, "TBRDEDDATE", intColCnt, i, OPTIONAL_STR) <> "" Then strTBRDEDDATE = GetElement(vntData, "TBRDEDDATE", intColCnt, i, OPTIONAL_STR).Substring(0, 4) & GetElement(vntData, "TBRDEDDATE", intColCnt, i, OPTIONAL_STR).Substring(5, 2) & GetElement(vntData, "TBRDEDDATE", intColCnt, i, OPTIONAL_STR).Substring(8, 2)

                            strSEQ = GetElement(vntData, "SEQ", intColCnt, i, OPTIONAL_STR)

                            intRtn = UpdateRtn_MD_OUTDOOR_REPORT(vntData, intColCnt, i, strYEARMON, strSEQ, strDEMANDDAY, strTBRDSTDATE, strTBRDEDDATE)
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
                mobjceMD_OUTDOOR_REPORT.Dispose()
            End Try
        End With
    End Function

    ' =============== DeleteRtn Sample Code
    Public Function DeleteRtn(ByVal strInfoXML As String, _
                              ByVal strYEARMON As String, _
                              ByVal strSEQ As Integer) As Integer   '데이터 DELETE

        Dim intRtn As Integer
        Dim intRtn2 As Integer

        SetConfig(strInfoXML)    '기본정보 Setting
        With mobjSCGLConfig    '기본정보 Config 개체
            Try
                ' 사용할Entity 개체생성(Config 정보를 넘겨생성)
                mobjceMD_OUTDOOR_REPORT = New ceMD_OUTDOOR_REPORT(mobjSCGLConfig)
                ' DB 접속 및 트랜잭션 시작
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                ' 엔티티 오브젝트의 Delete 메소드 호출
                intRtn = mobjceMD_OUTDOOR_REPORT.DeleteDo(strYEARMON, strSEQ)
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
                mobjceMD_OUTDOOR_REPORT.Dispose()
            End Try
        End With
    End Function
#End Region

#Region "GROUP BLOCK : 외부에 비공개 Method"
    Private Function InsertRtn_MD_OUTDOOR_REPORT(ByVal vntData As Object, _
                                                 ByVal intColCnt As Integer, _
                                                 ByVal intRow As Integer, _
                                                 ByRef strYEARMON As String, _
                                                 ByRef strSEQ As Integer, _
                                                 ByRef strDEMANDDAY As String, _
                                                 ByRef strTBRDSTDATE As String, _
                                                 ByRef strTBRDEDDATE As String) As Integer

        Dim intRtn As Integer
        intRtn = mobjceMD_OUTDOOR_REPORT.InsertDo( _
                                       strYEARMON, _
                                       strSEQ, _
                                       GetElement(vntData, "CLIENTCODE", intColCnt, intRow), _
                                       GetElement(vntData, "TIMCODE", intColCnt, intRow), _
                                       GetElement(vntData, "REAL_MED_CODE", intColCnt, intRow), _
                                       GetElement(vntData, "REAL_MED_CODE", intColCnt, intRow), _
                                       "D", _
                                       GetElement(vntData, "HIGHSUBSEQ", intColCnt, intRow), _
                                       GetElement(vntData, "DEPT_CD", intColCnt, intRow), _
                                       GetElement(vntData, "MATTERNAME", intColCnt, intRow), _
                                       strDEMANDDAY, _
                                       GetElement(vntData, "TITLE", intColCnt, intRow), _
                                       GetElement(vntData, "MED_GBN", intColCnt, intRow), _
                                       GetElement(vntData, "LOCATION", intColCnt, intRow), _
                                       strTBRDSTDATE, _
                                       strTBRDEDDATE, _
                                       "", _
                                       GetElement(vntData, "TOTALAMT", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "AMT", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "COMMI_RATE", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "COMMISSION", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "OUT_AMT", intColCnt, intRow, NULL_NUM, True), _
                                       "1", _
                                       "", _
                                       GetElement(vntData, "MEMO", intColCnt, intRow), _
                                       GetElement(vntData, "CONTIDX", intColCnt, intRow), _
                                       GetElement(vntData, "MDIDX", intColCnt, intRow), _
                                       GetElement(vntData, "CYEAR", intColCnt, intRow), _
                                       GetElement(vntData, "CMONTH", intColCnt, intRow), _
                                       GetElement(vntData, "SIDE", intColCnt, intRow), _
                                       GetElement(vntData, "PORTAL_SEQ", intColCnt, intRow), _
                                       GetElement(vntData, "VOCH_TYPE", intColCnt, intRow), _
                                       "Y", _
                                       "Y", _
                                       GetElement(vntData, "TRU_TRANS_NO", intColCnt, intRow), _
                                       GetElement(vntData, "TRU_TAX_NO", intColCnt, intRow), _
                                       GetElement(vntData, "TRU_VOCH_NO", intColCnt, intRow), _
                                       GetElement(vntData, "COMMI_TRANS_NO", intColCnt, intRow), _
                                       GetElement(vntData, "COMMI_TAX_NO", intColCnt, intRow), _
                                       GetElement(vntData, "COMMI_VOCH_NO", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR01", intColCnt, intRow), _
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

    Private Function UpdateRtn_MD_OUTDOOR_REPORT(ByVal vntData As Object, _
                                                 ByVal intColCnt As Integer, _
                                                 ByVal intRow As Integer, _
                                                 ByRef strYEARMON As String, _
                                                 ByRef strSEQ As Integer, _
                                                 ByRef strDEMANDDAY As String, _
                                                 ByRef strTBRDSTDATE As String, _
                                                 ByRef strTBRDEDDATE As String) As Integer


        Dim intRtn As Integer
        intRtn = mobjceMD_OUTDOOR_REPORT.UpdateDo( _
                                       strYEARMON, _
                                       strSEQ, _
                                       GetElement(vntData, "CLIENTCODE", intColCnt, intRow), _
                                       GetElement(vntData, "TIMCODE", intColCnt, intRow), _
                                       GetElement(vntData, "REAL_MED_CODE", intColCnt, intRow), _
                                       GetElement(vntData, "REAL_MED_CODE", intColCnt, intRow), _
                                       "D", _
                                       GetElement(vntData, "HIGHSUBSEQ", intColCnt, intRow), _
                                       GetElement(vntData, "DEPT_CD", intColCnt, intRow), _
                                       GetElement(vntData, "MATTERNAME", intColCnt, intRow), _
                                       strDEMANDDAY, _
                                       GetElement(vntData, "TITLE", intColCnt, intRow), _
                                       GetElement(vntData, "MED_GBN", intColCnt, intRow), _
                                       GetElement(vntData, "LOCATION", intColCnt, intRow), _
                                       strTBRDSTDATE, _
                                       strTBRDEDDATE, _
                                       "", _
                                       GetElement(vntData, "TOTALAMT", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "AMT", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "COMMI_RATE", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "COMMISSION", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "OUT_AMT", intColCnt, intRow, NULL_NUM, True), _
                                       "1", _
                                       "", _
                                       GetElement(vntData, "MEMO", intColCnt, intRow), _
                                       GetElement(vntData, "VOCH_TYPE", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR01", intColCnt, intRow), _
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
#End Region

End Class