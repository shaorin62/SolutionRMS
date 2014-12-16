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
Public Class ccSCCOFEE
    Inherits ccControl

#Region "GROUP BLOCK : 전역 또는 모듈레벨의 변수/상수 선언"
    Private CLASS_NAME = "ccSCCOFEE"                  '자신의 클래스명
    Private mobjceSC_FEE_MST As eSCCO.ceSC_FEE_MST            '사용할 Entity 변수 선언
#End Region

#Region "GROUP BLOCK : Event 선언"
    '============== 콤보값을 세팅하기위한 Function
    Public Function Get_COMBO_VALUE(ByVal strInfoXML As String, _
                                    ByRef intRowCnt As Integer, _
                                    ByRef intColCnt As Integer, _
                                    ByVal strCLASS_CODE As String) As Object

        Dim strSQL As String
        Dim vntData As Object

        SetConfig(strInfoXML)   '기본정보 설정					

        'SQL문 생성
        strSQL = "SELECT "
        strSQL = strSQL & " CODE, CODE_NAME,SORT_SEQ "
        strSQL = strSQL & " FROM SC_CODE"
        strSQL = strSQL & " WHERE CLASS_CODE = '" & strCLASS_CODE & "' "
        strSQL = strSQL & " ORDER BY SORT_SEQ "

        With mobjSCGLConfig
            ''데이터 조회
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

#Region "GROUP BLOCK : 외부에 공개 Method"
    'Fee거래처 조회
    Public Function SelectRtn(ByVal strInfoXML As String, _
                              ByRef intRowCnt As Integer, _
                              ByRef intColCnt As Integer, _
                              ByVal strYEARMON As String, _
                              ByVal strCONFIRMFLAG As String, _
                              ByVal strCLIENTCODE As String, _
                              ByVal strCLIENTNAME As String) As Object

        Dim strWhere As String       'Where조건 변수
        Dim strFormat As String      'SQL Format 변수
        Dim strSQL As String         'SQL 변수
        Dim vntData As Object        'Array Return 변수(Array 를사용할 때 선언)
        Dim Con1, Con2, Con3, Con4 As String       '조건 String 변수

        With mobjSCGLConfig '기본정보 Config 개체
            Try
                SetConfig(strInfoXML) '기본정보 Setting

                Con1 = "" : Con2 = "" : Con3 = "" : Con4 = ""
                If strYEARMON <> "" Then Con4 = String.Format(" AND A.YEARMON = '{0}'", strYEARMON)

                If strCONFIRMFLAG = "X" Then

                ElseIf strCONFIRMFLAG = "1" Then
                    Con1 = String.Format(" AND A.CONFIRMFLAG = '{0}'", strCONFIRMFLAG)
                ElseIf strCONFIRMFLAG = "2" Then
                    Con1 = String.Format(" AND ISNULL(A.CONFIRMFLAG,'') = '{0}'", "")
                End If

                If strCLIENTCODE <> "" Then Con2 = String.Format(" AND A.CLIENTCODE like '%{0}%'", strCLIENTCODE)
                If strCLIENTNAME <> "" Then Con3 = String.Format(" AND DBO.SC_GET_HIGHCUSTNAME_FUN(A.CLIENTCODE) like '%{0}%'", strCLIENTNAME)

                strWhere = BuildFields(" ", Con1, Con2, Con3, Con4)

                strFormat = strFormat & " SELECT "
                strFormat = strFormat & " A.YEARMON, A.SEQ,"
                strFormat = strFormat & " A.CLIENTCODE, "
                strFormat = strFormat & " DBO.SC_GET_HIGHCUSTNAME_FUN(A.CLIENTCODE) CLIENTNAME, "
                strFormat = strFormat & " A.DEPT_CD,"
                strFormat = strFormat & " DBO.SC_DEPT_NAME_FUN(A.DEPT_CD) DEPT_NAME ,"
                strFormat = strFormat & " A.FDATE,A.EDATE,"
                strFormat = strFormat & " A.DEMANDDAY,"
                strFormat = strFormat & " DBO.SC_FEESUMAMT_FUN('" & strYEARMON & "',A.CLIENTCODE) MONTHAMT, "
                strFormat = strFormat & " DBO.SC_FEERATE_FUN('" & strYEARMON & "',A.CLIENTCODE) * 100 SUSURATE,"
                strFormat = strFormat & " ISNULL(B.ALLAMT,0) FEEAMT,"
                strFormat = strFormat & " ISNULL((B.TV + B.SBS_TV),0) MED_TV,"
                strFormat = strFormat & " ISNULL((B.RADIO + B.SBS_RADIO),0) MED_RD,"
                strFormat = strFormat & " ISNULL((B.DMB + B.SBS_DMB),0) MED_DMB,"
                strFormat = strFormat & " ISNULL(B.CABLETV,0) + ISNULL(B.TOTALTV,0) MED_CATV,"
                strFormat = strFormat & " ISNULL(B.BOOK,0) MED_PAP,"
                strFormat = strFormat & " ISNULL(B.OUT,0) MED_OUT,"
                strFormat = strFormat & " A.CONFIRMFLAG, "
                strFormat = strFormat & " A.VOCHNO "
                strFormat = strFormat & " FROM SC_FEE_MST A LEFT JOIN ("
                strFormat = strFormat & " 	SELECT"
                strFormat = strFormat & " 	YEARMON,CLIENTCODE,"
                strFormat = strFormat & " 	SUM(ALLAMT) ALLAMT,"
                strFormat = strFormat & " 	SUM(TV_AMT) TV,SUM(SBS_TV_AMT) SBS_TV ,SUM(RADIO_AMT) RADIO,"
                strFormat = strFormat & " 	SUM(SBS_RADIO_AMT) SBS_RADIO,SUM(DMB_AMT) DMB,SUM(SBS_DMB_AMT) SBS_DMB,"
                strFormat = strFormat & " 	SUM(CABLETV_AMT) CABLETV,SUM(TOTALTV_AMT) TOTALTV,SUM(BOOK_AMT) BOOK,SUM(OUT_AMT) OUT"
                strFormat = strFormat & " 	FROM( "
                strFormat = strFormat & " 		SELECT "
                strFormat = strFormat & " 		YEARMON, CLIENTCODE, "
                strFormat = strFormat & " 		CASE GUBUN WHEN '대변' THEN SUM(AMT) ELSE 0 END  ALLAMT,"
                strFormat = strFormat & " 		CASE GUBUN WHEN 'TV' THEN SUM(AMT) ELSE 0 END TV_AMT,"
                strFormat = strFormat & " 		CASE GUBUN WHEN 'SBS_TV' THEN SUM(AMT) ELSE 0 END SBS_TV_AMT,"
                strFormat = strFormat & " 		CASE GUBUN WHEN 'RADIO' THEN SUM(AMT) ELSE 0 END RADIO_AMT,"
                strFormat = strFormat & " 		CASE GUBUN WHEN 'SBS_RADIO' THEN SUM(AMT) ELSE 0 END SBS_RADIO_AMT,"
                strFormat = strFormat & " 		CASE GUBUN WHEN 'DMB' THEN SUM(AMT) ELSE 0 END DMB_AMT,"
                strFormat = strFormat & " 		CASE GUBUN WHEN 'SBS_DMB' THEN SUM(AMT) ELSE 0 END SBS_DMB_AMT,"
                strFormat = strFormat & " 		CASE GUBUN WHEN 'CABLETV' THEN SUM(AMT) ELSE 0 END CABLETV_AMT,"
                strFormat = strFormat & " 		CASE GUBUN WHEN '종합편성채널' THEN SUM(AMT) ELSE 0 END TOTALTV_AMT,"
                strFormat = strFormat & " 		CASE GUBUN WHEN '신문' THEN SUM(AMT) ELSE 0 END BOOK_AMT,"
                strFormat = strFormat & " 		CASE GUBUN WHEN '옥외' THEN SUM(AMT) ELSE 0 END OUT_AMT"
                strFormat = strFormat & " 		FROM V_FEE('" & strYEARMON & "')"
                strFormat = strFormat & " 		GROUP BY YEARMON,CLIENTCODE,GUBUN"
                strFormat = strFormat & " 	)A"
                strFormat = strFormat & " 	GROUP BY YEARMON,CLIENTCODE"
                strFormat = strFormat & " )B "
                strFormat = strFormat & " ON A.CLIENTCODE = B.CLIENTCODE"
                strFormat = strFormat & " WHERE 1=1"
                strFormat = strFormat & " {0} "

                strSQL = String.Format(strFormat, strWhere)
                .mobjSCGLSql.SQLConnect(.DBConnStr)
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

    ''============== ProcessRtn Fee관리 거래처 저장 
    Public Function ProcessRtn(ByVal strInfoXML As String, _
                               ByVal vntData As Object) As Object '데이터 INSERT/UPDATE
        Dim intRtn As Integer
        Dim i, intColCnt, intRows As Integer
        Dim strSQL As String
        Dim dblSEQ
        Dim strDEMANDDAY, strFDATE, strEDATE

        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                If IsArray(vntData) Then
                    '''사용할 Entity 개체생성(Config 정보를 넘겨생성)
                    mobjceSC_FEE_MST = New ceSC_FEE_MST(mobjSCGLConfig)
                    'vntData의 컬럼수, 로우수를 변수입력
                    intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)

                    For i = 1 To intRows
                        'If GetElement(vntData, "FDATE", intColCnt, i, OPTIONAL_STR) <> "" Then strFDATE = GetElement(vntData, "FDATE", intColCnt, i, OPTIONAL_STR).Substring(0, 4) & GetElement(vntData, "FDATE", intColCnt, i, OPTIONAL_STR).Substring(5, 2) & GetElement(vntData, "FDATE", intColCnt, i, OPTIONAL_STR).Substring(8, 2)
                        'If GetElement(vntData, "EDATE", intColCnt, i, OPTIONAL_STR) <> "" Then strEDATE = GetElement(vntData, "EDATE", intColCnt, i, OPTIONAL_STR).Substring(0, 4) & GetElement(vntData, "EDATE", intColCnt, i, OPTIONAL_STR).Substring(5, 2) & GetElement(vntData, "EDATE", intColCnt, i, OPTIONAL_STR).Substring(8, 2)
                        If GetElement(vntData, "DEMANDDAY", intColCnt, i, OPTIONAL_STR) <> "" Then strDEMANDDAY = GetElement(vntData, "DEMANDDAY", intColCnt, i, OPTIONAL_STR).Substring(0, 4) & GetElement(vntData, "DEMANDDAY", intColCnt, i, OPTIONAL_STR).Substring(5, 2) & GetElement(vntData, "DEMANDDAY", intColCnt, i, OPTIONAL_STR).Substring(8, 2)

                        If GetElement(vntData, "SEQ", intColCnt, i, NULL_NUM, True) = -999999 Then
                            intRtn = InsertRtn(vntData, intColCnt, i, strDEMANDDAY)
                            '수정일때
                        Else
                            dblSEQ = GetElement(vntData, "SEQ", intColCnt, i, NULL_NUM, True)
                            intRtn = UpdateRtn(vntData, intColCnt, i, dblSEQ, strDEMANDDAY)
                        End If
                    Next
                End If
                .mobjSCGLSql.SQLCommitTrans()
                Return intRtn
            Catch err As Exception
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".ProcessRtn")
            Finally
                .mobjSCGLSql.SQLDisconnect()
                mobjceSC_FEE_MST.Dispose()
            End Try
        End With
    End Function

#End Region

#Region "GROUP BLOCK : 외부에 비공개 Method"
    Private Function InsertRtn(ByVal vntData As Object, _
                               ByVal intColCnt As Integer, _
                               ByVal intRow As Integer, _
                               ByVal strDEMANDDAY As String) As Integer

        Dim intRtn As Integer
        intRtn = mobjceSC_FEE_MST.InsertDo( _
                                       GetElement(vntData, "YEARMON", intColCnt, intRow), _
                                       GetElement(vntData, "CLIENTCODE", intColCnt, intRow), _
                                       GetElement(vntData, "DEPT_CD", intColCnt, intRow), _
                                       strDEMANDDAY, _
                                       GetElement(vntData, "FDATE", intColCnt, intRow), _
                                       GetElement(vntData, "EDATE", intColCnt, intRow), _
                                       GetElement(vntData, "MONTHAMT", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "SUSURATE", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "MED_TV", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "MED_RD", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "MED_DMB", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "MED_CATV", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "MED_PAP", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "MED_OUT", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "CONFIRMFLAG", intColCnt, intRow), _
                                       GetElement(vntData, "VOCHNO", intColCnt, intRow), _
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

    Private Function UpdateRtn(ByVal vntData As Object, _
                               ByVal intColCnt As Integer, _
                               ByVal intRow As Integer, _
                               ByVal dblSEQ As Integer, _
                               ByVal strDEMANDDAY As Integer) As Integer
        Dim intRtn As Integer

        intRtn = mobjceSC_FEE_MST.UpdateDo( _
                                       GetElement(vntData, "YEARMON", intColCnt, intRow), _
                                       GetElement(vntData, "CLIENTCODE", intColCnt, intRow), _
                                       dblSEQ, _
                                       GetElement(vntData, "DEPT_CD", intColCnt, intRow), _
                                       strDEMANDDAY, _
                                       GetElement(vntData, "FDATE", intColCnt, intRow), _
                                       GetElement(vntData, "EDATE", intColCnt, intRow), _
                                       GetElement(vntData, "MONTHAMT", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "SUSURATE", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "MED_TV", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "MED_RD", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "MED_DMB", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "MED_CATV", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "MED_PAP", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "MED_OUT", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "CONFIRMFLAG", intColCnt, intRow))

        Return intRtn
    End Function
#End Region
End Class
