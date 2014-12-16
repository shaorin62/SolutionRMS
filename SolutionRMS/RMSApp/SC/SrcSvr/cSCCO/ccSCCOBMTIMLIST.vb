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

Public Class ccSCCOBMTIMLIST
    Inherits ccControl

#Region "GROUP BLOCK : 전역 또는 모듈레벨의 변수/상수 선언"
    Private CLASS_NAME = "ccSCCOBMTIMLIST"                  '자신의 클래스명
    Private mobjceSC_CCTR As eSCCO.ceSC_CCTR     'TYPE 헤더내역

#End Region

#Region "GROUP BLOCK : Function Section"
    Public Function Get_COMBO_VALUE(ByVal strInfoXML As String, _
                                     ByRef intRowCnt As Integer, _
                                     ByRef intColCnt As Integer) As Object

        Dim strSQL As String
        Dim vntData As Object

        SetConfig(strInfoXML)   '기본정보 설정					

        With mobjSCGLConfig

            strSQL = " SELECT "
            strSQL = strSQL & " BMORDER, BMNAME+ ' - ' + BMDEFINE  BMDEFINE"
            strSQL = strSQL & " FROM SC_BM  "
            strSQL = strSQL & " WHERE 1=1  "
            strSQL = strSQL & " ORDER BY BMCODE "

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

    '=============== bm별 조회
    Public Function SelectRtn(ByVal strInfoXML As String, _
                              ByRef intRowCnt As Integer, _
                              ByRef intColCnt As Integer, _
                              ByVal strBMNAME As String, _
                              ByVal strDEPTNAME As String, _
                              ByVal strUSE_YN As String) As Object

        Dim strCols As String         '컬럼변수
        Dim strWhere As String       'Where조건 변수
        Dim strFormat As String      'SQL Format 변수
        Dim strSQL As String          'SQL 변수
        Dim vntData As Object        'Array Return 변수(Array 를사용할 때 선언)
        Dim strXMLData As String    'XML  Return 변수(XML  을 사용할 때 선언)
        Dim Con1, Con2, Con3 As String


        If strBMNAME <> "" Then Con1 = String.Format(" AND (DBO.SC_GET_BMNAME_FUN(BMCODE) like '%{0}%')", strBMNAME)
        If strDEPTNAME <> "" Then Con2 = String.Format(" AND (DBO.SC_DEPT_NAME_FUN(DEPT_CD) like '%{0}%')", strDEPTNAME)
        If strUSE_YN <> "" Then Con3 = String.Format(" AND USE_YN = '{0}'", strUSE_YN)

        strWhere = BuildFields(" ", Con1, Con2, Con3)

        strFormat = "SELECT SEQ,"
        strFormat = strFormat & " BMCODE, "
        strFormat = strFormat & " HIGHDEPT_CD, "
        strFormat = strFormat & " DBO.SC_DEPT_NAME_FUN(HIGHDEPT_CD) HIGHDEPT_NAME,"
        strFormat = strFormat & " CCTR,"
        strFormat = strFormat & " BA,"
        strFormat = strFormat & " DEPT_CD,"
        strFormat = strFormat & " DBO.SC_DEPT_NAME_FUN(DEPT_CD) DEPT_NAME, "
        strFormat = strFormat & " FDATE, TDATE,"
        strFormat = strFormat & "  CASE USE_YN WHEN 'Y' THEN '1' ELSE '0' END USE_YN "
        strFormat = strFormat & " FROM SC_CCTR   "
        strFormat = strFormat & " WHERE 1=1 {0}"
        strFormat = strFormat & " ORDER BY BMCODE, CCTR, HIGHDEPT_NAME, DEPT_NAME"

        SetConfig(strInfoXML) '기본정보 Setting
        With mobjSCGLConfig '기본정보 Config 개체
            strSQL = String.Format(strFormat, strWhere)
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

    '부서 BM 매핑 저장 
    Public Function ProcessRtn(ByVal strInfoXML As String, _
                               ByVal vntData As Object) As Integer '데이터 INSERT/UPDATE
        Dim intRtn As Integer
        Dim i, intColCnt, intRows As Integer
        Dim strCODE As String
        Dim intCnt As Integer
        Dim strFDATE, strTDATE
        Dim strUSE_YN


        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                '하단그리드 저장처리
                If IsArray(vntData) Then
                    '''사용할 Entity 개체생성(Config 정보를 넘겨생성)
                    mobjceSC_CCTR = New ceSC_CCTR(mobjSCGLConfig)
                    '''vntData의 컬럼수, 로우수를 변수입력
                    intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)
                    intCnt = 0

                    For i = 1 To intRows
                        strFDATE = "" : strTDATE = ""
                        If GetElement(vntData, "FDATE", intColCnt, i, OPTIONAL_STR) <> "" Then strFDATE = GetElement(vntData, "FDATE", intColCnt, i, OPTIONAL_STR).Substring(0, 4) & GetElement(vntData, "FDATE", intColCnt, i, OPTIONAL_STR).Substring(5, 2) & GetElement(vntData, "FDATE", intColCnt, i, OPTIONAL_STR).Substring(8, 2)
                        If GetElement(vntData, "TDATE", intColCnt, i, OPTIONAL_STR) <> "" Then strTDATE = GetElement(vntData, "TDATE", intColCnt, i, OPTIONAL_STR).Substring(0, 4) & GetElement(vntData, "TDATE", intColCnt, i, OPTIONAL_STR).Substring(5, 2) & GetElement(vntData, "TDATE", intColCnt, i, OPTIONAL_STR).Substring(8, 2)

                        If GetElement(vntData, "USE_YN", intColCnt, i, OPTIONAL_NUM) = 1 Then
                            strUSE_YN = "Y"
                        Else
                            strUSE_YN = "N"
                        End If

                        If GetElement(vntData, "SEQ", intColCnt, i, NULL_NUM, True) = -999999 Then
                            intRtn = InsertRtn_SC_CCTR(vntData, intColCnt, i, strFDATE, strTDATE, strUSE_YN)
                        Else
                            intRtn = UpdateRtn_SC_CCTR(vntData, intColCnt, i, strFDATE, strTDATE, strUSE_YN)

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
                mobjceSC_CCTR.Dispose()

            End Try
        End With
    End Function

#End Region

#Region "GROUP BLOCK : Entity Function Section"
    'CE단에 보내기 위해 펑션을 만듬
    Private Function InsertRtn_SC_CCTR(ByVal vntData As Object, _
                                       ByVal intColCnt As Integer, _
                                       ByVal intRow As Integer, _
                                       ByVal strFDATE As String, _
                                       ByVal strTDATE As String, _
                                       ByVal strUSE_YN As String) As Integer

        Dim intRtn As Integer

        intRtn = mobjceSC_CCTR.InsertDo( _
                                GetElement(vntData, "BMCODE", intColCnt, intRow), _
                                GetElement(vntData, "HIGHDEPT_CD", intColCnt, intRow), _
                                GetElement(vntData, "DEPT_CD", intColCnt, intRow), _
                                GetElement(vntData, "CCTR", intColCnt, intRow), _
                                GetElement(vntData, "BA", intColCnt, intRow), _
                                strFDATE, _
                                strTDATE, _
                                strUSE_YN)
        Return intRtn
    End Function

    Private Function UpdateRtn_SC_CCTR(ByVal vntData As Object, _
                                       ByVal intColCnt As Integer, _
                                       ByVal intRow As Integer, _
                                       ByVal strFDATE As String, _
                                       ByVal strTDATE As String, _
                                       ByVal strUSE_YN As String) As Integer

        Dim intRtn As Integer

        intRtn = mobjceSC_CCTR.UpdateDo( _
                                GetElement(vntData, "SEQ", intColCnt, intRow, NULL_NUM, True), _
                                GetElement(vntData, "BMCODE", intColCnt, intRow), _
                                GetElement(vntData, "HIGHDEPT_CD", intColCnt, intRow), _
                                GetElement(vntData, "DEPT_CD", intColCnt, intRow), _
                                GetElement(vntData, "CCTR", intColCnt, intRow), _
                                GetElement(vntData, "BA", intColCnt, intRow), _
                                strFDATE, _
                                strTDATE, _
                                strUSE_YN)
        Return intRtn
    End Function


#End Region
End Class