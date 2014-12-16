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

Public Class ccSCCOBANKTYPE
    Inherits ccControl

#Region "GROUP BLOCK : 전역 또는 모듈레벨의 변수/상수 선언"
    Private CLASS_NAME = "ccSCCOBANKTYPE"                  '자신의 클래스명
    Private mobjceSC_BANKTYPE_MST As eSCCO.ceSC_BANKTYPE_MST     'TYPE 헤더내역

#End Region

#Region "GROUP BLOCK : Function Section"
   
    '=============== BANK_TYPE 자료 조회
    Public Function SelectRtn(ByVal strInfoXML As String, _
                              ByRef intRowCnt As Integer, _
                              ByRef intColCnt As Integer, _
                              ByVal strBUSINO As String) As Object

        Dim strWhere As String       'Where조건 변수
        Dim strFormat As String      'SQL Format 변수
        Dim strSQL As String          'SQL 변수
        Dim vntData As Object        'Array Return 변수(Array 를사용할 때 선언)
        Dim Con1 As String

        If strBUSINO <> "" Then Con1 = String.Format(" AND (BUSINO = '{0}')", strBUSINO)

        strWhere = BuildFields(" ", Con1)

        strFormat = strFormat & " SELECT "
        strFormat = strFormat & " DBO.SC_BUSINO_CUSTNAME_FUN(BUSINO) CUSTNAME,"
        strFormat = strFormat & " substring(BUSINO,1,3) + '-' + substring(BUSINO,4,2) + '-' + substring(BUSINO,6,5) BUSINO,"
        strFormat = strFormat & " BANK_KEY,"
        strFormat = strFormat & " BANK_NUM,"
        strFormat = strFormat & " BANK_TYPE,"
        strFormat = strFormat & " BANK_USER,"
        strFormat = strFormat & " CASE USE_YN WHEN 'Y' THEN '1' ELSE '0' END USE_YN"
        strFormat = strFormat & " FROM SC_BANKTYPE_MST"
        strFormat = strFormat & " WHERE 1=1"
        strFormat = strFormat & " {0}"
        strFormat = strFormat & " ORDER BY DBO.SC_BUSINO_CUSTNAME_FUN(BUSINO) "

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


    ' =============== ProcessRtn_CUSTDTL    거래처 디테일 저장
    Public Function ProcessRtn(ByVal strInfoXML As String, _
                               ByVal vntData As Object) As Object

        Dim intRtn As Integer
        Dim i, intColCnt, intRows As Integer
        Dim strBUSINO
        Dim strUSEYN

        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                If IsArray(vntData) Then
                    '''사용할 Entity 개체생성(Config 정보를 넘겨생성)
                    mobjceSC_BANKTYPE_MST = New ceSC_BANKTYPE_MST(mobjSCGLConfig)
                    '''vntData의 컬럼수, 로우수를 변수입력
                    intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)

                    '1번째 로우의 사업자번호로 이전 BANKTYPE 을 삭제 후에 새로운 데이터로 업데이트 한다.
                    strBUSINO = GetElement(vntData, "BUSINO", intColCnt, 1, OPTIONAL_STR)
                    intRtn = mobjceSC_BANKTYPE_MST.DeleteDo(strBUSINO)

                    For i = 1 To intRows
                        strUSEYN = GetElement(vntData, "USE_YN", intColCnt, 1, OPTIONAL_STR)

                        If strUSEYN = "1" Then
                            strUSEYN = "Y"
                        Else
                            strUSEYN = "N"
                        End If

                        intRtn = InsertRtn(vntData, intColCnt, i, strUSEYN)
                    Next
                End If
                .mobjSCGLSql.SQLCommitTrans()
                Return intRows
            Catch err As Exception
                .mobjSCGLSql.SQLRollbackTrans()
                Throw RaiseSysErr(err, CLASS_NAME & ".ProcessRtn")
            Finally
                .mobjSCGLSql.SQLDisconnect()
                mobjceSC_BANKTYPE_MST.Dispose()
            End Try
        End With
    End Function


#End Region

#Region "GROUP BLOCK : Entity Function Section"

    Private Function InsertRtn(ByVal vntData As Object, _
                               ByVal intColCnt As Integer, _
                               ByVal intRow As Integer, _
                               ByVal strUSEYN As String) As Integer
        Dim intRtn As Integer
        intRtn = mobjceSC_BANKTYPE_MST.InsertDo( _
                                       GetElement(vntData, "BUSINO", intColCnt, intRow), _
                                       GetElement(vntData, "BANK_KEY", intColCnt, intRow), _
                                       GetElement(vntData, "BANK_NUM", intColCnt, intRow), _
                                       GetElement(vntData, "BANK_TYPE", intColCnt, intRow), _
                                       GetElement(vntData, "BANK_USER", intColCnt, intRow), _
                                       strUSEYN, _
                                       GetElement(vntData, "ATTR01", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR02", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR03", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR04", intColCnt, intRow), _
                                       GetElement(vntData, "ATTR05", intColCnt, intRow))
        Return intRtn
    End Function
#End Region
End Class