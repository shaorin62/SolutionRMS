
'****************************************************************************************
'Generated By  : Kim Tae Ho 
'시스템구분    : RMS/PD/Server Control Class
'실행   환경   : COM+ Service Server Package
'프로그램명    : ccPDCMPONO.vb
'기         능 : - Project Number 생성
'특이  사항    : - 
'                -
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008.11.07 Kim Tae Ho
'            2) 
'****************************************************************************************

Imports System.Xml                  ' XML처리
Imports SCGLControl                 ' ControlClass의 Base Class
Imports SCGLUtil.cbSCGLConfig       ' ConfigurationClass
Imports SCGLUtil.cbSCGLErr          '오류처리 클래스
Imports SCGLUtil.cbSCGLXml          'XML처리 클래스
Imports SCGLUtil.cbSCGLUtil         '기타유틸리티 클래스
Imports ePDCO                       '엔터티 추가

' 엔티티 클래스 사용시 해당 엔티티 클래스의 프로젝트를 참조한 후 Imports 하십시요. 
' Imports 엔티티프로젝트
Public Class ccPDCODIVAMT
    Inherits ccControl
#Region "GROUP BLOCK : 전역 또는 모듈레벨의 변수/상수 선언"
    Private CLASS_NAME = "ccPDCODIVAMT"                  '자신의 클래스명
    Private mobjcePD_DIVAMT As ePDCO.cePD_DIVAMT           '사용할 Entity 변수 선언
    'Private Const .DBConnStr = "Provider=SQLOLEDB;Data Source=10.110.10.86;Initial Catalog=MCDEV;DSN=MCDEV;UID=devadmin;Pwd=password"
#End Region

#Region "GROUP BLOCK : Function Section"
    '=============== 분할대상 내역 조회
    Public Function SelectRtn(ByVal strInfoXML As String, _
                              ByRef intRowCnt As Integer, _
                              ByRef intColCnt As Integer, _
                              ByVal strYEARMON As String, _
                              ByVal strJOBNAME As String, _
                              ByVal strJOBNO As String, _
                              ByVal strCMBYN As String) As Object

        Dim strCols As String         '컬럼변수
        Dim strWhere As String       'Where조건 변수
        Dim strFormat As String      'SQL Format 변수
        Dim strSQL As String          'SQL 변수
        Dim vntData As Object        'Array Return 변수(Array 를사용할 때 선언)
        Dim strXMLData As String    'XML  Return 변수(XML  을 사용할 때 선언)
        Dim Con1, Con2, Con3, Con4
        strCols = " A.PREESTNO,"
        strCols = strCols & " SUBSTRING(A.CONFIRMFLAG,1,6) YEARMON,"
        strCols = strCols & " A.JOBNO,"
        strCols = strCols & " DBO.PD_JOBNAME_FUN(A.JOBNO)JOBNAME,"
        strCols = strCols & " A.SUMAMT DIVAMT,"
        strCols = strCols & " A.CLIENTCODE,"
        strCols = strCols & " DBO.MD_GET_CUSTNAME_FUN(A.CLIENTCODE) CLIENTNAME,"
        strCols = strCols & " A.CLIENTSUBCODE,"
        strCols = strCols & " DBO.MD_GET_CUSTNAME_FUN(A.CLIENTSUBCODE) CLIENTSUBNAME,CASE A.SUMAMT WHEN B.DIVAMT THEN 'Y' ELSE 'N' END AS INYN,A.CONFIRMFLAG CREDAY,CASE A.SUMAMT WHEN B.DIVAMT THEN '완료' ELSE '미완료' END AS INYNNM,B.ADJAMT"



        If strYEARMON <> "" Then Con1 = String.Format(" AND (SUBSTRING(A.CONFIRMFLAG,1,6) like '%{0}%')", strYEARMON)

        strJOBNAME = Replace(strJOBNAME, "'", "''")

        If strJOBNAME <> "" Then Con2 = String.Format(" AND (LTRIM(DBO.PD_JOBNAME_FUN(A.JOBNO)) like '%{0}%')", strJOBNAME)
        If strJOBNO <> "" Then Con3 = String.Format(" AND (A.JOBNO = '{0}')", strJOBNO)
        If strCMBYN <> "" Then Con4 = String.Format(" AND (CASE A.SUMAMT WHEN B.DIVAMT THEN 'Y' ELSE 'N' END = '{0}')", strCMBYN)
        strWhere = BuildFields(" ", Con1, Con2, Con3, Con4)
        strFormat = "SELECT {0} FROM PD_PREEST_HDR A LEFT JOIN (SELECT SUM(DIVAMT) DIVAMT,SUM(ADJAMT) ADJAMT, JOBNO FROM PD_DIVAMT GROUP BY JOBNO) B ON A.JOBNO = B.JOBNO  WHERE isnull(confirmflag,'') <> '' {1} ORDER BY PREESTNO"

        SetConfig(strInfoXML) '기본정보 Setting
        With mobjSCGLConfig '기본정보 Config 개체
            strSQL = String.Format(strFormat, strCols, strWhere)
            Try
                ' DB 접속
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                ' ------ Array 데이터 조회 (True 일때 헤더정보 포함 조회(Sheet Data Binding 할 경우 사용), False 일때 데이터만 조회)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
                ' ------ XML 데이터 조회
                'strXMLData = .mobjSCGLSql.SQLSelectXml(strSQL, intRowCnt, intColCnt)
                'Return strXMLData
            Catch err As Exception
                ' 오류 전송
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn")
            Finally
                ' DB 접속 종료
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function
    '=============== 분할 내역 조회
    Public Function SelectRtn_DIV(ByVal strInfoXML As String, _
                                  ByRef intRowCnt As Integer, _
                                  ByRef intColCnt As Integer, _
                                  ByVal strJOBNO As String) As Object

        Dim strCols As String         '컬럼변수
        Dim strWhere As String       'Where조건 변수
        Dim strFormat As String      'SQL Format 변수
        Dim strSQL As String          'SQL 변수
        Dim vntData As Object        'Array Return 변수(Array 를사용할 때 선언)
        Dim strXMLData As String    'XML  Return 변수(XML  을 사용할 때 선언)
        Dim Con1
        strCols = " PREESTNO,SEQ,JOBNO,YEARMON,CREDAY,CLIENTCODE,DBO.MD_GET_CUSTNAME_FUN(CLIENTCODE) CLIENTNAME,CLIENTSUBCODE,DBO.MD_GET_CUSTNAME_FUN(CLIENTSUBCODE) CLIENTSUBNAME,DIVAMT,JOBNAME,ADJAMT,SUBSEQ,DBO.PD_JOBCUST_NAME_FUN(SUBSEQ) SUBSEQNAME,ISNULL(ATTR02,'') ATTR02"

        If strJOBNO <> "" Then Con1 = String.Format(" AND (JOBNO = '{0}')", strJOBNO)

        strWhere = BuildFields(" ", Con1)
        strFormat = "SELECT {0} FROM PD_DIVAMT WHERE 1=1 {1} ORDER BY SEQ"

        SetConfig(strInfoXML) '기본정보 Setting
        With mobjSCGLConfig '기본정보 Config 개체
            strSQL = String.Format(strFormat, strCols, strWhere)
            Try
                ' DB 접속
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                ' ------ Array 데이터 조회 (True 일때 헤더정보 포함 조회(Sheet Data Binding 할 경우 사용), False 일때 데이터만 조회)
                vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
                Return vntData
                ' ------ XML 데이터 조회
                'strXMLData = .mobjSCGLSql.SQLSelectXml(strSQL, intRowCnt, intColCnt)
                'Return strXMLData
            Catch err As Exception
                ' 오류 전송
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_DIV")
            Finally
                ' DB 접속 종료
                .mobjSCGLSql.SQLDisconnect()
            End Try
        End With
    End Function

    '' ============== ProcessRtn PROJECT 코드 등록
    'Public Function ProcessRtn(ByVal strInfoXML As String, _
    '                           ByVal strMasterXML As String, _
    '                           ByVal strSEQFLAG As String) As Integer '데이터 INSERT/UPDATE
    '    Dim intRtn As Integer '결과값 변수
    '    Dim i, intColCnt, intRows As Integer '루프, 컬럼Cnt, 로우Cnt 변수

    '    SetConfig(strInfoXML) '기본정보 Setting
    '    With mobjSCGLConfig '기본정보를 가지고 있는 Config 개체
    '        Try
    '            'XML Element 변수 선언 (strMasterXML을 변환)
    '            Dim xmlRoot As XmlElement
    '            xmlRoot = XMLGetRoot(strMasterXML) 'XML 데이터

    '            'DB접속 및 트랜잭션 시작
    '            .mobjSCGLSql.SQLConnect(.DBConnStr)
    '            .mobjSCGLSql.SQLBeginTrans()
    '            'Master 데이터 처리
    '            intRtn = ProcessRtn_Seq(xmlRoot, strSEQFLAG)

    '            .mobjSCGLSql.SQLCommitTrans()
    '            Return intRtn
    '        Catch err As Exception
    '            '트랜잭션RollBack 및 오류 전송
    '            .mobjSCGLSql.SQLRollbackTrans()
    '            Throw RaiseSysErr(err, CLASS_NAME & ".ProcessRtn")
    '        Finally
    '            'Resource해제
    '            .mobjSCGLSql.SQLDisconnect()
    '        End Try
    '    End With
    'End Function

    ' =============== DeleteRtn Sample Code
    'strJOBNO,strPREESTNO,strSEQ
    Public Function DeleteRtn(ByVal strInfoXML As String, _
                              ByVal strJOBNO As String, _
                              ByVal strPREESTNO As String, _
                              ByVal dblSEQ As Integer) As Integer '데이터 DELETE
        Dim intRtn As Integer      'Return변수( 처리건수 또는 0 )
        SetConfig(strInfoXML)    '기본정보 Setting
        With mobjSCGLConfig    '기본정보 Config 개체
            Try
                ' 사용할Entity 개체생성(Config 정보를 넘겨생성)
                mobjcePD_DIVAMT = New cePD_DIVAMT(mobjSCGLConfig)
                ' DB 접속 및 트랜잭션 시작
                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                ' 엔티티 오브젝트의 Delete 메소드 호출
                intRtn = mobjcePD_DIVAMT.DeleteDo(dblSEQ, strJOBNO, strPREESTNO)
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
                mobjcePD_DIVAMT.Dispose()
            End Try
        End With
    End Function
    ''============== ProcessRtn PROJECT 코드 수정
    Public Function ProcessRtn(ByVal strInfoXML As String, _
                               ByVal vntData As Object, _
                               ByVal strJOBNO As String) As Object '데이터 INSERT/UPDATE
        Dim intRtn As Integer
        Dim i, intColCnt, intRows As Integer
        Dim strCREDAY As String
        Dim strYEARMON As String
        Dim dblSEQ As Integer

        SetConfig(strInfoXML)
        With mobjSCGLConfig
            Try

                .mobjSCGLSql.SQLConnect(.DBConnStr)
                .mobjSCGLSql.SQLBeginTrans()
                If IsArray(vntData) Then
                    '''사용할 Entity 개체생성(Config 정보를 넘겨생성)
                    mobjcePD_DIVAMT = New cePD_DIVAMT(mobjSCGLConfig)
                    '''vntData의 컬럼수, 로우수를 변수입력
                    intColCnt = UBound(vntData, 1) : intRows = UBound(vntData, 2)
                    For i = 1 To intRows
                        strYEARMON = ""
                        strCREDAY = ""
                        If GetElement(vntData, "SEQ", intColCnt, i, OPTIONAL_STR) = "" Then
                            dblSEQ = SelectRtn_SEQNO(strJOBNO)
                            strCREDAY = Mid(GetElement(vntData, "CREDAY", intColCnt, i, OPTIONAL_STR), 1, 4) & Mid(GetElement(vntData, "CREDAY", intColCnt, i, OPTIONAL_STR), 6, 2) & Mid(GetElement(vntData, "CREDAY", intColCnt, i, OPTIONAL_STR), 9, 2)
                            strYEARMON = Mid(strCREDAY, 1, 6)
                            intRtn = InsertRtn(vntData, intColCnt, i, dblSEQ, strYEARMON, strCREDAY)
                        Else
                            strCREDAY = Mid(GetElement(vntData, "CREDAY", intColCnt, i, OPTIONAL_STR), 1, 4) & Mid(GetElement(vntData, "CREDAY", intColCnt, i, OPTIONAL_STR), 6, 2) & Mid(GetElement(vntData, "CREDAY", intColCnt, i, OPTIONAL_STR), 9, 2)
                            strYEARMON = Mid(strCREDAY, 1, 6)
                            intRtn = UpdateRtn(vntData, intColCnt, i, strYEARMON, strCREDAY)
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
                mobjcePD_DIVAMT.Dispose()
            End Try
        End With
    End Function

    ''============== ProcessRtn PROJECT 신규 실제 저장 로직
    'Public Function ProcessRtn_Seq(ByVal xmlRoot As XmlElement, _
    '                               ByVal strSEQFLAG As String) As Integer
    '    Dim strPONO
    '    Dim intRtn
    '    Dim strCUSTCODE As String
    '    Dim strYEAR As String
    '    Dim strCREDAY As String

    '    With mobjSCGLConfig

    '        Try
    '            If XMLGetElement(xmlRoot, "CREDAY") <> "" Then strCREDAY = XMLGetElement(xmlRoot, "CREDAY").Substring(0, 4) & XMLGetElement(xmlRoot, "CREDAY").Substring(5, 2) & XMLGetElement(xmlRoot, "CREDAY").Substring(8, 2)
    '            ''' 사용할 Entity 개체생성(Config 정보를 넘겨생성)
    '            mobjcePD_PONO = New cePD_PONO(mobjSCGLConfig)


    '            strYEAR = Mid(XMLGetElement(xmlRoot, "CREDAY"), 3, 2)
    '            If strSEQFLAG = "new" Then
    '                strPONO = SelectRtn_SEQNO(strYEAR)
    '                intRtn = InsertRtn_PONO(xmlRoot, strPONO, strCREDAY)
    '            Else
    '                'intRtn = UpdateRtn_SEQ(xmlRoot)

    '            End If

    '            Return intRtn
    '        Catch err As Exception

    '            Throw RaiseSysErr(err, CLASS_NAME & ".ProcessRtn_Seq")
    '        Finally

    '            mobjcePD_PONO.Dispose()
    '        End Try
    '    End With
    'End Function

    '============== ProcessRtn PROJECT CODE 생성
    Public Function SelectRtn_SEQNO(ByVal strJOBNO As String) As String
        '여기부터 단순조회
        Dim strSQL, strFormat, strRtn As String
        'SetConfig(strInfoXML) '기본정보 Setting

        With mobjSCGLConfig '기본정보 Config 개체

            Try
                strSQL = "SELECT ISNULL(MAX(SEQ),0)+1 FROM PD_DIVAMT WHERE JOBNO = '" & strJOBNO & "'"
                strRtn = .mobjSCGLSql.SQLSelectOneScalar(strSQL)
                Return strRtn
            Catch err As Exception
                ' 오류 전송
                Throw RaiseSysErr(err, CLASS_NAME & ".SelectRtn_SEQNO")
            Finally
            End Try
        End With
        '여기까지 단순조회
    End Function
    '' =============== Project Code 삭제여부 판단 조회
    'Public Function GetPONODELSELECT(ByVal strInfoXML As String, _
    '                          ByRef intRowCnt As Integer, _
    '                          ByRef intColCnt As Integer, _
    '                          ByVal strCODE As String) As Object

    '    Dim strCols As String         '컬럼변수
    '    Dim strWhere As String       'Where조건 변수
    '    Dim strFormat As String      'SQL Format 변수
    '    Dim strSQL As String          'SQL 변수
    '    Dim vntData As Object        'Array Return 변수(Array 를사용할 때 선언)
    '    Dim strXMLData As String    'XML  Return 변수(XML  을 사용할 때 선언)



    '    SetConfig(strInfoXML) '기본정보 Setting
    '    With mobjSCGLConfig '기본정보 Config 개체
    '        strSQL = "SELECT PROJECTNO FROM PD_JOBNO WHERE PROJECTNO = '" & strCODE & "'"
    '        Try
    '            ' DB 접속
    '            .mobjSCGLSql.SQLConnect(.DBConnStr)
    '            ' ------ Array 데이터 조회 (True 일때 헤더정보 포함 조회(Sheet Data Binding 할 경우 사용), False 일때 데이터만 조회)
    '            vntData = .mobjSCGLSql.SQLSelectArr(strSQL, intRowCnt, intColCnt, , True)
    '            Return vntData
    '        Catch err As Exception
    '            ' 오류 전송
    '            Throw RaiseSysErr(err, CLASS_NAME & ".GetPONODELSELECT")
    '        Finally
    '            ' DB 접속 종료
    '            .mobjSCGLSql.SQLDisconnect()
    '        End Try
    '    End With
    'End Function
    ''프로젝트 삭제
    'Public Function DeleteRtn(ByVal strInfoXML As String, _
    '                          ByVal strCODE As String) As Integer

    '    Dim intRtn As Integer      'Return변수( 처리건수 또는 0 )

    '    SetConfig(strInfoXML)    '기본정보 Setting
    '    With mobjSCGLConfig    '기본정보 Config 개체
    '        Try
    '            ' 사용할Entity 개체생성(Config 정보를 넘겨생성)
    '            mobjcePD_PONO = New cePD_PONO(mobjSCGLConfig)
    '            ' DB 접속 및 트랜잭션 시작
    '            .mobjSCGLSql.SQLConnect(.DBConnStr)
    '            .mobjSCGLSql.SQLBeginTrans()
    '            ' 엔티티 오브젝트의 Delete 메소드 호출
    '            intRtn = mobjcePD_PONO.DeleteDo(strCODE)
    '            ' 트랜잭션 Commit
    '            .mobjSCGLSql.SQLCommitTrans()
    '            Return intRtn
    '        Catch err As Exception
    '            '트랜잭션 RollBack 및 오류 전송
    '            .mobjSCGLSql.SQLRollbackTrans()
    '            Throw RaiseSysErr(err, CLASS_NAME & "DeleteRtn")
    '        Finally
    '            'DB접속 종료
    '            .mobjSCGLSql.SQLDisconnect()
    '            '사용한 Entity(개체Dispose)
    '            mobjcePD_PONO.Dispose()
    '        End Try
    '    End With
    'End Function
#End Region
    'PROJECTNO,PROJECTNM,CLIENTCODE,CLIENTSUBCODE,SUBSEQ,GROUPGBN,CREDAY,CPDEPTCD,CPEMPNO,MEMO
#Region "GROUP BLOCK : Entity Function Section"
    Private Function InsertRtn(ByVal vntData As Object, _
                                    ByVal intColCnt As Integer, _
                                    ByVal intRow As Integer, _
                                    ByVal dblSEQ As Double, _
                                    ByVal strYEARMON As String, _
                                    ByVal strCREDAY As String) As Integer
        Dim intRtn As Integer
        'PREESTNO,SEQ,JOBNO,YEARMON,CREDAY,CLIENTCODE,CLIENTSUBCODE,DIVAMT,JOBNAME
        intRtn = mobjcePD_DIVAMT.InsertDo( _
                                       GetElement(vntData, "PREESTNO", intColCnt, intRow), _
                                       dblSEQ, _
                                       GetElement(vntData, "JOBNO", intColCnt, intRow), _
                                       strYEARMON, _
                                       strCREDAY, _
                                       GetElement(vntData, "CLIENTCODE", intColCnt, intRow), _
                                       GetElement(vntData, "CLIENTSUBCODE", intColCnt, intRow), _
                                       GetElement(vntData, "DIVAMT", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "JOBNAME", intColCnt, intRow), _
                                       GetElement(vntData, "SUBSEQ", intColCnt, intRow), _
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
    End Function
    Private Function UpdateRtn(ByVal vntData As Object, _
                               ByVal intColCnt As Integer, _
                               ByVal intRow As Integer, _
                               ByVal strYEARMON As String, _
                               ByVal strCREDAY As String) As Integer
        'PREESTNO,SEQ,JOBNO,YEARMON,CREDAY,CLIENTCODE,CLIENTSUBCODE,DIVAMT,JOBNAME
        Dim intRtn As Integer
        intRtn = mobjcePD_DIVAMT.UpdateDo( _
                                       GetElement(vntData, "PREESTNO", intColCnt, intRow), _
                                       GetElement(vntData, "SEQ", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "JOBNO", intColCnt, intRow), _
                                       strYEARMON, _
                                       strCREDAY, _
                                       GetElement(vntData, "CLIENTCODE", intColCnt, intRow), _
                                       GetElement(vntData, "CLIENTSUBCODE", intColCnt, intRow), _
                                       GetElement(vntData, "DIVAMT", intColCnt, intRow, NULL_NUM, True), _
                                       GetElement(vntData, "JOBNAME", intColCnt, intRow), _
                                       GetElement(vntData, "SUBSEQ", intColCnt, intRow), _
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
